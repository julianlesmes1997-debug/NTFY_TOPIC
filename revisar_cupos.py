"""Revision automatica de cupos DIAN: Persona Natural -> Videoatencion -> Devoluciones.

Funciona igual en el PC (Programador de tareas) y en la nube (GitHub Actions).
No usa datos personales: solo mira si hay cupos y avisa por ntfy.
  - NTFY_TOPIC: canal de ntfy (variable de entorno o archivo ntfy_topic.txt)
  - ORIGEN: texto para identificar quien avisa ("PC" o "Nube")
"""
import csv
import json
import os
import subprocess
import sys
import time
import unicodedata
import urllib.request
from datetime import datetime, timedelta, timezone
from pathlib import Path

from playwright.sync_api import sync_playwright

URL = "https://agendamiento.dian.gov.co"
PASOS = ("Agendar cita", "Persona Natural", "Videoatención")
OPCIONES = ("Persona Natural", "Persona Jurídica", "Presencial", "Videoatención")
SERVICIO = "Devoluciones."
SIN_CUPO = "no se encontraron especialidades"
# Tramite buscado: Bogota "Solicitud de devolucion y/o compensacion Vehiculos electricos o
# Hibridos persona natural". Debe contener todas estas palabras (sin tildes, en minuscula).
OBJETIVO = ("bogota", "vehiculo")
BASE = Path(__file__).resolve().parent
HIST = BASE / "historial_revisiones.csv"
ANALISIS = BASE / "analisis_revisiones.md"
ORIGEN = os.environ.get("ORIGEN", "PC")
COLOMBIA = timezone(timedelta(hours=-5))

JS_BUSCAR = """(t) => { const n = s => (s||'').normalize('NFD').replace(/[\\u0300-\\u036f]/g,'').replace(/\\s+/g,' ').trim().toLowerCase();
  const c = [...document.querySelectorAll('body *')].filter(e => { const r = e.getBoundingClientRect();
    return r.width > 0 && r.height > 0 && getComputedStyle(e).visibility !== 'hidden' && n(e.innerText) === n(t); });
  c.sort((a,b) => { const ra = a.getBoundingClientRect(), rb = b.getBoundingClientRect(); return ra.width*ra.height - rb.width*rb.height; });
  if (!c.length) return null; c[0].scrollIntoView({block:'center'}); const r = c[0].getBoundingClientRect();
  return [r.x + r.width/2, r.y + r.height/2]; }"""

# Color de fondo de la tarjeta que contiene el texto (primer ancestro con fondo)
JS_FONDO = """(t) => { const n = s => (s||'').replace(/\\s+/g,' ').trim();
  const el = [...document.querySelectorAll('body *')].filter(e => n(e.innerText) === t && e.getBoundingClientRect().height > 0)
    .sort((a,b) => a.getBoundingClientRect().height - b.getBoundingClientRect().height)[0];
  for (let p = el, k = 0; p && k < 8; p = p.parentElement, k++) { const c = getComputedStyle(p).backgroundColor;
    if (c && c !== 'rgba(0, 0, 0, 0)') return c; }
  return ''; }"""


def ntfy_topic():
    t = os.environ.get("NTFY_TOPIC", "").strip()
    f = BASE / "ntfy_topic.txt"
    return t or (f.read_text(encoding="utf-8").strip() if f.exists() else "")


def notificar(titulo, mensaje, prioridad="default"):
    topic = ntfy_topic()
    if topic:
        req = urllib.request.Request(
            f"https://ntfy.sh/{topic}", data=mensaje.encode("utf-8"), method="POST",
            headers={"Title": titulo.encode("utf-8").decode("latin-1", "ignore"), "Priority": prioridad,
                     "Tags": "calendar", "Click": URL})
        try:
            urllib.request.urlopen(req, timeout=20)
        except Exception as e:
            print(f"No se pudo enviar ntfy: {e}")
    if os.name == "nt":  # notificacion de Windows, sin ventanas
        ps = ("[Windows.UI.Notifications.ToastNotificationManager,Windows.UI.Notifications,ContentType=WindowsRuntime]|Out-Null;"
              "$x=[Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent(1);"
              f"$t=$x.GetElementsByTagName('text');$t.Item(0).InnerText='{titulo}';$t.Item(1).InnerText='{mensaje}';"
              "[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier('DIAN Citas').Show("
              "[Windows.UI.Notifications.ToastNotification]::new($x))")
        subprocess.run(["powershell", "-NoProfile", "-WindowStyle", "Hidden", "-Command", ps],
                       creationflags=0x08000000, check=False)


def clic_texto(page, texto, espera=20):
    for _ in range(espera):
        try:
            xy = page.evaluate(JS_BUSCAR, texto)
        except Exception as e:  # la DIAN a veces recarga la pagina sola: esperar y reintentar
            if "context was destroyed" not in str(e) and "navigat" not in str(e):
                raise
            xy = None
        if xy:
            page.mouse.click(*xy)
            page.wait_for_timeout(2500)
            return True
        page.wait_for_timeout(1000)
    return False


def servicios_visibles(texto):
    ini = texto.find("Seleccione el tipo de servicio")
    if ini < 0:
        return []
    fin = texto.find("Anterior", ini)
    bloque = texto[ini + len("Seleccione el tipo de servicio"):fin if fin > 0 else None]
    return [" ".join(l.split()) for l in bloque.split("\n") if l.strip()]


def es_objetivo(tramite):
    sin_tildes = "".join(c for c in unicodedata.normalize("NFD", tramite.lower()) if unicodedata.category(c) != "Mn")
    return all(k in sin_tildes for k in OBJETIVO)


def seleccionada(page, texto):
    """La tarjeta elegida se pinta azul (rgb 57,183,217); las demas quedan blancas."""
    return page.evaluate(JS_FONDO, texto) == "rgb(57, 183, 217)"


def tramites_visibles(texto):
    """Opciones listadas bajo el encabezado 'Trámite' (solo aparecen cuando hay cupo)."""
    ini = texto.rfind("Trámite")
    if ini < 0:
        return []
    fin = texto.find("Anterior", ini)
    bloque = texto[ini + len("Trámite"):fin if fin > 0 else None]
    return [" ".join(l.split()) for l in bloque.split("\n") if l.strip()][:6]


def esperar_resultado(page, segundos=60):
    """Tras elegir Devoluciones la DIAN muestra 'Cargando'. Solo hay dos respuestas validas:
    el aviso de 'sin especialidades' o la lista de tramites. Mientras no llegue ninguna, se espera;
    si no llega, devuelve None (nunca se asume cupo por una pagina que no termino de cargar)."""
    for _ in range(segundos):
        try:
            texto = page.inner_text("body")
        except Exception:  # recarga de la pagina en curso
            page.wait_for_timeout(1000)
            continue
        if SIN_CUPO in texto.lower():
            return "sin_cupo"
        tramites = tramites_visibles(texto)
        if tramites:
            return tramites
        page.wait_for_timeout(1000)
    return None


def revisar(nav=None, agendar=False):
    """Una revision. Con `nav` reutiliza un navegador ya abierto (vigilante del PC: mas rapido).
    Con `agendar`, si aparece el tramite objetivo sigue en la misma pagina hasta agendar."""
    if nav is None:
        with sync_playwright() as p:
            navegador = p.chromium.launch(headless=True)
            try:
                return revisar(navegador, agendar)
            finally:
                navegador.close()
    r = {"fecha": datetime.now(COLOMBIA).strftime("%Y-%m-%d %H:%M"), "origen": ORIGEN,
         "hay_cupo": False, "detalle": "", "seleccion": [], "servicios": []}
    page = nav.new_page(viewport={"width": 1366, "height": 900}, locale="es-CO")
    try:
        page.goto(URL, wait_until="networkidle", timeout=90000)
        for paso in PASOS:
            for _ in range(3):  # reintenta si la tarjeta no quedo marcada
                # la portada de la DIAN a veces tarda mas de 20 s en mostrar el menu
                if not clic_texto(page, paso, espera=60 if paso == PASOS[0] else 20):
                    raise RuntimeError(f"No se encontro la opcion '{paso}'")
                if paso == PASOS[0] or seleccionada(page, paso):
                    break
            else:
                raise RuntimeError(f"La opcion '{paso}' no quedo seleccionada")
        r["seleccion"] = [t for t in OPCIONES if seleccionada(page, t)]
        if r["seleccion"] != list(PASOS[1:]):
            raise RuntimeError(f"Seleccion incorrecta: {r['seleccion']}")
        page.screenshot(path=str(BASE / "seleccion.png"), full_page=True)
        r["servicios"] = servicios_visibles(page.inner_text("body"))
        if not clic_texto(page, SERVICIO):
            raise RuntimeError(f"'{SERVICIO}' ya no aparece en Videoatencion")
        resultado = esperar_resultado(page)
        if resultado == "sin_cupo":
            r["detalle"] = "Sin cupos (No se encontraron especialidades)"
        elif resultado:
            r["hay_cupo"] = True
            r["tramites"] = resultado
            r["objetivo"] = any(es_objetivo(t) for t in resultado)
            r["detalle"] = "HAY CUPO de " + ("tu tramite (Bogota vehiculos)" if r["objetivo"] else "otro tramite")
            page.screenshot(path=str(BASE / "cupo_encontrado.png"), full_page=True)
            if r["objetivo"] and agendar:  # sin perder segundos: seguir en esta misma pagina
                notificar("Hay cita DIAN de TU trámite (Bogotá vehículos)",
                          f"[{ORIGEN}] {r['fecha']} Apareció el cupo. Intentando agendar ya en tu PC...", "urgent")
                if str(BASE) not in sys.path:  # el lanzador corre con -I: la carpeta no esta en la ruta
                    sys.path.insert(0, str(BASE))
                import agendar_auto
                r["agendamiento"] = agendar_auto.continuar(page, visible=False)
        else:
            page.screenshot(path=str(BASE / "sin_respuesta.png"), full_page=True)
            raise RuntimeError("La DIAN no respondio en 60 s (seguia 'Cargando')")
    except Exception as e:
        r["detalle"] = f"ERROR: {str(e).splitlines()[0][:200]}"
    finally:
        page.close()
    return r


def leer_historial():
    if not HIST.exists():
        return []
    with HIST.open(encoding="utf-8-sig") as f:
        return list(csv.DictReader(f, delimiter=";"))


def guardar(r, hist):
    nuevo = not HIST.exists()
    with HIST.open("a", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        if nuevo:
            w.writerow(["fecha", "origen", "hay_cupo", "detalle", "servicios"])
        w.writerow([r["fecha"], r["origen"], "SI" if r["hay_cupo"] else "NO", r["detalle"], " | ".join(r["servicios"])])

    # Con revisiones cada 5 min, el analisis solo se escribe si cambia el resultado o cambia la hora
    if hist and hist[-1]["detalle"] == r["detalle"] and hist[-1]["fecha"][:13] == r["fecha"][:13]:
        return
    total = len(hist) + 1
    con_cupo = sum(h["hay_cupo"] == "SI" for h in hist) + r["hay_cupo"]
    errores = sum(h["detalle"].startswith("ERROR") for h in hist) + r["detalle"].startswith("ERROR")
    lineas = [f"### {r['fecha']} ({r['origen']})",
              f"- Resultado: **{r['detalle']}**",
              f"- Revisiones: {total} desde {hist[0]['fecha'] if hist else r['fecha']} | con cupo: {con_cupo} | con error: {errores}"]
    if hist and r["servicios"] and hist[-1].get("servicios") and hist[-1]["servicios"] != " | ".join(r["servicios"]):
        antes, ahora = set(hist[-1]["servicios"].split(" | ")), set(r["servicios"])
        lineas.append(f"- Cambio en servicios de Videoatencion: nuevos {sorted(ahora - antes)}, retirados {sorted(antes - ahora)}")
    with ANALISIS.open("a", encoding="utf-8") as f:
        f.write("\n".join(lineas) + "\n\n")


# minutos entre revisiones del ciclo (debe dividir 60): la nube usa 3, por defecto 5
INTERVALO = int(os.environ.get("INTERVALO_MIN") or 5)


def esperar_marca(minutos=None):
    """Espera hasta la siguiente marca de `minutos` (por defecto INTERVALO), hora Colombia."""
    minutos = minutos or INTERVALO
    ahora = datetime.now(COLOMBIA)
    siguiente = ahora.replace(second=0, microsecond=0) + timedelta(minutes=minutos - ahora.minute % minutos)
    espera = (siguiente - ahora).total_seconds()
    print(f"Esperando {espera / 60:.1f} min hasta las {siguiente:%H:%M}...", flush=True)
    time.sleep(espera)


def toca_resumen():
    """Resumen 'sin citas': '1' siempre; 'hora' solo en la revision de la hora en punto."""
    modo = os.environ.get("NOTIFICAR_SIEMPRE", "")
    return modo == "1" or (modo == "hora" and datetime.now(COLOMBIA).minute < INTERVALO)


def horario_oficina(ahora):
    """Lunes a viernes de 7:00 a 18:00: cuando la DIAN ha publicado todos los cupos vistos."""
    return ahora.weekday() < 5 and 7 <= ahora.hour < 18


def vigilar():
    """Vigilante del PC: proceso continuo que revisa cada 1 min en horario de oficina y cada 5 min
    el resto, reutilizando el navegador. Si aparece el tramite objetivo agenda en la misma pagina
    (no vuelve a intentarlo antes de 2 min, ni nunca mas si ya quedo agendada)."""
    ultimo_intento = 0.0
    with sync_playwright() as p:
        nav, abierto_desde = p.chromium.launch(headless=True), time.monotonic()
        while True:
            agendar = (not (BASE / "cita_agendada.json").exists()
                       and time.monotonic() - ultimo_intento > 120)
            try:
                r = una_revision(nav, agendar)
                if r.get("agendamiento"):
                    ultimo_intento = time.monotonic()
            except Exception as e:  # navegador caido: se reabre
                print(f"Reiniciando navegador: {e}", flush=True)
                abierto_desde = 0.0
            if time.monotonic() - abierto_desde > 2 * 3600:  # renovar el navegador cada 2 h
                try:
                    nav.close()
                except Exception:
                    pass
                nav, abierto_desde = p.chromium.launch(headless=True), time.monotonic()
            esperar_marca(1 if horario_oficina(datetime.now(COLOMBIA)) else 5)


def main():
    """PC: vigilante continuo. Nube (BUCLE_MINUTOS): una revision en cada marca durante ese tiempo,
    porque GitHub no cumple horarios cortos. Otros casos (pruebas): una sola revision."""
    bucle = int(os.environ.get("BUCLE_MINUTOS") or 0)
    if not bucle:
        if vigilando_pc():
            return vigilar()
        return 1 if una_revision()["detalle"].startswith("ERROR") else 0
    fin = time.monotonic() + bucle * 60
    while True:
        esperar_marca()
        if time.monotonic() > fin:
            return 0
        una_revision()


def una_revision(nav=None, agendar=False):
    hist = leer_historial()
    for intento in range(2):  # la pagina de la DIAN a veces no carga a la primera
        r = revisar(nav, agendar)
        if not r["detalle"].startswith("ERROR"):
            break
        print(f"Intento {intento + 1} con error: {r['detalle']}")
        time.sleep(15)
    guardar(r, hist)
    print(json.dumps(r, ensure_ascii=False, indent=2), flush=True)

    # Avisar solo cuando cambia: si la revision anterior ya veia el mismo cupo, no se repite
    previo = next((h["detalle"] for h in reversed(hist) if h.get("origen", "PC") == ORIGEN), "")
    if r["hay_cupo"] and previo != r["detalle"]:
        tramites = f" Trámites: {', '.join(r['tramites'])}." if r.get("tramites") else ""
        if r.get("objetivo"):  # alarma solo para el tramite objetivo
            if not r.get("agendamiento"):  # si se intento agendar, ese flujo ya aviso
                notificar("Hay cita DIAN de TU trámite (Bogotá vehículos)",
                          f"[{ORIGEN}] {r['fecha']} Videoatención > Devoluciones.{tramites} Agenda YA: dura segundos.",
                          "urgent")
        else:
            notificar("Cupo DIAN de otro trámite",
                      f"[{ORIGEN}] {r['fecha']} Videoatención > Devoluciones.{tramites}")
    elif r["detalle"].startswith("ERROR"):
        # avisar una sola vez por racha: justo en el 3er error seguido, no en los siguientes
        ultimas = [h["detalle"].startswith("ERROR") for h in hist if h.get("origen", "PC") == ORIGEN][-3:]
        if ultimas[-2:] == [True, True] and (len(ultimas) < 3 or not ultimas[0]):
            notificar("Revisión DIAN fallando", f"[{ORIGEN}] 3 revisiones seguidas con error: {r['detalle'][:120]}", "high")
    elif not r["hay_cupo"] and toca_resumen():
        notificar("DIAN: sin citas todavía",
                  f"[{ORIGEN}] {r['fecha']} - {' > '.join(r['seleccion'])} > Devoluciones: sin cupos. "
                  f"Reviso cada {INTERVALO} min; próximo resumen en 1 hora.")
    return r


def vigilando_pc():
    return ORIGEN == "PC" and os.environ.get("GITHUB_ACTIONS") != "true"


if __name__ == "__main__":
    sys.exit(main())
