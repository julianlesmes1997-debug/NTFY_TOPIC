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
import urllib.request
from datetime import datetime, timedelta, timezone
from pathlib import Path

from playwright.sync_api import sync_playwright

URL = "https://agendamiento.dian.gov.co"
PASOS = ("Agendar cita", "Persona Natural", "Videoatención")
OPCIONES = ("Persona Natural", "Persona Jurídica", "Presencial", "Videoatención")
SERVICIO = "Devoluciones."
SIN_CUPO = "no se encontraron especialidades"
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
        xy = page.evaluate(JS_BUSCAR, texto)
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


def seleccionada(page, texto):
    """La tarjeta elegida se pinta azul (rgb 57,183,217); las demas quedan blancas."""
    return page.evaluate(JS_FONDO, texto) == "rgb(57, 183, 217)"


def revisar():
    r = {"fecha": datetime.now(COLOMBIA).strftime("%Y-%m-%d %H:%M"), "origen": ORIGEN,
         "hay_cupo": False, "detalle": "", "seleccion": [], "servicios": []}
    with sync_playwright() as p:
        nav = p.chromium.launch(headless=True)
        page = nav.new_page(viewport={"width": 1366, "height": 900}, locale="es-CO")
        try:
            page.goto(URL, wait_until="networkidle", timeout=90000)
            for paso in PASOS:
                for _ in range(3):  # reintenta si la tarjeta no quedo marcada
                    if not clic_texto(page, paso):
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
            if SIN_CUPO in page.inner_text("body").lower():
                r["detalle"] = "Sin cupos (No se encontraron especialidades)"
            else:
                r["hay_cupo"] = True
                r["detalle"] = "HAY CUPO: Devoluciones ya no muestra el aviso de sin especialidades"
                page.screenshot(path=str(BASE / "cupo_encontrado.png"), full_page=True)
                texto = page.inner_text("body")
                ini = texto.rfind("Trámite")
                if ini >= 0:  # opciones de tramite que ofrece la DIAN en ese momento
                    bloque = texto[ini + len("Trámite"):texto.find("Anterior", ini)]
                    r["tramites"] = [" ".join(l.split()) for l in bloque.split("\n") if l.strip()][:6]
        except Exception as e:
            r["detalle"] = f"ERROR: {str(e).splitlines()[0][:200]}"
        finally:
            nav.close()
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


INTERVALO = 5  # minutos entre revisiones: :00, :05, :10 ...


def esperar_marca():
    """Espera hasta la siguiente marca de INTERVALO minutos (hora Colombia)."""
    ahora = datetime.now(COLOMBIA)
    siguiente = ahora.replace(second=0, microsecond=0) + timedelta(minutes=INTERVALO - ahora.minute % INTERVALO)
    espera = (siguiente - ahora).total_seconds()
    print(f"Esperando {espera / 60:.1f} min hasta las {siguiente:%H:%M}...", flush=True)
    time.sleep(espera)


def toca_resumen():
    """Resumen 'sin citas': '1' siempre; 'hora' solo en la revision de la hora en punto."""
    modo = os.environ.get("NOTIFICAR_SIEMPRE", "")
    return modo == "1" or (modo == "hora" and datetime.now(COLOMBIA).minute < INTERVALO)


def main():
    """Una revision; o, con BUCLE_MINUTOS (nube), una revision en cada marca durante ese tiempo.
    En la nube GitHub no cumple horarios cortos, por eso el ritmo lo lleva este ciclo."""
    bucle = int(os.environ.get("BUCLE_MINUTOS") or 0)
    if not bucle:
        return una_revision()
    fin = time.monotonic() + bucle * 60
    while True:
        esperar_marca()
        if time.monotonic() > fin:
            return 0
        una_revision()


def una_revision():
    hist = leer_historial()
    for intento in range(2):  # la pagina de la DIAN a veces no carga a la primera
        r = revisar()
        if not r["detalle"].startswith("ERROR"):
            break
        print(f"Intento {intento + 1} con error: {r['detalle']}")
        time.sleep(15)
    guardar(r, hist)
    print(json.dumps(r, ensure_ascii=False, indent=2), flush=True)

    if r["hay_cupo"]:
        tramites = f" Trámites: {', '.join(r['tramites'])}." if r.get("tramites") else ""
        notificar("Hay citas DIAN - Devoluciones",
                  f"[{ORIGEN}] {r['fecha']} Hay disponibilidad en Videoatención > Devoluciones.{tramites} "
                  "Agenda YA con ejecutar.bat: los cupos se agotan rápido.",
                  "urgent")
    elif r["detalle"].startswith("ERROR"):
        ultimas = [h["detalle"] for h in hist if h.get("origen", "PC") == ORIGEN][-2:]
        if len(ultimas) == 2 and all(d.startswith("ERROR") for d in ultimas):
            notificar("Revisión DIAN fallando", f"[{ORIGEN}] 3 revisiones seguidas con error: {r['detalle'][:120]}", "high")
        return 1
    elif toca_resumen():
        notificar("DIAN: sin citas todavía",
                  f"[{ORIGEN}] {r['fecha']} - {' > '.join(r['seleccion'])} > Devoluciones: sin cupos. "
                  f"Reviso cada {INTERVALO} min; próximo resumen en 1 hora.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
