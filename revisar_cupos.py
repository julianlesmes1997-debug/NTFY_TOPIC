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
import urllib.request
from datetime import datetime, timedelta, timezone
from pathlib import Path

from playwright.sync_api import sync_playwright

URL = "https://agendamiento.dian.gov.co"
PASOS = ("Agendar cita", "Persona Natural", "Videoatención")
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


def revisar():
    r = {"fecha": datetime.now(COLOMBIA).strftime("%Y-%m-%d %H:%M"), "origen": ORIGEN,
         "hay_cupo": False, "detalle": "", "servicios": []}
    with sync_playwright() as p:
        nav = p.chromium.launch(headless=True)
        page = nav.new_page(viewport={"width": 1366, "height": 900}, locale="es-CO")
        try:
            page.goto(URL, wait_until="networkidle", timeout=90000)
            for paso in PASOS:
                if not clic_texto(page, paso):
                    raise RuntimeError(f"No se encontro la opcion '{paso}'")
            r["servicios"] = servicios_visibles(page.inner_text("body"))
            if not clic_texto(page, SERVICIO):
                raise RuntimeError(f"'{SERVICIO}' ya no aparece en Videoatencion")
            if SIN_CUPO in page.inner_text("body").lower():
                r["detalle"] = "Sin cupos (No se encontraron especialidades)"
            else:
                r["hay_cupo"] = True
                r["detalle"] = "HAY CUPO: Devoluciones ya no muestra el aviso de sin especialidades"
                page.screenshot(path=str(BASE / "cupo_encontrado.png"), full_page=True)
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


def main():
    hist = leer_historial()
    r = revisar()
    guardar(r, hist)
    print(json.dumps(r, ensure_ascii=False, indent=2))

    if r["hay_cupo"]:
        notificar("Hay citas DIAN - Devoluciones",
                  f"[{ORIGEN}] Apareció disponibilidad en Videoatención > Devoluciones. Ejecuta ejecutar.bat para agendar.",
                  "urgent")
    elif r["detalle"].startswith("ERROR"):
        ultimas = [h["detalle"] for h in hist if h.get("origen", "PC") == ORIGEN][-2:]
        if len(ultimas) == 2 and all(d.startswith("ERROR") for d in ultimas):
            notificar("Revisión DIAN fallando", f"[{ORIGEN}] 3 revisiones seguidas con error: {r['detalle'][:120]}", "high")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
