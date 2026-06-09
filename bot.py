import os, requests
from datetime import datetime, timezone, timedelta
import openpyxl
from openpyxl import Workbook

API_KEY = os.environ["API_FOOTBALL_KEY"]
TELEGRAM_TOKEN = os.environ["TELEGRAM_TOKEN"]
CHAT_ID = "691112681"
BASE = "https://v3.football.api-sports.io"
HEADERS = {"x-apisports-key": API_KEY}

PALABRAS_LIGAS = ["World Cup","Brazil","Argentina","Colombia","MLS","USA",
                  "Norway","Sweden","Finland","Japan","Korea"]
BANDA_MIN, BANDA_MAX = 2.5, 2.9
MAX_ODDS = 30
EXCEL = "registro_empates.xlsx"

BOGOTA = timezone(timedelta(hours=-5))
hoy = datetime.now(BOGOTA).strftime("%Y-%m-%d")

def enviar_telegram(texto):
    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    requests.post(url, data={"chat_id": CHAT_ID, "text": texto})

usados = 0
r = requests.get(f"{BASE}/fixtures", headers=HEADERS,
                 params={"date": hoy, "timezone": "America/Bogota"})
usados += 1
fixtures = r.json().get("response", [])

def nos_interesa(fx):
    t = ((fx["league"]["name"] or "")+" "+(fx["league"]["country"] or "")).lower()
    return any(p.lower() in t for p in PALABRAS_LIGAS)
candidatos = [fx for fx in fixtures if nos_interesa(fx)]

filas = []
for fx in candidatos[:MAX_ODDS]:
    fid = fx["fixture"]["id"]
    local = fx["teams"]["home"]["name"]
    visita = fx["teams"]["away"]["name"]
    liga = fx["league"]["name"]
    hora = fx["fixture"]["date"][11:16]
    ro = requests.get(f"{BASE}/odds", headers=HEADERS, params={"fixture": fid})
    usados += 1
    cuotas = []
    for bloque in ro.json().get("response", []):
        for casa in bloque.get("bookmakers", []):
            for bet in casa.get("bets", []):
                if bet.get("name") == "Match Winner":
                    for v in bet.get("values", []):
                        if v.get("value") == "Draw":
                            try: cuotas.append(float(v["odd"]))
                            except: pass
    if not cuotas: continue
    cuota = round(sum(cuotas)/len(cuotas), 2)
    filas.append({"liga":liga,"partido":f"{local} vs {visita}","hora":hora,
                  "cuota":cuota,"prob":round(100/cuota,1)})

filas.sort(key=lambda x:-x["prob"])
en_banda = [x for x in filas if BANDA_MIN <= x["cuota"] <= BANDA_MAX]

if en_banda:
    seleccion = en_banda; tipo = "OFICIAL"
    titulo = f"⚽ ALERTAS DE EMPATE — {hoy}"
else:
    seleccion = filas[:3]; tipo = "muestra"
    titulo = f"(Prueba nube) Sin candidatos en banda hoy {hoy}. Mas cercanos:"

if seleccion:
    msg = titulo + "\n"
    for x in seleccion:
        msg += f"\n• {x['partido']}\n  {x['liga']} | {x['hora']}\n  Empate: cuota {x['cuota']} ({x['prob']}%)\n"
else:
    msg = f"(Prueba nube) Hoy {hoy} no hubo partidos con cuota."
enviar_telegram(msg)

if os.path.exists(EXCEL):
    wb = openpyxl.load_workbook(EXCEL); ws = wb.active
else:
    wb = Workbook(); ws = wb.active; ws.title = "Registro"
    ws.append(["Fecha","Partido","Liga","Hora","Cuota empate","Prob %",
               "Tipo","Resultado real","Marcador","Acerto?"])
for x in seleccion:
    ws.append([hoy, x["partido"], x["liga"], x["hora"], x["cuota"], x["prob"],
               tipo, "", "", ""])
wb.save(EXCEL)

print(f"OK. Enviado a Telegram. Filas: {len(seleccion)} ({tipo}). API usados: {usados}")
