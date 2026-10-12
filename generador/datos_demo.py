"""
Genera los datos ficticios del plan a partir de datos-demo/config.json:

  datos-demo/movimientos.json   los movimientos de las cuatro cuentas, ya categorizados
  datos-demo/saldos.json        el saldo de cada cuenta a fin de mes
  ../extractos-ejemplo/         los mismos movimientos en el formato que descarga cada banco

Los extractos llevan las cabeceras que busca el importador (BANCOS en apps-script/importar.js).
El resto de la estructura (filas de título, columnas que el importador no usa) es aproximada.

Uso:  python generador/datos_demo.py        (necesita xlwt:  pip install xlwt)
"""
import datetime as dt
import json
import random
import unicodedata
from pathlib import Path

AQUI = Path(__file__).resolve().parent
DATOS = AQUI / "datos-demo"
EXTRACTOS = AQUI.parent / "extractos-ejemplo"
CFG = json.loads((DATOS / "config.json").read_text(encoding="utf-8"))
POR_REVISAR = "❓ Por revisar"

random.seed(2026)


# ---------------------------------------------------------------- reglas (igual que en importar.js)
def norm(s):
    s = unicodedata.normalize("NFD", str(s or ""))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return " ".join(s.upper().split())


def aplica_reglas(m):
    c, cu = norm(m["concepto"]), norm(m["cuenta"])
    for pat, cuenta, signo, cat, _ in CFG["reglas"]:
        if not pat or not cat:
            continue
        if cuenta and norm(cuenta) != cu:
            continue
        if signo == "+" and not m["importe"] > 0:
            continue
        if signo == "-" and not m["importe"] < 0:
            continue
        if norm(pat) in c:
            return cat
    return POR_REVISAR


# ---------------------------------------------------------------- movimientos
desde = dt.date.fromisoformat(CFG["periodo"]["desde"])
hasta = dt.date.fromisoformat(CFG["periodo"]["hasta"])
movs = []   # cuenta, fecha, hora, concepto, importe


def add(cuenta, fecha, concepto, importe, hora=""):
    if desde <= fecha <= hasta:
        movs.append({"cuenta": cuenta, "fecha": fecha, "hora": hora, "concepto": concepto, "importe": round(importe, 2)})


def meses():
    m = desde.replace(day=1)
    while m <= hasta:
        yield m
        m = (m.replace(day=28) + dt.timedelta(days=4)).replace(day=1)


def dia(m, d):
    ultimo = ((m.replace(day=28) + dt.timedelta(days=4)).replace(day=1) - dt.timedelta(days=1)).day
    return m.replace(day=min(d, ultimo))


C, O, T, I = "CaixaBank", "Openbank", "Openbank tarjeta", "ING"
TARJETA = [  # concepto, mín, máx, veces al mes
    ("SUPERMERCADO EL BARRIO", 18, 70, 4), ("FRUTERIA LA HUERTA", 8, 22, 2), ("CAFETERIA CENTRAL", 2.5, 9, 5),
    ("RESTAURANTE LA PLAZA", 22, 58, 2), ("TRANSPORTE PUBLICO RECARGA", 20, 20, 1), ("GASOLINERA RUTA", 35, 55, 1),
    ("FARMACIA SAN JOSE", 6, 28, 1), ("TIENDA ONLINE", 12, 60, 1), ("CINE ESTRENOS", 9, 18, 1),
]
tarjeta_mes = {}
for m in meses():
    # CaixaBank: la nómina y los recibos de casa
    add(C, dia(m, 28), "NOMINA EMPRESA EJEMPLO SL", 2150)
    if m.month in (6, 12):
        add(C, dia(m, 20), "NOMINA PAGA EXTRA EMPRESA EJEMPLO SL", 2150)
    add(C, dia(m, 1), f"ALQUILER VIVIENDA {m:%m/%Y}", -850)
    add(C, dia(m, 2), "TRASPASO A CUENTA AHORRO ING", -300)
    add(C, dia(m, 2), "TRASPASO A OPENBANK", -480)
    if m >= dt.date.fromisoformat(CFG["viajes"]["primera_aportacion"]):
        add(C, dia(m, 3), "TRANSFERENCIA CUENTA DE VIAJES", -CFG["viajes"]["aportacion"])
    add(C, dia(m, 8), "RECIBO LUZ HOGAR ENERGIA", -random.uniform(42, 78) if m.month in (12, 1, 2, 7, 8) else -random.uniform(35, 55))
    if m.month % 2 == 0:
        add(C, dia(m, 12), "RECIBO AGUAS MUNICIPALES", -random.uniform(32, 41))
    add(C, dia(m, 5), "RECIBO FIBRANET", -39.90)
    add(C, dia(m, 6), "RECIBO MOVILCOM", -14.95)
    add(C, dia(m, 9), "RECIBO STREAMFLIX", -12.99)
    if m.month == 3:
        add(C, dia(m, 15), "RECIBO SEGURO HOGAR INQUILINO", -148.00)
    if m.month == 7:
        add(C, dia(m, 18), "DEVOLUCION AEAT IRPF", 412.36)
    if m.month in (11, 4):
        add(C, dia(m, 14), "MODA CENTRO", -random.uniform(45, 110))
    if m.month % 2:
        add(C, dia(m, 22), "PELUQUERIA ESTILO", -22)
    # Bizum de ajustes de comida (sin regla: quedan por revisar)
    for d in (11, 24):
        add(C, dia(m, d), "BIZUM " + random.choice(["CENA", "COMPRA", "COMIDA DOMINGO"]), random.choice([-1, 1]) * random.randint(12, 38))
    # Openbank: recibe dinero y paga la tarjeta del mes anterior
    add(O, dia(m, 2), "TRASPASO DESDE CAIXABANK", 480)
    if tarjeta_mes:
        anterior = max(tarjeta_mes)
        add(O, dia(m, 5), "LIQUIDACION TARJETA CREDITO", -round(tarjeta_mes.pop(anterior), 2))
    # Tarjeta de Openbank: el día a día
    total = 0
    for concepto, a, b, n in TARJETA:
        for _ in range(n):
            f = dia(m, random.randint(1, 28))
            imp = -round(random.uniform(a, b), 2)
            add(T, f, concepto, imp, hora=f"{random.randint(8, 21):02d}:{random.choice([0, 7, 15, 22, 31, 45, 52]):02d}")
            total += -imp
    tarjeta_mes[m] = total
    # ING: el ahorro
    add(I, dia(m, 2), "TRASPASO DESDE CAIXABANK", 300)

# Intereses de ING, sobre el saldo que haya
saldo_ing = CFG["saldos_iniciales"][I]
for m in meses():
    saldo_ing += 300
    add(I, dia(m, 31), "INTERESES CUENTA AHORRO", round(saldo_ing * 0.0075 / 12, 2))
    saldo_ing += round(saldo_ing * 0.0075 / 12, 2)

# ---------------------------------------------------------------- saldos y claves (igual que el importador)
movs.sort(key=lambda m: (m["cuenta"], m["fecha"], m["hora"], -m["importe"]))
saldo = dict(CFG["saldos_iniciales"])
vistos = {}
for m in movs:
    if m["cuenta"] == T:
        m["saldo"] = ""
    else:
        saldo[m["cuenta"]] = round(saldo[m["cuenta"]] + m["importe"], 2)
        m["saldo"] = saldo[m["cuenta"]]
    base = "|".join([m["cuenta"], m["fecha"].isoformat(), m["hora"], f'{m["importe"]:.2f}',
                     "" if m["saldo"] == "" else f'{m["saldo"]:.2f}'])
    vistos[base] = vistos.get(base, 0) + 1
    m["clave"] = f"{base}#{vistos[base]}"
    m["categoria"] = aplica_reglas(m)

# Los Bizum de antes de septiembre ya están revisados; los últimos se quedan pendientes
for m in movs:
    if m["categoria"] == POR_REVISAR and m["fecha"] < dt.date(hasta.year, hasta.month, 1):
        m["categoria"] = "Ajustes de gastos compartidos (comida)"

for cu in CFG["cuentas"]:
    if any(m["saldo"] != "" and m["saldo"] < 0 for m in movs if m["cuenta"] == cu):
        raise SystemExit(f"{cu} se queda en negativo: revisa los importes del generador")

meses_l = list(meses())
saldos = {"meses": [m.isoformat() for m in meses_l], "saldos": {}}
for cu in CFG["cuentas"]:
    fila = []
    for m in meses_l:
        fin = dia(m, 31)
        hechos = [x for x in movs if x["cuenta"] == cu and x["fecha"] <= fin]
        fila.append(hechos[-1]["saldo"] if hechos else "")
    saldos["saldos"][cu] = fila

out = sorted(movs, key=lambda m: (m["fecha"], m["hora"]), reverse=True)
(DATOS / "movimientos.json").write_text(json.dumps(
    [{**m, "fecha": m["fecha"].isoformat()} for m in out], ensure_ascii=False, indent=1), encoding="utf-8")
(DATOS / "saldos.json").write_text(json.dumps(saldos, ensure_ascii=False, indent=1), encoding="utf-8")


# ---------------------------------------------------------------- extractos de ejemplo
def es(n):  # 1234.5 -> "1.234,50"
    s = f"{abs(n):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return ("-" if n < 0 else "") + s


def ddmmaaaa(d):
    return d.strftime("%d/%m/%Y")


def html(titulo, cab, filas):
    tr = lambda cs, t: "<tr>" + "".join(f"<{t}>{c}</{t}>" for c in cs) + "</tr>"
    cuerpo = [tr([titulo], "td"), tr(["Cuenta: ES00 0000 0000 0000 0000 0000"], "td"), tr([""], "td"), tr(cab, "th")]
    cuerpo += [tr(f, "td") for f in filas]
    return ('<html><head><meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1"></head>'
            "<body><table>" + "\n".join(cuerpo) + "</table></body></html>")


def de(cuenta):
    return [m for m in out if m["cuenta"] == cuenta]


EXTRACTOS.mkdir(exist_ok=True)
(EXTRACTOS / "openbank-cuenta.xls").write_bytes(html(
    "Movimientos de cuenta", ["Fecha Operación", "Fecha Valor", "Concepto", "Importe", "Saldo"],
    [[ddmmaaaa(m["fecha"]), ddmmaaaa(m["fecha"]), m["concepto"], es(m["importe"]), es(m["saldo"])] for m in de(O)]
).encode("latin-1"))
(EXTRACTOS / "openbank-tarjeta.xls").write_bytes(html(
    "Movimientos de tarjeta", ["Fecha operación", "Hora", "Concepto", "Situación", "Importe"],
    [[ddmmaaaa(m["fecha"]), m["hora"], m["concepto"], "Liquidado", es(m["importe"])] for m in de(T)]
).encode("latin-1"))

import xlwt  # noqa: E402  (solo hace falta para estos dos)

FECHA = xlwt.easyxf(num_format_str="DD/MM/YYYY")
EUROS = xlwt.easyxf(num_format_str="#,##0.00")


def xls(nombre, preambulo, cab, filas, tipos):
    wb = xlwt.Workbook(encoding="utf-8")
    ws = wb.add_sheet("Movimientos")
    r = 0
    for linea in preambulo:
        ws.write(r, 0, linea); r += 1
    r += 1
    for j, c in enumerate(cab):
        ws.write(r, j, c)
    for f in filas:
        r += 1
        for j, (v, t) in enumerate(zip(f, tipos)):
            if t == "f":
                ws.write(r, j, v, FECHA)
            elif t == "e":
                ws.write(r, j, v, EUROS)
            else:
                ws.write(r, j, v)
    wb.save(str(EXTRACTOS / nombre))


xls("caixabank.xls", ["Movimientos de la cuenta", "ES00 0000 0000 0000 0000 0000"],
    ["Fecha", "Fecha valor", "Movimiento", "Más datos", "Importe", "Saldo"],
    [[m["fecha"], m["fecha"], m["concepto"], "", m["importe"], m["saldo"]] for m in de(C)],
    "ffssee")
SUB = {"Intereses": ("Ingresos", "Intereses"), "Traspaso entre mis cuentas": ("Movimientos excluidos", "Traspaso")}
xls("ing.xls", ["Cuenta Ahorro", "ES00 0000 0000 0000 0000 0000"],
    ["F. VALOR", "CATEGORÍA", "SUBCATEGORÍA", "DESCRIPCIÓN", "COMENTARIO", "IMAGEN", "IMPORTE (€)", "SALDO (€)"],
    [[m["fecha"], *SUB.get(m["categoria"], ("Otros", "Otros")), m["concepto"], "", "", m["importe"], m["saldo"]] for m in de(I)],
    "fsssssee")

n = {cu: len(de(cu)) for cu in (C, O, T, I)}
print("movimientos:", len(out), n)
print("saldos a fin de periodo:", {cu: saldos["saldos"][cu][-1] for cu in CFG["cuentas"]})
print("por revisar:", sum(1 for m in out if m["categoria"] == POR_REVISAR))
