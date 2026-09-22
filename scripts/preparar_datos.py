"""Genera dashboard/index.html a partir del CSV de consumo de lotes de producción.

Uso:
    python3 scripts/preparar_datos.py [ruta_csv] [ruta_maestro_xlsx]

Reglas:
- Rend. Met (%) = Kgs producidos / Kgs consumidos * 100 (no usa despunte ni laminilla).
- Órdenes con rendimiento > 100 % se excluyen del cálculo y se listan como "a corregir".
- Los productos "SEMI" se unifican con su equivalente (ANGULO SEMI = ANGULO, etc.).
- Si existe el maestro de productos (datos/Datos_productos.xlsx) se usa su columna
  Tipo para el grupo y CORTE T4 para el largo; si un producto no está, se deduce del nombre.
- El CSV exportado trae varias tablas apiladas con los mismos datos; se usa solo la primera.
"""
import csv
import json
import re
import sys
from datetime import datetime
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
CSV_DEFAULT = RAIZ / "datos" / "Consumo_lotes_de_produccion.csv"
MAESTRO_DEFAULT = RAIZ / "datos" / "Datos_productos.xlsx"
PLANTILLA = RAIZ / "scripts" / "plantilla_dashboard.html"
SALIDA = RAIZ / "dashboard" / "index.html"

PULGADA_MM = 25.4
LARGO_ANGULO_DEFAULT = 6.0  # los ANGULO SEMI ... A270ES no indican largo


def num(s):
    s = (s or "").strip().replace(",", "")
    return float(s) if s else None


def leer_primera_tabla(ruta):
    with open(ruta, encoding="utf-8-sig", newline="") as f:
        lineas = f.read().splitlines()
    tabla = []
    for linea in lineas:
        if not linea.strip():
            break  # fin de la primera tabla
        tabla.append(linea)
    return list(csv.DictReader(tabla))


def pulgadas_a_mm(txt):
    """'1 3/4' -> 44.45 ; '1/2' -> 12.7 ; '1' -> 25.4"""
    total = 0.0
    for parte in txt.split():
        if "/" in parte:
            a, b = parte.split("/")
            total += float(a) / float(b)
        else:
            total += float(parte)
    return round(total * PULGADA_MM, 1)


def dec(s):
    return float(s.replace(",", "."))


def largo_m(nombre):
    # 12M, 10,5M, 10.5M, 7,315m, 6m — sin confundir con MM
    m = re.search(r"(?<![\d,.])(\d+(?:[.,]\d+)?)\s*M(?!M)\b", nombre, re.I)
    return dec(m.group(1)) if m else None


def clasificar(nombre):
    n = " ".join(nombre.upper().split())
    info = {"grupo": "Otros", "diametro": None, "ancho": None, "espesor": None,
            "lado": None, "largo": None, "largo_asumido": False, "pulgadas": None}

    if n.startswith("B HORMIGON"):
        info["grupo"] = "B Hormigón"
    elif n.startswith("REDONDO LISO") or n.startswith("BARRA TRASPASO"):
        info["grupo"] = "Redondo liso"
    elif n.startswith("PLANA"):
        info["grupo"] = "Plana"
    elif n.startswith("ANGULO") or n.startswith("ANG "):
        info["grupo"] = "Ángulo"
    elif n.startswith("CUADRADO"):
        info["grupo"] = "Cuadrado"
    elif n.startswith("ESTRELLA"):
        info["grupo"] = "Estrella"
    elif n.startswith("SAFEROCK"):
        info["grupo"] = "Saferock"
    elif n.startswith("HEXAGONO"):
        info["grupo"] = "Hexágono"

    g = info["grupo"]
    if g in ("B Hormigón", "Saferock", "Estrella", "Otros") or n.startswith("BARRA TRASPASO"):
        m = re.search(r"(\d+(?:[.,]\d+)?)\s*MM", n)
        if m:
            info["diametro"] = dec(m.group(1))
    if (g == "Redondo liso" and not n.startswith("BARRA")) or g == "Hexágono":
        m = re.search(r"(?:LISO|HEXAGONO)\s+((?:\d+\s+)?\d+(?:/\d+)?)\s*'", n)
        if m:
            info["pulgadas"] = m.group(1) + '"'
            info["diametro"] = pulgadas_a_mm(m.group(1))
        else:
            m = re.search(r"(\d+(?:[.,]\d+)?)\s*MM", n)
            if m:
                info["diametro"] = dec(m.group(1))
        # 'HEXAGONO SEMI 1 1/8' 8MM' -> 8MM es el largo (8 m)
        n = re.sub(r"(\d)MM\s+SAE", r"\1M SAE", n) if n.startswith("HEXAGONO") else n
    if g == "Cuadrado":
        m = re.search(r"(\d+(?:[.,]\d+)?)\s*MM", n)
        if m:
            info["lado"] = dec(m.group(1))
    if g == "Plana":
        m = re.search(r'(\d+)"X(\d+/\d+)"', n)
        if m:
            info["ancho"] = pulgadas_a_mm(m.group(1))
            info["espesor"] = pulgadas_a_mm(m.group(2))
            info["pulgadas"] = f'{m.group(1)}"x{m.group(2)}"'
        else:
            m = re.search(r"(\d+(?:[.,]\d+)?)\s*X\s*(\d+(?:[.,]\d+)?)\s*MM", n)
            if m:
                info["ancho"], info["espesor"] = dec(m.group(1)), dec(m.group(2))
    if g == "Ángulo":
        m = re.search(r"(\d+)X(\d+)X(\d+(?:[.,]\d+)?)\s*MM", n)
        if m:
            info["ancho"], info["espesor"] = dec(m.group(1)), dec(m.group(3))

    info["largo"] = largo_m(n)
    if info["largo"] is None and g == "Ángulo":
        info["largo"], info["largo_asumido"] = LARGO_ANGULO_DEFAULT, True
    return info


TIPO_A_GRUPO = {"PLANA": "Plana", "ÁNGULO": "Ángulo", "R. LISO": "Redondo liso",
                "CUADRADO": "Cuadrado", "ESTRELLA": "Estrella", "HEXÁGONO": "Hexágono"}


def clave(nombre):
    n = " ".join(str(nombre).upper().replace("''", "'").split()).rstrip(".")
    return n.replace(" SEMI ", " ")


def leer_maestro(ruta):
    if not ruta.exists():
        return {}
    try:
        import openpyxl
    except ImportError:
        print("Aviso: falta openpyxl (pip install openpyxl); se clasifica solo por nombre.")
        return {}
    ws = openpyxl.load_workbook(ruta, read_only=True, data_only=True).active
    filas = ws.iter_rows(values_only=True)
    cab = [str(c).strip() if c else "" for c in next(filas)]
    maestro = {}
    for f in filas:
        r = dict(zip(cab, f))
        if not r.get("Producto"):
            continue
        tipo = str(r.get("Tipo") or "").strip().upper()
        if tipo.startswith("BH"):
            grupo = "B Hormigón"
        elif tipo.startswith("SR"):
            grupo = "Saferock"
        else:
            grupo = TIPO_A_GRUPO.get(tipo)
        maestro.setdefault(clave(r["Producto"]), {
            "grupo": grupo,
            "largo": r.get("CORTE T4") if isinstance(r.get("CORTE T4"), (int, float)) else None,
            "formato": str(r.get("Grupo") or "").strip() or None,
            "codigo": r.get("Código"),
        })
    return maestro


def fmt(v):
    return f"{v:g}".replace(".", ",")


def descripcion(nombre, i):
    """Etiqueta corta: medidas en mm, largo en m."""
    g = i["grupo"]
    partes = []
    if g == "Plana" or g == "Ángulo":
        if i["ancho"] is not None:
            base = f"{fmt(i['ancho'])}×{fmt(i['ancho'])}×{fmt(i['espesor'])} mm" if g == "Ángulo" \
                else f"{fmt(i['ancho'])}×{fmt(i['espesor'])} mm"
            partes.append(base)
    elif g == "Cuadrado" and i["lado"]:
        partes.append(f"{fmt(i['lado'])} mm")
    elif g == "Hexágono" and i["diametro"]:
        partes.append(f"{fmt(i['diametro'])} mm e/c")
    elif i["diametro"]:
        partes.append(f"Ø{fmt(i['diametro'])} mm")
    if i["largo"]:
        partes.append(f"{fmt(i['largo'])} m" + ("*" if i["largo_asumido"] else ""))
    return " · ".join(partes)


def cargar(ruta, ruta_maestro):
    """Devuelve (productos, ordenes, maestro, sin_maestro)."""
    filas = leer_primera_tabla(ruta)
    maestro = leer_maestro(ruta_maestro)
    sin_maestro = set()

    productos, ordenes = {}, []
    for f in filas:
        # los SEMI son el mismo producto: se unifican
        nombre = " ".join(f["Material_prod"].replace(" SEMI ", " ").split()).rstrip(".")
        prod, cons = num(f["qty_prod"]), num(f["qty_cons"])
        if not prod or not cons:
            continue
        ini = datetime.strptime(f["Fecha_inicio"].strip(), "%m/%d/%Y %I:%M:%S %p")
        if nombre not in productos:
            i = clasificar(nombre)
            mm = maestro.get(clave(f["Material_prod"])) or maestro.get(clave(nombre))
            if mm:
                if mm["grupo"]:
                    i["grupo"] = mm["grupo"]
                if mm["largo"]:
                    i["largo"], i["largo_asumido"] = float(mm["largo"]), False
                i["formato"], i["codigo"] = mm["formato"], mm["codigo"]
            elif maestro:
                sin_maestro.add(nombre)
            i["desc"] = descripcion(nombre, i)
            productos[nombre] = i
        ordenes.append({
            "o": f["Orden_prod"].strip(),
            "p": nombre,
            "original": " ".join(f["Material_prod"].split()),
            "d": ini.strftime("%Y-%m-%d"),
            "ini": ini,
            "kp": prod,
            "kc": cons,
            "meta": num(f["Rend_Met"]),
        })
    return productos, ordenes, maestro, sin_maestro


def main():
    ruta = Path(sys.argv[1]) if len(sys.argv) > 1 else CSV_DEFAULT
    ruta_maestro = Path(sys.argv[2]) if len(sys.argv) > 2 else MAESTRO_DEFAULT
    productos, ordenes, maestro, sin_maestro = cargar(ruta, ruta_maestro)

    lista_prod = sorted(productos)
    idx = {p: k for k, p in enumerate(lista_prod)}
    datos = {
        "generado": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "fuente": ruta.name,
        "maestro": ruta_maestro.name if maestro else None,
        "sin_maestro": sorted(sin_maestro),
        "productos": [{"n": p, **productos[p]} for p in lista_prod],
        # orden: [orden, idx_producto, fecha, kg_prod, kg_cons, meta]
        "ordenes": [[o["o"], idx[o["p"]], o["d"], o["kp"], o["kc"], o["meta"]] for o in ordenes],
    }

    html = PLANTILLA.read_text(encoding="utf-8").replace(
        "/*__DATOS__*/null", json.dumps(datos, ensure_ascii=False, separators=(",", ":")))
    SALIDA.parent.mkdir(exist_ok=True)
    SALIDA.write_text(html, encoding="utf-8")

    sobre = sum(1 for o in ordenes if o["kp"] / o["kc"] > 1)
    if sin_maestro:
        print(f"{len(sin_maestro)} productos no están en el maestro (clasificados por nombre):")
        for n in sorted(sin_maestro):
            print("  -", n, "->", productos[n]["grupo"], productos[n]["desc"])
    print(f"{len(ordenes)} órdenes, {len(productos)} productos, {sobre} sobre 100 % -> {SALIDA}")


if __name__ == "__main__":
    main()
