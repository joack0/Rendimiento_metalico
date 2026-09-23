"""Prepara el dashboard de palanquillas fuera de grado.

Uso:
  python3 scripts/preparar_fuera_grado.py [datos/palanquillas_fuera_grado.tsv]

Lee el registro manual (Fecha, Colada, Largo, Cantidad, Producto, Calidad),
normaliza los nombres de producto escritos a mano en grupo + medida y genera
dashboard/fuera_grado.html a partir de scripts/plantilla_fuera_grado.html.
"""
import csv
import json
import re
import sys
from datetime import date
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
ENTRADA = Path(sys.argv[1]) if len(sys.argv) > 1 else RAIZ / "datos" / "palanquillas_fuera_grado.tsv"
PLANTILLA = RAIZ / "scripts" / "plantilla_fuera_grado.html"
SALIDA = RAIZ / "dashboard" / "fuera_grado.html"

# Fracciones de pulgada tal como se nombran las planas en mm (1 1/4" = 32, 3/16" = 4,85 ...)
PULG_PLANA = {"1": 25, "1 1/4": 32, "3/16": 4.85, "1/8": 3, "1/4": 6}


def num(s):
    return float(s.replace(",", "."))


def fmt(x):
    return f"{x:g}".replace(".", ",")


def pulgadas(ent, frac):
    """('1', '1/2') -> '1 1/2"'"""
    partes = [p for p in (ent, frac) if p]
    return " ".join(partes) + '"'


def pulg_a_mm(txt):
    total = 0.0
    for p in txt.replace('"', "").split():
        if "/" in p:
            a, b = p.split("/")
            total += int(a) / int(b)
        else:
            total += int(p)
    return round(total * 25.4, 1)


def grupo_de(n):
    if re.match(r"^(HEX|HES|EXAG)", n):
        return "Hexágono"
    if re.match(r"^CU(A|D)", n):
        return "Cuadrado"
    if re.match(r"^(SAF|SF)", n):
        return "Saferock"
    if n.startswith("ESTRELLA"):
        return "Estrella"
    if n.startswith("HORM") or n.startswith("REND"):
        return "B Hormigón"
    if re.match(r"^(ANG|AGU)", n):
        return "Ángulo"
    if re.match(r"^(PL|PAL|PÑ|PANO)", n):
        return "Plana"
    if re.match(r"^(RED|RDC|RCR|R/C|C/R|CR\b)", n) and re.search(r"C/R|R/C|^RCR|^CR\b", n):
        return "B Hormigón"
    if re.match(r"^(RED|R\s*\.?\s*LISO|LIS|B TRASPASO|BARRA TRASPASO)", n):
        return "Redondo liso"
    return None


def medida_plana_angulo(n, grupo):
    # Planas en pulgadas: 1 1/4 x 3/16, 1X1/4X3/16, 1/1/4X3/16, 1.1/4X3/16, 11/4X1/8, 1/14X1/1/8, 1X3/16
    n = n.replace("1X1/4X", "1 1/4X")  # 'PLANA 1X1/4X3/16'
    m = re.search(r"(\d[\d/., ]*)X\s*(\d?/?\d+/\d+)", n)
    if grupo == "Plana" and m and "/" in m.group(1) + m.group(2):
        a, b = m.group(1).strip(), m.group(2)
        ancho = "1" if a == "1" else "1 1/4"
        esp = b.split("/", 1)[1] if b.count("/") == 2 else b  # 1/1/8 -> 1/8
        return f"{fmt(PULG_PLANA[ancho])}x{fmt(PULG_PLANA[esp])}"
    m = re.search(r"(\d+(?:[.,]\d+)?)\s*X\s*(\d+(?:[.,]\d+)?)", n)
    if not m:
        return None
    a, b = m.group(1), m.group(2)
    if grupo == "Plana" and re.fullmatch(r"[24]85", b):  # 32X485 -> 32x4,85
        b = b[0] + ",85"
    return f"{fmt(num(a))}x{fmt(num(b))}"


def medida_redonda(n, grupo):
    """Devuelve (etiqueta, mm) para hormigón, saferock, liso, cuadrado y hexágono."""
    t = n.replace("C/R", " ").replace("R/C", " ")
    prefijo = re.compile(r"^[\s.,=/]*(REDONDO|REDOLISO|REDLISO|RED|RDC|RD|RCR|CR|R\s*\.?\s*LISO|LISO|LIS0|HORMIGON|HORMIG|HORM|REND|"
                         r"SAFERROK|SAFERRO|SAFEROCK|SAFF|SAF|SFR|SF|CUADRADO|CUDRADO|CUAD|HEXAGONO|HESAGONO|EXAGONO|"
                         r"B TRASPASO|BARRA TRASPASO|L)")
    while True:  # 'RED LISO 12', 'RED,LISO 19M/M', 'REDONDOLISO 12M/M'
        t2 = prefijo.sub("", t, count=1)
        if t2 == t:
            break
        t = t2
    t = re.sub(r"\b(1020|ASTM|A520D|SOLDA|SOL|SL|COM|SAE)\b", " ", t)
    t = re.sub(r"(FG|COM)$", "", t)
    t = t.strip(" .,=/")
    # mm explícitos: 25MM, 25 M/M, 28/M, 18M, 25,4
    if re.fullmatch(r"25[,.]4.*", t):
        return '1"', 25.4
    m = re.match(r"^(\d+)\s*(MM|M/M|/M|M\b|/MM|/M/M)", t)
    if m:
        return f"{m.group(1)} mm", float(m.group(1))
    tt = t.replace(" ", "").replace(",", "/").replace(".", "/").replace('"', "")
    m = re.fullmatch(r"1/?(\d)/(\d)", tt)  # 1/3/4, 13/4, 1/1/2, 1,1/2
    if m:
        txt = pulgadas("1", f"{m.group(1)}/{m.group(2)}")
        return txt, pulg_a_mm(txt)
    if tt == "1/14":  # typo de 1 1/4
        return '1 1/4"', pulg_a_mm("1 1/4")
    m = re.fullmatch(r"(\d)/(\d)", tt)
    if m:
        txt = pulgadas(None, tt)
        return txt, pulg_a_mm(txt)
    if tt == "1" and grupo in ("Redondo liso", "Hexágono"):
        return '1"', 25.4
    m = re.match(r"^(\d+)", tt)
    if m and int(m.group(1)) >= 8:
        return f"{m.group(1)} mm", float(m.group(1))
    return None, None


def leer():
    filas = []
    with open(ENTRADA, encoding="utf-8") as f:
        for r in csv.DictReader(f, delimiter="\t"):
            filas.append({k.strip(): (v or "").strip() for k, v in r.items()})
    return filas


def procesar(filas):
    regs, prev_fecha, prev_grupo = [], None, None
    for i, r in enumerate(filas):
        obs = []
        d, mth, y = (int(x) for x in r["Fecha"].split("-"))
        f = date(y, mth, d)
        if prev_fecha and (prev_fecha - f).days > 200:
            f = date(prev_fecha.year, mth, d)
            obs.append(f"Fecha corregida (año {y} → {f.year})")
        prev_fecha = f

        largo = None
        if r["Largo"]:
            largo = num(r["Largo"])
            if largo > 10:
                obs.append(f"Largo corregido ({r['Largo']} → {fmt(largo / 10)})")
                largo = largo / 10
        else:
            obs.append("Sin largo")

        cant = int(r["Cantidad"]) if r["Cantidad"] else None
        if cant is None:
            obs.append("Sin cantidad")

        orig = r["Producto"]
        n = " ".join(orig.upper().split())
        grupo = grupo_de(n) if n else None
        if grupo is None and re.match(r"^\d+(?:[.,]\d+)?\s*X", n):
            grupo = prev_grupo
            obs.append("Grupo deducido de la fila anterior")
        medida = None
        if grupo in ("Plana", "Ángulo"):
            medida = medida_plana_angulo(n, grupo)
        elif grupo:
            medida, _ = medida_redonda(n, grupo)
        if grupo is None:
            grupo = "Sin identificar" if not n else "Otros"
            obs.append("Producto vacío" if not n else "Producto fuera de las familias")
        if grupo in ("Plana", "Ángulo", "B Hormigón", "Saferock", "Redondo liso", "Cuadrado", "Hexágono"):
            prev_grupo = grupo
            if medida is None:
                obs.append("Sin medida")
        prod = f"{grupo} {medida}" if medida else (grupo if grupo not in ("Otros",) else orig.upper())

        regs.append({
            "i": i + 2, "f": f.isoformat(), "col": r["Colada"], "l": largo, "c": cant,
            "orig": orig, "g": grupo, "m": medida, "p": prod, "cal": r["Calidad"], "obs": obs,
        })
    return regs


def main():
    regs = procesar(leer())
    html = PLANTILLA.read_text(encoding="utf-8").replace("/*__DATOS__*/[]", json.dumps(regs, ensure_ascii=False))
    SALIDA.write_text(html, encoding="utf-8")
    from collections import Counter
    print(len(regs), "registros ->", SALIDA)
    print(Counter(r["g"] for r in regs))
    for r in regs:
        if r["obs"]:
            print(" ", r["i"], r["f"], repr(r["orig"]), "->", r["p"], "|", "; ".join(r["obs"]))


if __name__ == "__main__":
    main()
