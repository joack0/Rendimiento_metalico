"""Genera dashboard/timeline.html a partir de los CSV de interrupciones (Hecho causa acción).

Uso:
    python3 scripts/preparar_timeline.py [archivo_o_carpeta ...]

Sin argumentos toma todos los .csv de datos/interrupciones/ (uno por mes, o exportaciones
sucesivas: si dos archivos cubren las mismas fechas, manda el más nuevo).

El HTML resultante es un archivo único con los CSV incrustados. Además, desde el navegador:
- "Cargar CSV…" o arrastrar archivos: se agregan y se recuerdan en ese navegador.
- "Vincular CSV…" (Chrome/Edge): elige el CSV una vez y la página lo relee al abrirse y cada
  minuto; basta con sobrescribir ese archivo con la nueva exportación.

Reglas:
- El CSV exportado trae tablas de resumen arriba; se usa la tabla que empieza en Fecha_Inicio.
- Cada declaración ocupa [Fecha_Inicio, Fecha_Inicio + DuracionMinutos]. Si cruza la medianoche
  se reparte entre ambos días.
- Tiempo sin declaración = tren disponible / en producción.
- Tipo: Parada programada, Parada padrón y Libre según Disciplina; Retraso si Zona, Equipo u
  Ocurrencia dicen "retraso"; el resto es Interrupción/falla.
- Lo que no cubre ningún archivo (antes de su primer día o después de su última declaración) se muestra como "Sin registro".
"""
import json
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
CARPETA_DEFAULT = RAIZ / "datos" / "interrupciones"
PLANTILLA = RAIZ / "scripts" / "plantilla_timeline.html"
SALIDA = RAIZ / "dashboard" / "timeline.html"


def archivos(args):
    rutas = [Path(a) for a in args] or [CARPETA_DEFAULT]
    for r in rutas:
        yield from (sorted(r.glob("*.csv")) if r.is_dir() else [r])


def main():
    lotes = []
    for ruta in archivos(sys.argv[1:]):
        texto = ruta.read_text(encoding="utf-8-sig")
        if "Fecha_Inicio" not in texto:
            print(f"  se omite {ruta.name}: no tiene la columna Fecha_Inicio")
            continue
        # la fecha de modificación decide qué archivo manda si se solapan
        lotes.append({"nombre": ruta.name, "cargado": int(ruta.stat().st_mtime * 1000), "texto": texto})
    if not lotes:
        sys.exit("No se encontraron CSV de interrupciones")
    # </ dentro de un <script> cerraría la etiqueta
    datos = json.dumps(lotes, ensure_ascii=False).replace("</", "<\\/")
    html = PLANTILLA.read_text(encoding="utf-8").replace("__CSV__", datos)
    SALIDA.write_text(html, encoding="utf-8")
    print(f"{SALIDA.relative_to(RAIZ)} generado con {', '.join(l['nombre'] for l in lotes)}")


if __name__ == "__main__":
    main()
