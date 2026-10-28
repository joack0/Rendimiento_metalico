"""Genera dashboard/timeline.html a partir del CSV de interrupciones (Hecho causa acción).

Uso:
    python3 scripts/preparar_timeline.py [ruta_csv]

El HTML resultante es un archivo único: trae el CSV incrustado y además permite cargar
otro CSV desde el navegador (botón "Cargar CSV…" o arrastrando el archivo).

Reglas:
- El CSV exportado trae tablas de resumen arriba; se usa la tabla que empieza en Fecha_Inicio.
- Cada declaración ocupa [Fecha_Inicio, Fecha_Inicio + DuracionMinutos]. Si cruza la medianoche
  se reparte entre ambos días.
- Tiempo sin declaración = tren disponible / en producción.
- Tipo: Parada programada, Parada padrón y Libre según Disciplina; Retraso si Zona, Equipo u
  Ocurrencia dicen "retraso"; el resto es Interrupción/falla.
- Después del fin de la última declaración del archivo el tiempo se muestra como "Sin registro".
"""
import json
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
CSV_DEFAULT = RAIZ / "datos" / "Hecho_causa_accion.csv"
PLANTILLA = RAIZ / "scripts" / "plantilla_timeline.html"
SALIDA = RAIZ / "dashboard" / "timeline.html"


def main():
    ruta = Path(sys.argv[1]) if len(sys.argv) > 1 else CSV_DEFAULT
    texto = ruta.read_text(encoding="utf-8-sig")
    if "Fecha_Inicio" not in texto:
        sys.exit(f"{ruta}: no se encontró la columna Fecha_Inicio")
    # </ dentro de un <script> cerraría la etiqueta
    datos = json.dumps(texto, ensure_ascii=False).replace("</", "<\\/")
    html = (
        PLANTILLA.read_text(encoding="utf-8")
        .replace("__CSV__", datos)
        .replace("__NOMBRE__", json.dumps(ruta.name, ensure_ascii=False))
    )
    SALIDA.write_text(html, encoding="utf-8")
    print(f"{SALIDA.relative_to(RAIZ)} generado con {ruta.name}")


if __name__ == "__main__":
    main()
