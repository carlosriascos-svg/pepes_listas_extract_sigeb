"""
Pipeline completo: descomprime los .zip de "Archivo/" y luego procesa cada reporte
(PDF + Excel) extraído, dejando todo organizado bajo una carpeta `results/`:

    results/
      reportes/<NOMBRE_REPORTE>/        <- PDF + Excel descomprimidos, más el
                                            scraping de ESE reporte puntual:
                                            *.extraido.json, *.extraido.csv, *.hoja2.csv
      consolidados/
        consolidado_excel.xlsx          <- hoja 2 de todos los Excel concatenada
        consolidado_pdfs.xlsx           <- campos extraídos de todos los PDF concatenados
        errores.log                     <- advertencias/errores de todo el pipeline

Uso:
    python pipeline.py                      # usa "Archivo/" -> "results/"
    python pipeline.py --origen Archivo --resultados results
"""

import argparse
from pathlib import Path

from descomprimir_zips import descomprimir_zips
from procesar_reportes import consolidar_excel, consolidar_pdfs, procesar_carpeta


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Descomprime los reportes en .zip y extrae sus campos (PDF + Excel)."
    )
    parser.add_argument(
        "--origen",
        type=str,
        default="Archivo",
        help="Carpeta con los .zip de reportes (por defecto: 'Archivo')",
    )
    parser.add_argument(
        "--resultados",
        type=str,
        default="results",
        help="Carpeta raíz donde se guarda todo el pipeline (por defecto: 'results')",
    )
    args = parser.parse_args()

    origen = Path(args.origen)
    resultados = Path(args.resultados)
    carpeta_reportes = resultados / "reportes"
    carpeta_consolidados = resultados / "consolidados"

    print(f"=== Paso 1: Descomprimir '{origen}' -> '{carpeta_reportes}' ===")
    carpetas = descomprimir_zips(origen, carpeta_reportes, eliminar_zip=False)

    if not carpetas:
        print("No se descomprimió ningún reporte, se detiene el pipeline.")
        return

    print(f"\n=== Paso 2: Extraer PDF + hoja 2 de Excel de cada reporte ===")
    errores: list[str] = []
    resultados_por_carpeta = []
    for carpeta in sorted(carpetas):
        print(f"Procesando: {carpeta}")
        resultados_por_carpeta.append(procesar_carpeta(carpeta, errores))

    print(f"\n=== Paso 3: Consolidar resultados en '{carpeta_consolidados}' ===")
    n_excel = consolidar_excel(
        resultados_por_carpeta, carpeta_consolidados / "consolidado_excel.xlsx"
    )
    n_pdf = consolidar_pdfs(
        resultados_por_carpeta, carpeta_consolidados / "consolidado_pdfs.xlsx"
    )

    print(f"Consolidado Excel (hoja 2): {n_excel} fila(s) -> {carpeta_consolidados / 'consolidado_excel.xlsx'}")
    print(f"Consolidado PDFs: {n_pdf} fila(s) -> {carpeta_consolidados / 'consolidado_pdfs.xlsx'}")

    if errores:
        carpeta_consolidados.mkdir(parents=True, exist_ok=True)
        errores_path = carpeta_consolidados / "errores.log"
        errores_path.write_text("\n".join(errores), encoding="utf-8")
        print(f"\n{len(errores)} advertencia(s)/error(es). Ver detalle en {errores_path}")
    else:
        print("\nSin errores.")

    print(f"\nRevisa el scraping por reporte en: {carpeta_reportes}")


if __name__ == "__main__":
    main()
