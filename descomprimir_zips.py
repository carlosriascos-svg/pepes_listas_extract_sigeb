"""
Descomprime todos los archivos .zip encontrados dentro de un directorio (incluyendo subcarpetas).

Uso:
    python descomprimir_zips.py "Archivo/"
    python descomprimir_zips.py "Archivo/" --destino "Salida/"
    python descomprimir_zips.py "Archivo/" --eliminar-zip
"""

from __future__ import annotations

import argparse
import zipfile
from pathlib import Path


def descomprimir_zips(
    origen: Path, destino: Path | None, eliminar_zip: bool
) -> list[Path]:
    """Descomprime los .zip encontrados en `origen`. Devuelve la lista de carpetas
    destino generadas exitosamente (útil para encadenar con otro proceso)."""
    if not origen.exists() or not origen.is_dir():
        raise NotADirectoryError(f"La ruta '{origen}' no existe o no es un directorio.")

    zips = sorted(
        p
        for p in origen.rglob("*.zip")
        if "__MACOSX" not in p.parts and not p.name.startswith("._")
    )

    if not zips:
        print(f"No se encontraron archivos .zip en '{origen}'.")
        return []

    print(f"Se encontraron {len(zips)} archivo(s) .zip.")

    exitosos = 0
    fallidos = 0
    carpetas_generadas: list[Path] = []

    for zip_path in zips:
        carpeta_destino = (destino if destino else zip_path.parent) / zip_path.stem
        try:
            with zipfile.ZipFile(zip_path, "r") as zf:
                carpeta_destino.mkdir(parents=True, exist_ok=True)
                zf.extractall(carpeta_destino)
            print(f"[OK] {zip_path} -> {carpeta_destino}")
            exitosos += 1
            carpetas_generadas.append(carpeta_destino)

            if eliminar_zip:
                zip_path.unlink()
        except zipfile.BadZipFile:
            print(f"[ERROR] Archivo corrupto o inválido: {zip_path}")
            fallidos += 1
        except Exception as e:
            print(f"[ERROR] {zip_path}: {e}")
            fallidos += 1

    print(f"\nResumen: {exitosos} descomprimido(s), {fallidos} fallido(s).")
    return carpetas_generadas


def main() -> None:
    parser = argparse.ArgumentParser(description="Descomprime todos los .zip dentro de una carpeta.")
    parser.add_argument("origen", type=str, help="Carpeta donde buscar los archivos .zip")
    parser.add_argument(
        "--destino",
        type=str,
        default=None,
        help="Carpeta base donde extraer el contenido (por defecto, junto a cada .zip)",
    )
    parser.add_argument(
        "--eliminar-zip",
        action="store_true",
        help="Elimina el .zip original después de descomprimirlo exitosamente",
    )

    args = parser.parse_args()

    origen = Path(args.origen)
    destino = Path(args.destino) if args.destino else None

    descomprimir_zips(origen, destino, args.eliminar_zip)


if __name__ == "__main__":
    main()
