# Pepes Listas Extract SIGEB

Extrae información de las declaraciones de bienes/rentas y conflictos de interés
(PDF + Excel) que llegan comprimidas en `Archivo/`.

## Configuración

1. Crear y activar el entorno virtual:

   ```powershell
   python -m venv .venv
   .\.venv\Scripts\Activate.ps1
   ```

2. Instalar dependencias:

   ```powershell
   pip install -r requirements.txt
   ```

## Ejecución

Con los `.zip` de reportes dentro de `Archivo/`, correr el pipeline completo:

```powershell
python pipeline.py
```

Esto hace, en orden:

1. Descomprime cada `.zip` de `Archivo/`.
2. Por cada reporte, extrae los campos del PDF y la segunda hoja del Excel.
3. Genera los consolidados finales.

Todo queda organizado en `results/`:

```
results/
  reportes/<NOMBRE_REPORTE>/   # PDF + Excel descomprimidos, y su scraping individual
  consolidados/
    consolidado_excel.xlsx     # hoja 2 de todos los Excel, concatenada
    consolidado_pdfs.xlsx      # campos extraídos de todos los PDF, concatenados
    errores.log                # solo si hubo advertencias/errores
```

### Parámetros opcionales

```powershell
python pipeline.py --origen Archivo --resultados results
```

### Ejecutar los pasos por separado

```powershell
python descomprimir_zips.py "Archivo/"
python procesar_reportes.py "Archivo/" --output results
```

## Notas

- `results/`, así como los `.xlsx`, `.json` y `.csv` generados, están en `.gitignore`
  (son datos, no código).
- La extracción del PDF asume el layout del formulario oficial (Ley 2013 de 2019);
  si el formato cambia, algún campo puntual podría quedar vacío sin detener el proceso.
