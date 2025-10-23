
# SIAF Dashboard – Perú Compras

Este app en **Streamlit** carga el Excel de SIAF (rango A:CH) y genera resúmenes de PIA, PIM, Certificado, Comprometido, Devengado, Girado, Pagado,
con filtros por UE, programa, función, fuente, genérica, específica y meta. También crea un pivote y una serie mensual del devengado.

## Cómo ejecutarlo (local)

1) Instala las dependencias (ideal en un entorno virtual):
```
pip install -r requirements.txt
```

2) Ejecuta el servidor de Streamlit:
```
streamlit run siaf_dashboard.py
```

3) Abre el navegador en la URL que te muestre Streamlit (por defecto http://localhost:8501).

4) Coloca el Excel SIAF en la carpeta `data/siaf/` del repositorio (puedes versionarla o sincronizarla). Al iniciar el dashboard tomará automáticamente el archivo `.xlsx` más reciente y, si hay varios, podrás elegirlo desde la barra lateral. Por defecto, el app intentará:
   - Detectar la **hoja** con los datos (busca columnas como `ano_eje`, `mto_pim`).
   - Leer el rango **A:CH**.
   - Detectar automáticamente la fila de **encabezados** (puedes desactivar y fijarlo, p. ej. fila 4).
   
## Notas

- El app suma los **mto_devenga_01..12** para calcular el **devengado (YTD)**. También calcula el **saldo PIM** y el **avance %** (= devengado/PIM).
- Puedes agrupar por **generica, especifica, fuente_financ, unidad_ejecutora, función, programa**, etc.
- Descarga un Excel con el **resumen** (pivot) y los **datos filtrados**.

## ¿Cómo obtengo el código completo para copiarlo manualmente?

Desde la raíz del proyecto o directamente en la interfaz puedes elegir entre estas alternativas:

1. **Guardar una copia en un archivo nuevo**

   ```bash
   python export_siaf_dashboard.py -o ruta/destino/siaf_dashboard.py
   ```

   Sustituye `ruta/destino/` por la carpeta donde quieres recibir la copia. Se creará automáticamente si no existe.

2. **Mostrar todo el contenido en pantalla para copiarlo**

   ```bash
   python export_siaf_dashboard.py --stdout
   ```

   También puedes redirigir la salida a un archivo con `>` si te resulta más cómodo:

   ```bash
   python export_siaf_dashboard.py --stdout > siaf_dashboard.py
   ```

Estas opciones reproducen los comandos compartidos previamente y permiten obtener el código manualmente cuando lo necesites.

