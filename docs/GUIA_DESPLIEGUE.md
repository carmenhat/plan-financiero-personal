# Guía de despliegue

Montarlo desde cero: primero con los datos de ejemplo, para ver que todo funciona, y después con los tuyos.

Hace falta una cuenta de Google. Python, LibreOffice y Node solo si vas a regenerar el sheet o pasar las pruebas.

---

## 1. El sheet

1. Descarga `plantilla/Plan_financiero_demo.xlsx`.
2. En Google Drive: *Nuevo → Subir archivo*. Ábrelo con Google Sheets y *Archivo → Guardar como hoja de cálculo de Google*. Puedes borrar el `.xlsx` subido.
3. Ponle el nombre que quieras y déjalo en la carpeta donde vayas a guardar tus extractos.
4. *Archivo → Configuración*: configuración regional **España** y tu zona horaria (por ejemplo, `(GMT+01:00) Madrid`).

Ya se ve todo con los datos ficticios. Lo que falta es el importador.

## 2. El importador

1. En el sheet: *Extensiones → Apps Script*.
2. Borra lo que haya en `Código.gs` y pega `apps-script/importar.js`. Guarda.
3. *Servicios* (el `+` de la izquierda) → **Drive API** → versión v3, identificador `Drive` → *Añadir*. Sin esto no se pueden leer los `.xls` binarios (CaixaBank, ING).
4. *Configuración del proyecto* (el engranaje) → **Zona horaria**: la misma que pusiste en el sheet. Si no coinciden, los meses no casan y nada avisa.
   Si prefieres hacerlo de una vez: marca *Mostrar el archivo de manifiesto «appsscript.json»* y pega `apps-script/appsscript.json`, que ya trae la Drive API y la zona `Europe/Madrid`.
5. Vuelve al sheet y recárgalo. Aparece el menú **Plan financiero**.

La primera vez que uses el menú, Google pide permiso para el sheet y para Drive (la carpeta `Movimientos` y la conversión de los `.xls`).

## 3. Probarlo con los extractos de ejemplo

Los datos del sheet de ejemplo salen de los mismos movimientos que `extractos-ejemplo/`, así que importarlos es la prueba de que no duplica:

1. *Plan financiero → Importar movimientos*. La primera vez crea la carpeta `Movimientos` junto al sheet y dice que no hay archivos.
2. Sube a esa carpeta los cuatro archivos de `extractos-ejemplo/`.
3. *Plan financiero → Importar movimientos* otra vez.

Debería salir, para cada archivo, **0 nuevos** y todos «ya estaban» (las pruebas de `test/` lo comprueban fuera de Google con los mismos archivos). Los archivos pasan a `Movimientos/Procesados`.

## 4. Pasar a tus datos

1. **Borra lo ficticio:**
   - MOVIMIENTOS: todas las filas desde la 4.
   - FLEXIBLE: las filas de ejemplo (deja vacías las celdas amarillas).
   - HISTORIAL: las filas de ejemplo.
   - VIAJES: los viajes de ejemplo.
   - MIS DATOS conserva los saldos de ejemplo hasta tu primera importación, que los reescribe. Mientras, ingresos y gastos salen a cero.
2. **Tus cuentas.** El plan lleva tres cuentas con saldo (MIS DATOS, columnas B, C y D) y una tarjeta. Si las tuyas no son CaixaBank, Openbank e ING, cambia los nombres en `CUENTAS_SALDO` y en `cuenta` de cada entrada de `BANCOS`, en el importador. Los títulos de MIS DATOS y del panel solo son texto: cámbialos a mano o regenera el sheet (apartado 7).
3. **Tus supuestos:** PARÁMETROS (edad, jubilación, pensión, rentabilidades), VIVIENDA, VIAJES (aportación y primer mes) y, si no tienes retribución flexible, deja FLEXIBLE vacía.
4. **Tus reglas y categorías:** REGLAS trae ejemplos que no casarán con tus conceptos. Mejor empezar con pocas e ir añadiendo según lo que salga por revisar.
5. Descarga tus extractos, déjalos en `Movimientos` e **importa**. Después, revisa lo amarillo de MOVIMIENTOS.

⚠️ Tus extractos no van nunca al repositorio. Si trabajas con un clon, guárdalos fuera.

## 5. Si tu banco no está

El importador busca las cabeceras de la tabla. Abre un extracto tuyo, mira cómo se llaman las columnas y añade una entrada a `BANCOS` (en [`COMO_FUNCIONA.md`](COMO_FUNCIONA.md) está el formato). Para probarla sin Google, añade un caso a `test/test-importar.js` con unas filas inventadas con esas cabeceras y pasa:

```
node test/test-importar.js
```

Si el extracto es un CSV, el importador detecta solo si va separado por «;» o por «,».

## 6. El dashboard en directo

Necesita Claude con el conector de **Google Drive**.

1. Copia el `fileId` de tu sheet: lo que va entre `/d/` y `/edit` en su dirección.
2. En `dashboard/datasets.json`, cambia `TU_FILE_ID` por ese valor en los nueve datasets.
3. Pídele a Claude que cree un dashboard del tipo Dashboard con `index.html`, `page.js` y `loaders.js` como archivos de la página, y los datasets de `datasets.json`.

Cada vez que lo abras, descarga el sheet con tu conexión y lee PANEL_DATOS. Quien abra el enlace sin acceso a tu sheet no ve nada.

## 7. Regenerar el sheet (opcional)

Solo hace falta si cambias la estructura o quieres empezar con tus cuentas y tus supuestos ya puestos. Edita `generador/datos-demo/config.json` y:

```
pip install openpyxl xlwt
python generador/datos_demo.py
python generador/build.py
soffice --headless --calc --convert-to xlsx --outdir plantilla/recalculado plantilla/Plan_financiero_demo.xlsx
python generador/exportar_demo.py
```

- `datos_demo.py` inventa los movimientos y los extractos de ejemplo a partir de `config.json`.
- `build.py` escribe el `.xlsx` con todas las fórmulas, pero sin valores calculados.
- La conversión con LibreOffice lo recalcula. Google también recalcula al subirlo, así que este paso solo hace falta para revisar errores o para `exportar_demo.py`, que lee los valores (apunta `XLSX` al archivo recalculado si no sustituyes el original).

Para cambiar un sheet que ya usas sin perder su ID ni su script: en el sheet, *Archivo → Importar → Subir* el `.xlsx` nuevo y elige **Sustituir hoja de cálculo**. Ojo: sustituye también tus movimientos. Para un sheet en uso, mejor un script de una sola vez que añada lo nuevo sin tocar lo demás (ver [`APRENDIZAJES.md`](APRENDIZAJES.md)).

Última actualización: 12/10/2026
