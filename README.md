# plan-financiero-personal

Un plan financiero personal en Google Sheets que se alimenta de los extractos de tus bancos: cuánto tienes, cuánto ahorras de verdad, en qué se va, si llegas a los viajes que tienes pensados y qué patrimonio tendrás al jubilarte. Con un dashboard que lo lee.

> **Todo lo que hay aquí es ficticio.** Los movimientos, los saldos, los viajes y la vivienda son inventados para que se pueda probar sin datos de nadie.
>
> **No es asesoramiento financiero ni fiscal.** Las rentabilidades, la pensión y los impuestos son supuestos que tienes que cambiar por los tuyos.

## Cómo está montado

| Pieza | Dónde manda | Qué hace |
|---|---|---|
| El sheet | Google Sheets, en tu Drive | Los datos y todas las cuentas: 16 pestañas. Es la fuente de verdad |
| El importador | Apps Script ligado al sheet (`apps-script/`) | Lee los extractos de los bancos, los categoriza con tus reglas y los pasa a MOVIMIENTOS sin duplicar |
| La carpeta `Movimientos` | Junto al sheet, en Drive | Donde dejas los extractos. El importador la crea si no existe y mueve lo importado a `Movimientos/Procesados` |
| El dashboard | Un artifact de tipo Dashboard en Claude (`dashboard/`) | Lee la pestaña PANEL_DATOS del sheet y la enseña en gráficos |
| El generador | Python, en tu ordenador (`generador/`) | Construye el sheet desde cero con los datos de ejemplo. Solo hace falta para cambiar la estructura |

**Demo del dashboard con los datos ficticios:** https://claude.ai/artifact/9iReaYTAhHryHZHxVPR8gS

## Qué hay aquí

| Carpeta | Qué hace | Estado |
|---|---|---|
| `plantilla/` | `Plan_financiero_demo.xlsx`: el sheet ya generado y recalculado, con los datos de ejemplo. Se sube a Google tal cual | ✅ |
| `apps-script/` | `importar.js` (el importador y el menú *Plan financiero*) y `appsscript.json` (el manifiesto, con la Drive API activada) | ✅ |
| `extractos-ejemplo/` | Cuatro extractos ficticios, uno por formato: Openbank cuenta y tarjeta (`.xls` que por dentro es HTML), CaixaBank e ING (`.xls` binario) | ✅ |
| `generador/` | `datos_demo.py` (inventa los movimientos y los extractos), `build.py` (genera el sheet), `exportar_demo.py` (saca los datos del dashboard demo) y `datos-demo/` | ✅ |
| `dashboard/` | El código del dashboard: `index.html`, `page.js`, `loaders.js`, `lector-xlsx.js`, `datasets.json`. En `demo/`, los datos que lleva dentro la demo | ✅ |
| `test/` | `test-importar.js`: las pruebas del importador contra los extractos de ejemplo | ✅ |
| `docs/` | Cómo funciona, cómo montarlo, el día a día y lo que costó | ✅ |

Leyenda: ✅ funciona · 🧪 en pruebas · ⏸️ pausado · 🗄️ de una sola vez · 💀 muerto

## Stack

- **Google Sheets + Apps Script** (V8). Servicio avanzado **Drive API v3**, para convertir los `.xls` binarios.
- **Python 3** con `openpyxl` y `xlwt`, solo para regenerar el sheet o los extractos.
- **LibreOffice**, solo para recalcular el `.xlsx` en tu ordenador.
- **Node**, solo para las pruebas. Sin dependencias ni `package.json`.
- **Claude**, para el dashboard (tipo Dashboard, con el conector de Google Drive).

Sin coste: nada de esto necesita un plan de pago, salvo la cuenta de Claude para el dashboard.

## Variables de entorno

Ninguna. El importador no guarda IDs: busca la carpeta `Movimientos` junto al propio sheet. El único identificador es el `fileId` de tu sheet en `dashboard/datasets.json`, y solo si montas el dashboard en directo.

## Comandos

```
node test/test-importar.js            # pruebas del importador
python generador/datos_demo.py        # rehace los datos y los extractos de ejemplo
python generador/build.py             # rehace plantilla/Plan_financiero_demo.xlsx
python generador/exportar_demo.py     # rehace dashboard/demo/*.json (después de recalcular)
```

El recálculo y la subida a Google están en [`docs/GUIA_DESPLIEGUE.md`](docs/GUIA_DESPLIEGUE.md).

## Documentación completa

| Archivo | Qué contiene |
|---|---|
| [`docs/COMO_FUNCIONA.md`](docs/COMO_FUNCIONA.md) | Las pestañas, el importador, cómo no duplica, los saldos, las categorías, la vivienda, la jubilación y el dashboard |
| [`docs/GUIA_DESPLIEGUE.md`](docs/GUIA_DESPLIEGUE.md) | Montarlo desde cero con tus datos: el sheet, el importador, tus bancos y el dashboard |
| [`docs/OPERATIVA.md`](docs/OPERATIVA.md) | El día a día: cada mes, cómo se cambia una cifra, una regla o la estructura, y qué mirar si algo falla |
| [`docs/APRENDIZAJES.md`](docs/APRENDIZAJES.md) | Los fallos que ya pasaron y lo que se descartó. Leer primero si algo no cuadra |

## Lo que está roto y ya se sabe

| Qué | Severidad |
|---|---|
| Un movimiento de tarjeta que cambie de importe al liquidarse entraría dos veces: la clave lleva el importe. | Baja. Si la tarjeta no cuadra, es lo primero que mirar |
| La estructura de los extractos de ejemplo es aproximada. Las cabeceras son las que busca el importador; las filas de título y las columnas que no usa pueden no ser idénticas a las de cada banco | Baja. El importador busca las cabeceras en las 40 primeras filas |
| Las fórmulas de «mes completo» dependen de la fecha de hoy. Con los datos de ejemplo, que acaban en septiembre de 2026, todos los meses cuentan como completos a partir de octubre de 2026 | Ninguna con datos reales |

## Decisiones cerradas — no volver a proponer

| Decisión | Motivo |
|---|---|
| El banco se reconoce por las cabeceras de la tabla, no por el nombre del archivo | Los nombres cambian con cada descarga; las cabeceras, casi nunca |
| La clave de un movimiento no lleva el concepto | Es lo que más cambia entre formatos del mismo banco. Lleva cuenta, fecha, importe y saldo (en la tarjeta, la hora) |
| Los saldos de fin de mes los da el extracto, no se calculan sumando movimientos | Sumar arrastra cualquier hueco del histórico. El saldo del banco es el dato |
| La categoría decide si un movimiento es ingreso, gasto o traspaso. El signo no | Un dinero que te devuelven de una compra resta del gasto; un regalo que te hacen es ingreso. El signo no distingue los dos |
| Las fórmulas que escribe un script se escriben con la hoja en inglés un instante | Una hoja en español separa los argumentos con «;»: una fórmula con comas da `#ERROR!` |
| Las fechas pasan entre script y hoja como texto `aaaa-mm-dd` | La zona horaria del proyecto de Apps Script y la de la hoja son independientes |
| El dashboard lee el `.xlsx` entero y busca PANEL_DATOS | La exportación CSV de Drive solo da la primera pestaña, y con números en formato español |
| El dashboard en directo no guarda cifras; la demo sí | Con datos reales, una copia dentro de la página la vería cualquiera con el enlace. Con datos ficticios, llevarlos dentro es lo que hace que la demo funcione para cualquiera |
| Los cambios de estructura del sheet se hacen con un script de una sola vez, no regenerándolo | Regenerar pierde lo que has tocado a mano. `build.py` es para empezar, no para mantener |

## Lo que se me olvida

- **Probar el lector del dashboard contra la exportación de Google**, no contra LibreOffice. Google cierra las filas vacías en sí mismas (`<row r="38"/>`) y LibreOffice no.
- **Un loader del tipo Dashboard solo puede usar lo que lleva dentro.** Por eso `lector-xlsx.js` va copiado en cada loader de `loaders.js`: si cambias el lector, cámbialo en los nueve.
- Si cambias el nombre de una cuenta, hay que cambiarlo en tres sitios: `CUENTAS_SALDO` y `BANCOS` en `importar.js`, y `cuentas` en `generador/datos-demo/config.json`.
- Una tabla nueva en PANEL_DATOS va **al final**.
- La zona horaria del proyecto de Apps Script (en `appsscript.json`) y la de la hoja (*Archivo → Configuración*) se configuran por separado.

## Pendiente

| Qué | Por qué está pendiente |
|---|---|
| Licencia | Sin decidir. Mientras no haya, el código se puede leer pero no reutilizar legalmente |
| Más formatos de banco | Solo están los formatos que tienen extracto de ejemplo. Para añadir el tuyo, ver [`docs/GUIA_DESPLIEGUE.md`](docs/GUIA_DESPLIEGUE.md) |

Última actualización: 12/10/2026
