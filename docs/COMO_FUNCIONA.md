# Cómo funciona

Qué hace cada pieza y por qué está hecha así. El día a día está en [`OPERATIVA.md`](OPERATIVA.md); cómo montarlo, en [`GUIA_DESPLIEGUE.md`](GUIA_DESPLIEGUE.md).

---

## Qué resuelve

Saber, con los movimientos reales de tus bancos y no con cifras a ojo:

- cuánto tienes y cuánto ahorras de verdad al mes;
- en qué se te va el dinero, y en concreto los gastos de la vivienda;
- si lo que apartas para viajes llega para los viajes que tienes pensados;
- qué patrimonio tendrás al jubilarte y hasta cuándo te dura después;
- qué cambia si tomas una decisión: comprar vivienda ahora o dentro de unos años, invertir parte del efectivo, recortar gastos, apartar más o menos para viajes.

## Las pestañas

| Pestaña | Qué es | Quién la toca |
|---|---|---|
| LEEME | Cómo se usa | Nadie |
| MI FUTURO | El panel dentro del sheet: nivel, patrimonio, jubilación, hitos, medallas de ahorro y gráficos | Nadie, son fórmulas |
| MIS DATOS | Saldo de cada cuenta a fin de mes, ingresos, gastos y ahorro | El importador (saldos) y fórmulas |
| MOVIMIENTOS | Cada movimiento de los bancos con su categoría | El importador; tú solo Categoría y Nota |
| FLEXIBLE | La retribución flexible de la nómina, si la tienes | Tú, una fila al mes |
| CASA | Gastos de la vivienda por concepto y mes | Nadie |
| DECISIONES | Las palancas del «¿y si…?» | Tú |
| HISTORIAL | Registro con fecha de lo que cambias del plan | Tú |
| VIVIENDA | ¿Comprar ahora o esperar? | Tú |
| INMUEBLES | Lo que tienes en ladrillo | Tú |
| VIAJES | Cuenta de viajes, viajes y «¿y si aparto otra cantidad?» | Tú |
| PARÁMETROS | Edad, jubilación, rentabilidades, inflación, pensión | Tú; gastos y ahorro salen solos |
| REGLAS | Qué concepto va a qué categoría | Tú |
| CATEGORÍAS | Qué es ingreso, gasto o traspaso, y en qué grupo | Tú |
| PROYECCIÓN | El cálculo año a año, antes y después de la jubilación | Nadie |
| PANEL_DATOS | Lo que lee el dashboard | Nadie. No se mueve ni se renombra |

Colores: amarillo con texto azul es lo que cambias tú; verde, lo que viene de otra pestaña o de tus datos (puedes escribir encima); negro, fórmula.

## El importador

### Cada banco, un formato

Ningún banco da un CSV de verdad:

| Banco | Qué descargas | Cómo se lee |
|---|---|---|
| Openbank, cuenta y tarjeta | `.xls` que por dentro es una tabla HTML | Se lee como texto y se parsea la tabla |
| CaixaBank | `.xls` binario | Se convierte a Google Sheet con la Drive API, se leen los valores y se tira la copia |
| ING | `.xls` binario | Igual que CaixaBank |
| Cualquier otro | CSV, `.xlsx` o Google Sheet | Se lee igual; basta con añadir el banco a `BANCOS` |

El banco se reconoce por **las cabeceras de la tabla**, no por el nombre del archivo: el importador busca, en las 40 primeras filas, una que tenga todas las cabeceras de `reconoce`. Las cabeceras se comparan sin acentos ni mayúsculas, y basta con el principio («IMPORTE» encaja con «IMPORTE (€)»).

Para añadir un banco, una entrada más en `BANCOS` (`apps-script/importar.js`):

```js
{ cuenta: 'Mi banco', reconoce: ['FECHA', 'CONCEPTO', 'IMPORTE', 'SALDO'],
  fecha: 'FECHA', concepto: ['CONCEPTO'], importe: 'IMPORTE', saldo: 'SALDO' }
```

`reconoce` tiene que distinguirlo de los demás: el orden de `BANCOS` importa, gana el primero que encaja.

### No duplica

Cada movimiento lleva una clave: **cuenta, fecha, hora (solo en la tarjeta), importe y saldo**, más un contador si hay dos idénticos en el mismo archivo. El concepto no entra a propósito: es lo que más cambia entre formatos. Por eso puedes descargar extractos que se solapan: lo que ya estaba no entra otra vez.

### Saldos de fin de mes

Salen del propio extracto, no de sumar movimientos. Dentro de un mismo día, el saldo final es el del movimiento cuyo saldo no es el «saldo anterior» de otro del mismo día: así no depende del orden en que venga el archivo.

MIS DATOS empieza en el **primer mes con datos de las tres cuentas**. Un mes con una cuenta sin datos daría un patrimonio falso.

### Categorías y reglas

La primera regla de REGLAS que encaja decide la categoría: el concepto contiene el texto, y opcionalmente solo en una cuenta o solo con un signo. Lo que no encaja entra como «❓ Por revisar» y se pinta de amarillo.

**La categoría decide si un movimiento es ingreso, gasto o traspaso. El signo no.** Un dinero que te devuelven de algo que pagaste lleva la misma categoría de gasto y resta del gasto. Un regalo que te hacen va a «Regalos recibidos», de ingreso.

Los **traspasos** no cuentan ni como ingreso ni como gasto: lo que pasas entre tus cuentas, el cargo mensual de la tarjeta (las compras ya vienen del extracto de la tarjeta) y el efectivo propio que ingresas en el banco.

## Retribución flexible

Lo que pagas con una tarjeta de retribución flexible se descuenta de la nómina antes de cobrarla, así que no pasa por ningún banco. FLEXIBLE lo suma a la vez a ingresos y a gastos (grupos Comida y Transporte): el ahorro no cambia, pero tus gastos reales sí.

## VIVIENDA: ¿comprar ahora o esperar?

Compara dos caminos **con el mismo esfuerzo al mes**:

- **A, comprar ya.** Pones tus ahorros, el resto es hipoteca. Cuando la terminas, lo que pagabas de cuota pasa a invertirse.
- **B, esperar unos años.** Sigues pagando alquiler, inviertes lo que te sobra de la cuota que pagarías en A, y compras después, al precio revalorizado y con lo que hayas juntado.

Al horizonte que elijas compara el patrimonio ligado a la vivienda en los dos: vivienda menos deuda más lo invertido. El plazo se acorta solo si con tu edad te pasarías de la edad máxima que pone el banco.

Lo que decide casi siempre: la revalorización de la vivienda y los intereses de la hipoteca frente a lo que rinde el dinero invertido y lo que se va en alquiler. Los impuestos de la compra cambian según la comunidad autónoma y según si es nueva o de segunda mano: el 6 % del ejemplo es solo un ejemplo.

En DECISIONES puedes meter la compra (ahora o esperando) en tu plan: sale de tus ahorros, la cuota sale de lo que ibas a invertir, y desde que compras el alquiler deja de contar como gasto.

## La proyección y la jubilación

**Hasta la jubilación**, año a año: el efectivo, la inversión y el plan de pensiones crecen con su rentabilidad y con lo que aportas (que sube un poco cada año). Los gastos suben con la inflación. La **libertad financiera** es qué parte de tus gastos cubriría tu dinero con la regla de retirada (por defecto, el 4 %).

**Después**, cada año sacas lo que no cubren la pensión y las rentas de tus inmuebles, y el capital que queda renta algo menos que antes. «Te dura hasta los…» es la edad a la que se acaba.

Todo se calcula dos veces: **sin decisiones** (siguiendo como hoy) y **con tus decisiones** (lo que pongas en DECISIONES).

## VIAJES: ¿llega?

No es un presupuesto anual: es una cuenta a la que apartas cada mes y de la que pagas tu parte de cada viaje. Cada viaje dice cuánto queda en la cuenta después y si llega. El «¿y si aparto otra cantidad?» compara tres opciones: el saldo de la cuenta, el viaje más justo, lo que te queda de ahorro y cómo cambian el patrimonio y la libertad a la jubilación.

## El dashboard

Hay dos formas de montarlo, con el mismo `index.html` y el mismo `page.js`:

| | Lee de | Para qué |
|---|---|---|
| **En directo** | Tu sheet, con tu conexión de Google Drive | Tu plan. El dashboard descarga el sheet como `.xlsx`, lo abre en el navegador y lee solo PANEL_DATOS. No guarda cifras: sin acceso a tu archivo no se ve nada |
| **Demo** | Los JSON de `dashboard/demo/`, adjuntos al dashboard | Enseñarlo con los datos ficticios a quien no tiene el sheet |

En directo, cada dataset de `datasets.json` es una descarga del sheet y cada loader de `loaders.js` convierte el `.xlsx` en filas. `lector-xlsx.js` es la pieza común, copiada dentro de cada loader porque cada loader solo puede usar lo que lleva dentro.

En la demo no hay loaders: los datasets son archivos JSON con las filas ya hechas, sacados de PANEL_DATOS con `generador/exportar_demo.py`. El `<script src="loaders.js">` de `index.html` se ignora cuando el dashboard no tiene ese archivo.

PANEL_DATOS está hecha para que el dashboard no dependa de dónde caen las cosas en el resto del sheet: cada tabla empieza con una marca (`## KPI`, `## MESES`…), la fila siguiente son las cabeceras y la tabla acaba en la primera fila vacía.

Última actualización: 12/10/2026
