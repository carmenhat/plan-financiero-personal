# Aprendizajes

Lo que ya falló, lo que se descartó y por qué. Leer primero si algo no cuadra.

---

## Fallos que ya pasaron

**La zona horaria del script y la del sheet no coincidían.** El script escribía «1 de octubre» como el 30 de septiembre a las 15:00, los meses dejaban de casar y MIS DATOS salía a cero, sin ningún error. Ahora todo lo que se lee se pasa a texto `aaaa-mm-dd` en la zona del sheet de origen, y lo que se escribe se convierte a medianoche en la zona del sheet (`deSheet_` y `aSheet_` en el importador).

**Las fórmulas que escribe un script se interpretan con la configuración regional del sheet.** En español los argumentos van separados por «;», así que una fórmula con comas da `#ERROR!`, hasta la más simple. Las del `.xlsx` importado funcionaban porque llegan ya traducidas, y eso despista. Los scripts cambian la configuración a inglés un instante, escriben y la devuelven (`enIngles_`).

**La exportación de Google cierra las filas vacías en sí mismas** (`<row r="38"/>`). El lector del dashboard se tragaba esa fila junto con la siguiente, todo se descolocaba una fila y el panel enseñaba «undefined». En LibreOffice no pasaba, porque escribe las filas de otra forma: **la prueba tiene que ser contra la exportación real de Google**.

**Un regalo recibido contaba como gasto negativo.** No era un fallo del cálculo: la categoría «Regalos» es de gasto, y un ingreso con categoría de gasto resta del gasto. Se añadió «Regalos recibidos», de ingreso. Es la consecuencia directa de que decida la categoría y no el signo.

**El efectivo propio ingresado en el banco inflaba el ahorro.** Contaba como ingreso, y es dinero que ya tenías. Ahora es la categoría «Ingreso de efectivo propio», de traspaso.

## Descartado

| Qué | Por qué |
|---|---|
| Exportar en CSV la pestaña del dashboard | La exportación CSV de Drive solo da la primera pestaña, y con los números en formato español. Se exporta el `.xlsx` entero y se lee la pestaña que toca |
| Fórmulas en notación R1C1 desde Apps Script | Google no acepta la forma corta (`RC1`) y deja `#ERROR!`. Se escriben en notación A1 |
| Regenerar el sheet entero en cada cambio | Pierde lo que has tocado a mano. Los cambios de estructura se hacen con un script de una sola vez |
| Calcular los saldos sumando movimientos | Cualquier hueco en el histórico se arrastra para siempre. El saldo que da el banco es el dato |
| Identificar el banco por el nombre del archivo | Cambia con cada descarga. Las cabeceras no |
| Meter el concepto en la clave de un movimiento | Es lo que más cambia entre formatos del mismo banco: el mismo movimiento entraría dos veces |
| Guardar IDs en Propiedades del script | No hay ninguno que guardar: la carpeta `Movimientos` se busca junto al sheet y se crea si falta |
| Copiar las cifras dentro del dashboard en directo | Con datos reales, cualquiera con el enlace las vería. Solo la demo, con datos ficticios, las lleva dentro |

## Lo que costó entender

- **Los `.xls` de los bancos casi nunca son lo que dicen.** Los de Openbank son una tabla HTML con extensión `.xls`; los de CaixaBank e ING, `.xls` binarios de verdad. El importador mira los primeros bytes, no la extensión.
- **Un movimiento de tarjeta pendiente puede cambiar de importe al liquidarse.** Si pasa, entra dos veces, porque la clave lleva el importe.
- **La media de gastos no ve lo que no pasa por tus cuentas:** lo que paga otra persona, el efectivo, la retribución flexible si no la apuntas. Por eso PARÁMETROS deja escribir encima.

Última actualización: 12/10/2026
