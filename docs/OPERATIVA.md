# Operativa

El día a día con el plan ya montado. Cómo está hecho y por qué, en [`COMO_FUNCIONA.md`](COMO_FUNCIONA.md).

---

## Cada mes

1. **Descarga los extractos** de tus bancos, en el formato que dé cada uno. Da igual que se solapen con lo ya importado: no duplica.
2. **Déjalos en la carpeta `Movimientos`**, junto al sheet.
3. En el sheet, **Plan financiero → Importar movimientos.** Sale un resumen por archivo: cuántos nuevos, cuántos ya estaban y cuántos quedan por revisar.
4. **Revisa lo amarillo en MOVIMIENTOS:** elige categoría y explica en Nota. Si algo se va a repetir, crea la regla en REGLAS y pulsa *Plan financiero → Reaplicar reglas a lo pendiente*.
5. Si tienes retribución flexible, **apunta en FLEXIBLE** lo de la nómina del mes: una fila por nómina.

El dashboard se actualiza solo la próxima vez que lo abres.

## Cómo se cambia algo del plan

Hay tres clases de cambio, y solo la última toca código.

### 1. Una cifra o una decisión → en el sheet, y una línea en HISTORIAL

Ejemplo: *desde noviembre aparto 200 € al mes para viajes.*

1. VIAJES → **Aportación mensual** → `200`.
2. HISTORIAL → fila nueva arriba: fecha, «Cuenta de viajes», `150 €/mes`, `200 €/mes desde noviembre`, el porqué, `VIAJES`.

Nada más: MI FUTURO, la cuenta de viajes, el «¿y si…?» y el dashboard se recalculan solos.

| Cambio | Dónde |
|---|---|
| Aportación a viajes, viajes nuevos, fechas, costes | VIAJES |
| Pensión, edad de jubilación, rentabilidades, inflación | PARÁMETROS |
| Gastos reales, si sabes que gastas más de lo que ve el plan (lo que paga otra persona, el efectivo) | PARÁMETROS → escribe tu cifra encima de la fórmula |
| Precio de la vivienda, ahorros que pondrías, hipoteca, alquiler | VIVIENDA |
| Lo que tienes en ladrillo | INMUEBLES |
| Probar algo sin decidirlo («¿y si invierto 20.000 €?») | DECISIONES. No va a HISTORIAL hasta que lo decidas |

### 2. Una categoría o una regla → en el sheet

- **Categoría nueva:** fila al final de CATEGORÍAS, con Tipo (Ingreso, Gasto o Traspaso) y Grupo. Sale sola en el desplegable de MOVIMIENTOS. Si es del grupo Casa, añade también su fila en CASA.
- **Regla nueva:** fila en REGLAS. Gana la primera que encaja de arriba abajo: las específicas van antes que las generales.
- **La categoría decide si es ingreso o gasto, no el signo.** Dinero que te dan (un regalo, una venta) → categoría de ingreso. Dinero que te devuelven de algo que pagaste → la misma categoría de gasto que el pago.

### 3. Algo de estructura → código

Un banco nuevo, una pestaña nueva, un gráfico nuevo en el dashboard, una columna nueva.

- **En el importador:** cámbialo en Apps Script y pruébalo ahí, que es lo que corre. Si mantienes un clon de este repositorio, copia después el cambio a `apps-script/importar.js` y pasa `node test/test-importar.js`.
- **En el sheet:** con un script de una sola vez que añada lo nuevo sin tocar lo demás. Regenerar el sheet con `build.py` pierde lo que has tocado a mano.
- **Una fórmula escrita por un script** se escribe con la hoja en inglés un instante: copia `enIngles_` del importador.
- **Una tabla nueva en PANEL_DATOS** va al final. El dashboard busca cada tabla por su marca (`## NOMBRE`), pero mover las de en medio no aporta nada.

## Si algo falla

| Síntoma | Qué mirar |
|---|---|
| `#ERROR!` en celdas que ha escrito un script | La configuración regional: el script tiene que escribir las fórmulas en inglés (`enIngles_`) |
| MIS DATOS con ingresos y gastos a cero | Los meses de la columna A no caen en día 1 a medianoche: zona horaria del script distinta de la del sheet. Iguálalas y pulsa *Plan financiero → Recalcular saldos de MIS DATOS* |
| «Formato no reconocido» | El banco ha cambiado las cabeceras, o no está en `BANCOS`. El archivo se queda en la carpeta |
| La importación falla con un `.xls` binario | Que el servicio **Drive API** siga activado en Apps Script |
| El dashboard enseña «—» o no carga | Que PANEL_DATOS exista y no se haya renombrado. En el dashboard, *Fuentes* dice qué bloque falla |
| El dashboard enseña todo descolocado una fila | El lector no ha previsto una fila vacía cerrada en sí misma (`<row r="38"/>`). Ver [`APRENDIZAJES.md`](APRENDIZAJES.md) |
| La tarjeta no cuadra con el banco | Un movimiento que cambió de importe al liquidarse puede haber entrado dos veces. Búscalo en MOVIMIENTOS |

Última actualización: 12/10/2026
