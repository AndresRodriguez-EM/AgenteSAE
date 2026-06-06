# AgenteSAE — Actualización de Notas a los Estados Financieros

Automatiza la actualización de las **Notas (Word)** de las sociedades en liquidación
administradas por la SAE, tomando como fuente de verdad (auxiliares) los
**Balances de Comprobación en PDF** exportados del ERP.

El objetivo es **actualizar únicamente la información de las tablas del Word
conservando intacto su formato** (estilos, bordes, sombreados, celdas combinadas,
fuentes, encabezados y toda la narrativa).

Sociedades soportadas: **IRCA, SANTA, MONTOYA, ZARLHA, INVERMAP, CIA**.
La **Nota 3** (Propiedad, planta y equipo) se deja intacta a propósito (se
actualiza manualmente).

## Aplicación de escritorio

Hay una interfaz gráfica lista para usar mes a mes (`app.py`, Tkinter):
seleccionas la sociedad, el PDF del período, el PDF comparativo del año anterior
y el Word de notas → **Procesar** → obtienes el Word actualizado y un informe.
Ver **[INSTALL.md](INSTALL.md)** para ejecutarla o generar un `.exe` de escritorio.

## Qué hace

1. **Lee los auxiliares (PDF).** `pdf_parser.py` reconstruye el balance de
   comprobación (código PUC, tercero, *Saldo Anterior · Débito · Crédito · Nuevo
   Saldo*) directamente desde el PDF, sin servicios externos.
2. **Actualiza el Word.** `docx_updater.py` recorre las tablas de las notas y,
   por cada fila, la empareja con el auxiliar y actualiza solo las celdas de valor.
3. **Agrega y elimina terceros.** Agrega los terceros que están en el auxiliar y
   faltan en el Word; elimina las filas cuyo tercero ya no existe en ningún
   auxiliar. Recalcula los subtotales y los cuadros de "VALOR TOTAL".
4. **Genera un informe** (`.md`) con los cambios, altas/bajas, observaciones y la
   verificación de cuadre.

### Mapeo auxiliar → nota

| Columna en la nota | Fuente en el auxiliar |
|---|---|
| Saldo / Acumulado **mes actual** | Nuevo Saldo del auxiliar del periodo |
| Columna **… 2025** (mes anterior / año) | Nuevo Saldo del auxiliar comparativo |
| **Mov / Movimiento mes** | \|Débito − Crédito\| del periodo |
| Terceros (filas) | Cada NIT/cédula de la cuenta |
| Valor Total Nota X | Saldo de la clase/cuenta correspondiente |

Los valores se muestran en **positivo** salvo la Nota 5 (Patrimonio), que usa la
convención de signo propia (aportes en positivo, pérdidas en negativo).

### Criterios de seguridad ("do no harm")

- Solo se tocan **celdas de tabla**; la narrativa y los estilos nunca se modifican.
- Una celda **no se cambia** si su valor ya coincide con una cifra válida del
  auxiliar (a nivel de tercero **o** de cuenta), o si la diferencia es de redondeo
  (≤ 1–2 pesos; los auxiliares traen decimales y el Word muestra enteros).
- El emparejamiento es por **NIT** y se desempata por **valor** y por la
  **descripción de cuenta**; reconoce filas que el Word presenta consolidadas a
  nivel de cuenta (p. ej. retención, IVA, impuestos).
- Si un tercero coincide por **nombre** pero con NIT distinto, se conserva el del
  Word y se reporta como **observación** (posible error de digitación del NIT).

## Uso

```bash
pip install -r requirements.txt

python -m agentesae.cli \
  --actual       data/input/AUX_2026.pdf \
  --comparativo  data/input/AUX_2025.pdf \
  --notas        data/input/NOTAS.docx \
  --salida       output/NOTAS_ACTUALIZADO.docx \
  --informe      output/informe.md \
  --sociedad     "NOMBRE DE LA SOCIEDAD"
```

Para procesar **varias sociedades**, repetir el comando con sus respectivos
archivos (mismo template de notas).

## Estructura

```
agentesae/
  pdf_parser.py     # PDF (balance de comprobación) -> datos estructurados
  docx_updater.py   # actualiza las tablas del Word preservando el formato
  cli.py            # orquestador + informe de conciliación
data/input/         # auxiliares y notas de entrada  (no versionado)
output/             # notas actualizadas + informes   (no versionado)
```

> Los archivos con datos financieros de las sociedades **no se versionan**
> (ver `.gitignore`). El repositorio contiene solo el motor reutilizable.

## Validación del piloto (INVERSIONES IRCA C.I. S.A.S. EN LIQUIDACIÓN, abril 2026)

- Estructura preservada: 47 tablas, 524 párrafos, narrativa idéntica.
- Todos los elementos de formato (bordes, sombreado, negrita, color, tamaños,
  celdas combinadas) con conteo idéntico; los `.docx` solo difieren en la
  declaración XML al reserializar (sin efecto visual).
- Único cambio material detectado y corregido: **Nota 9, Valor Total
  $ 1.361 → $ 1.667** (error de digitación; el resto del documento ya estaba
  conciliado con los auxiliares).
- Observación reportada: el NIT de un tercero difiere entre el Word y el auxiliar.

## Notas y límites

- Las tablas no contables (jurídico, contratos de arrendamiento, presupuesto) no
  se tocan: no provienen del balance de comprobación.
- La planilla mensual de honorarios del depositario (tabla anidada) se valida pero
  su actualización mes a mes es una mejora pendiente.
- La columna de **edad/días** de cartera no se calcula desde el auxiliar.
- La configuración de tablas (`docx_updater.TABLES`) asume el template vigente de
  notas; para otro template basta ajustar esa configuración.
