# Oro XAU/USD — Tendencia + Pullback + RSI + ATR (Pine Script v6)

Indicador para **TradingView** diseñado para operar **Oro (XAU/USD)** en temporalidades de
**5 minutos** y **15 minutos**. Fusiona tres conceptos para filtrar falsas rupturas y
capturar movimientos de alta probabilidad.

Archivo: [`XAUUSD_TrendPullback_RSI_ATR.pine`](./XAUUSD_TrendPullback_RSI_ATR.pine)

## Lógica del indicador

1. **Estructura / Tendencia (filtro principal)**
   - `EMA 200` define la tendencia macro. Precio por encima → solo **compras**; por debajo → solo **ventas**.
   - `EMA 50` marca la zona de **retroceso (pullback)** de medio plazo.

2. **Momento (gatillo)**
   - `RSI 14`.
   - **LONG**: precio > EMA200, retroceso que **toca la EMA50**, y el RSI **sale de sobreventa**.
   - **SHORT**: precio < EMA200, pullback que **toca la EMA50**, y el RSI **sale de sobrecompra**.

3. **Gestión de riesgo (visual)**
   - Etiquetas **BUY** (verde) / **SELL** (rojo) al **cierre** de la vela.
   - Cajas de **Stop Loss** y **Take Profit** dimensionadas por **ATR 14**:
     - `Stop Loss = 1.5 × ATR` desde la entrada.
     - `Take Profit = Ratio × riesgo` (por defecto **1:2**, el doble del SL).

## Sin repintado (no-repaint)

- Las condiciones usan solo datos pasados/cerrados (EMA, RSI, ATR, máximos/mínimos).
- La señal **solo se confirma al cierre** de la vela (`barstate.isconfirmed`).
- No usa `request.security` con lookahead ni referencias a barras futuras → las señales y las
  cajas SL/TP **no se mueven** una vez impresas.

## Cómo instalarlo

1. Abre TradingView → gráfico de **XAU/USD** en **5m** o **15m**.
2. Abre el **Pine Editor** (parte inferior).
3. Copia y pega el contenido de `XAUUSD_TrendPullback_RSI_ATR.pine`.
4. Pulsa **Add to chart / Guardar**.
5. (Opcional) Crea alertas con las condiciones *BUY XAUUSD*, *SELL XAUUSD* o *Cualquier señal XAUUSD*.

## Parámetros (menú de configuración)

| Grupo | Parámetro | Por defecto | Descripción |
|------|-----------|-------------|-------------|
| Tendencia | EMA Tendencia Macro | 200 | Filtro de tendencia principal |
| Tendencia | EMA Medio Plazo / Pullback | 50 | Zona de retroceso |
| Tendencia | Exigir alineación EMA50 vs EMA200 | off | Tendencia más estricta |
| RSI | Periodo RSI | 14 | |
| RSI | RSI Sobreventa / Sobrecompra | 30 / 70 | Usa **20 / 80** para señales más estrictas |
| RSI | Gatillo RSI | Salida de zona | Salida (recomendado) o entrada en zona extrema |
| ATR | Periodo ATR | 14 | |
| ATR | Multiplicador SL (x ATR) | 1.5 | Distancia del Stop Loss |
| ATR | Ratio Riesgo/Beneficio (TP) | 2.0 | 1:2 |
| Pullback | Ventana de Pullback | 6 | Velas para validar el toque de EMA50 |
| Pullback | Tolerancia toque EMA50 (x ATR) | 0.5 | Banda alrededor de la EMA50 |
| Pullback | Velas mínimas entre señales | 5 | Cooldown anti-ruido |
| Visual | EMAs / Cajas / Panel / Colores | — | Personalización |

> **Nota sobre RSI 20/80**: la especificación original menciona cruces en 20/80. Por
> defecto se usa **30/70** porque en Oro a 5m/15m los extremos 20/80 son poco frecuentes y
> generarían muy pocas señales. Cambia los umbrales a **20/80** si prefieres el comportamiento
> estricto descrito originalmente (en modo *"Salida de zona"*, `crossover(RSI, 20)` equivale a
> "saliendo de sobreventa profunda").

## Aviso

Herramienta educativa de análisis técnico. No es asesoría financiera. Realiza tu propio
*backtesting* y gestión de capital antes de operar con dinero real.
