# 📦 warehouseMos · Módulo Guías · Análisis y Rediseño

> **Fecha**: 2026-05-20
> **Autor**: Claude + Luis
> **Estado**: Análisis pendiente de aprobación · NO se tocó código todavía

---

## 🩺 Diagnóstico — por qué se siente lento y desordenado

### Frontend
| Síntoma | Causa raíz |
|---|---|
| **3-4 clicks** para crear una guía simple | Picker tipo → subtipo → entidad → form. Cada paso abre un sheet nuevo |
| **Picker de entidad demora** (proveedor/zona/jefatura) | Se carga **on-demand** desde MOS, sin precarga al boot |
| **Comentario y tags densos** | 4 botones toggle (comp/compl) sin agrupación visual + textarea sin contador |
| **Sin feedback al agregar ítem** | No hay spinner ni animación de "escaneado" tras tocar producto en scanner |
| **Pantalla congelada al cerrar guía** | Espera en serie la respuesta de GAS (5-30s si es INGRESO_PROVEEDOR con sync a MOS) |
| **Modal PIN admin lento** | `_verificarAdminPin` no tiene timeout → si MOS está lento, cuelga 30-60s sin feedback |
| **"Abrir guía cerrada" se demora** | El botón lock dispara: validar PIN online (~2s MOS) + reabrirGuia con `_sheetToObjects` (~800ms) + revertir stock por cada detalle. **Sin loading visible durante esos pasos** |

### Backend (GAS)
| Bottleneck | Impacto |
|---|---|
| **LockService global** 60s timeout | Si 10 users hacen POST en paralelo en pico, 9 esperan 60s c/u |
| **UrlFetchApp serial a MOS** en `cerrarGuia` (sync productos prov) | 50 items × 2-5s c/u = potencialmente 100-250s |
| **`flush()` + `sleep(1500) × 3`** en `imprimirTicketGuia` | **4.5s garantizados** de espera ciega |
| **`_sheetToObjects`** en cada GET (sin índices) | 1-5s por endpoint si tablas grandes |
| **Validación admin sin caché en GAS** | Cada PIN va round-trip a MOS aunque WH ya tenga el caché local |

### El núcleo del problema
> **La app habla con GAS en tiempo real para CADA cosa, y GAS habla con Sheets sin índices y a veces con MOS sobre HTTP. Es síncrona end-to-end. El user paga la latencia en cada click.**

---

## 🎨 Mockups — cómo podría verse el módulo rediseñado

### A) Pantalla principal de guías — "Centro de Operaciones"

```
┌─────────────────────────────────────────────────────────┐
│ 🏭 ALMACEN MOS · MIÉ 20 may                  🟢 Online │
│ Jorgenis · Almacenero · 4h 22m de turno                 │
├─────────────────────────────────────────────────────────┤
│ ⚡ HOY  · 12 guías · S/ 4,820 movido · 3 abiertas      │
│ ┌───────────┬───────────┬───────────┬───────────┐      │
│ │ 🟢 ABIERTAS│ ⏳ POR     │ 📥 PREING.│ 🏷 ETIQS │      │
│ │     3     │ CERRAR · 1 │     5     │  12 nuev. │      │
│ └───────────┴───────────┴───────────┴───────────┘      │
│                                                          │
│ 🔥 NECESITAN TU ATENCIÓN                                │
│ ┌──────────────────────────────────────────────────┐   │
│ │ 🟡 G-1729 · SAL_ZONA Pampa Hermosa · 4 items     │   │
│ │    abierta hace 2h · sin movimiento desde 14:32  │   │
│ │    [✖ Cancelar]  [📦 Continuar]                  │   │
│ └──────────────────────────────────────────────────┘   │
│                                                          │
│ 📅 AYER · 18 guías · ▶ ver                              │
│ 📅 15 may · 22 guías · ▶ ver                            │
│                                                          │
│           ╭──────────────────────╮                      │
│           │  ➕  NUEVA GUÍA       │  ← FAB grande       │
│           ╰──────────────────────╯                      │
└─────────────────────────────────────────────────────────┘
```

**Cambios**:
- KPI cards arriba (clickeables → filtran lista) en lugar de solo lista
- Sección "Necesitan tu atención" prioriza guías estancadas (>2h sin movimiento) o preingresos sin guía
- FAB grande y central, no escondido en la esquina

### B) Crear guía — flujo de 1 paso, no 3

```
┌─────────────────────────────────────────────────┐
│ ➕ Nueva guía                            ✕      │
├─────────────────────────────────────────────────┤
│ ¿Qué vas a hacer?                                │
│                                                  │
│ ┌─────────────────┐  ┌─────────────────┐        │
│ │ 📥 INGRESO      │  │ 📤 SALIDA       │        │
│ │ (recibo merca)  │  │ (despacho merca)│        │
│ └─────────────────┘  └─────────────────┘        │
│                                                  │
│ ─── O elegí lo más frecuente ───                │
│                                                  │
│ 🚀 Salida → Pampa Hermosa  (12 esta semana) →   │
│ 🚀 Ingreso ← Don Lucho     (4 esta semana)  →   │
│ 🚀 Ingreso ← Marvisur      (3 esta semana)  →   │
│                                                  │
│ Cada chip te lleva directo al scanner.          │
└─────────────────────────────────────────────────┘
```

**Cambios**:
- Solo 2 botones grandes ingreso/salida + shortcuts inteligentes (top 3 entidades más usadas del último mes, calculadas en cliente con datos locales)
- Tocar un shortcut **te lleva DIRECTO al scanner** (saltea el form) — para el 70% de los casos
- Sin elegir entidad si el contexto lo deduce

### C) Scanner / agregar ítems — feedback inmediato

```
┌─────────────────────────────────────────────────┐
│ G-1729 · Pampa Hermosa · 7 items     [✓ Cerrar]│
├─────────────────────────────────────────────────┤
│         ╭──────────────────────╮                │
│         │  📷 ESCÁNER ACTIVO    │  fps:30        │
│         │  ║║│║│║║║│║║│         │                │
│         ╰──────────────────────╯                │
│                                                  │
│ ─── PARA OREO ROLLO 108GR ─── ✨ +1 (3 total)   │
│ ✓                                                │
│ Aparece banner verde 800ms, sonido beep doble   │
│                                                  │
│ ⬇ ITEMS EN LA GUÍA                              │
│ ┌───┬─────────────────────────────────┬──────┐ │
│ │ 3 │ OREO ROLLO 108GR        ✨pulse │  ⋮   │ │ ← anima al +1
│ │ 5 │ KR GASEOSA COLA 350ML           │  ⋮   │ │
│ │ 2 │ BEARY PANCO BLANCO 1KG          │  ⋮   │ │
│ └───┴─────────────────────────────────┴──────┘ │
│                                                  │
│ 💬 dictar producto: 🎤 "diez kr gaseosa"        │
└─────────────────────────────────────────────────┘
```

**Cambios**:
- Banner "✨ +1" verde brillante 800ms con el producto recién escaneado (feedback de ganga)
- La fila del producto hace **pulse glow ámbar** 600ms al sumar — confirmación visual
- Sonido distinto: nuevo (`beep`) vs +1 a existente (`beepDouble`)
- Botón micrófono para dictar producto + cantidad ("diez kr gaseosa")

### D) Modal admin PIN — moderno y rápido

```
┌─────────────────────────────────────────────────┐
│            🔐                                     │
│         AUTORIZACIÓN                             │
│   Reabrir guía G-1729                            │
│                                                   │
│  ╔════ Global ════╗  ╔══ PIN Jorgenis ══╗       │
│  │  ●   ●   ●   ●  │  │  ○   ○   ○   ○   │       │
│  ╚═══════════════╝  ╚═════════════════╝          │
│                                                   │
│         ┌─────┬─────┬─────┐                      │
│         │  1  │  2  │  3  │                      │
│         ├─────┼─────┼─────┤                      │
│         │  4  │  5  │  6  │  → cada tap          │
│         ├─────┼─────┼─────┤    hace vibrate(15) │
│         │  7  │  8  │  9  │    + click sonido    │
│         ├─────┼─────┼─────┤                      │
│         │  ⌫  │  0  │  ✓  │                      │
│         └─────┴─────┴─────┘                      │
│                                                   │
│  ⚡ Validando... 1.2s                            │
│  [████████████░░░░] 75%                          │
│                                                   │
│  💡 Tip: si lo dictás, decí "uno dos tres..."    │
└─────────────────────────────────────────────────┘
```

**Cambios**:
- Visual claro: dos bloques de 4 dots (global vs personal), no 8 corridos
- Avatar del admin que está logueado: "PIN Jorgenis"
- Progress bar real mientras valida (no spinner ambiguo)
- Timeout 5s explícito con mensaje "MOS lento, intentando local..." → fallback a cache automático
- Sonido distinto por estado: tap (clic), success (campanita), error (buzzer)
- Dictado por voz como bonus

### E) Detalle de guía cerrada con timeline

```
┌─────────────────────────────────────────────────┐
│  ✕     G-1729 · CERRADA  · monto S/ 247.50      │
├─────────────────────────────────────────────────┤
│  📤 Salida a Pampa Hermosa                      │
│  Creada hace 2h por Jorgenis · 12 items         │
│  Cerrada hace 30min por Jorgenis                │
│                                                  │
│  ⏱ TIMELINE (cronología)                        │
│  │                                               │
│  ● 14:32 · Jorgenis creó la guía                │
│  │                                               │
│  ● 14:33 · escaneó OREO ROLLO ×2                │
│  ● 14:34 · escaneó KR COLA ×5                   │
│  ● 14:35 · ✏ editó cantidad OREO 2→3            │
│  ● 14:48 · cerró guía · stock aplicado          │
│  │                                               │
│  ● 15:02 · 🖨 ticket reimpreso                  │
│  │                                               │
│                                                  │
│  ─────────────────────────────────────          │
│  📦 ITEMS (12)                                  │
│  ...                                             │
│                                                  │
│  [🔐 Reabrir]  [🖨 Reimprimir]  [📊 Reporte]   │
└─────────────────────────────────────────────────┘
```

**Cambios**:
- Timeline visual desde OPS_LOG (ya existe el dato) — el admin entiende qué pasó sin ir a "Historial"
- Acciones admin destacadas con candado + tooltip "Requiere PIN"
- Resumen ejecutivo arriba (cuántos items, cuánto monto, quién hizo qué)

---

## ⚡ Mejoras de performance — UI Optimista

### 1. Optimistic UI universal
Hoy: tocás "agregar item" → spinner → 800ms → aparece. Mejora: agregás item **localmente en la UI INMEDIATA** y la cola idempotente lo sincroniza con GAS en background. Ya existe `localId` + `idOp` + `OPS_LOG` (todo el plumbing está hecho). Solo falta **no esperar la respuesta para renderizar**.

### 2. Pre-carga al boot
Al hacer login, en paralelo (no en serie):
- Proveedores top-20 (más usados últimos 30 días)
- Zonas activas
- Jefatura/admins
- Caché de PINs admin
- Productos master (delta sync — solo cambios desde último login)

Eso elimina el "Cargando..." del entity picker.

### 3. PIN admin con caché optimista
Hoy: cada PIN va a MOS via UrlFetchApp serial. Mejora:
- Validar local PRIMERO (cache TTL 1h del PIN bundle)
- Si pasa local → ejecutar la acción
- En background, revalidar con MOS (sin bloquear UI)
- Si MOS dice "PIN cambió", invalidar caché y mostrar toast "tu PIN se actualizó, vuelve a entrar"

Esto reduce el "abrir guía con PIN" de **2-3s a 200ms percibido**.

### 4. Cerrar guía → response inmediato + sync diferida
- Click "Cerrar" → marca CERRADA local + muestra toast "🎉 cerrada"
- En background: sync a MOS (skipMosSync por defecto, hace batch nocturno)
- Si falla, badge naranja "1 sync pendiente" en header

### 5. Imprimir ticket sin retry blocking
Eliminar `Utilities.sleep(1500) × 3` del backend. En su lugar:
- Cliente manda `flush=true` solo cuando agrega items críticos
- Imprimir = job en cola PrintNode + retry async hasta 3x
- El user no espera = el ticket sale cuando sale (típico 2-4s en background)

### 6. Batch endpoints en GAS
- `_sheetToObjects(GUIA_DETALLE)` se llama 3-4 veces en `cerrarGuia` + `imprimirTicketGuia`. Cachear en memoria del mismo request.
- Endpoint `getModuloGuiasContexto(idGuia)` que devuelve **TODO** lo que el frontend necesita en 1 request (guía + detalles + productos + entidad + historial top 10). Hoy son 3-5 endpoints.

### 7. Index dinámico en Sheets
Crear hojas espejo "GUIAS_INDEX" mantenidas por trigger:
- 1 fila por guía con campos enriquecidos (estado, monto, lastUpdate, ...)
- `getGuias` lee SOLO esa hoja → O(1) por filtro
- Migration: backfill 1 vez

---

## 🎵 Mejoras visuales y sonoras

### Animaciones
- Slide-up smoother con `cubic-bezier(.16,1,.3,1)` en lugar del ease-out actual
- Item pulse glow al agregar ítem (ya propuesto)
- Confetti al cerrar guía con > 50 items o monto > S/1000 (logro!)
- Wave shimmer en cards mientras cargan (no spinner solo)
- Transición page→page con cross-fade (no cut seco)

### Sonidos (ampliar los actuales)
- `coin.mp3` o sintético al sumar monto > S/500 (gratificante)
- `scanFail` distinto del scan correcto (tres notas bajas)
- `lockOpen` (clack metálico) cuando se aprueba el PIN admin → feedback de éxito tactil
- `priceUpdate` (campanita) al editar precio masivo
- Toda configurable en preferencias del user

### Feedback visual de estado
- Badge "🔴 Sin red · 3 pendientes" persistente en header — hoy es pequeño y discreto
- Indicador de cola con animación de progreso real (no solo número)
- Auto-reconexión visible: cuando vuelve red, banner verde 2s "🟢 Reconectado · sincronizando 3 ops..." con barra de progreso

### Modo "Operación rápida"
Toggle en settings que activa:
- Auto-cerrar guía al primer ítem si tipo == ENVASADO simple
- Auto-aplicar tag "completo" si todos los items se escanearon
- Auto-imprimir ticket sin preguntar al cerrar
- Atajos de teclado (tablets POS conectados a teclado físico)

---

## 🤖 Oportunidades de IA

### 1. OCR de boleta del proveedor → preingreso auto (ya en roadmap)
- Foto de la boleta → Gemini/Claude vision extrae: proveedor, items con cantidades, precio total
- Pre-llena el preingreso con confianza score por campo
- User solo revisa y confirma

### 2. Sugerencia de productos al scanner
- Cuando ingresa "ALOE BEBI..." en búsqueda, el modelo sugiere "Probablemente quieres KERO ALOE PIÑA" basado en histórico
- Local con un fuzzy matcher + ranking por frecuencia

### 3. Validación inteligente al cerrar guía
- "⚠ Estás cerrando una INGRESO_PROVEEDOR sin foto de boleta. ¿Continuar?"
- "⚠ Esta guía tiene 23 items pero el preingreso esperaba 28. ¿Falta algo?"
- "✓ Todo coincide con el preingreso original"

### 4. Predicción de cantidad faltante
- "Para Pampa Hermosa, en promedio en miércoles despachás 14 cajas de OREO. Hoy llevas 3."
- Usa histórico de últimos 4 miércoles, normalizado por temporada

### 5. Chat almacén (también en roadmap)
- "Cuánto stock tengo de KR gaseosa cola 350?" → "23 unidades, vencen 12 nov 2026"
- "¿Qué guías tengo abiertas?" → lista
- "Mostrame el último despacho a Pampa" → abre detalle
- Endpoint en GAS que recibe NL, llama a Gemini con tools del sistema

### 6. Anti-fraude soft
- Si un cajero anula 3+ items en 5min de una guía cerrada → alerta a admin master (sin bloquear, solo flag)
- Si reabren la misma guía 2× en un turno → log estructurado a OPS_LOG con flag de revisión

### 7. Dictado de items por voz
Web Speech API (`es-PE`):
- "diez kr gaseosa cola" → busca match, agrega 10
- "borrar último" → quita la última fila
- "cerrar guía" → confirma y cierra
- Trabaja **junto** al scanner, no en lugar de

---

## 🗺 Roadmap sugerido (por impacto/esfuerzo)

### 🔥 Tier 1 — Quick wins (1-2 días c/u)
1. **Pre-carga al boot** de proveedores/zonas/admin cache (elimina lentitud en pickers)
2. **PIN admin con caché optimista local + revalidación bg** (resuelve "demora al abrir guía")
3. **Timeout 5s explícito** en fetch del PIN con fallback a local + toast claro
4. **Eliminar `sleep(1500)×3`** de `imprimirTicketGuia` (4.5s ganados sí o sí)
5. **Loading visible** en abrir detalle / cerrar guía / reabrir (no más "congelado")

### ⚡ Tier 2 — Rediseño UX (3-5 días c/u)
6. **Smart Quick Create** (1 paso en vez de 3, con shortcuts top entidades)
7. **Banner +1 verde con pulse glow** al escanear (feedback tactil)
8. **Modal PIN moderno** (dos bloques de dots + progress bar + sonidos)
9. **Centro de Operaciones** (dashboard con KPI cards arriba + "necesitan atención")
10. **Timeline en detalle** de guía cerrada

### 🚀 Tier 3 — Optimismo UI + Backend (1-2 semanas)
11. **Optimistic UI universal** (sin esperas en agregar item, cerrar guía, etc.)
12. **Endpoint compuesto** `getModuloGuiasContexto` (1 req en vez de 5)
13. **Hojas índice mantenidas por trigger** (queries O(1))
14. **Skip MOS sync por default + batch nocturno** (cerrar guía es instantáneo)

### 🧠 Tier 4 — IA (cada feature 1 semana)
15. **OCR boleta → preingreso** (mayor ROI: ahorra 5min por preingreso)
16. **Dictado de items por voz** (alternativa al scanner cuando manos están ocupadas)
17. **Chat almacén** (admin pregunta stock/guías por NL)
18. **Validación inteligente al cerrar** (avisos pre-cierre)

---

## 💡 Recomendaciones finales

1. **No reescribir todo** — el plumbing (idempotencia, OPS_LOG, offline queue) ya es bueno. Hay que **usarlo bien** desde el frontend (optimistic UI).
2. **Empezar por Tier 1** — los 5 quick wins resuelven el 80% de la sensación de lentitud sin tocar la arquitectura.
3. **Tier 2 da el "wow"** — el dashboard + flow rápido + modal PIN moderno son lo que se siente al abrir la app.
4. **Tier 3 y 4 son inversiones** — pagan a mediano plazo. El OCR de boleta solo ahorra **horas/semana** en el equipo.
5. **Medir** — agregar un endpoint simple `_logUx({accion, tiempo_ms})` para ver realmente dónde se va el tiempo. Sin métricas, optimizamos a ciegas.

---

## 📌 Estado al momento de guardar este doc

- v ME en producción: **2.5.38** (wizard moderno + verif pre-Vue + FAB etiquetas + DataCloneError + scroll detalle + fixes _wizBeep/_pinPress/_fmtCountdown + banner zombie)
- v MOS en producción: **2.41.88** (cobros en vuelo + scroll moderno carta detalle)
- v WH en producción: **2.13.40** (ticket preingreso imprime comentario completo)
- Pantallas blancas: ✅ resueltas
- Cambios pendientes en este análisis: **NINGUNO en código** — esperando aprobación de Luis para empezar Tier 1
