# Plan de Modernización · Módulo Guías (warehouseMos)

> Snapshot completo del plan acordado entre Luis y Claude.
> Fecha: 2026-05-10 · sábado · 9 mayo
> Sesión: conversación previa al primer commit.
> **Estado:** plan aprobado en su mayoría, falta confirmar orden de fases (F0+F5 vs F0+F2).
> No se ha tocado código todavía.

---

## Índice

1. Contexto y dolor #1
2. Decisiones del usuario (corregidas / finales)
3. Modelo conceptual final
4. Wireframes finales
5. Plan de implementación — 10 fases
6. Archivos a tocar
7. Dependencias entre fases
8. Pregunta pendiente
9. Reglas duras a respetar

---

## 1. Contexto y dolor #1

### 1.1 Qué es el módulo
"Guías" es el corazón de la PWA warehouseMos. Vive en `C:\Users\ISO\warehouseMos`. Hoy tiene 3 vistas hermanas:

- `view-guias` — listado + detalle de GUIA_INGRESO / GUIA_SALIDA / MERMA / ENVASADO
- `view-despacho` — despacho rápido + atender pickup ME→WH (mezclados)
- `view-preingresos` — recepción de proveedor (con fotos + cargadores)

Comparten bottom-nav y tienen flujos cruzados (crear guía desde preingreso, etc.).

### 1.2 Dolor #1 — pérdida de información
**El reclamo real del operador:** cards que muestran "guardando…" se quedan colgadas; al volver a la guía las cantidades aparecen modificadas o desaparecidas; a veces el double-click crea 2 guías.

**Hipótesis técnica (5 fallas):**
1. Race condition por debounce — múltiples PUT con state distinto, último gana.
2. No hay idempotency — cada llamada es set absoluto, no delta.
3. "Guardando" sin timeout ni retry visible.
4. Refetch destructivo — server sobreescribe ops pending locales.
5. Crear guía sin idempotency key — double-click crea 2.

**La fix raíz** es migrar de "sincronización de estado" a **log de operaciones idempotente**. Ese es el corazón de la fase 5.

### 1.3 Contexto mobile-first
- Uso principal: móvil.
- Uso ocasional: tablet.
- Almacén ruidoso → sonidos/haptics deben ser diferenciados e intuitivos.
- ~10 guías por día, 1-3 preingresos por día → cards NO se achican.

---

## 2. Decisiones del usuario (corregidas / finales)

| Tema | Decisión final |
|---|---|
| Reabrir guía | Clave admin de MOS (8 dígitos, compartida admin+master) — NO PIN propio |
| Reabrir < 5 min post-cierre | Sin clave (ventana de gracia) |
| Recordar clave admin 30 min en dispositivo | Sí, toggle en modal |
| TTS post-cierre | Lee solo `cantidad + nombre producto`. Operador escucha mientras llena carreta. Si oye algo raro → reedita en ventana 5 min |
| Multi-device del mismo operador | Debe sincronizar sin lock, vía op-log |
| Multi-operador en misma guía | No existe caso, pero banner ámbar discreto si pasara |
| Despacho Rápido | **Vista hermana** (no FAB). Rebrand a "⚡ Zona" |
| Cargadores | **100% independientes del preingreso**. Solo se agregan desde icono 🛺 del día |
| Origen cargadores | `PROVEEDORES_MASTER` de MOS, prefijo "CARGADOR" |
| Buscador inteligente + 🎤 | En toolbars de **Guías y Preingresos** (NO en detalle) |
| Búsqueda principalmente por | Nombre de proveedor + voz |
| Detalle de guía | Sin métricas, solo botones de acción rápida |
| Items en detalle | Orden de escaneo, más reciente arriba |
| Item layout | Nombre producto grande arriba · skuBase chico debajo |
| Agrupación canónico+equivalente | Colapsada por default, badge `▾Ncod` para expandir |
| Fotos preingreso | Mín 1 obligatoria · miniruleta en card · autoplay 4s + swipe |
| Fotos guía | Máx 1 · ícono pequeño · tap lightbox |
| Crear guía desde preingreso | Modal moderno ofrece: usar foto del preingreso / cámara / galería / sin foto |
| Storage de fotos | Drive (decisión Claude) — no base64 en Sheets |
| Día header | Mismo en Guías y Preingresos · `📋 N guías · 📥 N pre · 🛺 N` |
| Tap chip 🛺 | Abre modal cargadores (resumen + buscador) |
| Action sheet `[➕]` | Nueva guía / Nuevo preingreso / Cargador |
| Día agrupa | Sí, header por día en ambas vistas |
| Sonidos extra | 5 nuevos diferenciados, mismo toggle existente |
| Haptics | Diferenciados por estado de scan |
| Métricas (KPIs) | NO en ningún lado |

---

## 3. Modelo conceptual final

### 3.1 Las 3 vistas hermanas

```
┌───────────────────────────────────────────────────┐
│            MÓDULO GUÍAS (corazón de WH)           │
└───────────────────────────────────────────────────┘
        │                  │                 │
        ▼                  ▼                 ▼
   📋 GUÍAS          ⚡ ZONA             📥 PRE-INGRESOS
   (view-guias)     (view-zona,         (view-preingresos)
                     ex-despacho)
   Listado +        Zonas +             Recepción
   detalle +        pickup +            proveedor +
   acciones         lista + libre       fotos +
   rápidas          + cámara 3 estados  cargadores indep.
```

Bottom-nav: 3 tabs.

### 3.2 Día como unidad transversal

Mismo header `Día` en Guías y Preingresos:

```
══════════════════════════════════════════
  sábado · 9 mayo                  [➕]
  📋 6 guías · 📥 2 pre · 🛺 5
══════════════════════════════════════════
   │            │           │       │
   filtra día  scrolls      tap →   abre sheet
   solo        a pre del    modal   "Nueva guía /
               día          carga-  Nuevo preing /
                            dores   Cargador"
```

### 3.3 Cargadores — sistema independiente

- Catálogo: `PROVEEDORES_MASTER` de MOS, prefijo "CARGADOR".
- Log diario: nueva tabla `CARGADORES_LOG` en WH (idLog, idCargador, fecha, ts, addedBy, deviceId).
- Resumen del día computado.
- UI: modal con buscador arriba + lista resumen abajo (cargador + contador + [-]).
- Cada tap en resultado: +1 al cargador en resumen. Tap en [-]: -1.
- Independencia total de preingresos.

### 3.4 Fotos

| Entidad | Mín | Máx | UI |
|---|---|---|---|
| Preingreso | 1 | sin tope | Miniruleta autoplay 4s + swipe + dots · lightbox grid |
| Guía | 0 | 1 | Ícono cuadrado · lightbox fullscreen · placeholder `📷+` si vacío |

Almacenamiento: **Google Drive** en estructura `WH_FOTOS/<YYYY-MM>/<idEntidad>/<file>.jpg` + thumbnail 200×200. Sheet guarda solo `fileId`.

### 3.5 Op-log (la pieza crítica del problema #1)

Cada scan/edit es una **operación atómica** con `idOp` único, en cola persistente IndexedDB. Estados:
1. `escaneado` — dot azul fluo, UI optimista
2. `pending` — gris, sin red o esperando turno
3. `saving` — ámbar pulsante (solo si >600ms)
4. `saved` — verde, server respondió OK
5. `reconciled` — verde+✓, no parpadea

Retry exponencial: 1s · 3s · 9s · 27s · pausa. Si > 3s en saving → dot rojo + botón "Reintentar" + sonido warn.

**Diff-on-refresh:** al recargar, no sobreescribe — toast "🔄 +3 cambios desde el server · ver".

**Idempotency key** en crear-guía, cerrar-pickup, anular.

---

## 4. Wireframes finales

### 4.1 Vista Guías (móvil)

```
┌─────────────────────────────────────┐
│ [🌗]  WH · Almacén           🛒DSP  │
├─────────────────────────────────────┤
│ ⚙▾  🔎 Buscar proveedor…    🎤      │
├─────────────────────────────────────┤
│ ══════════════════════════════════  │
│   sábado · 9 mayo            [➕]   │
│   📋 6 guías · 📥 2 pre · 🛺 5      │
│ ══════════════════════════════════  │
│                                     │
│ ╭ guia-card moderna ──────────────╮ │
│ │ ┌──┐ ↓ INGRESO · ABIERTA 🔔     │ │
│ │ │📸│ Cabanossi · 18:30          │ │
│ │ └──┘ 12 items · ●●●●○ 80% sync  │ │
│ │     [💬][🖨]            [▸ ver] │ │
│ ╰─────────────────────────────────╯ │
│ ╭ guia-card ──────────────────────╮ │
│ │ ┌──┐ ↑ SALIDA · CERRADA          │ │
│ │ │📷+│ ME pickup · 17:02         │ │
│ │ └──┘ 4 items · ●●●●● ✓           │ │
│ ╰─────────────────────────────────╯ │
│                                     │
│ ══════════════════════════════════  │
│   viernes · 8 mayo                  │
│   📋 4 · 📥 1 · 🛺 3                │
│ ══════════════════════════════════  │
└─────────────────────────────────────┘
```

### 4.2 Vista Preingresos (móvil)

```
┌─────────────────────────────────────┐
│ ⚙▾  🔎 Buscar proveedor…    🎤      │
├─────────────────────────────────────┤
│ ══════════════════════════════════  │
│   sábado · 9 mayo            [➕]   │
│   📋 6 guías · 📥 2 pre · 🛺 5      │
│ ══════════════════════════════════  │
│ ╭ pre-card grande ────────────────╮ │
│ │ ┌────────────┐ Sin guía          │ │
│ │ │            │ Cabanossi · 18:00 │ │
│ │ │ [foto 1/3] │ S/. 1240 · ⚠Inc  │ │
│ │ │ ●  ○  ○    │ 🛺 Juan P.        │ │
│ │ └────────────┘ [💬][+Guía][▸]    │ │
│ ╰─────────────────────────────────╯ │
│ ╭ pre-card grande ────────────────╮ │
│ │ ┌────────────┐ Con guía ✓       │ │
│ │ │ [foto 1/1] │ Proveedor Y      │ │
│ │ │            │ S/. 890 · ✓Comp  │ │
│ │ └────────────┘ [💬][VerGuía][▸] │ │
│ ╰─────────────────────────────────╯ │
└─────────────────────────────────────┘
```

### 4.3 Sheet `[➕]` acción rápida

```
╭ ¿Qué quieres crear? ──────── ✕ ╮
│                                 │
│  📋  Nueva guía                 │
│  📥  Nuevo preingreso           │
│  🛺  Cargador del día           │
│                                 │
╰─────────────────────────────────╯
```

### 4.4 Modal cargadores

```
╭ 🛺 Cargadores · 9 mayo ──────── ✕ ╮
│                                    │
│ 🔎 [ Buscar cargador… ]       🎤   │
│                                    │
│ ── Coincidencias ─────────         │
│  • Juan Pérez                       │
│  • Carlos Mendoza                   │
│  • Andrea Ríos                      │
│                                    │
│ ════════════════════════════════   │
│ ── Resumen de hoy ── total 🛺 5 ── │
│                                    │
│  Juan Pérez          ×3      [-]   │
│  Carlos Mendoza      ×2      [-]   │
│                                    │
│  [✓ Listo]                         │
╰────────────────────────────────────╯
```

Comportamiento: tap resultado → +1 + `savedTick` + haptic 8ms. Tap `[-]` → -1 + toast undo 5s.

### 4.5 Detalle de guía

```
╭ Detalle · GUI-2026-0142 ──────── ✕ ╮
│                                     │
│  ┌────┐  ↓ INGRESO · ABIERTA 🔔     │
│  │📸  │  Cabanossi · 9 may 18:30   │
│  └────┘  ●●●●○ 80% sync             │
│                                     │
│  [📷] [🖨] [💬] [⤓] [🔒] [⚙]       │
│                                     │
│ ── Items · orden de scan ──         │
│ ╭───────────────────────────────╮   │
│ │ 🥬 Vinagre de arroz 500ml     │   │
│ │    LEV796 · KG                │   │
│ │ ████████  ×6  ●●●●● ▾2cod     │   │
│ ╰───────────────────────────────╯   │
│ ╭───────────────────────────────╮   │
│ │ 🧂 Sal de mesa fina           │   │
│ │    SAL123 · NIU               │   │
│ │ ████████  ×3  ●●●●●           │   │
│ ╰───────────────────────────────╯   │
╰─────────────────────────────────────╯
```

Tap `▾2cod` expande:
```
│ │    └─ 6959749711163 (canónico)  ×4 ●●●●● │
│ │    └─ EAN-NUEVO-001 (equiv)     ×2 ●●●●● │
```

### 4.6 Vista Zona (ex Despacho Rápido)

```
┌─────────────────────────────────────┐
│ [🌗]  ⚡ ZONA              🏠 ⚙    │
├─────────────────────────────────────┤
│  ─── Selecciona zona ───            │
│  ┌─────┐ ┌─────┐ ┌─────┐            │
│  │ Z-1 │ │ Z-2 │ │ Z-3 │            │
│  └─────┘ └─────┘ └─────┘            │
├─────────────────────────────────────┤
│  ─── Modo ───                       │
│  ◉ Pickup pendiente (1)             │
│  ◯ Lista cargada                    │
│  ◯ Libre (sin lista)                │
├─────────────────────────────────────┤
│  ┌─────────────────────────────┐    │
│  │   [video preview]           │    │
│  │   ╔═══════╗  🟢 LISTO        │    │
│  │   ║  ◯ mira║                │    │
│  │   ╚═══════╝                  │    │
│  │   [💡 luz]  [✕ cerrar]      │    │
│  └─────────────────────────────┘    │
│  Items escaneados: 4                │
└─────────────────────────────────────┘
```

3 estados cámara:
- 🟢 LISTO: borde verde glow, sin sonido bg
- ⏳ PROCESANDO: borde ámbar pulse, ping bajo opcional
- 🚫 BLOQUEADO: borde rojo opaco, buzz al intentar

### 4.7 Picker de foto al crear guía desde preingreso

```
╭ Foto de la guía ─────────── ✕ ╮
│                                │
│ ── Del preingreso ──           │
│  ┌────┐ ┌────┐ ┌────┐          │
│  │ 📸 │ │ 📸 │ │ 📸 │          │
│  └────┘ └────┘ └────┘          │
│                                │
│ ── O usa otra fuente ──        │
│  [📷 Cámara] [🖼 Galería]      │
│                                │
│ ── Sin foto ──                 │
│  [continuar sin foto]          │
│                                │
│  [Confirmar →]                 │
╰────────────────────────────────╯
```

### 4.8 Reabrir guía

```
╭ 🔒 Guía cerrada ──────── ✕ ╮
│                             │
│  Reabrir requiere clave     │
│  de administrador           │
│                             │
│  ┌─────────────────────┐    │
│  │ • • • • • • • •     │    │
│  └─────────────────────┘    │
│                             │
│  ☐ Recordar 30 min en       │
│     este dispositivo        │
│                             │
│  [Cancelar]  [Reabrir →]    │
╰─────────────────────────────╯
```

Si cierre fue < 5 min: bypass + toast "🕒 Reabriendo en gracia · 4:32 restantes".

### 4.9 Lightbox foto

```
╭ Foto ──────────────────── ✕ ╮
│                              │
│ ┌────────────────────────┐  │
│ │                        │  │
│ │   [foto fullscreen]    │  │
│ │   pinch-zoom · swipe   │  │
│ │                        │  │
│ └────────────────────────┘  │
│                              │
│  [📥 Descargar] [🗑]         │
╰──────────────────────────────╯
```

### 4.10 Banner multi-device

```
┌─────────────────────────────────────┐
│ Detalle · GUI-2026-0142             │
│ 📱 También editas en "Pixel 7" · 12s│
│ ●●●●○ sincronizando…                │
└─────────────────────────────────────┘
```

### 4.11 5 estados de sync por item

```
╭ ITEM CARD ────────────────────────────────────╮
│ 🥬 Vinagre de arroz 500ml                      │
│    LEV796 · KG                                 │
│ ████████░░░░░  ×6                              │
│                                                │
│ ● ● ● ○ ○   ← timeline                         │
│ │ │ │ │ └─ reconciled                          │
│ │ │ │ └─── saved                               │
│ │ │ └───── saving                              │
│ │ └─────── pending                             │
│ └───────── escaneado                           │
╰────────────────────────────────────────────────╯
```

---

## 5. Plan de implementación — 10 fases

### Fase 0 — Backend foundation

| Cambio | Archivo |
|---|---|
| Tabla `OPS_LOG` (idOp, idGuia, tipo, payload, ts, estado, deviceId) | `gas/Setup.gs` |
| Endpoint `aplicarOp` idempotente por `idOp` | `gas/Guias.gs` |
| Endpoint `listarOpsPendientes(idGuia)` | `gas/Guias.gs` |
| Endpoint `verificarClaveAdmin` (puente a MOS) | NEW `gas/Auth.gs` |
| Endpoint `listarCargadores` (filtra PROVEEDORES_MASTER prefijo CARGADOR vía bridge MOS) | NEW `gas/Cargadores.gs` |
| Endpoints `addCargadorDia` / `removeCargadorDia` / `getResumenDia` | `gas/Cargadores.gs` |
| Tabla `CARGADORES_LOG` | `gas/Setup.gs` |
| Endpoint `migrarFotosABase64Drive` (one-shot manual) | NEW `gas/Fotos.gs` |
| Endpoints `subirFoto`, `getFotoUrl`, `eliminarFoto` + thumbnails 200×200 | `gas/Fotos.gs` |

### Fase 1 — Migración fotos a Drive

| Acción | Detalle |
|---|---|
| Script migración | Lee fotos base64 PREINGRESOS / GUIAS · sube Drive · reemplaza con fileId |
| Estructura Drive | `WH_FOTOS/<YYYY-MM>/<idEntidad>/<file>.jpg` + `_thumb_<file>.jpg` |
| Validación | Re-leer 50 muestras antes/después · cero pérdida |

### Fase 2 — Visual base (lenguaje único)

| Componente | Archivo |
|---|---|
| CSS module "modern card v3" | `index.html` (style) |
| Card guía modernizada + foto icono | `index.html` + `js/app.js:_renderGuiaCard` |
| Card preingreso grande con foto carousel | `index.html` + `js/app.js:_renderCard` |
| Day header compartido | NEW helper `_renderDayHeader` |
| Bottom-sheet `[➕]` acción rápida | NEW sheet en `index.html` |

### Fase 3 — Sistema cargadores

| Componente | Archivo |
|---|---|
| Modal cargadores (buscador + resumen + contadores) | NEW sheet + NEW `js/cargadores.js` |
| Lectura PROVEEDORES_MASTER vía bridge MOS | `js/api.js` |
| Día header chip 🛺 N | `_renderDayHeader` |
| Polling de cargadores día actual | `js/app.js` |

### Fase 4 — Sistema fotos completo

| Componente | Archivo |
|---|---|
| Miniruleta pre-card (autoplay 4s + swipe + dots) | NEW `js/photos.js` |
| Lightbox fullscreen (pinch-zoom + swipe) | `js/photos.js` |
| Picker fuente al crear guía desde preingreso | NEW sheet + `js/photos.js` |
| Validación: pre ≥1 · guía ≤1 | `js/app.js` + `gas/Guias.gs` |
| Compresión client-side max 1600px | `js/photos.js` |

### Fase 5 — Op-log frontend (LA CLAVE)

| Componente | Archivo |
|---|---|
| Op-log persistente en IndexedDB | NEW `js/oplog.js` |
| Cola con retry exponencial 1s·3s·9s·27s·pausa | `js/oplog.js` |
| 5 dots de sync por item | `js/app.js:_renderItem` + CSS |
| Heartbeat header guía 🟢🟡🔴 con N en cola | `js/app.js` |
| Diff visual al recargar | `js/app.js` |
| Idempotency key en crear-guía + lock visual botón | `js/app.js:crearGuia` |
| Multi-device awareness (poll heartbeat) | `js/app.js` + `gas/Guias.gs` |
| Reemplazo `actualizarGuia` por ops `SCAN`, `EDIT_QTY`, `DELETE_ITEM`, `ANULAR` | `js/app.js` + `gas/Guias.gs` |

### Fase 6 — Despacho Zona (rebrand + reorganización)

| Componente | Archivo |
|---|---|
| Rebrand `view-despacho` → `view-zona` | `index.html` + `DespachoView`→`ZonaView` |
| Tabs internos: Pickup / Lista cargada / Libre | `js/app.js` |
| 3 estados cámara (LISTO/PROCESANDO/BLOQUEADO) | `js/app.js` + `js/scanner.js` |
| 5 sonidos nuevos: `scanReady`, `scanLocked`, `savedTick`, `saveRetry`, `saveLost` | `js/sounds.js` |
| Bottom-nav: 3 tabs Guías / Zona / Pre-ing | `index.html` |

### Fase 7 — Detalle de guía rediseñado

| Componente | Archivo |
|---|---|
| Header con acciones rápidas (sin métricas, sin buscador) | `index.html:sheetGuiaDetalle` + `js/app.js` |
| Items ordenados por scan (más reciente arriba) | `js/app.js:_renderItem` |
| Item con nombre grande / sku chico | CSS + `_renderItem` |
| Canónico+equivalente colapsado con badge `▾Ncod` | `_renderItem` + expand handler |
| Reabrir con clave admin + toggle "recordar 30 min" | NEW sheet + `gas/Auth.gs` bridge |
| Ventana 5 min gracia post-cierre | `js/app.js` (timestamp local) |

### Fase 8 — Buscador inteligente + 🎤

| Componente | Archivo |
|---|---|
| Wrapper Web Speech API | NEW `js/voice.js` |
| Botón 🎤 en toolbar Guías + Preingresos | `index.html` |
| Fuzzy match proveedor | `js/app.js` |
| Parser intención básica | `js/voice.js` |
| Feedback visual "escuchando" | CSS |

### Fase 9 — TTS post-cierre + ventana 5 min

| Componente | Archivo |
|---|---|
| Wrapper Speech Synthesis API | `js/voice.js` (mismo módulo) |
| TTS lee items: "6 vinagre, 3 sal, 2 aceite" | `js/app.js:cerrarGuia` |
| Botón "▶ Reproducir nuevo" / "🔁 Repetir" | `index.html` |
| Banner "🕒 Puedes reabrir sin clave: 4:32 restantes" | `js/app.js` |

### Fase 10 — Pulido final

| Componente | Archivo |
|---|---|
| Sincronizar sonido + haptic + visual <200ms | `js/app.js` |
| Reemplazar `confirm()` nativos por bottom-sheets con undo | `js/app.js` |
| Modo nocturno auto (ambient light) | `index.html` + `js/app.js` |
| Pull-to-refresh = forzar reconciliación op-log | `js/app.js` |
| Bump VERSION en `sw.js` + `version.json` | (siempre al final) |

---

## 6. Archivos a tocar — vista panorámica

```
warehouseMos/
├── index.html              · ALTO (CSS + nuevos sheets + estructura views)
├── js/
│   ├── app.js              · ALTO (refactor GuiasView/PreingresosView/ZonaView)
│   ├── sounds.js           · BAJO (+5 sonidos)
│   ├── scanner.js          · MEDIO (3 estados cámara)
│   ├── offline.js          · MEDIO (integra oplog)
│   ├── oplog.js            · NUEVO
│   ├── cargadores.js       · NUEVO
│   ├── photos.js           · NUEVO
│   ├── voice.js            · NUEVO (STT + TTS)
│   └── api.js              · MEDIO (nuevos endpoints)
├── gas/
│   ├── Guias.gs            · ALTO (aplicarOp, listarOpsPendientes, idempotency)
│   ├── Setup.gs            · MEDIO (tablas OPS_LOG + CARGADORES_LOG)
│   ├── Auth.gs             · NUEVO (bridge clave admin a MOS)
│   ├── Cargadores.gs       · NUEVO
│   └── Fotos.gs            · NUEVO (Drive ops + thumbnails)
├── sw.js                   · bump VERSION al final
└── version.json            · bump al final
```

---

## 7. Dependencias entre fases

```
F0 ──┬─→ F1 ──→ F4 ───────────────────────┐
     │                                     │
     ├─→ F2 ──→ F3 ──┐                     │
     │              ├──→ F5 ──┬─→ F6 ──┐   │
     └──────────────┘         │        ├──→ F7 ──→ F8 ──→ F9 ──→ F10
                              └────────┘

F0  Backend foundation (op-log endpoints, auth bridge, cargadores, fotos)
F1  Migración fotos a Drive (one-shot)
F2  Visual base (cards modernas, day header, sheet acción rápida)
F3  Cargadores end-to-end
F4  Sistema fotos (carousel, lightbox, picker)
F5  Op-log frontend ← LA CLAVE del problema #1
F6  Despacho Zona rebrand + 3 estados cámara + 5 sonidos
F7  Detalle de guía rediseñado + reabrir admin
F8  Buscador 🎤 en listas
F9  TTS post-cierre + ventana 5 min
F10 Pulido + bump versión
```

---

## 8. Pregunta pendiente (única antes de arrancar)

¿Por dónde empezar?

**Opción A — F0+F5 primero (recomendada por Claude):**
- El operador no nota cambio visual inmediato.
- Los reclamos de "se borró mi guía" desaparecen desde el día 1.
- Tarda 1-2 semanas en mostrar valor visual.

**Opción B — F0+F2 primero:**
- Cards bonitas desde el día 1.
- El bug de pérdida persiste hasta F5.
- "Ya se ve mejor" rápido pero deuda funcional.

**El usuario debe decidir.** No tocar código hasta confirmación.

---

## 9. Reglas duras a respetar

1. **codigoBarra siempre texto** — formato '@' a nivel columna + `String()` antes de escribir. NUNCA número en Sheets.
2. **Regla WH canónico+equivalente** — WH solo maneja canónicos (factor=1) y equivalentes activos. Guías SIEMPRE registran `codigoBarra` real, nunca `skuBase`. Stock descuenta por `codigoBarra` específico. (Ver `architecture_wh_codigos_canonico_equivalente.md` en memoria.)
3. **Verificación de dispositivo bloqueante** — endpoint `consultarEstadoDispositivo` debe completarse antes de cargar app.
4. **Bump VERSION** al final de cada cambio frontend de WH (en `sw.js` + `version.json`).
5. **No achicar guías ni preingresos** — son 10 y 1-3 por día respectivamente, hay espacio.
6. **Fotos preingreso ≥1, fotos guía ≤1** — validar tanto en frontend como en GAS.
7. **Cargadores independientes** — NUNCA enganchar el flujo cargador al preingreso. Es solo desde icono 🛺 del día.
8. **Sin métricas** — ni tiempo de apertura, ni KPIs, ni rankings. Solo acciones rápidas.
9. **Buscador en listas, no en detalle.**
10. **TTS lee solo cantidad + nombre producto. Nada más.**

---

## 10. Estado de avance — 11 fases completas

- [x] F0 · Backend foundation — Auth.gs, Cargadores.gs, Mermas.gs, Fotos.gs, OpLog.gs · tablas OPS_LOG + CARGADORES_LOG · getRolUsuario · getProductosCambiadosDesde
- [x] F1 · Migración fotos Drive — `migrarFotosABase64Drive()` en Fotos.gs (one-shot)
- [x] F2 · Visual base — CSS modules + sheets HTML (action sheet, type picker, subtype picker, entidad picker, modales fotos/cargadores/mermas) + App helpers + FAB pill ahora abre action sheet + day-header con [+] y count chip
- [x] F2.5 · Sistema mermas — cesta + agregar + solucionar (sliders) + procesar eliminación (clave admin) + ícono topbar 🗑 + badge + polling 90s
- [x] F3 · Sistema cargadores — modal moderno + buscador + 🎤 + filtro prefijo CARGADOR + log idempotente add/remove
- [x] F4 · Sistema fotos — lightbox fullscreen + carousel autoplay + picker fuente (preingreso/cámara/galería) + compresión client-side 1600px
- [x] F5 · Op-log frontend — IndexedDB persistente + retry exponencial (1s·3s·9s·27s) + multi-device id estable + 13 tipos de op soportados → **resuelve problema #1**
- [x] F6 · Camera states + sonidos — 5 estados de cámara aplicados en `_setScanStatus` (listo/procesando/preguntando/descubriendo/bloqueado) + 8 sonidos nuevos (scanReady, scanLocked, savedTick, saveRetry, saveLost, scanIncompleto, scanNuevo, productoVerificado)
- [x] F7 · Detalle + caso 2/3 — Reabrir con clave admin + recordar 30 min + ventana 5 min gracia + caso 2 (prefijo) ahora con sonido scanIncompleto + caso 3 (nuevo) auto-dispara scanNuevo en INGRESO + camPicker incluye "Ninguno · es otro código" que abre modal PN
- [x] F8 · Buscador + 🎤 — `Voice` wrapper + 🎤 en inputBuscarGuia + inputBuscarPre + dictado modal cargadores + dictado entidad picker
- [x] F9 · TTS post-cierre — `Voice.leerItems` se dispara automáticamente tras cerrarGuia exitoso (cantidad + nombre producto) + `App._registrarCierre` marca ventana 5 min gracia local
- [x] F10 · Pulido — Pull-to-refresh dispara `OpLog.flush()` + refresh badge mermas + bump VERSION 2.0.4 → 2.1.1

## 10.1. Trabajo opcional para más adelante

1. Modernizar `_renderGuiaCard` y `_renderCard` con foto icono grande y carousel embedded (CSS ya disponible).
2. Rebrand `view-despacho` → `view-zona` con tabs Pickup/Lista/Libre.
3. Item card en detalle con mini-badge tipo match (✓/↕E/↕C/🆕) y agrupación canónico+equiv `▾Ncod`.
4. Reemplazar los 9 `confirm()` nativos restantes por bottom-sheets con undo de 5s.
5. Diff visual al recargar guía (toast "+3 cambios desde server").
6. Multi-device awareness banner (poll heartbeat per idGuia + deviceId).
7. Modo nocturno auto con `ambient-light-sensor`.

Estos no son críticos para el funcionamiento — el sistema queda completamente operativo con lo entregado.

> Actualizar este bloque a medida que se cierre cada fase.

---

## 11. Casos de escaneo en INGRESO (los 4 que el usuario verificó)

### 11.0 — Regla consolidada por tipo de guía

| Tipo | Exacto | Prefijo | Nuevo |
|---|:---:|:---:|:---:|
| ↓ INGRESO · Proveedor | ✓ | ✓ | ✓ |
| ↓ INGRESO · Jefatura | ✓ | ✓ | ✓ |
| ↓ INGRESO · Devolución de zona | ✓ | ✓ | ✓ |
| ↑ SALIDA · Despacho zona | ✓ | ✗ | ✗ |
| ↑ SALIDA · A jefatura | ✓ | ✗ | ✗ |
| ↑ SALIDA · Devolución a proveedor | ✓ | ✗ | ✗ |
| ⚙ MERMA | ✓ | ✗ | ✗ |
| 📊 AJUSTE +/− | ✓ | ✗ | ✗ |

**Regla simple para el operador:**
- Si **algo entra** al almacén → puedes resolver código incompleto y registrar nuevo.
- Si **algo sale o se ajusta** → el producto debe existir ya, sí o sí.

**Justificación MERMA / AJUSTE:** el catálogo crece solo por la puerta de INGRESO. Si en conteo encuentras producto desconocido, hacer guía de ingreso (no ajuste).

### 11.1 — Caso 1: código exacto (canónico o equivalente)
- Aplica a **TODOS los tipos** (ingreso, salida, merma, ajuste).
- Buscar en lista canónicos + equivalentes activos.
- Match directo → suma al item.
- Item card muestra mini-badge tipo de match:
  - `✓` directo canónico
  - `↕E` fue equivalente
- Sonido: `beep` · Haptic: 15ms.

### 11.2 — Caso 2: código incompleto (prefijo)
- Operador escanea `12345`. Existen en catálogo `12345A`, `12345B`.
- WH detecta prefijo, abre modal moderno con candidatos.
- Cada candidato muestra: codigoBarra completo, skuBase, nombre, unidad, factor.
- Operador tap → se registra el `codigoBarra` REAL (`12345A`) en GUIA_DETALLE.
- Item card badge: `↕C` (completado desde prefijo).
- Si tap "Ninguno · es otro código" → abre flujo Caso 3.
- **Aplica en TODAS las guías de INGRESO (proveedor, jefatura, devolución de zona).**
- **NO aplica en SALIDA, MERMA ni AJUSTE.** En esos casos, prefijo dispara error duro.
- Sonido: `scanIncompleto` (NUEVO).
- Haptic: pulso curioso 30ms.
- Visual: borde cámara → ámbar pulsante.

### 11.3 — Caso 3: código nuevo (no existe ni exacto ni prefijo)
- Modal moderno: "¿registrar como nuevo?"
- Campos: código (auto-bloqueado del scan), descripción (con 🎤 dictado), unidad, factor, categoría, foto opcional.
- Optimista: se agrega al detalle de guía con badge `🆕 PENDIENTE APROBACIÓN`.
- Operador puede seguir escaneando sin esperar.
- En paralelo: WH crea el producto + request a MOS para aprobación.
- **Aplica en TODAS las guías de INGRESO (proveedor, jefatura, devolución de zona).**
- **NO aplica en SALIDA, MERMA ni AJUSTE.** En esos casos, código inexistente → error duro con sugerencia: "este código no existe en el catálogo. Si lo encontraste físicamente, regístralo con una guía de INGRESO."
- Sonido: `scanNuevo` (NUEVO).
- Haptic: doble medio [40, 30, 40].
- Visual: borde cámara → violeta "modo descubrimiento".

### 11.4 — Caso 4: reconciliación cuando MOS aprueba
- WH polling consulta MOS por productos pendientes aprobados.
- 4 sub-casos:
  - **4.A** Aprobado tal cual: badge `🆕` → `✓` con animación flip.
  - **4.B** Aprobado con corrección de nombre: nombre cambia con scramble→reveal. Badge `✓ CORREGIDO`.
  - **4.C** Identificado como equivalente de canónico existente: skuBase apunta al canónico, nombre muestra el canónico, si ya existía item del canónico en la guía → se agrupan con `▾Ncod`. Badge `↕E APROBADO`.
  - **4.D** Prefijo aprobado: transparente, no requiere reconciliación (ya se guardó el codigoBarra real).
- Sonido: `productoVerificado` (NUEVO).
- Visual: card brilla verde 800ms, badge animación flip, toast con detalle.

### 11.5 — Mini-badges de tipo de match en item card

| Badge | Significado |
|---|---|
| `✓` | Match directo en canónico |
| `↕E` | Match en equivalente |
| `↕C` | Completado desde prefijo |
| `🆕` | Nuevo pendiente aprobación |
| `✓ CORREGIDO` | Aprobado con cambio de nombre |
| `↕E APROBADO` | Vinculado a canónico existente por admin |

### 11.6 — Estados extra de la cámara (5 totales, antes 3)

| Estado | Color borde | Sonido | Cuándo |
|---|---|---|---|
| 🟢 LISTO | verde glow | sin sonido bg | esperando scan |
| ⏳ PROCESANDO | ámbar pulse | ping bajo opcional | resolviendo match |
| ↕ PREGUNTANDO | ámbar fuerte | scanIncompleto | modal prefijo abierto |
| 🆕 DESCUBRIENDO | violeta | scanNuevo | modal producto nuevo abierto |
| 🚫 BLOQUEADO | rojo opaco | buzz al intentar | modal qty granel u otro abierto |

### 11.7 — Cambios al plan derivados

**Sonidos nuevos (total 7, antes 5):**
- `scanReady` · `scanLocked` · `savedTick` · `saveRetry` · `saveLost`
- **+** `scanIncompleto` · `scanNuevo` · `productoVerificado`

**Modales modernos nuevos:**
- Modal prefijo (caso 2) con lista candidatos
- Modal producto nuevo (caso 3) con 🎤 dictado + foto opcional

**Backend GAS nuevo:**
- Endpoint WH `consultarProductosPendientesAprobacion`
- Endpoint MOS bridge `getProductosCambiadosDesde(ts)` para polling
- Sync inverso: cuando MOS cambia un producto, WH actualiza items abiertos

**Fases impactadas:**
- F0 — agregar endpoints aprobación + polling
- F5 — agregar reconciliación inversa como tipo de op
- F6 — agregar 2 estados cámara más (PREGUNTANDO, DESCUBRIENDO)
- F7 — modales prefijo y producto nuevo modernos
- Sonidos: pasan de 5 a 7 nuevos en `js/sounds.js`

---

## 12. Tipos de guías y selector visual

### 12.1 — Los 6 tipos de guía creables

Patrón simétrico: 3 entidades × 2 direcciones.

| Dirección | Entidad | Origen del catálogo |
|---|---|---|
| ↓ INGRESO | 🛺 Proveedor | PROVEEDORES_MASTER de MOS (excluir prefijo CARGADOR) |
| ↓ INGRESO | 👤 Jefatura | PERSONAL de MOS (roles admin + master) |
| ↓ INGRESO | 📍 Devolución de zona | ZONAS (zonas son "clientes" del almacén) |
| ↑ SALIDA | 📍 Despacho a zona | ZONAS |
| ↑ SALIDA | 👤 A jefatura | PERSONAL de MOS (roles admin + master) |
| ↑ SALIDA | 🛺 Devolución a proveedor | PROVEEDORES_MASTER de MOS (excluir CARGADOR) |

### 12.2 — MERMAS: separación, no guía

**MERMAS no son una guía**, son una **tabla de separación** del stock real.

Flujo completo:
1. Operador detecta problema (vencido, dañado, etc.) → agrega a cesta de mermas (no descuenta stock todavía).
2. La merma queda en estado `PENDIENTE` con responsable=zona.
3. Operador puede **solucionar**:
   - Recuperar (vuelven al andamio = al stock real).
   - Descartar (quedan acumuladas esperando proceso).
   - Mixto: ej 4 aceites → 2 recupera, 2 descarta.
4. Cuando los `DESCARTADO` acumulan, alguien procesa eliminación → genera UNA guía SALIDA agrupada con todos los descartados.
5. Las mermas pasan a `ELIMINADO` con `idGuiaSalida` apuntando a la guía generada.

**Tabla `MERMAS`:** idMerma, codigoBarra, skuBase, cantidadOriginal, recuperado, descartado, pendiente, estado, zonaResponsable, motivo, foto, idGuiaSalida, timestamps.

**Estados:** `PENDIENTE` · `SOLUCIONADO_PARCIAL` · `SOLUCIONADO_TOTAL` · `DESCARTADO_TOTAL` · `ELIMINADO`.

**Acceso:** ícono 🗑 en topbar global (no en day header — mermas se acumulan entre días) con badge rojo si hay pendientes.

**Scan en mermas:** solo exacto (canónico + equivalente). Mermas no crean productos nuevos.

**Vista cesta:** bottom-sheet 95vh con 3 secciones:
- Pendientes (sin decisión).
- Descartado (esperando proceso).
- Solucionado (historial últimos 7 días).

**Modales modernos:**
- "Agregar a mermas" (scan + cantidad + zona + motivo + foto opcional).
- "Solucionar" (sliders recuperar/descartar, suma debe igualar original).
- "Procesar eliminación" (genera guía SALIDA agrupada).

### 12.3 — ENVASADOS: 2 guías auto-generadas bloqueadas

Cada operación de envasado dispara **automáticamente** 2 guías:
- ↑ SALIDA · ENVASADO · −X del granel.
- ↓ INGRESO · ENVASADO · +Y de la presentación.

Ambas con `estado: BLOQUEADA`, color gris, read-only. Aparecen en la lista de Guías para consulta/auditoría pero no se editan. Se crean desde la estación de envasado (otro módulo, no este).

### 12.4 — AJUSTES: tabla aparte fuera del picker

**Stock = Σ INGRESOS − Σ SALIDAS + Σ AJUSTES.**

Los ajustes (+/−) vienen de auditoría/conteo (otro módulo). No se crean desde el picker de Guías.

**Devolución de cliente = devolución de zona** (las zonas son los clientes del almacén). No es un tipo aparte.

**Consumo interno y traslado entre almacenes:** archivados, no aplican.

### 12.5 — Picker rediseñado: solo 2 cards grandes

```
╭ ¿Qué movimiento haces? ── ✕ ╮
│                              │
│  ┌──────────────────────┐    │
│  │       ↓              │    │
│  │   INGRESO            │    │ ← verde, gradient
│  │   recibo al almacén  │    │
│  └──────────────────────┘    │
│                              │
│  ┌──────────────────────┐    │
│  │       ↑              │    │
│  │   SALIDA             │    │ ← azul, gradient
│  │   saco del almacén   │    │
│  └──────────────────────┘    │
╰──────────────────────────────╯
```

Más simple, más rápido de decidir con guantes.

Cesta de mermas se accede desde topbar (🗑), no desde este picker.

**Paso 2** (sub-tipo de entidad) sin cambio: proveedor / jefatura / zona según dirección.

### 12.6 — Selectores de entidad modernos

- **Proveedor**: buscador + "Más usados (este mes)" + "Todos" + botón `+ Nuevo`. Filtro automático excluye `CARGADOR`.
- **Jefatura**: solo PERSONAL con rol admin/master. Buscador + 🎤.
- **Zona**: botones grandes tap-friendly. ME aparece como zona especial.

### 12.7 — Shortcuts por contexto

- Action sheet del día tiene **shortcut directo "⚡ Despacho a zona"** (lo más frecuente → 2 taps).
- **Vista ⚡ Zona** del bottom-nav asume directamente "despacho a zona" — sin picker.

### 12.8 — Filtro por rol (anti-confusión)

| Rol | Puede crear |
|---|---|
| operador zona | Despacho a zona · Devolución de zona |
| operador almacén | Todos los INGRESO · Despacho a zona · Merma |
| admin / master | Todos |
| jefe de turno | Configurable |

Los tipos no permitidos no aparecen en el picker.

### 12.9 — Visualización de tipo en cards

```
Border-izq por dirección + ícono por sub-tipo:

↓ INGRESO  · 🛺 proveedor      verde + mototaxi
↓ INGRESO  · 👤 jefatura       verde + persona
↓ INGRESO  · 📍 devol zona     verde + ubicación
↑ SALIDA   · 📍 despacho zona  azul + ubicación
↑ SALIDA   · 👤 jefatura       azul + persona
↑ SALIDA   · 🛺 devol proveed  azul + mototaxi
↑ SALIDA   · 🗑 merma procesada azul + cesta (auto al procesar)
⏚ ENVASADO                     gris (read-only, auto-generado)
```

Nota: MERMA y AJUSTE no aparecen como tipo de guía. Mermas viven en su propia tabla; cuando se procesan, generan una guía SALIDA con sub-tipo "merma procesada".

Color = dirección · ícono = entidad/origen. Operador entiende sin leer letra.

### 12.10 — Reglas anti-confusión del selector

1. Solo 4 cards en paso 1. No más.
2. Iconos grandes + 1 línea de descripción.
3. Botón "← Volver" siempre visible en paso 2.
4. Estado vacío con CTA (no listas mudas).
5. Sin pantalla de confirmación final. Chip clickeable en header del sheet permite cambiar tipo.
6. Memoria del último uso por operador → ordena "Más usados".
7. Voz desde paso 1: "ingreso de cabanossi" salta los 2 pasos.

### 12.11 — Cesta de mermas en topbar

```
┌─────────────────────────────────────────────┐
│ [🌗]  WH · Almacén       🗑 23  🛒DSP        │
│                            │                 │
│                            └─ tap → cesta    │
│                               badge rojo si  │
│                               hay pendientes │
└─────────────────────────────────────────────┘
```

- Pendientes = 0 → ícono gris, sin badge.
- Pendientes > 0 → ícono blanco + badge rojo con N.
- Pendientes > threshold (ej 50) → badge pulsa ámbar.

Acceso global, no per-día.

### 12.12 — Cambios al plan derivados

**Backend GAS nuevo:**
- Endpoint `getRolUsuario(deviceId|email)` para filtrar tipos disponibles.
- Endpoint `getMasUsados(usuario, tipo)` para ordenar selectores por frecuencia.
- Validación server-side: el tipo declarado debe ser permitido al rol del usuario.

**Frontend nuevo:**
- Nuevo componente `GuiaTypePicker` (paso 1 + paso 2).
- Nuevo componente `EntidadPicker` (proveedor / jefatura / zona).
- Refactor del `sheetGuia` actual para recibir `tipo + entidad` pre-poblados.
- Comando de voz "crear ingreso de X" / "crear salida a Y" pasa por `js/voice.js`.

**Fases impactadas:**
- F0 — agregar endpoints rol + más usados + tabla MERMAS + endpoints merma.
- F2 — agregar `GuiaTypePicker` (2 cards) y `EntidadPicker` al sistema visual base.
- F2.5 (NUEVA) — sistema completo de mermas: tabla, ícono topbar, cesta sheet, modales agregar/solucionar/procesar.
- F5 — ops nuevas: `MERMA_AGREGAR`, `MERMA_SOLUCIONAR`, `MERMA_PROCESAR_ELIMINACION`.
- F8 — agregar parser de voz multi-paso ("ingreso de cabanossi", "agregar a mermas X").

### 12.13 — Fase nueva: F2.5 Sistema Mermas

| Componente | Archivo |
|---|---|
| Tabla `MERMAS` en GAS Setup | `gas/Setup.gs` |
| Endpoints `agregarMerma`, `solucionarMerma`, `procesarEliminacionMermas`, `listarMermas`, `contadorMermasPendientes` | NEW `gas/Mermas.gs` |
| Ícono 🗑 + badge en topbar global | `index.html` |
| Bottom-sheet "Cesta de mermas" 95vh con 3 secciones | NEW sheet en `index.html` |
| Modal "Agregar a mermas" | NEW sheet en `index.html` |
| Modal "Solucionar" con sliders recuperar/descartar | NEW sheet en `index.html` |
| Modal "Procesar eliminación" → genera guía SALIDA agrupada | NEW sheet en `index.html` |
| Lógica frontend cesta | NEW `js/mermas.js` |
| Polling de contador mermas pendientes (cada 60s) | `js/app.js` |
| Integración con op-log (3 ops nuevas) | `js/oplog.js` |
| Stock update inmediato al recuperar (op `MERMA_SOLUCIONAR` con delta+) | `gas/Mermas.gs` + `js/oplog.js` |

---

## 13. Ideas extra para considerar más adelante

1. Drag & drop foto desde escritorio en tablet/desktop.
2. Tab comparador foto guía vs fotos preingreso.
3. Comparador visual de cargas del cargador (mes, semana).
4. Re-OCR del ticket impreso vía QR para reabrir guía.
5. Marker "este preingreso ya tiene guía" clickable que navega al detalle.
6. Cargadores agrupados en el chip del día con conteo individual.
7. Modo "vista de auditoría de ops" escondida para diagnóstico.
8. Pre-warm de estado al volver a la app (reconciliar pendientes en background).
9. Notificación push a admin si una op queda en `lost` > 5 min.
10. Replay de la última guía cerrada (botón ▶ continuar guía cerrada 5 min) — ya cubierto por la ventana de gracia.

---

**Fin del plan. Última actualización: antes del primer commit. Próximo paso: confirmación del usuario sobre Opción A / B en sección 8.**
