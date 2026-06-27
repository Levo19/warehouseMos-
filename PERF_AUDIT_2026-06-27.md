# WH — Auditoría de lentitud + cero-GAS (2026-06-27)

Auditoría multi-agente (56 agentes, 6 hot-paths, verificación contra código + DB en vivo).
**37 hallazgos: 8 CRÍTICOS · 12 HIGH · 13 MED · 4 LOW · 16 fugas GAS.**

## Causa raíz (resumen en una frase)
Todo converge en **dos** causas:
- **(A) `mos.catalogo_wh_rls` re-descarga ~1.9–2.2 MB COMPLETOS en CADA bump de `catalogo_version`** (que sube decenas/cientos de veces al día — 9 triggers, varios irrelevantes para WH). Hoy bumpeé la versión ~170 veces → cascada de re-descargas.
- **(B) El arranque/login dispara ~8–13 RPCs pesadas en paralelo** (catálogo 1.9MB + guia_detalle 1.68MB + getDashboard 7 + descargarOperacional 6) → **encolan** en el pool PostgREST → la cola es lo que infló `catalogo_wh_rls` de ~1.2s a los **22.46s observados** (la query en DB es rápida).

## ✅ Hecho en esta tanda (WH 2.13.354) — máximo alivio, bajo riesgo
- **COALESCING de la re-descarga**: una ráfaga de bumps → **1 sola descarga diferida** (ventana 20s, toma la versión más alta), en vez de re-bajar 1.9MB por cada bump. *(El #1 alivio: mata la cascada.)*
- **THROTTLE del chequeo de versión**: foco/visibility/timer no chequean más de 1×/30s.
- **TIMEOUT 12→25s** en RPC pesadas: ya no aborta a GAS bajo cola (que pagaba directo-abortado + cold-start GAS).
- **PIN no bloqueante**: el teclado ya no espera la descarga de ~1.9MB (validación es server-side).
- **GAS `registrarUbicacion`**: fire-and-forget real + timeout 8s (no cuelga conexión ~18s).

## ⏳ Pendiente (orden de impacto/esfuerzo)
### Estructural (la causa raíz de fondo)
1. **CD1 — Catálogo DELTA**: RPC `mos.catalogo_wh_delta(desde_version)` que devuelva solo filas cambiadas + tombstones; el front mergea. Elimina >90% del peso. *(El fix de fondo.)*
2. **CD2 — Adelgazar `catalogo_wh_rls`**: proyectar solo las 32 columnas del spec (`_CAT_SPECS_LEC`) + filtrar inactivos (hoy hace `to_jsonb` de la fila completa: created_at/updated_at viajan de gratis, ~150KB). Server-side, sin tocar el front.
3. **CD3 — Filtrar triggers de `catalogo_version`**: que precio_tramos/proveedores_productos (que WH no usa) NO disparen re-descarga.
4. **QW2 — Escalonar el arranque**: limitar concurrencia (2–3 RPC), diferir getDashboard/operacional ~500ms tras el primer render, o 1 RPC agregadora `wh_arranque`.
5. **CD4 — Lazy `guia_detalle_operacional`** (60d=1.68MB cada 60s): bajar a 7–14d + detalle on-demand al abrir cada guía.

### Fugas GAS a migrar a Supabase (D1: cero-GAS)
- **G1 (HIGH)** `BloqueoRemoto._check` → GET GAS `getEstadoBloqueoUsuario` cada **120s todo el turno**. Migrar a RPC/Realtime. *(La peor recurrente — consume cuota urlfetch que frena los POST de operaciones.)*
- **G5 (HIGH)** verificar clave admin (REABRIR_GUIA) → GAS con await sin timeout (~18s congela el modal). Usar `mos.verificar_clave_admin` (ya existe).
- **G2** `registrarUbicacion` → migrar a RPC Supabase (ya endurecido, falta migrar).
- **G3** `notificarInicioSesionVendedor`, **G4** `getAdminPinsCache`, **G6** push doble-write (Supabase YA + GAS redundante), **G7** `detenerEscuchaAudio`.

Detalle completo: workflow `w2y0h5j1c` output.

## ✅ Lo de fondo COMPLETADO (WH 2.13.355/356 + SQL 277)
- **CD1 Catálogo DELTA** (la causa raíz): `mos.catalogo_wh_delta(desde)` → solo productos cambiados + tablas chicas + `eliminados` (tombstones) + `server_ts`. Un bump típico baja **~42KB vs 1.9MB (97% menos)**. Front: `_refrescarCatalogoDelta` mergea por idProducto sobre copia, quita borrados, fallback a full. Trigger `updated_at` (INSERT+UPDATE) + `mos.catalogo_tombstones` (AFTER DELETE).
- **CD2** full sin created_at/updated_at. **G6** push GAS solo-fallback.

## 🔍 500x del delta (workflow `wm12v8yph`) — 8 hallazgos, 3 HIGH, TODOS corregidos
- **HIGH deletes**: producto borrado quedaba en el cache WH → tombstones + `delta.eliminados` + filtro en el merge. ✅
- **HIGH ventana de borde**: `>` + server_ts=now() podía perder un cambio de precio para siempre → `>=` + margen -2s (solape idempotente). ✅
- **HIGH baseline**: avanzaba aunque la descarga fallara (catálogo stale indefinido) → `precargar`/`_refrescarCatalogoDelta` devuelven `{ok}`; baseline avanza SOLO si ok. ✅
- **MED** aliasing del parse-cache → `.slice()`. **LOW** tablas chicas → `_guardarSiCambia`. ✅

## ⏳ Migraciones GAS pendientes (pasada dedicada — tocan auth/login)
G1 heartbeat 120s, G2 registrarUbicacion, G4 getAdminPinsCache, G5 admin-PIN reabrir-guía → migrar a RPC Supabase con fallback GAS. G3, G7 menores. (G6 ya hecho.)
También pendiente: QW2 escalonar arranque (limitar concurrencia), CD4 lazy guia_detalle_operacional.
