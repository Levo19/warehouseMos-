/**
 * ============================================================
 * MIGRACIÓN SUPABASE — FASE 1 · Backfill de warehouseMos (esquema wh)
 * ============================================================
 * Vive en el GAS de warehouseMos. Requiere:
 *   - Supabase.gs (helper _sb) copiado aquí.
 *   - Script Properties: SUPABASE_URL, SUPABASE_SERVICE_KEY (legacy JWT eyJ…).
 *   - Haber corrido 01_schema_compartido.sql y 03_schema_wh.sql + exponer esquema `wh`.
 *   - SPREADSHEET_ID (ya configurado en producción).
 *
 * Características (heredadas del patrón ME, probado):
 *   - REANUDABLE: checkpoint por tabla en Script Properties (límite 6 min GAS).
 *   - `linea` DETERMINISTA para guia_detalle / pedidos_cliente_adj (orden de hoja).
 *   - Idempotente: upsert por clave natural (on_conflict).
 *   - Lectura CRUDA (_whSheetRows) → preserva timestamps completos (no los aplasta a fecha).
 *   - dryRun para validar headers sin escribir.
 *
 * Uso (desde el editor):
 *   dryRunWH()            // valida headers/mapeo, NO escribe
 *   backfillWH()          // backfill real (re-correr hasta que todo diga ok:true)
 *   verificarCuadreWH()   // compara conteos sheet vs supabase
 *   resetCheckpointsWH()  // borra checkpoints (para reempezar limpio)
 *
 * Diagnóstico (solo lectura de hojas):
 *   dumpHeadersWH() · inspeccionarTodoWH() · inspeccionarRestoWH() · inspeccionarRestoWH2() · chequearPKsWH() · _sbPingWH()
 */

// ---------- conversores defensivos ----------
function _whText(v){ return (v==null||v==='')?null:String(v); }
function _whNum(v){ if(v==null||v==='')return null; if(typeof v==='number')return isNaN(v)?null:v; var s=String(v).trim(); if(s.charAt(0)==='#')return null; var n=parseFloat(s.replace(',','.')); return isNaN(n)?null:n; }   // "#NUM!"→null
function _whInt(v){ var n=_whNum(v); return n==null?null:Math.round(n); }
function _whDate(v){ if(v==null||v==='')return null;
  if(!(v instanceof Date)){
    var sv=String(v).trim();
    // date-only STRING → new Date lo lee como UTC y al formatear en Lima cae al día anterior. Anclar a medianoche Lima.
    if(/^\d{4}-\d{2}-\d{2}$/.test(sv)) v=sv+'T00:00:00-05:00';
    // dd/MM/yyyy (formato peruano de fechas OCR ingresadas a mano) → ISO anclado a Lima.
    else if(/^\d{2}\/\d{2}\/\d{4}$/.test(sv)){ var p=sv.split('/'); v=p[2]+'-'+p[1]+'-'+p[0]+'T00:00:00-05:00'; }
  }
  var d=(v instanceof Date)?v:new Date(v); if(isNaN(d.getTime()))return null; return Utilities.formatDate(d,'America/Lima',"yyyy-MM-dd'T'HH:mm:ssXXX"); }
function _whHora(v){ if(v==null||v==='')return null; if(v instanceof Date)return Utilities.formatDate(v,Session.getScriptTimeZone(),'HH:mm:ss'); return String(v); }   // time-serial 1899 → HH:mm:ss en TZ del script (coherente con _sheetToObjects)
function _whBool(v){ if(v==null||v==='')return null; if(typeof v==='boolean')return v; var s=String(v).trim().toLowerCase(); if(s==='true'||s==='1'||s==='si'||s==='sí'||s==='verdadero'||s==='x')return true; if(s==='false'||s==='0'||s==='no'||s==='falso')return false; return null; }
function _whJson(v){ if(v==null||v==='')return null; if(typeof v==='object')return v; try{var p=JSON.parse(String(v)); return (p&&typeof p==='object')?p:null;}catch(e){return null;} }

function _whVal(raw,t){
  if(t==='text')return _whText(raw);
  if(t==='num') return _whNum(raw);
  if(t==='int') return _whInt(raw);
  if(t==='date')return _whDate(raw);
  if(t==='hora')return _whHora(raw);
  if(t==='bool')return _whBool(raw);
  if(t==='sino')return _whBool(raw);   // "SI/NO" en hoja → boolean en pg (mismo forward que bool)
  if(t==='json')return _whJson(raw);
  return _whText(raw);
}
function _whRowMap(obj,spec){ var r={}; for(var i=0;i<spec.length;i++){ r[spec[i][0]]=_whVal(obj[spec[i][1]],spec[i][2]); } return r; }

/** Lector CRUDO por header (preserva Date/valores → _whDate formatea el timestamp completo). */
function _whSheetRows(name){
  var sh=getSheet(name);
  if(!sh) throw new Error('Hoja no encontrada: '+name);
  var data=sh.getDataRange().getValues();
  if(data.length<2) return [];
  var headers=data[0], out=[];
  for(var r=1;r<data.length;r++){
    var row=data[r], obj={}, any=false;
    for(var c=0;c<headers.length;c++){
      var h=headers[c]; if(h===''||h==null) continue;   // ignora columnas con header vacío (ej. PREINGRESOS)
      var v=row[c]; obj[h]=v;
      if(v!==''&&v!==null&&v!==undefined) any=true;
    }
    if(any) out.push(obj);
  }
  return out;
}

// ---------- specs por tabla ----------
// flags: onConflict (pg cols), keyHeader (header para contar/filtrar), lineaBy (genera linea), insertOnly, post
var _WH_SPECS = {
  guias: { sheet:'GUIAS', onConflict:'id_guia', keyHeader:'idGuia', spec:[
    ['id_guia','idGuia','text'],['tipo','tipo','text'],['fecha','fecha','date'],['usuario','usuario','text'],
    ['id_proveedor','idProveedor','text'],['id_zona','idZona','text'],['numero_documento','numeroDocumento','text'],
    ['comentario','comentario','text'],['monto_total','montoTotal','num'],['estado','estado','text'],
    ['id_preingreso','idPreingreso','text'],['foto','foto','text'],
    ['ocr_estado','OCR_Estado','text'],['ocr_tipo','OCR_Tipo','text'],['ocr_ruc_emisor','OCR_RUC_Emisor','text'],
    ['ocr_razon_social','OCR_Razon_Social','text'],['ocr_serie','OCR_Serie','text'],['ocr_numero','OCR_Numero','text'],
    ['ocr_fecha_comprobante','OCR_Fecha_Comprobante','date'],['ocr_total','OCR_Total','num'],['ocr_subtotal','OCR_Subtotal','num'],
    ['igv_recuperable','IGV_Recuperable','num'],['ocr_confidence','OCR_Confidence','num'],['ocr_notas','OCR_Notas','text'],
    ['ocr_fecha_proceso','OCR_Fecha_Proceso','date']
  ]},
  guia_detalle: { sheet:'GUIA_DETALLE', onConflict:'id_guia,linea', keyHeader:'idGuia', lineaBy:'id_guia', spec:[
    ['id_guia','idGuia','text'],['cod_producto','codigoProducto','text'],['cant_esperada','cantidadEsperada','num'],
    ['cant_recibida','cantidadRecibida','num'],['precio_unitario','precioUnitario','num'],['id_lote','idLote','text'],
    ['observacion','observacion','text'],['id_producto_nuevo','idProductoNuevo','text'],
    ['id_detalle','idDetalle','text'],['fecha_vencimiento','fechaVencimiento','date']
  ]},
  stock: { sheet:'STOCK', onConflict:'id_stock', keyHeader:'idStock', spec:[
    ['id_stock','idStock','text'],['cod_producto','codigoProducto','text'],['cantidad_disponible','cantidadDisponible','num'],
    ['ultima_actualizacion','ultimaActualizacion','date']
  ]},
  stock_movimientos: { sheet:'STOCK_MOVIMIENTOS', onConflict:'id_mov', keyHeader:'idMov', big:true, spec:[
    ['id_mov','idMov','text'],['fecha','fecha','date'],['cod_producto','codigoProducto','text'],['delta','delta','num'],
    ['stock_antes','stockAntes','num'],['stock_despues','stockDespues','num'],['tipo_operacion','tipoOperacion','text'],
    ['origen','origen','text'],['usuario','usuario','text']
  ]},
  lotes_vencimiento: { sheet:'LOTES_VENCIMIENTO', onConflict:'id_lote', keyHeader:'idLote', spec:[
    ['id_lote','idLote','text'],['cod_producto','codigoProducto','text'],['fecha_vencimiento','fechaVencimiento','date'],
    ['cantidad_inicial','cantidadInicial','num'],['cantidad_actual','cantidadActual','num'],['id_guia','idGuia','text'],
    ['estado','estado','text'],['fecha_creacion','fechaCreacion','date']
  ]},
  mermas: { sheet:'MERMAS', onConflict:'id_merma', keyHeader:'idMerma', spec:[
    ['id_merma','idMerma','text'],['fecha_ingreso','fechaIngreso','date'],['origen','origen','text'],['cod_producto','codigoProducto','text'],
    ['id_lote','idLote','text'],['cantidad_original','cantidadOriginal','num'],['cantidad_pendiente','cantidadPendiente','num'],
    ['motivo','motivo','text'],['usuario','usuario','text'],['id_guia','idGuia','text'],['estado','estado','text'],
    ['responsable','responsable','text'],['cantidad_reparada','cantidadReparada','num'],['cantidad_desechada','cantidadDesechada','num'],
    ['foto','foto','text'],['fecha_resolucion','fechaResolucion','date'],['observacion_resolucion','observacionResolucion','text'],
    ['id_guia_salida','idGuiaSalida','text']
  ]},
  auditorias: { sheet:'AUDITORIAS', onConflict:'id_auditoria', keyHeader:'idAuditoria', spec:[
    ['id_auditoria','idAuditoria','text'],['fecha_asignacion','fechaAsignacion','date'],['cod_producto','codigoProducto','text'],
    ['usuario','usuario','text'],['stock_sistema','stockSistema','num'],['stock_fisico','stockFisico','num'],
    ['diferencia','diferencia','num'],['resultado','resultado','text'],['observacion','observacion','text'],
    ['estado','estado','text'],['fecha_ejecucion','fechaEjecucion','date']
  ]},
  ajustes: { sheet:'AJUSTES', onConflict:'id_ajuste', keyHeader:'idAjuste', spec:[
    ['id_ajuste','idAjuste','text'],['cod_producto','codigoProducto','text'],['tipo_ajuste','tipoAjuste','text'],
    ['cantidad_ajuste','cantidadAjuste','num'],['motivo','motivo','text'],['usuario','usuario','text'],
    ['id_auditoria','idAuditoria','text'],['fecha','fecha','date']
  ]},
  envasados: { sheet:'ENVASADOS', onConflict:'id_envasado', keyHeader:'idEnvasado', spec:[
    ['id_envasado','idEnvasado','text'],['cod_producto_base','codigoProductoBase','text'],['cantidad_base','cantidadBase','num'],
    ['unidad_base','unidadBase','text'],['cod_producto_envasado','codigoProductoEnvasado','text'],['unidades_esperadas','unidadesEsperadas','num'],
    ['unidades_producidas','unidadesProducidas','num'],['merma_real','mermaReal','num'],['eficiencia_pct','eficienciaPct','num'],
    ['fecha','fecha','date'],['usuario','usuario','text'],['estado','estado','text'],['id_guia_salida','idGuiaSalida','text'],
    ['id_guia_ingreso','idGuiaIngreso','text'],['observacion','observacion','text']
  ]},
  preingresos: { sheet:'PREINGRESOS', onConflict:'id_preingreso', keyHeader:'idPreingreso', spec:[
    ['id_preingreso','idPreingreso','text'],['fecha','fecha','date'],['id_proveedor','idProveedor','text'],['cargadores','cargadores','text'],
    ['usuario','usuario','text'],['monto','monto','num'],['fotos','fotos','text'],['comentario','comentario','text'],
    ['estado','estado','text'],['id_guia','idGuia','text'],['snapshot_aviso','snapshotAviso','json']
  ]},
  producto_nuevo: { sheet:'PRODUCTO_NUEVO', onConflict:'id_producto_nuevo', keyHeader:'idProductoNuevo', spec:[
    ['id_producto_nuevo','idProductoNuevo','text'],['id_guia','idGuia','text'],['marca','marca','text'],['descripcion','descripcion','text'],
    ['codigo_barra','codigoBarra','text'],['id_categoria','idCategoria','text'],['unidad','unidad','text'],['cantidad','cantidad','num'],
    ['fecha_vencimiento','fechaVencimiento','date'],['foto','foto','text'],['estado','estado','text'],['usuario','usuario','text'],
    ['fecha_registro','fechaRegistro','date'],['aprobado_por','aprobadoPor','text'],['fecha_aprobacion','fechaAprobacion','date'],
    ['observacion','observacion','text']
  ]},
  sesiones: { sheet:'SESIONES', onConflict:'id_sesion', keyHeader:'idSesion', spec:[
    ['id_sesion','idSesion','text'],['id_personal','idPersonal','text'],['fecha_inicio','fechaInicio','date'],['hora_inicio','horaInicio','hora'],
    ['fecha_fin','fechaFin','date'],['hora_fin','horaFin','hora'],['minutos_activos','minutosActivos','num'],['estado','estado','text']
  ]},
  desempeno: { sheet:'DESEMPENO', onConflict:'id_desempeno', keyHeader:'idDesempeno', spec:[
    ['id_desempeno','idDesempeno','text'],['id_personal','idPersonal','text'],['id_sesion','idSesion','text'],['fecha','fecha','date'],
    ['minutos_activos','minutosActivos','num'],['horas_trabajadas','horasTrabajadas','num'],['guias_creadas','guiasCreadas','num'],
    ['guias_cerradas','guiasCerradas','num'],['envasados_registrados','envasadosRegistrados','num'],['unidades_envasadas','unidadesEnvasadas','num'],
    ['mermas_registradas','mermasRegistradas','num'],['auditoria_ejecutadas','auditoriaEjecutadas','num'],['preingreso_creados','preingresoCreados','num'],
    ['ajustes_realizados','ajustesRealizados','num'],['total_actividades','totalActividades','num'],['actividades_por_hora','actividadesPorHora','num'],
    ['puntuacion','puntuacion','num'],['calificacion','calificacion','text'],['monto_base','montoBase','num'],['monto_bonus','montoBonus','num'],
    ['monto_total','montoTotal','num'],['estado','estado','text']
  ]},
  pickups: { sheet:'PICKUPS', onConflict:'id_pickup', keyHeader:'idPickup', spec:[
    ['id_pickup','idPickup','text'],['fuente','fuente','text'],['estado','estado','text'],['items','items','json'],['id_zona','idZona','text'],
    ['notas','notas','text'],['creado_por','creadoPor','text'],['fecha_creado','fechaCreado','date'],['fecha_atendido','fechaAtendido','date'],
    ['atendido_por','atendidoPor','text'],['ultima_actividad','ultimaActividad','date']
  ]},
  ops_log: { sheet:'OPS_LOG', onConflict:'id_op', keyHeader:'idOp', spec:[
    ['id_op','idOp','text'],['id_guia','idGuia','text'],['tipo','tipo','text'],['payload','payload','json'],['estado','estado','text'],
    ['device_id','deviceId','text'],['usuario','usuario','text'],['fecha_creado','fechaCreado','date'],['fecha_aplicado','fechaAplicado','date'],
    ['error','error','text'],['resultado','resultado','json']
  ]},
  cargadores_log: { sheet:'CARGADORES_LOG', onConflict:'id_log', keyHeader:'idLog', spec:[
    ['id_log','idLog','text'],['fecha','fecha','date'],['id_cargador','idCargador','text'],['nombre','nombre','text'],
    ['added_by','addedBy','text'],['device_id','deviceId','text'],['ts','ts','date'],['estado','estado','text']
  ]},
  listas_sombra: { sheet:'LISTAS_SOMBRA', onConflict:'id_lista', keyHeader:'idLista', spec:[
    ['id_lista','idLista','text'],['fecha_creacion','fechaCreacion','date'],['usuario_creador','usuarioCreador','text'],['items','items','json'],
    ['estado','estado','text'],['usuario_tomada','usuarioTomada','text'],['fecha_tomada','fechaTomada','date'],
    ['fecha_completada','fechaCompletada','date'],['nota','nota','text']
  ]},
  lotes_adhesivo: { sheet:'LOTES_ADHESIVO', onConflict:'id_lote', keyHeader:'idLote', spec:[
    ['id_lote','idLote','text'],['fecha_creacion','fechaCreacion','date'],['fecha_ultimo_update','fechaUltimoUpdate','date'],
    ['usuario','usuario','text'],['origen','origen','text'],['codigo_barra','codigoBarra','text'],['descripcion','descripcion','text'],
    ['vto','vto','text'],['total_etq','totalEtq','num'],['completadas','completadas','num'],['sub_job_size','subJobSize','num'],
    ['status','status','text'],['ultimo_error','ultimoError','text'],['ultimo_printnode_job_id','ultimoPrintNodeJobId','text'],
    ['printer_id','printerId','text'],['tipo_etiqueta','tipoEtiqueta','text'],['items_json','itemsJson','json']
  ]},
  alertas_stock: { sheet:'ALERTAS_STOCK', onConflict:'id_alerta', keyHeader:'idAlerta', spec:[
    ['id_alerta','idAlerta','text'],['fecha','fecha','date'],['cod_producto','codigoProducto','text'],['descripcion','descripcion','text'],
    ['stock_real','stockReal','num'],['stock_teorico','stockTeorico','num'],['diferencia','diferencia','num'],
    ['revisado','revisado','sino'],['fecha_revision','fechaRevision','date']
  ]},
  config: { sheet:'CONFIG', onConflict:'clave', keyHeader:'clave', spec:[
    ['clave','clave','text'],['valor','valor','text'],['descripcion','descripcion','text']
  ]},
  clientes: { sheet:'Clientes', onConflict:'token', keyHeader:'token', spec:[
    ['token','token','text'],['nombre','nombre','text'],['telefono','telefono','text'],['tipo','tipo','text'],
    ['premium','premium','bool'],['fecha_alta','fechaAlta','date'],['ultimo_pedido','ultimoPedido','date']
  ]},
  pedidos_cliente: { sheet:'PedidosCliente', onConflict:'id_pedido', keyHeader:'idPedido', spec:[
    ['id_pedido','idPedido','text'],['token','token','text'],['ts','ts','date'],['estado','estado','text'],
    ['id_lista_sombra','idListaSombra','text'],['total_estimado','totalEstimado','num'],['notas','notas','text']
  ]},
  pedidos_cliente_items: { sheet:'PedidosClienteItems', onConflict:'id_pedido,idx', keyHeader:'idPedido', spec:[
    ['id_pedido','idPedido','text'],['idx','idx','int'],['nombre','nombre','text'],['cantidad','cantidad','num'],
    ['unidad','unidad','text'],['precio_est','precioEst','num'],['duda','duda','text']
  ]},
  pedidos_cliente_adj: { sheet:'PedidosClienteAdj', onConflict:'id_pedido,linea', keyHeader:'idPedido', lineaBy:'id_pedido', spec:[
    ['id_pedido','idPedido','text'],['tipo','tipo','text'],['nombre_archivo','nombreArchivo','text'],
    ['url_drive','urlDrive','text'],['ts','ts','date']
  ]}
};

var _WH_ORDEN=['guias','guia_detalle','stock','stock_movimientos','lotes_vencimiento','mermas','auditorias',
  'ajustes','envasados','preingresos','producto_nuevo','sesiones','desempeno','pickups','ops_log','cargadores_log',
  'listas_sombra','lotes_adhesivo','alertas_stock','config','clientes','pedidos_cliente','pedidos_cliente_items','pedidos_cliente_adj'];

var _WH_TIME_BUDGET = 4.5*60*1000;   // < 6 min límite GAS
var _WH_BATCH = 100;

/** Construye las filas pg de una tabla (mapeo + linea + dedupe + filtro). */
function _whBuildRows(tabla){
  var cfg=_WH_SPECS[tabla];
  var objs=_whSheetRows(cfg.sheet);
  var rows=objs.map(function(o){ var r=_whRowMap(o,cfg.spec); if(cfg.post) r=cfg.post(r,o); return r; });

  if(cfg.lineaBy){ // linea determinista por grupo (orden de hoja)
    var cnt={};
    rows=rows.filter(function(r){ return r[cfg.lineaBy]!=null && r[cfg.lineaBy]!==''; });
    rows.forEach(function(r){ var k=String(r[cfg.lineaBy]); cnt[k]=(cnt[k]||0)+1; r.linea=cnt[k]; });
  } else if(!cfg.insertOnly){ // pk simple O COMPUESTO: filtra sin pk + dedupe (gana el último)
    var pkCols=String(cfg.onConflict).split(',').map(function(c){ return c.trim(); });
    rows=rows.filter(function(r){ return pkCols.every(function(c){ return r[c]!=null && r[c]!==''; }); });
    var seen={}; rows.forEach(function(r){ var k=pkCols.map(function(c){ return String(r[c]); }).join('||'); seen[k]=r; });
    rows=Object.keys(seen).map(function(k){ return seen[k]; });
  }
  return rows;
}

/** Backfill principal. opts: {dryRun, soloTabla} */
function migrarWH(opts){
  opts=opts||{};
  var props=PropertiesService.getScriptProperties();
  var t0=Date.now();
  var tablas=opts.soloTabla?[opts.soloTabla]:_WH_ORDEN;
  var resumen={};

  for(var ti=0; ti<tablas.length; ti++){
    var tabla=tablas[ti], cfg=_WH_SPECS[tabla];
    if(!cfg){ resumen[tabla]={error:'spec desconocida'}; continue; }
    try{
      if(!getSheet(cfg.sheet)){ resumen[tabla]={saltado:'hoja no existe: '+cfg.sheet}; continue; }
      if(!opts.dryRun && !opts.soloTabla && props.getProperty('WHBF_DONE_'+tabla)==='1'){
        resumen[tabla]={saltado:'ya completada (resetCheckpointsWH para rehacer)'}; continue;
      }
      var rows=_whBuildRows(tabla);

      if(opts.dryRun){ resumen[tabla]={dryRun:true, filasValidas:rows.length, muestra:rows[0]||null}; continue; }

      var ckKey='WHBF_'+tabla;
      var start=parseInt(props.getProperty(ckKey)||'0',10);
      var errores=[], upserted=0, corto=false;
      for(var i=start; i<rows.length; i+=_WH_BATCH){
        if(Date.now()-t0 > _WH_TIME_BUDGET){
          props.setProperty(ckKey,String(i));
          resumen[tabla]={incompleto:true, desde:i, total:rows.length, nota:'re-corre backfillWH para continuar'};
          Logger.log(JSON.stringify(resumen,null,2));
          return resumen;
        }
        var lote=rows.slice(i,i+_WH_BATCH);
        if(JSON.stringify(lote).length>10000000){ errores.push('lote '+i+': payload muy grande, omitido'); props.setProperty(ckKey,String(i+_WH_BATCH)); continue; }
        var r=_sbUpsert('wh.'+tabla,lote,cfg.onConflict);
        if(r.ok){ upserted+=lote.length; props.setProperty(ckKey,String(i+_WH_BATCH)); }   // checkpoint SOLO en éxito
        else { errores.push('lote '+i+': HTTP '+r.code+' '+(r.error||'')); corto=true; break; }
      }
      if(errores.length===0){ props.deleteProperty(ckKey); props.setProperty('WHBF_DONE_'+tabla,'1'); }
      resumen[tabla]={filas:rows.length, upserted:upserted, errores:errores, ok:errores.length===0, incompleto:corto};
    }catch(e){ resumen[tabla]={error:String(e&&e.message||e)}; }
  }
  Logger.log(JSON.stringify(resumen,null,2));
  return resumen;
}

/** Compara conteos sheet vs supabase. */
function verificarCuadreWH(){
  var out={};
  _WH_ORDEN.forEach(function(tabla){
    var cfg=_WH_SPECS[tabla], nSheet=-1;
    try{
      if(!getSheet(cfg.sheet)){ out[tabla]={sheet:'(hoja no existe)', supabase:_sbCount('wh.'+tabla,null)}; return; }
      var rows=_whBuildRows(tabla);   // cuenta filas REALES a migrar (post dedupe/linea) → cuadre exacto
      nSheet=rows.length;
    }catch(e){ nSheet=-1; }
    var nPg=_sbCount('wh.'+tabla,null);
    out[tabla]={sheet:nSheet, supabase:nPg, cuadra:(nSheet===nPg)};
  });
  Logger.log(JSON.stringify(out,null,2));
  return out;
}

// ============================================================
// FASE 1.C — Doble escritura vía SYNC INCREMENTAL en segundo plano.
// No toca los endpoints de WH (cero latencia/riesgo). Re-upsertea (idempotente)
// las filas RECIENTES de cada tabla; las nuevas se agregan al final → quedan en la cola.
// ============================================================
var _WH_SYNC_TAILS = {
  guias:400, guia_detalle:1500, stock:99999, stock_movimientos:1500, lotes_vencimiento:99999,
  mermas:99999, auditorias:600, ajustes:600, envasados:400, preingresos:300, producto_nuevo:99999,
  sesiones:400, desempeno:400, pickups:99999, ops_log:500, cargadores_log:99999, listas_sombra:99999,
  lotes_adhesivo:300, alertas_stock:600, config:99999, clientes:99999, pedidos_cliente:99999,
  pedidos_cliente_items:99999, pedidos_cliente_adj:99999
};

function _syncWHImpl(full){
  var resumen={};
  _WH_ORDEN.forEach(function(tabla){
    var cfg=_WH_SPECS[tabla];
    try{
      if(!getSheet(cfg.sheet)){ return; }
      var rows=_whBuildRows(tabla);
      var tail=_WH_SYNC_TAILS[tabla]||300;
      var slice = (full || rows.length<=tail) ? rows : rows.slice(rows.length-tail);
      var err=[], up=0;
      for(var i=0;i<slice.length;i+=100){
        var lote=slice.slice(i,i+100);
        var r=_sbUpsert('wh.'+tabla,lote,cfg.onConflict);
        if(r.ok) up+=lote.length; else err.push('lote '+i+': HTTP '+r.code+' '+(r.error||''));
      }
      resumen[tabla]={sync:up, de:slice.length, errores:err};
    }catch(e){ resumen[tabla]={error:String(e&&e.message||e)}; }
  });
  Logger.log(JSON.stringify(resumen,null,2));
  return resumen;
}
function syncWHReciente(){ return _syncWHImpl(false); }   // 15 min: solo cola reciente (barato)
function syncWHCompleto(){ var r=_syncWHImpl(true); try{ reconciliarDiarioWH(); }catch(e){ Logger.log('recon WH falló: '+e); } return r; }   // recon pegada al sync nocturno (sin trigger extra)

/** Instala (idempotente) AMBOS triggers: incremental 15 min + completo nocturno (3:30am). Ejecutar 1 vez. */
function instalarTriggersSyncWH(){
  ScriptApp.getProjectTriggers().forEach(function(t){
    var h=t.getHandlerFunction(); if(h==='syncWHReciente'||h==='syncWHCompleto') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncWHReciente').timeBased().everyMinutes(15).create();
  ScriptApp.newTrigger('syncWHCompleto').timeBased().everyDays(1).atHour(3).nearMinute(30).create();
  Logger.log('Triggers instalados: syncWHReciente (15min) + syncWHCompleto (3:30am)');
  return {ok:true};
}
function desinstalarTriggersSyncWH(){
  var n=0; ScriptApp.getProjectTriggers().forEach(function(t){
    var h=t.getHandlerFunction(); if(h==='syncWHReciente'||h==='syncWHCompleto'){ ScriptApp.deleteTrigger(t); n++; }
  });
  return {ok:true, eliminados:n};
}

// ---------- wrappers para el editor ----------
function dryRunWH(){ return migrarWH({dryRun:true}); }
function backfillWH(){ return migrarWH(); }
function resetCheckpointsWH(){
  var props=PropertiesService.getScriptProperties();
  var n=0; _WH_ORDEN.forEach(function(t){
    ['WHBF_'+t,'WHBF_DONE_'+t].forEach(function(k){ if(props.getProperty(k)!=null){ props.deleteProperty(k); n++; } });
  });
  Logger.log('Checkpoints/flags borrados: '+n); return {ok:true, borrados:n};
}

// ============================================================
// RECONCILIACIÓN v2 — drift dashboard (conteo + SUMA de columnas clave)
// Detecta drift de VALORES (ediciones/anulaciones) que el solo conteo no ve.
// 100% lectura. Suma paginada ordenada por PK (cada fila exactamente 1 vez).
// ============================================================
var _WH_SUMCOLS = {
  guias:['monto_total'], guia_detalle:['cant_recibida'], stock:['cantidad_disponible'],
  stock_movimientos:['delta'], lotes_vencimiento:['cantidad_actual'], mermas:['cantidad_pendiente'],
  auditorias:['diferencia'], ajustes:['cantidad_ajuste'], envasados:['unidades_producidas'],
  preingresos:['monto'], producto_nuevo:['cantidad'], sesiones:['minutos_activos'], desempeno:['monto_total'],
  pickups:[], ops_log:[], cargadores_log:[], listas_sombra:[], lotes_adhesivo:['total_etq'],
  alertas_stock:[], config:[], clientes:[], pedidos_cliente:['total_estimado'],
  pedidos_cliente_items:['cantidad'], pedidos_cliente_adj:[]
};
var _WH_PRUNE = { alertas_stock:true };   // se auto-podan → el shadow conserva huérfanos (supabase ≥ sheet esperado)

/** Lee TODAS las filas de una tabla de Supabase paginando por PK estable (evita el cap db-max-rows=1000). */
function _sbSelectAll(schemaTable, order, select){
  var out=[], offset=0, PAGE=1000;
  while(true){
    var r=_sbSelect(schemaTable,{ select: select||'*', order:order, limit:PAGE, offset:offset });
    if(!r.ok) return { ok:false, error:'HTTP '+r.code+' '+(r.error||''), code:r.code };
    var rows=r.data||[];
    for(var i=0;i<rows.length;i++) out.push(rows[i]);
    if(rows.length<PAGE) break;
    offset+=PAGE;
    if(offset>200000) break; // backstop anti-bucle
  }
  return { ok:true, data:out };
}

/** Suma columnas de una tabla de Supabase, paginando ordenado por PK (estable). */
function _sbSumCols(schemaTable, cols, order){
  var sums={}; cols.forEach(function(c){ sums[c]=0; });
  var n=0, offset=0, PAGE=1000;
  while(true){
    var r=_sbSelect(schemaTable,{select:cols.join(',')||order.split(',')[0], order:order, limit:PAGE, offset:offset});
    if(!r.ok) return {error:'HTTP '+r.code+' '+(r.error||'')};
    var rows=r.data||[];
    rows.forEach(function(row){ cols.forEach(function(c){ var num=parseFloat(row[c]); if(!isNaN(num)) sums[c]+=num; }); });   // numeric puede venir como string desde PostgREST
    n+=rows.length;
    if(rows.length<PAGE) break;
    offset+=PAGE;
  }
  return {n:n, sums:sums};
}

function reconciliarWH(){
  var out={}, problemas=0;
  _WH_ORDEN.forEach(function(tabla){
    var cfg=_WH_SPECS[tabla], cols=_WH_SUMCOLS[tabla]||[], info={};
    try{
      var rows=getSheet(cfg.sheet)?_whBuildRows(tabla):[];
      info.sheet_n=rows.length;
      var ss={}; cols.forEach(function(c){ ss[c]=0; });
      rows.forEach(function(r){ cols.forEach(function(c){ var v=r[c]; if(typeof v==='number'&&!isNaN(v)) ss[c]+=v; }); });
      var sb=_sbSumCols('wh.'+tabla, cols, cfg.onConflict);
      if(sb.error){ info.error=sb.error; out[tabla]=info; problemas++; return; }
      info.sb_n=sb.n;
      info.n_ok=(info.sheet_n===info.sb_n);
      var sumOk=true; info.sums={};
      cols.forEach(function(c){ var a=ss[c]||0, b=sb.sums[c]||0, ok=Math.abs(a-b)<0.01; if(!ok)sumOk=false;
        info.sums[c]={sheet:Math.round(a*1000)/1000, sb:Math.round(b*1000)/1000, ok:ok}; });
      info.ok=(info.n_ok||_WH_PRUNE[tabla]) && sumOk;
      if(_WH_PRUNE[tabla] && !info.n_ok) info.nota='poda esperada (shadow conserva huérfanos)';
      if(!info.ok) problemas++;
    }catch(e){ info.error=String(e&&e.message||e); problemas++; }
    out[tabla]=info;
  });
  out._resumen={problemas:problemas, veredicto: problemas===0?'✓ SIN DRIFT':'⚠ revisar '+problemas+' tabla(s)'};
  Logger.log(JSON.stringify(out,null,2));
  return out;
}

/** Corre reconciliarWH y registra una fila en la hoja RECON_LOG (lo dispara el trigger diario). */
function reconciliarDiarioWH(){
  var res=reconciliarWH(), r=res._resumen||{};
  var probs={}; Object.keys(res).forEach(function(k){ if(k!=='_resumen' && res[k] && res[k].ok===false) probs[k]=res[k]; });
  var sh=getSheet('RECON_LOG') || getSpreadsheet().insertSheet('RECON_LOG');
  if(sh.getLastRow()===0) sh.appendRow(['fecha','app','problemas','veredicto','tablas_con_drift']);
  sh.appendRow([Utilities.formatDate(new Date(),'America/Lima','yyyy-MM-dd HH:mm'),'WH', r.problemas||0, r.veredicto||'', JSON.stringify(probs).slice(0,45000)]);
  return res;
}
/** La recon ahora va PEGADA a syncWHCompleto (sin trigger propio, por el límite de 20 triggers).
 *  Esta función solo LIMPIA un trigger de recon separado si lo instalaste antes. */
function desinstalarTriggerReconWH(){
  var n=0; ScriptApp.getProjectTriggers().forEach(function(t){ if(t.getHandlerFunction()==='reconciliarDiarioWH'){ ScriptApp.deleteTrigger(t); n++; } });
  Logger.log('Triggers recon separados eliminados: '+n+' (la recon corre dentro de syncWHCompleto)'); return {ok:true, eliminados:n};
}

// ============================================================
// DIAGNÓSTICO (solo lectura de hojas)
// ============================================================
var _WH_HOJAS = [
  'GUIAS','GUIA_DETALLE','STOCK','STOCK_MOVIMIENTOS','LOTES_VENCIMIENTO','MERMAS',
  'AUDITORIAS','AJUSTES','ENVASADOS','PREINGRESOS','PRODUCTO_NUEVO','SESIONES','DESEMPENO',
  'PICKUPS','OPS_LOG','CARGADORES_LOG','LISTAS_SOMBRA','LOTES_HISTORIAL','LOTES_ADHESIVO',
  'ALERTAS_STOCK','TICKETS_IMPRESOS','DEVOLUCIONES_ZONA','SYNC_LOG','JORNADAS',
  'Clientes','PedidosCliente','PedidosClienteItems','PedidosClienteAdj',
  'PRODUCTOS','PROVEEDORES','PERSONAL','ZONAS','CATEGORIAS','CONFIG'
];

function dumpHeadersWH(){
  var ss=getSpreadsheet(), out={};
  _WH_HOJAS.forEach(function(n){
    var sh=ss.getSheetByName(n);
    if(!sh){ out[n]='(NO EXISTE)'; return; }
    var lc=sh.getLastColumn();
    out[n]= lc<1 ? '(vacía)' : sh.getRange(1,1,1,lc).getValues()[0];
  });
  Logger.log(JSON.stringify(out,null,2));
  return out;
}

function inspeccionarTodoWH(){
  var ss=getSpreadsheet(), out={};
  _WH_HOJAS.forEach(function(n){
    var sh=ss.getSheetByName(n);
    if(!sh){ out[n]='(NO EXISTE)'; return; }
    var lc=sh.getLastColumn(), lr=sh.getLastRow();
    out[n]={ columnas:lc, filas:(lr-1),
      headers: lc>0 ? sh.getRange(1,1,1,lc).getValues()[0] : [],
      primeras: lr>1 ? sh.getRange(2,1,Math.min(2,lr-1),lc).getValues() : [],
      ultimas:  lr>2 ? sh.getRange(lr-1,1,2,lc).getValues() : []
    };
  });
  Logger.log(JSON.stringify(out,null,2));
  return out;
}

// Candidatos de PK a verificar por hoja. '|' = clave compuesta.
var _WH_PK_CANDIDATOS = {
  GUIAS:['idGuia'], GUIA_DETALLE:['idDetalle','idGuia|codigoProducto','idGuia|codigoProducto|idLote'],
  STOCK:['idStock','codigoProducto'], STOCK_MOVIMIENTOS:['idMov'], LOTES_VENCIMIENTO:['idLote'],
  MERMAS:['idMerma'], AUDITORIAS:['idAuditoria','idAuditoria|codigoProducto'], AJUSTES:['idAjuste'],
  ENVASADOS:['idEnvasado'], PREINGRESOS:['idPreingreso'], PRODUCTO_NUEVO:['idProductoNuevo'],
  SESIONES:['idSesion'], DESEMPENO:['idDesempeno','idPersonal|fecha'], PICKUPS:['idPickup'],
  OPS_LOG:['idOp'], CARGADORES_LOG:['idLog'], LISTAS_SOMBRA:['idLista'],
  LOTES_HISTORIAL:['idLote|ts','ts|idLote|codigoProducto|accion']
};

function chequearPKsWH(){
  var ss=getSpreadsheet(), out={};
  Object.keys(_WH_PK_CANDIDATOS).forEach(function(hoja){
    var sh=ss.getSheetByName(hoja);
    if(!sh){ out[hoja]='(NO EXISTE)'; return; }
    var rows=_sheetToObjects(sh);
    var info={ filas:rows.length, candidatos:{} };
    _WH_PK_CANDIDATOS[hoja].forEach(function(cand){
      var cols=cand.split('|'), seen={}, dups=0, ej=null, vacios=0;
      rows.forEach(function(r){
        var falta=cols.some(function(c){ return r[c]==null||r[c]===''; });
        if(falta){ vacios++; }
        var k=cols.map(function(c){ return String(r[c]==null?'':r[c]); }).join('||');
        if(seen[k]){ dups++; if(!ej) ej=k; } else seen[k]=true;
      });
      info.candidatos[cand]={ distintos:Object.keys(seen).length, duplicados:dups, conVacio:vacios, unico:(dups===0&&vacios===0), ejDup:ej };
    });
    out[hoja]=info;
  });
  Logger.log(JSON.stringify(out,null,2));
  return out;
}

function inspeccionarRestoWH2(){
  var ss=getSpreadsheet();
  var faltan=['OPS_LOG','CARGADORES_LOG','LISTAS_SOMBRA','LOTES_ADHESIVO','ALERTAS_STOCK',
    'TICKETS_IMPRESOS','DEVOLUCIONES_ZONA','SYNC_LOG','JORNADAS',
    'Clientes','PedidosCliente','PedidosClienteItems','PedidosClienteAdj',
    'PRODUCTOS','PROVEEDORES','PERSONAL','ZONAS','CATEGORIAS','CONFIG'];
  var out={};
  faltan.forEach(function(n){
    var sh=ss.getSheetByName(n);
    if(!sh){ out[n]='(NO EXISTE)'; return; }
    var lc=sh.getLastColumn(), lr=sh.getLastRow();
    out[n]={ cols:lc, filas:(lr-1), headers: lc>0 ? sh.getRange(1,1,1,lc).getValues()[0] : [] };
  });
  Logger.log(JSON.stringify(out,null,2));
  return out;
}

/** Prueba de conexión a Supabase desde WH. */
function _sbPingWH(){
  var out={ok:false, pasos:[]};
  try{ var cfg=_sbCfg_(); out.pasos.push('✓ Credenciales presentes ('+cfg.url+')'); }
  catch(e){ out.error=String(e.message); out.pasos.push('✗ '+out.error); Logger.log(JSON.stringify(out,null,2)); return out; }
  var t0=new Date().getTime();
  var r=_sbSelect('wh.config',{select:'clave',limit:1});
  out.latencia_ms=new Date().getTime()-t0;
  if(r.ok){ out.ok=true; out.pasos.push('✓ GET wh.config OK ('+out.latencia_ms+' ms, HTTP '+r.code+') — vacío [] es normal antes del backfill'); }
  else{ out.pasos.push('✗ GET wh.config falló: HTTP '+r.code+' — '+(r.error||'')); out.pasos.push('  Revisa: esquema wh EXPUESTO · 03_schema_wh.sql corrido · service_role key'); }
  Logger.log(JSON.stringify(out,null,2)); return out;
}

// ============================================================
// FASE 1.D (canary WH) — comparador de paridad getStock vs wh.stock_enriquecido()
// 100% shadow: llama a getStock (producción) y a la RPC, compara por idStock. NO toca el endpoint.
// ============================================================
function _numEq(a,b){ var na=parseFloat(a), nb=parseFloat(b); if(!isNaN(na)&&!isNaN(nb)) return Math.abs(na-nb)<0.01; return String(a)===String(b); }
function _dateEqSec(a,b){ var sa=String(a||''), sb=String(b||''); if(sa===''||sb==='') return sa===sb; var ta=new Date(sa).getTime(), tb=new Date(sb).getTime(); if(isNaN(ta)||isNaN(tb)) return sa===sb; return Math.floor(ta/1000)===Math.floor(tb/1000); }
function _diffStock(label,a,b,diffs){
  ['codigoProducto','descripcion','unidad'].forEach(function(f){ if(String(a[f])!==String(b[f])) diffs.push(label+'.'+f+': sheets="'+a[f]+'" sb="'+b[f]+'"'); });
  ['cantidadDisponible','stockMinimo','stockMaximo'].forEach(function(f){ if(!_numEq(a[f],b[f])) diffs.push(label+'.'+f+': sheets='+a[f]+' sb='+b[f]); });
  if(String(!!a.alertaMinimo)!==String(!!b.alertaMinimo)) diffs.push(label+'.alertaMinimo: sheets='+a.alertaMinimo+' sb='+b.alertaMinimo);
  if(!_dateEqSec(a.ultimaActualizacion,b.ultimaActualizacion)) diffs.push(label+'.ultimaActualizacion: sheets="'+a.ultimaActualizacion+'" sb="'+b.ultimaActualizacion+'"');
}
function compararStockWH(){
  var escenarios=[{n:'todos', solo:null}, {n:'soloAlertas', solo:'true'}];
  var salida={ok:true, escenarios:[]};
  escenarios.forEach(function(esc){
    var t0=Date.now(); var sh=getStock({soloAlertas:esc.solo}); var tS=Date.now()-t0;
    var t1=Date.now(); var r=_sbRpc('wh','stock_enriquecido',{solo_alertas:(esc.solo==='true')}); var tB=Date.now()-t1;
    var res={escenario:esc.n};
    if(!r.ok){ res.error='RPC falló: HTTP '+r.code+' — '+(r.error||''); res.nota='¿corriste 10_fase1d_wh_stock.sql?'; salida.ok=false; salida.escenarios.push(res); return; }
    var sd=(sh&&sh.data)||[], bd=(r.data&&r.data.data)||[], diffs=[];
    function byId(arr){ var m={}; arr.forEach(function(x){ m[String(x.idStock)]=x; }); return m; }
    var ms=byId(sd), mb=byId(bd), ids={};
    Object.keys(ms).forEach(function(k){ids[k]=1;}); Object.keys(mb).forEach(function(k){ids[k]=1;});
    Object.keys(ids).forEach(function(id){
      if(!ms[id]){ diffs.push(id+': falta en SHEETS'); return; }
      if(!mb[id]){ diffs.push(id+': falta en SUPABASE'); return; }
      _diffStock(id, ms[id], mb[id], diffs);
    });
    res.ok=diffs.length===0; res.filas={sheets:sd.length, sb:bd.length};
    res.velocidad={sheets_ms:tS, supabase_ms:tB, speedup:(tS&&tB)?(Math.round(tS/tB*10)/10+'x'):'n/a'};
    res.diferencias=diffs.slice(0,30); if(!res.ok) salida.ok=false;
    salida.escenarios.push(res);
  });
  salida.veredicto = salida.ok?'✓ PARIDAD EXACTA en ambos escenarios — listo para flip':'⚠ revisar diferencias';
  Logger.log(JSON.stringify(salida,null,2)); return salida;
}

// ============================================================
// FASE 1.D — FLIP con feature flag FUENTE_DATOS (Script Property de WH; default 'sheets').
// El router (Code.gs) llama getStockFlip. Ante CUALQUIER fallo de Supabase cae a Sheets.
// Encender: activarSupabaseWH() · Apagar: desactivarSupabaseWH() · granular: desactivarUnoWH('getStock')
// ============================================================
var _FLIP_CACHE_SEG = 15;
function _fuenteDatos(key){
  try{
    var p=PropertiesService.getScriptProperties();
    if(String(p.getProperty('FUENTE_DATOS')||'sheets').toLowerCase()!=='supabase') return 'sheets';
    var off=String(p.getProperty('FUENTE_DATOS_OFF')||'').toLowerCase();
    if(key && off){ var arr=off.split(',').map(function(s){return s.trim();}); if(arr.indexOf(String(key).toLowerCase())>=0) return 'sheets'; }
    return 'supabase';
  }catch(e){ return 'sheets'; }
}
function getStockFlip(params){
  params=params||{};
  if(_fuenteDatos('getStock')==='supabase'){
    try{
      var solo=(String(params.soloAlertas)==='true');
      var cache=CacheService.getScriptCache(), ckey='SB_STOCK_'+(solo?'A':'T');
      var hit=cache.get(ckey);
      if(hit) return JSON.parse(hit);
      var r=_sbRpc('wh','stock_enriquecido',{solo_alertas:solo});
      if(r.ok && r.data && Array.isArray(r.data.data)){
        try{ cache.put(ckey, JSON.stringify(r.data), _FLIP_CACHE_SEG); }catch(eC){}
        return r.data;   // {ok:true, data:[...]} (plano, igual que getStock)
      }
    }catch(e){ /* cae a Sheets */ }
  }
  return getStock(params);   // Sheets: default y fallback
}
// ---- controles del flip WH ----
function activarSupabaseWH(){ PropertiesService.getScriptProperties().setProperty('FUENTE_DATOS','supabase'); Logger.log('✅ FUENTE_DATOS(WH) = supabase — getStock lee de Supabase (fallback a Sheets si falla)'); return {ok:true, fuente:'supabase'}; }
function desactivarSupabaseWH(){ PropertiesService.getScriptProperties().setProperty('FUENTE_DATOS','sheets'); try{ CacheService.getScriptCache().removeAll(['SB_STOCK_T','SB_STOCK_A']); }catch(e){} Logger.log('↩️ FUENTE_DATOS(WH) = sheets — rollback instantáneo'); return {ok:true, fuente:'sheets'}; }
function estadoFuenteDatosWH(){ var p=PropertiesService.getScriptProperties(); var o={master:String(p.getProperty('FUENTE_DATOS')||'sheets'), off:String(p.getProperty('FUENTE_DATOS_OFF')||'')}; Logger.log(JSON.stringify(o)); return o; }
function desactivarUnoWH(ep){ var p=PropertiesService.getScriptProperties(); var off=(p.getProperty('FUENTE_DATOS_OFF')||'').split(',').map(function(s){return s.trim();}).filter(Boolean); if(off.indexOf(ep)<0) off.push(ep); p.setProperty('FUENTE_DATOS_OFF',off.join(',')); Logger.log('🔻 '+ep+' forzado a Sheets. OFF=['+off.join(',')+']'); return {ok:true,off:off}; }
function reactivarUnoWH(ep){ var p=PropertiesService.getScriptProperties(); var off=(p.getProperty('FUENTE_DATOS_OFF')||'').split(',').map(function(s){return s.trim();}).filter(Boolean).filter(function(e){return e!==ep;}); p.setProperty('FUENTE_DATOS_OFF',off.join(',')); Logger.log('🔼 '+ep+' reactivado a Supabase. OFF=['+off.join(',')+']'); return {ok:true,off:off}; }

// ---------- Canary WH #2: getRotacionSemanal (Sheets vs wh.rotacion_semanal) ----------
function compararRotacionWH(){
  var t0=Date.now(); var sh=getRotacionSemanal({semanas:8}); var tS=Date.now()-t0;
  var t1=Date.now(); var r=_sbRpc('wh','rotacion_semanal',{semanas:8, codigos_producto:null}); var tB=Date.now()-t1;
  if(!r.ok){ var e={ok:false, error:'RPC falló: HTTP '+r.code+' — '+(r.error||''), nota:'¿corriste 11_fase1d_wh_rotacion.sql?'}; Logger.log(JSON.stringify(e,null,2)); return e; }
  var sd=(sh&&sh.data)||{}, bd=(r.data&&r.data.data)||{}, diffs=[];
  var ea=sd.etiquetas||[], eb=bd.etiquetas||[];
  if(ea.join(',')!==eb.join(',')) diffs.push('etiquetas: sheets='+JSON.stringify(ea)+' sb='+JSON.stringify(eb));
  if(String(sd.semanas)!==String(bd.semanas)) diffs.push('semanas: sheets='+sd.semanas+' sb='+bd.semanas);
  var pa=sd.productos||{}, pb=bd.productos||{}, cbs={};
  Object.keys(pa).forEach(function(k){cbs[k]=1;}); Object.keys(pb).forEach(function(k){cbs[k]=1;});
  Object.keys(cbs).forEach(function(cb){
    if(!pa[cb]){ diffs.push(cb+': falta en SHEETS'); return; }
    if(!pb[cb]){ diffs.push(cb+': falta en SUPABASE'); return; }
    var sa=pa[cb], sb2=pb[cb];
    if(sa.length!==sb2.length){ diffs.push(cb+'.length: sheets='+sa.length+' sb='+sb2.length); return; }
    for(var i=0;i<sa.length;i++){
      if(String(sa[i].semana)!==String(sb2[i].semana)) diffs.push(cb+'['+i+'].semana: sheets="'+sa[i].semana+'" sb="'+sb2[i].semana+'"');
      if(!_numEq(sa[i].unidades,sb2[i].unidades)) diffs.push(cb+'['+sa[i].semana+'].unidades: sheets='+sa[i].unidades+' sb='+sb2[i].unidades);
    }
  });
  var out={ ok:diffs.length===0,
    veredicto: diffs.length===0?'✓ PARIDAD EXACTA — listo para flip':'⚠ '+diffs.length+' diferencias',
    velocidad:{sheets_ms:tS, supabase_ms:tB, speedup:(tS&&tB)?(Math.round(tS/tB*10)/10+'x'):'n/a'},
    conteos:{ productos:{sheets:Object.keys(pa).length, sb:Object.keys(pb).length}, etiquetas:ea.length },
    diferencias: diffs.slice(0,40) };
  Logger.log(JSON.stringify(out,null,2)); return out;
}

function getRotacionSemanalFlip(params){
  params=params||{};
  if(_fuenteDatos('getRotacionSemanal')==='supabase'){
    try{
      var semanas=parseInt(params.semanas,10)||8;
      var codigos=(params.codigosProducto!=null && String(params.codigosProducto).trim()!=='') ? String(params.codigosProducto) : null;
      var cache=CacheService.getScriptCache(), ckey=('SB_ROTACION_'+semanas+'_'+(codigos||'')).slice(0,240);
      var hit=cache.get(ckey);
      if(hit) return JSON.parse(hit);
      var r=_sbRpc('wh','rotacion_semanal',{semanas:semanas, codigos_producto:codigos});
      if(r.ok && r.data && r.data.data && Array.isArray(r.data.data.etiquetas)){
        var resp=r.data;
        resp.data.generadoEn = new Date().toISOString();   // la RPC no lo trae; lo agrega GAS igual que el original
        try{ cache.put(ckey, JSON.stringify(resp), _FLIP_CACHE_SEG); }catch(eC){}
        return resp;
      }
    }catch(e){ /* cae a Sheets */ }
  }
  return getRotacionSemanal(params);   // Sheets: default y fallback
}

// ════════════════════════════════════════════════════════════════════
// [Migración WH · Fase 2 · GATE] verificarParidadWH — SOLO LECTURA.
// Mide qué tan completa está la sombra de WH en Supabase comparándola contra
// Sheets (la fuente de verdad hoy). Es el prerequisito antes de habilitar
// cualquier lectura/escritura directa de WH — mismo gate que usamos en ME.
// Hoy la sombra se llena por sync BATCH (cada 15min), así que es esperable un
// pequeño hueco de lo creado en los últimos ~15min; este check lo cuantifica.
// GET ?action=verificarParidadWH&dias=3&tabla=guias  (tabla: guias|stock)
// ════════════════════════════════════════════════════════════════════
function verificarParidadWH(diasAtras, tabla){
  tabla = String(tabla || 'guias');
  var dias = parseInt(diasAtras, 10); if(!dias || dias < 1) dias = 3;

  if(tabla === 'stock'){
    // STOCK es estado actual (1 fila por producto): comparar presencia + cantidad.
    // ⚠️ ~1348 filas > db-max-rows=1000 de PostgREST → PAGINAR o se trunca silenciosamente.
    var shS = getSheet('STOCK'); if(!shS) return { ok:false, error:'STOCK no existe' };
    var dS = shS.getDataRange().getValues(); var hS = dS[0].map(function(h){return String(h||'').trim();});
    var iIdS = hS.indexOf('idStock'), iCod = hS.indexOf('codigoProducto'), iCant = hS.indexOf('cantidadDisponible');
    var shStock = {}, nSh = 0;
    for(var i=1;i<dS.length;i++){ var id=String(dS[i][iIdS]||'').trim(); if(!id) continue; shStock[id]={cant:parseFloat(dS[i][iCant])||0}; nSh++; }
    var rS = _sbSelectAll('wh.stock', 'id_stock.asc', 'id_stock,cantidad_disponible');
    if(!rS.ok) return { ok:false, error:'no se pudo leer wh.stock: '+(rS.error||'') };
    var supStock={}; (rS.data||[]).forEach(function(s){ supStock[String(s.id_stock||'').trim()]={cant:parseFloat(s.cantidad_disponible)||0}; });
    var faltan=[], difCant=[];
    Object.keys(shStock).forEach(function(id){
      if(!supStock[id]) faltan.push(id);
      else if(Math.abs(shStock[id].cant - supStock[id].cant) > 0.001) difCant.push({id:id, sheet:shStock[id].cant, supa:supStock[id].cant});
    });
    return { ok:true, data:{ tabla:'stock', sheets_total:nSh, supabase_total:(rS.data||[]).length,
      solo_en_sheets_count:faltan.length, solo_en_sheets:faltan.slice(0,30),
      cantidad_difiere_count:difCant.length, cantidad_difiere:difCant.slice(0,30) }};
  }

  // ── Verificador UNIVERSAL por presencia de PK (cubre las tablas dual-written) ──
  // Para cualquier tabla de _WH_SPECS distinta de guias/stock: id en col 0 de la hoja vs PK en Supabase.
  // Si la spec tiene una columna 'date', filtra la ventana de N días por esa fecha; si no, compara todo.
  if(tabla !== 'guias' && _WH_SPECS[tabla]){
    var cfg = _WH_SPECS[tabla];
    var pkCol = cfg.onConflict.split(',')[0];          // PK pg (1ra col)
    // Guard: solo tablas con PK de UNA columna que además es la 1ra col de la hoja (col 0 = id).
    // Excluye PK compuestas (guia_detalle, pedidos_cliente_*) donde col0 ≠ pkCol → comparación inválida.
    if(cfg.onConflict.indexOf(',') >= 0 || cfg.spec[0][0] !== pkCol){
      return { ok:false, error:'verificador universal no soporta '+tabla+' (PK compuesta o no es la 1ra col)' };
    }
    var shT = getSheet(cfg.sheet); if(!shT) return { ok:false, error:cfg.sheet+' no existe' };
    var dT = shT.getDataRange().getValues();
    if(dT.length < 2) return { ok:true, data:{ tabla:tabla, sheets_total:0, supabase_total:0, solo_en_sheets_count:0, solo_en_sheets:[] } };
    // Elegir columna de fecha para la ventana: priorizar fecha de CREACIÓN/registro, nunca vencimiento/aprobación.
    var bestEntry = null, bestScore = -99;
    for(var s=0;s<cfg.spec.length;s++){
      if(cfg.spec[s][2] !== 'date') continue;
      var pg = cfg.spec[s][0];
      if(pg.indexOf('venc') >= 0) continue;                 // vencimiento NO sirve como ventana temporal
      var sc = 0;
      if(/registro|asignacion|creacion/.test(pg)) sc = 3;
      else if(pg === 'fecha' || /_mov$/.test(pg)) sc = 2;
      else if(/ejecucion/.test(pg)) sc = 1;
      else if(/aprob/.test(pg)) sc = -1;
      if(sc > bestScore){ bestScore = sc; bestEntry = cfg.spec[s]; }
    }
    var dateHeader = bestEntry ? bestEntry[1] : null;
    var hT = dT[0].map(function(h){ return String(h||'').trim(); });
    var iDate = dateHeader ? hT.indexOf(dateHeader) : -1;
    var desdeT = new Date(Date.now() - dias*86400000);
    var shIdsT = {}, shTotalT = 0;
    for(var t=1;t<dT.length;t++){
      if(iDate >= 0){
        var fv = dT[t][iDate]; var fd = (fv instanceof Date) ? fv : new Date(fv);
        if(isNaN(fd.getTime()) || fd < desdeT) continue;
      }
      var idT = String(dT[t][0]||'').trim();
      if(idT){ shIdsT[idT]=1; shTotalT++; }
    }
    var rT = _sbSelectAll('wh.'+tabla, pkCol+'.asc', pkCol);
    if(!rT.ok) return { ok:false, error:'no se pudo leer wh.'+tabla+': '+(rT.error||'') };
    var enSupaT = {}; (rT.data||[]).forEach(function(row){ var v=String(row[pkCol]||'').trim(); if(v) enSupaT[v]=1; });
    var soloSheetsT = Object.keys(shIdsT).filter(function(id){ return !enSupaT[id]; });
    return { ok:true, data:{
      tabla:tabla, dias: (iDate>=0 ? dias : 'todo'), filtrado_por: dateHeader || '(sin fecha)',
      sheets_total:shTotalT, supabase_total:(rT.data||[]).length,
      solo_en_sheets_count: soloSheetsT.length, solo_en_sheets: soloSheetsT.slice(0,30)
    }};
  }

  // GUIAS (time-series): hueco = guía en Sheets ausente de Supabase
  var sh = getSheet('GUIAS'); if(!sh) return { ok:false, error:'GUIAS no existe' };
  var data = sh.getDataRange().getValues();
  var hdrs = data[0].map(function(h){ return String(h||'').trim(); });
  var iId = hdrs.indexOf('idGuia'), iFecha = hdrs.indexOf('fecha');
  if(iId < 0) return { ok:false, error:'col idGuia no encontrada' };
  var desde = new Date(Date.now() - dias*86400000);
  var shIds = {}, shTotal = 0;
  for(var k=1;k<data.length;k++){
    var f = iFecha>=0 ? data[k][iFecha] : null;
    var fecha = (f instanceof Date) ? f : new Date(f);
    if(iFecha>=0 && (isNaN(fecha.getTime()) || fecha < desde)) continue;
    var id = String(data[k][iId]||'').trim();
    if(id){ shIds[id]=1; shTotal++; }
  }
  var desdeIso = Utilities.formatDate(desde, 'America/Lima', "yyyy-MM-dd'T'00:00:00XXX");
  var r = _sbSelect('wh.guias', { filters:{ fecha:'gte.'+desdeIso }, order:'fecha.asc', limit:5000 });
  if(!r.ok) return { ok:false, error:'no se pudo leer wh.guias: '+(r.error||'') };
  var enSupa = {};
  (r.data||[]).forEach(function(g){ var id=String(g.id_guia||'').trim(); if(id) enSupa[id]=1; });
  var soloSheets = Object.keys(shIds).filter(function(id){ return !enSupa[id]; });
  return { ok:true, data:{
    tabla:'guias', dias:dias, sheets_total:shTotal, supabase_total:(r.data||[]).length,
    solo_en_sheets_count: soloSheets.length, solo_en_sheets: soloSheets.slice(0,30)
  }};
}

// ════════════════════════════════════════════════════════════════════
// [Migración WH · Fase 2 · PASO 2] Dual-write en TIEMPO REAL (best-effort).
// Espeja UNA fila a wh.<tabla> apenas se escribe en Sheets, reusando el mapeo del
// sync batch (_WH_SPECS + _whRowMap). Idempotente (upsert por PK). NUNCA lanza:
// un fallo de Supabase no debe romper la escritura a Sheets. El sync batch (15min)
// + reconciliación siguen como red de seguridad si un dual-write se pierde.
// `o` = objeto keyed por las cabeceras de la hoja (igual que produce _whSheetRows).
// ════════════════════════════════════════════════════════════════════
function _dualWriteWH(tabla, o){
  try {
    var cfg = _WH_SPECS[tabla];
    if(!cfg) { Logger.log('[dualWriteWH] spec desconocida: '+tabla); return { ok:false, error:'spec' }; }
    var row = _whRowMap(o, cfg.spec);
    if(cfg.post) row = cfg.post(row, o);
    // PK completa (sin pk → omitir; el batch lo levantará)
    var pkCols = String(cfg.onConflict).split(',').map(function(c){ return c.trim(); });
    for(var i=0;i<pkCols.length;i++){ if(row[pkCols[i]]==null || row[pkCols[i]]==='') return { ok:false, error:'falta pk '+pkCols[i] }; }
    var r = _sbUpsert('wh.'+tabla, [row], cfg.onConflict);
    if(!r.ok) Logger.log('[dualWriteWH '+tabla+'] upsert falló: HTTP '+(r.code)+' '+(r.error||''));
    return r;
  } catch(e){ Logger.log('[dualWriteWH '+tabla+'] '+(e&&e.message)); return { ok:false, error:String(e&&e.message||e) }; }
}


// [WH Fase 2 · PASO 2] PATCH best-effort de campos puntuales a wh.<tabla> (para cambios de estado, sin
// pisar el resto de la fila — p.ej. estado de guía sin tocar los campos OCR). pkFilters: {col:'eq.valor'}.
function _dualWritePatchWH(tabla, pkFilters, patch){
  try {
    // [20x] defensa extra: nunca PATCH con filtro vacío/sin valor (sino afectaría toda la tabla).
    if(!pkFilters || !Object.keys(pkFilters).length) return { ok:false, error:'PATCH sin filtros — abortado' };
    var malo = Object.keys(pkFilters).some(function(k){ var v=String(pkFilters[k]||''); return v==='' || /^[a-z]+\.$/.test(v); });
    if(malo) return { ok:false, error:'PATCH con filtro vacío — abortado' };
    // [Hardening 50x] patch vacío = no-op inútil (o PATCH degenerado). Abortar antes del HTTP.
    if(!patch || typeof patch !== 'object' || !Object.keys(patch).length) return { ok:false, error:'PATCH sin campos — abortado' };
    var r = _sb('PATCH', 'wh.'+tabla, { data: patch, filters: pkFilters, maxRetry: 1 });
    if(!r.ok) Logger.log('[dualWritePatchWH '+tabla+'] falló: HTTP '+(r.code)+' '+(r.error||''));
    return r;
  } catch(e){ Logger.log('[dualWritePatchWH '+tabla+'] '+(e&&e.message)); return { ok:false, error:String(e&&e.message||e) }; }
}

// [WH Fase 2 · PASO 2] Re-lee una fila de PREINGRESOS por id y la espeja a wh.preingresos (best-effort).
// Robusto: refleja SIEMPRE el estado actual en Sheets → sirve para crear/actualizar/aprobar sin reconstruir
// el objeto en cada caller. PREINGRESOS es chica (preingresos activos). NUNCA lanza.
function _dualWritePreingresoWH(idPreingreso){
  try {
    var id = String(idPreingreso || ''); if(!id) return { ok:false, error:'sin id' };
    var sh = getSheet('PREINGRESOS'); if(!sh) return { ok:false, error:'PREINGRESOS no existe' };
    var data = sh.getDataRange().getValues();
    var hdrs = data[0].map(function(h){ return String(h||'').trim(); });
    for(var i=1;i<data.length;i++){
      if(String(data[i][0]) !== id) continue;
      var o = {}; for(var c=0;c<hdrs.length;c++){ o[hdrs[c]] = data[i][c]; }
      return _dualWriteWH('preingresos', o);
    }
    return { ok:false, error:'preingreso no encontrado: '+id };
  } catch(e){ Logger.log('[dualWritePreingresoWH] '+(e&&e.message)); return { ok:false, error:String(e&&e.message||e) }; }
}

// [WH Fase 2 · PASO 2] Re-sincroniza TODAS las líneas de UNA guía a wh.guia_detalle (best-effort).
// La PK es (id_guia, linea) posicional → NO se puede dual-write una línea suelta sin re-numerar. Esta
// función reproduce EXACTAMENTE la numeración del batch (linea = N-ésima fila de la guía en orden de hoja),
// así el upsert pisa las filas correctas. Pensada para llamarse al CERRAR la guía (ítems finales = cuando
// se leen para despacho/auditoría). NUNCA lanza. Nota: si una fila se BORRÓ físicamente y N decrece, queda
// un huérfano de linea alta en la sombra (misma limitación que el batch; anularDetalle marca, no borra).
function _dualWriteDetallesGuiaWH(idGuia){
  try {
    var id = String(idGuia||''); if(!id) return { ok:false, error:'sin id' };
    var sh = getSheet('GUIA_DETALLE'); if(!sh) return { ok:false, error:'GUIA_DETALLE no existe' };
    var data = sh.getDataRange().getValues();
    var hdrs = data[0].map(function(h){ return String(h||'').trim(); });
    var iIdG = hdrs.indexOf('idGuia'); if(iIdG<0) return { ok:false, error:'col idGuia' };
    var cfg = _WH_SPECS.guia_detalle;
    var rows = [], linea = 0;
    for(var i=1;i<data.length;i++){
      if(String(data[i][iIdG]) !== id) continue;
      linea++;
      var o = {}; for(var c=0;c<hdrs.length;c++){ o[hdrs[c]] = data[i][c]; }
      var r = _whRowMap(o, cfg.spec);
      if(cfg.post) r = cfg.post(r, o);
      r.linea = linea;   // PK posicional — idéntico criterio que el batch (lineaBy id_guia, orden de hoja)
      rows.push(r);
    }
    if(!rows.length) return { ok:true, nada:true };
    var res = _sbUpsert('wh.guia_detalle', rows, cfg.onConflict);
    if(!res.ok) Logger.log('[dualWriteDetallesGuiaWH '+id+'] '+(res.error||''));
    return res;
  } catch(e){ Logger.log('[dualWriteDetallesGuiaWH] '+(e&&e.message)); return { ok:false, error:String(e&&e.message||e) }; }
}

// [WH Fase 2 · PASO 2] Re-lee un lote por id y lo espeja a wh.lotes_vencimiento (best-effort).
function _dualWriteLoteWH(idLote){
  try {
    var id = String(idLote||''); if(!id) return { ok:false, error:'sin id' };
    var sh = getSheet('LOTES_VENCIMIENTO'); if(!sh) return { ok:false, error:'LOTES_VENCIMIENTO no existe' };
    var data = sh.getDataRange().getValues();
    var hdrs = data[0].map(function(h){ return String(h||'').trim(); });
    for(var i=1;i<data.length;i++){
      if(String(data[i][0]) !== id) continue;
      var o = {}; for(var c=0;c<hdrs.length;c++){ o[hdrs[c]] = data[i][c]; }
      return _dualWriteWH('lotes_vencimiento', o);
    }
    return { ok:false, error:'lote no encontrado: '+id };
  } catch(e){ Logger.log('[dualWriteLoteWH] '+(e&&e.message)); return { ok:false, error:String(e&&e.message||e) }; }
}

// [WH Fase 2 · PASO 2] Re-lee una merma por id y la espeja a wh.mermas (best-effort).
function _dualWriteMermaWH(idMerma){
  try {
    var id = String(idMerma||''); if(!id) return { ok:false, error:'sin id' };
    var sh = getSheet('MERMAS'); if(!sh) return { ok:false, error:'MERMAS no existe' };
    var data = sh.getDataRange().getValues();
    var hdrs = data[0].map(function(h){ return String(h||'').trim(); });
    for(var i=1;i<data.length;i++){
      if(String(data[i][0]) !== id) continue;
      var o = {}; for(var c=0;c<hdrs.length;c++){ o[hdrs[c]] = data[i][c]; }
      return _dualWriteWH('mermas', o);
    }
    return { ok:false, error:'merma no encontrada: '+id };
  } catch(e){ Logger.log('[dualWriteMermaWH] '+(e&&e.message)); return { ok:false, error:String(e&&e.message||e) }; }
}

// [WH Fase 2 · PASO 2 · R4] Re-lee una auditoría por id y la espeja a wh.auditorias (best-effort).
function _dualWriteAuditoriaWH(idAuditoria){
  try {
    var id = String(idAuditoria||''); if(!id) return { ok:false, error:'sin id' };
    var sh = getSheet('AUDITORIAS'); if(!sh) return { ok:false, error:'AUDITORIAS no existe' };
    var data = sh.getDataRange().getValues();
    var hdrs = data[0].map(function(h){ return String(h||'').trim(); });
    for(var i=1;i<data.length;i++){
      if(String(data[i][0]) !== id) continue;
      var o = {}; for(var c=0;c<hdrs.length;c++){ o[hdrs[c]] = data[i][c]; }
      return _dualWriteWH('auditorias', o);
    }
    return { ok:false, error:'auditoria no encontrada: '+id };
  } catch(e){ Logger.log('[dualWriteAuditoriaWH] '+(e&&e.message)); return { ok:false, error:String(e&&e.message||e) }; }
}

// [WH Fase 2 · PASO 2 · R4] Re-lee un producto nuevo por id y lo espeja a wh.producto_nuevo (best-effort).
function _dualWriteProductoNuevoWH(idProductoNuevo){
  try {
    var id = String(idProductoNuevo||''); if(!id) return { ok:false, error:'sin id' };
    var sh = getSheet('PRODUCTO_NUEVO'); if(!sh) return { ok:false, error:'PRODUCTO_NUEVO no existe' };
    var data = sh.getDataRange().getValues();
    var hdrs = data[0].map(function(h){ return String(h||'').trim(); });
    for(var i=1;i<data.length;i++){
      if(String(data[i][0]) !== id) continue;
      var o = {}; for(var c=0;c<hdrs.length;c++){ o[hdrs[c]] = data[i][c]; }
      return _dualWriteWH('producto_nuevo', o);
    }
    return { ok:false, error:'producto nuevo no encontrado: '+id };
  } catch(e){ Logger.log('[dualWriteProductoNuevoWH] '+(e&&e.message)); return { ok:false, error:String(e&&e.message||e) }; }
}

// ════════════════════════════════════════════════════════════════════
// [Migración WH · PASO 3] LECTURA DIRECTA genérica desde la sombra Supabase.
// Invierte _sheetToObjects: reconstruye objetos con el MISMO shape (claves camelCase, fechas
// 'yyyy-MM-dd' en TZ del script, celdas vacías='') a partir de wh.<tabla>, reusando _WH_SPECS.
// Un solo helper sirve para TODAS las tablas → cada lectura nueva = un getXxxFlip + cablear.
// El gate compararLecturaWH(tabla) valida paridad EXACTA por id ANTES de cablear (regla de oro).
// ════════════════════════════════════════════════════════════════════

// Convierte un valor pg al valor que _sheetToObjects habría producido de la celda original.
function _sbValToSheet(v, t, tz){
  if(v==null) return '';                                   // celda vacía en Sheets = ''
  if(t==='num' || t==='int'){ var n=(typeof v==='number')?v:parseFloat(v); return isNaN(n)?'':n; }
  if(t==='date'){ var d=(v instanceof Date)?v:new Date(v); return isNaN(d.getTime())?'':Utilities.formatDate(d,tz,'yyyy-MM-dd'); }
  if(t==='bool'){ return (v===true||v==='true'||v===1||v==='1'); }
  if(t==='sino'){ return (v===true||v==='true'||v===1||v==='1') ? 'SI' : 'NO'; }   // boolean pg → "SI/NO" (como la celda)
  if(t==='hora'){ return String(v); }
  if(t==='json'){ return (typeof v==='object') ? JSON.stringify(v) : String(v); }  // string JSON, como _sheetToObjects (celda texto)
  return String(v);                                        // text
}

// Reconstruye los objetos de UNA tabla desde filas pg (mismo shape que _sheetToObjects).
function _sbRowsToObjsWH(tabla, rows){
  var cfg=_WH_SPECS[tabla]; if(!cfg) throw new Error('sin spec: '+tabla);
  var tz=Session.getScriptTimeZone(), out=[];
  for(var i=0;i<rows.length;i++){
    var row=rows[i], o={};
    for(var s=0;s<cfg.spec.length;s++){ o[cfg.spec[s][1]]=_sbValToSheet(row[cfg.spec[s][0]], cfg.spec[s][2], tz); }
    out.push(o);
  }
  return out;
}

// Lee TODA una tabla wh.<tabla> (paginado, ordenado por PK) y la mapea a shape Sheets. Lanza si falla.
function _leerTablaWH(tabla){
  var cfg=_WH_SPECS[tabla]; if(!cfg) throw new Error('sin spec: '+tabla);
  var pk=cfg.onConflict.split(',')[0];
  var r=_sbSelectAll('wh.'+tabla, pk+'.asc');
  if(!r.ok) throw new Error('lectura wh.'+tabla+' falló: '+(r.error||''));
  return _sbRowsToObjsWH(tabla, r.data||[]);
}

// PUNTO ÚNICO de lectura de una tabla para las funciones de API: sombra Supabase si el flip está ON
// (key 'lectura_<tabla>'), con FALLBACK automático a Sheets ante cualquier fallo. Reemplaza
// `_sheetToObjects(getSheet(X))` en las funciones getXxx de presentación (NO en lecturas internas).
function _filasLecturaWH(tabla, sheetName){
  if(_fuenteDatos('lectura_'+tabla)==='supabase'){
    try{ return _leerTablaWH(tabla); }
    catch(e){ Logger.log('[filasLecturaWH '+tabla+'] cae a Sheets: '+(e&&e.message)); }
  }
  return _sheetToObjects(getSheet(sheetName));
}

// Comparación numérica tolerante (vacío==vacío, ±0.001).
function _numEqLoose(a,b){
  var va=(a===''||a==null), vb=(b===''||b==null);
  if(va&&vb) return true; if(va!==vb) return false;
  var x=parseFloat(a), y=parseFloat(b); if(isNaN(x)&&isNaN(y)) return true;
  return Math.abs((x||0)-(y||0))<0.001;
}

// Ordena claves recursivamente para comparar JSON por CONTENIDO (pg jsonb reordena las claves).
function _sortKeysDeep(o){
  if(Array.isArray(o)) return o.map(_sortKeysDeep);
  if(o&&typeof o==='object'){ var r={}; Object.keys(o).sort().forEach(function(k){ r[k]=_sortKeysDeep(o[k]); }); return r; }
  return o;
}
// Normaliza a 'yyyy-MM-dd' (igual que _sheetToObjects trunca toda Date). Tolera string ISO con hora.
function _norm10(v){
  if(v==null||v==='') return '';
  var s=String(v).trim();
  if(/^\d{4}-\d{2}-\d{2}/.test(s)) return s.substring(0,10);
  if(/^\d{2}\/\d{2}\/\d{4}$/.test(s)){ var p=s.split('/'); return p[2]+'-'+p[1]+'-'+p[0]; }  // dd/MM/yyyy → yyyy-MM-dd
  var d=new Date(s); return isNaN(d.getTime())?s:Utilities.formatDate(d,Session.getScriptTimeZone(),'yyyy-MM-dd');
}
// Igualdad de fecha tolerante: ambos lados se comparan como 'yyyy-MM-dd' (la celda puede ser Date o string ISO).
function _dateEqLoose(a,b){ return _norm10(a)===_norm10(b); }

// Igualdad de JSON tolerante al orden de claves y a string-vs-objeto (mismo contenido = igual).
function _jsonEqLoose(a,b){
  var pa, pb;
  try{ pa=(typeof a==='string')?JSON.parse(a||'null'):a; }catch(e){ return String(a)===String(b); }
  try{ pb=(typeof b==='string')?JSON.parse(b||'null'):b; }catch(e){ return String(a)===String(b); }
  return JSON.stringify(_sortKeysDeep(pa))===JSON.stringify(_sortKeysDeep(pb));
}

// Diffs por id entre dataset Sheets y dataset Supabase de UNA tabla (campo a campo según spec).
function _diffDatasetWH(tabla, sheetData, supaData){
  var cfg=_WH_SPECS[tabla], key=cfg.keyHeader, diffs=[];
  function byKey(arr){ var m={}; for(var i=0;i<arr.length;i++){ m[String(arr[i][key])]=arr[i]; } return m; }
  var ms=byKey(sheetData), mb=byKey(supaData), ids={};
  Object.keys(ms).forEach(function(k){ids[k]=1;}); Object.keys(mb).forEach(function(k){ids[k]=1;});
  Object.keys(ids).forEach(function(id){
    if(!ms[id]){ diffs.push(id+': falta en SHEETS'); return; }
    if(!mb[id]){ diffs.push(id+': falta en SUPABASE'); return; }
    var a=ms[id], b=mb[id];
    for(var s=0;s<cfg.spec.length;s++){
      var h=cfg.spec[s][1], t=cfg.spec[s][2];
      var eq=(t==='num'||t==='int') ? _numEqLoose(a[h],b[h])
           : (t==='json')          ? _jsonEqLoose(a[h],b[h])
           : (t==='date')          ? _dateEqLoose(a[h],b[h])
           :                         (String(a[h])===String(b[h]));
      if(!eq) diffs.push(id+'.'+h+': sheets='+JSON.stringify(a[h])+' sb='+JSON.stringify(b[h]));
    }
  });
  return diffs;
}

// Map tabla → lectura CRUDA de la hoja (sin filtros) para el gate. Se extiende por ronda.
var _LECTURA_SHEET_FN = {
  mermas:         function(){ return _sheetToObjects(getSheet('MERMAS')); },
  auditorias:     function(){ return _sheetToObjects(getSheet('AUDITORIAS')); },
  ajustes:        function(){ return _sheetToObjects(getSheet('AJUSTES')); },
  envasados:      function(){ return _sheetToObjects(getSheet('ENVASADOS')); },
  producto_nuevo: function(){ return _sheetToObjects(getSheet('PRODUCTO_NUEVO')); },
  guias:          function(){ return _sheetToObjects(getSheet('GUIAS')); },
  preingresos:    function(){ return _sheetToObjects(getSheet('PREINGRESOS')); },
  lotes_vencimiento: function(){ return _sheetToObjects(getSheet('LOTES_VENCIMIENTO')); },
  stock_movimientos: function(){ return _sheetToObjects(getSheet('STOCK_MOVIMIENTOS')); },
  alertas_stock:     function(){ return _sheetToObjects(getSheet('ALERTAS_STOCK')); }
};

// [Revisión 20x] AUDITOR de cobertura: columnas en la hoja que el spec NO mapea → se PIERDEN en la
// lectura directa (el gate de paridad NO las detecta porque solo compara campos del spec). Para cada
// tabla flipeada, lista las columnas huérfanas que algún consumidor del frontend podría necesitar.
function auditarColumnasSpecWH(){
  var out={};
  ['mermas','auditorias','ajustes','envasados','producto_nuevo','preingresos','lotes_vencimiento','stock_movimientos','guias'].forEach(function(t){
    try{
      var cfg=_WH_SPECS[t]; var sh=getSheet(cfg.sheet);
      var data=sh.getDataRange().getValues();
      var hdrs=(data[0]||[]).map(function(h){return String(h||'').trim();}).filter(function(h){return h!=='';});
      var specH={}; cfg.spec.forEach(function(s){ specH[s[1]]=1; });
      out[t]={ hoja:cfg.sheet, cols_hoja:hdrs.length, cols_spec:cfg.spec.length,
        faltan_en_spec: hdrs.filter(function(h){ return !specH[h]; }) };
    }catch(e){ out[t]={error:String(e&&e.message||e)}; }
  });
  return out;
}

// GATE genérico de paridad de lectura: Sheets crudo vs sombra Supabase, por id.
function compararLecturaWH(tabla){
  tabla=String(tabla||'');
  if(!_WH_SPECS[tabla]) return { ok:false, error:'tabla sin spec: '+tabla };
  if(!_LECTURA_SHEET_FN[tabla]) return { ok:false, error:'tabla sin lectura sheet registrada: '+tabla };
  try{
    var t0=Date.now(); var sh=_LECTURA_SHEET_FN[tabla](); var tS=Date.now()-t0;
    var t1=Date.now(); var sb=_leerTablaWH(tabla); var tB=Date.now()-t1;
    var diffs=_diffDatasetWH(tabla, sh, sb);
    return { ok:diffs.length===0, tabla:tabla,
      filas:{sheets:sh.length, sb:sb.length},
      velocidad:{sheets_ms:tS, supabase_ms:tB, speedup:(tS&&tB)?(Math.round(tS/tB*10)/10+'x'):'n/a'},
      veredicto: diffs.length===0?'✓ PARIDAD EXACTA — listo para flip':'⚠ '+diffs.length+' diferencias',
      diferencias: diffs.slice(0,40) };
  }catch(e){ return { ok:false, tabla:tabla, error:String(e&&e.message||e) }; }
}

// [Regla: 1 fila por producto en STOCK] Deduplica la hoja STOCK: por cada codigoProducto con >1 fila,
// conserva la de ultimaActualizacion MÁS RECIENTE (la vigente) y borra las demás. dryRun por defecto.
function dedupStockSheet(opts){
  opts = opts || {};
  var dry = String(opts.dryRun) !== 'false';   // default DRY-RUN (no borra)
  return _conLock('dedupStockSheet', function(){
    var sheet = getSheet('STOCK');
    var data = sheet.getDataRange().getValues();
    var hdrs = data[0];
    var iId = hdrs.indexOf('idStock'), iCod = hdrs.indexOf('codigoProducto'),
        iCant = hdrs.indexOf('cantidadDisponible'), iAct = hdrs.indexOf('ultimaActualizacion');
    var porCod = {};
    for(var r=1;r<data.length;r++){
      var cod = String(data[r][iCod]||'').trim(); if(!cod) continue;
      (porCod[cod] = porCod[cod] || []).push({ row:r+1, idStock:String(data[r][iId]||''), cant:data[r][iCant], act:data[r][iAct] });
    }
    var grupos = [], rowsBorrar = [];
    Object.keys(porCod).forEach(function(cod){
      var fs = porCod[cod]; if(fs.length < 2) return;
      fs.sort(function(a,b){ return new Date(b.act).getTime() - new Date(a.act).getTime(); });  // más reciente primero
      var keep = fs[0], del = fs.slice(1);
      grupos.push({ cod:cod, conserva:{idStock:keep.idStock, cant:keep.cant, act:String(keep.act)},
        borra: del.map(function(x){ return {idStock:x.idStock, cant:x.cant, act:String(x.act)}; }) });
      del.forEach(function(x){ rowsBorrar.push(x.row); });
    });
    var postCount = null;
    if(!dry){
      var nCols = hdrs.length;
      // VACIAR la fila huérfana (clearContent) en vez de deleteRow: si la hoja tiene protección de estructura
      // o filtro, deleteRow es no-op pero setValue/clearContent sí funciona. _sheetToObjects descarta filas
      // 100% vacías y el batch no sube filas sin idStock → la huérfana desaparece lógicamente.
      rowsBorrar.sort(function(a,b){ return b-a; }).forEach(function(rw){ sheet.getRange(rw, 1, 1, nCols).clearContent(); });
      SpreadsheetApp.flush();
      // re-leer DENTRO de la misma ejecución para confirmar el borrado real (diagnóstico)
      var d2 = sheet.getDataRange().getValues(); var c2 = {};
      for(var k=1;k<d2.length;k++){ var cc=String(d2[k][iCod]||'').trim(); if(cc) c2[cc]=(c2[cc]||0)+1; }
      postCount = {}; grupos.forEach(function(g){ postCount[g.cod] = c2[g.cod]||0; });
    }
    return { ok:true, dryRun:dry, gruposDuplicados:grupos.length, filasBorradas:(dry?0:rowsBorrar.length),
             filasABorrar:rowsBorrar.length, postConteoPorCod:postCount, lastRow:sheet.getLastRow(), detalle:grupos.slice(0,30) };
  });
}

// [PASO 3 fix] Re-sincroniza estado+detalle de guías AUTOCERRADA recientes a la sombra (históricas que
// autoCloseDay marcó sin dual-write). Cubre la ventana que la rotación cuenta. Best-effort por guía.
function resyncDetalleAutocerradas(diasAtras){
  var dias = parseInt(diasAtras, 10) || 60;
  var sh = getSheet('GUIAS'); var data = sh.getDataRange().getValues(); var h = data[0];
  var iId = h.indexOf('idGuia'), iFec = h.indexOf('fecha'), iEst = h.indexOf('estado');
  var desde = new Date(Date.now() - dias*86400000);
  var n = 0, errs = 0;
  for (var i = 1; i < data.length; i++){
    if (String(data[i][iEst] || '').toUpperCase() !== 'AUTOCERRADA') continue;
    var fv = data[i][iFec]; var f = (fv instanceof Date) ? fv : new Date(fv);
    if (isNaN(f.getTime()) || f < desde) continue;
    var id = String(data[i][iId] || ''); if (!id) continue;
    try {
      if (typeof _dualWritePatchWH === 'function') _dualWritePatchWH('guias', { id_guia: 'eq.' + id }, { estado: 'AUTOCERRADA' });
      if (typeof _dualWriteDetallesGuiaWH === 'function') _dualWriteDetallesGuiaWH(id);
      n++;
    } catch(e){ errs++; }
  }
  return { ok: true, resincronizadas: n, errores: errs, dias: dias };
}

// [PASO 3 fix] Re-sincroniza estado+detalle de guías CERRADA/AUTOCERRADA recientes (ventana de rotación) a la
// sombra — cubre cualquier guía de salida con detalle incompleto que descuadre la rotación. Best-effort por guía.
function resyncDetalleGuiasRecientes(diasAtras){
  var dias = parseInt(diasAtras, 10) || 60;
  var sh = getSheet('GUIAS'); var data = sh.getDataRange().getValues(); var h = data[0];
  var iId = h.indexOf('idGuia'), iFec = h.indexOf('fecha'), iEst = h.indexOf('estado'), iTipo = h.indexOf('tipo');
  var desde = new Date(Date.now() - dias*86400000);
  var n = 0, errs = 0, salidas = 0;
  for (var i = 1; i < data.length; i++){
    var est = String(data[i][iEst] || '').toUpperCase();
    if (est !== 'CERRADA' && est !== 'AUTOCERRADA') continue;
    var fv = data[i][iFec]; var f = (fv instanceof Date) ? fv : new Date(fv);
    if (isNaN(f.getTime()) || f < desde) continue;
    var id = String(data[i][iId] || ''); if (!id) continue;
    if (String(data[i][iTipo] || '').toUpperCase().indexOf('SALIDA') === 0) salidas++;
    try {
      if (typeof _dualWritePatchWH === 'function') _dualWritePatchWH('guias', { id_guia: 'eq.' + id }, { estado: est });
      if (typeof _dualWriteDetallesGuiaWH === 'function') _dualWriteDetallesGuiaWH(id);
      n++;
    } catch(e){ errs++; }
  }
  return { ok: true, resincronizadas: n, salidas: salidas, errores: errs, dias: dias };
}
