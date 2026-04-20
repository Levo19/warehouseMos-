// ============================================================
// warehouseMos — Dashboard.gs  KPIs y alertas
// ============================================================

function getDashboard() {
  var diasAlerta       = parseInt(_getConfigValue('DIAS_ALERTA_VENC')) || 30;
  var diasAlertaCrit   = parseInt(_getConfigValue('DIAS_ALERTA_VENC_CRITICO')) || 7;
  var hoy              = new Date();
  var limAlerta        = new Date(hoy); limAlerta.setDate(hoy.getDate() + diasAlerta);
  var limCritico       = new Date(hoy); limCritico.setDate(hoy.getDate() + diasAlertaCrit);

  var productos  = _sheetToObjects(getProductosSheet());
  var stock      = _sheetToObjects(getSheet('STOCK'));
  var lotes      = _sheetToObjects(getSheet('LOTES_VENCIMIENTO'));
  var guias      = _sheetToObjects(getSheet('GUIAS'));
  var preingresos= _sheetToObjects(getSheet('PREINGRESOS'));
  var mermas     = _sheetToObjects(getSheet('MERMAS'));
  var auditorias = _sheetToObjects(getSheet('AUDITORIAS'));
  var envasados  = _sheetToObjects(getSheet('ENVASADOS'));

  // Índices rápidos
  var stockMap = {};
  stock.forEach(function(s) { stockMap[s.codigoProducto] = parseFloat(s.cantidadDisponible) || 0; });

  var productosMap = {};
  productos.forEach(function(p) { productosMap[p.idProducto] = p; });

  // ── 1. Vencimientos críticos (< diasAlertaCrit días) ───────
  var vencCriticos = [];
  var vencAlertas  = [];
  lotes.forEach(function(l) {
    if (l.estado !== 'ACTIVO') return;
    var fv = new Date(l.fechaVencimiento);
    var cant = parseFloat(l.cantidadActual) || 0;
    if (cant <= 0) return;
    var prod = productosMap[l.codigoProducto] || {};
    var entry = {
      idLote: l.idLote,
      codigoProducto: l.codigoProducto,
      descripcion: prod.descripcion || l.codigoProducto,
      fechaVencimiento: l.fechaVencimiento,
      cantidadActual: cant,
      diasRestantes: Math.ceil((fv - hoy) / (1000 * 60 * 60 * 24))
    };
    if (fv <= limCritico) {
      vencCriticos.push(entry);
    } else if (fv <= limAlerta) {
      vencAlertas.push(entry);
    }
  });
  vencCriticos.sort(function(a,b){ return a.diasRestantes - b.diasRestantes; });
  vencAlertas.sort(function(a,b){ return a.diasRestantes - b.diasRestantes; });

  // ── 2. Stock bajo mínimo ───────────────────────────────────
  var bajominimo = [];
  productos.forEach(function(p) {
    if (p.estado !== '1') return;
    var actual = stockMap[p.idProducto] || 0;
    var minimo = parseFloat(p.stockMinimo) || 0;
    if (actual < minimo) {
      bajominimo.push({
        codigo: p.idProducto,
        descripcion: p.descripcion,
        stockActual: actual,
        stockMinimo: minimo,
        diferencia: minimo - actual
      });
    }
  });
  bajominimo.sort(function(a,b){ return (b.diferencia/b.stockMinimo) - (a.diferencia/a.stockMinimo); });

  // ── 3. Pendientes de envasado ──────────────────────────────
  var pendientesEnvasado = _calcularPendientesEnvasado(productos, stockMap);

  // ── 4. Preingresos pendientes ──────────────────────────────
  var presPendientes = preingresos.filter(function(p){ return p.estado === 'PENDIENTE'; });

  // ── 5. Guías abiertas > 48h ────────────────────────────────
  var hace48h = new Date(hoy); hace48h.setHours(hoy.getHours() - 48);
  var guiasAbiertas = guias.filter(function(g) {
    return g.estado === 'ABIERTA' && new Date(g.fecha) < hace48h;
  });

  // ── 6. Mermas pendientes ───────────────────────────────────
  var mermasPendientes = mermas.filter(function(m){ return m.estado === 'PENDIENTE'; });

  // ── 7. Auditorías pendientes ───────────────────────────────
  var audPendientes = auditorias.filter(function(a){
    return a.estado === 'PENDIENTE' || a.estado === 'ASIGNADA';
  });

  // ── 8. KPIs métricas ──────────────────────────────────────
  var hace30 = new Date(hoy); hace30.setDate(hoy.getDate() - 30);
  var mermasMes = mermas.filter(function(m){
    return new Date(m.fechaIngreso) >= hace30;
  }).reduce(function(acc, m){ return acc + (parseFloat(m.cantidadOriginal) || 0); }, 0);

  var envasadosMes = envasados.filter(function(e){
    return new Date(e.fecha) >= hace30 && e.estado === 'COMPLETADO';
  });
  var eficienciaPromedio = 0;
  if (envasadosMes.length > 0) {
    var sumEfic = envasadosMes.reduce(function(acc,e){ return acc + (parseFloat(e.eficienciaPct) || 0); }, 0);
    eficienciaPromedio = Math.round(sumEfic / envasadosMes.length * 10) / 10;
  }

  // Guías de salida del mes → rotación
  var salidasMes = guias.filter(function(g){
    return new Date(g.fecha) >= hace30 && g.tipo.startsWith('SALIDA');
  }).length;

  // ── 9. Resumen totales ─────────────────────────────────────
  var totalProductos  = productos.filter(function(p){ return p.estado === '1'; }).length;
  var totalBases      = productos.filter(function(p){ return p.esEnvasable === '1' && p.estado === '1'; }).length;
  var totalDerivados  = productos.filter(function(p){ return p.codigoProductoBase !== '' && p.estado === '1'; }).length;

  return {
    ok: true,
    data: {
      alertas: {
        vencimientosCriticos:  vencCriticos,
        vencimientosAlertas:   vencAlertas,
        stockBajoMinimo:       bajominimo,
        pendientesEnvasado:    pendientesEnvasado,
        preingresosPendientes: presPendientes,
        guiasAbiertasTardias:  guiasAbiertas,
        mermasPendientes:      mermasPendientes,
        auditoriasPendientes:  audPendientes
      },
      kpis: {
        totalProductosActivos: totalProductos,
        totalProductosBase:    totalBases,
        totalProductosDerivados: totalDerivados,
        mermasTotalMes:        mermasMes,
        eficienciaEnvasadoPct: eficienciaPromedio,
        salidasUltimos30dias:  salidasMes,
        lotesCriticos:         vencCriticos.length,
        lotesEnAlerta:         vencAlertas.length
      },
      contadores: {
        alertasTotal: vencCriticos.length + bajominimo.length + pendientesEnvasado.length
                    + presPendientes.length + mermasPendientes.length,
        criticos: vencCriticos.length + bajominimo.filter(function(b){
          return b.stockActual === 0;
        }).length
      },
      generadoEn: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
    }
  };
}

function _calcularPendientesEnvasado(productos, stockMap) {
  var pendientes = [];
  productos.forEach(function(p) {
    if (p.esEnvasable !== '1' || p.estado !== '1') return;
    // Para cada derivado que use este producto como base
    // No — buscamos productos derivados con stock bajo
  });

  // Productos WH envasados: codigoBarra empieza con WH + tienen codigoProductoBase
  productos.forEach(function(p) {
    if (!p.codigoProductoBase || p.codigoProductoBase === '') return;
    if (p.estado !== '1') return;
    // Solo productos envasados en almacén (prefijo WH en barcode o idProducto)
    var esWH = String(p.codigoBarra || p.idProducto || '').toUpperCase().indexOf('WH') === 0;
    if (!esWH) return;

    var stockDerivado = stockMap[p.idProducto] || 0;
    var minDerivado   = parseFloat(p.stockMinimo) || 0;
    var maxDerivado   = parseFloat(p.stockMaximo) || 0;
    // Incluir siempre que esté bajo el máximo (objetivo es llegar al max, mínimo es alerta)
    if (maxDerivado > 0 && stockDerivado >= maxDerivado) return;

    var stockBase = stockMap[p.codigoProductoBase] || 0;
    // factorConversionBase: unidades WH por 1 unidad de granel
    // factorConversion: fallback si aún no se migró el campo
    var factor = parseFloat(p.factorConversionBase) || parseFloat(p.factorConversion) || 1;
    var merma  = parseFloat(p.mermaEsperadaPct) || 0;
    var maxProducibles = Math.floor(stockBase * factor * (1 - merma / 100));

    if (stockBase <= 0) return; // Sin base, no se puede envasar

    var necesita = maxDerivado > 0
      ? Math.max(0, maxDerivado - stockDerivado)
      : Math.max(0, minDerivado - stockDerivado);

    pendientes.push({
      codigoDerivado:       p.idProducto,
      descripcion:          p.descripcion,
      codigoBase:           p.codigoProductoBase,
      stockDerivado:        stockDerivado,
      stockMinimo:          minDerivado,
      stockMaximo:          maxDerivado,
      necesitaProducir:     necesita,
      stockBase:            stockBase,
      factorConversionBase: factor,
      mermaEsperadaPct:     merma,
      maxProducibles:       maxProducibles,
      // Granel necesario para llenar hasta max
      granelNecesario:      maxDerivado > 0 ? Math.ceil(necesita / factor) : null,
      urgencia:             stockDerivado === 0 ? 'CRITICA' : (stockDerivado < minDerivado ? 'ALTA' : 'MEDIA')
    });
  });

  pendientes.sort(function(a,b) {
    if (a.urgencia === 'CRITICA' && b.urgencia !== 'CRITICA') return -1;
    if (b.urgencia === 'CRITICA' && a.urgencia !== 'CRITICA') return 1;
    return b.faltan - a.faltan;
  });

  return pendientes;
}
