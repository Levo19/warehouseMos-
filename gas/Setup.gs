// ============================================================
// warehouseMos — Setup.gs
// Ejecutar setupWarehouse() UNA sola vez para crear el
// Spreadsheet, hojas, cabeceras y datos de muestra.
// ============================================================

var SHEET_NAMES = {
  CONFIG:            'CONFIG',
  CATEGORIAS:        'CATEGORIAS',
  PRODUCTOS:         'PRODUCTOS',
  STOCK:             'STOCK',
  LOTES_VENCIMIENTO: 'LOTES_VENCIMIENTO',
  PROVEEDORES:       'PROVEEDORES',
  PREINGRESOS:       'PREINGRESOS',
  GUIAS:             'GUIAS',
  GUIA_DETALLE:      'GUIA_DETALLE',
  MERMAS:            'MERMAS',
  AUDITORIAS:        'AUDITORIAS',
  AJUSTES:           'AJUSTES',
  ENVASADOS:         'ENVASADOS',
  PRODUCTO_NUEVO:    'PRODUCTO_NUEVO',
  ZONAS:             'ZONAS',
  PERSONAL:          'PERSONAL',
  SESIONES:          'SESIONES',
  DESEMPENO:         'DESEMPENO',
  SYNC_LOG:          'SYNC_LOG',
  PICKUPS:           'PICKUPS'
};

var HEADERS = {
  CONFIG:            ['clave','valor','descripcion'],
  CATEGORIAS:        ['idCategoria','nombre','descripcion','icono','activo'],
  PRODUCTOS:         ['idProducto','skuBase','codigoBarra','descripcion','marca','idCategoria',
                      'unidad','stockMinimo','stockMaximo','precioCompra',
                      'Cod_Tributo','IGV_Porcentaje','Cod_SUNAT','Tipo_IGV',
                      'esEnvasable','codigoProductoBase','factorConversion',
                      'mermaEsperadaPct','foto','estado','fechaCreacion'],
  STOCK:             ['idStock','codigoProducto','cantidadDisponible','ultimaActualizacion'],
  LOTES_VENCIMIENTO: ['idLote','codigoProducto','fechaVencimiento','cantidadInicial',
                      'cantidadActual','idGuia','estado','fechaCreacion'],
  PROVEEDORES:       ['idProveedor','nombre','ruc','imagen','telefono','banco',
                      'numeroCuenta','cci','email','diaPedido','diaPago','diaEntrega',
                      'formaPago','plazoCredito','responsable','categoriaProducto','estado'],
  PREINGRESOS:       ['idPreingreso','fecha','idProveedor','usuario',
                      'monto','fotos','comentario','estado','idGuia'],
  GUIAS:             ['idGuia','tipo','fecha','usuario','idProveedor','idZona',
                      'numeroDocumento','comentario','montoTotal','estado','idPreingreso','foto'],
  GUIA_DETALLE:      ['idDetalle','idGuia','codigoProducto','cantidadEsperada',
                      'cantidadRecibida','precioUnitario','idLote','observacion'],
  MERMAS:            ['idMerma','fechaIngreso','origen','codigoProducto','idLote',
                      'cantidadOriginal','cantidadPendiente','motivo','usuario','idGuia','estado'],
  AUDITORIAS:        ['idAuditoria','fechaAsignacion','codigoProducto','usuario',
                      'stockSistema','stockFisico','diferencia','resultado',
                      'observacion','estado','fechaEjecucion'],
  AJUSTES:           ['idAjuste','codigoProducto','tipoAjuste','cantidadAjuste',
                      'motivo','usuario','idAuditoria','fecha'],
  ENVASADOS:         ['idEnvasado','codigoProductoBase','cantidadBase','unidadBase',
                      'codigoProductoEnvasado','unidadesEsperadas','unidadesProducidas',
                      'mermaReal','eficienciaPct','fecha','usuario','estado',
                      'idGuiaSalida','idGuiaIngreso','observacion'],
  PRODUCTO_NUEVO:    ['idProductoNuevo','idGuia','marca','descripcion','codigoBarra',
                      'idCategoria','unidad','cantidad','fechaVencimiento','foto',
                      'estado','usuario','fechaRegistro','aprobadoPor','fechaAprobacion'],
  ZONAS:             ['idZona','nombre','descripcion','responsable','estado'],
  PERSONAL:          ['idPersonal','nombre','apellido','pin','rol','tarifaHora',
                      'montoBase','estado','fechaIngreso','foto','color'],
  SESIONES:          ['idSesion','idPersonal','fechaInicio','horaInicio','fechaFin',
                      'horaFin','minutosActivos','estado'],
  SYNC_LOG:          ['localId','action','resultado','fecha'],
  DESEMPENO:         ['idDesempeno','idPersonal','idSesion','fecha',
                      'minutosActivos','horasTrabajadas',
                      'guiasCreadas','guiasCerradas',
                      'envasadosRegistrados','unidadesEnvasadas',
                      'mermasRegistradas','auditoriaEjecutadas',
                      'preingresoCreados','ajustesRealizados',
                      'totalActividades','actividadesPorHora',
                      'puntuacion','calificacion',
                      'montoBase','montoBonus','montoTotal','estado'],
  PICKUPS:           ['idPickup','fuente','estado','items','idZona',
                      'notas','creadoPor','fechaCreado','fechaAtendido']
};

// ============================================================
// FUNCIÓN PRINCIPAL — ejecutar una vez
// ============================================================
function setupWarehouse() {
  var ss = SpreadsheetApp.create('warehouseMos_DB');
  var ssId = ss.getId();

  // Guardar propiedades del script
  PropertiesService.getScriptProperties().setProperties({
    'SPREADSHEET_ID':      ssId,
    'MOS_SS_ID':           '',   // ← ID Spreadsheet ProyectoMOS (dejar vacío hasta que MOS esté listo)
    'PRINTNODE_API_KEY':   '',   // ← Ingresar en Project Settings > Script Properties
    'PRINTER_ETIQUETAS_ID':'',   // ← ID impresora etiquetas adhesivas
    'PRINTER_TICKETS_ID':  ''    // ← ID impresora tickets/reportes
  });

  Logger.log('Spreadsheet creado: ' + ssId);
  Logger.log('URL: ' + ss.getUrl());

  // Crear todas las hojas
  _crearHojas(ss);

  // Cargar datos de muestra
  _seedCategorias(ss);
  _seedZonas(ss);
  _seedProveedores(ss);
  _seedProductos(ss);
  _seedStock(ss);
  _seedLotes(ss);
  _seedGuias(ss);
  _seedGuiaDetalle(ss);
  _seedPreingresos(ss);
  _seedMermas(ss);
  _seedAuditorias(ss);
  _seedEnvasados(ss);
  _seedConfig(ss);
  _seedPersonal(ss);

  // Formatear cabeceras
  _formatearCabeceras(ss);

  Logger.log('✅ Setup completado. ID: ' + ssId);
  Logger.log('⚠️  Ir a Project Settings > Script Properties para ingresar:');
  Logger.log('   PRINTNODE_API_KEY, PRINTER_ETIQUETAS_ID, PRINTER_TICKETS_ID');

  return ssId;
}

// ============================================================
// Crear hojas
// ============================================================
function _crearHojas(ss) {
  // Renombrar la hoja por defecto
  var defaultSheet = ss.getSheets()[0];
  defaultSheet.setName(SHEET_NAMES.CONFIG);
  defaultSheet.getRange(1, 1, 1, HEADERS.CONFIG.length).setValues([HEADERS.CONFIG]);

  // Crear el resto
  Object.keys(SHEET_NAMES).forEach(function(key) {
    if (key === 'CONFIG') return;
    var name = SHEET_NAMES[key];
    var sheet = ss.insertSheet(name);
    if (HEADERS[key]) {
      sheet.getRange(1, 1, 1, HEADERS[key].length).setValues([HEADERS[key]]);
    }
  });
}

function _formatearCabeceras(ss) {
  ss.getSheets().forEach(function(sheet) {
    var lastCol = sheet.getLastColumn();
    if (lastCol < 1) return;
    var headerRange = sheet.getRange(1, 1, 1, lastCol);
    headerRange.setBackground('#1e3a5f')
               .setFontColor('#ffffff')
               .setFontWeight('bold')
               .setFontSize(10);
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, lastCol, 140);
  });
}

// ============================================================
// SEED — CONFIG
// ============================================================
function _seedConfig(ss) {
  var sheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
  var rows = [
    ['PRINTNODE_API_KEY',   '',          'API key de PrintNode para etiquetas'],
    ['PRINTNODE_PRINTER_ID','',          'ID de impresora en PrintNode'],
    ['DIAS_ALERTA_VENC',    '30',        'Días antes del vencimiento para alertar'],
    ['DIAS_ALERTA_VENC_CRITICO','7',     'Días críticos antes del vencimiento'],
    ['EMPRESA_NOMBRE',      'InversionMos', 'Nombre para etiquetas'],
    ['EMPRESA_RUC',         '',          'RUC para documentos'],
    ['TARIFA_CARRETA',      '10',        'Pago por carreta a cargadores (S/.)'],
    ['VERSION',             '1.0.0',     'Versión del sistema']
  ];
  sheet.getRange(2, 1, rows.length, 3).setValues(rows);
}

// ============================================================
// SEED — CATEGORIAS
// ============================================================
function _seedCategorias(ss) {
  var sheet = ss.getSheetByName(SHEET_NAMES.CATEGORIAS);
  var rows = [
    ['CAT001','Granos y Cereales','Arroz, quinua, cebada, maíz','🌾','1'],
    ['CAT002','Especias y Condimentos','Comino, pimienta, orégano, ají','🌶️','1'],
    ['CAT003','Azúcares y Endulzantes','Azúcar, chancaca, stevia','🍬','1'],
    ['CAT004','Sal y Minerales','Sal de mesa, sal gruesa','🧂','1'],
    ['CAT005','Aceites y Grasas','Aceite vegetal, manteca','🫙','1'],
    ['CAT006','Harinas','Harina de trigo, maíz, yuca','🌽','1'],
    ['CAT007','Menestras','Frijol, lenteja, garbanzo, arveja','🫘','1'],
    ['CAT008','Limpieza','Detergente, lejía, jabón','🧴','1']
  ];
  sheet.getRange(2, 1, rows.length, HEADERS.CATEGORIAS.length).setValues(rows);
}

// ============================================================
// SEED — ZONAS
// ============================================================
function _seedZonas(ss) {
  var sheet = ss.getSheetByName(SHEET_NAMES.ZONAS);
  var rows = [
    ['Z001','Zona Norte','Distribución sector norte','Carlos Ramos','1'],
    ['Z002','Zona Sur','Distribución sector sur','Ana Torres','1'],
    ['Z003','Zona Centro','Distribución sector centro','Luis Medina','1'],
    ['Z004','Zona Este','Distribución sector este','María Quispe','1'],
    ['ALMACEN','Almacén Central','Almacén principal InversionMos','Pedro Huanca','1']
  ];
  sheet.getRange(2, 1, rows.length, HEADERS.ZONAS.length).setValues(rows);
}

// ============================================================
// SEED — PROVEEDORES
// ============================================================
function _seedProveedores(ss) {
  var sheet = ss.getSheetByName(SHEET_NAMES.PROVEEDORES);
  var rows = [
    ['PROV001','Molinos El Sol SAC','20100456789','','987654321','BCP',
     '19200345678','00219200345678901','ventas@molinoselsol.pe',
     1,15,3,'CREDITO',30,'Juan Perez','Granos y Cereales','1'],
    ['PROV002','Especias del Norte EIRL','20234567890','','976543210','Interbank',
     '06012345678','00306012345678901','pedidos@especiasnorte.pe',
     2,16,4,'CONTADO',0,'Rosa Flores','Especias','1'],
    ['PROV003','Distribuidora Azucarera Lima SA','20345678901','','965432109','BBVA',
     '00110012345','00100110012345001','lima@distazucarera.pe',
     3,17,5,'CREDITO',45,'Miguel Castro','Azúcares','1']
  ];
  sheet.getRange(2, 1, rows.length, HEADERS.PROVEEDORES.length).setValues(rows);
}

// ============================================================
// SEED — PRODUCTOS
// idProducto | skuBase | codigoBarra | descripcion | marca | idCategoria |
// unidad | stockMinimo | stockMaximo | precioCompra |
// esEnvasable | codigoProductoBase | factorConversion |
// mermaEsperadaPct | foto | estado | fechaCreacion
//
// skuBase: vacío en bases/granel (no se venden en POS)
//          = codigoBarra en derivados (enlace con MosExpress/MOS)
// ============================================================
function _seedProductos(ss) {
  var sheet = ss.getSheetByName(SHEET_NAMES.PRODUCTOS);
  var hoy = new Date();
  var rows = [
    // ── BASES granel — no se venden en POS, skuBase vacío ──
    // cols: id,skuBase,codBarra,desc,marca,cat,unidad,stkMin,stkMax,precioCompra, Cod_Tributo,IGV_Pct,Cod_SUNAT,Tipo_IGV, esEnv,codBase,factor,merma,foto,estado,fecha
    ['P001','','','Arroz Extra Granel','El Sol','CAT001','SACO',
     10,100,85,'','','','','1','','','','','1',hoy],
    ['P002','','','Comino Molido Granel','Del Norte','CAT002','KG',
     5,50,18,'','','','','1','','','','','1',hoy],
    ['P003','','','Azúcar Rubia Granel','Paramonga','CAT003','SACO',
     8,80,120,'','','','','1','','','','','1',hoy],
    ['P004','','','Sal de Mesa Granel','Emsal','CAT004','SACO',
     5,60,12,'','','','','1','','','','','1',hoy],
    ['P005','','','Pimienta Negra Granel','Del Norte','CAT002','KG',
     2,30,45,'','','','','1','','','','','1',hoy],
    ['P006','','','Orégano Seco Granel','Andino','CAT002','KG',
     3,30,22,'','','','','1','','','','','1',hoy],
    ['P007','','','Frijol Canario Granel','Selecto','CAT007','SACO',
     6,60,95,'','','','','1','','','','','1',hoy],
    // ── DERIVADOS envasados — skuBase = codigoBarra (enlace MOS/MosExpress) ──
    // Arroz Extra: 1 saco 50kg → bolsas de 1kg = 50 uds (merma 2%)
    ['P101','7501234100011','7501234100011','Arroz Extra 1kg','El Sol','CAT001','BOLSA',
     200,2000,2.20,'','','','','0','P001',50,2,'','1',hoy],
    // Arroz Extra: 1 saco 50kg → bolsas de 5kg = 10 uds (merma 1%)
    ['P102','7501234100012','7501234100012','Arroz Extra 5kg','El Sol','CAT001','BOLSA',
     50,500,10.50,'','','','','0','P001',10,1,'','1',hoy],
    // Comino: 1 kg → bolsas de 100g = 10 uds (merma 3%)
    ['P201','7501234200011','7501234200011','Comino Molido 100g','Del Norte','CAT002','BOLSA',
     100,1000,2.30,'','','','','0','P002',10,3,'','1',hoy],
    // Comino: 1 kg → bolsas de 500g = 2 uds (merma 3%)
    ['P202','7501234200012','7501234200012','Comino Molido 500g','Del Norte','CAT002','BOLSA',
     50,500,10.80,'','','','','0','P002',2,3,'','1',hoy],
    // Azúcar: 1 saco 50kg → bolsas de 1kg = 50 uds (merma 1%)
    ['P301','7501234300011','7501234300011','Azúcar Rubia 1kg','Paramonga','CAT003','BOLSA',
     200,2000,2.90,'','','','','0','P003',50,1,'','1',hoy],
    // Azúcar: 1 saco 50kg → bolsas de 2kg = 25 uds (merma 1%)
    ['P302','7501234300012','7501234300012','Azúcar Rubia 2kg','Paramonga','CAT003','BOLSA',
     100,1000,5.60,'','','','','0','P003',25,1,'','1',hoy],
    // Sal: 1 saco 25kg → bolsas de 1kg = 25 uds (merma 0.5%)
    ['P401','7501234400011','7501234400011','Sal de Mesa 1kg','Emsal','CAT004','BOLSA',
     150,1500,0.55,'','','','','0','P004',25,0.5,'','1',hoy],
    // Pimienta: 1 kg → sobres 50g = 20 uds (merma 4%)
    ['P501','7501234500011','7501234500011','Pimienta Negra 50g','Del Norte','CAT002','SOBRE',
     80,800,2.70,'','','','','0','P005',20,4,'','1',hoy],
    // Orégano: 1 kg → sobres 25g = 40 uds (merma 5%)
    ['P601','7501234600011','7501234600011','Orégano Seco 25g','Andino','CAT002','SOBRE',
     100,1000,0.80,'','','','','0','P006',40,5,'','1',hoy],
    // Frijol: 1 saco 50kg → bolsas de 500g = 100 uds (merma 2%)
    ['P701','7501234700011','7501234700011','Frijol Canario 500g','Selecto','CAT007','BOLSA',
     200,2000,1.20,'','','','','0','P007',100,2,'','1',hoy]
  ];
  sheet.getRange(2, 1, rows.length, HEADERS.PRODUCTOS.length).setValues(rows);
}

// ============================================================
// SEED — STOCK
// ============================================================
function _seedStock(ss) {
  var sheet = ss.getSheetByName(SHEET_NAMES.STOCK);
  var now = new Date();
  var stockData = [
    // Bases
    ['STK001','P001',25,now],
    ['STK002','P002',18,now],
    ['STK003','P003',12,now],
    ['STK004','P004',20,now],
    ['STK005','P005',8,now],
    ['STK006','P006',5,now],
    ['STK007','P007',15,now],
    // Derivados
    ['STK101','P101',320,now],
    ['STK102','P102',45,now],
    ['STK201','P201',180,now],
    ['STK202','P202',30,now],
    ['STK301','P301',85,now],   // bajo mínimo → alerta
    ['STK302','P302',92,now],
    ['STK401','P401',420,now],
    ['STK501','P501',60,now],
    ['STK601','P601',95,now],
    ['STK701','P701',175,now]
  ];
  sheet.getRange(2, 1, stockData.length, HEADERS.STOCK.length).setValues(stockData);
}

// ============================================================
// SEED — LOTES VENCIMIENTO
// ============================================================
function _seedLotes(ss) {
  var sheet = ss.getSheetByName(SHEET_NAMES.LOTES_VENCIMIENTO);
  var now = new Date();
  var hoy = new Date();
  var en5dias  = new Date(hoy); en5dias.setDate(hoy.getDate() + 5);
  var en15dias = new Date(hoy); en15dias.setDate(hoy.getDate() + 15);
  var en45dias = new Date(hoy); en45dias.setDate(hoy.getDate() + 45);
  var en6meses = new Date(hoy); en6meses.setMonth(hoy.getMonth() + 6);
  var en1anio  = new Date(hoy); en1anio.setFullYear(hoy.getFullYear() + 1);

  var rows = [
    ['LOT001','P101',en5dias,  500,320,'G001','ACTIVO',now],  // ¡Crítico!
    ['LOT002','P201',en15dias, 300,180,'G002','ACTIVO',now],  // Próximo
    ['LOT003','P301',en45dias, 200,85, 'G003','ACTIVO',now],  // Alerta
    ['LOT004','P102',en6meses, 100,45, 'G001','ACTIVO',now],
    ['LOT005','P202',en6meses, 80, 30, 'G002','ACTIVO',now],
    ['LOT006','P401',en1anio,  600,420,'G004','ACTIVO',now],
    ['LOT007','P701',en6meses, 200,175,'G005','ACTIVO',now],
    ['LOT008','P501',en45dias, 100,60, 'G003','ACTIVO',now]
  ];
  sheet.getRange(2, 1, rows.length, HEADERS.LOTES_VENCIMIENTO.length).setValues(rows);
}

// ============================================================
// SEED — GUIAS
// ============================================================
function _seedGuias(ss) {
  var sheet = ss.getSheetByName(SHEET_NAMES.GUIAS);
  var hoy = new Date();
  var ayer = new Date(hoy); ayer.setDate(hoy.getDate() - 1);
  var hace3 = new Date(hoy); hace3.setDate(hoy.getDate() - 3);
  var rows = [
    ['G001','INGRESO_PROVEEDOR',hace3,'almacenero1','PROV001','ALMACEN','GR-2024-001',
     'Ingreso mensual arroz','2125.00','CERRADA','PI001',''],
    ['G002','INGRESO_PROVEEDOR',hace3,'almacenero1','PROV002','ALMACEN','GR-2024-002',
     'Ingreso especias','720.00','CERRADA','PI002',''],
    ['G003','SALIDA_ZONA',ayer,'almacenero2','','Z001','GS-2024-001',
     'Pedido zona norte','0','CERRADA','',''],
    ['G004','INGRESO_PROVEEDOR',hoy,'almacenero1','PROV003','ALMACEN','GR-2024-003',
     'Ingreso azúcar y sal','3850.00','ABIERTA','PI003',''],
    ['G005','INGRESO_PROVEEDOR',hoy,'almacenero2','PROV001','ALMACEN','GR-2024-004',
     'Reposición frijol','1425.00','CERRADA','',''],
    ['G006','SALIDA_ENVASADO',hoy,'envasador1','','ALMACEN','GS-2024-002',
     'Salida granel para envasado arroz','0','CERRADA','',''],
    ['G007','SALIDA_ENVASADO',hoy,'envasador1','','ALMACEN','GS-2024-003',
     'Salida granel para envasado comino','0','CERRADA','','']
  ];
  sheet.getRange(2, 1, rows.length, HEADERS.GUIAS.length).setValues(rows);
}

// ============================================================
// SEED — GUIA DETALLE
// ============================================================
function _seedGuiaDetalle(ss) {
  var sheet = ss.getSheetByName(SHEET_NAMES.GUIA_DETALLE);
  var rows = [
    ['DET001','G001','P001',25,25,85,'LOT004',''],
    ['DET002','G001','P101',500,500,2.20,'LOT001',''],
    ['DET003','G002','P002',20,18,18,'LOT002','Faltaron 2kg en saco'],
    ['DET004','G002','P201',300,300,2.30,'LOT002',''],
    ['DET005','G002','P202',80,80,10.80,'LOT005',''],
    ['DET006','G003','P101',100,100,0,'LOT001',''],
    ['DET007','G003','P201',80,80,0,'LOT002',''],
    ['DET008','G004','P003',12,0,120,'','Pendiente recepción'],
    ['DET009','G005','P007',15,15,95,'LOT007',''],
    ['DET010','G006','P001',5,5,0,'','Salida para envasado'],
    ['DET011','G007','P002',3,3,0,'','Salida para envasado comino']
  ];
  sheet.getRange(2, 1, rows.length, HEADERS.GUIA_DETALLE.length).setValues(rows);
}

// ============================================================
// SEED — PREINGRESOS
// ============================================================
function _seedPreingresos(ss) {
  var sheet = ss.getSheetByName(SHEET_NAMES.PREINGRESOS);
  var hace3 = new Date(); hace3.setDate(hace3.getDate() - 3);
  var hoy = new Date();
  var rows = [
    ['PI001',hace3,'PROV001','almacenero1','FAC-001-00234','2125.00',
     '','foto_pi001.jpg','Arroz mensual Molinos El Sol','ET-001','PROCESADO','G001'],
    ['PI002',hace3,'PROV002','almacenero1','FAC-002-00089','720.00',
     '','foto_pi002.jpg','Especias mensual','ET-002','PROCESADO','G002'],
    ['PI003',hoy,'PROV003','almacenero1','FAC-003-00412','3850.00',
     '','foto_pi003.jpg','Azúcar y sal — urgente','ET-003','PENDIENTE','']
  ];
  sheet.getRange(2, 1, rows.length, HEADERS.PREINGRESOS.length).setValues(rows);
}

// ============================================================
// SEED — MERMAS
// ============================================================
function _seedMermas(ss) {
  var sheet = ss.getSheetByName(SHEET_NAMES.MERMAS);
  var hace2 = new Date(); hace2.setDate(hace2.getDate() - 2);
  var hoy = new Date();
  var rows = [
    ['M001',hace2,'RECEPCION','P002','LOT002',20,2,
     'Saco roto en recepción — 2kg derramados','almacenero1','G002','PROCESADA'],
    ['M002',hoy,'VENCIMIENTO','P101','LOT001',50,50,
     'Lote próximo a vencer — retirar del almacén','almacenero2','','PENDIENTE'],
    ['M003',hoy,'ENVASADO','P002','LOT002',3,0.09,
     'Merma normal proceso envasado','envasador1','G007','PROCESADA']
  ];
  sheet.getRange(2, 1, rows.length, HEADERS.MERMAS.length).setValues(rows);
}

// ============================================================
// SEED — AUDITORIAS
// ============================================================
function _seedAuditorias(ss) {
  var sheet = ss.getSheetByName(SHEET_NAMES.AUDITORIAS);
  var ayer = new Date(); ayer.setDate(ayer.getDate() - 1);
  var hoy = new Date();
  var manana = new Date(); manana.setDate(manana.getDate() + 1);
  var rows = [
    ['AUD001',ayer,'P101','almacenero2',420,320,-100,
     'DIFERENCIA','Revisar guías de salida — falta cuadrar 100 unidades','EJECUTADA',ayer],
    ['AUD002',hoy,'P002','almacenero1',18,18,0,
     'OK','Sin diferencias','EJECUTADA',hoy],
    ['AUD003',hoy,'P003','almacenero1',12,0,0,
     '','Guía G004 aún no cerrada','PENDIENTE',''],
    ['AUD004',manana,'P701','','15,0,0',0,0,
     '','Por ejecutar mañana','ASIGNADA','']
  ];
  sheet.getRange(2, 1, rows.length, HEADERS.AUDITORIAS.length).setValues(rows);
}

// ============================================================
// SEED — ENVASADOS
// ============================================================
function _seedEnvasados(ss) {
  var sheet = ss.getSheetByName(SHEET_NAMES.ENVASADOS);
  var hoy = new Date();
  var ayer = new Date(); ayer.setDate(ayer.getDate() - 1);
  var rows = [
    // 5 sacos arroz → 250 bolsas 1kg esperadas (factor 50, merma 2%) → prod real 245
    ['ENV001','P001',5,'SACO','P101',245,245,5,98.0,ayer,'envasador1','COMPLETADO','G006','',''],
    // 3 kg comino → 60 sobres 100g esperados (factor 10, merma 3%) → prod real 58
    ['ENV002','P002',3,'KG','P201',58,58,2,96.7,hoy,'envasador1','COMPLETADO','G007','',''],
    // 3 kg comino → 6 bolsas 500g (factor 2, merma 3%) → prod real 6
    ['ENV003','P002',3,'KG','P202',5.82,6,0,103.0,hoy,'envasador1','COMPLETADO','G007','','Sobraron 30g de otro lote']
  ];
  sheet.getRange(2, 1, rows.length, HEADERS.ENVASADOS.length).setValues(rows);
}

// ============================================================
// Agregar tablas de Personal al Spreadsheet existente
// Ejecutar UNA vez si ya corriste setupWarehouse() antes
// ============================================================
function setupAgregarPersonal() {
  var ss = SpreadsheetApp.openById(
    PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID')
  );

  var tablasFaltantes = ['PERSONAL', 'SESIONES', 'DESEMPENO'];
  var existentes = ss.getSheets().map(function(s){ return s.getName(); });

  tablasFaltantes.forEach(function(nombre) {
    if (existentes.indexOf(nombre) >= 0) {
      Logger.log('⏭ Ya existe: ' + nombre);
      return;
    }
    var sheet = ss.insertSheet(nombre);
    var hdrs  = HEADERS[nombre];
    if (hdrs) sheet.getRange(1, 1, 1, hdrs.length).setValues([hdrs]);

    // Formato cabecera
    sheet.getRange(1, 1, 1, hdrs.length)
         .setBackground('#1e3a5f').setFontColor('#ffffff')
         .setFontWeight('bold').setFontSize(10);
    sheet.setFrozenRows(1);
    Logger.log('✅ Creada: ' + nombre);
  });

  // Seed personal + config extra
  _seedPersonal(ss);

  Logger.log('✅ Tablas de personal listas. Revisa el Sheet.');
}

// ============================================================
// SEED — PERSONAL + CONFIG adicional
// ============================================================
function _seedPersonal(ss) {
  // PERSONAL — 3 operadores de almacén
  var sheetP = ss.getSheetByName(SHEET_NAMES.PERSONAL);
  var hoy = new Date();
  // idPersonal | nombre | apellido | pin | rol | tarifaHora | montoBase | estado | fechaIngreso | foto | color
  var personal = [
    ['OP001','Carlos','Ramos','1234','ALMACENERO',5.00,1200,'1',hoy,'','#3b82f6'],
    ['OP002','Ana','Torres','5678','ENVASADOR', 4.50,1100,'1',hoy,'','#22c55e'],
    ['OP003','Luis','Medina','9012','ALMACENERO',5.00,1200,'1',hoy,'','#f59e0b']
  ];
  sheetP.getRange(2, 1, personal.length, HEADERS.PERSONAL.length).setValues(personal);

  // Agregar config adicional en CONFIG
  var sheetC = ss.getSheetByName(SHEET_NAMES.CONFIG);
  var extraConfig = [
    ['HORA_CIERRE_FORZADO', '22:00', 'Hora de cierre automático de turno (HH:MM)'],
    ['MIN_INACTIVIDAD_BLOQUEO', '5',  'Minutos sin actividad para bloquear pantalla'],
    ['BONUS_PUNTUACION_MIN',   '8',   'Puntuación mínima para bonus (actividades/hora)'],
    ['BONUS_PORCENTAJE',       '10',  'Porcentaje de bonus sobre monto base al superar meta']
  ];
  var lastRow = sheetC.getLastRow();
  sheetC.getRange(lastRow + 1, 1, extraConfig.length, 3).setValues(extraConfig);
}
