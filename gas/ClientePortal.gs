// ============================================================
// warehouseMos — ClientePortal.gs
// Portal público de pedidos para clientes (pedido.html)
//
// Endpoints expuestos (vía router en Code.gs):
//   - clienteInfo            → datos del cliente por token (nombre bonito)
//   - clienteRegistrar       → alta/edición desde el modal admin O auto-alta
//   - clienteListar          → lista para el modal admin
//   - clienteRecibirPedido   → recibe foto/audio/excel/texto + procesa IA
//   - clienteConfirmarPedido → cliente confirma → crea lista sombra real
//   - clienteEstadoPedido    → estado + timeline para el cliente
//   - clienteInboxPolling    → WH consulta cada 20s nuevos pedidos
//
// Hojas que crea automáticamente:
//   - Clientes            [token, nombre, telefono, tipo, premium, fechaAlta, ultimoPedido]
//   - PedidosCliente      [idPedido, token, ts, estado, idListaSombra, totalEstimado, notas]
//   - PedidosClienteItems [idPedido, idx, nombre, cantidad, unidad, precioEst, duda]
//   - PedidosClienteAdj   [idPedido, tipo, nombreArchivo, urlDrive, ts]
//
// Reusa _llamarClaude e analizarListaSombra de IA.gs.
// ============================================================

var CLI_FOLDER_NAME = 'PedidosClientes_WH';
var CLI_TIMELINE = ['Recibido','Cotizando','Despachando','Listo','En camino'];

// ── Helpers hojas ──────────────────────────────────────────
function _cliSheet(name, headers) {
  var ss = getSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold');
    // Forzar columnas de IDs como texto ([[feedback_codigobarra_texto]])
    sh.getRange(1, 1, sh.getMaxRows(), 1).setNumberFormat('@');
  }
  return sh;
}
function _shClientes()      { return _cliSheet('Clientes',      ['token','nombre','telefono','tipo','premium','fechaAlta','ultimoPedido']); }
function _shPedidos()       { return _cliSheet('PedidosCliente',['idPedido','token','ts','estado','idListaSombra','totalEstimado','notas']); }
function _shPedidoItems()   { return _cliSheet('PedidosClienteItems', ['idPedido','idx','nombre','cantidad','unidad','precioEst','duda']); }
function _shPedidoAdj()     { return _cliSheet('PedidosClienteAdj',   ['idPedido','tipo','nombreArchivo','urlDrive','ts']); }

function _cliBuscarToken(token) {
  if (!token) return null;
  var sh = _shClientes(), data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(token)) {
      return {
        row: i + 1,
        token: data[i][0], nombre: data[i][1], telefono: data[i][2],
        tipo: data[i][3], premium: data[i][4] === true || data[i][4] === 'TRUE',
        fechaAlta: data[i][5], ultimoPedido: data[i][6]
      };
    }
  }
  return null;
}

function _slugToken(nombre) {
  var s = String(nombre || 'CLI').toUpperCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
  s = s.replace(/[^A-Z0-9]/g, '').substring(0, 8);
  if (!s) s = 'CLI';
  return s + Math.floor(Math.random() * 900 + 100); // 3 dígitos
}

// ── 1. INFO de cliente (frontend lo llama al cargar) ───────
function clienteInfo(params) {
  var token = String(params && params.token || '').toUpperCase();
  if (!token) return { ok: true, data: { token: '', nombre: 'Cliente', existe: false } };
  var c = _cliBuscarToken(token);
  if (!c) return { ok: true, data: { token: token, nombre: '', existe: false } };
  return { ok: true, data: {
    token: c.token, nombre: c.nombre, telefono: c.telefono,
    tipo: c.tipo, premium: c.premium, existe: true
  }};
}

// ── 2. REGISTRAR cliente (manual o auto desde portal) ──────
function clienteRegistrar(params) {
  var nombre   = String(params.nombre || '').trim();
  var telefono = String(params.telefono || '').trim();
  var tipo     = String(params.tipo || 'minorista');
  var premium  = !!params.premium;
  if (!nombre) return { ok: false, error: 'NOMBRE_REQUERIDO' };

  var token = String(params.token || '').toUpperCase();
  var sh = _shClientes();

  // Si viene token y existe → update
  if (token) {
    var c = _cliBuscarToken(token);
    if (c) {
      sh.getRange(c.row, 2, 1, 4).setValues([[nombre, telefono, tipo, premium]]);
      return { ok: true, data: { token: token, nombre: nombre, actualizado: true } };
    }
  }
  // Alta nueva → generar token único
  if (!token) token = _slugToken(nombre);
  while (_cliBuscarToken(token)) token = _slugToken(nombre);
  sh.appendRow([token, nombre, telefono, tipo, premium, new Date(), '']);
  return { ok: true, data: { token: token, nombre: nombre, creado: true } };
}

// ── 3. LISTAR clientes (para el modal admin de WH) ─────────
function clienteListar(params) {
  var chk = _requireAdmin(params || {});            // gate admin: la PII de clientes (nombre/teléfono) solo para admin (fix C6)
  if (!chk.ok) return { ok: false, error: chk.error || 'clave admin requerida' };
  var sh = _shClientes(), data = sh.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < data.length; i++) {
    out.push({
      token: data[i][0], nombre: data[i][1], telefono: data[i][2],
      tipo: data[i][3], premium: data[i][4] === true || data[i][4] === 'TRUE',
      fechaAlta: data[i][5], ultimoPedido: data[i][6]
    });
  }
  return { ok: true, data: { clientes: out, total: out.length } };
}

// ── 4. RECIBIR PEDIDO (foto/audio/excel/texto) ─────────────
function clienteRecibirPedido(params) {
  var token = String(params.token || '').toUpperCase() || 'ANON';
  var cli   = _cliBuscarToken(token);
  if (!cli && token !== 'ANON') {
    // Auto-alta silenciosa si llega un token desconocido
    clienteRegistrar({ token: token, nombre: token });
    cli = _cliBuscarToken(token);
  }
  var nombreCli = cli ? cli.nombre : 'Cliente anónimo';

  var adjuntos = Array.isArray(params.adjuntos) ? params.adjuntos : [];
  var textoUser = String(params.texto || '').trim();

  // [Lote4 · M3-WH] sufijo aleatorio: 'PC'+getTime() colisionaba si dos pedidos
  // caían en el mismo ms → items/adjuntos se mezclaban entre pedidos (keyed por idPedido).
  var idPedido = 'PC' + new Date().getTime() + Math.floor(Math.random() * 1000);
  var folder   = _cliGetFolder(token);

  // ── Guardar adjuntos en Drive y construir texto consolidado para la IA ──
  var textosParaIA = [];
  if (textoUser) textosParaIA.push(textoUser);
  var adjMeta = [];

  adjuntos.forEach(function(a) {
    try {
      var bytes  = _cliB64ToBytes(a.b64);
      var mime   = a.mime || 'application/octet-stream';
      var nombre = a.nombre || (a.tipo + '_' + Date.now());
      var blob   = Utilities.newBlob(bytes, mime, idPedido + '_' + nombre);
      var f      = folder.createFile(blob);
      adjMeta.push({ tipo: a.tipo, nombre: nombre, url: f.getUrl() });

      // Si es imagen → llamar a Claude Vision para extraer items
      if (a.tipo === 'foto' && mime.indexOf('image/') === 0) {
        var ext = _cliAnalizarImagenLista(a.b64, mime);
        if (ext.ok && ext.text) textosParaIA.push('[de imagen ' + nombre + ']\n' + ext.text);
      }
      // Audio / Excel → quedan como adjunto, marcados "needs review"
    } catch(e) {
      adjMeta.push({ tipo: a.tipo, nombre: a.nombre || '?', url: '', error: e.message });
    }
  });

  // ── Pasar el texto consolidado por la IA de listas (la que ya existe) ──
  var textoFinal = textosParaIA.join('\n');
  var items = [];
  var notaProc = '';
  if (textoFinal) {
    var res = analizarListaSombra({ texto: textoFinal });
    if (res.ok && res.data && Array.isArray(res.data.items) && res.data.items.length > 0) {
      items = res.data.items;
    } else {
      // Fallback regex: listas estructuradas tipo "N unidad nombre" — IA no respondió bien
      items = _cliFallbackParse(textoFinal);
      notaProc = res.ok ? 'fallback regex (IA sin items)' : ('IA falló: ' + (res.error || '?'));
    }
  }
  // Marcar audio/excel para review humano
  var hayAdjuntoNoIA = adjMeta.some(function(a) { return a.tipo === 'audio' || a.tipo === 'excel'; });
  if (hayAdjuntoNoIA && items.length === 0) {
    items.push({ nombre: 'Adjunto sin procesar IA — revisar manualmente', cantidad: 1, unidad: 'rev', duda: 'audio/excel adjunto' });
  }

  // Normalizar a la shape que espera el frontend (solo nombre+cantidad+unidad)
  var itemsFront = items.map(function(it) {
    return {
      nombre: String(it.nombre || '').toUpperCase().trim(),
      cantidad: Math.round((parseFloat(it.cantidad) || 0) * 10) / 10,
      unidad: it.unidad || 'unidad',
      codigoVisto: it.codigoVisto || '',
      duda: it.duda || ''
    };
  }).filter(function(it) { return it.nombre && it.cantidad > 0; });

  // Guardar registros
  _shPedidos().appendRow([idPedido, token, new Date(), 'PREVIEW', '', 0, notaProc || '']);
  itemsFront.forEach(function(it, i) {
    _shPedidoItems().appendRow([idPedido, i, it.nombre, it.cantidad, it.unidad, 0, it.duda]);
  });
  adjMeta.forEach(function(a) {
    _shPedidoAdj().appendRow([idPedido, a.tipo, a.nombre, a.url || '', new Date()]);
  });

  return { ok: true, data: { idPedido: idPedido, items: itemsFront, nombreCliente: nombreCli, nota: notaProc, textoOriginal: textoFinal } };
}

// ── Fallback regex parser para listas tipo "N unidad nombre" ─
// Cuando la IA no responde o devuelve 0 items, intentamos un parser
// simple por línea: número + unidad opcional + nombre.
// Ej: "10 ajinomoto kilo" → { nombre:'AJINOMOTO', cantidad:10, unidad:'kilo' }
function _cliFallbackParse(texto) {
  var lines = String(texto || '').split(/\r?\n/);
  var items = [];
  var unidades = ['kg','kilo','kilos','gr','gramo','gramos','lt','litro','litros','saco','sacos','caja','cajas','paquete','paquetes','unidad','unidades','und','u','botella','botellas','lata','latas','bolsa','bolsas','docena','docenas','tarro','tarros','frasco','frascos','pack'];
  lines.forEach(function(line) {
    var l = String(line).trim();
    if (!l) return;
    var m = l.match(/^(\d+(?:[.,]\d+)?)\s+(.+)$/);
    if (!m) return;
    var qty = parseFloat(m[1].replace(',', '.'));
    if (!(qty > 0)) return;
    var rest = m[2].trim();
    // Buscar unidad al inicio o al final
    var palabras = rest.split(/\s+/);
    var unidad = 'unidad';
    if (palabras.length >= 2) {
      var primera = palabras[0].toLowerCase().replace(/[.,]$/, '');
      var ultima  = palabras[palabras.length - 1].toLowerCase().replace(/[.,]$/, '');
      if (unidades.indexOf(primera) >= 0) {
        unidad = primera;
        palabras.shift();
      } else if (unidades.indexOf(ultima) >= 0) {
        unidad = ultima;
        palabras.pop();
      }
    }
    var nombre = palabras.join(' ').trim().toUpperCase();
    if (nombre) items.push({ nombre: nombre, cantidad: qty, unidad: unidad });
  });
  return items;
}

// ── 5. CONFIRMAR PEDIDO → crea lista sombra real ───────────
function clienteConfirmarPedido(params) {
  var idPedido = String(params.idPedido || '');
  var items    = Array.isArray(params.items) ? params.items : [];
  if (!idPedido) return { ok: false, error: 'ID_FALTANTE' };

  var shP = _shPedidos(), data = shP.getDataRange().getValues();
  var row = -1, tokenCli = '';
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === idPedido) { row = i + 1; tokenCli = data[i][1]; break; }
  }
  if (row < 0) return { ok: false, error: 'PEDIDO_NO_ENCONTRADO' };

  // IDOR guard: quien confirma debe ser el DUEÑO del pedido (su token == token del pedido). (fix C6)
  // Sin esto, cualquiera con un idPedido enumerable (PC+timestamp) confirmaba pedidos ajenos.
  // Normaliza vacío→'ANON' espejando clienteRecibirPedido (preserva pedidos anónimos sin ?c=).
  var tokenReq = String(params.token || '').toUpperCase() || 'ANON';
  if (String(tokenCli).toUpperCase() !== tokenReq) {
    return { ok: false, error: 'PEDIDO_NO_ENCONTRADO' };   // no distinguir: no revelar existencia del pedido
  }

  var cli = _cliBuscarToken(tokenCli);
  var nombreCli = cli ? cli.nombre : tokenCli;

  // Crear lista sombra usando la función que ya existe.
  // OJO: crearListaSombra requiere `usuario`. Pasamos el nombre del cliente
  // como usuario para que el almacenero vea "lista de Don Pepe" en su panel.
  // Se crea como DISPONIBLE (compartida) para que cualquier almacenero pueda
  // tomarla. El matching real contra el catálogo lo hace el operador al escanear.
  var itemsParaLista = items.map(function(it) {
    return {
      nombre: String(it.nombre || '').toUpperCase().trim(),
      cantidad: parseFloat(it.cantidad) || 1,
      unidad: it.unidad || 'unidad',
      codigoVisto: it.codigoVisto || ''
    };
  });
  var idListaSombra = '';
  try {
    var ls = crearListaSombra({
      usuario:   'Cliente: ' + nombreCli,
      idLista:   'LSCLI' + idPedido.substring(2), // prefijo CLI para distinguir
      items:     itemsParaLista,
      compartir: true,
      nota:      'Pedido portal cliente — ' + nombreCli + ' (' + tokenCli + ') · #' + idPedido
    });
    idListaSombra = (ls && ls.data && ls.data.idLista) || (ls && ls.idLista) || '';
  } catch(e) {
    // Si crearListaSombra falla, igual confirmamos el pedido — el operador puede crearla a mano
  }

  // Actualizar pedido
  shP.getRange(row, 4).setValue('CONFIRMADO');
  shP.getRange(row, 5).setValue(idListaSombra);
  // Actualizar último pedido del cliente
  if (cli) _shClientes().getRange(cli.row, 7).setValue(new Date());

  // Marcar en inbox para que el polling de WH lo detecte y suene la alerta
  _cliInboxAgregar({ idPedido: idPedido, cliente: nombreCli, token: tokenCli, items: items.length, idListaSombra: idListaSombra });

  return { ok: true, data: { idPedido: idPedido, idListaSombra: idListaSombra, eta: 25 } };
}

// ── 6. ESTADO del pedido (timeline) ────────────────────────
function clienteEstadoPedido(params) {
  var idPedido = String(params.idPedido || '');
  var shP = _shPedidos(), data = shP.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === idPedido) {
      // [Lote2-C · fix M1 revisión 2026-06-12] IDOR guard — mismo patrón que
      // clienteConfirmarPedido (fix C6): el token del request debe ser el del
      // DUEÑO del pedido. idPedido es enumerable ('PC'+timestamp); sin esto
      // cualquiera leía el estado de pedidos ajenos. Mismo error opaco (no
      // revelar existencia). Sin caller activo en pedido.html hoy → no rompe nada.
      var tokenReqE = String(params.token || '').toUpperCase() || 'ANON';
      if (String(data[i][1] || '').toUpperCase() !== tokenReqE) {
        return { ok: false, error: 'PEDIDO_NO_ENCONTRADO' };
      }
      var estado = data[i][3] || 'PREVIEW';
      var pasos = CLI_TIMELINE.map(function(p, idx) {
        var idxEstado = ({ 'PREVIEW':0,'CONFIRMADO':1,'EN_DESPACHO':2,'LISTO':3,'EN_CAMINO':4,'ENTREGADO':5 })[estado] || 0;
        return { paso: p, done: idx < idxEstado, now: idx === idxEstado };
      });
      return { ok: true, data: { idPedido: idPedido, estado: estado, timeline: pasos } };
    }
  }
  return { ok: false, error: 'PEDIDO_NO_ENCONTRADO' };
}

// ── 7. INBOX POLLING (WH consulta cada 20s) ────────────────
function clienteInboxPolling(params) {
  var desde = parseFloat(params && params.desde) || 0; // timestamp ms
  var inbox = _cliInboxLeer();
  var nuevos = inbox.filter(function(it) { return it.ts > desde; });
  return { ok: true, data: { nuevos: nuevos, ahora: new Date().getTime() } };
}

// ── Inbox interna (PropertiesService) ──────────────────────
// Buffer rotativo de últimos 20 pedidos confirmados.
function _cliInboxAgregar(item) {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty('CLI_INBOX') || '[]';
  var arr = [];
  try { arr = JSON.parse(raw); } catch(e) { arr = []; }
  item.ts = new Date().getTime();
  arr.unshift(item);
  arr = arr.slice(0, 20);
  props.setProperty('CLI_INBOX', JSON.stringify(arr));
}
function _cliInboxLeer() {
  var raw = PropertiesService.getScriptProperties().getProperty('CLI_INBOX') || '[]';
  try { return JSON.parse(raw); } catch(e) { return []; }
}

// ── Drive folder por cliente ───────────────────────────────
function _cliGetFolder(token) {
  var root = _cliGetOrCreateFolder(DriveApp.getRootFolder(), CLI_FOLDER_NAME);
  return _cliGetOrCreateFolder(root, token || 'ANON');
}
function _cliGetOrCreateFolder(parent, name) {
  var it = parent.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return parent.createFolder(name);
}

// ── Base64 (data URL) → bytes ──────────────────────────────
function _cliB64ToBytes(b64) {
  var s = String(b64 || '');
  var i = s.indexOf('base64,');
  if (i >= 0) s = s.substring(i + 7);
  return Utilities.base64Decode(s);
}

// ── Analizar imagen con Claude Vision → texto de items ─────
function _cliAnalizarImagenLista(b64, mime) {
  var s = String(b64 || '');
  var i = s.indexOf('base64,');
  if (i >= 0) s = s.substring(i + 7);
  var system = [
    'Recibes la foto de una LISTA de productos escrita a mano o impresa.',
    'Extrae los productos UNO POR LÍNEA en este formato exacto:',
    'CANTIDAD UNIDAD NOMBRE',
    'Ejemplo:',
    '2 saco arroz costeño',
    '6 lt aceite primor',
    '4 unidad coca cola 3L',
    '',
    'Solo escribe las líneas, sin encabezados, sin comentarios.'
  ].join('\n');
  var res = _llamarClaude({
    max_tokens: 2048,
    system: system,
    messages: [{
      role: 'user',
      content: [
        { type: 'image', source: { type: 'base64', media_type: mime || 'image/jpeg', data: s } },
        { type: 'text', text: 'Transcribe esta lista en el formato indicado.' }
      ]
    }]
  });
  return { ok: !!res.ok, text: res.text || '' };
}
