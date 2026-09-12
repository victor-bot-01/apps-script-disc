/************** CONFIG **************/
const CONFIG = {
  API_KEY: "victorsevero15", // a mesma que você usa no Render (SHEETS_API_KEY)

  SHEET_NAME: "Conferencia",
  LIST_SHEET_NAME: "Lista",

  HEADER_ROW: 1,
  DATA_START_ROW: 2,
};

/************** WEBAPP ENTRY **************/
function doGet(e) {
  try {
    const key = (e.parameter.key || "").trim();
    if (!isValidKey_(key)) return json_({ ok: false, error: "Invalid key" });

    const action = (e.parameter.action || "").trim();

    if (action === "list_pending") {
      const out = listPending_();
      return json_({ ok: true, ...out });
    }

    if (action === "list_confirmed") {
      const out = listConfirmed_();
      return json_({ ok: true, ...out });
    }

    // ✅ lista confirmados com OutrosMessageId (coluna M) agregado
    if (action === "list_confirmed_with_others") {
      const out = listConfirmedWithOthers_();
      return json_({ ok: true, ...out });
    }

    // ✅ lista todos os IDs (únicos) presentes na coluna M (OutrosMessageId)
    if (action === "list_outros_message_ids") {
      const out = listOutrosMessageIds_();
      return json_({ ok: true, ...out });
    }

    // ✅ pesquisa na aba Lista (Lista completa)
    if (action === "search_lista_completa") {
      const q = String(e.parameter.q || "").trim();
      const out = searchListaCompleta_(q);
      return json_({ ok: true, ...out });
    }

    // ✅ obter item da Lista por idKey (usaremos o número da linha)
    if (action === "get_lista_item_by_idkey") {
      const idKey = String(e.parameter.idKey || "").trim();
      const out = getListaItemByIdKey_(idKey);
      return json_({ ok: true, ...out });
    }

    // ✅ obter order completo pelo DiscordMessageId original (Conferencia!I)
    if (action === "get_order_by_message_id") {
      const messageId = String(e.parameter.messageId || "").trim();
      const out = getOrderByMessageId_(messageId);
      return json_({ ok: true, ...out });
    }

    // ✅ NOVO (para /pendentes)
    // GET ?action=list_pendentes_status_vazio&limit=10
    if (action === "list_pendentes_status_vazio") {
      const limit = Math.max(1, Math.min(50, parseInt(String(e.parameter.limit || "10"), 10) || 10));
      const out = listPendentesStatusVazio_(limit);
      return json_({ ok: true, ...out });
    }

    // ✅ NOVO (para /codigo)
    // GET ?action=search_pedidos_conferencia_contains&q=YK3H
    if (action === "search_pedidos_conferencia_contains") {
      const q = String(e.parameter.q || "").trim();
      const out = searchPedidosConferenciaContains_(q);
      return json_({ ok: true, ...out });
    }

    // ✅ NOVO (para /codigo)
    // GET ?action=get_order_by_pedido&pedido=ORDBKF5SN58NYK3H
    if (action === "get_order_by_pedido") {
      const pedido = String(e.parameter.pedido || "").trim();
      const out = getOrderByPedido_(pedido);
      return json_({ ok: true, ...out });
    }

    return json_({ ok: false, error: "Unknown action" });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData?.contents || "{}");
    const key = String(body.key || "").trim();
    if (!isValidKey_(key)) return json_({ ok: false, error: "Invalid key" });

    const action = String(body.action || "").trim();

    if (action === "set_message_id") {
      const pedido = String(body.pedido || "").trim();
      const messageId = String(body.messageId || "").trim();
      if (!pedido || !messageId) return json_({ ok: false, error: "pedido/messageId required" });

      setMessageId_(pedido, messageId);
      return json_({ ok: true });
    }

    if (action === "set_item_status") {
      const itemKey = String(body.itemKey || "").trim();
      const status = String(body.status || "").trim(); // TENHO / FALTA / FALTA - obs
      const conferidoPor = String(body.conferidoPor || "").trim();
      const conferidoEmISO = String(body.conferidoEmISO || "").trim();

      if (!itemKey || !status) return json_({ ok: false, error: "itemKey/status required" });

      setItemStatus_(itemKey, status, conferidoPor, conferidoEmISO);
      return json_({ ok: true });
    }

    if (action === "delete_order_by_message_id") {
      const messageId = String(body.messageId || "").trim();
      if (!messageId) return json_({ ok: false, error: "messageId required" });

      const deletedRows = deleteOrderByMessageId_(messageId);
      return json_({ ok: true, deletedRows });
    }

    // ✅ append do novo messageId na coluna OutrosMessageId (Conferencia!M)
    if (action === "append_outros_message_id") {
      const originalMessageId = String(body.originalMessageId || "").trim();
      const newMessageId = String(body.newMessageId || "").trim();
      if (!originalMessageId || !newMessageId) {
        return json_({ ok: false, error: "originalMessageId/newMessageId required" });
      }

      const updatedRows = appendOutrosMessageId_(originalMessageId, newMessageId);
      return json_({ ok: true, updatedRows });
    }

    // ✅ remove 1 id específico de dentro da coluna M (OutrosMessageId)
    if (action === "clear_outros_message_id") {
      const messageId = String(body.messageId || "").trim();
      if (!messageId) return json_({ ok: false, error: "messageId required" });

      const cleared = clearOutrosMessageId_(messageId);
      return json_({ ok: true, cleared });
    }

    // ✅ NOVO (para /pendentes e /codigo)
    // POST { action:"append_outros_message_id_by_pedido", pedido:"123", newMessageId:"..." }
    if (action === "append_outros_message_id_by_pedido") {
      const pedido = String(body.pedido || "").trim();
      const newMessageId = String(body.newMessageId || "").trim();
      if (!pedido || !newMessageId) {
        return json_({ ok: false, error: "pedido/newMessageId required" });
      }

      const updated = appendOutrosMessageIdByPedido_(pedido, newMessageId);
      return json_({ ok: true, updated });
    }

    // ✅ NOVO (para /codigo)
    // POST { action:"append_outros_message_id_m2", newMessageId:"..." }
    if (action === "append_outros_message_id_m2") {
      const newMessageId = String(body.newMessageId || "").trim();
      if (!newMessageId) {
        return json_({ ok: false, error: "newMessageId required" });
      }

      const updated = appendOutrosMessageIdM2_(newMessageId);
      return json_({ ok: true, updated });
    }

    // ✅ NOVO: /importacao (dispara o Apps Script "Sincronização" dentro do mesmo projeto)
    if (action === "importacao") {
      const out = runImportacaoComLock_();
      return json_({ ok: true, ...out });
    }

    return json_({ ok: false, error: "Unknown action" });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

/************** ACTIONS **************/
function listPending_() {
  const sh = getSheetByName_(CONFIG.SHEET_NAME);
  const { idx } = getHeaderMap_(sh);

  assertCols_(idx, ["pedido","cliente","marketplace","produto","qtd","itemkey","statuspedido","discordmessageid","status"]);

  const colPedido = idx["pedido"];
  const colCliente = idx["cliente"];
  const colMarketplace = idx["marketplace"];
  const colProduto = idx["produto"];
  const colQtd = idx["qtd"];
  const colItemKey = idx["itemkey"];
  const colStatusPedido = idx["statuspedido"];
  const colDiscordMessageId = idx["discordmessageid"];
  const colStatusItem = idx["status"];

  const lastRow = sh.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return { orders: [] };

  const numRows = lastRow - CONFIG.DATA_START_ROW + 1;
  const data = sh.getRange(CONFIG.DATA_START_ROW, 1, numRows, sh.getLastColumn()).getValues();

  const map = new Map();
  for (let r = 0; r < data.length; r++) {
    const row = data[r];
    const pedido = String(row[colPedido] || "").trim();
    if (!pedido) continue;

    const statusPedido = String(row[colStatusPedido] || "").trim().toUpperCase();
    const messageId = String(row[colDiscordMessageId] || "").trim();

    // Somente pendentes e ainda não postados
    if (statusPedido !== "PENDENTE") continue;
    if (messageId) continue;

    const cliente = String(row[colCliente] || "").trim();
    const marketplace = String(row[colMarketplace] || "").trim();

    const produto = String(row[colProduto] || "").trim();
    const qtd = String(row[colQtd] || "").trim();
    const itemKey = String(row[colItemKey] || "").trim();
    const statusItem = String(row[colStatusItem] || "").trim();

    if (!map.has(pedido)) {
      map.set(pedido, { pedido, cliente, marketplace, items: [] });
    }

    map.get(pedido).items.push({ produto, qtd, itemKey, status: statusItem });
  }

  return { orders: Array.from(map.values()) };
}

function listConfirmed_() {
  const sh = getSheetByName_(CONFIG.SHEET_NAME);
  const { idx } = getHeaderMap_(sh);

  assertCols_(idx, ["pedido","cliente","marketplace","produto","qtd","itemkey","discordmessageid","confirmado"]);

  const colPedido = idx["pedido"];
  const colCliente = idx["cliente"];
  const colMarketplace = idx["marketplace"];
  const colProduto = idx["produto"];
  const colQtd = idx["qtd"];
  const colItemKey = idx["itemkey"];
  const colDiscordMessageId = idx["discordmessageid"];
  const colConfirmado = idx["confirmado"];

  const lastRow = sh.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return { orders: [] };

  const numRows = lastRow - CONFIG.DATA_START_ROW + 1;
  const data = sh.getRange(CONFIG.DATA_START_ROW, 1, numRows, sh.getLastColumn()).getValues();

  const map = new Map(); // messageId -> order
  for (let r = 0; r < data.length; r++) {
    const row = data[r];
    const confirmado = String(row[colConfirmado] || "").trim().toUpperCase();
    if (confirmado !== "SIM") continue;

    const messageId = String(row[colDiscordMessageId] || "").trim();
    if (!messageId) continue;

    const pedido = String(row[colPedido] || "").trim();
    const cliente = String(row[colCliente] || "").trim();
    const marketplace = String(row[colMarketplace] || "").trim();
    const produto = String(row[colProduto] || "").trim();
    const qtd = String(row[colQtd] || "").trim();
    const itemKey = String(row[colItemKey] || "").trim();

    if (!map.has(messageId)) {
      map.set(messageId, {
        messageId,
        pedido,
        cliente,
        marketplace,
        items: []
      });
    }

    map.get(messageId).items.push({ produto, qtd, itemKey });
  }

  return { orders: Array.from(map.values()) };
}

/**
 * ✅ Lista confirmados agregando também os IDs de reenvio da coluna OutrosMessageId (M).
 */
function listConfirmedWithOthers_() {
  const sh = getSheetByName_(CONFIG.SHEET_NAME);
  const { idx } = getHeaderMap_(sh);

  assertCols_(idx, ["pedido","cliente","marketplace","discordmessageid","confirmado","outrosmessageid"]);

  const colPedido = idx["pedido"];
  const colCliente = idx["cliente"];
  const colMarketplace = idx["marketplace"];
  const colDiscordMessageId = idx["discordmessageid"];
  const colConfirmado = idx["confirmado"];
  const colOutros = idx["outrosmessageid"];

  const lastRow = sh.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return { orders: [] };

  const numRows = lastRow - CONFIG.DATA_START_ROW + 1;
  const data = sh.getRange(CONFIG.DATA_START_ROW, 1, numRows, sh.getLastColumn()).getValues();

  const map = new Map(); // discordMessageId -> agregado

  for (let r = 0; r < data.length; r++) {
    const row = data[r];
    const confirmado = String(row[colConfirmado] || "").trim().toUpperCase();
    if (confirmado !== "SIM") continue;

    const discordMessageId = String(row[colDiscordMessageId] || "").trim();
    if (!discordMessageId) continue;

    const pedido = String(row[colPedido] || "").trim();
    const cliente = String(row[colCliente] || "").trim();
    const marketplace = String(row[colMarketplace] || "").trim();

    const outrosRaw = String(row[colOutros] || "").trim();
    const outrosIds = outrosRaw
      ? outrosRaw.split("\n").map(s => s.trim()).filter(Boolean)
      : [];

    if (!map.has(discordMessageId)) {
      map.set(discordMessageId, {
        discordMessageId,
        pedido,
        cliente,
        marketplace,
        outrosMessageIds: new Set(),
      });
    }

    const obj = map.get(discordMessageId);
    for (const id of outrosIds) obj.outrosMessageIds.add(id);
  }

  const orders = Array.from(map.values()).map(o => ({
    discordMessageId: o.discordMessageId,
    pedido: o.pedido,
    cliente: o.cliente,
    marketplace: o.marketplace,
    outrosMessageIds: Array.from(o.outrosMessageIds),
  }));

  return { orders };
}

/**
 * ✅ Lista todos os IDs únicos presentes na coluna OutrosMessageId (M).
 */
function listOutrosMessageIds_() {
  const sh = getSheetByName_(CONFIG.SHEET_NAME);
  const { idx } = getHeaderMap_(sh);

  assertCols_(idx, ["outrosmessageid"]);

  const colOutros = idx["outrosmessageid"];

  const lastRow = sh.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return { ids: [] };

  const numRows = lastRow - CONFIG.DATA_START_ROW + 1;
  const data = sh.getRange(CONFIG.DATA_START_ROW, 1, numRows, sh.getLastColumn()).getValues();

  const set = new Set();

  for (let r = 0; r < data.length; r++) {
    const raw = String(data[r][colOutros] || "").trim();
    if (!raw) continue;

    raw.split("\n")
      .map(s => s.trim())
      .filter(Boolean)
      .forEach(id => set.add(id));
  }

  return { ids: Array.from(set) };
}

/**
 * ✅ Lista pedidos PENDENTE onde exista ao menos 1 linha do pedido com Status (coluna "status") vazio.
 */
function listPendentesStatusVazio_(limit) {
  const sh = getSheetByName_(CONFIG.SHEET_NAME);
  const { idx } = getHeaderMap_(sh);

  assertCols_(idx, ["pedido","cliente","marketplace","produto","qtd","itemkey","statuspedido","confirmado","status"]);

  const colPedido = idx["pedido"];
  const colCliente = idx["cliente"];
  const colMarketplace = idx["marketplace"];
  const colProduto = idx["produto"];
  const colQtd = idx["qtd"];
  const colItemKey = idx["itemkey"];
  const colStatusPedido = idx["statuspedido"];
  const colConfirmado = idx["confirmado"];
  const colStatusItem = idx["status"];

  const lastRow = sh.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return { orders: [] };

  const numRows = lastRow - CONFIG.DATA_START_ROW + 1;
  const data = sh.getRange(CONFIG.DATA_START_ROW, 1, numRows, sh.getLastColumn()).getValues();

  const map = new Map();
  const orderInSheet = [];

  for (let r = 0; r < data.length; r++) {
    const row = data[r];

    const pedido = String(row[colPedido] || "").trim();
    if (!pedido) continue;

    const statusPedido = String(row[colStatusPedido] || "").trim().toUpperCase();
    if (statusPedido !== "PENDENTE") continue;

    const confirmado = String(row[colConfirmado] || "").trim().toUpperCase();
    if (confirmado === "SIM") continue;

    const cliente = String(row[colCliente] || "").trim();
    const marketplace = String(row[colMarketplace] || "").trim();

    const produto = String(row[colProduto] || "").trim();
    const qtd = String(row[colQtd] || "").trim();
    const itemKey = String(row[colItemKey] || "").trim();
    const statusItem = String(row[colStatusItem] || "").trim();

    if (!map.has(pedido)) {
      map.set(pedido, {
        pedido,
        cliente,
        marketplace,
        items: [],
        hasBlank: false,
      });
      orderInSheet.push(pedido);
    }

    const obj = map.get(pedido);

    if (produto) {
      obj.items.push({ produto, qtd, itemKey, status: statusItem });
    }

    if (!statusItem) obj.hasBlank = true;
  }

  const out = [];
  for (const pedido of orderInSheet) {
    const o = map.get(pedido);
    if (!o) continue;
    if (!o.hasBlank) continue;
    if (!o.items || !o.items.length) continue;

    out.push({ pedido: o.pedido, cliente: o.cliente, marketplace: o.marketplace, items: o.items });

    if (out.length >= limit) break;
  }

  return { orders: out };
}

/**
 * ✅ Busca pedidos na aba Conferencia cuja coluna "pedido" contenha o termo (partial match).
 */
function searchPedidosConferenciaContains_(q) {
  const termRaw = String(q || "").trim();
  if (!termRaw) return { pedidos: [] };

  const sh = getSheetByName_(CONFIG.SHEET_NAME);
  const { idx } = getHeaderMap_(sh);

  assertCols_(idx, ["pedido"]);
  const colPedido = idx["pedido"];

  const lastRow = sh.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return { pedidos: [] };

  const numRows = lastRow - CONFIG.DATA_START_ROW + 1;
  const data = sh.getRange(CONFIG.DATA_START_ROW, 1, numRows, sh.getLastColumn()).getValues();

  const needle = normText_(termRaw);
  const out = [];
  const seen = new Set();

  for (let r = 0; r < data.length; r++) {
    const pedido = String(data[r][colPedido] || "").trim();
    if (!pedido) continue;

    if (!normText_(pedido).includes(needle)) continue;

    if (seen.has(pedido)) continue;
    seen.add(pedido);

    out.push(pedido);
  }

  return { pedidos: out };
}

/**
 * ✅ Obtém o "order" completo pelo PEDIDO (coluna A).
 */
function getOrderByPedido_(pedido) {
  const p0 = String(pedido || "").trim();
  if (!p0) return { order: null };

  const sh = getSheetByName_(CONFIG.SHEET_NAME);
  const { idx } = getHeaderMap_(sh);

  assertCols_(idx, ["pedido","cliente","marketplace","produto","qtd","itemkey","status"]);

  const colPedido = idx["pedido"];
  const colCliente = idx["cliente"];
  const colMarketplace = idx["marketplace"];
  const colProduto = idx["produto"];
  const colQtd = idx["qtd"];
  const colItemKey = idx["itemkey"];
  const colStatusItem = idx["status"];

  const lastRow = sh.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return { order: null };

  const numRows = lastRow - CONFIG.DATA_START_ROW + 1;
  const data = sh.getRange(CONFIG.DATA_START_ROW, 1, numRows, sh.getLastColumn()).getValues();

  let order = null;

  for (let r = 0; r < data.length; r++) {
    const row = data[r];
    const thisPedido = String(row[colPedido] || "").trim();
    if (thisPedido !== p0) continue;

    const cliente = String(row[colCliente] || "").trim();
    const marketplace = String(row[colMarketplace] || "").trim();

    if (!order) {
      order = { pedido: thisPedido, cliente, marketplace, items: [] };
    }

    const produto = String(row[colProduto] || "").trim();
    const qtd = String(row[colQtd] || "").trim();
    const itemKey = String(row[colItemKey] || "").trim();
    const status = String(row[colStatusItem] || "").trim();

    if (produto) {
      order.items.push({ produto, qtd, itemKey, status });
    }
  }

  return { order };
}

function setMessageId_(pedido, messageId) {
  const sh = getSheetByName_(CONFIG.SHEET_NAME);
  const { idx } = getHeaderMap_(sh);

  assertCols_(idx, ["pedido", "discordmessageid"]);

  const colPedido = idx["pedido"];
  const colDiscordMessageId = idx["discordmessageid"];

  const lastRow = sh.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return;

  const numRows = lastRow - CONFIG.DATA_START_ROW + 1;
  const range = sh.getRange(CONFIG.DATA_START_ROW, 1, numRows, sh.getLastColumn());
  const values = range.getValues();

  let changed = 0;
  for (let i = 0; i < values.length; i++) {
    const p = String(values[i][colPedido] || "").trim();
    if (p === pedido) {
      values[i][colDiscordMessageId] = messageId;
      changed++;
    }
  }
  if (changed) range.setValues(values);
}

function setItemStatus_(itemKey, status, conferidoPor, conferidoEmISO) {
  const sh = getSheetByName_(CONFIG.SHEET_NAME);
  const { idx } = getHeaderMap_(sh);

  assertCols_(idx, ["itemkey", "status", "conferidopor", "conferidoem"]);

  const colItemKey = idx["itemkey"];
  const colStatus = idx["status"];
  const colConferidoPor = idx["conferidopor"];
  const colConferidoEm = idx["conferidoem"];

  const lastRow = sh.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return;

  const numRows = lastRow - CONFIG.DATA_START_ROW + 1;
  const range = sh.getRange(CONFIG.DATA_START_ROW, 1, numRows, sh.getLastColumn());
  const values = range.getValues();

  let changed = 0;
  for (let i = 0; i < values.length; i++) {
    const k = String(values[i][colItemKey] || "").trim();
    if (k === itemKey) {
      values[i][colStatus] = status;
      if (conferidoPor) values[i][colConferidoPor] = conferidoPor;
      if (conferidoEmISO) values[i][colConferidoEm] = conferidoEmISO;
      changed++;
      break;
    }
  }
  if (changed) range.setValues(values);
}

function deleteOrderByMessageId_(messageId) {
  const sh = getSheetByName_(CONFIG.SHEET_NAME);
  const { idx } = getHeaderMap_(sh);
  assertCols_(idx, ["discordmessageid"]);

  const colDiscordMessageId = idx["discordmessageid"];

  const lastRow = sh.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return 0;

  const numRows = lastRow - CONFIG.DATA_START_ROW + 1;
  const values = sh.getRange(CONFIG.DATA_START_ROW, 1, numRows, sh.getLastColumn()).getValues();

  const rowsToDelete = [];
  for (let i = 0; i < values.length; i++) {
    const mid = String(values[i][colDiscordMessageId] || "").trim();
    if (mid === messageId) {
      rowsToDelete.push(CONFIG.DATA_START_ROW + i);
    }
  }

  for (let i = rowsToDelete.length - 1; i >= 0; i--) {
    sh.deleteRow(rowsToDelete[i]);
  }

  return rowsToDelete.length;
}

/************** ACTIONS FOR /pesquisar **************/
function searchListaCompleta_(q) {
  const keyword = String(q || "").trim();
  if (!keyword) return { options: [] };

  const sh = getSheetByName_(CONFIG.LIST_SHEET_NAME);

  const lastRow = sh.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return { options: [] };

  const numRows = lastRow - CONFIG.DATA_START_ROW + 1;
  const vals = sh.getRange(CONFIG.DATA_START_ROW, 4, numRows, 3).getValues(); // D..F

  const kw = normText_(keyword);
  const out = [];

  for (let i = 0; i < vals.length; i++) {
    const rowNum = CONFIG.DATA_START_ROW + i;
    const nome = String(vals[i][0] || "").trim();
    const quant = String(vals[i][1] || "").trim();
    const idsText = String(vals[i][2] || "").trim();

    if (!nome) continue;
    if (!normText_(nome).includes(kw)) continue;

    out.push({
      idKey: String(rowNum),
      nome,
      quant: quant ? Number(quant) || quant : 0,
    });
  }

  return { options: out };
}

function getListaItemByIdKey_(idKey) {
  const rowNum = parseInt(String(idKey || "").trim(), 10);
  if (!rowNum || rowNum < CONFIG.DATA_START_ROW) {
    return { item: null };
  }

  const sh = getSheetByName_(CONFIG.LIST_SHEET_NAME);
  const lastRow = sh.getLastRow();
  if (rowNum > lastRow) return { item: null };

  const v = sh.getRange(rowNum, 4, 1, 3).getValues()[0];
  const nome = String(v[0] || "").trim();
  const quant = String(v[1] || "").trim();
  const idsText = String(v[2] || "").trim();

  if (!nome) return { item: null };

  return {
    item: {
      nome,
      quant: quant ? Number(quant) || quant : 0,
      idsText,
    }
  };
}

function getOrderByMessageId_(messageId) {
  const mid = String(messageId || "").trim();
  if (!mid) return { order: null };

  const sh = getSheetByName_(CONFIG.SHEET_NAME);
  const { idx } = getHeaderMap_(sh);

  assertCols_(idx, ["pedido","cliente","marketplace","produto","qtd","itemkey","discordmessageid","status"]);

  const colPedido = idx["pedido"];
  const colCliente = idx["cliente"];
  const colMarketplace = idx["marketplace"];
  const colProduto = idx["produto"];
  const colQtd = idx["qtd"];
  const colItemKey = idx["itemkey"];
  const colDiscordMessageId = idx["discordmessageid"];
  const colStatusItem = idx["status"];

  const lastRow = sh.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return { order: null };

  const numRows = lastRow - CONFIG.DATA_START_ROW + 1;
  const data = sh.getRange(CONFIG.DATA_START_ROW, 1, numRows, sh.getLastColumn()).getValues();

  let order = null;

  for (let r = 0; r < data.length; r++) {
    const row = data[r];
    const thisMid = String(row[colDiscordMessageId] || "").trim();
    if (thisMid !== mid) continue;

    const pedido = String(row[colPedido] || "").trim();
    const cliente = String(row[colCliente] || "").trim();
    const marketplace = String(row[colMarketplace] || "").trim();

    if (!order) {
      order = { pedido, cliente, marketplace, items: [] };
    }

    const produto = String(row[colProduto] || "").trim();
    const qtd = String(row[colQtd] || "").trim();
    const itemKey = String(row[colItemKey] || "").trim();
    const status = String(row[colStatusItem] || "").trim();

    if (produto) {
      order.items.push({ produto, qtd, itemKey, status });
    }
  }

  return { order };
}

function appendOutrosMessageId_(originalMessageId, newMessageId) {
  const sh = getSheetByName_(CONFIG.SHEET_NAME);
  const { idx } = getHeaderMap_(sh);

  assertCols_(idx, ["discordmessageid", "outrosmessageid"]);

  const colDiscordMessageId = idx["discordmessageid"];
  const colOutros = idx["outrosmessageid"];

  const lastRow = sh.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return 0;

  const numRows = lastRow - CONFIG.DATA_START_ROW + 1;
  const range = sh.getRange(CONFIG.DATA_START_ROW, 1, numRows, sh.getLastColumn());
  const values = range.getValues();

  let changed = 0;

  for (let i = 0; i < values.length; i++) {
    const mid = String(values[i][colDiscordMessageId] || "").trim();
    if (mid !== originalMessageId) continue;

    const cur = String(values[i][colOutros] || "").trim();
    values[i][colOutros] = cur ? (cur + "\n" + newMessageId) : newMessageId;
    changed++;
  }

  if (changed) range.setValues(values);
  return changed;
}

function appendOutrosMessageIdByPedido_(pedido, newMessageId) {
  const sh = getSheetByName_(CONFIG.SHEET_NAME);
  const { idx } = getHeaderMap_(sh);

  assertCols_(idx, ["pedido", "outrosmessageid"]);

  const colPedido = idx["pedido"];
  const colOutros = idx["outrosmessageid"];

  const lastRow = sh.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return 0;

  const numRows = lastRow - CONFIG.DATA_START_ROW + 1;
  const range = sh.getRange(CONFIG.DATA_START_ROW, 1, numRows, sh.getLastColumn());
  const values = range.getValues();

  for (let i = 0; i < values.length; i++) {
    const p = String(values[i][colPedido] || "").trim();
    if (p !== pedido) continue;

    const cur = String(values[i][colOutros] || "").trim();
    values[i][colOutros] = cur ? (cur + "\n" + newMessageId) : newMessageId;

    range.setValues(values);
    return 1;
  }

  return 0;
}

/**
 * ✅ NOVO (pedido do usuário)
 * Append do newMessageId diretamente na célula "OutrosMessageId" da LINHA 2 (M2),
 * respeitando a coluna pelo cabeçalho.
 * - Coloca um embaixo do outro
 * - Evita duplicar o mesmo ID
 * Retorna 1 se alterou
 */
function appendOutrosMessageIdM2_(newMessageId) {
  const sh = getSheetByName_(CONFIG.SHEET_NAME);
  const { idx } = getHeaderMap_(sh);

  assertCols_(idx, ["outrosmessageid"]);
  const colOutros0 = idx["outrosmessageid"]; // 0-based

  const row = CONFIG.DATA_START_ROW; // 2
  const col = colOutros0 + 1;        // Range é 1-based

  const cell = sh.getRange(row, col);
  const cur = String(cell.getValue() || "").trim();

  const arr = cur ? cur.split("\n").map(s => s.trim()).filter(Boolean) : [];
  if (arr.includes(newMessageId)) return 1;

  const next = cur ? (cur + "\n" + newMessageId) : newMessageId;
  cell.setValue(next);
  return 1;
}

function clearOutrosMessageId_(messageId) {
  const sh = getSheetByName_(CONFIG.SHEET_NAME);
  const { idx } = getHeaderMap_(sh);

  assertCols_(idx, ["outrosmessageid"]);
  const colOutros = idx["outrosmessageid"];

  const lastRow = sh.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return 0;

  const numRows = lastRow - CONFIG.DATA_START_ROW + 1;
  const range = sh.getRange(CONFIG.DATA_START_ROW, 1, numRows, sh.getLastColumn());
  const values = range.getValues();

  let changed = 0;

  for (let i = 0; i < values.length; i++) {
    const raw = String(values[i][colOutros] || "").trim();
    if (!raw) continue;

    const arr = raw.split("\n").map(s => s.trim()).filter(Boolean);
    const filtered = arr.filter(id => id !== messageId);

    if (filtered.length !== arr.length) {
      values[i][colOutros] = filtered.join("\n");
      changed++;
    }
  }

  if (changed) range.setValues(values);
  return changed;
}

/************** HELPERS **************/
function getSheetByName_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(String(name || ""));
  if (!sh) throw new Error(`Sheet not found: ${name}`);
  return sh;
}

function getHeaderMap_(sh) {
  const lastCol = sh.getLastColumn();
  const header = sh.getRange(CONFIG.HEADER_ROW, 1, 1, lastCol).getValues()[0];

  const idx = {};
  for (let c = 0; c < header.length; c++) {
    const key = String(header[c] || "").trim().toLowerCase();
    if (!key) continue;
    idx[key.replace(/\s+/g, "")] = c;
  }
  return { headers: header, idx };
}

function assertCols_(idx, names) {
  const missing = names.filter(n => idx[n] === undefined);
  if (missing.length) throw new Error("Missing columns in header: " + missing.join(", "));
}

function isValidKey_(key) {
  return key && key === CONFIG.API_KEY;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Normaliza texto: remove acentos + uppercase + colapsa espaços
function normText_(s) {
  return String(s || "")
    .trim()
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ");
}

/**
 * Executa a importação com Lock para evitar rodar duas vezes ao mesmo tempo
 * (ex.: alguém clicar no comando duas vezes no Discord).
 *
 * Requer que a função importarDaPlanilhaB_Julho_Setembro_Agosto() exista no projeto
 * (ela está no seu arquivo Sincronização.gs).
 */
function runImportacaoComLock_() {
  const lock = LockService.getScriptLock();
  const got = lock.tryLock(1000); // tenta por 1s

  if (!got) {
    return {
      summary: "Importação já está em execução (Lock ativo). Tente novamente em alguns segundos.",
    };
  }

  const t0 = new Date();
  try {
    // Chama a função principal do seu Sincronização.gs
    importarDaPlanilhaB_Julho_Setembro_Agosto();

    const secs = Math.round((new Date() - t0) / 1000);
    return {
      summary: `Importação concluída. Tempo: ${secs}s.`,
    };
  } catch (err) {
    const msg = String(err && err.message ? err.message : err);
    return {
      summary: `❌ Importação falhou: ${msg}`,
    };
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}
