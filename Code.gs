/**************************************************************************
 * API do Discord  ->  Google Sheets
 *
 * Este arquivo expõe o Web App (doGet/doPost) consumido pelo bot.
 * A importação/sincronização com a Planilha B fica em "Sincronização.gs".
 *
 * ⚠️ CONFIGURAÇÃO OBRIGATÓRIA (uma única vez):
 *    Configurações do projeto → Propriedades do script → adicionar:
 *       SHEETS_API_KEY = <a mesma chave usada no Render>
 *    A chave NÃO fica mais no código (ela ia parar no GitHub).
 *    Rode verificarConfiguracao() no editor para conferir.
 **************************************************************************/

/************** CONFIG **************/
const SCRIPT_VERSION = "2.1.0";

const CONFIG = {
  /**
   * Script Properties que podem guardar a chave da API, em ordem de
   * preferência: a primeira que existir e tiver valor é usada.
   *
   * O valor precisa ser EXATAMENTE o que o bot envia no parâmetro "key"
   * (no Render normalmente a variável SHEETS_API_KEY).
   */
  API_KEY_PROPS: ["SHEETS_API_KEY", "BOT_SECRET", "API_KEY"],

  SHEET_NAME: "Conferencia",
  LIST_SHEET_NAME: "Lista",

  HEADER_ROW: 1,
  DATA_START_ROW: 2,
};

/**
 * Colunas que guardam IDs de mensagem do Discord (snowflakes de 18-19 dígitos).
 * Precisam ficar formatadas como TEXTO, senão o Sheets converte para número
 * e perde os últimos dígitos (acima de 2^53 a precisão do double acaba).
 */
const ID_COLUMNS = ["discordmessageid", "outrosmessageid"];

/************** WEBAPP ENTRY **************/
function doGet(e) {
  return handleRequest_("GET", e);
}

function doPost(e) {
  return handleRequest_("POST", e);
}

function handleRequest_(method, e) {
  let action = "";
  let rawAction = "";

  try {
    const params = extractParams_(method, e);

    rawAction = String(params.action || "").trim();
    action = normAction_(rawAction);

    const keyCheck = checkApiKey_(String(params.key || "").trim());

    // ping responde mesmo sem chave válida (versão reduzida), para permitir
    // descobrir qual versão está publicada sem precisar de credencial.
    if (action === "ping") {
      return json_({ ok: true, ...ping_(method, keyCheck.ok) });
    }

    if (!keyCheck.ok) return json_(keyCheck);

    const table = method === "GET" ? GET_ACTIONS_ : POST_ACTIONS_;
    const handler = table[action];

    if (handler) {
      const out = handler(params) || {};
      return json_({ ok: true, ...out });
    }

    return json_(unknownAction_(method, rawAction, action));
  } catch (err) {
    return json_({
      ok: false,
      error: String(err && err.message ? err.message : err),
      method,
      receivedAction: rawAction,
      version: SCRIPT_VERSION,
    });
  }
}

/**
 * Resposta de "Unknown action" com diagnóstico: diz o que chegou, por qual
 * método, e se a ação existe no outro verbo HTTP.
 */
function unknownAction_(method, rawAction, action) {
  const other = method === "GET" ? POST_ACTIONS_ : GET_ACTIONS_;
  const otherName = method === "GET" ? "POST" : "GET";
  const table = method === "GET" ? GET_ACTIONS_ : POST_ACTIONS_;

  const payload = {
    ok: false,
    error: "Unknown action",
    receivedAction: rawAction,
    normalizedAction: action,
    method,
    version: SCRIPT_VERSION,
  };

  if (other[action]) {
    payload.hint = `A ação "${action}" existe, mas somente via ${otherName}. O bot está chamando por ${method}.`;
  } else {
    payload.hint = `Ação não existe nesta versão do script. Disponíveis em ${method}: ${Object.keys(table).sort().join(", ")}`;
  }

  return payload;
}

/************** TABELA DE AÇÕES **************/
/** Somente leitura. */
const GET_ACTIONS_ = {
  list_pending: () => listPending_(),
  list_confirmed: () => listConfirmed_(),
  list_confirmed_with_others: () => listConfirmedWithOthers_(),
  list_outros_message_ids: () => listOutrosMessageIds_(),
  search_lista_completa: (p) => searchListaCompleta_(str_(p.q)),
  get_lista_item_by_idkey: (p) => getListaItemByIdKey_(str_(p.idKey)),
  get_order_by_message_id: (p) => getOrderByMessageId_(str_(p.messageId)),
  list_pendentes_status_vazio: (p) => listPendentesStatusVazio_(parseLimit_(p.limit)),
  search_pedidos_conferencia_contains: (p) => searchPedidosConferenciaContains_(str_(p.q)),
  get_order_by_pedido: (p) => getOrderByPedido_(str_(p.pedido)),

  // Aceita também por GET para não quebrar se o bot chamar sem POST.
  importacao: () => runImportacaoComLock_(),
};

/** Escrita. */
const POST_ACTIONS_ = {
  set_message_id: (p) => {
    const pedido = req_(p.pedido, "pedido/messageId required");
    const messageId = req_(p.messageId, "pedido/messageId required");
    return { updatedRows: setMessageId_(pedido, messageId) };
  },

  set_item_status: (p) => {
    const itemKey = req_(p.itemKey, "itemKey/status required");
    const status = req_(p.status, "itemKey/status required");
    const updated = setItemStatus_(itemKey, status, str_(p.conferidoPor), str_(p.conferidoEmISO));
    return { updated };
  },

  delete_order_by_message_id: (p) => {
    const messageId = req_(p.messageId, "messageId required");
    return { deletedRows: deleteOrderByMessageId_(messageId) };
  },

  append_outros_message_id: (p) => {
    const originalMessageId = req_(p.originalMessageId, "originalMessageId/newMessageId required");
    const newMessageId = req_(p.newMessageId, "originalMessageId/newMessageId required");
    return { updatedRows: appendOutrosMessageId_(originalMessageId, newMessageId) };
  },

  clear_outros_message_id: (p) => {
    const messageId = req_(p.messageId, "messageId required");
    return { cleared: clearOutrosMessageId_(messageId) };
  },

  append_outros_message_id_by_pedido: (p) => {
    const pedido = req_(p.pedido, "pedido/newMessageId required");
    const newMessageId = req_(p.newMessageId, "pedido/newMessageId required");
    return { updated: appendOutrosMessageIdByPedido_(pedido, newMessageId) };
  },

  append_outros_message_id_m2: (p) => {
    const newMessageId = req_(p.newMessageId, "newMessageId required");
    return { updated: appendOutrosMessageIdM2_(newMessageId, p.row) };
  },

  importacao: () => runImportacaoComLock_(),
};

/************** DIAGNÓSTICO **************/
function ping_(method, authenticated) {
  const info = getApiKeyInfo_();

  const base = {
    version: SCRIPT_VERSION,
    method,
    authenticated: !!authenticated,
    apiKeyConfigured: !!info.value,
    apiKeyProp: info.prop, // qual Script Property está sendo usada
    serverTime: new Date().toISOString(),
  };

  if (!authenticated) return base;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return {
    ...base,
    spreadsheet: ss.getName(),
    sheets: {
      conferencia: !!ss.getSheetByName(CONFIG.SHEET_NAME),
      lista: !!ss.getSheetByName(CONFIG.LIST_SHEET_NAME),
    },
    getActions: Object.keys(GET_ACTIONS_).sort(),
    postActions: Object.keys(POST_ACTIONS_).sort(),
  };
}

/************** AÇÕES DE LEITURA **************/
function listPending_() {
  const sh = getSheetByName_(CONFIG.SHEET_NAME);
  const { idx } = getHeaderMap_(sh);

  assertCols_(idx, ["pedido","cliente","marketplace","produto","qtd","itemkey","statuspedido","discordmessageid","status"]);

  const data = readAllData_(sh);
  if (!data.length) return { orders: [] };

  const map = new Map();
  for (let r = 0; r < data.length; r++) {
    const row = data[r];
    const pedido = str_(row[idx["pedido"]]);
    if (!pedido) continue;

    const statusPedido = str_(row[idx["statuspedido"]]).toUpperCase();
    const messageId = readId_(row[idx["discordmessageid"]]);

    // Somente pendentes e ainda não postados
    if (statusPedido !== "PENDENTE") continue;
    if (messageId) continue;

    if (!map.has(pedido)) {
      map.set(pedido, {
        pedido,
        cliente: str_(row[idx["cliente"]]),
        marketplace: str_(row[idx["marketplace"]]),
        items: [],
      });
    }

    map.get(pedido).items.push({
      produto: str_(row[idx["produto"]]),
      qtd: str_(row[idx["qtd"]]),
      itemKey: str_(row[idx["itemkey"]]),
      status: str_(row[idx["status"]]),
    });
  }

  return { orders: Array.from(map.values()) };
}

function listConfirmed_() {
  const sh = getSheetByName_(CONFIG.SHEET_NAME);
  const { idx } = getHeaderMap_(sh);

  assertCols_(idx, ["pedido","cliente","marketplace","produto","qtd","itemkey","discordmessageid","confirmado"]);

  const data = readAllData_(sh);
  if (!data.length) return { orders: [] };

  const map = new Map(); // messageId -> order
  for (let r = 0; r < data.length; r++) {
    const row = data[r];
    if (str_(row[idx["confirmado"]]).toUpperCase() !== "SIM") continue;

    const messageId = readId_(row[idx["discordmessageid"]]);
    if (!messageId) continue;

    if (!map.has(messageId)) {
      map.set(messageId, {
        messageId,
        pedido: str_(row[idx["pedido"]]),
        cliente: str_(row[idx["cliente"]]),
        marketplace: str_(row[idx["marketplace"]]),
        items: [],
      });
    }

    map.get(messageId).items.push({
      produto: str_(row[idx["produto"]]),
      qtd: str_(row[idx["qtd"]]),
      itemKey: str_(row[idx["itemkey"]]),
    });
  }

  return { orders: Array.from(map.values()) };
}

/** Lista confirmados agregando também os IDs de reenvio (coluna OutrosMessageId). */
function listConfirmedWithOthers_() {
  const sh = getSheetByName_(CONFIG.SHEET_NAME);
  const { idx } = getHeaderMap_(sh);

  assertCols_(idx, ["pedido","cliente","marketplace","discordmessageid","confirmado","outrosmessageid"]);

  const data = readAllData_(sh);
  if (!data.length) return { orders: [] };

  const map = new Map(); // discordMessageId -> agregado

  for (let r = 0; r < data.length; r++) {
    const row = data[r];
    if (str_(row[idx["confirmado"]]).toUpperCase() !== "SIM") continue;

    const discordMessageId = readId_(row[idx["discordmessageid"]]);
    if (!discordMessageId) continue;

    if (!map.has(discordMessageId)) {
      map.set(discordMessageId, {
        discordMessageId,
        pedido: str_(row[idx["pedido"]]),
        cliente: str_(row[idx["cliente"]]),
        marketplace: str_(row[idx["marketplace"]]),
        outrosMessageIds: new Set(),
      });
    }

    const obj = map.get(discordMessageId);
    for (const id of splitIds_(row[idx["outrosmessageid"]])) obj.outrosMessageIds.add(id);
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

/** Lista todos os IDs únicos presentes na coluna OutrosMessageId. */
function listOutrosMessageIds_() {
  const sh = getSheetByName_(CONFIG.SHEET_NAME);
  const { idx } = getHeaderMap_(sh);

  assertCols_(idx, ["outrosmessageid"]);

  const data = readAllData_(sh);
  if (!data.length) return { ids: [] };

  const set = new Set();
  for (let r = 0; r < data.length; r++) {
    for (const id of splitIds_(data[r][idx["outrosmessageid"]])) set.add(id);
  }

  return { ids: Array.from(set) };
}

/** Pedidos PENDENTE com ao menos 1 item sem status. */
function listPendentesStatusVazio_(limit) {
  const sh = getSheetByName_(CONFIG.SHEET_NAME);
  const { idx } = getHeaderMap_(sh);

  assertCols_(idx, ["pedido","cliente","marketplace","produto","qtd","itemkey","statuspedido","confirmado","status"]);

  const data = readAllData_(sh);
  if (!data.length) return { orders: [] };

  const map = new Map();
  const orderInSheet = [];

  for (let r = 0; r < data.length; r++) {
    const row = data[r];

    const pedido = str_(row[idx["pedido"]]);
    if (!pedido) continue;

    if (str_(row[idx["statuspedido"]]).toUpperCase() !== "PENDENTE") continue;
    if (str_(row[idx["confirmado"]]).toUpperCase() === "SIM") continue;

    if (!map.has(pedido)) {
      map.set(pedido, {
        pedido,
        cliente: str_(row[idx["cliente"]]),
        marketplace: str_(row[idx["marketplace"]]),
        items: [],
        hasBlank: false,
      });
      orderInSheet.push(pedido);
    }

    const obj = map.get(pedido);
    const produto = str_(row[idx["produto"]]);
    const statusItem = str_(row[idx["status"]]);

    if (produto) {
      obj.items.push({
        produto,
        qtd: str_(row[idx["qtd"]]),
        itemKey: str_(row[idx["itemkey"]]),
        status: statusItem,
      });
    }

    if (!statusItem) obj.hasBlank = true;
  }

  const out = [];
  for (const pedido of orderInSheet) {
    const o = map.get(pedido);
    if (!o || !o.hasBlank || !o.items.length) continue;

    out.push({ pedido: o.pedido, cliente: o.cliente, marketplace: o.marketplace, items: o.items });
    if (out.length >= limit) break;
  }

  return { orders: out };
}

/** Busca parcial no número do pedido (usado pelo /codigo). */
function searchPedidosConferenciaContains_(q) {
  const termRaw = str_(q);
  if (!termRaw) return { pedidos: [] };

  const sh = getSheetByName_(CONFIG.SHEET_NAME);
  const { idx } = getHeaderMap_(sh);

  assertCols_(idx, ["pedido"]);

  const data = readAllData_(sh);
  if (!data.length) return { pedidos: [] };

  const needle = normText_(termRaw);
  const out = [];
  const seen = new Set();

  for (let r = 0; r < data.length; r++) {
    const pedido = str_(data[r][idx["pedido"]]);
    if (!pedido) continue;
    if (!normText_(pedido).includes(needle)) continue;
    if (seen.has(pedido)) continue;

    seen.add(pedido);
    out.push(pedido);
  }

  return { pedidos: out };
}

/** Pedido completo a partir do número do pedido. */
function getOrderByPedido_(pedido) {
  const p0 = str_(pedido);
  if (!p0) return { order: null };

  const sh = getSheetByName_(CONFIG.SHEET_NAME);
  const { idx } = getHeaderMap_(sh);

  assertCols_(idx, ["pedido","cliente","marketplace","produto","qtd","itemkey","status"]);

  const data = readAllData_(sh);
  if (!data.length) return { order: null };

  let order = null;

  for (let r = 0; r < data.length; r++) {
    const row = data[r];
    if (str_(row[idx["pedido"]]) !== p0) continue;

    if (!order) {
      order = {
        pedido: p0,
        cliente: str_(row[idx["cliente"]]),
        marketplace: str_(row[idx["marketplace"]]),
        items: [],
      };
    }

    const produto = str_(row[idx["produto"]]);
    if (produto) {
      order.items.push({
        produto,
        qtd: str_(row[idx["qtd"]]),
        itemKey: str_(row[idx["itemkey"]]),
        status: str_(row[idx["status"]]),
      });
    }
  }

  return { order };
}

/** Pedido completo a partir do DiscordMessageId. */
function getOrderByMessageId_(messageId) {
  const mid = str_(messageId);
  if (!mid) return { order: null };

  const sh = getSheetByName_(CONFIG.SHEET_NAME);
  const { idx } = getHeaderMap_(sh);

  assertCols_(idx, ["pedido","cliente","marketplace","produto","qtd","itemkey","discordmessageid","status"]);

  const data = readAllData_(sh);
  if (!data.length) return { order: null };

  let order = null;

  for (let r = 0; r < data.length; r++) {
    const row = data[r];
    if (readId_(row[idx["discordmessageid"]]) !== mid) continue;

    if (!order) {
      order = {
        pedido: str_(row[idx["pedido"]]),
        cliente: str_(row[idx["cliente"]]),
        marketplace: str_(row[idx["marketplace"]]),
        items: [],
      };
    }

    const produto = str_(row[idx["produto"]]);
    if (produto) {
      order.items.push({
        produto,
        qtd: str_(row[idx["qtd"]]),
        itemKey: str_(row[idx["itemkey"]]),
        status: str_(row[idx["status"]]),
      });
    }
  }

  return { order };
}

/************** AÇÕES DE LEITURA: aba Lista (/pesquisar) **************/
function searchListaCompleta_(q) {
  const keyword = str_(q);
  if (!keyword) return { options: [] };

  const sh = getSheetByName_(CONFIG.LIST_SHEET_NAME);

  const lastRow = sh.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return { options: [] };

  const numRows = lastRow - CONFIG.DATA_START_ROW + 1;
  const vals = sh.getRange(CONFIG.DATA_START_ROW, 4, numRows, 2).getValues(); // D..E

  const kw = normText_(keyword);
  const out = [];

  for (let i = 0; i < vals.length; i++) {
    const nome = str_(vals[i][0]);
    if (!nome) continue;
    if (!normText_(nome).includes(kw)) continue;

    const quant = str_(vals[i][1]);
    out.push({
      idKey: String(CONFIG.DATA_START_ROW + i), // idKey = número da linha
      nome,
      quant: quant ? Number(quant) || quant : 0,
    });
  }

  return { options: out };
}

function getListaItemByIdKey_(idKey) {
  const rowNum = parseInt(str_(idKey), 10);
  if (!rowNum || rowNum < CONFIG.DATA_START_ROW) return { item: null };

  const sh = getSheetByName_(CONFIG.LIST_SHEET_NAME);
  if (rowNum > sh.getLastRow()) return { item: null };

  const v = sh.getRange(rowNum, 4, 1, 3).getValues()[0]; // D..F
  const nome = str_(v[0]);
  if (!nome) return { item: null };

  const quant = str_(v[1]);

  return {
    item: {
      nome,
      quant: quant ? Number(quant) || quant : 0,
      idsText: splitIds_(v[2]).join("\n"),
    }
  };
}

/************** AÇÕES DE ESCRITA **************/
/**
 * Todas as escritas abaixo gravam SOMENTE as células alvo.
 * (A versão anterior lia e regravava a planilha inteira a cada chamada, o que
 *  apagava fórmulas e fazia duas conferências simultâneas se sobrescreverem.)
 */

function setMessageId_(pedido, messageId) {
  const sh = getSheetByName_(CONFIG.SHEET_NAME);
  const { idx } = getHeaderMap_(sh);

  assertCols_(idx, ["pedido", "discordmessageid"]);

  const col = idx["pedido"] + 1;
  const lastRow = sh.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return 0;

  const numRows = lastRow - CONFIG.DATA_START_ROW + 1;
  const pedidos = sh.getRange(CONFIG.DATA_START_ROW, col, numRows, 1).getValues();

  const rows = [];
  for (let i = 0; i < pedidos.length; i++) {
    if (str_(pedidos[i][0]) === pedido) rows.push(CONFIG.DATA_START_ROW + i);
  }

  return writeColumnCells_(sh, idx["discordmessageid"] + 1, rows, () => messageId, true);
}

function setItemStatus_(itemKey, status, conferidoPor, conferidoEmISO) {
  const sh = getSheetByName_(CONFIG.SHEET_NAME);
  const { idx } = getHeaderMap_(sh);

  assertCols_(idx, ["itemkey", "status", "conferidopor", "conferidoem"]);

  const lastRow = sh.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return 0;

  const numRows = lastRow - CONFIG.DATA_START_ROW + 1;
  const keys = sh.getRange(CONFIG.DATA_START_ROW, idx["itemkey"] + 1, numRows, 1).getValues();

  let targetRow = 0;
  for (let i = 0; i < keys.length; i++) {
    if (str_(keys[i][0]) === itemKey) {
      targetRow = CONFIG.DATA_START_ROW + i;
      break;
    }
  }

  if (!targetRow) return 0;

  // Grava célula a célula: não toca em nenhuma outra linha/coluna.
  sh.getRange(targetRow, idx["status"] + 1).setValue(status);
  if (conferidoPor) sh.getRange(targetRow, idx["conferidopor"] + 1).setValue(conferidoPor);
  if (conferidoEmISO) sh.getRange(targetRow, idx["conferidoem"] + 1).setValue(conferidoEmISO);

  return 1;
}

function deleteOrderByMessageId_(messageId) {
  const sh = getSheetByName_(CONFIG.SHEET_NAME);
  const { idx } = getHeaderMap_(sh);

  assertCols_(idx, ["discordmessageid"]);

  const lastRow = sh.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return 0;

  const numRows = lastRow - CONFIG.DATA_START_ROW + 1;
  const ids = sh.getRange(CONFIG.DATA_START_ROW, idx["discordmessageid"] + 1, numRows, 1).getValues();

  const rows = [];
  for (let i = 0; i < ids.length; i++) {
    if (readId_(ids[i][0]) === messageId) rows.push(CONFIG.DATA_START_ROW + i);
  }

  if (!rows.length) return 0;

  // Apaga em blocos contíguos, de baixo para cima (1 chamada por bloco em vez
  // de 1 por linha).
  const blocks = contiguousBlocks_(rows);
  for (let b = blocks.length - 1; b >= 0; b--) {
    sh.deleteRows(blocks[b].start, blocks[b].count);
  }

  return rows.length;
}

function appendOutrosMessageId_(originalMessageId, newMessageId) {
  const sh = getSheetByName_(CONFIG.SHEET_NAME);
  const { idx } = getHeaderMap_(sh);

  assertCols_(idx, ["discordmessageid", "outrosmessageid"]);

  const lastRow = sh.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return 0;

  const numRows = lastRow - CONFIG.DATA_START_ROW + 1;
  const colOutros1 = idx["outrosmessageid"] + 1;

  const ids = sh.getRange(CONFIG.DATA_START_ROW, idx["discordmessageid"] + 1, numRows, 1).getValues();
  const outros = sh.getRange(CONFIG.DATA_START_ROW, colOutros1, numRows, 1).getValues();

  const rows = [];
  const novoPorLinha = {};

  for (let i = 0; i < ids.length; i++) {
    if (readId_(ids[i][0]) !== originalMessageId) continue;

    const atual = splitIds_(outros[i][0]);
    if (atual.includes(newMessageId)) continue; // já está lá

    const row = CONFIG.DATA_START_ROW + i;
    atual.push(newMessageId);
    novoPorLinha[row] = atual.join("\n");
    rows.push(row);
  }

  return writeColumnCells_(sh, colOutros1, rows, (row) => novoPorLinha[row], true);
}

function appendOutrosMessageIdByPedido_(pedido, newMessageId) {
  const sh = getSheetByName_(CONFIG.SHEET_NAME);
  const { idx } = getHeaderMap_(sh);

  assertCols_(idx, ["pedido", "outrosmessageid"]);

  const lastRow = sh.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return 0;

  const numRows = lastRow - CONFIG.DATA_START_ROW + 1;
  const colOutros1 = idx["outrosmessageid"] + 1;

  const pedidos = sh.getRange(CONFIG.DATA_START_ROW, idx["pedido"] + 1, numRows, 1).getValues();

  for (let i = 0; i < pedidos.length; i++) {
    if (str_(pedidos[i][0]) !== pedido) continue;

    // Mantém o comportamento original: grava apenas na PRIMEIRA linha do pedido.
    const row = CONFIG.DATA_START_ROW + i;
    const cell = sh.getRange(row, colOutros1);
    const atual = splitIds_(cell.getValue());

    if (atual.includes(newMessageId)) return 1;

    atual.push(newMessageId);
    cell.setNumberFormat("@");
    cell.setValue(atual.join("\n"));
    return 1;
  }

  return 0;
}

/**
 * Append do newMessageId na coluna OutrosMessageId de UMA linha fixa.
 *
 * ⚠️ Por padrão essa linha é a DATA_START_ROW (linha 2), independente de qual
 * pedido seja — esse era o comportamento original e o bot depende dele.
 * Agora dá para escolher outra linha passando "row" no POST, sem mudar o padrão.
 */
function appendOutrosMessageIdM2_(newMessageId, rowParam) {
  const sh = getSheetByName_(CONFIG.SHEET_NAME);
  const { idx } = getHeaderMap_(sh);

  assertCols_(idx, ["outrosmessageid"]);

  const row = Math.max(CONFIG.DATA_START_ROW, parseInt(str_(rowParam), 10) || CONFIG.DATA_START_ROW);
  const cell = sh.getRange(row, idx["outrosmessageid"] + 1);

  const atual = splitIds_(cell.getValue());
  if (atual.includes(newMessageId)) return 1;

  atual.push(newMessageId);
  cell.setNumberFormat("@");
  cell.setValue(atual.join("\n"));
  return 1;
}

function clearOutrosMessageId_(messageId) {
  const sh = getSheetByName_(CONFIG.SHEET_NAME);
  const { idx } = getHeaderMap_(sh);

  assertCols_(idx, ["outrosmessageid"]);

  const lastRow = sh.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return 0;

  const numRows = lastRow - CONFIG.DATA_START_ROW + 1;
  const colOutros1 = idx["outrosmessageid"] + 1;
  const outros = sh.getRange(CONFIG.DATA_START_ROW, colOutros1, numRows, 1).getValues();

  const rows = [];
  const novoPorLinha = {};

  for (let i = 0; i < outros.length; i++) {
    const arr = splitIds_(outros[i][0]);
    if (!arr.length) continue;

    const filtered = arr.filter(id => id !== messageId);
    if (filtered.length === arr.length) continue;

    const row = CONFIG.DATA_START_ROW + i;
    novoPorLinha[row] = filtered.join("\n");
    rows.push(row);
  }

  return writeColumnCells_(sh, colOutros1, rows, (row) => novoPorLinha[row], true);
}

/**
 * Executa a importação com Lock para evitar rodar duas vezes ao mesmo tempo
 * (ex.: alguém clicar no comando duas vezes no Discord).
 */
function runImportacaoComLock_() {
  const lock = LockService.getScriptLock();
  const got = lock.tryLock(1000);

  if (!got) {
    return { summary: "Importação já está em execução (Lock ativo). Tente novamente em alguns segundos." };
  }

  const t0 = new Date();
  try {
    const stats = importarDaPlanilhaB() || {};
    const secs = Math.round((new Date() - t0) / 1000);

    let summary = `Importação concluída. Tempo: ${secs}s.`;
    if (stats.abas && stats.abas.length) summary += ` Abas: ${stats.abas.join(", ")}.`;
    if (typeof stats.importadas === "number") summary += ` Linhas novas: ${stats.importadas}.`;
    if (stats.avisos && stats.avisos.length) summary += `\n⚠️ ${stats.avisos.join("\n⚠️ ")}`;

    return { summary, ...stats };
  } catch (err) {
    return { summary: `❌ Importação falhou: ${String(err && err.message ? err.message : err)}` };
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

/************** HELPERS: PLANILHA **************/
function getSheetByName_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(String(name || ""));
  if (!sh) throw new Error(`Sheet not found: ${name}`);
  return sh;
}

/** Lê todas as linhas de dados (da DATA_START_ROW até a última). */
function readAllData_(sh) {
  const lastRow = sh.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return [];

  const numRows = lastRow - CONFIG.DATA_START_ROW + 1;
  return sh.getRange(CONFIG.DATA_START_ROW, 1, numRows, sh.getLastColumn()).getValues();
}

/**
 * Mapeia cabeçalho -> índice (0-based).
 * A normalização remove acentos e qualquer caractere que não seja letra/número,
 * então "Qtd.", "Conferido Por" e "Discord Message ID" continuam funcionando.
 */
function getHeaderMap_(sh) {
  const lastCol = sh.getLastColumn();
  const header = sh.getRange(CONFIG.HEADER_ROW, 1, 1, lastCol).getValues()[0];

  const idx = {};
  for (let c = 0; c < header.length; c++) {
    const key = normHeader_(header[c]);
    if (!key) continue;
    idx[key] = c;
  }
  return { headers: header, idx };
}

function normHeader_(s) {
  return String(s || "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]/g, "");
}

function assertCols_(idx, names) {
  const missing = names.filter(n => idx[n] === undefined);
  if (!missing.length) return;

  throw new Error(
    "Missing columns in header: " + missing.join(", ") +
    " | Cabeçalhos encontrados: " + Object.keys(idx).sort().join(", ")
  );
}

/**
 * Grava valores em células específicas de UMA coluna, agrupando linhas
 * contíguas para reduzir o número de chamadas à API do Sheets.
 * asText = formata como texto (obrigatório para IDs do Discord).
 */
function writeColumnCells_(sh, col1Based, rows, valueFn, asText) {
  if (!rows || !rows.length) return 0;

  const blocks = contiguousBlocks_(rows);
  for (const b of blocks) {
    const vals = [];
    for (let r = b.start; r < b.start + b.count; r++) vals.push([valueFn(r)]);

    const rng = sh.getRange(b.start, col1Based, b.count, 1);
    if (asText) rng.setNumberFormat("@");
    rng.setValues(vals);
  }

  return rows.length;
}

/** [2,3,4,9,10] -> [{start:2,count:3},{start:9,count:2}] */
function contiguousBlocks_(rows) {
  const sorted = rows.slice().sort((a, b) => a - b);
  const blocks = [];

  let i = 0;
  while (i < sorted.length) {
    let j = i;
    while (j + 1 < sorted.length && sorted[j + 1] === sorted[j] + 1) j++;
    blocks.push({ start: sorted[i], count: sorted[j] - sorted[i] + 1 });
    i = j + 1;
  }

  return blocks;
}

/************** HELPERS: TEXTO / IDs **************/
function str_(v) {
  return String(v === null || v === undefined ? "" : v).trim();
}

/**
 * Lê um ID de mensagem do Discord de uma célula.
 * Se a célula estiver como NÚMERO (formatação errada), evita ao menos a
 * notação científica. A precisão perdida acima de 2^53 é irrecuperável — por
 * isso as colunas de ID precisam estar formatadas como texto.
 */
function readId_(v) {
  if (v === null || v === undefined) return "";
  if (typeof v === "number") {
    return Number.isFinite(v) ? v.toLocaleString("fullwide", { useGrouping: false }) : "";
  }
  return String(v).trim();
}

/** Quebra a célula OutrosMessageId (vários IDs separados por quebra de linha). */
function splitIds_(v) {
  return readId_(v)
    .split("\n")
    .map(s => s.trim())
    .filter(Boolean);
}

/** Normaliza o nome da ação: minúsculas, sem acento, espaços/hífens viram "_". */
function normAction_(s) {
  return String(s || "")
    .trim()
    .replace(/^\/+/, "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[\s-]+/g, "_");
}

/** Normaliza texto para busca: sem acento, maiúsculo, espaços colapsados. */
function normText_(s) {
  return String(s || "")
    .trim()
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ");
}

function parseLimit_(v) {
  return Math.max(1, Math.min(50, parseInt(str_(v) || "10", 10) || 10));
}

function req_(v, msg) {
  const s = str_(v);
  if (!s) throw new Error(msg);
  return s;
}

/************** HELPERS: HTTP **************/
function extractParams_(method, e) {
  if (!e) {
    throw new Error("Sem contexto HTTP. Esta função é chamada pelo Web App — não pelo botão ▶ do editor. Use verificarConfiguracao() para testar.");
  }

  const qs = e.parameter || {};
  if (method === "GET") return qs;

  const raw = (e.postData && e.postData.contents) ? e.postData.contents : "";
  if (!raw) return qs;

  let body;
  try {
    body = JSON.parse(raw);
  } catch (err) {
    throw new Error("Body não é JSON válido: " + String(err && err.message ? err.message : err));
  }

  // Aceita parâmetros pela querystring também; o body tem prioridade.
  return Object.assign({}, qs, body || {});
}

/**
 * Procura a chave nas Script Properties listadas em CONFIG.API_KEY_PROPS.
 * Retorna { prop, value } da primeira que tiver valor, ou { prop:"", value:"" }.
 */
function getApiKeyInfo_() {
  const props = PropertiesService.getScriptProperties();

  for (const nome of CONFIG.API_KEY_PROPS || []) {
    const v = String(props.getProperty(nome) || "").trim();
    if (v) return { prop: nome, value: v };
  }

  return { prop: "", value: "" };
}

function checkApiKey_(key) {
  const info = getApiKeyInfo_();

  if (!info.value) {
    return {
      ok: false,
      error: "API key não configurada",
      hint: `Configurações do projeto → Propriedades do script → adicionar uma destas: ${(CONFIG.API_KEY_PROPS || []).join(", ")}.`,
      version: SCRIPT_VERSION,
    };
  }

  if (!key || key !== info.value) return { ok: false, error: "Invalid key" };

  return { ok: true };
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Prévia mascarada de um segredo, para conferência sem expor o valor. */
function mascarar_(v) {
  const s = String(v || "");
  if (s.length <= 4) return "*".repeat(s.length);
  return s.slice(0, 2) + "*".repeat(Math.max(3, s.length - 4)) + s.slice(-2);
}

/************** MANUTENÇÃO (rodar pelo editor) **************/
/**
 * Confere se o projeto está pronto: chave configurada, abas existentes,
 * cabeçalhos reconhecidos e colunas de ID formatadas como texto.
 */
function verificarConfiguracao() {
  const linhas = [];
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const info = getApiKeyInfo_();
  linhas.push(`Versão do script: ${SCRIPT_VERSION}`);

  if (info.value) {
    linhas.push(`✅ Chave lida da Script Property "${info.prop}" (${info.value.length} caracteres, ${mascarar_(info.value)})`);
    linhas.push(`   Confira se bate EXATAMENTE com o valor que o bot envia em "key" (no Render, SHEETS_API_KEY).`);
  } else {
    linhas.push(`❌ Nenhuma chave encontrada. Configurações do projeto → Propriedades do script → criar uma destas: ${(CONFIG.API_KEY_PROPS || []).join(", ")}`);
  }

  // Avisa se houver mais de uma propriedade de chave preenchida.
  const preenchidas = (CONFIG.API_KEY_PROPS || []).filter(
    (n) => String(PropertiesService.getScriptProperties().getProperty(n) || "").trim()
  );
  if (preenchidas.length > 1) {
    linhas.push(`⚠️ Mais de uma propriedade de chave preenchida (${preenchidas.join(", ")}). Vale a primeira da lista: "${info.prop}".`);
  }

  const shConf = ss.getSheetByName(CONFIG.SHEET_NAME);
  const shLista = ss.getSheetByName(CONFIG.LIST_SHEET_NAME);

  linhas.push(shConf ? `✅ Aba "${CONFIG.SHEET_NAME}" encontrada` : `❌ Aba "${CONFIG.SHEET_NAME}" não encontrada`);
  linhas.push(shLista ? `✅ Aba "${CONFIG.LIST_SHEET_NAME}" encontrada` : `❌ Aba "${CONFIG.LIST_SHEET_NAME}" não encontrada`);

  if (shConf) {
    const { idx } = getHeaderMap_(shConf);
    const esperadas = ["pedido","cliente","marketplace","produto","qtd","status","conferidopor","conferidoem","discordmessageid","itemkey","statuspedido","confirmado","outrosmessageid"];
    const faltando = esperadas.filter(n => idx[n] === undefined);

    linhas.push(faltando.length
      ? `❌ Cabeçalhos faltando: ${faltando.join(", ")}`
      : "✅ Todos os cabeçalhos esperados foram encontrados");
    linhas.push(`   Encontrados: ${Object.keys(idx).sort().join(", ")}`);

    for (const nome of ID_COLUMNS) {
      if (idx[nome] === undefined) continue;
      const fmt = shConf.getRange(CONFIG.DATA_START_ROW, idx[nome] + 1).getNumberFormat();
      linhas.push(fmt === "@"
        ? `✅ Coluna "${nome}" formatada como texto`
        : `❌ Coluna "${nome}" NÃO é texto (formato "${fmt}") — rode formatarColunasDeIdComoTexto()`);
    }
  }

  const msg = linhas.join("\n");
  Logger.log(msg);
  return msg;
}

/**
 * Formata as colunas de ID como texto simples.
 * Isso impede que novos IDs do Discord sejam convertidos em número (o que
 * corrompe os últimos dígitos). Rode UMA vez.
 *
 * ⚠️ IDs que já foram gravados como número perderam dígitos e não têm como ser
 * recuperados — precisam ser regravados pelo bot.
 */
function formatarColunasDeIdComoTexto() {
  const sh = getSheetByName_(CONFIG.SHEET_NAME);
  const { idx } = getHeaderMap_(sh);

  const maxRows = sh.getMaxRows();
  const feitas = [];

  for (const nome of ID_COLUMNS) {
    if (idx[nome] === undefined) continue;
    sh.getRange(CONFIG.DATA_START_ROW, idx[nome] + 1, maxRows - CONFIG.DATA_START_ROW + 1, 1)
      .setNumberFormat("@");
    feitas.push(nome);
  }

  const msg = feitas.length
    ? `✅ Colunas formatadas como texto: ${feitas.join(", ")}`
    : "⚠️ Nenhuma coluna de ID encontrada no cabeçalho.";

  Logger.log(msg);
  return msg;
}

/** Lista os IDs suspeitos de terem sido corrompidos (gravados como número). */
function listarIdsSuspeitos() {
  const sh = getSheetByName_(CONFIG.SHEET_NAME);
  const { idx } = getHeaderMap_(sh);

  const data = readAllData_(sh);
  const suspeitos = [];

  for (let i = 0; i < data.length; i++) {
    for (const nome of ID_COLUMNS) {
      if (idx[nome] === undefined) continue;
      const v = data[i][idx[nome]];
      if (typeof v === "number") {
        suspeitos.push(`linha ${CONFIG.DATA_START_ROW + i} / ${nome}: ${readId_(v)}`);
      }
    }
  }

  const msg = suspeitos.length
    ? `❌ ${suspeitos.length} ID(s) gravados como número (provavelmente corrompidos):\n` + suspeitos.join("\n")
    : "✅ Nenhum ID gravado como número.";

  Logger.log(msg);
  return msg;
}
