/******************* CONFIG *******************/
const IMPORT_CONFIG = {
  // ====== PLANILHA B (ORIGEM) ======
  SOURCE_SPREADSHEET_ID: "1RHTxClkkFamZUE4hM8romnZfWWTecNwvkvI4SO7rYmU",

  // Considera múltiplas abas
  SOURCE_SHEET_NAMES: ["Setembro", "Julho", "Agosto"],

  // Dados começam na linha 6
  DATA_START_ROW: 6,

  // Colunas fixas (por número) na Planilha B
  COL_B: 2, // Pedido
  COL_C: 3, // Marketplace
  COL_D: 4, // Nome
  COL_E: 5, // Qt
  COL_F: 6, // Produto
  COL_G: 7, // Status (ok / cor verde etc.)

  // Coluna na Planilha B onde refletimos o Status da Planilha A (H)
  SOURCE_WRITEBACK_COL_J: 8, // H

  // ====== marketplaces com regras de corte do pedido ======

  // 13 PRIMEIROS caracteres
  SHORT_PEDIDO_FIRST_MARKETS: [
    "Essência do Brasil - Site",
    "Ponte Vecchio (Site)",
  ],

  // 13 ÚLTIMOS caracteres
  SHORT_PEDIDO_LAST_MARKETS: [
    "TTeste",
  ],

  SHORT_PEDIDO_MAXLEN: 13,

  // ====== PLANILHA A (DESTINO) ======
  DEST_SHEET_NAME: "Conferencia",

  // ====== TERMOS PROIBIDOS (na linha inteira A..G da Planilha B) ======
  FORBIDDEN_TERMS: ["fulfillment", "full", "cancelado"],

  // ====== COLUNAS DESTINO ======
  DEST_COL_PEDIDO: 1, // A
  DEST_COL_MARKET: 2, // B
  DEST_COL_NOME: 3, // C
  DEST_COL_PRODUTO: 4, // D
  DEST_COL_QT: 5, // E
  DEST_COL_STATUS_F: 6, // F
  DEST_COL_ITEMKEY_J: 10, // J => "PEDIDO#01"
  DEST_COL_STATUS_K: 11, // K => PENDENTE
  DEST_COL_FLAG_L: 12, // L => SIM (quando entrar no critério proibido na B)
};

/******************* FUNÇÃO PRINCIPAL *******************/
function importarDaPlanilhaB_Julho_Setembro_Agosto() {
  const ssDest = SpreadsheetApp.getActiveSpreadsheet();
  const shDest = getOrCreateSheet_(ssDest, IMPORT_CONFIG.DEST_SHEET_NAME);

  // 1) Carrega chaves já existentes na Conferencia
  const importedSet = loadExistingKeysFromConferencia_(shDest);

  // 2) Abre origem
  const ssSrc = SpreadsheetApp.openById(IMPORT_CONFIG.SOURCE_SPREADSHEET_ID);

  let totalAppended = 0;

  for (const sheetName of IMPORT_CONFIG.SOURCE_SHEET_NAMES || []) {
    const shSrc = ssSrc.getSheetByName(sheetName);
    if (!shSrc) {
      Logger.log(`Aba de origem não encontrada (pulando): ${sheetName}`);
      continue;
    }

    const lastRow = shSrc.getLastRow();
    if (lastRow < IMPORT_CONFIG.DATA_START_ROW) {
      Logger.log(`Sem dados para ler na origem (${sheetName}).`);
      continue;
    }

    // Ler A..G para varrer a linha inteira por termos proibidos
    const numRows = lastRow - IMPORT_CONFIG.DATA_START_ROW + 1;
    const values = shSrc
      .getRange(IMPORT_CONFIG.DATA_START_ROW, 1, numRows, IMPORT_CONFIG.COL_G)
      .getValues(); // A..G
    const bgG = shSrc
      .getRange(IMPORT_CONFIG.DATA_START_ROW, IMPORT_CONFIG.COL_G, numRows, 1)
      .getBackgrounds();

    const rowsToAppend = [];

    for (let i = 0; i < values.length; i++) {
      const row = values[i]; // A..G

      const srcB = row[IMPORT_CONFIG.COL_B - 1]; // pedido (original)
      const srcC = row[IMPORT_CONFIG.COL_C - 1]; // marketplace
      const srcD = row[IMPORT_CONFIG.COL_D - 1]; // nome
      const srcF = row[IMPORT_CONFIG.COL_F - 1]; // produto
      const srcE = row[IMPORT_CONFIG.COL_E - 1]; // qt

      const pedidoId = normalize_(srcB);
      if (!pedidoId) continue;

      // === FILTRO PELA COLUNA G ===
      const statusG = normalize_(row[IMPORT_CONFIG.COL_G - 1]).toLowerCase(); // G
      const bgColor = (bgG[i][0] || "").toLowerCase();

      // 1) Se status "ok" -> ignora
      if (statusG === "ok") continue;

      // 2) Se cor verde #00ff00 -> ignora
      if (bgColor === "#00ff00") continue;

      // 3) Se linha inteira contiver termos proibidos -> ignora
      const rowText = row.map((v) => normalize_(v).toLowerCase()).join(" | ");
      if (containsForbidden_(rowText)) continue;

      // pedido exportado (reduz para os marketplaces definidos)
      const pedidoOut = normalizePedidoForExport_(srcB, srcC);

      // Chave para evitar duplicar: PedidoOut + Marketplace + Produto + Nome + Qt
      const key = makeKey_(pedidoOut, srcC, srcF, srcD, srcE);
      if (importedSet.has(key)) continue;

      // MAPEAR COLUNAS: (pedidoOut), C, D, F, E -> A,B,C,D,E
      const destRow = [pedidoOut, srcC, srcD, srcF, srcE]; // A..E

      rowsToAppend.push(destRow);
      importedSet.add(key);
    }

    if (rowsToAppend.length > 0) {
      shDest
        .getRange(shDest.getLastRow() + 1, 1, rowsToAppend.length, 5)
        .setValues(rowsToAppend);
      totalAppended += rowsToAppend.length;
      Logger.log(
        `Importadas ${rowsToAppend.length} linhas de "${sheetName}" para "${IMPORT_CONFIG.DEST_SHEET_NAME}".`
      );
    } else {
      Logger.log(`Nenhuma linha elegível encontrada para importar em "${sheetName}".`);
    }
  }

  if (totalAppended === 0) {
    Logger.log("Nenhuma linha elegível encontrada para importar (somando todas as abas).");
  }

  // 3) Pós-processos (ordem importa)
  preencherItemKeyConferencia_(); // J
  preencherPendentesConferencia_(); // K
  marcarPedidosQueViraramProibidos_(); // L

  // >>> NOVO: B -> A (preenche SOMENTE Status vazios na Conferencia usando TENHO/FALTA da coluna H na Planilha B)
  refletirStatusDaPlanilhaBNaConferencia_SomenteVazios();

  // Mantém: A -> B (reflete status da Conferencia na Planilha B - coluna H)
  refletirStatusDaConferenciaNaPlanilhaB_();

  // 4) Executa no final: gerarListaFaltas
  try {
    gerarListaFaltas();
    Logger.log("✅ gerarListaFaltas() executado com sucesso ao final do import.");
  } catch (e) {
    Logger.log(
      "⚠️ gerarListaFaltas() falhou, mas o import já terminou. Erro: " +
        (e && e.message ? e.message : e)
    );
  }
}

/******************* PÓS-PROCESSO 0: preencher coluna J com PEDIDO#NN *******************/
function preencherItemKeyConferencia_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(IMPORT_CONFIG.DEST_SHEET_NAME);
  if (!sh) throw new Error(`Aba destino não encontrada: ${IMPORT_CONFIG.DEST_SHEET_NAME}`);

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const startRow = 2; // cabeçalho na linha 1
  const numRows = lastRow - startRow + 1;

  const dataAE = sh.getRange(startRow, 1, numRows, 5).getValues(); // A..E

  const outJ = [];
  const counterByPedido = new Map(); // pedido -> contador

  for (let i = 0; i < numRows; i++) {
    const pedido = normalize_(dataAE[i][0]); // A
    const hasAny = dataAE[i].some((v) => String(v || "").trim() !== "");

    if (!hasAny || !pedido) {
      outJ.push([""]);
      continue;
    }

    const cur = (counterByPedido.get(pedido) || 0) + 1;
    counterByPedido.set(pedido, cur);

    const suffix = String(cur).padStart(2, "0");
    outJ.push([`${pedido}#${suffix}`]);
  }

  sh.getRange(startRow, IMPORT_CONFIG.DEST_COL_ITEMKEY_J, numRows, 1).setValues(outJ);
}

/******************* PÓS-PROCESSO 1: preencher K com PENDENTE *******************/
function preencherPendentesConferencia_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(IMPORT_CONFIG.DEST_SHEET_NAME);
  if (!sh) throw new Error(`Aba destino não encontrada: ${IMPORT_CONFIG.DEST_SHEET_NAME}`);

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const startRow = 2; // cabeçalho na linha 1
  const numRows = lastRow - startRow + 1;

  const dataAE = sh.getRange(startRow, 1, numRows, 5).getValues();
  const colK = sh.getRange(startRow, IMPORT_CONFIG.DEST_COL_STATUS_K, numRows, 1).getValues();

  const outK = [];

  for (let i = 0; i < numRows; i++) {
    const row = dataAE[i];
    const hasAny = row.some((v) => String(v || "").trim() !== "");
    const currentK = String(colK[i][0] || "").trim();

    if (hasAny) {
      outK.push([currentK ? currentK : "PENDENTE"]);
    } else {
      outK.push([""]);
    }
  }

  sh.getRange(startRow, IMPORT_CONFIG.DEST_COL_STATUS_K, numRows, 1).setValues(outK);
}

/******************* PÓS-PROCESSO 2: marcar L="SIM" se pedido na A virou proibido na B Setembro+Julho+Agosto) *******************/
function marcarPedidosQueViraramProibidos_() {
  const ssDest = SpreadsheetApp.getActiveSpreadsheet();
  const shA = ssDest.getSheetByName(IMPORT_CONFIG.DEST_SHEET_NAME);
  if (!shA) throw new Error(`Aba destino não encontrada: ${IMPORT_CONFIG.DEST_SHEET_NAME}`);

  const lastRowA = shA.getLastRow();
  if (lastRowA < 2) return;

  const startRowA = 2;
  const numRowsA = lastRowA - startRowA + 1;

  const pedidosA = shA
    .getRange(startRowA, IMPORT_CONFIG.DEST_COL_PEDIDO, numRowsA, 1)
    .getValues();
  const colL = shA
    .getRange(startRowA, IMPORT_CONFIG.DEST_COL_FLAG_L, numRowsA, 1)
    .getValues();

  const pedidoToIndexes = new Map();
  for (let i = 0; i < numRowsA; i++) {
    const pedido = normalize_(pedidosA[i][0]);
    if (!pedido) continue;
    if (!pedidoToIndexes.has(pedido)) pedidoToIndexes.set(pedido, []);
    pedidoToIndexes.get(pedido).push(i);
  }
  if (pedidoToIndexes.size === 0) return;

  const ssSrc = SpreadsheetApp.openById(IMPORT_CONFIG.SOURCE_SPREADSHEET_ID);

  const proibidos = new Set();

  for (const sheetName of IMPORT_CONFIG.SOURCE_SHEET_NAMES || []) {
    const shB = ssSrc.getSheetByName(sheetName);
    if (!shB) {
      Logger.log(`Aba de origem não encontrada (pulando): ${sheetName}`);
      continue;
    }

    const lastRowB = shB.getLastRow();
    if (lastRowB < IMPORT_CONFIG.DATA_START_ROW) continue;

    const numRowsB = lastRowB - IMPORT_CONFIG.DATA_START_ROW + 1;
    const valuesB = shB
      .getRange(IMPORT_CONFIG.DATA_START_ROW, 1, numRowsB, IMPORT_CONFIG.COL_G)
      .getValues(); // A..G
    const bgG = shB
      .getRange(IMPORT_CONFIG.DATA_START_ROW, IMPORT_CONFIG.COL_G, numRowsB, 1)
      .getBackgrounds();

    for (let i = 0; i < valuesB.length; i++) {
      const row = valuesB[i];

      const pedidoIdRaw = row[IMPORT_CONFIG.COL_B - 1]; // B
      const marketRaw = row[IMPORT_CONFIG.COL_C - 1]; // C
      const pedidoId = normalizePedidoForExport_(pedidoIdRaw, marketRaw); // considera corte
      if (!pedidoId) continue;

      if (!pedidoToIndexes.has(pedidoId)) continue;

      const statusG = normalize_(row[IMPORT_CONFIG.COL_G - 1]).toLowerCase();
      const bgColor = (bgG[i][0] || "").toLowerCase();
      const rowText = row.map((v) => normalize_(v).toLowerCase()).join(" | ");

      const isOkText = statusG === "ok";
      const isGreen = bgColor === "#00ff00";
      const hasForbidden = containsForbidden_(rowText);

      if (isOkText || isGreen || hasForbidden) {
        proibidos.add(pedidoId);
      }
    }
  }

  if (proibidos.size === 0) {
    Logger.log(
      "Nenhum pedido da Conferencia entrou em critério proibido na Planilha B (Julho/Setembro/Agosto)."
    );
    return;
  }

  const outL = colL.map((r) => [String(r[0] || "").trim()]);

  proibidos.forEach((pedido) => {
    const idxs = pedidoToIndexes.get(pedido) || [];
    for (const idx of idxs) {
      outL[idx][0] = "SIM";
    }
  });

  shA.getRange(startRowA, IMPORT_CONFIG.DEST_COL_FLAG_L, numRowsA, 1).setValues(outL);
  Logger.log(`Marcados ${proibidos.size} pedidos com "SIM" na coluna L.`);
}

/******************* NOVO (OTIMIZADO): B -> A SOMENTE STATUS VAZIO *******************/
function refletirStatusDaPlanilhaBNaConferencia_SomenteVazios() {
  const ssA = SpreadsheetApp.getActiveSpreadsheet();
  const shA = ssA.getSheetByName(IMPORT_CONFIG.DEST_SHEET_NAME);
  if (!shA) throw new Error(`Aba destino não encontrada: ${IMPORT_CONFIG.DEST_SHEET_NAME}`);

  const lastRowA = shA.getLastRow();
  if (lastRowA < 2) {
    Logger.log("B->A (vazios): Conferencia sem linhas para atualizar.");
    return;
  }

  const startRowA = 2;
  const numRowsA = lastRowA - startRowA + 1;

  // Lê A..F (Pedido, Market, Nome, Produto, Qt, Status)
  const dataAF = shA.getRange(startRowA, 1, numRowsA, 6).getValues();

  // key -> índices na Conferencia que estão com Status vazio
  const pendingKeyToIdxs = new Map();

  for (let i = 0; i < dataAF.length; i++) {
    const pedido = dataAF[i][0];
    const market = dataAF[i][1];
    const nome = dataAF[i][2];
    const produto = dataAF[i][3];
    const qt = dataAF[i][4];
    const statusA = dataAF[i][5];

    if (normalize_(statusA)) continue; // só vazios

    const key = makeKey_(pedido, market, produto, nome, qt);
    if (!pendingKeyToIdxs.has(key)) pendingKeyToIdxs.set(key, []);
    pendingKeyToIdxs.get(key).push(i);
  }

  if (pendingKeyToIdxs.size === 0) {
    Logger.log("B->A (vazios): não há Status vazio na Conferencia.");
    return;
  }

  const ssB = SpreadsheetApp.openById(IMPORT_CONFIG.SOURCE_SPREADSHEET_ID);

  // key -> status (TENHO / FALTA / FALTA - ...)
  const statusByKey = new Map();

  // Para early-stop: quantas chaves pendentes ainda faltam encontrar
  let remaining = pendingKeyToIdxs.size;

  for (const sheetName of IMPORT_CONFIG.SOURCE_SHEET_NAMES || []) {
    if (remaining <= 0) break;

    const shB = ssB.getSheetByName(sheetName);
    if (!shB) {
      Logger.log(`B->A (vazios): Aba de origem não encontrada (pulando): ${sheetName}`);
      continue;
    }

    const lastRowB = shB.getLastRow();
    if (lastRowB < IMPORT_CONFIG.DATA_START_ROW) continue;

    const numRowsB = lastRowB - IMPORT_CONFIG.DATA_START_ROW + 1;

    // lê até H (8)
    const maxCol = Math.max(IMPORT_CONFIG.COL_F, IMPORT_CONFIG.SOURCE_WRITEBACK_COL_J); // 8
    const valuesB = shB.getRange(IMPORT_CONFIG.DATA_START_ROW, 1, numRowsB, maxCol).getValues();

    for (let i = 0; i < valuesB.length; i++) {
      if (remaining <= 0) break;

      const row = valuesB[i];

      const pedidoRaw = row[IMPORT_CONFIG.COL_B - 1]; // B
      const market = row[IMPORT_CONFIG.COL_C - 1]; // C
      const nome = row[IMPORT_CONFIG.COL_D - 1]; // D
      const qt = row[IMPORT_CONFIG.COL_E - 1]; // E
      const produto = row[IMPORT_CONFIG.COL_F - 1]; // F

      const statusH = row[IMPORT_CONFIG.SOURCE_WRITEBACK_COL_J - 1]; // H
      const s = normalize_(statusH).toUpperCase();

      // Só aceita TENHO ou FALTA (inclui "FALTA - ...")
      if (!(s === "TENHO" || s.indexOf("FALTA") === 0)) continue;

      const pedido = normalizePedidoForExport_(pedidoRaw, market);
      if (!normalize_(pedido)) continue;

      const key = makeKey_(pedido, market, produto, nome, qt);

      // Otimização: só considera se está pendente na A
      if (!pendingKeyToIdxs.has(key)) continue;

      // Se ainda não tínhamos esse key, diminui remaining (early stop)
      if (!statusByKey.has(key)) remaining--;

      // Último prevalece conforme ordem das abas
      statusByKey.set(key, s);
    }
  }

  if (statusByKey.size === 0) {
    Logger.log("B->A (vazios): nenhum TENHO/FALTA encontrado na B para os pendentes.");
    return;
  }

  // Atualiza somente os vazios encontrados
  const outF = dataAF.map((r) => [r[5]]);
  let updated = 0;

  for (const [key, newStatus] of statusByKey.entries()) {
    const idxs = pendingKeyToIdxs.get(key) || [];
    for (const idx of idxs) {
      const cur = normalize_(outF[idx][0]).toUpperCase();
      if (cur) continue; // segurança

      outF[idx][0] = newStatus;
      updated++;
    }
  }

  if (updated > 0) {
    shA.getRange(startRowA, IMPORT_CONFIG.DEST_COL_STATUS_F, numRowsA, 1).setValues(outF);
    Logger.log(`B->A (vazios): preenchidas ${updated} células na Conferencia!F.`);
  } else {
    Logger.log("B->A (vazios): nada para atualizar.");
  }
}

/******************* PÓS-PROCESSO 3: A -> B (reflete Status (F) da Conferencia na col H da Planilha B) *******************/
function refletirStatusDaConferenciaNaPlanilhaB_() {
  const ssA = SpreadsheetApp.getActiveSpreadsheet();
  const shA = ssA.getSheetByName(IMPORT_CONFIG.DEST_SHEET_NAME);
  if (!shA) throw new Error(`Aba destino não encontrada: ${IMPORT_CONFIG.DEST_SHEET_NAME}`);

  const lastRowA = shA.getLastRow();
  if (lastRowA < 2) return;

  const startRowA = 2;
  const numRowsA = lastRowA - startRowA + 1;

  // Lê A..F (Pedido, Market, Nome, Produto, Qt, Status)
  const dataAF = shA.getRange(startRowA, 1, numRowsA, 6).getValues(); // A..F

  // Mapa: key -> status (somente status preenchido)
  const statusByKey = new Map();
  for (const r of dataAF) {
    const pedido = r[0];
    const market = r[1];
    const nome = r[2];
    const produto = r[3];
    const qt = r[4];
    const status = r[5]; // F

    const s = normalize_(status);
    if (!s) continue;

    const key = makeKey_(pedido, market, produto, nome, qt);
    statusByKey.set(key, s);
  }

  if (statusByKey.size === 0) {
    Logger.log("A->B: nenhum status preenchido na Conferencia (coluna F).");
    return;
  }

  const ssB = SpreadsheetApp.openById(IMPORT_CONFIG.SOURCE_SPREADSHEET_ID);

  for (const sheetName of IMPORT_CONFIG.SOURCE_SHEET_NAMES || []) {
    const shB = ssB.getSheetByName(sheetName);
    if (!shB) {
      Logger.log(`A->B: Aba de origem não encontrada (pulando): ${sheetName}`);
      continue;
    }

    const lastRowB = shB.getLastRow();
    if (lastRowB < IMPORT_CONFIG.DATA_START_ROW) continue;

    const numRowsB = lastRowB - IMPORT_CONFIG.DATA_START_ROW + 1;

    const maxColToRead = Math.max(IMPORT_CONFIG.COL_F, IMPORT_CONFIG.SOURCE_WRITEBACK_COL_J); // 8 (até H)
    const valuesB = shB.getRange(IMPORT_CONFIG.DATA_START_ROW, 1, numRowsB, maxColToRead).getValues();

    const currentH = shB
      .getRange(IMPORT_CONFIG.DATA_START_ROW, IMPORT_CONFIG.SOURCE_WRITEBACK_COL_J, numRowsB, 1)
      .getValues();

    const outH = currentH.map((r) => [r[0]]);
    let updated = 0;

    for (let i = 0; i < valuesB.length; i++) {
      const row = valuesB[i];

      const pedidoRaw = row[IMPORT_CONFIG.COL_B - 1]; // B
      const market = row[IMPORT_CONFIG.COL_C - 1]; // C
      const nome = row[IMPORT_CONFIG.COL_D - 1]; // D
      const qt = row[IMPORT_CONFIG.COL_E - 1]; // E
      const produto = row[IMPORT_CONFIG.COL_F - 1]; // F

      const pedido = normalizePedidoForExport_(pedidoRaw, market); // considera corte
      const key = makeKey_(pedido, market, produto, nome, qt);

      if (!statusByKey.has(key)) continue;

      const newStatus = statusByKey.get(key);
      const oldStatus = normalize_(outH[i][0]);

      if (normalize_(newStatus) !== oldStatus) {
        outH[i][0] = newStatus;
        updated++;
      }
    }

    if (updated > 0) {
      shB
        .getRange(IMPORT_CONFIG.DATA_START_ROW, IMPORT_CONFIG.SOURCE_WRITEBACK_COL_J, numRowsB, 1)
        .setValues(outH);
      Logger.log(`A->B (${sheetName}): atualizadas ${updated} células na coluna H.`);
    } else {
      Logger.log(`A->B (${sheetName}): nada para atualizar (H já estava igual).`);
    }
  }
}

/******************* HELPERS *******************/
function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function normalize_(v) {
  if (v === null || v === undefined) return "";
  return String(v).trim();
}

// corta pedido conforme o marketplace
function normalizePedidoForExport_(pedido, marketplace) {
  const p = normalize_(pedido);
  const m = normalize_(marketplace);

  const maxLen = Number(IMPORT_CONFIG.SHORT_PEDIDO_MAXLEN || 13);

  // Marketplaces que usam os 13 PRIMEIROS caracteres
  if (
    (IMPORT_CONFIG.SHORT_PEDIDO_FIRST_MARKETS || []).indexOf(m) >= 0
  ) {
    return p.length > maxLen ? p.slice(0, maxLen) : p;
  }

  // Marketplaces que usam os 13 ÚLTIMOS caracteres
  if (
    (IMPORT_CONFIG.SHORT_PEDIDO_LAST_MARKETS || []).indexOf(m) >= 0
  ) {
    return p.length > maxLen ? p.slice(-maxLen) : p;
  }

  // Demais marketplaces: pedido completo
  return p;
}

// inclui marketplace na chave para evitar colisões entre sistemas
function makeKey_(pedido, marketplace, produto, nome, qt) {
  const parts = [pedido, marketplace, produto, nome, qt].map((v) => normalize_(v).toLowerCase());
  return parts.join("||");
}

function containsForbidden_(rowTextLower) {
  for (const t of IMPORT_CONFIG.FORBIDDEN_TERMS || []) {
    if (rowTextLower.includes(String(t).toLowerCase())) return true;
  }
  return false;
}

function loadExistingKeysFromConferencia_(shDest) {
  const lastRow = shDest.getLastRow();
  if (lastRow < 2) return new Set();

  const startRow = 2;
  const numRows = lastRow - startRow + 1;

  // A..E (Pedido, Market, Nome, Produto, Qt)
  const data = shDest.getRange(startRow, 1, numRows, 5).getValues();
  const set = new Set();

  for (const r of data) {
    const pedido = normalize_(r[0]); // A
    const market = normalize_(r[1]); // B
    const nome = normalize_(r[2]); // C
    const produto = normalize_(r[3]); // D
    const qt = normalize_(r[4]); // E

    if (!pedido) continue;
    const key = makeKey_(pedido, market, produto, nome, qt);
    set.add(key);
  }

  return set;
}

/**********************************************************************
 *  SCRIPT ADICIONADO: gerarListaFaltas (executa ao final do import)
 *  ✅ AJUSTADO para:
 *  - "FALTA" => inclui todos os itens do pedido (como antes)
 *  - "FALTA - termo1; termo2; ..." => inclui SOMENTE os itens que contêm qualquer termo (termo inteiro)
 *    • termo não encontrado: ignora e segue
 *    • nenhum termo bateu: não inclui todos, apenas pula (e registra como não encontrado do produto original)
 *
 *  ✅ NOVO AJUSTE pedido por você:
 *  - Considera a coluna E (Qtd) da aba "Conferencia" para somar quantidades na aba "Lista".
 *    Ex.: se Conferencia!E=2, conta 2x aquele produto nas somas (A:B e D:E).
 **********************************************************************/
function gerarListaFaltas() {
  const ss = SpreadsheetApp.getActive();

  const shConf = ss.getSheetByName("Conferencia");
  const shDados = ss.getSheetByName("Dados");
  const shLista = ss.getSheetByName("Lista") || ss.insertSheet("Lista");

  if (!shConf) throw new Error('Aba "Conferencia" não encontrada.');
  if (!shDados) throw new Error('Aba "Dados" não encontrada.');

  // ===== 0) Limpa saída mantendo cabeçalhos (linha 1) =====
  limparSaidaLista_(shLista);

  // ===== 1) Lê Conferencia (Produto D, Qtd E, Status F, DiscordMessageId I) =====
  const lastConfRow = shConf.getLastRow();
  if (lastConfRow < 2) {
    shLista.getRange("C2").setValue("Sem dados em Conferencia.");
    return;
  }

  const numConfRows = lastConfRow - 1;

  const confProdutos = shConf.getRange(2, 4, numConfRows, 1).getValues(); // D
  const confQtd = shConf.getRange(2, 5, numConfRows, 1).getValues(); // E
  const confStatus = shConf.getRange(2, 6, numConfRows, 1).getValues(); // F
  const confMsgId = shConf.getRange(2, 9, numConfRows, 1).getValues(); // I

  const faltasItens = []; // [{prod, terms:[], qty}]
  const todosItens = []; // [{prod,id, qty}]

  for (let i = 0; i < numConfRows; i++) {
    const prod = String(confProdutos[i][0] ?? "").trim();
    if (!prod) continue;

    const qty = parseQtdConferencia_(confQtd[i][0]); // NOVO
    const id = String(confMsgId[i][0] ?? "").trim();

    todosItens.push({ prod, id, qty });

    const stRaw = String(confStatus[i][0] ?? "").trim();
    const stUp = stRaw.toUpperCase();

    if (stUp.includes("FALTA")) {
      const terms = extrairFaltaTermos_(stRaw);
      faltasItens.push({ prod, terms, qty });
    }
  }

  if (todosItens.length === 0) {
    shLista.getRange("C2").setValue("Nenhum produto encontrado em Conferencia.");
    return;
  }

  // ===== 2) Varrendo Dados A:Z e montando mapa Produto -> valor à direita =====
  const lastDadosRow = shDados.getLastRow();
  const lastDadosCol = Math.min(shDados.getLastColumn(), 26); // até Z

  if (lastDadosRow < 1 || lastDadosCol < 2) {
    shLista.getRange("C2").setValue('Aba "Dados" vazia ou com poucas colunas.');
    return;
  }

  const dadosVals = shDados.getRange(1, 1, lastDadosRow, lastDadosCol).getValues();

  const wanted = new Set(
    [...faltasItens.map((x) => x.prod), ...todosItens.map((x) => x.prod)].map((p) =>
      normalizarBusca_(p)
    )
  );

  const mapSubseq = new Map(); // normalizedProduct -> subseqValue (à direita)

  for (let r = 0; r < dadosVals.length; r++) {
    const row = dadosVals[r];
    for (let c = 0; c < row.length - 1; c++) {
      const cell = String(row[c] ?? "").trim();
      if (!cell) continue;

      const key = normalizarBusca_(cell);
      if (!wanted.has(key)) continue;
      if (mapSubseq.has(key)) continue;

      const rightVal = String(row[c + 1] ?? "").trim();
      mapSubseq.set(key, rightVal);
    }
  }

  // ===== 3) Processadores =====
  const notFoundAll = []; // Lista!C (união dos não encontrados)

  // 3.1) A:B (somente FALTA) — com filtro por termos quando houver "FALTA - ..."
  function processarFaltas_(listaItens) {
    const counts = new Map(); // nomeExato -> qtd

    for (const it of listaItens) {
      const prod = it.prod;
      const terms = Array.isArray(it.terms) ? it.terms : [];
      const mult = Number(it.qty || 1); // NOVO: multiplicador pela Qtd da Conferencia

      const key = normalizarBusca_(prod);
      const subseq = mapSubseq.get(key);

      if (!subseq) {
        notFoundAll.push(prod);
        continue;
      }

      const partsAll = subseq
        .split(";")
        .map((s) => s.trim())
        .filter(Boolean);

      if (partsAll.length === 0) {
        notFoundAll.push(prod);
        continue;
      }

      let partsToCount = partsAll;

      if (terms.length) {
        const matched = [];
        const matchedSet = new Set();

        for (const p of partsAll) {
          const pNorm = normalizarBusca_(p);

          for (const term of terms) {
            const needle = normalizarBusca_(term);
            if (needle && pNorm.includes(needle)) {
              if (!matchedSet.has(p)) {
                matchedSet.add(p);
                matched.push(p);
              }
              break;
            }
          }
        }

        partsToCount = matched;

        // nenhum termo bateu => pula (NÃO inclui todos)
        if (!partsToCount.length) {
          notFoundAll.push(prod);
          continue;
        }
      }

      for (const p of partsToCount) {
        counts.set(p, (counts.get(p) || 0) + mult); // NOVO: soma pela qtd
      }
    }

    return Array.from(counts.entries())
      .sort((a, b) => a[0].localeCompare(b[0], "pt-BR"))
      .map(([name, qty]) => [name, qty]);
  }

  // 3.2) D:E:F (Lista completa) — com IDs por item
  // ✅ Mantém IDs únicos (não repete ID pela Qtd) para não poluir, mas a quantidade (col E) soma com Qtd.
  function processarTodosComIds_(itens) {
    const counts = new Map();
    const idsMap = new Map();

    for (const item of itens) {
      const prod = item.prod;
      const id = item.id;
      const mult = Number(item.qty || 1); // NOVO

      const key = normalizarBusca_(prod);
      const subseq = mapSubseq.get(key);

      if (!subseq) {
        notFoundAll.push(prod);
        continue;
      }

      const parts = subseq
        .split(";")
        .map((s) => s.trim())
        .filter(Boolean);

      if (parts.length === 0) {
        notFoundAll.push(prod);
        continue;
      }

      for (const p of parts) {
        counts.set(p, (counts.get(p) || 0) + mult); // NOVO: soma pela qtd

        if (!idsMap.has(p)) idsMap.set(p, { seen: new Set(), list: [] });

        if (id) {
          const bucket = idsMap.get(p);
          if (!bucket.seen.has(id)) {
            bucket.seen.add(id);
            bucket.list.push(id);
          }
        }
      }
    }

    const namesSorted = Array.from(counts.keys()).sort((a, b) =>
      a.localeCompare(b, "pt-BR")
    );

    return namesSorted.map((name) => {
      const qty = counts.get(name) || 0;
      const bucket = idsMap.get(name);
      const idsText = bucket && bucket.list.length ? bucket.list.join("\n") : "";
      return [name, qty, idsText];
    });
  }

  const foundRowsFaltas = processarFaltas_(faltasItens); // A:B
  const foundRowsTodos = processarTodosComIds_(todosItens); // D:E:F

  // ===== 4) Escreve na aba Lista =====
  if (foundRowsFaltas.length) {
    shLista.getRange(2, 1, foundRowsFaltas.length, 2).setValues(foundRowsFaltas);
  } else {
    shLista.getRange("A2").setValue("Nenhum item com FALTA.");
  }

  if (foundRowsTodos.length) {
    shLista.getRange(2, 4, foundRowsTodos.length, 3).setValues(foundRowsTodos); // D:E:F
  }

  if (notFoundAll.length) {
    const seenNorm = new Set();
    const nf = [];

    for (const x of notFoundAll) {
      const k = normalizarBusca_(x);
      if (seenNorm.has(k)) continue;
      seenNorm.add(k);
      nf.push([x]);
    }

    shLista.getRange(2, 3, nf.length, 1).setValues(nf); // C
  }
}

function limparSaidaLista_(shLista) {
  const lastRow = shLista.getMaxRows();
  if (lastRow >= 2) {
    shLista.getRange(2, 1, lastRow - 1, 6).clearContent();
  }
}

/**
 * Extrai termos após "FALTA -", aceitando múltiplos separados por ";"
 * Ex:
 *  "FALTA - Dark Tabc; Lavanda; Terre" -> ["Dark Tabc","Lavanda","Terre"]
 *  "FALTA - Dark Tabc" -> ["Dark Tabc"]
 *  "FALTA" -> []
 */
function extrairFaltaTermos_(statusRaw) {
  const s = String(statusRaw || "").trim();
  const up = s.toUpperCase();

  if (up.indexOf("FALTA") < 0) return [];

  const pos = up.indexOf("FALTA -");
  if (pos < 0) return []; // "FALTA" sem complemento

  const after = s.slice(pos + "FALTA -".length).trim();
  if (!after) return [];

  return after
    .split(";")
    .map((x) => String(x || "").trim())
    .filter(Boolean);
}

// Normalização para busca (case-insensitive + remove acentos + colapsa espaços)
function normalizarBusca_(s) {
  return String(s ?? "")
    .trim()
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ");
}

// Mantive seu normalizar_ original (usado em outras partes do script)
function normalizar_(s) {
  return String(s ?? "")
    .trim()
    .toUpperCase()
    .replace(/\s+/g, " ");
}

/**
 * NOVO: interpreta Qtd da Conferencia!E
 * - Aceita número ou texto ("2", "2,0", etc.)
 * - Qualquer coisa inválida/vazia vira 1
 * - Decimais: piso (2.7 -> 2) para evitar inflar sem querer
 */
function parseQtdConferencia_(v) {
  if (v === null || v === undefined || v === "") return 1;

  const s = String(v).trim().replace(",", ".");
  const n = Number(s);

  if (!isFinite(n) || n <= 0) return 1;

  return Math.floor(n);
}
