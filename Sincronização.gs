/**************************************************************************
 * Importação / sincronização entre a Planilha B (vendas) e a Planilha A
 * (aba "Conferencia"), mais a geração da aba "Lista".
 *
 * Chamada pelo Web App através da ação "importacao" (ver Code.gs).
 **************************************************************************/

/******************* CONFIG *******************/
const IMPORT_CONFIG = {
  // ====== PLANILHA B (ORIGEM) ======
  SOURCE_SPREADSHEET_ID: "1RHTxClkkFamZUE4hM8romnZfWWTecNwvkvI4SO7rYmU",

  /**
   * Abas de origem.
   *
   * Deixe VAZIO para detectar automaticamente os últimos N meses (não precisa
   * mais editar o script todo mês). Preencha para fixar manualmente — a lista
   * manual sempre tem prioridade.
   *
   * ⚠️ A ordem importa: em caso de conflito, a ÚLTIMA aba da lista prevalece.
   * Na detecção automática a ordem é cronológica (mês mais recente por último).
   */
  SOURCE_SHEET_NAMES: [],
  SOURCE_SHEET_AUTO: true,
  SOURCE_SHEET_AUTO_MESES: 3,

  // Dados começam na linha 6
  DATA_START_ROW: 6,

  // Colunas fixas (por número) na Planilha B
  COL_B: 2, // Pedido
  COL_C: 3, // Marketplace
  COL_D: 4, // Nome
  COL_E: 5, // Qt
  COL_F: 6, // Produto
  COL_G: 7, // Status (ok / cor verde etc.)

  // Coluna H da Planilha B: onde refletimos o Status da Planilha A
  SOURCE_WRITEBACK_COL_H: 8,

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

  // ====== TERMOS PROIBIDOS (varridos na linha inteira A..G da Planilha B) ======
  FORBIDDEN_TERMS: ["fulfillment", "full", "cancelado"],

  /**
   * "word"      = só casa a palavra inteira (padrão).
   *               Evita descartar produtos como "Fullaroma" por causa de "full".
   * "substring" = comportamento antigo (casa em qualquer pedaço do texto).
   */
  FORBIDDEN_MATCH: "word",

  // ====== COLUNAS DESTINO (aba Conferencia) ======
  DEST_COL_PEDIDO: 1,        // A
  DEST_COL_MARKET: 2,        // B
  DEST_COL_NOME: 3,          // C
  DEST_COL_PRODUTO: 4,       // D
  DEST_COL_QT: 5,            // E
  DEST_COL_STATUS_ITEM: 6,   // F  => TENHO / FALTA / "FALTA - obs"
  DEST_COL_ITEMKEY: 10,      // J  => "PEDIDO#01"
  DEST_COL_STATUS_PEDIDO: 11,// K  => PENDENTE

  /**
   * L => "Confirmado"
   *
   * ⚠️ ATENÇÃO: esta coluna tem DOIS significados que são o mesmo sinal:
   *   1) Aqui (importação) ela recebe "SIM" quando o pedido passou a atender
   *      um critério de exclusão na Planilha B (status "ok", fundo verde ou
   *      termo proibido como "cancelado").
   *   2) No Code.gs ela é lida como "Confirmado = SIM" para considerar o
   *      pedido encerrado (list_confirmed, list_pendentes_status_vazio...).
   *
   * Ou seja: "resolvido/cancelado na origem" e "confirmado" são a mesma coisa
   * para este sistema. Se um dia precisarem ser distintos, crie uma coluna nova
   * em vez de mudar o significado desta.
   */
  DEST_COL_CONFIRMADO: 12,   // L
};

const CONFIRMADO_SIM = "SIM";

/******************* FUNÇÃO PRINCIPAL *******************/
/**
 * Importa da Planilha B para a aba Conferencia e roda os pós-processos.
 * Retorna estatísticas usadas no resumo mostrado no Discord.
 */
function importarDaPlanilhaB() {
  const ssDest = SpreadsheetApp.getActiveSpreadsheet();
  const shDest = getOrCreateSheet_(ssDest, IMPORT_CONFIG.DEST_SHEET_NAME);

  const avisos = [];

  // 1) Carrega chaves já existentes na Conferencia
  const importedSet = loadExistingKeysFromConferencia_(shDest);

  // 2) Abre origem e resolve quais abas serão lidas
  const ssSrc = SpreadsheetApp.openById(IMPORT_CONFIG.SOURCE_SPREADSHEET_ID);
  const { nomes: sheetNames, faltando } = resolveSourceSheetNames_(ssSrc);

  if (faltando.length) {
    avisos.push(`Abas não encontradas na Planilha B: ${faltando.join(", ")}.`);
  }
  if (!sheetNames.length) {
    throw new Error("Nenhuma aba de origem encontrada na Planilha B. Preencha SOURCE_SHEET_NAMES manualmente.");
  }

  let totalAppended = 0;

  // Detecta pedidos diferentes que colidem depois do corte de 13 caracteres
  const cortePorPedido = new Map(); // pedidoCortado -> Set(pedidoOriginal)

  for (const sheetName of sheetNames) {
    const shSrc = ssSrc.getSheetByName(sheetName);
    if (!shSrc) continue;

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
      const srcE = row[IMPORT_CONFIG.COL_E - 1]; // qt
      const srcF = row[IMPORT_CONFIG.COL_F - 1]; // produto

      const pedidoId = normalize_(srcB);
      if (!pedidoId) continue;

      // === FILTRO PELA COLUNA G ===
      const statusG = normalize_(row[IMPORT_CONFIG.COL_G - 1]).toLowerCase();
      const bgColor = (bgG[i][0] || "").toLowerCase();

      // 1) Se status "ok" -> ignora
      if (statusG === "ok") continue;

      // 2) Se cor verde #00ff00 -> ignora
      if (bgColor === "#00ff00") continue;

      // 3) Se linha inteira contiver termos proibidos -> ignora
      if (containsForbidden_(rowTextParaBusca_(row))) continue;

      // pedido exportado (reduz para os marketplaces definidos)
      const pedidoOut = normalizePedidoForExport_(srcB, srcC);

      if (pedidoOut !== pedidoId) {
        if (!cortePorPedido.has(pedidoOut)) cortePorPedido.set(pedidoOut, new Set());
        cortePorPedido.get(pedidoOut).add(pedidoId);
      }

      // Chave para evitar duplicar: PedidoOut + Marketplace + Produto + Nome + Qt
      const key = makeKey_(pedidoOut, srcC, srcF, srcD, srcE);
      if (importedSet.has(key)) continue;

      // MAPEAR COLUNAS: (pedidoOut), C, D, F, E -> A,B,C,D,E
      rowsToAppend.push([pedidoOut, srcC, srcD, srcF, srcE]);
      importedSet.add(key);
    }

    if (rowsToAppend.length > 0) {
      shDest
        .getRange(shDest.getLastRow() + 1, 1, rowsToAppend.length, 5)
        .setValues(rowsToAppend);
      totalAppended += rowsToAppend.length;
      Logger.log(`Importadas ${rowsToAppend.length} linhas de "${sheetName}".`);
    } else {
      Logger.log(`Nenhuma linha elegível para importar em "${sheetName}".`);
    }
  }

  // Aviso de colisão no corte de 13 caracteres
  for (const [curto, originais] of cortePorPedido.entries()) {
    if (originais.size > 1) {
      avisos.push(
        `Pedidos diferentes viraram o mesmo código "${curto}" após o corte de ` +
        `${IMPORT_CONFIG.SHORT_PEDIDO_MAXLEN} caracteres: ${Array.from(originais).join(", ")}.`
      );
    }
  }

  if (totalAppended === 0) {
    Logger.log("Nenhuma linha elegível encontrada para importar (somando todas as abas).");
  }

  // 3) Pós-processos (ordem importa)
  preencherItemKeyConferencia_();      // J
  preencherPendentesConferencia_();    // K
  marcarPedidosQueViraramProibidos_(sheetNames); // L

  // B -> A: preenche SOMENTE Status vazios usando TENHO/FALTA da coluna H
  refletirStatusDaPlanilhaBNaConferencia_SomenteVazios(sheetNames);

  // A -> B: reflete o status da Conferencia na coluna H da Planilha B
  refletirStatusDaConferenciaNaPlanilhaB_(sheetNames);

  // 4) Executa no final: gerarListaFaltas
  try {
    gerarListaFaltas();
    Logger.log("✅ gerarListaFaltas() executado com sucesso ao final do import.");
  } catch (e) {
    const msg = String(e && e.message ? e.message : e);
    avisos.push(`gerarListaFaltas() falhou: ${msg}`);
    Logger.log("⚠️ gerarListaFaltas() falhou, mas o import já terminou. Erro: " + msg);
  }

  return { importadas: totalAppended, abas: sheetNames, avisos };
}

/**
 * Alias mantido para não quebrar acionadores (triggers) já configurados na
 * interface do Apps Script com o nome antigo.
 */
function importarDaPlanilhaB_Julho_Setembro_Agosto() {
  return importarDaPlanilhaB();
}

/******************* RESOLUÇÃO DAS ABAS DE ORIGEM *******************/
const MESES_PT = [
  "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
  "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro",
];

/** Últimos N meses incluindo o atual, em ordem cronológica. */
function mesesRecentes_(n) {
  const qtd = Math.max(1, Math.min(12, Number(n) || 3));
  const hoje = new Date();
  const out = [];

  for (let i = qtd - 1; i >= 0; i--) {
    const d = new Date(hoje.getFullYear(), hoje.getMonth() - i, 1);
    out.push(MESES_PT[d.getMonth()]);
  }

  return out;
}

/**
 * Decide quais abas da Planilha B serão lidas.
 * Lista manual tem prioridade; senão usa os últimos N meses.
 * A comparação ignora acentos e maiúsculas ("marco" acha "Março").
 */
function resolveSourceSheetNames_(ssSrc) {
  const byNorm = new Map();
  for (const s of ssSrc.getSheets()) {
    byNorm.set(normalizarBusca_(s.getName()), s.getName());
  }

  const manual = (IMPORT_CONFIG.SOURCE_SHEET_NAMES || []).filter(Boolean);

  const desejados = manual.length
    ? manual
    : (IMPORT_CONFIG.SOURCE_SHEET_AUTO ? mesesRecentes_(IMPORT_CONFIG.SOURCE_SHEET_AUTO_MESES) : []);

  const nomes = [];
  const faltando = [];

  for (const d of desejados) {
    const real = byNorm.get(normalizarBusca_(d));
    if (real) nomes.push(real);
    else faltando.push(d);
  }

  Logger.log(`Abas de origem: ${nomes.join(", ") || "(nenhuma)"}${faltando.length ? " | ausentes: " + faltando.join(", ") : ""}`);

  return { nomes, faltando };
}

/******************* PÓS-PROCESSO 0: coluna J com PEDIDO#NN *******************/
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
  const counterByPedido = new Map();

  for (let i = 0; i < numRows; i++) {
    const pedido = normalize_(dataAE[i][0]); // A
    const hasAny = dataAE[i].some((v) => String(v || "").trim() !== "");

    if (!hasAny || !pedido) {
      outJ.push([""]);
      continue;
    }

    const cur = (counterByPedido.get(pedido) || 0) + 1;
    counterByPedido.set(pedido, cur);

    outJ.push([`${pedido}#${String(cur).padStart(2, "0")}`]);
  }

  sh.getRange(startRow, IMPORT_CONFIG.DEST_COL_ITEMKEY, numRows, 1).setValues(outJ);
}

/******************* PÓS-PROCESSO 1: coluna K com PENDENTE *******************/
function preencherPendentesConferencia_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(IMPORT_CONFIG.DEST_SHEET_NAME);
  if (!sh) throw new Error(`Aba destino não encontrada: ${IMPORT_CONFIG.DEST_SHEET_NAME}`);

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const startRow = 2;
  const numRows = lastRow - startRow + 1;

  const dataAE = sh.getRange(startRow, 1, numRows, 5).getValues();
  const colK = sh.getRange(startRow, IMPORT_CONFIG.DEST_COL_STATUS_PEDIDO, numRows, 1).getValues();

  const outK = [];

  for (let i = 0; i < numRows; i++) {
    const hasAny = dataAE[i].some((v) => String(v || "").trim() !== "");
    const currentK = String(colK[i][0] || "").trim();

    outK.push([hasAny ? (currentK ? currentK : "PENDENTE") : ""]);
  }

  sh.getRange(startRow, IMPORT_CONFIG.DEST_COL_STATUS_PEDIDO, numRows, 1).setValues(outK);
}

/******************* PÓS-PROCESSO 2: coluna L = "SIM" *******************/
/**
 * Marca "SIM" na coluna Confirmado (L) dos pedidos que passaram a atender um
 * critério de exclusão na Planilha B. Ver comentário em DEST_COL_CONFIRMADO.
 */
function marcarPedidosQueViraramProibidos_(sheetNames) {
  const ssDest = SpreadsheetApp.getActiveSpreadsheet();
  const shA = ssDest.getSheetByName(IMPORT_CONFIG.DEST_SHEET_NAME);
  if (!shA) throw new Error(`Aba destino não encontrada: ${IMPORT_CONFIG.DEST_SHEET_NAME}`);

  const lastRowA = shA.getLastRow();
  if (lastRowA < 2) return;

  const startRowA = 2;
  const numRowsA = lastRowA - startRowA + 1;

  const pedidosA = shA.getRange(startRowA, IMPORT_CONFIG.DEST_COL_PEDIDO, numRowsA, 1).getValues();
  const colL = shA.getRange(startRowA, IMPORT_CONFIG.DEST_COL_CONFIRMADO, numRowsA, 1).getValues();

  const pedidoToIndexes = new Map();
  for (let i = 0; i < numRowsA; i++) {
    const pedido = normalize_(pedidosA[i][0]);
    if (!pedido) continue;
    if (!pedidoToIndexes.has(pedido)) pedidoToIndexes.set(pedido, []);
    pedidoToIndexes.get(pedido).push(i);
  }
  if (pedidoToIndexes.size === 0) return;

  const ssSrc = SpreadsheetApp.openById(IMPORT_CONFIG.SOURCE_SPREADSHEET_ID);
  const nomes = sheetNames || resolveSourceSheetNames_(ssSrc).nomes;

  const proibidos = new Set();

  for (const sheetName of nomes) {
    const shB = ssSrc.getSheetByName(sheetName);
    if (!shB) continue;

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

      const pedidoId = normalizePedidoForExport_(
        row[IMPORT_CONFIG.COL_B - 1],
        row[IMPORT_CONFIG.COL_C - 1]
      );
      if (!pedidoId) continue;
      if (!pedidoToIndexes.has(pedidoId)) continue;

      const isOkText = normalize_(row[IMPORT_CONFIG.COL_G - 1]).toLowerCase() === "ok";
      const isGreen = (bgG[i][0] || "").toLowerCase() === "#00ff00";
      const hasForbidden = containsForbidden_(rowTextParaBusca_(row));

      if (isOkText || isGreen || hasForbidden) proibidos.add(pedidoId);
    }
  }

  if (proibidos.size === 0) {
    Logger.log("Nenhum pedido da Conferencia entrou em critério de exclusão na Planilha B.");
    return;
  }

  const outL = colL.map((r) => [String(r[0] || "").trim()]);

  proibidos.forEach((pedido) => {
    for (const idx of (pedidoToIndexes.get(pedido) || [])) {
      outL[idx][0] = CONFIRMADO_SIM;
    }
  });

  shA.getRange(startRowA, IMPORT_CONFIG.DEST_COL_CONFIRMADO, numRowsA, 1).setValues(outL);
  Logger.log(`Marcados ${proibidos.size} pedidos com "${CONFIRMADO_SIM}" na coluna Confirmado (L).`);
}

/******************* B -> A: SOMENTE STATUS VAZIO *******************/
function refletirStatusDaPlanilhaBNaConferencia_SomenteVazios(sheetNames) {
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
    if (normalize_(dataAF[i][5])) continue; // só vazios

    const key = makeKey_(dataAF[i][0], dataAF[i][1], dataAF[i][3], dataAF[i][2], dataAF[i][4]);
    if (!pendingKeyToIdxs.has(key)) pendingKeyToIdxs.set(key, []);
    pendingKeyToIdxs.get(key).push(i);
  }

  if (pendingKeyToIdxs.size === 0) {
    Logger.log("B->A (vazios): não há Status vazio na Conferencia.");
    return;
  }

  const ssB = SpreadsheetApp.openById(IMPORT_CONFIG.SOURCE_SPREADSHEET_ID);
  const nomes = sheetNames || resolveSourceSheetNames_(ssB).nomes;

  const statusByKey = new Map();
  let remaining = pendingKeyToIdxs.size; // early-stop

  for (const sheetName of nomes) {
    if (remaining <= 0) break;

    const shB = ssB.getSheetByName(sheetName);
    if (!shB) continue;

    const lastRowB = shB.getLastRow();
    if (lastRowB < IMPORT_CONFIG.DATA_START_ROW) continue;

    const numRowsB = lastRowB - IMPORT_CONFIG.DATA_START_ROW + 1;
    const maxCol = Math.max(IMPORT_CONFIG.COL_F, IMPORT_CONFIG.SOURCE_WRITEBACK_COL_H); // até H
    const valuesB = shB.getRange(IMPORT_CONFIG.DATA_START_ROW, 1, numRowsB, maxCol).getValues();

    for (let i = 0; i < valuesB.length; i++) {
      if (remaining <= 0) break;

      const row = valuesB[i];
      const s = normalize_(row[IMPORT_CONFIG.SOURCE_WRITEBACK_COL_H - 1]).toUpperCase();

      // Só aceita TENHO ou FALTA (inclui "FALTA - ...")
      if (!(s === "TENHO" || s.indexOf("FALTA") === 0)) continue;

      const market = row[IMPORT_CONFIG.COL_C - 1];
      const pedido = normalizePedidoForExport_(row[IMPORT_CONFIG.COL_B - 1], market);
      if (!normalize_(pedido)) continue;

      const key = makeKey_(
        pedido,
        market,
        row[IMPORT_CONFIG.COL_F - 1],
        row[IMPORT_CONFIG.COL_D - 1],
        row[IMPORT_CONFIG.COL_E - 1]
      );

      if (!pendingKeyToIdxs.has(key)) continue;
      if (!statusByKey.has(key)) remaining--;

      statusByKey.set(key, s); // último prevalece conforme ordem das abas
    }
  }

  if (statusByKey.size === 0) {
    Logger.log("B->A (vazios): nenhum TENHO/FALTA encontrado na B para os pendentes.");
    return;
  }

  const outF = dataAF.map((r) => [r[5]]);
  let updated = 0;

  for (const [key, newStatus] of statusByKey.entries()) {
    for (const idx of (pendingKeyToIdxs.get(key) || [])) {
      if (normalize_(outF[idx][0])) continue; // segurança
      outF[idx][0] = newStatus;
      updated++;
    }
  }

  if (updated > 0) {
    shA.getRange(startRowA, IMPORT_CONFIG.DEST_COL_STATUS_ITEM, numRowsA, 1).setValues(outF);
    Logger.log(`B->A (vazios): preenchidas ${updated} células na Conferencia!F.`);
  } else {
    Logger.log("B->A (vazios): nada para atualizar.");
  }
}

/******************* A -> B: reflete Status (F) na coluna H da Planilha B *******************/
function refletirStatusDaConferenciaNaPlanilhaB_(sheetNames) {
  const ssA = SpreadsheetApp.getActiveSpreadsheet();
  const shA = ssA.getSheetByName(IMPORT_CONFIG.DEST_SHEET_NAME);
  if (!shA) throw new Error(`Aba destino não encontrada: ${IMPORT_CONFIG.DEST_SHEET_NAME}`);

  const lastRowA = shA.getLastRow();
  if (lastRowA < 2) return;

  const startRowA = 2;
  const numRowsA = lastRowA - startRowA + 1;

  const dataAF = shA.getRange(startRowA, 1, numRowsA, 6).getValues(); // A..F

  const statusByKey = new Map();
  for (const r of dataAF) {
    const s = normalize_(r[5]); // F
    if (!s) continue;
    statusByKey.set(makeKey_(r[0], r[1], r[3], r[2], r[4]), s);
  }

  if (statusByKey.size === 0) {
    Logger.log("A->B: nenhum status preenchido na Conferencia (coluna F).");
    return;
  }

  const ssB = SpreadsheetApp.openById(IMPORT_CONFIG.SOURCE_SPREADSHEET_ID);
  const nomes = sheetNames || resolveSourceSheetNames_(ssB).nomes;

  for (const sheetName of nomes) {
    const shB = ssB.getSheetByName(sheetName);
    if (!shB) continue;

    const lastRowB = shB.getLastRow();
    if (lastRowB < IMPORT_CONFIG.DATA_START_ROW) continue;

    const numRowsB = lastRowB - IMPORT_CONFIG.DATA_START_ROW + 1;
    const maxColToRead = Math.max(IMPORT_CONFIG.COL_F, IMPORT_CONFIG.SOURCE_WRITEBACK_COL_H);
    const valuesB = shB.getRange(IMPORT_CONFIG.DATA_START_ROW, 1, numRowsB, maxColToRead).getValues();

    const currentH = shB
      .getRange(IMPORT_CONFIG.DATA_START_ROW, IMPORT_CONFIG.SOURCE_WRITEBACK_COL_H, numRowsB, 1)
      .getValues();

    const outH = currentH.map((r) => [r[0]]);
    let updated = 0;

    for (let i = 0; i < valuesB.length; i++) {
      const row = valuesB[i];
      const market = row[IMPORT_CONFIG.COL_C - 1];

      const key = makeKey_(
        normalizePedidoForExport_(row[IMPORT_CONFIG.COL_B - 1], market),
        market,
        row[IMPORT_CONFIG.COL_F - 1],
        row[IMPORT_CONFIG.COL_D - 1],
        row[IMPORT_CONFIG.COL_E - 1]
      );

      if (!statusByKey.has(key)) continue;

      const newStatus = statusByKey.get(key);
      if (normalize_(newStatus) !== normalize_(outH[i][0])) {
        outH[i][0] = newStatus;
        updated++;
      }
    }

    if (updated > 0) {
      shB
        .getRange(IMPORT_CONFIG.DATA_START_ROW, IMPORT_CONFIG.SOURCE_WRITEBACK_COL_H, numRowsB, 1)
        .setValues(outH);
      Logger.log(`A->B (${sheetName}): atualizadas ${updated} células na coluna H.`);
    } else {
      Logger.log(`A->B (${sheetName}): nada para atualizar.`);
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

/** Corta o pedido conforme o marketplace. */
function normalizePedidoForExport_(pedido, marketplace) {
  const p = normalize_(pedido);
  const m = normalize_(marketplace);

  const maxLen = Number(IMPORT_CONFIG.SHORT_PEDIDO_MAXLEN || 13);

  // 13 PRIMEIROS caracteres
  if ((IMPORT_CONFIG.SHORT_PEDIDO_FIRST_MARKETS || []).indexOf(m) >= 0) {
    return p.length > maxLen ? p.slice(0, maxLen) : p;
  }

  // 13 ÚLTIMOS caracteres
  if ((IMPORT_CONFIG.SHORT_PEDIDO_LAST_MARKETS || []).indexOf(m) >= 0) {
    return p.length > maxLen ? p.slice(-maxLen) : p;
  }

  return p;
}

/** Inclui marketplace na chave para evitar colisões entre sistemas. */
function makeKey_(pedido, marketplace, produto, nome, qt) {
  return [pedido, marketplace, produto, nome, qt]
    .map((v) => normalize_(v).toLowerCase())
    .join("||");
}

/** Junta a linha A..G num único texto normalizado (sem acento, minúsculo). */
function rowTextParaBusca_(row) {
  return row.map((v) => normalizarBusca_(v).toLowerCase()).join(" | ");
}

/**
 * Termos proibidos.
 * Em modo "word" (padrão) só casa a palavra inteira, então "full" não descarta
 * mais um produto chamado "Fullaroma" — mas ainda descarta "FULL" e
 * "Fulfillment", que é a intenção original do filtro.
 */
let _forbiddenRegex_ = undefined;

function forbiddenRegex_() {
  if (_forbiddenRegex_ !== undefined) return _forbiddenRegex_;

  const termos = (IMPORT_CONFIG.FORBIDDEN_TERMS || [])
    .map((t) => normalizarBusca_(t).toLowerCase())
    .filter(Boolean)
    .map(escapeRegex_);

  _forbiddenRegex_ = termos.length
    ? new RegExp("(^|[^a-z0-9])(" + termos.join("|") + ")([^a-z0-9]|$)")
    : null;

  return _forbiddenRegex_;
}

function escapeRegex_(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function containsForbidden_(rowTextLower) {
  const texto = String(rowTextLower || "");

  if (String(IMPORT_CONFIG.FORBIDDEN_MATCH || "word").toLowerCase() === "substring") {
    for (const t of IMPORT_CONFIG.FORBIDDEN_TERMS || []) {
      const termo = normalizarBusca_(t).toLowerCase();
      if (termo && texto.includes(termo)) return true;
    }
    return false;
  }

  const re = forbiddenRegex_();
  return re ? re.test(texto) : false;
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
    if (!normalize_(r[0])) continue;
    set.add(makeKey_(r[0], r[1], r[3], r[2], r[4]));
  }

  return set;
}

/**********************************************************************
 *  gerarListaFaltas — executa ao final do import
 *
 *  - "FALTA"                      => inclui todos os itens do pedido
 *  - "FALTA - termo1; termo2; ..." => inclui SOMENTE os itens que contêm
 *                                     algum dos termos
 *      • termo não encontrado: ignora e segue
 *      • nenhum termo bateu: não inclui nada e registra como não encontrado
 *
 *  A coluna E (Qtd) da Conferencia multiplica as somas (A:B e D:E).
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
  const confQtd = shConf.getRange(2, 5, numConfRows, 1).getValues();      // E
  const confStatus = shConf.getRange(2, 6, numConfRows, 1).getValues();   // F
  const confMsgId = shConf.getRange(2, 9, numConfRows, 1).getValues();    // I

  const faltasItens = []; // [{prod, terms:[], qty}]
  const todosItens = [];  // [{prod, id, qty}]

  for (let i = 0; i < numConfRows; i++) {
    const prod = String(confProdutos[i][0] ?? "").trim();
    if (!prod) continue;

    const qty = parseQtdConferencia_(confQtd[i][0]);
    const id = readId_(confMsgId[i][0]); // evita notação científica

    todosItens.push({ prod, id, qty });

    const stRaw = String(confStatus[i][0] ?? "").trim();
    if (stRaw.toUpperCase().includes("FALTA")) {
      faltasItens.push({ prod, terms: extrairFaltaTermos_(stRaw), qty });
    }
  }

  if (todosItens.length === 0) {
    shLista.getRange("C2").setValue("Nenhum produto encontrado em Conferencia.");
    return;
  }

  // ===== 2) Varre Dados A:Z montando mapa Produto -> valor à direita =====
  const lastDadosRow = shDados.getLastRow();
  const lastDadosCol = Math.min(shDados.getLastColumn(), 26); // até Z

  if (lastDadosRow < 1 || lastDadosCol < 2) {
    shLista.getRange("C2").setValue('Aba "Dados" vazia ou com poucas colunas.');
    return;
  }

  const dadosVals = shDados.getRange(1, 1, lastDadosRow, lastDadosCol).getValues();

  const wanted = new Set(
    [...faltasItens.map((x) => x.prod), ...todosItens.map((x) => x.prod)]
      .map((p) => normalizarBusca_(p))
  );

  const mapSubseq = new Map(); // produtoNormalizado -> valor da célula à direita

  for (let r = 0; r < dadosVals.length; r++) {
    const row = dadosVals[r];
    for (let c = 0; c < row.length - 1; c++) {
      const cell = String(row[c] ?? "").trim();
      if (!cell) continue;

      const key = normalizarBusca_(cell);
      if (!wanted.has(key)) continue;
      if (mapSubseq.has(key)) continue;

      mapSubseq.set(key, String(row[c + 1] ?? "").trim());
    }
  }

  // ===== 3) Processadores =====
  const notFoundAll = []; // Lista!C (união dos não encontrados)

  // 3.1) A:B (somente FALTA) — com filtro por termos quando houver "FALTA - ..."
  function processarFaltas_(listaItens) {
    const counts = new Map();

    for (const it of listaItens) {
      const prod = it.prod;
      const terms = Array.isArray(it.terms) ? it.terms : [];
      const mult = Number(it.qty || 1);

      const subseq = mapSubseq.get(normalizarBusca_(prod));
      if (!subseq) {
        notFoundAll.push(prod);
        continue;
      }

      const partsAll = subseq.split(";").map((s) => s.trim()).filter(Boolean);
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
        counts.set(p, (counts.get(p) || 0) + mult);
      }
    }

    return Array.from(counts.entries())
      .sort((a, b) => a[0].localeCompare(b[0], "pt-BR"))
      .map(([name, qty]) => [name, qty]);
  }

  // 3.2) D:E:F (Lista completa) — com IDs por item
  // IDs ficam únicos (não repetem pela Qtd); a quantidade (col E) soma com Qtd.
  function processarTodosComIds_(itens) {
    const counts = new Map();
    const idsMap = new Map();

    for (const item of itens) {
      const prod = item.prod;
      const id = item.id;
      const mult = Number(item.qty || 1);

      const subseq = mapSubseq.get(normalizarBusca_(prod));
      if (!subseq) {
        notFoundAll.push(prod);
        continue;
      }

      const parts = subseq.split(";").map((s) => s.trim()).filter(Boolean);
      if (parts.length === 0) {
        notFoundAll.push(prod);
        continue;
      }

      for (const p of parts) {
        counts.set(p, (counts.get(p) || 0) + mult);

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

    return Array.from(counts.keys())
      .sort((a, b) => a.localeCompare(b, "pt-BR"))
      .map((name) => {
        const bucket = idsMap.get(name);
        return [
          name,
          counts.get(name) || 0,
          bucket && bucket.list.length ? bucket.list.join("\n") : "",
        ];
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
    // Coluna F guarda IDs do Discord: precisa ser texto.
    shLista.getRange(2, 6, foundRowsTodos.length, 1).setNumberFormat("@");
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
 * Extrai termos após "FALTA -", aceitando múltiplos separados por ";".
 *  "FALTA - Dark Tabc; Lavanda" -> ["Dark Tabc", "Lavanda"]
 *  "FALTA"                      -> []
 */
function extrairFaltaTermos_(statusRaw) {
  const s = String(statusRaw || "").trim();
  const up = s.toUpperCase();

  if (up.indexOf("FALTA") < 0) return [];

  const pos = up.indexOf("FALTA -");
  if (pos < 0) return []; // "FALTA" sem complemento

  const after = s.slice(pos + "FALTA -".length).trim();
  if (!after) return [];

  return after.split(";").map((x) => String(x || "").trim()).filter(Boolean);
}

/** Normalização para busca: sem acento, maiúsculo, espaços colapsados. */
function normalizarBusca_(s) {
  return String(s ?? "")
    .trim()
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ");
}

/**
 * Interpreta Qtd da Conferencia!E
 * - Aceita número ou texto ("2", "2,0", ...)
 * - Vazio/inválido vira 1
 * - Decimais: piso (2.7 -> 2)
 */
function parseQtdConferencia_(v) {
  if (v === null || v === undefined || v === "") return 1;

  const n = Number(String(v).trim().replace(",", "."));
  if (!isFinite(n) || n <= 0) return 1;

  return Math.floor(n);
}
