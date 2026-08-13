/* ============================================================================
   STÄRKER CONSULTING — Motor de formatação de tabelas (Office.js)  v2
   Aplica o padrão visual STÄRKER a Tabelas Dinâmicas, Tabelas e Intervalos.
   Compatível com Excel para Windows, Mac e Web (Microsoft 365).
   ----------------------------------------------------------------------------
   Para ajustar as cores da identidade, altere APENAS o objeto PALETA abaixo.
   ========================================================================== */

const PALETA = {
  offwhite:   "#F7F6F2", // Fundo base / 1ª faixa de linha
  warmgray:   "#ECEBE7", // 2ª faixa de linha (cinza quente)
  graphite:   "#262626", // Cor de fonte padrão
  navy:       "#0B202D", // Cabeçalhos, sidebar de rótulos e Total Geral
  cobreClaro: "#D9B184", // Subtítulos / Subtotais
  cobre:      "#C6793C", // Cobre de destaque (bordas)
  white:      "#FFFFFF",
  gridline:   "#D8D5CE"  // Linhas internas discretas (só nas áreas claras)
};

/* --------------------------------------------------------------------------
   Handlers dos botões da faixa
   -------------------------------------------------------------------------- */

async function aplicarSelecao(event) {
  try {
    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      const sel = ctx.workbook.getSelectedRange();
      sel.load("address");

      const pivots = sheet.pivotTables;
      pivots.load("items/name");
      const tables = sheet.tables;
      tables.load("items/name");
      await ctx.sync();

      // 1) Seleção dentro de uma Tabela Dinâmica?
      let alvoPivot = null;
      for (const p of pivots.items) {
        const inter = sel.getIntersectionOrNullObject(p.layout.getRange());
        inter.load("isNullObject");
        p.__inter = inter;
      }
      await ctx.sync();
      for (const p of pivots.items) if (!p.__inter.isNullObject) { alvoPivot = p; break; }
      if (alvoPivot) {
        await formatarPivot(ctx, alvoPivot);
        await ctx.sync();
        notificar("Padrão STÄRKER aplicado à tabela dinâmica.");
        return;
      }

      // 2) Seleção dentro de uma Tabela?
      let alvoTabela = null;
      for (const t of tables.items) {
        const inter = sel.getIntersectionOrNullObject(t.getRange());
        inter.load("isNullObject");
        t.__inter = inter;
      }
      await ctx.sync();
      for (const t of tables.items) if (!t.__inter.isNullObject) { alvoTabela = t; break; }
      if (alvoTabela) {
        await formatarTabela(ctx, alvoTabela);
        await ctx.sync();
        notificar("Padrão STÄRKER aplicado à tabela.");
        return;
      }

      // 3) Intervalo comum
      await formatarIntervalo(ctx, sel);
      await ctx.sync();
      notificar("Padrão STÄRKER aplicado ao intervalo selecionado.");
    });
  } catch (e) {
    console.error(e);
    notificar("Erro ao aplicar o padrão: " + (e.message || e), true);
  } finally {
    if (event && event.completed) event.completed();
  }
}

async function aplicarPlanilha(event) {
  try {
    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      const pivots = sheet.pivotTables;
      pivots.load("items/name");
      const tables = sheet.tables;
      tables.load("items/name");
      await ctx.sync();

      for (const p of pivots.items) { await formatarPivot(ctx, p); await ctx.sync(); }
      for (const t of tables.items) { await formatarTabela(ctx, t); await ctx.sync(); }

      notificar("Padrão aplicado a " + (pivots.items.length + tables.items.length) + " tabela(s).");
    });
  } catch (e) {
    console.error(e);
    notificar("Erro ao aplicar o padrão: " + (e.message || e), true);
  } finally {
    if (event && event.completed) event.completed();
  }
}

/* --------------------------------------------------------------------------
   TABELA DINÂMICA
   -------------------------------------------------------------------------- */
async function formatarPivot(ctx, pivot) {
  const layout = pivot.layout;
  const full = layout.getRange();
  full.load("rowCount, columnCount, values");
  const colLabels = layout.getColumnLabelRange();
  colLabels.load("rowCount");
  const rowLabels = layout.getRowLabelRange();
  rowLabels.load("columnCount");
  await ctx.sync();

  const nRows = full.rowCount;
  const nCols = full.columnCount;
  const headerRows = Math.max(1, colLabels.rowCount || 1);
  const labelCols  = Math.max(1, rowLabels.columnCount || 1);
  const vals = full.values;

  // Base clara
  full.format.fill.color = PALETA.offwhite;
  full.format.font.color = PALETA.graphite;

  // ---- Cabeçalho (inclui o canto superior esquerdo) ----
  const header = full.getCell(0, 0).getResizedRange(headerRows - 1, nCols - 1);
  header.format.fill.color = PALETA.navy;
  header.format.font.color = PALETA.white;
  header.format.font.bold = true;
  fundirBordas(header, PALETA.navy);
  bordaAresta(header, "EdgeBottom", PALETA.cobre, "Medium");

  // Detecta coluna de Total Geral (por texto no cabeçalho)
  let colTotalGeral = -1;
  for (let c = labelCols; c < nCols; c++) {
    const txt = String(vals[headerRows - 1][c] || "").toLowerCase();
    if (txt.includes("total geral") || txt.includes("grand total")) colTotalGeral = c;
  }

  // ---- Linhas ----
  let dataIdx = 0;
  for (let r = headerRows; r < nRows; r++) {
    let rotulo = "";
    for (let c = 0; c < labelCols; c++) rotulo += " " + String(vals[r][c] || "");
    rotulo = rotulo.toLowerCase();

    const rowFull = full.getCell(r, 0).getResizedRange(0, nCols - 1);
    const isGrand = rotulo.includes("total geral") || rotulo.includes("grand total");
    const isSub   = !isGrand && rotulo.includes("total");

    if (isGrand) {
      rowFull.format.fill.color = PALETA.navy;
      rowFull.format.font.color = PALETA.white;
      rowFull.format.font.bold = true;
      fundirBordas(rowFull, PALETA.navy);
      bordaAresta(rowFull, "EdgeTop", PALETA.cobre, "Medium");
    } else if (isSub) {
      rowFull.format.fill.color = PALETA.cobreClaro;
      rowFull.format.font.color = PALETA.navy;
      rowFull.format.font.bold = true;
      fundirBordas(rowFull, PALETA.cobreClaro);
      bordaAresta(rowFull, "EdgeTop", PALETA.cobre, "Thin");
    } else {
      // Sidebar escuro: colunas de rótulo (Tipo/Subtipo/Grupo)
      const labelCells = full.getCell(r, 0).getResizedRange(0, labelCols - 1);
      labelCells.format.fill.color = PALETA.navy;
      labelCells.format.font.color = PALETA.white;
      fundirBordas(labelCells, PALETA.navy);

      // Área de dados: faixas claras + grade discreta
      const dataCells = full.getCell(r, labelCols).getResizedRange(0, nCols - 1 - labelCols);
      dataCells.format.fill.color = (dataIdx % 2 === 0) ? PALETA.offwhite : PALETA.warmgray;
      dataCells.format.font.color = PALETA.graphite;
      gradeClara(dataCells);
      dataIdx++;
    }
  }

  // Divisor cobre entre o sidebar escuro e a área de dados
  const sidebar = full.getCell(headerRows, 0).getResizedRange(nRows - 1 - headerRows, labelCols - 1);
  bordaAresta(sidebar, "EdgeRight", PALETA.cobre, "Thin");

  // Coluna de Total Geral (se existir) — azul-marinho
  if (colTotalGeral >= 0) {
    const col = full.getCell(headerRows, colTotalGeral).getResizedRange(nRows - headerRows - 1, 0);
    col.format.fill.color = PALETA.navy;
    col.format.font.color = PALETA.white;
    col.format.font.bold = true;
    fundirBordas(col, PALETA.navy);
  }

  // Borda externa cobre
  bordaExterna(full, PALETA.cobre, "Thin");
}

/* --------------------------------------------------------------------------
   TABELA
   -------------------------------------------------------------------------- */
async function formatarTabela(ctx, table) {
  table.style = "TableStyleLight1";
  table.showBandedRows = false;
  table.load("showTotals");

  const headerRange = table.getHeaderRowRange();
  const bodyRange = table.getDataBodyRange();
  bodyRange.load("rowCount, columnCount");
  await ctx.sync();

  headerRange.format.fill.color = PALETA.navy;
  headerRange.format.font.color = PALETA.white;
  headerRange.format.font.bold = true;
  fundirBordas(headerRange, PALETA.navy);
  bordaAresta(headerRange, "EdgeBottom", PALETA.cobre, "Medium");

  const nr = bodyRange.rowCount;
  for (let r = 0; r < nr; r++) {
    const linha = bodyRange.getRow(r);
    linha.format.fill.color = (r % 2 === 0) ? PALETA.offwhite : PALETA.warmgray;
    linha.format.font.color = PALETA.graphite;
  }
  gradeClara(bodyRange);

  if (table.showTotals) {
    const totalRange = table.getTotalRowRange();
    totalRange.format.fill.color = PALETA.navy;
    totalRange.format.font.color = PALETA.white;
    totalRange.format.font.bold = true;
    fundirBordas(totalRange, PALETA.navy);
    bordaAresta(totalRange, "EdgeTop", PALETA.cobre, "Medium");
  }

  bordaExterna(table.getRange(), PALETA.cobre, "Thin");
}

/* --------------------------------------------------------------------------
   INTERVALO comum (1ª linha = cabeçalho)
   -------------------------------------------------------------------------- */
async function formatarIntervalo(ctx, range) {
  range.load("rowCount, columnCount");
  await ctx.sync();
  const nr = range.rowCount;

  range.format.fill.color = PALETA.offwhite;
  range.format.font.color = PALETA.graphite;

  const header = range.getRow(0);
  header.format.fill.color = PALETA.navy;
  header.format.font.color = PALETA.white;
  header.format.font.bold = true;
  fundirBordas(header, PALETA.navy);
  bordaAresta(header, "EdgeBottom", PALETA.cobre, "Medium");

  const body = range.getCell(1, 0).getResizedRange(nr - 2, range.columnCount - 1);
  for (let r = 1; r < nr; r++) {
    const linha = range.getRow(r);
    linha.format.fill.color = ((r - 1) % 2 === 0) ? PALETA.offwhite : PALETA.warmgray;
    linha.format.font.color = PALETA.graphite;
  }
  if (nr > 1) gradeClara(body);
  bordaExterna(range, PALETA.cobre, "Thin");
}

/* --------------------------------------------------------------------------
   Bordas
   -------------------------------------------------------------------------- */
// Funde TODAS as bordas (internas e arestas) com a cor do fundo -> some no escuro
function fundirBordas(range, cor) {
  const b = range.format.borders;
  ["InsideHorizontal", "InsideVertical", "EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"]
    .forEach((e) => { const it = b.getItem(e); it.color = cor; it.weight = "Thin"; it.style = "Continuous"; });
}
// Grade discreta só nas áreas claras
function gradeClara(range) {
  const b = range.format.borders;
  ["InsideHorizontal", "InsideVertical"]
    .forEach((e) => { const it = b.getItem(e); it.color = PALETA.gridline; it.weight = "Thin"; it.style = "Continuous"; });
}
function bordaAresta(range, aresta, cor, peso) {
  const it = range.format.borders.getItem(aresta);
  it.color = cor; it.weight = peso; it.style = "Continuous";
}
function bordaExterna(range, cor, peso) {
  ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"].forEach((e) => bordaAresta(range, e, cor, peso));
}

/* --------------------------------------------------------------------------
   Status
   -------------------------------------------------------------------------- */
function notificar(msg, erro) {
  console.log(msg);
  if (window.__starkerStatus) window.__starkerStatus(msg, erro);
}

Office.onReady(() => {
  if (Office.actions && Office.actions.associate) {
    Office.actions.associate("aplicarSelecao", aplicarSelecao);
    Office.actions.associate("aplicarPlanilha", aplicarPlanilha);
  }
});

if (typeof window !== "undefined") {
  window.STARKER = { aplicarSelecao, aplicarPlanilha, PALETA };
}
