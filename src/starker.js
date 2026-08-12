/* ============================================================================
   STÄRKER CONSULTING — Motor de formatação de tabelas (Office.js)
   Aplica o padrão visual STÄRKER a Tabelas Dinâmicas, Tabelas e Intervalos.
   Compatível com Excel para Windows, Mac e Web (Microsoft 365).
   ----------------------------------------------------------------------------
   Para ajustar as cores da identidade, altere APENAS o objeto PALETA abaixo.
   ========================================================================== */

const PALETA = {
  offwhite:   "#F7F6F2", // Fundo base / 1ª faixa de linha
  warmgray:   "#ECEBE7", // 2ª faixa de linha (cinza quente)
  graphite:   "#262626", // Cor de fonte padrão
  navy:       "#0B202D", // Cabeçalhos e Total Geral
  cobreClaro: "#D9B184", // Subtítulos / Subtotais
  cobre:      "#C6793C", // Cobre de destaque (bordas)
  white:      "#FFFFFF",
  gridline:   "#D8D5CE"  // Linhas internas discretas
};

/* --------------------------------------------------------------------------
   Handlers dos botões da faixa (associados no final do arquivo)
   -------------------------------------------------------------------------- */

async function aplicarSelecao(event) {
  try {
    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      const sel = ctx.workbook.getSelectedRange();
      sel.load("address, rowIndex, columnIndex, rowCount, columnCount");

      const pivots = sheet.pivotTables;
      pivots.load("items/name");
      const tables = sheet.tables;
      tables.load("items/name");
      await ctx.sync();

      // 1) A seleção intersecta alguma Tabela Dinâmica?
      let alvoPivot = null;
      for (const p of pivots.items) {
        const pr = p.layout.getRange();
        const inter = sel.getIntersectionOrNullObject(pr);
        inter.load("isNullObject");
        p.__pr = pr;
        p.__inter = inter;
      }
      await ctx.sync();
      for (const p of pivots.items) {
        if (!p.__inter.isNullObject) { alvoPivot = p; break; }
      }
      if (alvoPivot) {
        await formatarPivot(ctx, alvoPivot);
        await ctx.sync();
        notificar("Padrão STÄRKER aplicado à tabela dinâmica.");
        return;
      }

      // 2) A seleção intersecta alguma Tabela?
      let alvoTabela = null;
      for (const t of tables.items) {
        const tr = t.getRange();
        const inter = sel.getIntersectionOrNullObject(tr);
        inter.load("isNullObject");
        t.__inter = inter;
      }
      await ctx.sync();
      for (const t of tables.items) {
        if (!t.__inter.isNullObject) { alvoTabela = t; break; }
      }
      if (alvoTabela) {
        await formatarTabela(ctx, alvoTabela);
        await ctx.sync();
        notificar("Padrão STÄRKER aplicado à tabela.");
        return;
      }

      // 3) Caso contrário: formata o intervalo selecionado (1ª linha = cabeçalho)
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

      const total = pivots.items.length + tables.items.length;
      notificar("Padrão aplicado a " + total + " tabela(s) na planilha ativa.");
    });
  } catch (e) {
    console.error(e);
    notificar("Erro ao aplicar o padrão: " + (e.message || e), true);
  } finally {
    if (event && event.completed) event.completed();
  }
}

/* --------------------------------------------------------------------------
   Formatação de TABELA DINÂMICA
   -------------------------------------------------------------------------- */
async function formatarPivot(ctx, pivot) {
  const layout = pivot.layout;
  const full = layout.getRange();
  full.load("rowCount, columnCount, values, address");

  const colLabels = layout.getColumnLabelRange();
  colLabels.load("rowCount, columnCount");
  const rowLabels = layout.getRowLabelRange();
  rowLabels.load("columnCount");
  await ctx.sync();

  const nRows = full.rowCount;
  const nCols = full.columnCount;
  const headerRows = Math.max(1, colLabels.rowCount || 1);
  const labelCols  = Math.max(1, rowLabels.columnCount || 1);
  const vals = full.values;

  // Base: fundo off-white + fonte grafite + linhas internas discretas
  full.format.fill.color = PALETA.offwhite;
  full.format.font.color = PALETA.graphite;
  aplicarGrade(full);

  // Bloco de cabeçalho (inclui o canto superior esquerdo dos campos)
  const header = full.getResizedRange(-(nRows - headerRows), 0); // primeiras headerRows linhas
  header.format.fill.color = PALETA.navy;
  header.format.font.color = PALETA.white;
  header.format.font.bold = true;
  header.getRow(headerRows - 1).format.borders.getItem("EdgeBottom").color = PALETA.cobre;
  header.getRow(headerRows - 1).format.borders.getItem("EdgeBottom").weight = "Medium";

  // Detecta coluna de Total Geral (por texto no cabeçalho)
  let colTotalGeral = -1;
  for (let c = labelCols; c < nCols; c++) {
    const txt = String(vals[headerRows - 1][c] || "").toLowerCase();
    if (txt.includes("total geral") || txt.includes("grand total")) colTotalGeral = c;
  }

  // Linhas de dados
  let dataIdx = 0;
  for (let r = headerRows; r < nRows; r++) {
    let rotulo = "";
    for (let c = 0; c < labelCols; c++) rotulo += " " + String(vals[r][c] || "");
    rotulo = rotulo.toLowerCase();

    const linha = full.getCell(r, 0).getResizedRange(0, nCols - 1); // linha inteira do pivot
    const isGrand = rotulo.includes("total geral") || rotulo.includes("grand total");
    const isSub   = !isGrand && rotulo.includes("total");

    if (isGrand) {
      linha.format.fill.color = PALETA.navy;
      linha.format.font.color = PALETA.white;
      linha.format.font.bold = true;
      linha.format.borders.getItem("EdgeTop").color = PALETA.cobre;
      linha.format.borders.getItem("EdgeTop").weight = "Medium";
    } else if (isSub) {
      linha.format.fill.color = PALETA.cobreClaro;
      linha.format.font.color = PALETA.navy;
      linha.format.font.bold = true;
      linha.format.borders.getItem("EdgeTop").color = PALETA.cobre;
    } else {
      linha.format.fill.color = (dataIdx % 2 === 0) ? PALETA.offwhite : PALETA.warmgray;
      linha.format.font.color = PALETA.graphite;
      linha.format.font.bold = false;
      dataIdx++;
    }
  }

  // Coluna de Total Geral (se existir) — azul-marinho
  if (colTotalGeral >= 0) {
    const col = full.getCell(headerRows, colTotalGeral)
                    .getResizedRange(nRows - headerRows - 1, 0);
    col.format.fill.color = PALETA.navy;
    col.format.font.color = PALETA.white;
    col.format.font.bold = true;
  }

  // Borda externa cobre
  bordaExterna(full, PALETA.cobre, "Thin");
}

/* --------------------------------------------------------------------------
   Formatação de TABELA (objeto Tabela do Excel)
   -------------------------------------------------------------------------- */
async function formatarTabela(ctx, table) {
  // Neutraliza o estilo nativo para termos controle total das cores
  table.style = "TableStyleLight1";
  table.showBandedRows = false;
  table.load("showTotals");

  const headerRange = table.getHeaderRowRange();
  const bodyRange = table.getDataBodyRange();
  bodyRange.load("rowCount, columnCount");
  await ctx.sync();

  // Cabeçalho
  headerRange.format.fill.color = PALETA.navy;
  headerRange.format.font.color = PALETA.white;
  headerRange.format.font.bold = true;
  headerRange.format.borders.getItem("EdgeBottom").color = PALETA.cobre;
  headerRange.format.borders.getItem("EdgeBottom").weight = "Medium";

  // Corpo com faixas
  const nr = bodyRange.rowCount;
  for (let r = 0; r < nr; r++) {
    const linha = bodyRange.getRow(r);
    linha.format.fill.color = (r % 2 === 0) ? PALETA.offwhite : PALETA.warmgray;
    linha.format.font.color = PALETA.graphite;
  }
  aplicarGrade(bodyRange);

  // Linha de totais (se ligada)
  if (table.showTotals) {
    const totalRange = table.getTotalRowRange();
    totalRange.format.fill.color = PALETA.navy;
    totalRange.format.font.color = PALETA.white;
    totalRange.format.font.bold = true;
    totalRange.format.borders.getItem("EdgeTop").color = PALETA.cobre;
    totalRange.format.borders.getItem("EdgeTop").weight = "Medium";
  }

  bordaExterna(table.getRange(), PALETA.cobre, "Thin");
}

/* --------------------------------------------------------------------------
   Formatação de INTERVALO comum (1ª linha tratada como cabeçalho)
   -------------------------------------------------------------------------- */
async function formatarIntervalo(ctx, range) {
  range.load("rowCount, columnCount");
  await ctx.sync();
  const nr = range.rowCount;

  range.format.fill.color = PALETA.offwhite;
  range.format.font.color = PALETA.graphite;
  aplicarGrade(range);

  const header = range.getRow(0);
  header.format.fill.color = PALETA.navy;
  header.format.font.color = PALETA.white;
  header.format.font.bold = true;
  header.format.borders.getItem("EdgeBottom").color = PALETA.cobre;
  header.format.borders.getItem("EdgeBottom").weight = "Medium";

  for (let r = 1; r < nr; r++) {
    const linha = range.getRow(r);
    linha.format.fill.color = ((r - 1) % 2 === 0) ? PALETA.offwhite : PALETA.warmgray;
    linha.format.font.color = PALETA.graphite;
  }
  bordaExterna(range, PALETA.cobre, "Thin");
}

/* --------------------------------------------------------------------------
   Utilitários de borda
   -------------------------------------------------------------------------- */
function aplicarGrade(range) {
  const b = range.format.borders;
  ["InsideHorizontal", "InsideVertical"].forEach((edge) => {
    b.getItem(edge).color = PALETA.gridline;
    b.getItem(edge).weight = "Thin";
    b.getItem(edge).style = "Continuous";
  });
}

function bordaExterna(range, cor, peso) {
  const b = range.format.borders;
  ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"].forEach((edge) => {
    b.getItem(edge).color = cor;
    b.getItem(edge).weight = peso;
    b.getItem(edge).style = "Continuous";
  });
}

/* --------------------------------------------------------------------------
   Notificação leve (barra do Excel) — ignora silenciosamente se indisponível
   -------------------------------------------------------------------------- */
function notificar(msg, erro) {
  try {
    Office.context.mailbox; // no-op guard
  } catch (e) {}
  try {
    if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.message) {
      // reservado para futuras versões
    }
  } catch (e) {}
  console.log(msg);
  if (window.__starkerStatus) window.__starkerStatus(msg, erro);
}

/* --------------------------------------------------------------------------
   Registro das ações (necessário para os botões ExecuteFunction)
   -------------------------------------------------------------------------- */
Office.onReady(() => {
  if (Office.actions && Office.actions.associate) {
    Office.actions.associate("aplicarSelecao", aplicarSelecao);
    Office.actions.associate("aplicarPlanilha", aplicarPlanilha);
  }
});

// Exporta para o painel de tarefas
if (typeof window !== "undefined") {
  window.STARKER = { aplicarSelecao, aplicarPlanilha, PALETA };
}
