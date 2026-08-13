/* ============================================================================
   STÄRKER CONSULTING — Motor de formatação de tabelas (Office.js)  v4
   Aplica o padrão visual STÄRKER a Tabelas Dinâmicas, Tabelas e Intervalos.
   Compatível com Excel para Windows, Mac e Web (Microsoft 365).
   ----------------------------------------------------------------------------
   • Colunas de rótulo: escala escura harmônica; grade SÓ horizontal (sem
     verticais que quebrem a cor).
   • Subtotais: tom diferente por nível hierárquico.
   • Dinâmica: preserva a formatação ao expandir/recolher/atualizar.
   ========================================================================== */

const PALETA = {
  offwhite:    "#F7F6F2", // Fundo base / 1ª faixa
  warmgray:    "#ECEBE7", // 2ª faixa
  graphite:    "#262626", // Fonte padrão
  navy:        "#0B202D", // Cabeçalho das colunas de dados e Total Geral
  slate:       "#2B4F63", // Fim da escala escura das colunas de rótulo
  linhaTitulo: "#48627A", // Grade clara (barra de títulos + horizontais do sidebar)
  subtotalTop: "#D9B184", // Subtotal de nível mais alto (mais forte)
  subtotalBot: "#F1E7D9", // Subtotal de nível mais baixo (mais claro)
  cobre:       "#C6793C", // Cobre de destaque (bordas)
  white:       "#FFFFFF",
  gridline:    "#D8D5CE"  // Grade discreta nas áreas claras
};

/* ---- Escalas de cor (interpolação linear entre dois hex) ---- */
function hexToRgb(h){ h=h.replace("#",""); return [0,2,4].map(i=>parseInt(h.substr(i,2),16)); }
function rgbToHex(r){ return "#"+r.map(x=>Math.max(0,Math.min(255,Math.round(x))).toString(16).padStart(2,"0")).join(""); }
function escala(c1, c2, n){
  const a=hexToRgb(c1), b=hexToRgb(c2), out=[];
  for(let i=0;i<n;i++){ const t=n<=1?0:i/(n-1);
    out.push(rgbToHex([a[0]+(b[0]-a[0])*t, a[1]+(b[1]-a[1])*t, a[2]+(b[2]-a[2])*t])); }
  return out;
}

/* --------------------------------------------------------------------------
   Handlers
   -------------------------------------------------------------------------- */
async function aplicarSelecao(event){
  try{
    await Excel.run(async (ctx)=>{
      const sheet=ctx.workbook.worksheets.getActiveWorksheet();
      const sel=ctx.workbook.getSelectedRange(); sel.load("address");
      const pivots=sheet.pivotTables; pivots.load("items/name");
      const tables=sheet.tables; tables.load("items/name");
      await ctx.sync();

      let alvoPivot=null;
      for(const p of pivots.items){ const i=sel.getIntersectionOrNullObject(p.layout.getRange()); i.load("isNullObject"); p.__i=i; }
      await ctx.sync();
      for(const p of pivots.items) if(!p.__i.isNullObject){ alvoPivot=p; break; }
      if(alvoPivot){ await formatarPivot(ctx,alvoPivot); await ctx.sync(); notificar("Padrão aplicado à tabela dinâmica."); return; }

      let alvoTabela=null;
      for(const t of tables.items){ const i=sel.getIntersectionOrNullObject(t.getRange()); i.load("isNullObject"); t.__i=i; }
      await ctx.sync();
      for(const t of tables.items) if(!t.__i.isNullObject){ alvoTabela=t; break; }
      if(alvoTabela){ await formatarTabela(ctx,alvoTabela); await ctx.sync(); notificar("Padrão aplicado à tabela."); return; }

      await formatarIntervalo(ctx,sel); await ctx.sync(); notificar("Padrão aplicado ao intervalo.");
    });
  }catch(e){ console.error(e); notificar("Erro: "+(e.message||e),true); }
  finally{ if(event&&event.completed) event.completed(); }
}

async function aplicarPlanilha(event){
  try{
    await Excel.run(async (ctx)=>{
      const sheet=ctx.workbook.worksheets.getActiveWorksheet();
      const pivots=sheet.pivotTables; pivots.load("items/name");
      const tables=sheet.tables; tables.load("items/name");
      await ctx.sync();
      for(const p of pivots.items){ await formatarPivot(ctx,p); await ctx.sync(); }
      for(const t of tables.items){ await formatarTabela(ctx,t); await ctx.sync(); }
      notificar("Padrão aplicado a "+(pivots.items.length+tables.items.length)+" tabela(s).");
    });
  }catch(e){ console.error(e); notificar("Erro: "+(e.message||e),true); }
  finally{ if(event&&event.completed) event.completed(); }
}

/* --------------------------------------------------------------------------
   TABELA DINÂMICA
   -------------------------------------------------------------------------- */
async function formatarPivot(ctx, pivot){
  const layout=pivot.layout;
  try{ layout.preserveFormatting=true; }catch(e){}  // mantém formatação ao expandir/recolher

  const full=layout.getRange();
  full.load("rowCount, columnCount, values");
  const colLabels=layout.getColumnLabelRange(); colLabels.load("rowCount");
  const rowLabels=layout.getRowLabelRange(); rowLabels.load("columnCount");
  await ctx.sync();

  const nRows=full.rowCount, nCols=full.columnCount;
  const headerRows=Math.max(1, colLabels.rowCount||1);
  const labelCols =Math.max(1, rowLabels.columnCount||1);
  const vals=full.values;
  const escalaRotulo=escala(PALETA.navy, PALETA.slate, labelCols);          // colunas de rótulo
  const escalaSub   =escala(PALETA.subtotalTop, PALETA.subtotalBot, labelCols); // subtotais por nível

  // 1) Base clara
  full.format.fill.color=PALETA.offwhite;
  full.format.font.color=PALETA.graphite;

  // 2) Colunas de rótulo em escala escura (coluna inteira)
  for(let c=0;c<labelCols;c++){
    const col=full.getCell(0,c).getResizedRange(nRows-1,0);
    col.format.fill.color=escalaRotulo[c];
    col.format.font.color=PALETA.white;
  }

  // 3) Cabeçalho das colunas de dados
  const hData=full.getCell(0,labelCols).getResizedRange(headerRows-1, nCols-1-labelCols);
  hData.format.fill.color=PALETA.navy;
  hData.format.font.color=PALETA.white;
  hData.format.font.bold=true;
  full.getCell(0,0).getResizedRange(headerRows-1, labelCols-1).format.font.bold=true; // rótulos do cabeçalho em negrito

  // 4) Canto do cabeçalho (rótulos) = sólido, sem linhas claras
  limparBordasInternas(full.getCell(0,0).getResizedRange(headerRows-1, labelCols-1));

  // 4b) Corpo do sidebar (colunas de rótulo): 100% sólido, SEM linhas claras
  const sidebarCorpo=full.getCell(headerRows,0).getResizedRange(nRows-1-headerRows, labelCols-1);
  limparBordasInternas(sidebarCorpo);

  // 5) Barra de títulos (colunas de dados): grade SÓ vertical + linha inferior cobre
  gradeTitulo(hData);
  bordaAresta(full.getCell(0,0).getResizedRange(headerRows-1,nCols-1), "EdgeBottom", PALETA.cobre, "Medium");

  // Detecta coluna de Total Geral
  let colTotalGeral=-1;
  for(let c=labelCols;c<nCols;c++){
    const txt=String(vals[headerRows-1][c]||"").toLowerCase();
    if(txt.includes("total geral")||txt.includes("grand total")) colTotalGeral=c;
  }

  // 6) Linhas
  let dataIdx=0;
  for(let r=headerRows;r<nRows;r++){
    let rotulo=""; let nivel=0, achouTotal=false;
    for(let c=0;c<labelCols;c++){
      const txt=String(vals[r][c]||"");
      rotulo+=" "+txt;
      if(!achouTotal && txt.toLowerCase().includes("total")){ nivel=c; achouTotal=true; }
    }
    rotulo=rotulo.toLowerCase();
    const rowFull=full.getCell(r,0).getResizedRange(0,nCols-1);
    const isGrand=rotulo.includes("total geral")||rotulo.includes("grand total");
    const isSub=!isGrand && rotulo.includes("total");

    if(isGrand){
      rowFull.format.fill.color=PALETA.navy;
      rowFull.format.font.color=PALETA.white;
      rowFull.format.font.bold=true;
      fundirBordas(rowFull, PALETA.navy);
      bordaAresta(rowFull,"EdgeTop",PALETA.cobre,"Medium");
    } else if(isSub){
      const tom=escalaSub[Math.min(nivel, escalaSub.length-1)];
      rowFull.format.fill.color=tom;
      rowFull.format.font.color=PALETA.navy;
      rowFull.format.font.bold=true;
      fundirBordas(rowFull, tom);
      bordaAresta(rowFull,"EdgeTop",PALETA.cobre,"Thin");
    } else {
      const dataCells=full.getCell(r,labelCols).getResizedRange(0, nCols-1-labelCols);
      dataCells.format.fill.color=(dataIdx%2===0)?PALETA.offwhite:PALETA.warmgray;
      dataCells.format.font.color=PALETA.graphite;
      gradeClara(dataCells);
      dataIdx++;
    }
  }

  // 7) Divisor cobre entre sidebar e dados
  bordaAresta(full.getCell(headerRows,0).getResizedRange(nRows-1-headerRows, labelCols-1), "EdgeRight", PALETA.cobre, "Thin");

  // 8) Coluna de Total Geral
  if(colTotalGeral>=0){
    const col=full.getCell(headerRows,colTotalGeral).getResizedRange(nRows-headerRows-1,0);
    col.format.fill.color=PALETA.navy;
    col.format.font.color=PALETA.white;
    col.format.font.bold=true;
    fundirBordas(col, PALETA.navy);
  }

  bordaExterna(full, PALETA.cobre, "Thin");
}

/* --------------------------------------------------------------------------
   TABELA
   -------------------------------------------------------------------------- */
async function formatarTabela(ctx, table){
  table.style="TableStyleLight1";
  table.showBandedRows=false;
  table.load("showTotals");
  const headerRange=table.getHeaderRowRange();
  const bodyRange=table.getDataBodyRange();
  bodyRange.load("rowCount, columnCount");
  await ctx.sync();

  headerRange.format.fill.color=PALETA.navy;
  headerRange.format.font.color=PALETA.white;
  headerRange.format.font.bold=true;
  gradeTitulo(headerRange);
  bordaAresta(headerRange,"EdgeBottom",PALETA.cobre,"Medium");

  const nr=bodyRange.rowCount;
  for(let r=0;r<nr;r++){
    const linha=bodyRange.getRow(r);
    linha.format.fill.color=(r%2===0)?PALETA.offwhite:PALETA.warmgray;
    linha.format.font.color=PALETA.graphite;
  }
  gradeClara(bodyRange);

  if(table.showTotals){
    const totalRange=table.getTotalRowRange();
    totalRange.format.fill.color=PALETA.navy;
    totalRange.format.font.color=PALETA.white;
    totalRange.format.font.bold=true;
    fundirBordas(totalRange, PALETA.navy);
    bordaAresta(totalRange,"EdgeTop",PALETA.cobre,"Medium");
  }
  bordaExterna(table.getRange(), PALETA.cobre, "Thin");
}

/* --------------------------------------------------------------------------
   INTERVALO
   -------------------------------------------------------------------------- */
async function formatarIntervalo(ctx, range){
  range.load("rowCount, columnCount");
  await ctx.sync();
  const nr=range.rowCount, ncc=range.columnCount;
  range.format.fill.color=PALETA.offwhite;
  range.format.font.color=PALETA.graphite;

  const header=range.getRow(0);
  header.format.fill.color=PALETA.navy;
  header.format.font.color=PALETA.white;
  header.format.font.bold=true;
  gradeTitulo(header);
  bordaAresta(header,"EdgeBottom",PALETA.cobre,"Medium");

  for(let r=1;r<nr;r++){
    const linha=range.getRow(r);
    linha.format.fill.color=((r-1)%2===0)?PALETA.offwhite:PALETA.warmgray;
    linha.format.font.color=PALETA.graphite;
  }
  if(nr>1) gradeClara(range.getCell(1,0).getResizedRange(nr-2, ncc-1));
  bordaExterna(range, PALETA.cobre, "Thin");
}

/* --------------------------------------------------------------------------
   Bordas
   -------------------------------------------------------------------------- */
function fundirBordas(range, cor){
  const b=range.format.borders;
  ["InsideHorizontal","InsideVertical","EdgeTop","EdgeBottom","EdgeLeft","EdgeRight"]
    .forEach(e=>{ const it=b.getItem(e); it.color=cor; it.weight="Thin"; it.style="Continuous"; });
}
function gradeClara(range){
  const b=range.format.borders;
  ["InsideHorizontal","InsideVertical"]
    .forEach(e=>{ const it=b.getItem(e); it.color=PALETA.gridline; it.weight="Thin"; it.style="Continuous"; });
}
// Barra de títulos: grade SÓ vertical clara (horizontais não aparecem)
function gradeTitulo(range){
  const b=range.format.borders;
  const v=b.getItem("InsideVertical"); v.color=PALETA.linhaTitulo; v.weight="Thin"; v.style="Continuous";
  b.getItem("InsideHorizontal").style="None";
}
// Colunas de rótulo: SÓ horizontal; vertical removida (não quebra a cor de fundo)
function gradeSomenteHorizontal(range){
  const b=range.format.borders;
  const h=b.getItem("InsideHorizontal"); h.color=PALETA.linhaTitulo; h.weight="Thin"; h.style="Continuous";
  b.getItem("InsideVertical").style="None";
}
// Remove todas as linhas internas (deixa o fundo escuro sólido)
function limparBordasInternas(range){
  const b=range.format.borders;
  b.getItem("InsideVertical").style="None";
  b.getItem("InsideHorizontal").style="None";
}
function bordaAresta(range, aresta, cor, peso){
  const it=range.format.borders.getItem(aresta);
  it.color=cor; it.weight=peso; it.style="Continuous";
}
function bordaExterna(range, cor, peso){
  ["EdgeTop","EdgeBottom","EdgeLeft","EdgeRight"].forEach(e=>bordaAresta(range,e,cor,peso));
}

/* -------------------------------------------------------------------------- */
function notificar(msg,erro){ console.log(msg); if(window.__starkerStatus) window.__starkerStatus(msg,erro); }

Office.onReady(()=>{
  if(Office.actions && Office.actions.associate){
    Office.actions.associate("aplicarSelecao", aplicarSelecao);
    Office.actions.associate("aplicarPlanilha", aplicarPlanilha);
  }
});
if(typeof window!=="undefined") window.STARKER={ aplicarSelecao, aplicarPlanilha, PALETA };
