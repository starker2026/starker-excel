/* ============================================================================
   STÄRKER CONSULTING — Motor de formatação de tabelas (Office.js)  v9
   ----------------------------------------------------------------------------
   IMPORTANTE (lição da v8): a Tabela Dinâmica tem um ESTILO interno que desenha
   bordas azuis por baixo. Definir a borda como "None" NÃO remove essas linhas —
   ele deixa o estilo aparecer. A solução é MASCARAR: pintar a borda com a MESMA
   cor do fundo (navy sobre navy = invisível e cobre o azul do estilo).
   Por isso, aqui NÃO usamos "None" em nenhuma parte escura — sempre mascaramos.
   ========================================================================== */

const PALETA = {
  offwhite:    "#F7F6F2",
  warmgray:    "#ECEBE7",
  graphite:    "#262626",
  navy:        "#0B202D",
  slate:       "#2B4F63",
  linhaTitulo: "#48627A", // linha horizontal clara da barra de títulos
  subtotalTop: "#D9B184", // subtotal nível mais alto
  subtotalBot: "#F1E7D9", // subtotal nível mais baixo
  cobre:       "#C6793C",
  white:       "#FFFFFF",
  gridline:    "#D8D5CE"  // grade discreta nas áreas claras
};

/* Escalas */
function hexToRgb(h){ h=h.replace("#",""); return [0,2,4].map(i=>parseInt(h.substr(i,2),16)); }
function rgbToHex(r){ return "#"+r.map(x=>Math.max(0,Math.min(255,Math.round(x))).toString(16).padStart(2,"0")).join(""); }
function escala(c1,c2,n){ const a=hexToRgb(c1),b=hexToRgb(c2),o=[]; for(let i=0;i<n;i++){const t=n<=1?0:i/(n-1); o.push(rgbToHex([a[0]+(b[0]-a[0])*t,a[1]+(b[1]-a[1])*t,a[2]+(b[2]-a[2])*t]));} return o; }

/* --------------------------- Handlers --------------------------- */
async function aplicarSelecao(event){
  try{
    await Excel.run(async (ctx)=>{
      const sheet=ctx.workbook.worksheets.getActiveWorksheet();
      const sel=ctx.workbook.getSelectedRange(); sel.load("address");
      const pivots=sheet.pivotTables; pivots.load("items/name");
      const tables=sheet.tables; tables.load("items/name");
      await ctx.sync();

      let ap=null;
      for(const p of pivots.items){ const i=sel.getIntersectionOrNullObject(p.layout.getRange()); i.load("isNullObject"); p.__i=i; }
      await ctx.sync();
      for(const p of pivots.items) if(!p.__i.isNullObject){ ap=p; break; }
      if(ap){ await formatarPivot(ctx,ap); await ctx.sync(); notificar("Padrão aplicado à tabela dinâmica."); return; }

      let at=null;
      for(const t of tables.items){ const i=sel.getIntersectionOrNullObject(t.getRange()); i.load("isNullObject"); t.__i=i; }
      await ctx.sync();
      for(const t of tables.items) if(!t.__i.isNullObject){ at=t; break; }
      if(at){ await formatarTabela(ctx,at); await ctx.sync(); notificar("Padrão aplicado à tabela."); return; }

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

/* --------------------------- TABELA DINÂMICA --------------------------- */
async function formatarPivot(ctx, pivot){
  const layout=pivot.layout;
  try{ layout.preserveFormatting=true; }catch(e){}

  const full=layout.getRange();
  full.load("rowCount, columnCount, values");
  const colLabels=layout.getColumnLabelRange(); colLabels.load("rowCount");
  const rowLabels=layout.getRowLabelRange(); rowLabels.load("columnCount");
  await ctx.sync();

  const nRows=full.rowCount, nCols=full.columnCount;
  const headerRows=Math.max(1, colLabels.rowCount||1);
  const labelCols =Math.max(1, rowLabels.columnCount||1);
  const vals=full.values;
  const escR=escala(PALETA.navy, PALETA.slate, labelCols);
  const escS=escala(PALETA.subtotalTop, PALETA.subtotalBot, labelCols);

  // Base clara
  full.format.fill.color=PALETA.offwhite;
  full.format.font.color=PALETA.graphite;

  // Sidebar (colunas de rótulo): escala escura + bordas MASCARADAS na cor do fundo
  for(let c=0;c<labelCols;c++){
    const col=full.getCell(0,c).getResizedRange(nRows-1,0);
    col.format.fill.color=escR[c];
    col.format.font.color=PALETA.white;
    mascarar(col, escR[c]);   // cobre o azul do estilo e não deixa linha visível
  }

  // Barra de títulos: cor única (navy) + máscara + SÓ linhas horizontais claras
  const header=full.getCell(0,0).getResizedRange(headerRows-1, nCols-1);
  header.format.fill.color=PALETA.navy;
  header.format.font.color=PALETA.white;
  header.format.font.bold=true;
  mascarar(header, PALETA.navy);                       // some tudo (inclusive azul do estilo)
  aresta(header,"InsideHorizontal",PALETA.linhaTitulo,"Thin");  // reintroduz só as horizontais
  aresta(header,"EdgeBottom",PALETA.cobre,"Medium");

  // Total Geral (coluna)
  let colTG=-1;
  for(let c=labelCols;c<nCols;c++){ const t=String(vals[headerRows-1][c]||"").toLowerCase(); if(t.includes("total geral")||t.includes("grand total")) colTG=c; }

  // Linhas
  let dataIdx=0;
  for(let r=headerRows;r<nRows;r++){
    let nivel=0, achou=false, rot="";
    for(let c=0;c<labelCols;c++){ const t=String(vals[r][c]||""); rot+=" "+t; if(!achou&&t.toLowerCase().includes("total")){nivel=c;achou=true;} }
    rot=rot.toLowerCase();
    const rowFull=full.getCell(r,0).getResizedRange(0,nCols-1);
    const isGrand=rot.includes("total geral")||rot.includes("grand total");
    const isSub=!isGrand && rot.includes("total");

    if(isGrand){
      rowFull.format.fill.color=PALETA.navy; rowFull.format.font.color=PALETA.white; rowFull.format.font.bold=true;
      mascarar(rowFull, PALETA.navy); aresta(rowFull,"EdgeTop",PALETA.cobre,"Medium");
    } else if(isSub){
      const tom=escS[Math.min(nivel,escS.length-1)];
      rowFull.format.fill.color=tom; rowFull.format.font.color=PALETA.navy; rowFull.format.font.bold=true;
      mascarar(rowFull, tom); aresta(rowFull,"EdgeTop",PALETA.cobre,"Thin");
    } else {
      const dcells=full.getCell(r,labelCols).getResizedRange(0,nCols-1-labelCols);
      const base=(dataIdx%2===0)?PALETA.offwhite:PALETA.warmgray;
      dcells.format.fill.color=base;
      dcells.format.font.color=PALETA.graphite;
      mascarar(dcells, base);                           // cobre o azul do estilo
      aresta(dcells,"InsideHorizontal",PALETA.gridline,"Thin");  // grade discreta
      aresta(dcells,"InsideVertical",PALETA.gridline,"Thin");
      dataIdx++;
    }
  }

  // Divisor cobre entre sidebar e dados (corpo)
  aresta(full.getCell(headerRows,0).getResizedRange(nRows-1-headerRows, labelCols-1), "EdgeRight", PALETA.cobre, "Thin");

  // Coluna de Total Geral
  if(colTG>=0){
    const col=full.getCell(headerRows,colTG).getResizedRange(nRows-headerRows-1,0);
    col.format.fill.color=PALETA.navy; col.format.font.color=PALETA.white; col.format.font.bold=true;
    mascarar(col, PALETA.navy);
  }

  // Moldura cobre
  bordaExterna(full, PALETA.cobre, "Thin");
}

/* --------------------------- TABELA --------------------------- */
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
  mascarar(headerRange, PALETA.navy);
  aresta(headerRange,"EdgeBottom",PALETA.cobre,"Medium");

  const nr=bodyRange.rowCount;
  for(let r=0;r<nr;r++){
    const linha=bodyRange.getRow(r);
    const base=(r%2===0)?PALETA.offwhite:PALETA.warmgray;
    linha.format.fill.color=base;
    linha.format.font.color=PALETA.graphite;
  }
  mascarar(bodyRange, PALETA.offwhite);
  aresta(bodyRange,"InsideHorizontal",PALETA.gridline,"Thin");
  aresta(bodyRange,"InsideVertical",PALETA.gridline,"Thin");

  if(table.showTotals){
    const totalRange=table.getTotalRowRange();
    totalRange.format.fill.color=PALETA.navy;
    totalRange.format.font.color=PALETA.white;
    totalRange.format.font.bold=true;
    mascarar(totalRange, PALETA.navy);
    aresta(totalRange,"EdgeTop",PALETA.cobre,"Medium");
  }
  bordaExterna(table.getRange(), PALETA.cobre, "Thin");
}

/* --------------------------- INTERVALO --------------------------- */
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
  mascarar(header, PALETA.navy);
  aresta(header,"EdgeBottom",PALETA.cobre,"Medium");

  for(let r=1;r<nr;r++){
    const linha=range.getRow(r);
    linha.format.fill.color=((r-1)%2===0)?PALETA.offwhite:PALETA.warmgray;
    linha.format.font.color=PALETA.graphite;
  }
  if(nr>1){
    const body=range.getCell(1,0).getResizedRange(nr-2, ncc-1);
    aresta(body,"InsideHorizontal",PALETA.gridline,"Thin");
    aresta(body,"InsideVertical",PALETA.gridline,"Thin");
  }
  bordaExterna(range, PALETA.cobre, "Thin");
}

/* --------------------------- Bordas --------------------------- */
// Pinta TODAS as bordas com a cor 'cor' (máscara). Cobre linhas do estilo por baixo.
function mascarar(range, cor){
  const b=range.format.borders;
  ["InsideHorizontal","InsideVertical","EdgeTop","EdgeBottom","EdgeLeft","EdgeRight"]
    .forEach(e=>{ const it=b.getItem(e); it.color=cor; it.weight="Thin"; it.style="Continuous"; });
}
function aresta(range, qual, cor, peso){
  const it=range.format.borders.getItem(qual);
  it.color=cor; it.weight=peso; it.style="Continuous";
}
function bordaExterna(range, cor, peso){
  ["EdgeTop","EdgeBottom","EdgeLeft","EdgeRight"].forEach(e=>aresta(range,e,cor,peso));
}

/* --------------------------- Util --------------------------- */
function notificar(msg,erro){ console.log(msg); if(window.__starkerStatus) window.__starkerStatus(msg,erro); }
Office.onReady(()=>{
  if(Office.actions && Office.actions.associate){
    Office.actions.associate("aplicarSelecao", aplicarSelecao);
    Office.actions.associate("aplicarPlanilha", aplicarPlanilha);
  }
});
if(typeof window!=="undefined") window.STARKER={ aplicarSelecao, aplicarPlanilha, PALETA };
