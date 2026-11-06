/* Orçamento de Chapas — gera o .xlsx a partir do modelo-orcamento.xlsx (mesmo layout do
   Pasta1.xlsx: logo no topo, 31 linhas de item, VALOR TOTAL / PESO TOTAL e rodapé com endereço).
   Mexe direto no XML da planilha para não perder nada do modelo (imagens, bordas, impressão).
   Mais de 31 itens: insere linhas e empurra o rodapé para baixo. */
(function (raiz) {
'use strict';
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const XMLNS = 'http://www.w3.org/XML/1998/namespace';
const PRIMEIRA = 32, ULTIMA = 62;           // linhas de item no modelo
const CAPACIDADE = ULTIMA - PRIMEIRA + 1;    // 31
const ESTILO_NUM = 56;                       // estilo extra do modelo: branco, negrito, fundo verde
const ESTILO_VINCO = 57, ESTILO_VINCO_ULT = 58;  // como os da linha de item, mas a fonte diminui p/ caber

const colDe = ref => ref.replace(/\d+$/, '');
const linDe = ref => +ref.replace(/^[A-Z]+/, '');
const colNum = c => c.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
const desloca = (ref, apartir, n) => (linDe(ref) >= apartir ? colDe(ref) + (linDe(ref) + n) : ref);
const deslocaIntervalo = (txt, apartir, n) => txt.split(' ').map(rg => rg.split(':').map(r => desloca(r, apartir, n)).join(':')).join(' ');
const filhos = (el, nome) => Array.from(el.getElementsByTagNameNS(NS, nome));

function linha(doc, sd, r) {
  const rows = filhos(sd, 'row');
  let row = rows.find(x => +x.getAttribute('r') === r);
  if (row) return row;
  row = doc.createElementNS(NS, 'row');
  row.setAttribute('r', r);
  const depois = rows.find(x => +x.getAttribute('r') > r);
  sd.insertBefore(row, depois || null);
  return row;
}
function celula(doc, sd, ref) {
  const row = linha(doc, sd, linDe(ref));
  const cs = filhos(row, 'c');
  let c = cs.find(x => x.getAttribute('r') === ref);
  if (c) return c;
  c = doc.createElementNS(NS, 'c');
  c.setAttribute('r', ref);
  const depois = cs.find(x => colNum(colDe(x.getAttribute('r'))) > colNum(colDe(ref)));
  row.insertBefore(c, depois || null);
  return c;
}
/* valor vazio = célula sem valor (nunca inlineStr vazio, que o Excel acusa como arquivo com problema) */
function poe(doc, sd, ref, v, estilo) {
  const c = celula(doc, sd, ref);
  while (c.firstChild) c.removeChild(c.firstChild);
  c.removeAttribute('t');
  if (estilo != null) c.setAttribute('s', estilo);
  if (v == null || v === '' || (typeof v === 'number' && !isFinite(v))) return;
  if (typeof v === 'number') {
    const e = doc.createElementNS(NS, 'v');
    e.textContent = String(+v.toFixed(10));
    c.appendChild(e);
  } else {
    c.setAttribute('t', 'inlineStr');
    const is = doc.createElementNS(NS, 'is'), t = doc.createElementNS(NS, 't');
    t.setAttributeNS(XMLNS, 'xml:space', 'preserve');
    t.textContent = String(v);
    is.appendChild(t); c.appendChild(is);
  }
}

/* insere n linhas de item antes da última (que tem a borda de baixo) */
function insereLinhas(doc, sd, n) {
  const rows = filhos(sd, 'row');
  rows.filter(r => +r.getAttribute('r') >= ULTIMA)
    .sort((a, b) => b.getAttribute('r') - a.getAttribute('r'))
    .forEach(r => {
      const nr = +r.getAttribute('r') + n;
      r.setAttribute('r', nr);
      filhos(r, 'c').forEach(c => c.setAttribute('r', colDe(c.getAttribute('r')) + nr));
    });
  const modelo = rows.find(r => +r.getAttribute('r') === ULTIMA - 1);
  const antes = filhos(sd, 'row').find(r => +r.getAttribute('r') === ULTIMA + n);
  for (let k = 0; k < n; k++) {
    const nr = ULTIMA + k;
    const nova = modelo.cloneNode(true);
    nova.setAttribute('r', nr);
    filhos(nova, 'c').forEach(c => {
      c.setAttribute('r', colDe(c.getAttribute('r')) + nr);
      while (c.firstChild) c.removeChild(c.firstChild);
      c.removeAttribute('t');
    });
    sd.insertBefore(nova, antes);
  }
  const ws = doc.documentElement;
  filhos(ws, 'mergeCell').forEach(m => m.setAttribute('ref', deslocaIntervalo(m.getAttribute('ref'), ULTIMA, n)));
  filhos(ws, 'conditionalFormatting').forEach(cf => {
    cf.setAttribute('sqref', cf.getAttribute('sqref').replace(/([A-Z]+)(\d+)(?=$|\s)/g, (m0, c, l) => (+l === ULTIMA ? c + (ULTIMA + n) : m0)));
  });
  const dim = filhos(ws, 'dimension')[0];
  if (dim) dim.setAttribute('ref', deslocaIntervalo(dim.getAttribute('ref'), ULTIMA, n));
}

/* mais de 25 itens não cabe na escala do modelo: imprime ajustando à largura da folha
   (até 60 itens numa página só; mais que isso, quantas páginas precisar) */
function ajustaImpressao(doc, n) {
  const ws = doc.documentElement;
  let sp = filhos(ws, 'sheetPr')[0];
  if (!sp) { sp = doc.createElementNS(NS, 'sheetPr'); ws.insertBefore(sp, ws.firstElementChild); }
  let pp = filhos(sp, 'pageSetUpPr')[0];
  if (!pp) { pp = doc.createElementNS(NS, 'pageSetUpPr'); sp.appendChild(pp); }
  pp.setAttribute('fitToPage', '1');
  const ps = filhos(ws, 'pageSetup')[0];
  if (ps) { ps.removeAttribute('scale'); ps.setAttribute('fitToWidth', '1'); ps.setAttribute('fitToHeight', n <= 60 ? '1' : '0'); }
}

/* dados: { num, data, cliente, cnpj, ie, endereco, bairro, cidade, email, contato,
            vendNome, vendFone, vendEmail, frete, prazo, entrega, ipi (fração),
            itens: [{ qtde, onda, larg, compr, vincos, vUnit, vIpi, vProd, vTot, cod, kgUn, kgTot }],
            vTot, kg } */
async function geraExcel(dados, urlModelo) {
  if (typeof JSZip === 'undefined') throw new Error('biblioteca JSZip não carregou (sem internet?)');
  const resp = await fetch(urlModelo || 'modelo-orcamento.xlsx', { cache: 'no-cache' });
  if (!resp.ok) throw new Error('não achei o modelo do Excel (' + resp.status + ')');
  const zip = await JSZip.loadAsync(await resp.arrayBuffer());
  const caminho = 'xl/worksheets/sheet1.xml';
  const doc = new DOMParser().parseFromString(await zip.file(caminho).async('string'), 'application/xml');
  if (doc.getElementsByTagName('parsererror').length) throw new Error('modelo do Excel ilegível');
  const sd = filhos(doc.documentElement, 'sheetData')[0];
  const n = dados.itens.length;
  // O modelo deixa as linhas 49 a 54 ocultas para caber numa página (25 linhas de item visíveis).
  // Até 25 itens: usa só as visíveis. Mais que isso: mostra as ocultas e, passando de 31, insere linhas.
  const ocultas = filhos(sd, 'row').filter(r => { const k = +r.getAttribute('r'); return k >= PRIMEIRA && k <= ULTIMA && r.getAttribute('hidden') === '1'; });
  const visiveis = [];
  for (let k = PRIMEIRA; k <= ULTIMA; k++) if (!ocultas.some(o => +o.getAttribute('r') === k)) visiveis.push(k);
  let extra = 0, linhasItem;
  if (n <= visiveis.length) linhasItem = visiveis.slice(0, n);
  else {
    ocultas.forEach(o => o.removeAttribute('hidden'));
    extra = Math.max(0, n - CAPACIDADE);
    if (extra) insereLinhas(doc, sd, extra);
    linhasItem = Array.from({ length: n }, (_, k) => PRIMEIRA + k);
    ajustaImpressao(doc, n);
  }
  const L = r => (r >= ULTIMA ? r + extra : r);   // linha do rodapé já deslocada
  const P = (ref, v, s) => poe(doc, sd, ref, v, s);

  // cabeçalho
  P('M15', dados.num != null && dados.num !== '' ? String(dados.num) : '', ESTILO_NUM);
  P('M17', dados.data || '');
  P('C19', dados.cliente); P('C20', dados.cnpj); P('H20', dados.ie);
  P('C21', dados.endereco); P('H21', dados.bairro); P('C22', dados.cidade);
  P('H22', dados.email); P('H23', dados.contato);
  P('C25', dados.vendNome); P('H25', dados.vendFone); P('C26', dados.vendEmail);

  // itens
  dados.itens.forEach((it, k) => {
    const r = linhasItem[k];
    P('A' + r, 'CHAPA'); P('B' + r, it.qtde); P('C' + r, 'UND'); P('D' + r, it.onda);
    P('E' + r, it.larg + 'x' + it.compr);
    P('F' + r, it.vincos || 'X', r === ULTIMA + extra ? ESTILO_VINCO_ULT : ESTILO_VINCO);
    P('G' + r, it.vUnit); P('H' + r, dados.ipi); P('I' + r, it.vIpi); P('J' + r, it.vProd);
    P('K' + r, it.vTot); P('L' + r, it.cod); P('M' + r, it.kgUn); P('N' + r, it.kgTot);
  });

  // totais e rodapé
  P('K' + L(63), dados.vTot);
  P('N' + L(63), dados.kg);
  P('C' + L(64), dados.frete); P('C' + L(65), dados.prazo); P('C' + L(66), dados.entrega);

  zip.file(caminho, new XMLSerializer().serializeToString(doc));

  if (extra) {
    // área de impressão e imagem do rodapé descem junto
    const wbP = 'xl/workbook.xml';
    let wb = await zip.file(wbP).async('string');
    wb = wb.replace(/(\$N\$)(\d+)(<\/definedName>)/, (m0, a, l, b) => a + (+l + extra) + b);
    zip.file(wbP, wb);
    const drP = 'xl/drawings/drawing1.xml';
    let dr = await zip.file(drP).async('string');
    dr = dr.replace(/<xdr:row>(\d+)<\/xdr:row>/g, (m0, l) => '<xdr:row>' + (+l >= ULTIMA - 1 ? +l + extra : +l) + '</xdr:row>');
    zip.file(drP, dr);
  }
  return zip.generateAsync({ type: 'blob', compression: 'DEFLATE',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

raiz.ExcelOrc = { geraExcel, CAPACIDADE };
})(typeof window !== 'undefined' ? window : globalThis);
