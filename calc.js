/* Orçamento de Chapas — Prime Ondulados
   Regras de cálculo, conferência do preço do vendedor e leitura do que vem colado do Excel.
   Sem DOM: usado pelo index.html e pelos testes (node teste.js).

   Regra da Prime (tabela de especificação):
     peso da chapa (kg)  = largura (m) × comprimento (m) × g/m² do papel ÷ 1000
     R$/m²               = g/m² ÷ 1000 × R$/kg que vale para o item (ver precoKgItem), em reais com 2 casas
                           (pela 3ª casa: 5 ou mais sobe, menos que 5 desce)
     V. unit. (R$)       = m² da chapa × R$/m², arredondado no centavo do mesmo jeito
     V. produtos         = V. unit. × quantidade
     IPI                 = V. produtos × 3,25%   (arredondado no centavo)
     V. total            = V. produtos + IPI
   Os preços por kg NÃO ficam no código (repositório público): ficam no navegador (localStorage). */
(function (raiz) {
'use strict';

/* ---------- tabela de papéis (gramaturas da Tabela de Especificação Prime — iguais às da
   Programação da Onduladeira). linha: parda | dupla | branca | kraft ---------- */
const LINHAS = { parda: 'Parda', dupla: 'Parede dupla', branca: 'Branca', kraft: 'Kraft (cotar)' };
const PAPEIS_PADRAO = [
  ['P25', 320, 330], ['P30', 330, 340], ['P35', 345, 354], ['P40', 355, 365], ['P45', 365, 375],
  ['P50', 388, 400], ['P55', 400, 409], ['P60', 428, 440], ['P70', 458, 470], ['P80', 514, 528],
  ['P85', 560, 576],
  ['PW40', 365, 375, 'branca'], ['PW50', 400, 420, 'branca'],
  ['PK40', 345, 355, 'kraft'], ['PK45', 378, 389, 'kraft'], ['PK50', 408, 419, 'kraft'],
  ['PK70', 448, 460, 'kraft'], ['PK80', 504, 518, 'kraft']
].flatMap(([b, gB, gC, linha]) => [
  { cod: b + 'B', base: b, onda: 'B', g: gB, linha: linha || 'parda' },
  { cod: b + 'C', base: b, onda: 'C', g: gC, linha: linha || 'parda' }
]).concat([['P50D', 538], ['P55D', 568], ['P70D', 632], ['P80D', 662]].map(([c, g]) =>
  ({ cod: c, base: c.slice(0, -1), onda: 'BC', g, linha: 'dupla' })));

const LIMITES = { largMin: 300, largMax: 1600, comprMin: 500, comprMax: 3000 };

/* ---------- utilidades ---------- */
const semAcento = s => String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '');
const normCod = s => semAcento(s).toUpperCase().replace(/[^A-Z0-9]/g, '');
function arred(v, casas) {
  if (!isFinite(v)) return v;
  const f = Math.pow(10, casas == null ? 2 : casas);
  return Math.round((v + Math.sign(v) * 1e-9) * f) / f;   // 5 na casa seguinte sobe (como o ROUND do Excel)
}
const fmtN = (v, casas) => isFinite(v) ? v.toLocaleString('pt-BR',
  { minimumFractionDigits: casas, maximumFractionDigits: casas }) : '—';
const fmtR = (v, casas) => isFinite(v) ? 'R$ ' + fmtN(v, casas == null ? 2 : casas) : '—';
function fmtMM(v) { return isFinite(v) ? (Number.isInteger(v) ? String(v) : fmtN(v, 1)) : ''; }

/* Número no formato brasileiro (ou do Excel). tipo: 'int' quantidade | 'mm' medida |
   'un' preço unitário ou por m² | 'tot' valor total */
function numBR(txt, tipo) {
  if (txt == null) return NaN;
  if (typeof txt === 'number') return txt;
  let s = String(txt).replace(/ /g, ' ').trim();
  s = s.replace(/R\$|reais|und\.?|un\.?|p[cç]s?\.?|kg|chapas?|%/gi, '').replace(/\s+/g, '');
  if (!s) return NaN;
  let neg = false;
  if (/^-/.test(s)) { neg = true; s = s.slice(1); }
  if (!/^[\d.,]+$/.test(s) || !/\d/.test(s)) return NaN;
  const temP = s.includes('.'), temV = s.includes(',');
  if (temP && temV) {
    if (s.lastIndexOf(',') > s.lastIndexOf('.')) s = s.replace(/\./g, '').replace(',', '.');
    else s = s.replace(/,/g, '');
  } else if (temV) {
    if (s.split(',').length > 2) s = s.replace(/,/g, '');
    else if (tipo === 'int' && /^\d{1,3},\d{3}$/.test(s)) s = s.replace(',', '');
    else s = s.replace(',', '.');
  } else if (temP) {
    if (s.split('.').length > 2) s = s.replace(/\./g, '');
    else if (tipo !== 'un' && /^[1-9]\d{0,2}\.\d{3}$/.test(s)) s = s.replace('.', ''); // 5.000 / 1.181 = milhar
  }
  const v = parseFloat(s);
  return isNaN(v) ? NaN : (neg ? -v : v);
}

/* medida em mm. Sem unidade: menor que 10 é metro (0,796), senão é mm. */
function paraMM(v, unid) {
  if (!isFinite(v) || v <= 0) return NaN;
  if (unid === 'mm') return v;
  if (unid === 'cm') return v * 10;
  if (unid === 'm') return v * 1000;
  return v < 10 ? v * 1000 : v;
}
function lerDim(txt) {
  if (typeof txt === 'number') return paraMM(txt);
  const m = String(txt == null ? '' : txt).trim().toLowerCase().match(/^([\d.,]+)\s*(mm|cm|m)?$/);
  if (!m) return NaN;
  return arred(paraMM(numBR(m[1], 'mm'), m[2]), 1);
}
const RE_MEDIDA = /^([\d.,]+)\s*(mm|cm|m)?\s*[x×*]\s*([\d.,]+)\s*(mm|cm|m)?$/;
function lerMedida(txt) {
  const m = String(txt == null ? '' : txt).trim().toLowerCase().replace(/\s+/g, ' ').match(RE_MEDIDA);
  if (!m) return null;
  const u1 = m[2] || m[4], u2 = m[4] || m[2];
  const l = paraMM(numBR(m[1], 'mm'), u1), c = paraMM(numBR(m[3], 'mm'), u2);
  if (!isFinite(l) || !isFinite(c)) return null;
  return { larg: arred(l, 1), compr: arred(c, 1) };
}
/* vincos: 118X560X118 (3 ou mais partes). "X", "-" ou vazio = sem vinco */
function lerVincos(txt) {
  const s = String(txt == null ? '' : txt).trim();
  if (!s || /^[x×*\-–]$/i.test(s)) return null;
  const partes = s.split(/\s*[x×*+]\s*/i);
  if (partes.length < 2) return null;
  const nums = partes.map(p => numBR(p, 'mm'));
  if (nums.some(n => !isFinite(n))) return null;
  return { partes: nums, soma: arred(nums.reduce((a, b) => a + b, 0), 1) };
}

/* ---------- papel ---------- */
function normOnda(t) {
  const s = normCod(t).replace(/^ONDA/, '');
  if (!s) return '';
  if (s === 'BC' || s === 'CB' || s === 'D' || s.startsWith('DUPLA') || s === 'PAREDEDUPLA') return 'BC';
  if (s === 'B' || s === 'C' || s === 'E') return s;
  return '?';
}
function achaPapel(txt, ondaTxt, papeis) {
  const s = normCod(String(txt == null ? '' : txt).replace(/onda|chapa|papel[aã]o|papel/gi, ''));
  if (!s) return { erro: 'falta o tipo de papel' };
  const oc = normOnda(ondaTxt);
  const conflito = p => (oc && oc !== '?' && oc !== p.onda)
    ? 'a coluna de onda diz ' + String(ondaTxt).trim().toUpperCase() + ', mas o papel ' + p.cod + ' é onda ' + p.onda : '';
  let p = papeis.find(x => normCod(x.cod) === s);
  if (p) return { papel: p, aviso: conflito(p) };
  const m = s.match(/^(P[A-Z]?\d{2})(BC|CB|D|B|C)?$/);
  if (m) {
    const base = m[1];
    let o = m[2] ? normOnda(m[2]) : oc, aviso = '';
    const daBase = papeis.filter(x => normCod(x.base) === base);
    if (!daBase.length) return { erro: 'papel ' + String(txt).trim() + ' não está na tabela' };
    if (!o || o === '?') {
      if (daBase.length === 1) return { papel: daBase[0] };
      o = 'B'; aviso = 'onda não informada — calculei como onda B';
    }
    p = daBase.find(x => x.onda === o);
    if (p) return { papel: p, aviso: aviso || (m[2] ? conflito(p) : ''), semOnda: !!aviso };
    return { erro: 'papel ' + base + ' não existe em onda ' + o };
  }
  return { erro: 'papel "' + String(txt).trim() + '" não está na tabela' };
}

/* ---------- cálculo ---------- */
/* ctx: { precoKg (geral), especiais:{ onda:{B,C,BC}, branco, papel:{cod:R$} }, ipi, tol, papeis, limites,
          precoBase, vendedores:[{nome,precoKg}] }
   R$/kg de cada item — vale o mais específico:
     preço do item (lápis) → preço do tipo de papel → branco → onda (B, C, dupla BC) → geral */
const NOME_ONDA = { B: 'onda B', C: 'onda C', BC: 'dupla (BC)' };
function precoKgItem(it, papel, ctx) {
  const e = ctx.especiais || {};
  if (it && it.precoKg > 0) return { v: +it.precoKg, fonte: 'item' };
  if (papel) {
    const pp = e.papel || {};
    const k = Object.keys(pp).find(c => normCod(c) === normCod(papel.cod));
    if (k && pp[k] > 0) return { v: +pp[k], fonte: papel.cod };
    if (papel.linha === 'branca' && e.branco > 0) return { v: +e.branco, fonte: 'branco' };
    const po = (e.onda || {})[papel.onda];
    if (po > 0) return { v: +po, fonte: NOME_ONDA[papel.onda] || papel.onda };
  }
  return { v: +ctx.precoKg || 0, fonte: 'geral' };
}
function calcItem(it, ctx) {
  const r = { erros: [], avisos: [] };
  const q = Number(it.qtde), L = Number(it.larg), C = Number(it.compr);
  const lim = ctx.limites || LIMITES;
  if (!(q > 0)) r.erros.push('falta a quantidade');
  else if (!Number.isInteger(q)) r.avisos.push('quantidade quebrada (' + fmtN(q, 2) + ')');
  if (!(L > 0) || !(C > 0)) r.erros.push(!(L > 0) && !(C > 0) ? 'falta a medida' : (!(L > 0) ? 'falta a largura' : 'falta o comprimento'));
  const ap = achaPapel(it.papel, it.onda, ctx.papeis);
  if (ap.erro) r.erros.push(ap.erro); else { r.papel = ap.papel; if (ap.aviso) r.avisos.push(ap.aviso); }
  if (L > 0 && C > 0) {
    r.m2 = L * C / 1e6;
    if (L < lim.largMin || L > lim.largMax)
      r.avisos.push('largura ' + fmtMM(L) + ' mm fora da tabela (' + lim.largMin + ' a ' + lim.largMax + ')' + (L < 200 ? ' — está em cm?' : ''));
    if (C < lim.comprMin || C > lim.comprMax)
      r.avisos.push('comprimento ' + fmtMM(C) + ' mm fora da tabela (' + lim.comprMin + ' a ' + lim.comprMax + ')' + (C < 300 ? ' — está em cm?' : ''));
    const v = lerVincos(it.vincos);
    if (v && Math.abs(v.soma - L) > 1 && Math.abs(v.soma - C) > 1)
      r.avisos.push('vincos somam ' + fmtMM(v.soma) + ' mm, mas a chapa é ' + fmtMM(L) + ' x ' + fmtMM(C));
  }
  const pk = precoKgItem(it, r.papel, ctx);
  r.fontePreco = pk.fonte;
  if (!(pk.v > 0)) { if (!r.erros.length) r.erros.push('falta o R$/kg'); }
  else if (r.papel && r.m2) {
    r.precoKg = pk.v;
    r.kgUn = r.m2 * r.papel.g / 1000;
    r.rm2 = arred(r.papel.g / 1000 * r.precoKg, 2);   // preço do m² em reais (2 casas)
    r.vUnitExato = r.m2 * r.rm2;
    r.vUnit = arred(r.vUnitExato, 2);
    if (q > 0) {
      r.m2Tot = r.m2 * q;
      r.kgTot = r.kgUn * q;
      r.vProd = arred(r.vUnit * q, 2);
      r.vIpi = arred(r.vProd * ctx.ipi, 2);
      r.vTot = arred(r.vProd + r.vIpi, 2);
    }
    if (r.papel.linha === 'kraft') r.avisos.push('papel kraft: na tabela o preço é "COTAR"');
  }
  conferir(it, r, ctx);
  r.ok = !r.erros.length && isFinite(r.vProd);
  return r;
}

/* ---------- conferência do preço do vendedor ---------- */
function conferir(it, r, ctx) {
  if (!isFinite(r.vUnitExato)) return;
  const q = Number(it.qtde), ipi = ctx.ipi;
  const fontes = [];
  // O Excel copia o que está na tela: V. unit. 6,625 vem como "R$ 6,63". Quando o total do item
  // também veio e bate com o unitário (dentro do arredondamento), o total é mais preciso.
  const meioCent = q > 0 ? 0.005 * q + 0.011 : 0.011;
  const totBate = it.vUn > 0 && it.vTot > 0 && q > 0 && Math.abs(it.vUn * q - it.vTot) <= meioCent;
  if (totBate) fontes.push(['tot', it.vTot / q]);
  if (it.vUn > 0) fontes.push(['un', +it.vUn]);
  if (it.vM2 > 0) fontes.push(['m2', it.vM2 * r.m2]);
  if (it.vTot > 0 && q > 0) fontes.push(['tot', it.vTot / q]);
  if (it.vTotIpi > 0 && q > 0) fontes.push(['totipi', it.vTotIpi / (1 + ipi) / q]);
  if (!fontes.length) return;
  const vUn = fontes[0][1];
  const c = r.conf = { vUn, fonte: fontes[0][0] };
  c.vProd = it.vTot > 0 && (totBate || !(it.vUn > 0)) ? +it.vTot
    : (it.vTotIpi > 0 && !(it.vUn > 0) ? it.vTotIpi / (1 + ipi) : (q > 0 ? vUn * q : NaN));
  c.dif = vUn - r.vUnitExato;
  c.pct = c.dif / r.vUnitExato;
  c.difTot = isFinite(r.vProd) && isFinite(c.vProd) ? c.vProd - r.vProd : NaN;
  c.kgEf = vUn / r.kgUn;                                 // R$/kg que o vendedor usou de fato
  c.gEf = vUn / r.m2 / r.precoKg * 1000;                 // g/m² que fecharia com o R$/kg combinado
  const folga = Math.max(ctx.tol || 0, 0.0051 / r.vUnitExato); // tolerância ou meio centavo
  c.status = Math.abs(c.pct) <= folga ? 'ok' : (c.pct > 0 ? 'acima' : 'abaixo');
  c.contas = [];
  if (it.vUn > 0 && it.vTot > 0 && q > 0 && !totBate)
    c.contas.push('V. unit. × quant. dá ' + fmtR(it.vUn * q) + ', mas o vendedor pôs ' + fmtR(+it.vTot) + ' no total');
  if (it.vTot > 0 && it.vTotIpi > 0 && Math.abs(it.vTot * (1 + ipi) - it.vTotIpi) > 0.011)
    c.contas.push('total com IPI do vendedor (' + fmtR(+it.vTotIpi) + ') não é o total + ' + fmtN(ipi * 100, 2) + '% (' + fmtR(it.vTot * (1 + ipi)) + ')');
  if (c.status !== 'ok') c.dica = diagnostico(c, r, ctx);
}
function diagnostico(c, r, ctx) {
  const perto = (a, b) => Math.abs(a - b) / b <= Math.max(0.003, 0.0051 / c.vUn);
  const e = ctx.especiais || {};
  const cands = [];
  if (ctx.precoKg > 0) cands.push([+ctx.precoKg, 'preço geral do orçamento']);
  if (ctx.precoBase > 0) cands.push([+ctx.precoBase, 'preço de tabela']);
  if (e.branco > 0) cands.push([+e.branco, 'preço do branco']);
  Object.entries(e.onda || {}).forEach(([o, v]) => { if (v > 0) cands.push([+v, 'preço da ' + (NOME_ONDA[o] || o)]); });
  (ctx.vendedores || []).forEach(v => { if (v.precoKg > 0) cands.push([+v.precoKg, 'R$/kg de ' + v.nome]); });
  for (const [p, nome] of cands)
    if (Math.abs(p - r.precoKg) > 0.0001 && perto(c.kgEf, p))
      return 'usou R$ ' + fmtN(p, 2) + '/kg (' + nome + ') em vez de R$ ' + fmtN(r.precoKg, 2) + '/kg';
  for (const p of ctx.papeis)
    if (p.cod !== r.papel.cod && perto(c.gEf, p.g))
      return 'parece calculado com ' + p.g + ' g/m² (' + p.cod + '); o ' + r.papel.cod + ' tem ' + r.papel.g + ' g/m²';
  for (const [preco, nome] of cands)
    for (const p of ctx.papeis)
      if (p.cod !== r.papel.cod && perto(c.vUn / r.m2 / preco * 1000, p.g))
        return 'parece calculado com ' + p.g + ' g/m² (' + p.cod + ') e ' + nome + ' (R$ ' + fmtN(preco, 2) + '/kg)';
  return 'equivale a R$ ' + fmtN(c.kgEf, 2) + '/kg — ou a ' + fmtN(c.gEf, 0) + ' g/m² no R$ ' +
    fmtN(r.precoKg, 2) + '/kg combinado (o ' + r.papel.cod + ' tem ' + r.papel.g + ' g/m²)';
}

function calcOrcamento(orc, ctx) {
  const rs = (orc.itens || []).map(it => calcItem(it, ctx));
  const t = { itens: rs.length, chapas: 0, m2: 0, kg: 0, vProd: 0, vIpi: 0, vTot: 0, erros: 0, avisos: 0,
    conf: { n: 0, div: 0, acima: 0, abaixo: 0, vendProd: 0, corretoProd: 0 } };
  rs.forEach((r, i) => {
    const q = Number(orc.itens[i].qtde) || 0;
    if (r.erros.length) t.erros++;
    if (r.avisos.length) t.avisos++;
    if (!r.ok) return;
    t.chapas += q; t.m2 += r.m2Tot; t.kg += r.kgTot;
    t.vProd += r.vProd; t.vIpi += r.vIpi; t.vTot += r.vTot;
    if (r.conf && isFinite(r.conf.vProd)) {
      t.conf.n++;
      if (r.conf.status !== 'ok') { t.conf.div++; t.conf[r.conf.status]++; }
      t.conf.vendProd += r.conf.vProd; t.conf.corretoProd += r.vProd;
    }
  });
  ['vProd', 'vIpi', 'vTot'].forEach(k => { t[k] = arred(t[k], 2); });
  if (orc.vendTotal > 0) {
    const vend = orc.vendTotalModo === 'c' ? orc.vendTotal : orc.vendTotal * (1 + ctx.ipi);
    t.vendTotal = { vTot: vend, dif: vend - t.vTot, pct: t.vTot ? (vend - t.vTot) / t.vTot : NaN };
  }
  return { rs, t };
}

/* ---------- leitura do que foi colado ---------- */
const CAMPOS = {
  ign: '— ignorar —', qtde: 'Quantidade', larg: 'Largura', compr: 'Comprimento', medida: 'Medida (L x C)',
  papel: 'Tipo de papel', onda: 'Onda', vincos: 'Vincos',
  vUn: 'Vendedor: R$ por chapa', vM2: 'Vendedor: R$ por m²', vTot: 'Vendedor: total s/ IPI', vTotIpi: 'Vendedor: total c/ IPI'
};
/* ordem em que o Marciliano costuma colar coluna por coluna */
const ORDEM_COLAGEM = ['larg', 'compr', 'qtde', 'papel', 'vincos', 'vUn'];

function lerTSV(txt, opc) {
  opc = opc || {};
  txt = String(txt == null ? '' : txt).replace(/\r\n?/g, '\n');
  const linhas = []; let lin = [], cel = '', aspas = false;
  for (let i = 0; i < txt.length; i++) {
    const ch = txt[i];
    if (aspas) {
      if (ch === '"') { if (txt[i + 1] === '"') { cel += '"'; i++; } else aspas = false; }
      else cel += ch;
      continue;
    }
    if (ch === '"' && cel === '') { aspas = true; continue; }
    if (ch === '\t') { lin.push(cel); cel = ''; continue; }
    if (ch === '\n') { lin.push(cel); linhas.push(lin); lin = []; cel = ''; continue; }
    cel += ch;
  }
  if (cel !== '' || lin.length) { lin.push(cel); linhas.push(lin); }
  let out = linhas.map(l => l.map(c => c.replace(/ /g, ' ').replace(/\s*\n\s*/g, ' ').trim()));
  const temTab = out.some(l => l.length > 1);
  if (!temTab && !opc.semEspacos) out = out.map(l => separaEspacos(l[0] || ''));
  out = out.filter(l => l.some(c => c !== ''));
  const n = out.reduce((m, l) => Math.max(m, l.length), 0);
  return out.map(l => { while (l.length < n) l.push(''); return l; });
}
/* colado sem tab (PDF, WhatsApp): separa por espaço, mas mantém "796 x 2380" e "R$ 6,82" juntos */
function separaEspacos(s) {
  s = String(s).replace(/R\$\s*/gi, '')
    .replace(/(\d)\s*(mm|cm)?\s*[x×*]\s*(?=\d)/gi, (m0, d, u) => d + (u || '') + 'x')
    .trim();
  if (!s) return [''];
  if (/^\d+([.,]\d+)?\s+(mm|cm|m)$/i.test(s)) return [s];
  return s.split(/\s+/);
}

const CABECALHOS = [
  [/^(QUANT|QTDE|QTD|QUANTIDADE|QTDADE|PECAS|QUANTIDADES)$/, 'qtde'],
  [/^(LARG|LARGURA|LARG(MM)?)$/, 'larg'],
  [/^(COMPR|COMP|COMPRIMENTO|FACAO|COMPRIMENTOFACAO)$/, 'compr'],
  [/^(TAMANHO|MEDIDA|MEDIDAS|DIMENSAO|DIMENSOES|FORMATO|TAM)$/, 'medida'],
  [/^(REFERENCIA|REF|PAPEL|TIPOPAPEL|TIPODEPAPEL|PO|COMPOSICAO|TIPO)$/, 'papel'],
  [/^ONDA$/, 'onda'],
  [/^VINCOS?$/, 'vincos'],
  [/^(VUNIT|VALORUNIT|VALORUNITARIO|PRECOUNIT|PRECOUNITARIO|UNIT|UNITARIO|PRECO|VLUNIT|PRECOUN|VUN)$/, 'vUn'],
  [/^(RM2|PRECOM2|VALORM2|M2R|PRECOPORM2)$/, 'vM2'],
  [/^(VPROD|VPRODUTOS|VALORPRODUTOS|SUBTOTAL|VPRODUTO|TOTALSIPI|TOTALSEMIPI)$/, 'vTot'],
  [/^(VTOTAL|VALORTOTAL|TOTAL|TOTALCIPI|TOTALCOMIPI)$/, 'vTotIpi'],
  [/^(IPI|VIPI|PESOUNIT|PESOTOTAL|PESO|DESCR|DESCRICAO|PRODUTO|ITEM|N|NO|UND|UNID|UN)$/, 'ign']
];
function campoDoCabecalho(txt) {
  const s = normCod(txt);
  if (!s) return null;
  for (const [re, campo] of CABECALHOS) if (re.test(s)) return campo;
  return null;
}
const RE_PAPEL = /^P[A-Z]?\s*-?\s*\d{2}\s*-?\s*(BC|D|B|C)?$/i;
function tipoValor(v, papeis) {
  const s = String(v).trim();
  if (!s) return '';
  if (lerMedida(s)) return 'medida';
  if (lerVincos(s) && s.split(/[x×*+]/i).length >= 3) return 'vincos';
  if (RE_PAPEL.test(s) || (papeis && papeis.some(p => normCod(p.cod) === normCod(s)))) return 'papel';
  const o = normOnda(s);
  if (o && o !== '?' && s.length <= 6) return 'onda';
  if (isFinite(numBR(s, 'un'))) return 'num';
  return 'txt';
}

/* Adivinha o que é cada coluna. linhas = matriz de texto. ctx.itensAtuais (nº de itens já na lista)
   e ctx.camposVazios (campos ainda vazios nos itens atuais, na ordem de colagem). */
function analisaColagem(linhas, ctx) {
  ctx = ctx || {};
  const papeis = ctx.papeis || PAPEIS_PADRAO;
  // cabeçalho: a linha com mais nomes de coluna conhecidos (2 ou mais)
  let iCab = -1, melhor = 1;
  linhas.forEach((l, i) => {
    const n = l.filter(c => { const k = campoDoCabecalho(c); return k && k !== 'ign'; }).length;
    if (n > melhor || (n === melhor && n >= 2 && iCab < 0)) { melhor = n; iCab = i; }
  });
  if (iCab < 0 && linhas.length && linhas[0].some(c => { const k = campoDoCabecalho(c); return k && k !== 'ign'; })
      && !linhas[0].some(c => ['medida', 'num', 'vincos', 'papel'].includes(tipoValor(c, papeis)))) iCab = 0;
  const cab = iCab >= 0 ? linhas[iCab] : null;
  const dados = iCab >= 0 ? linhas.slice(iCab + 1) : linhas.slice();
  const nCol = linhas.reduce((m, l) => Math.max(m, l.length), 0);

  const info = [];
  for (let j = 0; j < nCol; j++) {
    const vals = dados.map(l => l[j] || '').filter(v => v !== '');
    const tipos = {};
    vals.forEach(v => { const t = tipoValor(v, papeis); tipos[t] = (tipos[t] || 0) + 1; });
    const n = vals.length;
    const domina = t => n && (tipos[t] || 0) / n >= 0.6;
    // "5.000" / "1.650" copiados do Excel em pt-BR são milhar (o decimal vem com vírgula)
    const RE_MILHAR = /^\s*\d{1,3}(\.\d{3})+\s*$/;
    const nums = vals.map(v => numBR(v, RE_MILHAR.test(v) ? 'int' : 'un')).filter(isFinite);
    const inteiros = nums.length && nums.every(x => Number.isInteger(x));
    info.push({ j, cab: cab ? cab[j] : '', campoCab: cab ? campoDoCabecalho(cab[j]) : null, n, tipos, domina, nums, vals,
      inteiros: !!inteiros, amostra: vals.slice(0, 3) });
  }

  const mapa = new Array(nCol).fill('ign');
  const usado = new Set();
  const poe = (j, campo) => { mapa[j] = campo; if (campo !== 'ign') usado.add(campo); };

  // 1) conteúdo inconfundível
  info.forEach(c => {
    if (!c.n) return;
    if (c.domina('medida')) poe(c.j, usado.has('medida') ? 'ign' : 'medida');
    else if (c.domina('vincos')) poe(c.j, usado.has('vincos') ? 'ign' : 'vincos');
    else if (c.domina('papel')) poe(c.j, usado.has('papel') ? 'ign' : 'papel');
    else if (c.domina('onda') && c.campoCab !== 'papel') poe(c.j, usado.has('onda') ? 'ign' : 'onda');
  });
  // 2) cabeçalho para as colunas de número
  info.forEach(c => {
    if (!c.n || mapa[c.j] !== 'ign' || !c.domina('num')) return;
    if (c.campoCab && !['medida', 'papel', 'onda', 'vincos'].includes(c.campoCab) && !usado.has(c.campoCab))
      poe(c.j, c.campoCab);
  });
  // 3) números sem cabeçalho
  const numCols = info.filter(c => c.n && mapa[c.j] === 'ign' && c.domina('num') && !c.campoCab);
  const ehSequencia = c => c.nums.length >= 2 && c.nums.every((x, k) => x === k + 1);
  const ehIPI = c => c.nums.length && c.nums.every(x => Math.abs(x - 0.0325) < 1e-9 || Math.abs(x - 3.25) < 1e-9);
  const livres = numCols.filter(c => !ehSequencia(c) && !ehIPI(c));
  const ints = livres.filter(c => c.inteiros), decs = livres.filter(c => !c.inteiros);
  const temMedida = usado.has('medida') || (usado.has('larg') && usado.has('compr'));
  const faixa = (c, a, b) => c.nums.every(x => x >= a && x <= b);
  if (temMedida) {
    ints.forEach(c => { if (!usado.has('qtde')) poe(c.j, 'qtde'); });
  } else if (ints.length >= 2) {
    // procura um par vizinho largura × comprimento
    let par = null;
    for (let k = 0; k + 1 < ints.length && !par; k++) {
      const a = ints[k], b = ints[k + 1];
      if (b.j === a.j + 1 && faixa(a, 100, 2600) && faixa(b, 100, 3600)) par = [a, b];
    }
    if (par) {
      if (!usado.has('larg')) poe(par[0].j, 'larg');
      if (!usado.has('compr')) poe(par[1].j, 'compr');
    }
    ints.forEach(c => {
      if (mapa[c.j] !== 'ign') return;
      for (const campo of ['qtde', 'larg', 'compr']) if (!usado.has(campo)) { poe(c.j, campo); break; }
    });
  } else if (ints.length === 1) {
    // uma coluna só (colando coluna por coluna): o primeiro campo ainda vazio nos itens atuais
    const prox = (ctx.camposVazios || ORDEM_COLAGEM).find(f => ['larg', 'compr', 'qtde'].includes(f) && !usado.has(f));
    poe(ints[0].j, prox || 'qtde');
  }
  // preços: o menor é o unitário; os outros, total sem e com IPI
  if (decs.length) {
    const media = c => c.nums.reduce((a, b) => a + b, 0) / c.nums.length;
    const ord = decs.slice().sort((a, b) => media(a) - media(b));
    const camposPreco = ['vUn', 'vTot', 'vTotIpi'].filter(f => !usado.has(f));
    ord.forEach((c, k) => { if (camposPreco[k]) poe(c.j, camposPreco[k]); });
  }
  // inteiros que sobraram e parecem total (ex.: V. PROD. 1899)
  ints.forEach(c => {
    if (mapa[c.j] !== 'ign') return;
    if (usado.has('vUn') && !usado.has('vTot')) poe(c.j, 'vTot');
  });
  return { iCab, mapa, info, nCol, dados };
}

/* monta itens a partir das linhas e do mapa de colunas */
function montaItens(dados, mapa) {
  const out = [];
  let vazias = 0;
  dados.forEach((l, i) => {
    const it = {}, prob = [];
    let algum = false;
    mapa.forEach((campo, j) => {
      if (campo === 'ign') return;
      const v = (l[j] || '').trim();
      if (!v) return;
      algum = true;
      let x;
      switch (campo) {
        case 'qtde': x = numBR(v, 'int'); if (isFinite(x)) it.qtde = x; else prob.push('"' + v + '" não é quantidade'); break;
        case 'larg': case 'compr': {
          const md = lerMedida(v);
          if (md && campo === 'larg') { it.larg = md.larg; it.compr = md.compr; break; }
          x = lerDim(v); if (isFinite(x)) it[campo] = x; else prob.push('"' + v + '" não é ' + CAMPOS[campo].toLowerCase());
          break;
        }
        case 'medida': {
          const md = lerMedida(v);
          if (md) { it.larg = md.larg; it.compr = md.compr; } else prob.push('medida "' + v + '" não entendida');
          break;
        }
        case 'papel': it.papel = v.replace(/\s+/g, ' ').toUpperCase(); break;
        case 'onda': it.onda = v.toUpperCase(); break;
        case 'vincos': it.vincos = /^[x×*\-–]$/i.test(v) ? '' : v.replace(/\s+/g, ''); break;
        case 'vUn': case 'vM2': case 'vTot': case 'vTotIpi':
          x = numBR(v, campo === 'vUn' || campo === 'vM2' ? 'un' : 'tot');
          if (isFinite(x)) it[campo] = x; else prob.push('"' + v + '" não é valor');
          break;
      }
    });
    // linha que só tem onda/vincos (ex.: "FRETE  CIF  FOB" no rodapé do pedido) não é item
    const temDado = ['qtde', 'larg', 'compr', 'papel', 'vUn', 'vM2', 'vTot', 'vTotIpi'].some(k => it[k] != null);
    if (!algum || (!temDado && !prob.length)) { vazias++; return; }
    out.push({ linha: i, item: it, prob, orig: l });
  });
  return { linhas: out, vazias };
}

/* dados do cabeçalho/rodapé do pedido, quando a planilha inteira é colada
   (CLIENTE:, CPF/CNPJ:, VENDEDOR, FONE:, FRETE, PRAZO PGTO: ..., Entrega: ...) */
const ROTULOS = [
  [/^(CLIENTE|RAZAOSOCIAL)$/, 'cliente'], [/^(CPFCNPJ|CNPJ|CPF|CNPJCPF)$/, 'cnpj'],
  [/^(INSC|IE|INSCRICAO|INSCRICAOESTADUAL|INSCEST)$/, 'ie'], [/^(ENDERECO|END)$/, 'endereco'],
  [/^BAIRRO$/, 'bairro'], [/^(CIDADE|CIDADEUF|MUNICIPIO)$/, 'cidade'], [/^E?MAIL$/, 'email'],
  [/^(CONTATO|COMPRADOR|COMPRADORA|AC)$/, 'contato'], [/^(FONE|TELEFONE|TEL|CELULAR|WHATSAPP)$/, 'fone'],
  [/^(VENDEDOR|VENDEDORA|REPRESENTANTE|REPRES|REPRESENTANTECOMERCIAL)$/, 'vendNome'], [/^FRETE$/, 'frete'],
  [/^(PRAZO|PRAZOPGTO|PRAZOPAGTO|PRAZODEPAGAMENTO|CONDPAGTO|CONDICAODEPAGAMENTO|PAGAMENTO)$/, 'prazo'],
  [/^(ENTREGA|DATAENTREGA|DATADEENTREGA|PREVENTREGA|PREVISAODEENTREGA)$/, 'entrega']
];
function lerDadosPedido(linhas) {
  const d = {};
  let viuVend = false;
  linhas.forEach(l => {
    for (let j = 0; j < l.length; j++) {
      const cel = l[j];
      if (!cel) continue;
      let rot = cel, val = '';
      const m = cel.match(/^([^:]{2,30}):\s*(.*)$/);
      if (m) { rot = m[1]; val = m[2].trim(); }
      const achado = ROTULOS.find(([re]) => re.test(normCod(rot)));
      if (!achado) continue;
      let campo = achado[1];
      if (!val) for (let k = j + 1; k < l.length; k++) if (l[k]) { val = l[k]; j = k; break; }
      if (!val || ROTULOS.some(([re]) => re.test(normCod(val.replace(/:$/, ''))))) continue;
      if (campo === 'vendNome') viuVend = true;
      if (campo === 'fone' || campo === 'email') {
        if (viuVend) campo = campo === 'fone' ? 'vendFone' : 'vendEmail';
        else if (campo === 'fone') continue;          // o modelo não tem fone do cliente
      }
      if (campo === 'frete') val = /CIF/i.test(val) ? 'CIF' : (/FOB/i.test(val) ? 'FOB' : val);
      if (d[campo] == null) d[campo] = val.trim();
    }
  });
  return d;
}

/* ajeita o código do papel quando ele foi reconhecido sem dúvida (P25 + onda B → P25B) */
function normalizaPapel(it, papeis) {
  if (!it.papel) return it;
  const ap = achaPapel(it.papel, it.onda, papeis);
  if (ap.papel && !ap.semOnda && !ap.aviso) { it.papel = ap.papel.cod; delete it.onda; }
  return it;
}

const api = { LINHAS, PAPEIS_PADRAO, LIMITES, CAMPOS, ORDEM_COLAGEM, NOME_ONDA, precoKgItem, semAcento, normCod, arred, fmtN, fmtR, fmtMM,
  numBR, paraMM, lerDim, lerMedida, lerVincos, normOnda, achaPapel, calcItem, calcOrcamento, lerTSV,
  analisaColagem, montaItens, normalizaPapel, campoDoCabecalho, tipoValor, lerDadosPedido };
if (typeof module !== 'undefined' && module.exports) module.exports = api; else raiz.Calc = api;
})(typeof window !== 'undefined' ? window : globalThis);
