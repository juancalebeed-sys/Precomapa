import * as XLSX from 'xlsx';
import { writeFileSync } from 'node:fs';

const PAGE = 'https://www.gov.br/anp/pt-br/assuntos/precos-e-defesa-da-concorrencia/precos/levantamento-de-precos-de-combustiveis-ultimas-semanas-pesquisadas';
const HEAD = { 'User-Agent': 'Mozilla/5.0 (compatible; PrecoMapaBot/1.0; +https://precomapa.vercel.app)' };
const GRAVAR = process.env.GRAVAR === 'sim';
const SB = (() => { try { return new URL(process.env.SUPABASE_URL || '').origin; } catch { return ''; } })();
const KEY = process.env.SUPABASE_SERVICE_KEY || '';
const SOURCE_ID = 1;
const QTD = Number(process.env.ULTIMAS || 2);

const norm = (s) => String(s ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '')
  .toUpperCase().replace(/[-_]/g, ' ').replace(/\s+/g, ' ').trim();
const iso = (v) => {
  if (typeof v === 'number') return new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 864e5).toISOString().slice(0, 10);
  const s = String(v ?? '');
  const a = s.match(/(\d{4})-(\d{2})-(\d{2})/);
  if (a) return a[0];
  const b = s.match(/(\d{2})\/(\d{2})\/(\d{4})/);
  return b ? `${b[3]}-${b[2]}-${b[1]}` : null;
};
const num = (v) => {
  if (v === null || v === '') return null;
  const n = typeof v === 'number' ? v : Number(String(v).replace(',', '.'));
  return Number.isFinite(n) ? Math.round(n * 1000) / 1000 : null;
};

const H = { apikey: KEY, 'Content-Type': 'application/json' };
if (KEY.startsWith('eyJ')) H.Authorization = `Bearer ${KEY}`;
async function sb(path, opt = {}) {
  const r = await fetch(`${SB}/rest/v1/${path}`, { ...opt, headers: { ...H, ...(opt.headers || {}) } });
  const t = await r.text();
  if (!r.ok) throw new Error(`Supabase ${r.status}: ${t.slice(0, 300)}`);
  return t ? JSON.parse(t) : null;
}

const PROD = {
  'GASOLINA COMUM': 'gasolina-comum', 'GASOLINA ADITIVADA': 'gasolina-aditivada',
  'ETANOL HIDRATADO': 'etanol-hidratado', 'GLP': 'glp-13kg', 'GNV': 'gnv',
  'OLEO DIESEL': 'oleo-diesel', 'OLEO DIESEL S10': 'oleo-diesel-s10',
};

console.log(GRAVAR ? 'MODO: GRAVAR' : 'MODO: TESTE (nada sera gravado)');
if (!SB || !KEY) throw new Error('Faltam os segredos SUPABASE_URL e SUPABASE_SERVICE_KEY');

const produtos = {};
for (const p of await sb('products?select=id,slug,unit')) produtos[p.slug] = p;
const locs = await sb('locations?select=id,level,name,uf');
const mapa = {}, estUF = {};
for (const l of locs) {
  if (l.level === 'estado') { mapa['estado|' + norm(l.name)] = l.id; estUF[norm(l.name)] = String(l.uf).trim(); }
  if (l.level === 'regiao') mapa['regiao|' + norm(l.name)] = l.id;
  if (l.level === 'pais') mapa['pais|BRASIL'] = l.id;
}
for (const l of locs) {
  if (l.level === 'municipio') mapa['municipio|' + norm(l.name) + '|' + String(l.uf).trim()] = l.id;
}

const html = await (await fetch(PAGE, { headers: HEAD })).text();
const arquivos = [...new Set([...html.matchAll(/href=["']([^"']+?\.xlsx[^"']*)["']/gi)].map((m) => new URL(m[1], PAGE).href))]
  .map((u) => {
        const m = u.match(/resumo_semanal_lpc_\d{4}[-_]\d{2}[-_]\d{2}[-_](\d{4})[-_](\d{2})[-_](\d{2})\.xlsx/i);
    return m ? { u, fim: `${m[1]}-${m[2]}-${m[3]}` } : null;
  })
  .filter(Boolean)
  .sort((a, b) => b.fim.localeCompare(a.fim))
  .slice(0, QTD);
if (!arquivos.length) throw new Error('Nenhuma planilha resumo_semanal encontrada');

const registros = new Map();
const semfim = new Set();
const naoCasou = new Set();
let lidas = 0, ultimaSemana = '';

for (const arq of arquivos) {
  console.log('\nArquivo:', arq.u);
  const f = await fetch(arq.u, { headers: HEAD });
  if (!f.ok) throw new Error('Download falhou: ' + f.status);
  const wb = XLSX.read(Buffer.from(await f.arrayBuffer()), { type: 'buffer' });

  for (const nome of wb.SheetNames) {
    const n = norm(nome);
    const tipo = n.includes('CAPITAIS') ? 'municipio' : n.includes('ESTADOS') ? 'estado'
      : n.includes('REGIOES') ? 'regiao' : n.includes('BRASIL') ? 'pais' : null;
    if (!tipo) { console.log(`Aba ignorada: ${nome}`); continue; }

    const linhas = XLSX.utils.sheet_to_json(wb.Sheets[nome], { header: 1, raw: true, defval: null });
    const h = linhas.findIndex((l) => norm(l[0]) === 'DATA INICIAL');
    if (h < 0) throw new Error(`Cabecalho nao encontrado na aba ${nome}`);
    const cols = linhas[h].map(norm);
    const c = (p) => cols.findIndex((x) => x.startsWith(p));
    const ix = { ini: c('DATA INICIAL'), fim: c('DATA FINAL'), reg: c('REGIAO'), est: c('ESTADO'),
      mun: c('MUNICIPIO'), prod: c('PRODUTO'), n: c('NUMERO DE POSTOS'),
      med: c('PRECO MEDIO REVENDA'), min: c('PRECO MINIMO REVENDA'), max: c('PRECO MAXIMO REVENDA') };
    const exige = { pais: [], regiao: ['reg'], estado: ['est'], municipio: ['est', 'mun'] }[tipo]
      .concat(['ini', 'fim', 'prod', 'med']);
    for (const k of exige) {
      if (ix[k] < 0) throw new Error(`Aba ${nome}: coluna "${k}" nao encontrada. Cabecalho: ${cols.join(' | ')}`);
    }

    let ok = 0, total = 0;
    for (const l of linhas.slice(h + 1)) {
      if (l[ix.prod] == null) continue;
      total++; lidas++;
      const slug = PROD[norm(l[ix.prod])];
      if (!slug || !produtos[slug]) { semfim.add(String(l[ix.prod])); continue; }
      let locId;
      if (tipo === 'pais') locId = mapa['pais|BRASIL'];
      else if (tipo === 'regiao') locId = mapa['regiao|' + norm(l[ix.reg])];
      else if (tipo === 'estado') locId = mapa['estado|' + norm(l[ix.est])];
      else locId = mapa['municipio|' + norm(l[ix.mun]) + '|' + estUF[norm(l[ix.est])]];
      if (!locId) { naoCasou.add(`${tipo}:${l[ix.mun] ?? l[ix.est] ?? l[ix.reg] ?? 'Brasil'}`); continue; }

      const valor = num(l[ix.med]);
      const ini = iso(l[ix.ini]), fim = iso(l[ix.fim]);
      if (!valor || valor <= 0 || !ini || !fim) continue;
      const reg = {
        product_id: produtos[slug].id, source_id: SOURCE_ID, location_id: locId,
        price_type: 'calculado', value: valor, unit: 'R$/' + produtos[slug].unit,
        collected_at: fim, is_demo: false, period_start: ini,
        sample_size: ix.n >= 0 && num(l[ix.n]) != null ? Math.round(num(l[ix.n])) : null,
        min_value: ix.min >= 0 ? num(l[ix.min]) : null,
        max_value: ix.max >= 0 ? num(l[ix.max]) : null,
      };
      registros.set([reg.product_id, reg.location_id, reg.collected_at].join('|'), reg);
      ok++;
      if (fim > ultimaSemana) ultimaSemana = fim;
    }
    console.log(`Aba ${nome} (${tipo}): ${total} linhas, ${ok} aproveitadas`);
  }
}

const lista = [...registros.values()];
console.log(`\nTotal lido: ${lidas} | a gravar: ${lista.length} | ultima semana: ${ultimaSemana}`);
if (semfim.size) console.log('Produtos nao reconhecidos:', [...semfim].join(', '));
if (naoCasou.size) console.log(`Locais sem correspondencia (${naoCasou.size}):`, [...naoCasou].slice(0, 20).join(', '));
console.log('Exemplos:', JSON.stringify(lista.slice(0, 3)));

if (GRAVAR) {
  if (!lista.length) throw new Error('Nada para gravar');
  for (let i = 0; i < lista.length; i += 500) {
    await sb('price_records?on_conflict=product_id,source_id,location_id,price_type,collected_at', {
      method: 'POST',
      headers: { Prefer: 'resolution=merge-duplicates,return=minimal' },
      body: JSON.stringify(lista.slice(i, i + 500)),
    });
  }
  console.log('Gravado com sucesso:', lista.length);
  try {
    await sb('import_logs', { method: 'POST', headers: { Prefer: 'return=minimal' },
      body: JSON.stringify({ source_id: SOURCE_ID, rows_read: lidas, rows_saved: lista.length, status: 'ok',
        message: 'Semana ate ' + ultimaSemana }) });
  } catch (e) { console.log('Aviso: nao consegui registrar em import_logs:', e.message); }
  writeFileSync('dados-status.json', JSON.stringify({ ultima_semana: ultimaSemana }) + '\n');
                   }
