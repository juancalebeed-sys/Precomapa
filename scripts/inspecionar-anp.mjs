import * as XLSX from 'xlsx';

const PAGE = 'https://www.gov.br/anp/pt-br/assuntos/precos-e-defesa-da-concorrencia/precos/levantamento-de-precos-de-combustiveis-ultimas-semanas-pesquisadas';
const HEAD = { 'User-Agent': 'Mozilla/5.0 (compatible; PrecoMapaBot/1.0; +https://precomapa.vercel.app)' };

const r = await fetch(PAGE, { headers: HEAD });
console.log('Pagina:', r.status);
const html = await r.text();

const links = [...html.matchAll(/href=["']([^"']+?\.xlsx[^"']*)["']/gi)]
  .map((m) => new URL(m[1], PAGE).href);
console.log('Links .xlsx:', links.length);
links.slice(0, 12).forEach((l) => console.log(' -', l));

const alvo = links.find((l) => /resumo/i.test(l)) || links[0];
if (!alvo) throw new Error('Nenhuma planilha encontrada na pagina');
console.log('\nInspecionando:', alvo);

const f = await fetch(alvo, { headers: HEAD });
console.log('Download:', f.status);
const wb = XLSX.read(Buffer.from(await f.arrayBuffer()), { type: 'buffer' });

for (const nome of wb.SheetNames) {
  const ws = wb.Sheets[nome];
  const linhas = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  console.log(`\n=== ABA "${nome}" | ${linhas.length} linhas ===`);
  linhas.slice(0, 14).forEach((l, i) =>
    console.log(i, JSON.stringify(l).slice(0, 260))
  );
}
