// Downloads real, public 10-Q "Financial Report" Excel exports from SEC
// EDGAR for a handful of well-known companies -- genuine multi-sheet
// (40-50+ sheets per file) filings, not generated. Not committed to the
// repo (see .gitignore) since redistributing other companies' filings
// isn't this project's place, even though they're public; run this
// yourself to reproduce benchmark/run-real-world.js's result.
//
// SEC EDGAR requires a descriptive User-Agent identifying the requester
// (see https://www.sec.gov/os/webmaster-faq#developers) and asks for a
// light request rate, both of which this respects.

const fs = require('fs');
const path = require('path');
const https = require('https');

const OUT_DIR = path.join(__dirname, 'real-world');
const USER_AGENT = 'FastExcel-benchmark contact@example.com'; // replace with your own contact info

const COMPANIES = {
  apple: '0000320193',
  microsoft: '0000789019',
  tesla: '0001318605',
  cocacola: '0000021344',
  nike: '0000320187',
  starbucks: '0000829224',
  '3dsystems': '0000910638',
  boeing: '0000012927',
  pfizer: '0000078003',
  amd: '0000002488',
};

function get(url) {
  return new Promise((resolve, reject) => {
    https
      .get(url, { headers: { 'User-Agent': USER_AGENT } }, (res) => {
        if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
          resolve(get(res.headers.location));
          return;
        }
        const chunks = [];
        res.on('data', (c) => chunks.push(c));
        res.on('end', () => resolve({ status: res.statusCode, body: Buffer.concat(chunks) }));
      })
      .on('error', reject);
  });
}

async function recentTenQAccessions(cik) {
  const { body } = await get(`https://data.sec.gov/submissions/CIK${cik}.json`);
  const json = JSON.parse(body.toString('utf-8'));
  const r = json.filings.recent;
  const out = [];
  for (let i = 0; i < r.form.length && out.length < 5; i++) {
    // Filings from a few years back reliably have a static
    // Financial_Report.xlsx archived; very recent ones sometimes don't.
    if (r.form[i] === '10-Q' && r.filingDate[i] < '2024-01-01') out.push(r.accessionNumber[i]);
  }
  return out;
}

async function main() {
  fs.mkdirSync(OUT_DIR, { recursive: true });

  for (const [name, cik] of Object.entries(COMPANIES)) {
    const cikNoDash = String(Number(cik));
    process.stdout.write(`${name}... `);
    const accessions = await recentTenQAccessions(cik);
    let ok = false;
    for (const accn of accessions) {
      const accnNoDash = accn.replace(/-/g, '');
      const url = `https://www.sec.gov/Archives/edgar/data/${cikNoDash}/${accnNoDash}/Financial_Report.xlsx`;
      const { body } = await get(url);
      if (body.length > 10000) {
        fs.writeFileSync(path.join(OUT_DIR, `${name}.xlsx`), body);
        console.log(`OK (${accn}, ${body.length} bytes)`);
        ok = true;
        break;
      }
      await new Promise((r) => setTimeout(r, 200));
    }
    if (!ok) console.log('FAILED -- no static Financial_Report.xlsx found in recent filings');
    await new Promise((r) => setTimeout(r, 300));
  }
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
