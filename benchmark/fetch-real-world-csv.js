// Downloads a real, public CSV from Taiwan's official stock exchange open
// data platform: quarterly operating results for every listed company
// (~950 rows), published under Taiwan's government open-data license.
// Not committed to the repo (see .gitignore) so this always fetches a
// fresh copy; run this to reproduce benchmark/run-real-world-csv.js.

const fs = require('fs');
const path = require('path');
const https = require('https');

const OUT_DIR = path.join(__dirname, 'real-world');
const URL = 'https://mopsfin.twse.com.tw/opendata/t187ap17_L.csv';

function get(url) {
  return new Promise((resolve, reject) => {
    https
      .get(url, (res) => {
        if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
          resolve(get(res.headers.location));
          return;
        }
        const chunks = [];
        res.on('data', (c) => chunks.push(c));
        res.on('end', () => resolve(Buffer.concat(chunks)));
      })
      .on('error', reject);
  });
}

async function main() {
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const body = await get(URL);
  const outPath = path.join(OUT_DIR, 'tw-listed-companies.csv');
  fs.writeFileSync(outPath, body);
  console.log(`Downloaded ${body.length} bytes to ${outPath}`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
