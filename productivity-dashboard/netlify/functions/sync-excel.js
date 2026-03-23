const https = require('https');
const http = require('http');

// SharePoint public link with direct download
const SHAREPOINT_URL = 'https://aleprocaec-my.sharepoint.com/:x:/g/personal/rafael_luzuriaga_aleproca_com/IQDWUUHwP3kFQKVpuoWhcGCkATHAOrnr0bV7wAPQlRrvlMc?download=1';

/**
 * Follow redirects and return an ArrayBuffer of the downloaded file.
 */
function download(url, redirectCount = 0) {
  return new Promise((resolve, reject) => {
    if (redirectCount > 10) return reject(new Error('Too many redirects'));
    const lib = url.startsWith('https') ? https : http;
    lib.get(url, {
      headers: {
        'User-Agent': 'Mozilla/5.0 (compatible; Netlify-Function/1.0)',
        'Accept': '*/*',
      }
    }, (res) => {
      if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
        return resolve(download(res.headers.location, redirectCount + 1));
      }
      if (res.statusCode !== 200) {
        return reject(new Error(`HTTP ${res.statusCode} from SharePoint`));
      }
      const chunks = [];
      res.on('data', chunk => chunks.push(chunk));
      res.on('end', () => resolve(Buffer.concat(chunks)));
      res.on('error', reject);
    }).on('error', reject);
  });
}

/**
 * Parse rows from an Excel sheet object (SheetJS).
 * Sheet layout:
 *   Row 1  (idx 0) = Fecha
 *   Row 3  (idx 2) = Operadores nómina
 *   Row 4  (idx 3) = Operadores diario
 *   Row 5  (idx 4) = Operadores total
 *   Row 6  (idx 5) = Toneladas totales
 *   Row 8  (idx 7) = Horas trabajadas
 *   Row 9  (idx 8) = Productividad (Kg/p/h)
 *   Row 13 (idx 12)= Costo MOD total
 */
function parseSheet(ws, XLSX) {
  const range = XLSX.utils.decode_range(ws['!ref']);
  const results = [];

  const num = (cell) => {
    if (!cell || cell.v === undefined || cell.v === '') return 0;
    const v = parseFloat(String(cell.v).replace(',', '.'));
    return isNaN(v) ? 0 : v;
  };

  const cellAt = (r, c) => ws[XLSX.utils.encode_cell({ r, c })];

  for (let c = range.s.c + 1; c <= range.e.c; c++) {
    const prodCell = cellAt(8, c);
    if (!prodCell || !prodCell.v || num(prodCell) === 0) continue;

    // Parse date
    const dateCell = cellAt(0, c);
    let dateVal = `col-${c}`;
    if (dateCell) {
      if (dateCell.t === 'd' && dateCell.v instanceof Date) {
        const d = dateCell.v;
        dateVal = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
      } else if (dateCell.t === 'n') {
        const d = new Date(Math.round((dateCell.v - 25569) * 86400 * 1000));
        dateVal = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
      } else if (dateCell.w) {
        const parts = dateCell.w.split('/');
        if (parts.length === 3) {
          dateVal = `${parts[2]}-${parts[1].padStart(2,'0')}-${parts[0].padStart(2,'0')}`;
        }
      }
    }

    results.push({
      date:    dateVal,
      nominal: num(cellAt(2, c)),
      daily:   num(cellAt(3, c)),
      total:   num(cellAt(4, c)),
      tons:    num(cellAt(5, c)),
      hours:   num(cellAt(7, c)),
      prod:    num(prodCell),
      cost:    num(cellAt(12, c)),
    });
  }

  return results.sort((a, b) => a.date.localeCompare(b.date));
}

exports.handler = async (event) => {
  const headers = {
    'Access-Control-Allow-Origin': '*',
    'Content-Type': 'application/json',
    'Cache-Control': 'no-store',
  };

  // Handle CORS preflight
  if (event.httpMethod === 'OPTIONS') {
    return { statusCode: 200, headers, body: '' };
  }

  try {
    // Lazy-load SheetJS (available in Netlify Functions via npm)
    let XLSX;
    try {
      XLSX = require('xlsx');
    } catch {
      return {
        statusCode: 500,
        headers,
        body: JSON.stringify({ error: 'SheetJS (xlsx) not installed. Add it to package.json.' })
      };
    }

    console.log('Downloading Excel from SharePoint…');
    const buffer = await download(SHAREPOINT_URL);
    console.log(`Downloaded ${buffer.byteLength} bytes`);

    const wb = XLSX.read(buffer, { type: 'buffer', cellDates: true, raw: false });
    console.log('Sheets:', wb.SheetNames);

    const sheetName =
      wb.SheetNames.find(n => n.toLowerCase().includes('kpi') && n.toLowerCase().includes('empaque')) ||
      wb.SheetNames.find(n => n.toLowerCase().includes('kpi')) ||
      wb.SheetNames[0];

    const ws = wb.Sheets[sheetName];
    const data = parseSheet(ws, XLSX);

    console.log(`Parsed ${data.length} rows. Last:`, data[data.length - 1]);

    return {
      statusCode: 200,
      headers,
      body: JSON.stringify({ ok: true, count: data.length, sheet: sheetName, data }),
    };
  } catch (err) {
    console.error('sync-excel error:', err.message);
    return {
      statusCode: 500,
      headers,
      body: JSON.stringify({ ok: false, error: err.message }),
    };
  }
};
