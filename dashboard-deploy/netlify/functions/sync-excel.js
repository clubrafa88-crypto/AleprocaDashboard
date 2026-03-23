const https = require('https');
const http  = require('http');

const EXCEL_URL = 'https://aleprocaec-my.sharepoint.com/:x:/g/personal/rafael_luzuriaga_aleproca_com/IQDWUUHwP3kFQKVpuoWhcGCkATHAOrnr0bV7wAPQlRrvlMc?download=1';

function download(url, hops = 0) {
  return new Promise((resolve, reject) => {
    if (hops > 10) return reject(new Error('Too many redirects'));
    const lib = url.startsWith('https') ? https : http;
    const req = lib.get(url, {
      headers: { 'User-Agent': 'Mozilla/5.0', 'Accept': '*/*' }
    }, res => {
      if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
        return resolve(download(res.headers.location, hops + 1));
      }
      if (res.statusCode !== 200) {
        return reject(new Error(`HTTP ${res.statusCode}`));
      }
      const chunks = [];
      res.on('data', c => chunks.push(c));
      res.on('end', () => resolve(Buffer.concat(chunks)));
      res.on('error', reject);
    });
    req.on('error', reject);
    req.setTimeout(20000, () => { req.destroy(); reject(new Error('Timeout')); });
  });
}

exports.handler = async () => {
  const CORS = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, OPTIONS',
  };

  try {
    const buf = await download(EXCEL_URL);
    return {
      statusCode: 200,
      headers: { ...CORS, 'Content-Type': 'application/octet-stream' },
      body: buf.toString('base64'),
      isBase64Encoded: true,
    };
  } catch (err) {
    return {
      statusCode: 500,
      headers: CORS,
      body: JSON.stringify({ error: err.message }),
    };
  }
};
