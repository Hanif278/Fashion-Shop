// Vercel serverless proxy -> Google Apps Script Web App.
// The secret never reaches the browser. Set in Vercel: GAS_URL, GAS_SECRET
const ALLOWED = new Set([
  'authenticateUser', 'getShopData', 'getVersion', 'getCompanyInfo', 'saveAllShopData',
  'deleteSheetEntry', 'updateInvoiceInSheet', 'getUsersList', 'saveOrUpdateUser',
  'deleteUser', 'saveCompanyInfo', 'uploadFile'
]);

module.exports = async function handler(req, res) {
  res.setHeader('Cache-Control', 'no-store');
  if (req.method !== 'POST') return res.status(405).json({ ok: false, error: 'POST only' });
  const { fn, args, auth } = req.body || {};
  if (!ALLOWED.has(fn)) return res.status(400).json({ ok: false, error: 'Unknown function' });
  if (!process.env.GAS_URL || !process.env.GAS_SECRET) {
    return res.status(500).json({ ok: false, error: 'Server not configured (GAS_URL / GAS_SECRET)' });
  }
  try {
    const r = await fetch(process.env.GAS_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ secret: process.env.GAS_SECRET, fn, args: Array.isArray(args) ? args : [], auth: auth || null }),
      redirect: 'follow'
    });
    const text = await r.text();
    let data;
    try { data = JSON.parse(text); } catch (e) {
      return res.status(502).json({ ok: false, error: 'Apps Script returned an invalid response. Check the deployment access is "Anyone".' });
    }
    return res.status(200).json(data);
  } catch (e) {
    return res.status(502).json({ ok: false, error: 'Cannot reach Apps Script: ' + e.message });
  }
};
