const fs = require('fs');
const { Client } = require('pg');
const url = fs.readFileSync('C:/Users/ISO/.sb_db.url', 'utf8').trim();
(async () => {
  const c = new Client({ connectionString: url, ssl: { rejectUnauthorized: false } });
  await c.connect();
  const r = await c.query(`select proname, pg_get_functiondef(oid) as def from pg_proc where proname in ('catalogo_wh_delta','catalogo_wh_rls') and pronamespace='mos'::regnamespace`);
  for (const row of r.rows) {
    console.log('=== ' + row.proname + ' ===');
    console.log(row.def);
    console.log('');
  }
  await c.end();
})().catch(e => { console.error(e); process.exit(1); });
