/**
 * JSON endpoint for the standalone Food log app (GitHub Pages, works offline).
 * Deploy as a SEPARATE web app: Execute as me, Who has access: Anyone.
 * Every request must carry the key kept in Script Properties (API_KEY); get it from <"Only myself" web app url>?pair=1.
 *
 * The app POSTs text/plain JSON (no CORS preflight): { key, fn: 'sync', date: 'yyyy-MM-dd', ops: [...] }
 *   op = { id, type: 'add', date, items: [{ food, amount, per, kcalPer, proteinPer, save, note, amountText? }] }
 *      | { id, type: 'update', entryId, food, kcal, protein, amount? }
 *      | { id, type: 'delete', entryId, food }
 *      | { id, type: 'weight', date, kg }            (kg '' clears that day)
 *      | { id, type: 'reorder', names: [...] }       (these Foods rows swap into this order among themselves)
 * Ops run in order; each gets 'ok' or an error message in results[id]. Log rows carry an id in column G
 * (`<op id>-<n>` for app entries, assigned on first read for others), so retries never log twice.
 */

const ID_COL = 7;

function doPost(e) {
  let out;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const key = PropertiesService.getScriptProperties().getProperty('API_KEY');
    if (!key || req.key !== key) out = { error: 'bad key' };
    else if (req.fn === 'sync') out = apiSync_(req);
    else out = { error: 'Unknown request' };
  } catch (err) {
    out = { error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function apiSync_(req) {
  const ss = ss_();
  const tz = ss.getSpreadsheetTimeZone();
  const log = ss.getSheetByName(LOG);
  const ops = (req.ops || []).slice(0, 50);
  const results = {};
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    ensureIdColumn_(log);
    ops.forEach(op => {
      try {
        applyOp_(ss, log, tz, op);
        results[op.id] = 'ok';
      } catch (err) {
        results[op.id] = String((err && err.message) || err);
      }
    });
    SpreadsheetApp.flush();
    const date = /^\d{4}-\d{2}-\d{2}$/.test(req.date || '') ? req.date : Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    return { results: results, state: apiState_(ss, log, tz, date) };
  } finally {
    lock.releaseLock();
  }
}

function applyOp_(ss, log, tz, op) {
  if (!op || !op.id) throw new Error('Missing id');
  if (op.type === 'add') {
    if (findId_(log, op.id + '-0')) return; // already logged (an earlier reply got lost)
    const clean = (op.items || []).map(cleanItem_);
    if (!clean.length) throw new Error('Nothing to add.');
    const date = dateAt_(op.date, tz);
    const rows = clean.map((it, i) => {
      const k = scale_(it.per, it.amount);
      return [date, it.food, round1_(it.kcalPer * k), it.proteinPer === '' ? '' : round1_(it.proteinPer * k),
              it.amountText != null ? it.amountText : amountText_(it.per, it.amount), it.note, op.id + '-' + i];
    });
    const first = log.getLastRow() + 1;
    if (first + rows.length - 1 > log.getMaxRows()) log.insertRowsAfter(log.getMaxRows(), 200);
    log.getRange(first, 1, rows.length, ID_COL).setValues(rows);
    log.getRange(first, 1, rows.length, 1).setNumberFormat('dd/mm/yyyy');
    rememberFoods_(ss, clean);
  } else if (op.type === 'update') {
    const k = Number(op.kcal);
    const p = op.protein === '' || op.protein == null ? '' : Number(op.protein);
    if (op.kcal === '' || !isFinite(k) || k < 0) throw new Error('Enter the calories.');
    if (p !== '' && (!isFinite(p) || p < 0)) throw new Error('Protein needs to be a number.');
    const row = findId_(log, op.entryId);
    if (!row) throw new Error('That entry is no longer in the sheet.');
    log.getRange(row, 3, 1, 2).setValues([[round1_(k), p === '' ? '' : round1_(p)]]);
    if (op.amount != null) log.getRange(row, 5).setValue(String(op.amount).slice(0, 40));
    const note = String(log.getRange(row, 6).getValue());
    if (!/edited/i.test(note)) log.getRange(row, 6).setValue(note ? note + '; edited' : 'Edited');
  } else if (op.type === 'delete') {
    const row = findId_(log, op.entryId);
    if (row) log.deleteRow(row); // already gone counts as done
  } else if (op.type === 'weight') {
    const kg = op.kg === '' || op.kg == null ? '' : Number(op.kg);
    if (kg !== '' && !(kg > 20 && kg < 400)) throw new Error('Enter your weight in kg.');
    const date = dateAt_(op.date, tz);
    const sh = ss.getSheetByName(WEIGHT);
    const row = weightRow_(sh, tz, op.date);
    if (kg === '') {
      if (row) sh.deleteRow(row);
    } else if (row) {
      sh.getRange(row, 2).setValue(round1_(kg));
    } else {
      const r = sh.getLastRow() + 1;
      sh.getRange(r, 1, 1, 2).setValues([[date, round1_(kg)]]);
      sh.getRange(r, 1).setNumberFormat('dd/mm/yyyy');
      sh.getRange(r, 2).setNumberFormat('0.0');
    }
  } else if (op.type === 'reorder') {
    reorderFoods_(ss, op.names || []);
  } else {
    throw new Error('Unknown change');
  }
}

function cleanItem_(it) {
  const food = String(it.food || '').trim();
  const per = ['100g', 'piece', 'serve'].indexOf(it.per) === -1 ? 'serve' : it.per;
  const kcalPer = Number(it.kcalPer);
  const proteinPer = it.proteinPer === '' || it.proteinPer == null ? '' : Number(it.proteinPer);
  const amount = Number(it.amount);
  if (!food) throw new Error('Enter a food.');
  if (it.kcalPer === '' || !isFinite(kcalPer) || kcalPer < 0) throw new Error('Enter calories for ' + food + '.');
  if (proteinPer !== '' && (!isFinite(proteinPer) || proteinPer < 0)) throw new Error('Protein for ' + food + ' needs to be a number.');
  if (per === '100g' && !(amount > 0)) throw new Error('Enter the grams for ' + food + '.');
  return { food: food, per: per, kcalPer: kcalPer, proteinPer: proteinPer, amount: amount > 0 ? amount : 1, save: !!it.save,
           note: String(it.note || '').slice(0, 300),
           // set for quick calories and repeats of adjusted entries, which log their totals as they are
           amountText: typeof it.amountText === 'string' ? it.amountText.slice(0, 40) : null };
}

function weightRow_(sh, tz, ymd) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const dates = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = dates.length - 1; i >= 0; i--) {
    if (dates[i][0] instanceof Date && Utilities.formatDate(dates[i][0], tz, 'yyyy-MM-dd') === ymd) return i + 2;
  }
  return 0;
}

// The named foods trade places so they read in the given order; every other row stays where it is.
// Column H (Times logged) is one ARRAYFORMULA, so only A:G and I move.
function reorderFoods_(ss, names) {
  const sh = ss.getSheetByName(FOODS);
  const rng = sh.getRange(2, 1, sh.getMaxRows() - 1, 9);
  const vals = rng.getValues(), bgs = rng.getBackgrounds();
  const at = {};
  vals.forEach((r, i) => { const n = String(r[0]).trim().toLowerCase(); if (n && !(n in at)) at[n] = i; });
  const order = [];
  names.forEach(s => { const n = String(s).trim().toLowerCase(); if (n in at && order.indexOf(at[n]) === -1) order.push(at[n]); });
  const slots = order.slice().sort((a, b) => a - b);
  if (!order.length || order.every((src, k) => src === slots[k])) return;
  const first = slots[0], last = slots[slots.length - 1];
  const block = vals.slice(first, last + 1).map(r => r.slice()), blockBg = bgs.slice(first, last + 1).map(r => r.slice());
  slots.forEach((slot, k) => { block[slot - first] = vals[order[k]].slice(); blockBg[slot - first] = bgs[order[k]].slice(); });
  const rows = last - first + 1;
  sh.getRange(first + 2, 1, rows, 7).setValues(block.map(r => r.slice(0, 7)));
  sh.getRange(first + 2, 1, rows, 7).setBackgrounds(blockBg.map(r => r.slice(0, 7)));
  sh.getRange(first + 2, 9, rows, 1).setValues(block.map(r => [r[8]]));
  sh.getRange(first + 2, 9, rows, 1).setBackgrounds(blockBg.map(r => [r[8]]));
}

// One day's entries (they can sit anywhere near the bottom: entries for other days get added out of order)
function apiState_(ss, log, tz, ymd) {
  const last = log.getLastRow();
  const start = Math.max(2, last - 599);
  const vals = last >= 2 ? log.getRange(start, 1, last - start + 1, ID_COL).getValues() : [];
  const entries = [];
  vals.forEach((r, i) => {
    if (!(r[0] instanceof Date) || Utilities.formatDate(r[0], tz, 'yyyy-MM-dd') !== ymd || !String(r[1]).trim()) return;
    let id = String(r[6]);
    if (!id) { // typed into the sheet or logged from the old page
      id = 'r' + Utilities.getUuid().replace(/-/g, '').slice(0, 10);
      log.getRange(start + i, ID_COL).setValue(id);
    }
    entries.push({ id: id, food: String(r[1]), kcal: num_(r[2]), protein: r[3] === '' ? null : num_(r[3]),
                   amount: String(r[4]), note: String(r[5]) });
  });
  const set = ss.getSheetByName(SETTINGS);
  return {
    date: ymd,
    goalKcal: num_(set.getRange('B2').getValue()) || 2400,
    goalProtein: num_(set.getRange('B3').getValue()) || 110,
    entries: entries,
    foods: foodList_(ss),
    recent: recent_(vals, start, tz),
    days: dailyTotals_(ss, tz),
    weights: weights_(ss, tz),
    sheetUrl: ss.getUrl(),
  };
}

// What you've logged in the last two weeks, newest first, one per distinct food/meal and amounts
// (entries logged together, like a meal's parts, stay together)
function recent_(vals, start, tz) {
  const since = Utilities.formatDate(new Date(Date.now() - 14 * 86400000), tz, 'yyyy-MM-dd');
  const groups = [], byKey = {};
  for (let i = vals.length - 1; i >= 0 && groups.length < 80; i--) {
    const r = vals[i];
    if (!(r[0] instanceof Date) || !String(r[1]).trim()) continue;
    const date = Utilities.formatDate(r[0], tz, 'yyyy-MM-dd');
    if (date < since) continue;
    const id = String(r[6]), m = /^(.+)-\d+$/.exec(id);
    const key = m ? m[1] : 'row' + (start + i);
    if (!byKey[key]) { byKey[key] = { date: date, items: [] }; groups.push(byKey[key]); }
    byKey[key].items.unshift({ food: String(r[1]), kcal: num_(r[2]), protein: r[3] === '' ? null : num_(r[3]),
                               amount: String(r[4]), note: String(r[5]) });
  }
  const seen = {}, out = [];
  groups.forEach(g => {
    const sig = JSON.stringify(g.items.map(it => [it.food.toLowerCase(), it.amount]));
    if (seen[sig] || out.length >= 12) return;
    seen[sig] = 1;
    out.push(g);
  });
  return out;
}

// Daily totals (from the Daily tab) for the last few weeks
function dailyTotals_(ss, tz) {
  const sh = ss.getSheetByName(DAILY);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const start = Math.max(2, last - 44);
  return sh.getRange(start, 1, last - start + 1, 3).getValues()
    .filter(r => r[0] instanceof Date)
    .map(r => ({ date: Utilities.formatDate(r[0], tz, 'yyyy-MM-dd'), kcal: num_(r[1]), protein: num_(r[2]) }));
}

// Weigh-ins for the last few weeks, oldest first
function weights_(ss, tz) {
  const sh = ss.getSheetByName(WEIGHT);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const since = Utilities.formatDate(new Date(Date.now() - 45 * 86400000), tz, 'yyyy-MM-dd');
  return sh.getRange(2, 1, last - 1, 2).getValues()
    .filter(r => r[0] instanceof Date && r[1] !== '' && isFinite(Number(r[1])))
    .map(r => ({ date: Utilities.formatDate(r[0], tz, 'yyyy-MM-dd'), kg: Number(r[1]) }))
    .filter(w => w.date >= since)
    .sort((a, b) => a.date < b.date ? -1 : a.date > b.date ? 1 : 0);
}

function findId_(log, id) {
  if (!id || log.getLastRow() < 2) return 0;
  const hit = log.getRange(2, ID_COL, log.getLastRow() - 1, 1).createTextFinder(String(id)).matchEntireCell(true).findNext();
  return hit ? hit.getRow() : 0;
}

function ensureIdColumn_(log) {
  if (log.getMaxColumns() < ID_COL) log.insertColumnsAfter(log.getMaxColumns(), ID_COL - log.getMaxColumns());
  if (log.getRange(1, ID_COL).getValue() !== 'ID') {
    log.getRange(1, ID_COL).setValue('ID');
    header_(log.getRange(1, ID_COL));
    log.hideColumns(ID_COL);
  }
}

/* ---------- the key (shown only to you, from the "Only myself" web app) ---------- */

function apiKey_(fresh) {
  const props = PropertiesService.getScriptProperties();
  let key = props.getProperty('API_KEY');
  if (!key || fresh) {
    key = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '').slice(0, 40);
    props.setProperty('API_KEY', key);
  }
  return key;
}

function pairPage_(fresh) {
  const key = apiKey_(fresh);
  const html = `<!doctype html><html><head><meta charset="utf-8"><base target="_top">
<style>
body{margin:0;font:16px/1.45 system-ui,-apple-system,"Segoe UI",Roboto,sans-serif;background:#f6f5f1;color:#1d1c1a}
@media (prefers-color-scheme:dark){body{background:#141413;color:#f0efec}input{background:#21211e!important;color:#f0efec!important}}
main{max-width:520px;margin:0 auto;padding:20px 16px}
input{width:100%;box-sizing:border-box;font:15px ui-monospace,monospace;padding:12px;border:1px solid #c9c7bf;border-radius:10px;background:#fff;color:#1d1c1a}
button{width:100%;margin-top:12px;padding:13px;border:0;border-radius:12px;background:#2a78d6;color:#fff;font:inherit;font-weight:600}
p{color:#6b6a64;font-size:14px}
</style></head><body><main>
<h2>Your Food log key</h2>
<p>Copy this and paste it into the Food log app. Keep it private: anyone with it can add to your log.</p>
<input id="k" readonly value="${key}" onclick="this.select()">
<button id="c" type="button">Copy key</button>
<p id="m"></p>
<p>If the key ever gets out, open this page with <b>?pair=new</b> instead of ?pair=1 to make a new one. The old key stops working straight away.</p>
<script>
document.getElementById('c').onclick = function () {
  var k = document.getElementById('k'), m = document.getElementById('m');
  k.select(); k.setSelectionRange(0, 99);
  var ok = false;
  try { ok = document.execCommand('copy'); } catch (e) {}
  if (navigator.clipboard) navigator.clipboard.writeText(k.value).then(function () { m.textContent = 'Copied.'; }, function () {
    m.textContent = ok ? 'Copied.' : 'Press and hold the key to copy it.';
  });
  else m.textContent = ok ? 'Copied.' : 'Press and hold the key to copy it.';
};
</script></main></body></html>`;
  return HtmlService.createHtmlOutput(html).setTitle('Food log key')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
