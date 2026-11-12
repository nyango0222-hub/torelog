/**
 * トレログ用 GAS（スプレッドシートに紐づけて使う）
 *
 * 1. スプレッドシートを新規作成 →「拡張機能」→「Apps Script」にこのコードを貼る
 * 2. 下の TOKEN を自分だけの合言葉に書き換える
 * 3. 「デプロイ」→「新しいデプロイ」→ 種類：ウェブアプリ
 *    実行ユーザー：自分 ／ アクセスできるユーザー：全員
 * 4. 発行された …/exec のURLと合言葉を、アプリの「設定」に入れる
 * ※コードを直したら「デプロイを管理」→ 鉛筆 → 新バージョン で更新する
 *
 * シート
 *  「記録」… 1行＝1セット（Museや自分が読む用）
 *  「日別」… 1行＝1日（集計と、復元用のJSON）
 *
 * 読み取りAPI（Muse用・GET）
 *  ?token=合言葉&action=summary
 *  ?token=合言葉&action=days&from=2026-11-01&to=2026-11-30
 */
const TOKEN = 'ここを自分の合言葉に変える';

const SH_LOG = '記録';
const SH_DAY = '日別';
const HEAD_LOG = ['日付', '部位', '種目', 'セット', '重さkg', '回数', '推定RM', '有酸素分'];
const HEAD_DAY = ['日付', 'サウナ', '有酸素分', '種目数', 'セット数', 'レップ数', '負荷量kg', 'JSON'];

function doPost(e) {
  try {
    const p = JSON.parse(e.postData.contents);
    if (p.token !== TOKEN) return out_({ ok: false, error: '合言葉が違います' });
    if (p.action === 'ping') return out_({ ok: true });
    if (p.action === 'save') return out_(save_(p.days || []));
    if (p.action === 'all') return out_({ ok: true, days: readDays_() });
    return out_({ ok: false, error: 'actionが不明です' });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

function doGet(e) {
  try {
    const q = (e && e.parameter) || {};
    if (q.token !== TOKEN) return out_({ ok: false, error: '合言葉が違います' });
    if (q.action === 'days') return out_({ ok: true, days: readDays_(q.from, q.to) });
    if (q.action === 'summary') return out_(summary_());
    return out_({ ok: true, actions: ['summary', 'days'] });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function rm_(w, r) {
  if (w <= 0 || r <= 0) return 0;
  return r === 1 ? w : Math.round(w * (1 + r / 40) * 100) / 100;
}

function sheet_(name, head) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('@'); // 日付を文字のまま保つ
  }
  return sh;
}

function dateStr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v);
}

// 指定した日付の行を入れ替える（同じ日を送り直しても重複しない）
function replace_(sh, dates, rows, width) {
  const last = sh.getLastRow();
  let keep = [];
  if (last > 1) {
    const rng = sh.getRange(2, 1, last - 1, width);
    keep = rng.getValues().filter(function (r) { return !dates[dateStr_(r[0])]; });
    rng.clearContent();
  }
  const all = keep.concat(rows).map(function (r) { r[0] = dateStr_(r[0]); return r; });
  all.sort(function (a, b) { return a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0; });
  if (all.length) sh.getRange(2, 1, all.length, width).setValues(all);
}

function save_(days) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const log = sheet_(SH_LOG, HEAD_LOG);
    const dsh = sheet_(SH_DAY, HEAD_DAY);
    const dates = {}, logRows = [], dayRows = [];
    days.forEach(function (d) {
      if (!d || !d.date) return;
      dates[d.date] = true;
      if (d.deleted) return;
      let sets = 0, reps = 0, vol = 0, min = 0;
      (d.items || []).forEach(function (it) {
        if (it.min !== undefined) {
          const m = Number(it.min) || 0;
          if (it.part !== '自重トレーニング') min += m;
          logRows.push([d.date, it.part, it.name, '', '', '', '', m]);
        } else {
          (it.sets || []).forEach(function (s, j) {
            const w = Number(s.w) || 0, r = Number(s.r) || 0;
            if (r <= 0) return;
            sets++; reps += r; vol += w * r;
            logRows.push([d.date, it.part, it.name, j + 1, w, r, rm_(w, r), '']);
          });
        }
      });
      dayRows.push([d.date, d.sauna ? '○' : '', min, (d.items || []).length, sets, reps, vol, JSON.stringify(d)]);
    });
    replace_(log, dates, logRows, HEAD_LOG.length);
    replace_(dsh, dates, dayRows, HEAD_DAY.length);
    return { ok: true, saved: Object.keys(dates).length };
  } finally {
    lock.releaseLock();
  }
}

function readDays_(from, to) {
  const sh = sheet_(SH_DAY, HEAD_DAY);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const out = [];
  sh.getRange(2, 1, last - 1, HEAD_DAY.length).getValues().forEach(function (r) {
    const d = dateStr_(r[0]);
    if (!d || (from && d < from) || (to && d > to)) return;
    try { out.push(JSON.parse(r[7])); } catch (err) {}
  });
  return out;
}

function summary_() {
  const days = readDays_();
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const back = function (n) { return Utilities.formatDate(new Date(now.getTime() - n * 86400000), tz, 'yyyy-MM-dd'); };
  const d7 = back(6), d28 = back(27);
  const s = { ok: true, today: back(0), days_total: days.length,
    last7: { days: 0, volume_kg: 0, cardio_min: 0, bodyweight_min: 0, sauna: 0 },
    last28: { days: 0, volume_kg: 0, cardio_min: 0, bodyweight_min: 0, sauna: 0 },
    exercises: [] };
  const ex = {};
  days.forEach(function (d) {
    let vol = 0, min = 0, bw = 0;
    (d.items || []).forEach(function (it) {
      const key = it.part + '｜' + it.name;
      const e = ex[key] || (ex[key] = { part: it.part, name: it.name, last_date: '', times: 0, best_rm_kg: 0, best_rm_date: '', last: '' });
      e.times++;
      if (it.min !== undefined) {
        if (it.part === '自重トレーニング') bw += Number(it.min) || 0; else min += Number(it.min) || 0;
        if (d.date >= e.last_date) { e.last_date = d.date; e.last = (Number(it.min) || 0) + '分'; }
      } else {
        const txt = [];
        (it.sets || []).forEach(function (x) {
          const w = Number(x.w) || 0, r = Number(x.r) || 0;
          if (r <= 0) return;
          vol += w * r; txt.push(w + 'kg×' + r);
          const v = rm_(w, r);
          if (v > e.best_rm_kg) { e.best_rm_kg = v; e.best_rm_date = d.date; }
        });
        if (d.date >= e.last_date) { e.last_date = d.date; e.last = txt.join(', '); }
      }
    });
    [[d7, s.last7], [d28, s.last28]].forEach(function (p) {
      if (d.date >= p[0]) { p[1].days++; p[1].volume_kg += vol; p[1].cardio_min += min; p[1].bodyweight_min += bw; if (d.sauna) p[1].sauna++; }
    });
  });
  s.exercises = Object.keys(ex).map(function (k) { return ex[k]; });
  return s;
}
