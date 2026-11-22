/* 봉필네 데릴사위 · 순위표 (Google Apps Script)
 * 구글 시트의 [확장 프로그램 → Apps Script]에 이 내용을 통째로 붙여 넣고, 웹 앱으로 배포합니다.
 * 자세한 순서는 같은 폴더의 '설치 방법.txt'를 보세요. */

const TEACHER_CODE = '3719';      // 게임의 교사용 코드와 같게 둡니다. 바꾸려면 게임 쪽도 함께 바꿔야 해요.
const SHEET_NAME = '기록';
const PEN_WRONG = 60, PEN_MISS = 30;   // 틀린 눈치 1개당 60초, 못 찾은 조각 1개당 30초
const MIN_PLAY = 90;                   // 이보다 짧은 기록은 받지 않습니다(초)
const MAX_PLAY = 3 * 60 * 60;

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['시각', '반', '번호', '이름', '기록(초)', '플레이(초)', '틀린 눈치', '못 찾은 조각', '결말', '판 번호']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function mask_(n) {
  n = String(n);
  if (n.length <= 1) return n;
  if (n.length === 2) return n.charAt(0) + '*';
  return n.charAt(0) + '*'.repeat(n.length - 2) + n.charAt(n.length - 1);
}

/* 한 사람(반·번호·이름)당 가장 빠른 기록 하나 */
function best_() {
  const rows = sheet_().getDataRange().getValues().slice(1);
  const m = {};
  rows.forEach(function (r) {
    const c = Number(r[1]), no = Number(r[2]), n = String(r[3]).trim(), rec = Number(r[4]);
    if (!c || !n || !(rec > 0)) return;
    const k = c + '|' + no + '|' + n;
    const t = r[0] instanceof Date ? r[0].getTime() : 0;
    if (!m[k] || rec < m[k].rec || (rec === m[k].rec && t < m[k].t)) {
      m[k] = { c: c, no: no, n: n, rec: rec, play: Number(r[5]), wr: Number(r[6]), ms: Number(r[7]), end: String(r[8] || ''), t: t };
    }
  });
  return Object.keys(m).map(function (k) { return m[k]; }).sort(function (a, b) { return a.rec - b.rec || a.t - b.t; });
}

function doGet(e) {
  try {
    const code = String((e && e.parameter && e.parameter.code) || '');
    const teacher = code === TEACHER_CODE;
    const list = best_().map(function (x) {
      return teacher ? x : { c: x.c, no: x.no, n: mask_(x.n), rec: x.rec, play: x.play, wr: x.wr, ms: x.ms };
    });
    return out_({ ok: true, v: list, teacher: teacher });
  } catch (err) {
    return out_({ ok: false, e: String(err && err.message || err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const b = JSON.parse(e.postData.contents);
    if (b.a !== 'submit') throw new Error('알 수 없는 요청이에요.');
    const c = Math.round(Number(b.c)), no = Math.round(Number(b.no));
    const n = String(b.n || '').replace(/\s+/g, '').slice(0, 10);
    const play = Math.round(Number(b.play)), wr = Math.round(Number(b.wr)), ms = Math.round(Number(b.ms));
    const end = String(b.end || '').slice(0, 8), run = String(b.run || '').slice(0, 24);
    if (!(c >= 1 && c <= 20) || !(no >= 1 && no <= 60) || !n) throw new Error('반·번호·이름을 확인해 주세요.');
    if (!(play >= MIN_PLAY && play <= MAX_PLAY)) throw new Error('기록이 올바르지 않아요.');
    if (!(wr >= 0 && wr <= 9) || !(ms >= 0 && ms <= 6)) throw new Error('기록이 올바르지 않아요.');
    const rec = play + wr * PEN_WRONG + ms * PEN_MISS;

    lock.waitLock(10000);
    const sh = sheet_();
    const last = sh.getLastRow();
    let dup = false;
    if (run && last > 1) {
      const from = Math.max(2, last - 400);
      dup = sh.getRange(from, 10, last - from + 1, 1).getValues().some(function (r) { return String(r[0]) === run; });
    }
    if (!dup) sh.appendRow([new Date(), c, no, n, rec, play, wr, ms, end, run]);
    SpreadsheetApp.flush();
    lock.releaseLock();

    const list = best_();
    const i = list.findIndex(function (x) { return x.c === c && x.no === no && x.n === n; });
    const mine = list[i];
    const cls = list.filter(function (x) { return x.c === c; });
    const ci = cls.findIndex(function (x) { return x.no === no && x.n === n; });
    return out_({ ok: true, v: { rec: rec, rank: i + 1, total: list.length, cRank: ci + 1, cTotal: cls.length, best: mine && mine.rec === rec, bestRec: mine ? mine.rec : rec } });
  } catch (err) {
    try { lock.releaseLock(); } catch (e2) {}
    return out_({ ok: false, e: String(err && err.message || err) });
  }
}
