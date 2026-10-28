// node test/show_statement.js : 架空の2026年9月分の実績で、チェック・計算・PDF・下書きをモック上で通し、結果を表示する。
// 実行日時は 2026-10-01 10:00（日本時間）に固定した想定。実際の Apps Script・Gmail・ドライブでは実行していない（README の画像はこの出力）。
var fs = require('fs'), path = require('path'), vm = require('vm');
var mockApis = require('./gas_mock/mock_apis.js');
var ROOT = path.join(__dirname, '..');
var FIXED = Date.UTC(2026, 9, 1, 1, 0, 0); // 2026-10-01 10:00 JST
var RealDate = Date;
function FixedDate() { var a = [].slice.call(arguments); return a.length ? new (Function.prototype.bind.apply(RealDate, [null].concat(a)))() : new RealDate(FIXED); }
FixedDate.prototype = RealDate.prototype; FixedDate.now = function () { return FIXED; }; FixedDate.UTC = RealDate.UTC; FixedDate.parse = RealDate.parse;
var built = mockApis.buildMockGlobals();
built.globals.Date = FixedDate;
var c = vm.createContext(built.globals);
['Config.gs', 'Util.gs', 'Sheets.gs', 'State.gs', 'Drive.gs', 'Batch.gs', 'Validate.gs', 'Calculate.gs', 'Pdf.gs', 'Draft.gs', 'Finalize.gs', 'Dashboard.gs', 'Code.gs']
  .forEach(function (f) { vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), c, { filename: f }); });
function csv(p) { return fs.readFileSync(path.join(ROOT, 'data', p), 'utf8').split('\n').map(function (l) { return l.replace(/\r$/, ''); }).filter(Boolean).map(function (l) { return l.split(','); }); }
function load(name, rows, nums) {
  var sh = built.world.spreadsheet.insertSheet(name);
  rows = rows.map(function (r, i) { return r.map(function (v, j) { return i && nums.indexOf(j) !== -1 && v !== '' && !isNaN(Number(v)) ? Number(v) : v; }); });
  sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
}
load('設定', csv('settings.csv'), []); load('マスタ_講師', csv('master_instructors.csv'), []);
load('マスタ_品目', csv('master_items.csv'), [3, 4, 5]); load('実績_202608', csv('performance_202608.csv'), [3]);
load('実績_202609', csv('performance_202609.csv'), [3]);
var ss = built.world.spreadsheet;
var s = ss.getSheetByName('設定').__dumpRows(); s.forEach(function (r, i) { if (r[0] === '対象年月') ss.getSheetByName('設定').getRange(i + 1, 2).setValue('2026-08'); });
var ck = c.runCheck(); c.runCalculate(); c.runPdf(); c.runDraft();
console.log('対象 2026-08：チェックのエラー ' + ck.errorCount + '件。続けて計算・PDF・下書きを実行');
var calc = ss.getSheetByName('計算結果_202608').__dumpRows();
console.log('\n■ 計算結果_202608（先頭5名／全' + (calc.length - 1) + '名）');
calc.slice(0, 6).forEach(function (r) { console.log(r.slice(0, 7).join(' | ')); });
var pdfs = Object.keys(built.world.files).filter(function (k) { return built.world.files[k].mimeType === 'application/pdf'; }).length;
console.log('\nPDF ' + pdfs + '件・Gmail 下書き ' + Object.keys(built.world.drafts).length + '件');
console.log('\n■ エラーログ');
(ss.getSheetByName('エラーログ').__dumpRows()).forEach(function (r) { console.log(r.join(' | ')); });
