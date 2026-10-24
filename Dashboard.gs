/**
 * ダッシュボード：月別総額・講師別推移。「合計」シートの値から作り直す（script所有・値のみ）。
 * グラフ生成（Chartsサービス）が使えないテスト環境でも本処理が止まらないよう、try/catchで囲う。
 */
function refreshDashboard_() {
  try {
    buildDashboard_();
  } catch (err) {
    console.error('ダッシュボード更新に失敗しました（本処理は継続します）: ' + err);
  }
}

function buildDashboard_() {
  var summarySheet = getActiveSs_().getSheetByName(SHEET_SUMMARY);
  var rows = summarySheet ? getRowsBelowHeader_(summarySheet) : [];

  var sheet = getOrResetOwnedSheet_(SHEET_DASHBOARD, ['年月', '総額']);
  var byMonth = {};
  var byInstructorMonth = {};
  rows.forEach(function (r) {
    var month = String(r[COL_SUMMARY.MONTH - 1]);
    var instructorId = String(r[COL_SUMMARY.INSTRUCTOR_ID - 1]);
    var total = toNumber_(r[COL_SUMMARY.TOTAL - 1]);
    byMonth[month] = (byMonth[month] || 0) + total;
    byInstructorMonth[instructorId] = byInstructorMonth[instructorId] || {};
    byInstructorMonth[instructorId][month] = total;
  });

  var months = Object.keys(byMonth).sort();
  var monthRows = months.map(function (m) { return [formatMonthLabel_(m), byMonth[m]]; });
  if (monthRows.length) writeRows_(sheet, 2, 1, monthRows);

  var instructorIds = Object.keys(byInstructorMonth).sort();
  var tableStartCol = 4;
  var header = ['講師ID'].concat(months.map(formatMonthLabel_));
  sheet.getRange(1, tableStartCol, 1, header.length).setValues([header]);
  if (instructorIds.length) {
    var instructorRows = instructorIds.map(function (id) {
      return [id].concat(months.map(function (m) { return byInstructorMonth[id][m] || 0; }));
    });
    sheet.getRange(2, tableStartCol, instructorRows.length, header.length).setValues(instructorRows);
  }

  insertOrUpdateDashboardChart_(sheet, monthRows.length);
}

function insertOrUpdateDashboardChart_(sheet, numMonthRows) {
  sheet.getCharts().forEach(function (c) { sheet.removeChart(c); });
  if (numMonthRows < 1) return;
  var chartRange = sheet.getRange(1, 1, numMonthRows + 1, 2);
  var chart = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(chartRange)
    .setPosition(1, 5, 0, 0)
    .setOption('title', '月別総額')
    .build();
  sheet.insertChart(chart);
}
