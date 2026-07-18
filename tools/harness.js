// Прогоняет Code.gs через стаб Apps Script API и выгружает вычисленные данные в JSON
// для предпросмотра в чате (артефакт), без необходимости открывать Google Sheets.

function makeSheet(name) {
  return {
    name: name, data: {}, maxRow: 0, maxCol: 0, charts: [],
    clear: function () { this.data = {}; this.maxRow = 0; this.maxCol = 0; this.charts = []; },
    getCharts: function () { return this.charts; },
    removeChart: function (c) { this.charts = this.charts.filter(function (x) { return x !== c; }); },
    insertChart: function (c) { this.charts.push(c); },
    setFrozenRows: function () { return this; },
    setColumnWidth: function () { return this; },
    autoResizeColumns: function () { return this; },
    newChart: function () {
      var builder = {
        opts: {}, ranges: [],
        setChartType: function (t) { this.opts.type = t; return this; },
        addRange: function (r) { this.ranges.push(r); return this; },
        setOption: function (k, v) { this.opts[k] = v; return this; },
        setPosition: function () { return this; },
        build: function () { return this.opts; }
      };
      return builder;
    },
    getRange: function (row, col, numRows, numCols) {
      numRows = numRows || 1; numCols = numCols || 1;
      var self = this;
      return {
        getValues: function () {
          var out = [];
          for (var r = 0; r < numRows; r++) {
            var rowArr = [];
            for (var c = 0; c < numCols; c++) {
              var key = (row + r) + '_' + (col + c);
              rowArr.push(self.data.hasOwnProperty(key) ? self.data[key] : '');
            }
            out.push(rowArr);
          }
          return out;
        },
        setValues: function (vals) {
          for (var r = 0; r < vals.length; r++) {
            for (var c = 0; c < vals[r].length; c++) {
              var key = (row + r) + '_' + (col + c);
              self.data[key] = vals[r][c];
              self.maxRow = Math.max(self.maxRow, row + r);
              self.maxCol = Math.max(self.maxCol, col + c);
            }
          }
          return this;
        },
        setValue: function (v) { self.data[row + '_' + col] = v; self.maxRow = Math.max(self.maxRow, row); self.maxCol = Math.max(self.maxCol, col); return this; },
        setFontWeight: function () { return this; },
        setBackground: function () { return this; },
        setFontColor: function () { return this; },
        setNumberFormat: function () { return this; },
        setFontSize: function () { return this; }
      };
    },
    getLastRow: function () { return this.maxRow; },
    getLastColumn: function () { return this.maxCol; },
    toTable: function (nCols) {
      var out = [];
      for (var r = 1; r <= this.maxRow; r++) {
        var row = [];
        for (var c = 1; c <= nCols; c++) row.push(this.data[r + '_' + c]);
        out.push(row);
      }
      return out;
    }
  };
}

var SHEETS = {};
var SS = {
  getSheetByName: function (name) { return SHEETS[name] || null; },
  insertSheet: function (name) { var s = makeSheet(name); SHEETS[name] = s; return s; }
};

global.SpreadsheetApp = {
  getActiveSpreadsheet: function () { return SS; },
  getUi: function () {
    return { createMenu: function () { var m = { addItem: function () { return m; }, addSeparator: function () { return m; }, addToUi: function () { return m; } }; return m; },
      alert: function (msg) { console.error('[UI ALERT]', msg); } };
  }
};
global.Charts = { ChartType: { LINE: 'LINE', COLUMN: 'COLUMN', PIE: 'PIE' } };

var fs = require('fs');
var code = fs.readFileSync(__dirname + '/../apps-script/Code.gs', 'utf8');
eval(code);

generateSampleData();

// Демонстрационное значение для листа "CAPEX vs Аутсорс" — Light тир исторически сильно
// перегружен (высокая загрузка парка), так что докупка машин там показательна.
function setParam(name, value) {
  var sh = SHEETS['Params'];
  for (var r = 1; r <= sh.maxRow; r++) {
    if (sh.data[r + '_1'] === name) { sh.data[r + '_2'] = value; return; }
  }
  throw new Error('param not found: ' + name);
}
setParam('additional_vehicles_light', 3);

buildBudget();
runScenarioAnalysis();
buildTransportBreakdown();
buildDashboard();

console.error('Dashboard charts:', SHEETS['Дашборд'].charts.length);

var budget = SHEETS['Бюджет'].toTable(15);
var scenarios = SHEETS['Сценарии'].toTable(7);
var tiers = SHEETS['По видам транспорта'].toTable(14);
var regions = SHEETS['По регионам'].toTable(7);
var topCustomers = SHEETS['Топ клиентов'].toTable(7);
var bands = SHEETS['По дистанции'].toTable(7);
var segments = SHEETS['По сегментам клиентов'].toTable(8);
var capex = SHEETS['CAPEX vs Аутсорс'].toTable(7);

var out = { budget: budget, scenarios: scenarios, tiers: tiers, regions: regions,
  topCustomers: topCustomers, bands: bands, segments: segments, capex: capex };
fs.writeFileSync(__dirname + '/out.data.json', JSON.stringify(out));
console.error('Wrote tools/out.data.json');

// Sanity: bands and segments totals must reconcile with the overall annual budget revenue
var annualRevenue = budget[14][1];
var bandRevSum = bands.slice(1).reduce(function (s, r) { return s + r[2]; }, 0);
var segRevSum = segments.slice(1).reduce(function (s, r) { return s + r[2]; }, 0);
console.error('annual revenue:', annualRevenue.toFixed(0));
console.error('band sum:', bandRevSum.toFixed(0), 'diff%:', (100 * (bandRevSum - annualRevenue) / annualRevenue).toFixed(6));
console.error('segment sum:', segRevSum.toFixed(0), 'diff%:', (100 * (segRevSum - annualRevenue) / annualRevenue).toFixed(6));
console.error('bands:', JSON.stringify(bands.slice(1)));
console.error('segments:', JSON.stringify(segments.slice(1)));
console.error('tiers cost/km cols:', JSON.stringify(tiers.map(function(r){ return [r[0], r[11], r[12], r[13]]; })));
