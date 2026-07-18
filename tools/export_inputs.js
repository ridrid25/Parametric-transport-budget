var SHEETS={};
function makeSheet(name){return{name:name,data:{},maxRow:0,maxCol:0,charts:[],
  getRange:function(r,c,nr,nc){nr=nr||1;nc=nc||1;var self=this;return{
    getValues:function(){var o=[];for(var i=0;i<nr;i++){var row=[];for(var j=0;j<nc;j++){var k=(r+i)+'_'+(c+j);row.push(self.data.hasOwnProperty(k)?self.data[k]:'');}o.push(row);}return o;},
    setValues:function(v){for(var i=0;i<v.length;i++)for(var j=0;j<v[i].length;j++){var k=(r+i)+'_'+(c+j);self.data[k]=v[i][j];self.maxRow=Math.max(self.maxRow,r+i);self.maxCol=Math.max(self.maxCol,c+j);}return this;},
    setValue:function(v){self.data[r+'_'+c]=v;self.maxRow=Math.max(self.maxRow,r);self.maxCol=Math.max(self.maxCol,c);return this;},
    setFontWeight:function(){return this;},setBackground:function(){return this;},setFontColor:function(){return this;},setNumberFormat:function(){return this;},setFontSize:function(){return this;}};},
  getLastRow:function(){return this.maxRow;},getLastColumn:function(){return this.maxCol;},
  clear:function(){this.data={};this.maxRow=0;this.maxCol=0;this.charts=[];},setFrozenRows:function(){return this;},setColumnWidth:function(){return this;},autoResizeColumns:function(){return this;},
  getCharts:function(){return this.charts;},removeChart:function(){},insertChart:function(){},newChart:function(){var b={setChartType:function(){return b;},addRange:function(){return b;},setOption:function(){return b;},setPosition:function(){return b;},build:function(){return {};}};return b;},
  toTable:function(nc){var o=[];for(var r=1;r<=this.maxRow;r++){var row=[];for(var c=1;c<=nc;c++)row.push(this.data[r+'_'+c]);o.push(row);}return o;}};}
var SS={getSheetByName:function(n){return SHEETS[n]||null;},insertSheet:function(n){var s=makeSheet(n);SHEETS[n]=s;return s;}};
global.SpreadsheetApp={getActiveSpreadsheet:function(){return SS;},getUi:function(){var m={createMenu:function(){return m;},addItem:function(){return m;},addSeparator:function(){return m;},addToUi:function(){return m;},alert:function(){}};return m;}};
global.Charts={ChartType:{LINE:'LINE',COLUMN:'COLUMN',PIE:'PIE'}};
var fs=require('fs');
eval(fs.readFileSync(__dirname+'/../apps-script/Code.gs','utf8'));
generateSampleData();
buildBudget();
buildTransportBreakdown();
var out={
  Vehicles:SHEETS['Vehicles'].toTable(11),
  Seasonality:SHEETS['Seasonality'].toTable(4),
  Params:SHEETS['Params'].toTable(3),
  Customers:SHEETS['Customers'].toTable(10),
  Budget:SHEETS['Бюджет'].toTable(15)
};
fs.writeFileSync(__dirname+'/out.inputs.json',JSON.stringify(out));
console.error('rows: Vehicles',out.Vehicles.length,'Customers',out.Customers.length,'Params',out.Params.length,'Budget',out.Budget.length);
