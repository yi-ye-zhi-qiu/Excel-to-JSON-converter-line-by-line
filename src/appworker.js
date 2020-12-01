/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
const XLSX = require('../node_modules/xlsx')

//path to xlsx: C:\Users\立安\webpack-demo\node_modules\xlsx

postMessage({t:"ready"});

onmessage = function (evt) {
  var v;
  try {
    v = XLSX.read(evt.data.d, {type: evt.data.b});
postMessage({t:"xlsx", d:JSON.stringify(v)});
  } catch(e) { postMessage({t:"e",d:e.stack||e}); }
};
