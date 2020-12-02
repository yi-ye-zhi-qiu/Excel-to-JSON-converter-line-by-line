/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
/*jshint browser:true */
/*global XLSX */
/* use webpack, need build tool for JS in order to use require for client-side code*/
/* need to use webpack to load FileSaver, JSZip and run client-side without script tags in html */

var X = typeof require !== "undefined" && require('../node_modules/xlsx') || XLSX;

var JSZip = require('../node_modules/jszip/dist/jszip.min.js')
var FileSaver = require('../node_modules/file-saver/dist/FileSaver.js')

var global_wb;

var process_wb = (function() {
	
	var OUT = document.getElementById('out');
	var HTMLOUT = document.getElementById('htmlout');

	var get_format = (function() {
		var radios = document.getElementsByName( "format" );
		return function() {
			for(var i = 0; i < radios.length; ++i) if(radios[i].checked || radios.length === 1) return radios[i].value;
		};
	})();

	var to_json = function to_json(workbook) {
		var result = {};
		workbook.SheetNames.forEach(function(sheetName) {
			var roa = X.utils.sheet_to_json(workbook.Sheets[sheetName]);
			if(roa.length) result[sheetName] = roa;
		});
		return JSON.stringify(result, 2, 2);
	};

	var to_json_linebyline = function to_json_linebyline(wb){
	    //NOTICE: must be named Sheet1...will find a way around that soon enough.
	    var zip = new JSZip();
	    var sheet = wb.Sheets['Sheet1'];
	    var results = [];
	    var range = X.utils.decode_range(sheet['!ref']);
	    for(let rowNum = (range.s.r+1); rowNum <= range.e.r--; rowNum++){
	       let thisRow = {},
	           thisNode = '';
	       for(let colNum=range.s.c; colNum<=range.e.c; colNum++){
	          var thisHeader = sheet[X.utils.encode_cell({r: 0, c: colNum})].w
	          var thisCell = sheet[X.utils.encode_cell({r: rowNum, c: colNum})].w
	          if(colNum === 0){
	            thisNode = thisCell;
	          }
	          thisRow[thisHeader] = thisCell;
	       }
	       thisResult = {};
	       thisResult[thisNode] = [thisRow]
	       results.push(thisResult)

	       if(rowNum === 1){
		   console.log('The point of this program is to return a zip folder full of JSON-ified excel data rows, as an example, the JSON of row ' + rowNum + ' will show as: \n' + JSON.stringify(thisResult, 2, 2));
	       }
		   zip.file(''.concat(rowNum) + '.json', JSON.stringify(thisResult, 2, 2))
	       }

		zip.generateAsync({type: "blob"})
		  .then(function(content) {
			saveAs(content, "result_" + new Date() + ".zip");
		   });
		return 'full unsplit JSON: \n' + JSON.stringify(results, 2, 2)
	}

	return function process_wb(wb) {
		global_wb = wb;
		var output = "";
		var domlinebyline = document.getElementById("uselinebyline");
		if(domlinebyline.checked == false){
			console.log("using to_json...");
			output = to_json(wb);
		}
		if(domlinebyline.checked == true) {
			console.log("using to_json_linebyline...")
			output = to_json_linebyline(wb);
		}
		if(OUT.innerText === undefined) OUT.textContent = output;
		else OUT.innerText = output;
		if(typeof console !== 'undefined') console.log("output ", new Date());
	}

})();

var setfmt = window.setfmt = function setfmt() { if(global_wb) process_wb(global_wb); };

var do_file = (function() {
	var rABS = typeof FileReader !== "undefined" && (FileReader.prototype||{}).readAsBinaryString;
	var domrabs = document.getElementsByName("userabs")[0];
	if(!rABS) domrabs.disabled = !(domrabs.checked = false);

	var use_worker = typeof Worker !== 'undefined';
	var domwork = document.getElementsByName("useworker")[0];
	if(!use_worker) domwork.disabled = !(domwork.checked = false);

	var xw = function xw(data, cb) {
		var worker = new Worker(XW.worker);
		worker.onmessage = function(e) {
			switch(e.data.t) {
				case 'ready': break;
				case 'e': console.error(e.data.d); break;
				case XW.msg: cb(JSON.parse(e.data.d)); break;
			}
		};
		worker.postMessage({d:data,b:rABS?'binary':'array'});
	};

	return function do_file(files) {
		rABS = domrabs.checked;
		use_worker = domwork.checked;
		var f = files[0];
		var reader = new FileReader();
		reader.onload = function(e) {
			if(typeof console !== 'undefined') console.log("onload", new Date(), rABS, use_worker);
			var data = e.target.result;
			if(!rABS) data = new Uint8Array(data);
			if(use_worker) xw(data, process_wb);
			else process_wb(X.read(data, {type: rABS ? 'binary' : 'array'}));
		};
		if(rABS) reader.readAsBinaryString(f);
		else reader.readAsArrayBuffer(f);
	};
})();

(function() {
	var drop = document.getElementById('drop');
	if(!drop.addEventListener) return;

	function handleDrop(e) {
		e.stopPropagation();
		e.preventDefault();
		do_file(e.dataTransfer.files);
	}

	function handleDragover(e) {
		e.stopPropagation();
		e.preventDefault();
		e.dataTransfer.dropEffect = 'copy';
	}

	drop.addEventListener('dragenter', handleDragover, false);
	drop.addEventListener('dragover', handleDragover, false);
	drop.addEventListener('drop', handleDrop, false);
})();

(function() {
	var xlf = document.getElementById('xlf');
	if(!xlf.addEventListener) return;
	function handleFile(e) { do_file(e.target.files); }
	xlf.addEventListener('change', handleFile, false);
})();
