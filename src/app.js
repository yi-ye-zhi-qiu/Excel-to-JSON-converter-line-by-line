/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
/*jshint browser:true */
/*global XLSX */
/* use webpack, need build tool for JS in order to use require for client-side code*/

var X = typeof require !== "undefined" && require('../node_modules/xlsx') || XLSX;

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
		var sheet = wb.Sheets['Sheet1'];
		var result = {};
		var row, rowNum, colNum;
		var range = X.utils.decode_range(sheet['!ref']);
		for(rowNum = range.s.r; rowNum <= range.e.r-2; rowNum++){
			row = [];
			for(colNum=range.s.c; colNum<=range.e.c; colNum++){
			   var nextCell = sheet[
			      X.utils.encode_cell({r: rowNum, c: colNum})
			   ];
			   if( typeof nextCell === 'undefined' ){
			      row.push(void 0);
			   } else row.push(nextCell.w);
			}
			result[nextCell.v] = row;
		    }
		    return JSON.stringify(result, 2, 2);
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

		/*var datastr="data:text/json;charset=utf-8," + encodeURIComponent(output);
		var exportName = 'hello'
		var downloadAnchorNode = document.createElement('a');
		downloadAnchorNode.setAttribute("href",     datastr);
		downloadAnchorNode.setAttribute("download", exportName + ".json");
		document.body.appendChild(downloadAnchorNode); // required for firefox
		downloadAnchorNode.click();
		downloadAnchorNode.remove();*/
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
