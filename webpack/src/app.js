/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
/*jshint browser:true */
/*global XLSX */
/* use webpack, need build tool for JS in order to use require for client-side code (users do not need packages) & to write files from browser*/

var X = typeof require !== "undefined" && require('../node_modules/xlsx') || XLSX;
//require('../node_modules/jszip/dist/jszip.min.js')

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
			var zip = new JSZip();
			var sheet = wb.Sheets['Sheet1'];
			var range = X.utils.decode_range(sheet['!ref']);
			for(let rowNum = range.s.r+3; rowNum <= range.e.r--; rowNum++){
				 var results = {};
				 let thisRow = {},
						 thisNode = '';
				 for(let colNum=range.s.c; colNum<=range.e.c; colNum++){
						let eachCol = {},
								label = 'label',
								value = 'value';
						var sub_header = sheet[X.utils.encode_cell({r: 1, c: colNum})].w
						var sub_sub_header = sheet[X.utils.encode_cell({r: 2, c: colNum})].w
						var thisCell = sheet[X.utils.encode_cell({r: rowNum, c: colNum})].w
						eachCol[label] = sub_sub_header;
						eachCol[value] = thisCell;
						var this_header = sheet[X.utils.encode_cell({r: 0, c: colNum})].w
						if(colNum != 0){
							var previous_header = sheet[X.utils.encode_cell({r: 0, c: colNum-1})].w
							if(this_header != previous_header){
								new_object = {}
								entries_for_new = {}
								entries_for_new[sub_header] = eachCol
								const wantbrackets = ['row1needbrackets?']
								if(wantbrackets.some(needbrackets => needbrackets === this_header) === false){
									//if wantbrackers === false... create new_object as {entries_for_new}
									new_object[this_header] = entries_for_new
									Object.assign(results, new_object)
								} else {
									//if wantbrackets === true... create new_object as [{entries_for_new}]
									new_object[this_header] = [entries_for_new]
									Object.assign(results, new_object)
								}
							} else { //if this_header === previous_header, just add entries_for_new to the same new_object
								new_object = entries_for_new
								entries_for_new[sub_header] = eachCol
							}
						}	else {
							  //for column index-0, colNum-1 will return a negative value, so we have to hard-code c: 0
								var previous_header = sheet[X.utils.encode_cell({r: 0, c: 0})].w
								new_object = {}
								entries_for_new = {}
								entries_for_new[sub_header] = eachCol
								Object.assign(entries_for_new, hard_coded_general_variables)
								new_object[this_header] = entries_for_new
								Object.assign(results, new_object)
						}
				 }
				zip.file(''.concat(rowNum) + '.json', JSON.stringify(results, 2, 2))
				if(rowNum === 3){
					console.log('The point of this program is to return a zip folder full of JSON-ified excel data rows, as an example, the JSON of row ' + rowNum + ' will show as: \n' + JSON.stringify(results, 2, 2));
				}
			}

			zip.generateAsync({type: "blob"})
			.then(function(content) {
				saveAs(content, "result_" + new Date() + ".zip");
			});

			return 'example output JSON (will always be last row of data in input Excel): \n' + JSON.stringify(results, 2, 2)
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

const DARK_MODE = 'dark';
const LIGHT_MODE = 'light';
const THEME = 'mode';

document.addEventListener(
  'DOMContentLoaded', (event) => {
    applyTheme();
    const toggleSwitch = document.getElementById('toggle-switch');
    toggleSwitch.onclick = function() {
      let currentMode = localStorage.getItem(THEME);
      localStorage.setItem(
        THEME,
        currentMode === DARK_MODE ? LIGHT_MODE : DARK_MODE
      );
      applyTheme();
    }
  }
);

function applyTheme() {
  let html = document.documentElement;
  let currentMode = localStorage.getItem(THEME);
  if (currentMode === DARK_MODE) {
    html.classList.add(DARK_MODE);
    document.getElementById('toggle-switch').innerHTML =
      '<i class="fas fa-sun"></i>';
  }
  else {
    html.classList.remove(DARK_MODE);
    document.getElementById('toggle-switch').innerHTML =
      '<i class="fas fa-moon"></i>';
  }
}
