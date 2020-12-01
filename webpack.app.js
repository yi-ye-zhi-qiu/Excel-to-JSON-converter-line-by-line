/* xlsx.js (C) 2013-present  SheetJS -- http://sheetjs.com */
var commonops = {
	/* suppress node shims */
	node: {
		process: false,
		Buffer: false
	}
}

/* app.out.js */
var app_config = Object.assign({
	entry: './src/app.js',
	output: { path:__dirname, filename: './dist/app.out.js' }
}, commonops);

/* appworker.out.js */
var appworker_config = Object.assign({
	entry: './src/appworker.js',
	output: { path:__dirname, filename: './dist/appworker.out.js' }
}, commonops);

module.exports = [
	app_config,
	appworker_config
]
