/* webpack config for app.out.js */
const path = require('path');


module.exports = {
	mode: 'development',
	devtool: 'inline-source-map',

	/* entry point app.js */
	entry: './src/app.js',

	/* write to app.out.js */
	output: {
		filename: 'app.out.js',
		path: path.resolve(__dirname, 'dist'),
	}
	/*,
	module: {
		rules: [
			{ test: /\.xlsx$/, loader: "webpack-xlsx-loader" }
		]
	}
	*/


}
