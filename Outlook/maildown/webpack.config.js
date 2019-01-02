/*
 * Microsoft Office Add-ins: An add-in collection for the Microsoft Office Suite.
 * Copyright (C) 2018, Ruby Allison Rose (aka: M3TIOR)
 *
 * Original templates used are copyright of Microsoft 2017 under the MIT license.
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301
 * USA
 */

// Current Library Requirements
const HtmlWebpackPlugin = require('html-webpack-plugin');

// Standard Library imports
const fs = require('fs');
const path = require('path');


supported_locales = [
	// Update this list as new locales are added
	"en-US"
]


module.exports = env => { // function(env){...}
	// Convert module export to a function so webpack can pass in our env vars.

	// Conditional mode (might make more modular later)
	if (env.mode == "development")
		websource = "https://localhost:3000";

	if (env.mode == "production")
		websource = "https://m3tior.github.io/Microsoft-Office-Addins/Outlook/maildown";

	let locale = JSON.parse(fs.readFileSync(__dirname+"/locales/en-US.json","utf-8"))

	/*
	for (var key in locales){

	}
	*/

	return {
		entry: {
			polyfill: 'babel-polyfill',
			app: './src/index.js',
			'function-file': './function-file/function-file.js'
		},
		output: {
			publicPath: '/dist/'
		},
		module: {
			rules: [
				{
					test: /\.js$/,
					exclude: /node_modules/,
					use: 'babel-loader'
				},
				{
					test: /\.html$/,
					exclude: /node_modules/,
					use: 'html-loader'
				},
				{
					test: /\.(png|jpg|jpeg|gif)$/,
					use: 'file-loader'
				},
				{
					test: /\manifest.xml/,
					exclude: /node_modules/,
					use: [
						{
							loader: 'file-loader?name=manifest.xml'
						},
						{
							loader: 'render-template-loader',
							options: {
								engine: 'handlebars',
								locals: {
									websource: websource,
									locale: locale["manifest.xml"]
								},
								engineOptions: function (info) {
									// Ejs wants the template filename for partials rendering.
									// (Configuring a "views" option can also be done.)
									return { filename: info.filename }
								}
							}
						}
					]
				}
			]
		},
		plugins: [
			new HtmlWebpackPlugin({
				template: './index.html',
				chunks: ['polyfill', 'app']
			}),
			new HtmlWebpackPlugin({
				template: './function-file/function-file.html',
				filename: 'function-file/function-file.html',
				chunks: ['function-file']
			})
		],
	};
};
