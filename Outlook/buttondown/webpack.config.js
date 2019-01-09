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

// External Library Requirements
const HtmlWebpackPlugin = require('html-webpack-plugin');

// Internal Library Requirements
const localizer = require('./webpack-helpers/localizer.js')

// Standard Library imports
//...


// Convert module export to a function so webpack can pass in our env vars.
module.exports = env => { // function(env){...}
	let websource = "" // the location where we'll be serving our content from

	// Conditional mode (might make more modular later)
	if (env.mode == "development") websource="https://localhost:3000";
	if (env.mode == "production") websource="https://m3tior.github.io/Microsoft-Office-Addins/Outlook/maildown";

	// Update this list as new locales are added
	localizer.registerLocale("en-US");

	/* Add localized files here */
	localizer.stage("./manifest.xml",{
		template: {
			websource: websource,
			locales: localizer.supported_locales
		}
	})

	return {
		//stats: 'verbose',
		mode: env.mode,
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
				...localizer.rules
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
			}),
			...localizer.plugins
		],
	};
};
