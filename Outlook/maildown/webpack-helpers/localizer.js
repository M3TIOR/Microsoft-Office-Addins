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
//...

// Standard Library imports
const fs = require('fs');
const path = require('path');


// Globals
var supported_locales = [];
var	plugins = [];
var rules = [];


// Exports
var exports = module.exports = {
	supported_locales: supported_locales,
	plugins: plugins,
	rules: rules,
	registerLocale: (locale) => supported_locales.push(locale),
	stage: (file, options) => {
		/* Load all our localization text into memory before we loop over it. */
		supported_locales.map((key) => {
			let file = path.join(".","locales",key+".json");
			let contents = fs.readFileSync(file,"utf-8");
			return Object.assign(JSON.parse(contents), { name: key });

		}).forEach( (localization) => {
			let outputName = path.basename(file);
			let outputPath = path.dirname(file);
			let output = "locales/"+localization.name+"/"+outputPath+"/"+outputName;

			if (outputPath == ".") file = outputName;

			plugins.push(
				/* The HtmlWebpackPlugin will be responsible for controlling the
				 * extraction of our file data and output redirection.
				 */
				new HtmlWebpackPlugin({
					filename: output, // where we put it.
					template: file // where we load it from
				})
			);

			rules.push({
				/* Search for the filename explicitly since our loader is responsible
				 * for passing file specific locale data to our templates.
				 */
				test: new RegExp(file),
				exclude: /node_modules/,
				use: [
					{
						loader: 'render-template-loader',
						options: {
							engine: 'handlebars', /* favorite syntax & we can add helpers */
							locals: Object.assign(
								options.template, /* webpack specific data.*/
								localization[file], /* file localization data. */
								{
									/* And then the universal constants */
									currentLocale: localization.name,
									outputPath: outputPath
								}
							),
							init: function( engine, info ){
								engine.registerHelper('rpath', (target) => {
									/* Handlebars Helper: rpath
									 * ---------------------------------------------------------
									 * establishes a relative connection from the template file
									 * to an external resource by tracking the original path
									 * and changing it to accomidate for any change relative to
									 * the resource being imported.
									 */
									let location = path.relative(output, target);

									return engine.SafeString(location);
								});
							}
						}
					}
				]
			})
		})
	}
};
