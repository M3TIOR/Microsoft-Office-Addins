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

// The initialize function must be run each time a new page is loaded
Office.initialize = reason => {

};

// Add any ui-less function here
function convertMarkdown(event){
    // This code is ran on a button press only. (I hope)
    var item = Office.context.mailbox.item;
    var content = item.body;

	// Is enum, no instanceof (*u*)
	if (item.itemType == Office.MailboxEnums.ItemType.Message){

		// Load Showdown and make it spicy!
		var converter = new showdown.Converter();
		//converter.setFlavor('github');

		content.getAsync(Office.CoercionType.Text, {}, (result) => {
			if (result.status == "succeeded"){
				let html = converter.makeHtml(result.value);

				content.setAsync(html, {coercionType:"html", }, (result) => {
					// We have to have some kind of error handeling.
					if(result.status != "succeeded")
						console.log(result);
				});
			}
			else {
				console.log(result);
			}

		});
	}
}
