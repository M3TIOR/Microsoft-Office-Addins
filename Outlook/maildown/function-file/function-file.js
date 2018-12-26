/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
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
