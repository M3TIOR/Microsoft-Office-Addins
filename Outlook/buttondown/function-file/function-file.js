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

// This is where we'll store the button event so it doesn't get destroyed
// before execution of the asynchronous code is completed.
/* NOTE:
 *	Note to self: please remember to invoice the Microsoft Addin Development
 *	Documentation team about, or submit a pull request for, adding the event
 *	documentation for the function-file <ExecuteFunction> manifest API to
 *	the Git The Gist add-in tutorial. And a description of why it's used.
 */
// event.completed(); // Closes the tracked event processing loop in outlook
let btnEvent = null;
let unrendered = null; // stored pre-render md string for reverting rendered md.

// Load Showdown and make it spicy!
let converter = new showdown.Converter({
	strikethrough: true,
	tables: true,
	tasklist: true,
	emoji: true,
	underline: true,
	omitExtraWLInCodeBlocks: true,
	parseImgDimensions: true,
	ghCodeBlocks: true,
	smartIndentationFix: true,
	disableForced4SpacesIndentedSublists: true,
	requireSpaceBeforeHeadingText: true,
	encodeEmails: true,
	openLinksInNewWindow: true,
	backslashEscapesHTMLTags: true,
});

// -----------------------------------------------------------------------------
// The initialize function must be run each time a new page is loaded
Office.initialize = reason => {

};

// -----------------------------------------------------------------------------
// Add any ui-less function here
function convertMarkdown(event){
	// store the event for later
	btnEvent = event;

  let item = Office.context.mailbox.item;
  let body = item.body;
	let contents = null;
	let revert = null;

	// Is enum, no instanceof (*u*)
	if (item.itemType == Office.MailboxEnums.ItemType.Message){

		body.getAsync(Office.CoercionType.Text, {}, (result) => {
			if (result.status == "succeeded"){
				if (unrendered === null){
					unrendered = result.value;
					contents = converter.makeHtml(unrendered);
				}
				else {
					if (result.value != unrendered){
						revert = confirm(
							"You've modified the message since it was rendered and your changes "
							+ "will be lost if reverted, would you like to revert it anyway?"
						);
						if (!revert) {
							btnEvent.completed();
							return;
						}
					}

					contents = unrendered;
					unrendered = null;
				}


				body.setAsync(contents, {coercionType:"html", }, (result) => {
					// We have to have some kind of error handeling.
					if(result.status != "succeeded")
						console.log(result);

					btnEvent.completed();
				});
			}
			else {
				console.log(result);
				btnEvent.completed();
			}
		});
	}
}
