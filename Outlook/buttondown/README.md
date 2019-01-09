# It's Time To Button Down!

![Project Status][status]

This is a little office add-in I wrote to enable users to write their emails in
Markdown.

### Installation

### Features

The following features that are checked are features that I've implemented.
I assume you can infer what the unchecked boxes mean. Features I plan on
implementing may take some time since this isn't a priority project of mine.
So please be patient, or submit a pull request with the addition.

### History

My problem originated from a personal decision at my job, which was to write my
shift reports in Markdown. Since I like how easy it is to style text using the
markdown syntax, I thought, why don't I use Markdown and make my shift reports
exemplary? I looked all over the Outlook Add-in Store and couldn't find any
add-ins that were built to enable Markdown rendering within the email client,
so I turned to the Google Chrome App Store and *because I wasn't using the
right search parameters*, I couldn't find any app that supported what I wanted
to do there either.

I already had [HTML Inserter for Gmail][GIHTML] installed in chromium, so my
original solution comprised of writing my shift reports in [atom](https://atom.io);
since I could use the Live Markdown Preview extension to export my Markdown
as raw HTML, setting up my work email to act as a distributor of any emails
coming from my personal address with the title containing "Shift Report" to
all the necessary recipients, and using my personal email to dispatch the
initial copy. This was a bit excessive to say the least, and do to the fact
that the initial copy was being sent from my personal account, the company email
client would not properly categorize the emails into the appropriate inbox,
which led to me getting a really nasty email from our CEO about how
using our personal email addresses to handle any internal information is
forbidden.

![Nice Email](https://insert-my-url-here)

Now, I needed a new solution. Being the problem solver I am, I decompiled the
assets for GmailInsertHTML to use as a starting point for designing a Chrome App
that would fix my problem. I wanted to know the legalities of borrowing the
source code and looked for an open source license bundled with the software.
I couldn't find one, oh boy... I got the creators information out of the source
blob and looked all over for some sign; any, that I wouldn't have to get
permission from some random goober to use their source code to write a solution
to a problem that I'm sure others besides me have had. I was greeted by the
dis-satisfactory truth, that not everyone supports the open source ideology.
I found no Open Source License, and a website stating they were from
Washington DC. Now, there is the possibility that they'd just never heard of
Open Source Software, but due to my personal experiences with individuals
from the DC area, I came to the conclusion that they had no interest in making
*Their "Intellectual Property"* available to the public without compensation
in the form of cold; hard; cash. Before I tweeted them to confirm that and
either get them to open source it, or give me permission to use their source as
a primer for my own project; I discovered I wouldn't be able to use their work
anyway. It was built on [InboxSDKs][InSDK], which is a Gmail specific API.

With another potential solution stifled, I located the Microsoft Office
Extension API in hopes of using it to accomplish my task. Eureka, I'd found
my framework. Within a few days I had my first alpha version working, which
supplied a button to outlook that facilitated invoking my markdown conversion
function. It was everything I needed! But now I needed polish. I added an
internal Webpack framework for localizing the add-in to different regions,
and wrote the support page. Then, while seeking inspiration for the logo,
I found [Markdown Here][MDH].

Oh crap. Here I was writing an Outlook Exclusive feature, and these wonderful
open source supporting devs have made an all inclusive solution.

### Why This Project Wasn't Terminated

There are positives specifically related to having a solution implemented
off the current hardware you're running / connected via web protocol.

 * Less bloat, *which is especially helpful to people like me who's daily driver
   has less than 16G of soldered memory.*

 * No Intrusive Updates, any updates are handled externally.

 * Faster load times (depending on your hardware)

 * No matter what browser you're using; so long as it supports javascript,
   It'll be there.

There's also the fact that I can extend this software to encompass other
Microsoft Office Applications should I feel like it.


[InSDK]: http://www.inboxsdk.com/
[GIHTML]: https://chrome.google.com/webstore/detail/html-inserter-for-gmail/lkdchkblgffcinmodbodlkclphfldkll?hl=en
[MDH]: https://markdown-here.com/get.html

[status]: https://img.shields.io/badge/Project%20Status-On%20Hiatus-yellow.svg
