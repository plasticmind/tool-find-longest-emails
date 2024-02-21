# Find Longest Emails

**A Google Apps Script that searches the emails in your Gmail mailbox writes them to a Google Sheet with a character count (that can be sorted).**

⚠️ Note: You'll need to create a new script in [Google Apps Script](https://script.google.com/), copy/paste the code in from [Code.gs](https://github.com/plasticmind/tool-find-longest-emails/blob/main/Code.gs), and modify the constants at the beginning of the script with relevant information. You'll also need to create a new Google Sheets where the data will be written and set up permssions. Finally, Google App Scripts only lets you access 500 emails at a time, so you'll either need to manually run this over and over, or you can [set up a trigger](https://developers.google.com/apps-script/guides/triggers/installable) like a cron task to run the script recurrant until all emails have been sorted.
