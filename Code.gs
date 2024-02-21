const SPREADSHEET_ID = '#######'; // The Google Sheet ID to write results to
const FROM_EMAIL = 'your.email@example.com'; // The email address to filter by (presumably yours)
const MESSAGES_PER_BATCH = 500; // How many messages to fetch at one time (500 is the max)
const MIN_CHARACTERS = 2000; // Only log emails with more characters than this

function findLongestEmails() {
  let sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getActiveSheet();
  
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Subject', 'Date', 'Length', 'Email ID', 'URL']);
  }

  // Retrieve the last processed email date from properties
  let lastProcessedDate = PropertiesService.getScriptProperties().getProperty('LAST_PROCESSED_DATE');
  let searchQuery = 'from:' + FROM_EMAIL;

  if (lastProcessedDate) {
      searchQuery += ` before:${lastProcessedDate}`;
  }

  Logger.log("Using search query: " + searchQuery);

  // After retrieving threads
  let threads = GmailApp.search(searchQuery, 0, MESSAGES_PER_BATCH);

  if (threads.length === 0) {
      Logger.log("No new emails to process.");
      return;
  }

  // Verify the first and last thread
  Logger.log("First Thread's Date: " + (threads[0] ? threads[0].getLastMessageDate() : "No first thread found"));
  Logger.log("Last Thread's Date: " + (threads[threads.length - 1] ? threads[threads.length - 1].getLastMessageDate() : "No last thread found"));


  for (let thread of threads) {
    let messages = thread.getMessages();
    for (let message of messages) {
      let fromAddress = message.getFrom();
      if (messages.length === 0) continue;  // Skip threads without messages

      if (fromAddress.includes(FROM_EMAIL)) {
        let cleanedBody = getCleanedBody(message.getPlainBody());
        let bodyLength = cleanedBody.length;

        if (bodyLength > MIN_CHARACTERS) {
          let messageId = message.getHeader('Message-ID');
          // Creates a live link to the email in question for convenient browsing in context of a thread 
          let emailUrl = 'https://mail.google.com/mail/u/0/#search/rfc822msgid%3A' + encodeURIComponent(messageId);
          let rowData = [message.getSubject(), message.getDate(), bodyLength, messageId, emailUrl];
          sheet.appendRow(rowData);
        }
      }
    }
  }

  // Convert the date of the last email in this batch to a Date object, subtract one second, then store
  Logger.log("Last Thread's Subject: " + threads[threads.length - 1].getFirstMessageSubject());
  Logger.log("Last Thread's Date: " + threads[threads.length - 1].getLastMessageDate());
  
  // At the end of the function, after processing
  if (threads.length > 0) {
      // Convert the date of the last email in this batch to a Date object, subtract one second, then store
      let lastEmailDateObj = new Date(threads[threads.length - 1].getLastMessageDate());
      lastEmailDateObj.setSeconds(lastEmailDateObj.getSeconds() - 1);
      let adjustedLastEmailDate = Utilities.formatDate(lastEmailDateObj, "GMT", "yyyy/MM/dd");
      PropertiesService.getScriptProperties().setProperty('LAST_PROCESSED_DATE', adjustedLastEmailDate);
  }
  
  Logger.log("Processed " + MESSAGES_PER_BATCH + " emails and added qualifying emails to the Google Sheet.");
}

function getCleanedBody(body) {
  let lines = body.split('\n');
  let newLines = [];

  let isQuotedText = false;

  for (let line of lines) {
    // This checks if the line starts with ">" or "From:" which often indicates quoted text or a previous email header
    if (line.startsWith(">") || line.trim().startsWith("From:")) {
      isQuotedText = true;
    }

    if (!isQuotedText) {
      newLines.push(line);
    }

    // Look for empty lines or breaks between quoted sections to possibly resume reading the email
    if (isQuotedText && line.trim() === "") {
      isQuotedText = false;
    }
  }

  return newLines.join('\n');
}

function clearLastProcessedDate() {
  PropertiesService.getScriptProperties().deleteProperty('LAST_PROCESSED_DATE');
}

function clearSheetExceptHeader() {
  let sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getActiveSheet();
  let lastRow = sheet.getLastRow();
  
  if (lastRow > 1) { // If there are more than just the header rows
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }
}
