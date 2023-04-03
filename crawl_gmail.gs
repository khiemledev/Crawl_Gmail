/*
 * Script to crawl emails from Gmail and write to a Google sheet
 * Author: Khiem Le
 * Date: 03/04/2023
 */

function getEmails() {
  const max_messages = -1 // -1 means no limit

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()

  // If there is no active spreadsheet, create a new one
  if (spreadsheet == null) {
    spreadsheet = SpreadsheetApp.create('Crawled emails from Gmail')
  }

  var sheet = spreadsheet.getActiveSheet()

  // If there is no active sheet, create a new one
  if (sheet == null) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet()
  }

  sheet.clearContents() // Clear existing data in the sheet

  // Set the header row with the column names
  sheet.appendRow(['date', 'sender', 'subject', 'body_plain'])

  // Get all messages from Gmail
  var threads = GmailApp.getInboxThreads()
  var messages = GmailApp.getMessagesForThreads(threads)
  console.info(`Found ${messages.length} messages`)

  console.info('Writing messages to Google sheet...')
  // Loop through all messages and add them to the sheet
  let len =
    max_messages != -1
      ? Math.min(max_messages, messages.length)
      : messages.length
  for (var i = 0; i < len; i++) {
    console.info(`Writing message ${i}...`)
    var msg = messages[i]
    var date = msg[0].getDate()
    var sender = msg[0].getFrom()
    var subject = msg[0].getSubject()
    var body_plain = msg[0].getPlainBody()
    if (body_plain.length >= 50000) {
      // Skip cell that have more than 50000 characters (litmit of google sheet)
      console.info(`Skip mesage ${i}`)
      continue
    }
    var row = [date, sender, subject, body_plain]
    sheet.appendRow(row)
  }

  console.info('Completed!')
}

function myFunction() {
  getEmails()
}
