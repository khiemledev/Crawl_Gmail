/*
 * Script to crawl emails from Gmail and write to a Google sheet
 * Author: Khiem Le
 * Date: 03/04/2023
 */

function getEmails() {
  const maxThread = 30 // -1 means no limit
  const maxMessage = 10000 // -1 means no limit
  const spreadsheet_name = 'Crawled emails from Gmail'

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()

  // If there is no active spreadsheet, create a new one
  if (spreadsheet == null) {
    var now = new Date()
    var formattedDate = Utilities.formatDate(now, 'GMT', 'yyyy-MM-dd_HH-mm-ss')
    spreadsheet = SpreadsheetApp.create(spreadsheet_name + '_' + formattedDate)
  }

  var sheet = spreadsheet.getActiveSheet()

  // If there is no active sheet, create a new one
  if (sheet == null) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet()
  }

  sheet.clearContents() // Clear existing data in the sheet

  // Set the header row with the column names
  sheet.appendRow([
    'id',
    'thread_id',
    'date',
    'sender',
    'receiver',
    'cc',
    'bcc',
    'subject',
    'body_plain',
  ])

  // Get all messages from Gmail
  var threads = GmailApp.getInboxThreads()
  if (maxThread != -1) {
    threads = threads.slice(0, Math.min(maxThread, threads.length))
  }
  var messages = GmailApp.getMessagesForThreads(threads)
  console.info(`Found ${messages.length} messages`)

  console.info('Writing messages to Google sheet...')
  // Loop through all messages and add them to the sheet
  let len =
    maxMessage != -1 ? Math.min(maxMessage, messages.length) : messages.length
  for (var i = 0; i < len; i++) {
    console.info(`Writing message ${i}...`)
    var msg = messages[i][0]
    var id = msg.getId()
    var threadId = msg.getThread().getId()
    var receiver = msg.getTo()
    var cc = msg.getCc()
    var bcc = msg.getBcc()
    var date = msg.getDate()
    var sender = msg.getFrom()
    var subject = msg.getSubject()
    var bodyPlain = msg.getPlainBody()

    bodyPlain = bodyPlain.replace(/https?:\/\/[^\s]+/g, '<Link>') // Replace all link with <Link> token to reduce length
    // Replace multiple consecutive blank lines with 2 blank lines, while trimming whitespace from each line
    bodyPlain = bodyPlain.replace(/ *\n\s*( *\n\s*)+/g, '\n\n')
    bodyPlain = bodyPlain.trim()

    // Remove all &#847;&zwnj; tokens from the body html
    bodyPlain = bodyPlain.replace(/&#847;/g, '')
    bodyPlain = bodyPlain.replace(/&zwnj;/g, '')

    if (bodyPlain.length >= 50000) {
      // Skip cell that have more than 50000 characters (litmit of google sheet)
      console.info(`Skip mesage ${i}`)
      continue
    }
    var row = [
      id,
      threadId,
      date,
      sender,
      receiver,
      cc,
      bcc,
      subject,
      bodyPlain,
    ]
    sheet.appendRow(row)
  }

  console.info('Completed!')
}

function myFunction() {
  getEmails()
}
