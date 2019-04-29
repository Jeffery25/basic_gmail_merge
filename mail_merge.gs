function sendEmails() {
  // Fetch the data sheet ("Contact_List")
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = ss.getSheets()[0];
  var dataRange = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, dataSheet.getLastColumn());
  
  // Fetch the template sheet ("Email_Template")
  var templateSheet = ss.getSheets()[1];
  var templateRange = templateSheet.getRange(2, 1, templateSheet.getLastRow() - 1, 3);

  // find the label for programming emails and snoozing them
  var programLabel = GmailApp.getUserLabelByName("GmailDelaySend/ToSend");
  var snoozeLabel;
  
  // placeholders for email drafts and messages
  var draft;
  var message;
  
  // boolean for knowing if the email is a bump or not
  // and a placeholder for the thread and message to bump when found.
  // append 'bump' to status line to make the email a followup
  var bump = false;
  var threadToBump;
  var messageToBump;
  
  // boolean for knowing if we are testing or not
  // email address sent to if testing
  var test = false;
  var testEmail = "durand.jeffery@gmail.com";
  
  // Create one JavaScript object per row of data and row of templates.
  objects = getRowsData(dataSheet, dataRange);
  templates = getRowsData(templateSheet, templateRange);

  // Find the number of the column named "Status" and "Email ID"; program is halted if there is none
  var headersRow = dataSheet.getRange(1, 1, 1, dataSheet.getLastColumn()).getValues()[0];
  var statusColumn = findColumnWithHeader(headersRow, "Status");
  var threadIdColumn = findColumnWithHeader(headersRow, "Thread ID");
  var emailsSentColumn = findColumnWithHeader(headersRow, "Emails Sent");
  if (statusColumn == null || threadIdColumn == null || emailsSentColumn == null) {
    Logger.log("Either the 'Status,' 'Thread ID,' or 'Emails Sent' column is missing");
    return;
  }
  
  // For every row object, create a personalized email from a template and send
  // it to the appropriate person.
  for (var i = 0; i < objects.length; ++i) {
    // Get a row object
    var rowData = objects[i];
    
    // Set test variable to false as default
    test = false;
    
    // depending on the word in the status column, do the appropriate action
    switch(rowData.status) {
      // ready or draft: go on below
      case "test":
        test = true;
      case "ready":        
      case "draft":
        bump = false;
        break;
      // if bumping, make 'bump' true
      case "test bump" :
        test = true;
      case "bump":
      case "draft bump":
      case "force draft bump":
        bump = true;
        break;
      // other status: no action
      case "drafted":
      case "drafted bump":
      case "programmed":
      case "programmed bump":
      case "sent":        
      case "sent bump":
      case "standby":
      case "done":
      case "failed bump":
      case "input error":
      case "incorrect ID":
      case "tested":
      case "tested bump":
      case "programmed test":
      case "programmed test bump":
        continue;
      // if empty, set to standby
      case null:
        updateSheet(dataSheet, i, statusColumn, "standby");
        continue;
      // all other input is invalid
      default:
        updateSheet(dataSheet, i, statusColumn, "input error");
        Logger.log("input error: invalid input in 'Status' column");
        continue;
    }
    
    // checking "Email Address" contains something
    // checking if the "Template Row" is inputted and is a number in the right range
    // checking if the "Snooze" row inputted is a number in the right range
    // checking the "Thread ID" column is empty unless we're bumping an email
    // checking the "Emails Sent" column has the right input
    Logger.log(test);
    var templateNum = rowData.templateRow;
    var snoozeNum = rowData.snooze;
    if ((rowData.emailAddress == null && test == false)
      || (typeof templateNum != "number" || !(templateNum - 2 < templates.length && 0 <= templateNum - 2))
      || (snoozeNum != null && getDayLabelName(snoozeNum) == null)
      || (rowData.threadId != null && bump == false)
      || (rowData.emailsSent != null && bump == false)
      || (typeof rowData.emailsSent != "number" && rowData.emailsSent != null)
      || (rowData.emailsSent == null && rowData.threadId != null)) {
        updateSheet(dataSheet, i, statusColumn, "input error");
        Logger.log("input error: entries in 'template row' or 'Snooze' or 'emails sent' are invalid ; or 'email address' empty"
                   + " when this is not a test; or 'thread ID' or 'emails sent' not empty yet you are not bumping");
      continue;
    }
    
    // If we're bumping, increment the "Emails Sent" column, otherwise set to 1
    // placed here so that the right number can be used in the emails themselves
    if (bump) {
      rowData.emailsSent += 1;
    } else {
      rowData.emailsSent = 1;
    }
    if (!test) {
      updateSheet(dataSheet, i, emailsSentColumn, rowData.emailsSent);
    }
    
    // Generate a personalized email.
    // Get the right template content: subject, body and signature
    var rowTemplate = templates[templateNum - 2];
    var emailSubjectTemplate = rowTemplate.subject;
    var emailBodyTemplate = rowTemplate.body;
    var emailSignatureTemplate = rowTemplate.signature;
        
    // Given a template string, replace markers (for instance ${"First Name"}) with
    // the corresponding value in a row object (for instance rowData.firstName).
    var emailSubject = fillInTemplateFromObject(emailSubjectTemplate, rowData);
    var emailBody = fillInTemplateFromObject(emailBodyTemplate, rowData);
    var emailSignature = fillInTemplateFromObject(emailSignatureTemplate, rowData);
    var emailText = "";
    var htmlEmailText = "";
    
    // add the "re: " header gmail automatically adds
    if (bump)
    {
      emailSubject = "Re: " + emailSubject;
    }
    
    // Fill in the email Content
    // If we're using the "Date to Send" field, add the instuction line to program the email
    if (rowData.dateToSend != null) {
      emailText = "@" + rowData.dateToSend + "\n" + emailText;
    }
    
    // Add the email body
    emailText = emailText + emailBody;
    
    // add the signature together with the line breaks
    if (emailSignature != null)
    {
      emailText = emailText + "\n\n" + emailSignature;
    }
    
    // To make the text 'html' (otherwise unwanted linebreaks will be added), add tags
    htmlEmailText = emailText.replace(/\n/g,'\n<br>');
    
    // setup the necessary thread and make checks if we're bumping the recipient
    if (bump) {
      if (rowData.threadId == null) {
        threadToBump = findThreadToBump(emailSubject, rowData.emailAddress, rowData.status);
      } else {
        // IF THE LINE BELOW FAILS: THE THREAD ID GIVEN IN SHEETS IS INCORRECT
        // find thread with Thread ID ; if this fails do the search anyways
        threadToBump = GmailApp.getThreadById(rowData.threadId);
        if (threadToBump == null) {
          updateSheet(dataSheet, i, statusColumn, "incorrect ID");
        }
      }
      // get the message to bump: last one sent to the person
      messageToBump = findLastMessage(threadToBump, rowData.emailAddress);

      // check that the person did not respond in the thread
      if (personReplied(threadToBump, rowData.emailAddress, rowData.status) || messageToBump == false) {
        updateSheet(dataSheet, i, statusColumn, "failed bump");
        continue;
      }
    }
    
    // If we're bumping, append the previous message and use previous title unless otherwise specified
    if (bump) {
      if (emailSubject == null)
      {
        emailSubject = messageToBump.getSubject();
      }
      emailText += '\n\nOn ' + messageToBump.getDate().toString() + ' <' + messageToBump.getFrom() + '> wrote:\n\n' + messageToBump.getBody();
      htmlEmailText +='<br><br><div class="gmail_quote"><div dir="ltr">On ' + messageToBump.getDate().toString() + ' &lt;<a href="mailto:' + messageToBump.getFrom() + '">' + messageToBump.getFrom() + '</a>&gt; wrote:<br></div><blockquote class="gmail_quote" style="margin:0 0 0 .8ex;border-left:1px #ccc solid;padding-left:1ex">' + messageToBump.getBody() + '</blockquote></div>';
    }
    
    // make the draft
    if (test) {
      // make draft (testing: remove cc and send to test recipient)
      draft = GmailApp.createDraft(testEmail, emailSubject, emailText, 
                                   {
                                     htmlBody: htmlEmailText,
                                   });
    } else {
      // make draft (not a test: send normally)
      draft = GmailApp.createDraft(rowData.emailAddress, emailSubject, emailText, 
                                   {
                                     cc: rowData.cc,
                                     bcc: rowData.bcc,
                                     htmlBody: htmlEmailText,
                                   });
      // track the thread with its ID (if not testing)
      recordThreadId(dataSheet, i, threadIdColumn, draft.getMessage());
    }
      
    // add the relevant snooze label
    if (snoozeNum != null) {
      snoozeLabel = GmailApp.getUserLabelByName(getDayLabelName(snoozeNum));
      snoozeLabel.addToThread(draft.getMessage().getThread());
    }
    
    // if status is "ready", send the email or program it if "Date To Send" has an entry,
    if (rowData.status == "ready" || rowData.status == "bump" || rowData.status == "test" || rowData.status == "test bump") {
      if (rowData.dateToSend == null) {
        message = draft.send();
        // add the relevant snooze label again (sending takes labels off)
        if (snoozeNum != null) {
          snoozeLabel.addToThread(message.getThread());
        }
      } else {
        programLabel.addToThread(draft.getMessage().getThread());
      }
    }
    
    // update the status and the spreadsheet in case the script is interrupted
    updateSheet(dataSheet, i, statusColumn, nextStatus(rowData.status, rowData.dateToSend));
    // Make sure the cell is updated right away in case the script is interrupted
    SpreadsheetApp.flush();
  } 
}


//////////////////////////////////////////////////////////////////////////////////////////
//
// The functions by Jeffery Durand
//
//////////////////////////////////////////////////////////////////////////////////////////

// Updates the "Status" column of the Data Sheet with a message
// This means there can be no double sends even if the program is interrupted
function updateSheet(sheet, row, column, message) {
  sheet.getRange(row + 2, column).setValue(message);
}

// This writes the Email ID to the relevant column
// The "email" variable should be a GmailMessage type
function recordThreadId(sheet, row, column, email) {
  var threadID = email.getThread().getId();
  sheet.getRange(row + 2, column).setValue(threadID);
}

// Find the column number of a particular header
// This is useful when we need to write to the spreadsheet and 
// the columns may move around but keep the same name.
function findColumnWithHeader(headers, headerName) {
  for (i = 0; i < headers.length; i++) {
    if (normalizeHeader(headers[i]) == normalizeHeader(headerName)) {
      return i + 1;
    }
  }
  return null;
}

// Check if the thread being followed up on has a reply from the person
// Sends true if the person did reply, false otherwise
// Sends false if we're using "force draft bump"
function personReplied(thread, address, status) {
  if (status == "force draft bump" || thread == false) {
   return false; 
  }
  var messages = thread.getMessages();
  for (i = 0; i < messages.length; i++) {
    if (messages[i].getFrom().indexOf(address) >= 0) {
      return true;
    }
  }
  return false;
}

// finds the last message in a thread that was sent to a particular person
// this does not look at cc or bcc
function findLastMessage(thread, address) {
  if (thread == false || thread == null)
  {
   return false; 
  }
  var messages = thread.getMessages();
  for (i = messages.length - 1; i >= 0; i--) {
    var message = messages[i];
    if (message.getTo().indexOf(address) >= 0) {
      return message;
    }
  }
  return false;
}

// If we don't have the thread ID, go search for the message
// Returns false if none or several threads are found unless we're using "force draft bump"
function findThreadToBump(subject, address, status) {
  // search for the right thread, get the array of answers
  threads = GmailApp.search('subject:"' + subject + '" to:' + address);
  // check the result is correct (unless we're forcing through
  if (threads.length == 0 || (subject == null && status != "force draft bump") 
      || (threads.length > 1 && status != "force draft bump")) {
    return false
  }
  return threads[0];
}

// get the next status to put on the spreadsheet after an action
function nextStatus(status, dateToSend) {
  if (dateToSend == null) {
    switch(status) {
      case "ready":
        return "sent";
      case "bump":
        return "sent bump";
      case "test":
        return "tested";
      case "test bump":
        return "tested bump";
      case "draft":
        return "drafted";
      case "draft bump":
      case "force draft bump":
        return "drafted bump";
      default:
        return null;
    }
  }
  if (dateToSend != null) {
    switch(status) {
      case "ready":
        return "programmed";
      case "bump":
        return "programmed bump";
      case "test":
        return "programmed test";
      case "test bump":
        return "programmed test bump";
      case "draft":
        return "drafted";
      case "draft bump":
      case "force draft bump":
        return "drafted bump";
      default:
        return null;
    }
  }
}


// used for Snooze label name retrieving
function getDayLabelName(i) {
  // Retrieve name of the day label when given number 
  // note that 0 corresponds to Sunday so the getDay() function works
  var root = "Snooze/"
  switch(i) {
    case 1:
      return root + "1 Monday";
    case 2:
      return root + "2 Tuesday";
    case 3:
      return root + "3 Wednesday";
    case 4:
      return root + "4 Thursday";
    case 5:
      return root + "5 Friday";
    case 6:
      return root + "6 Saturday";
    case 0:
      return root + "7 Sunday";
    default:
      return null
  }
}

//////////////////////////////////////////////////////////////////////////////////////////
//
// The code below was in the original version found on internet
//
//////////////////////////////////////////////////////////////////////////////////////////

// Replaces markers in a template string with values define in a JavaScript data object.
// Arguments:
//   - template: string containing markers, for instance ${"Column name"}
//   - data: JavaScript object with values to that will replace markers. For instance
//           data.columnName will replace marker ${"Column name"}
// Returns a string without markers. If no data is found to replace a marker, it is
// simply removed.
function fillInTemplateFromObject(template, data) {
  
  // case if null string is sent
  if (template == null) {
   return null; 
  }
  
  var email = template;
  // Search for all the variables to be replaced, for instance ${"Column name"}
  var templateVars = template.match(/\$\{\"[^\"]+\"\}/g);

  // if no values to replace go back
  if (templateVars == null) {
    return template;
  }
  
  // Replace variables from the template with the actual values from the data object.
  // If no value is available, replace with the empty string.
  for (var i = 0; i < templateVars.length; ++i) {
    // normalizeHeader ignores ${"} so we can call it directly here.
    var variableData = data[normalizeHeader(templateVars[i])];
    email = email.replace(templateVars[i], variableData || "");
  }
  return email;
}

//////////////////////////////////////////////////////////////////////////////////////////
//
// The code below is reused from the 'Reading Spreadsheet data using JavaScript Objects'
// tutorial.
//
//////////////////////////////////////////////////////////////////////////////////////////

// getRowsData iterates row by row in the input range and returns an array of objects.
// Each object contains all the data for a given row, indexed by its normalized column name.
// Arguments:
//   - sheet: the sheet object that contains the data to be processed
//   - range: the exact range of cells where the data is stored
//   - columnHeadersRowIndex: specifies the row number where the column names are stored.
//       This argument is optional and it defaults to the row immediately above range; 
// Returns an Array of objects.
function getRowsData(sheet, range, columnHeadersRowIndex) {
  columnHeadersRowIndex = columnHeadersRowIndex || range.getRowIndex() - 1;
  var numColumns = range.getEndColumn() - range.getColumn() + 1;
  var headersRange = sheet.getRange(columnHeadersRowIndex, range.getColumn(), 1, numColumns);
  var headers = headersRange.getValues()[0];
  return getObjects(range.getValues(), normalizeHeaders(headers));
}

// For every row of data in data, generates an object that contains the data. Names of
// object fields are defined in keys.
// Arguments:
//   - data: JavaScript 2d array
//   - keys: Array of Strings that define the property names for the objects to create
function getObjects(data, keys) {
  var objects = [];
  for (var i = 0; i < data.length; ++i) {
    var object = {};
    var hasData = false;
    for (var j = 0; j < data[i].length; ++j) {
      var cellData = data[i][j];
      if (isCellEmpty(cellData)) {
        continue;
      }
      object[keys[j]] = cellData;
      hasData = true;
    }
    if (hasData) {
      objects.push(object);
    }
  }
  return objects;
}

// Returns an Array of normalized Strings.
// Arguments:
//   - headers: Array of Strings to normalize
function normalizeHeaders(headers) {
  var keys = [];
  for (var i = 0; i < headers.length; ++i) {
    var key = normalizeHeader(headers[i]);
    if (key.length > 0) {
      keys.push(key);
    }
  }
  return keys;
}

// Normalizes a string, by removing all alphanumeric characters and using mixed case
// to separate words. The output will always start with a lower case letter.
// This function is designed to produce JavaScript object property names.
// Arguments:
//   - header: string to normalize
// Examples:
//   "First Name" -> "firstName"
//   "Market Cap (millions) -> "marketCapMillions
//   "1 number at the beginning is ignored" -> "numberAtTheBeginningIsIgnored"
function normalizeHeader(header) {
  var key = "";
  var upperCase = false;
  for (var i = 0; i < header.length; ++i) {
    var letter = header[i];
    if (letter == " " && key.length > 0) {
      upperCase = true;
      continue;
    }
    if (!isAlnum(letter)) {
      continue;
    }
    if (key.length == 0 && isDigit(letter)) {
      continue; // first character must be a letter
    }
    if (upperCase) {
      upperCase = false;
      key += letter.toUpperCase();
    } else {
      key += letter.toLowerCase();
    }
  }
  return key;
}

// Returns true if the cell where cellData was read from is empty.
// Arguments:
//   - cellData: string
function isCellEmpty(cellData) {
  return typeof(cellData) == "string" && cellData == "";
}

// Returns true if the character char is alphabetical, false otherwise.
function isAlnum(char) {
  return char >= 'A' && char <= 'Z' ||
    char >= 'a' && char <= 'z' ||
    isDigit(char);
}

// Returns true if the character char is a digit, false otherwise.
function isDigit(char) {
  return char >= '0' && char <= '9';
}