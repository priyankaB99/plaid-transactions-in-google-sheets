/***********************************************************************************
 * PULL PLAID TRANSACTIONS IN GOOGLE SHEETS
 * ---
 * Author: Priyanka Iyer
 * -- Based off a Frank Harris repo: https://github.com/hirefrank/plaid-txns-google-sheets/blob/master/README.md
 * Initial Date: Feb 15, 2026
 * MIT License
 *
 ***********************************************************************************/

var SHEET = PropertiesService.getScriptProperties().getProperty('sheet'); // Add name of Google Sheet to Script Properties

function test() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET);
  var last_row = sheet.getLastRow();
  var txns_ids = getTransactionIds(sheet.getRange(2,7,last_row,1));

  console.log(txns_ids)
}

/**
 * Main function that grabs transactions
 */

function run() {

  const ACCESS_TOKEN1 = PropertiesService.getScriptProperties().getProperty('access_token_1');
  const ACCESS_TOKEN2 = PropertiesService.getScriptProperties().getProperty('access_token_2');

  getTransactionHistory(ACCESS_TOKEN1, "B1");

  // // 5s pause to prevent the functions running synchronously 
  Utilities.sleep(5000);

  getTransactionHistory(ACCESS_TOKEN2, "B2");

  // // 5s pause to prevent the functions running synchronously 
  Utilities.sleep(5000);

  cleanup();
}

/**
 * Get the Transactions via the Plaid API
 */

function getTransactionHistory(ACCESS_TOKEN, cursor_location) { 
  const CLIENT_ID = PropertiesService.getScriptProperties().getProperty('client_id');
  const SECRET = PropertiesService.getScriptProperties().getProperty('secret');
  const COUNT = 500;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lastCursor = ss.getSheetByName("Cursor").getRange(cursor_location).getValue();

  // headers are a parameter plaid requires for the post request
  // plaid takes a contentType parameter
  // google app script takes a content-type parameter
  var headers = {                                         
    'contentType': 'application/json',                                        
    'Content-Type': 'application/json',
  };
  
  // data is a parameter plaid requires for the post request
  // created via the plaid quickstart app (node)
  var data = { 
    'access_token': ACCESS_TOKEN,
    'client_id': CLIENT_ID,                                
    'secret': SECRET,
    'count': COUNT,
    'cursor': lastCursor
  };
  
  // pass in the necessary headers
  // pass the payload as a json object
  var parameters = {                                                                                                             
    'headers': headers,            
    'payload': JSON.stringify(data),                            
    'method': 'post',
    'muteHttpExceptions': true,
  };
  
  // api host + endpoint
  var url = "https://production.plaid.com/transactions/sync";
  var response = UrlFetchApp.fetch(url, parameters);
  
  // parse the response into a JSON object
  var json_data = JSON.parse(response);

  // get cursor to know what to update next time
  var nextCursor = json_data.next_cursor
  ss.getSheetByName("Cursor").getRange(cursor_location).setValue(nextCursor)
  
  // get the transactions from the JSON
  var transactions = json_data.added;

  console.log(json_data)
  console.log("Added ", json_data.added)
  console.log("Modified ", json_data.modified)
  console.log("Cursor ", json_data.next_cursor)

  var sheet = ss.getSheetByName(SHEET);
  var last_row = sheet.getLastRow();
  var txns_ids = getTransactionIds(sheet.getRange(2,7,last_row,1));
  
  // Pull new transactions and update modified/removed transactions
  for (i in transactions) {
    if (transactions[i].pending == false) { 
      let merchant = transactions[i].merchant_name ? transactions[i].merchant_name : transactions[i].name
      var row = [
        transactions[i].date,
        transactions[i].name,
        transactions[i].amount,
        merchant,
        transactions[i].personal_finance_category.primary,
        transactions[i].payment_channel,
        transactions[i].transaction_id,
        transactions[i].account_id,
      ]
      // Add below the lowest row in the sheet
      sheet.appendRow(row);
    }
  }

  // Get and update modified transactions from last time
  var modifiedTransactions = json_data.modified;
  for (i in modifiedTransactions) {
    if (modifiedTransactions[i].pending == false) {
      var txnID = txns_ids.indexOf(modifiedTransactions[i].transaction_id)
      if (txnID !== -1) {
        
        sheet.deleteRow(txnID + 2)
        txns_ids.splice(txnID, 1)

        let merchant = modifiedTransactions[i].merchant_name ? modifiedTransactions[i].merchant_name : modifiedTransactions[i].name

        var row = [
          modifiedTransactions[i].date,
          modifiedTransactions[i].name,
          modifiedTransactions[i].amount,
          merchant,
          modifiedTransactions[i].personal_finance_category.primary,
          modifiedTransactions[i].payment_channel,
          modifiedTransactions[i].transaction_id,
          modifiedTransactions[i].account_id,
        ]
        // Add below the lowest row in the sheet
        sheet.appendRow(row);
      }
    }
  }


  // Get and update removed transactions from last time
  var removedTransactions = json_data.modified;
  for (i in removedTransactions) {
    if (removedTransactions[i].pending == false) {
      var txnID = txns_ids.indexOf(modifiedTransactions[i].transaction_id)
      if (txnID !== -1) {
        sheet.deleteRow(txnID + 2)
        txns_ids.splice(txnID, 1)
      }
    }
  }
}

/**
 * Removes all transactions from the spreadsheet
 */
     
function reset() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET);

  var last_row = sheet.getLastRow();
  sheet.getRange('2:' + last_row).activate();
  sheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
};

/**
 * Returns transaction_ids
 */
        
function getTransactionIds(range) {
  var ids = [];
  
  var numRows = range.getNumRows();
  var numCols = range.getNumColumns();
        
  for (var i = 1; i <= numRows; i++) {
    for (var j = 1; j <= numCols; j++) {
      var cv = range.getCell(i,j).getValue();
      if (cv !== '') {
        ids.push(cv);
      }
    }
  }
  return ids;
};

/**
 * Left aligns all cells in the spreadsheet and sorts by date
 */
        
function cleanup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET);
  
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).activate();
  sheet.getActiveRangeList().setHorizontalAlignment('left');
  sheet.getRange('A:A').activate();
  sheet.sort(1, true);
};
        
/**
 * Returns the date in a Plaid friendly format, e.g. YYYY-MM-DD
 */

function formatDate(date) {
  var d = new Date(date),
  month = '' + (d.getMonth() + 1),
  day = '' + d.getDate(),
  year = d.getFullYear();

  if (month.length < 2) month = '0' + month;
  if (day.length < 2) day = '0' + day;

  return [year, month, day].join('-');
}
