/**
 * script will only run on data structured with
 * dates as rows and ID keys as columns
 * 
 * adjust dates to iterate through by changing the ranges
 * at lines 37-41 / 63-68 
 */
function onOpen(){
  var app = SpreadsheetApp.getUi();
  app.createMenu('QuickPivots')
  .addItem('Create', 'addPivotTable')
  .addItem('Return Missing Days','returnMissingLdap')
  .addItem('Send Email', 'sendEmail')
  .addToUi();
}

function sendEmail(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName("missingDates");
  var values = s.getRange(2,5,70,1).getValues();
  var message = composeMessage(values);
  var messageHTML = composeHtmlMsg(values);
  //Logger.log(messageHTML);
  
  var emailAddress = values[0][0];
  var subject = 'Missing EOD dates for ' + emailAddress;
  MailApp.sendEmail(emailAddress,subject,message);
}

function composeMessage(values){
  var message = 'please submit EOD forms for the following dates: \n'
  for(var c=0;c<values.length;++c){
    message+='\n'+values[c]
  }
  Logger.log(message);
  return message;
}

function composeHtmlMsg(values){
  var message = 'please submit dates:<br><br><table style="background-color:lightblue;border-collapse:collapse;" border = 1 cellpadding = 5><th>data</th><th>Values</th><tr>'
  for(var c=0;c<values.length;++c){
    message+='</td><td>'+values[c]+'</td></tr>'
  }
  return message+'</table>';
}



function returnMissingLdap(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName("Sheet34");
  var s1 = ss.getSheetByName("missingDates");

  //march2020
  // var dateRange = s.getRange(296,1,17,47).getValues();

  //august2019
  var dateRange = s.getRange(165,1,s.getLastRow(),47).getValues();
  var ldapIndex;
  var arrayIndex;

  //arrays to store empty cells
  var ldapList = [];
  var emptyDates = [];

  // Logger.log(dateRange);

  for(var i=0; i < dateRange.length; i++){
    ldapIndex = dateRange[i];
    for(var j=0; j < ldapIndex.length; j++){
      arrayIndex = ldapIndex[j]
      if(arrayIndex.length == 0){

        //return assoc. ldaps as headers
        var goldIndex = j-2;
        var ldapRange = s.getRange(2,3,1,s.getLastColumn()).getValues();
        var missingldap = ldapRange[0][goldIndex];
        ldapList.push([missingldap])

        //return assoc. dates for range
        //march2020 
        // var dateList = s.getRange(296,1,17,1).getValues();

        //august2019
        var dateList = s.getRange(165,1,s.getLastRow(),1).getValues();
        var missingDate = dateList[i];
        emptyDates.push([missingDate])


        Logger.log(ldapList)
        //Logger.log('ldap: ' + missingldap + ' | missing date: ' + missingDate);

        }
      }
      var newrange = s1.getRange(1,1,ldapList.length,1);
      var newrange2 = s1.getRange(1,2,emptyDates.length,1);
      newrange.setValues(ldapList);
      newrange2.setValues(emptyDates);
  }
}

function addPivotTable(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  //sheet with data
  var sheetName = "Sheet1";

  var pivotTableParams = {};
  //source indicating range of data for pivot_t in json structure
  pivotTableParams.source = {
    sheetId: ss.getSheetByName(sheetName).getSheetId()
  };

  //grouping by the column in data set, eg 0 = group by first column
  pivotTableParams.rows = [
    {
      sourceColumnOffset: 1,
      sortOrder: "ASCENDING"
    }
  ];

  pivotTableParams.columns = [
    {
      sourceColumnOffset: 8,
      sortOrder: 'ASCENDING'
    }
  ];

  //define value calculation for data
  pivotTableParams.values = [
    {
    summarizeFunction: "SUM",
    sourceColumnOffset: 5,
    }];

  //create new sheet for pivot_t
  var pivotTableSheet = ss.insertSheet();
  var pivotTableSheetId = pivotTableSheet.getSheetId();

  //add pivot_t to new sheet, which requires sending 'updateCells' request to sheets API
  //specify at 'start' where we want to put our pivot_t
  //add our parameters in 'rows'

  var requests = {
    'updateCells': {
      'rows': {
        'values': [{
          'pivotTable': pivotTableParams
          }
        ],
          },
          'start': {
            'sheetId': pivotTableSheetId,
          },
          'fields': 'pivotTable'
        }
      };
      Sheets.Spreadsheets.batchUpdate({'requests': requests}, ss.getId());
      }







