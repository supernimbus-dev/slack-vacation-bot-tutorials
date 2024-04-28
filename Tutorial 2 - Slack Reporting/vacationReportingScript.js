// CONSTS & VARIABLES //
const startingCellIndex = 3; // we start at A5
let currSheet = undefined;
let currDay = undefined;
let currDayCell = undefined;
const weekDayList = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

function onOpen() {
  loadSheet();
}

function checkToday(){
  if(!currSheet){
    loadSheet();
  }
  const todayRowIndex = (getCurrentDayInYear()+startingCellIndex);
  const valuesForDay = getValuesForDay(todayRowIndex);
  const dayName = weekDayList[new Date(Date.now()).getDay()];
  // If this is today or
  if(dayName === "Saturday" || dayName === "Sunday"){
    return;
  }
  const dayReport = {
    [dayName]:  valuesForDay
  }
  console.log(dayReport);
  sendReport(dayReport);
}

function checkWeek(){
  if(!currSheet){
    loadSheet();
  }
  postToSlack("[Weekly Report]");
  let mondayRowIndex = getRowForMonday();
  let weekReport = {};
  for(let i = 1; i < 6; i++){
    const dayReport = getValuesForDay(mondayRowIndex+(i-1));
    const dayString = weekDayList[i];
    weekReport[dayString] = dayReport;
  }
  console.log(JSON.stringify(weekReport));
  sendReport(weekReport);
}

////////// SPREADSHEET FUNCTIONS //////////////

function loadSheet(){
  // We want to get the list of all tabs in the sheet and focus on the newest sheet //
  // this will be the current year //
  let allTabs = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  let yearsList = [];
  for(let i = 0; i < allTabs.length; i++){
    let sheetData = allTabs[i];
    if(hasNumber(sheetData.getSheetName())){
      let year = parseInt(sheetData.getSheetName().split(' ')[0]);
      yearsList.push({ 
        sheet:  sheetData,
        year:   year
      });
    }
  }
  yearsList = yearsList.sort((a, b) => b.year - a.year);
  let currYear = yearsList[0];
  console.log(`Current Year [${currYear.year}] - ID: ${currYear.sheet.getSheetId()}`);
  currSheet = SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(currYear.sheet);

  // Now we want to focus on the current date //
  currDay = getCurrentDayInYear();
  currDayCell = "A"+(startingCellIndex+currDay);
  console.log(`Current Day's Cell [${currDayCell}]`);
  let currCell = currSheet.getRange(currDayCell);
  SpreadsheetApp.getActiveSpreadsheet().setCurrentCell(currCell);
}

function getValuesForDay(rowIndex){
  const currDayCell = "C"+(rowIndex);
  const finalCell = `${numberToEncodedLetter(getNoOfEmployees()+2)}${(rowIndex)}`;
  const rangeString = `${currDayCell}:${finalCell}`;
  const range = currSheet.getRange(rangeString);
  let summary = [];
  for(let i = 0; i < range.getValues().length; i++){
    for(let val = 0; val < range.getValues()[i].length; val++){
      const employeeIndex = `${numberToEncodedLetter(val+3)}3`;
      const employeeName =  currSheet.getRange(employeeIndex).getValues()[0][0];
      const valString = getLeaveContext(range.getValues()[i][val]);
      if(valString !== null){
        summary.push({
          name:   employeeName,
          value:  valString
        });
      }
    }
  }
  // Now we need to check if there is a bank holiday here //
  // if there is, we should set the report to one field instead of reporting "BH" for everyone //
  for(let i = 0; i < summary.length; i++){
    if(summary[i].value.indexOf("BH") > -1){
      summary = [{
        name:   "Report Bot",
        value:  "Monday is a Bank Holiday <!channel>"
      }]
      break;
    }
  }

  console.log("SUMMARY : ["+ rowIndex +"]"+JSON.stringify(summary));

  return summary;
}

function getNoOfEmployees(){
  const rangeString = `C3:Z3`; // If you have more than 26 employees then you have to update
  const rangeData = currSheet.getRange(rangeString).getValues();
  for(let i = 0; i < rangeData[0].length; i++){
    var name =  rangeData[0][i];
    if(name === undefined || name === ""){
      return (i+1);
    }
  }
}

function getRowForMonday(){
  let currDay = getCurrentDayInYear();
  // Get next Monday //
  for(let i = 0; i < 7; i++){
    const currCell =      (currDay+startingCellIndex+i);
    const rangeString =   `B${(currDay+startingCellIndex+i)}`;
    const currWeekDay =   currSheet.getRange(rangeString).getValues()[0];
    if(currWeekDay == 'Monday'){
      currWeekStartCell = currWeekDay;
      currDay = currCell;
      break;
    }
  }
  console.log("Monday is Cell: "+currDay);
  return currDay;
}

////////// HELPER FUNCTIONS //////////////

function hasNumber(myString) {
  return /\d/.test(myString);
}

function getCurrentDayInYear(){
  var now = new Date();
  var start = new Date(now.getFullYear(), 0, 0);
  var diff = now - start;
  var oneDay = 1000 * 60 * 60 * 24;
  var day = Math.floor(diff / oneDay);
  return day;
}

function numberToEncodedLetter(number) {
    //Takes any number and converts it into a base (dictionary length) letter combo. 0 corresponds to an empty string.
    //It converts any numerical entry into a positive integer.
    if (isNaN(number)) {return undefined}
    number = Math.abs(Math.floor(number))

    const dictionary = getDictionary()
    let index = number % dictionary.length
    let quotient = number / dictionary.length
    let result
    
    if (number <= dictionary.length) {return numToLetter(number)}  //Number is within single digit bounds of our encoding letter alphabet

    if (quotient >= 1) {
        //This number was bigger than our dictionary, recursively perform this function until we're done
        if (index === 0) {quotient--}   //Accounts for the edge case of the last letter in the dictionary string
        result = numberToEncodedLetter(quotient)
    }

    if (index === 0) {index = dictionary.length}   //Accounts for the edge case of the final letter; avoids getting an empty string
    
    return result + numToLetter(index)

    function numToLetter(number) {
        //Takes a letter between 0 and max letter length and returns the corresponding letter
        if (number > dictionary.length || number < 0) {return undefined}
        if (number === 0) {
            return ''
        } else {
            return dictionary.slice(number - 1, number)
        }
    }
}

function getDictionary() {
    return validateDictionary("ABCDEFGHIJKLMNOPQRSTUVWXYZ")

    function validateDictionary(dictionary) {
        for (let i = 0; i < dictionary.length; i++) {
            if(dictionary.indexOf(dictionary[i]) !== dictionary.lastIndexOf(dictionary[i])) {
                console.log('Error: The dictionary in use has at least one repeating symbol:', dictionary[i])
                return undefined
            }
        }
        return dictionary
    }
}

/////////// SLACK //////////////
const webhookUrl = "https://hooks.slack.com/services/<>/<>/<>";

function postToSlack(message) {

  let payload = {
    "channel":    "#int_general",
    "username":   "Annual Leave Calendar Bot",
    "text":       message,
  }
  let options = {
    "method":         "post",
    "contentType":    "application/json",
    "payload":        JSON.stringify(payload)
  };
  console.log(JSON.stringify(payload));
  return UrlFetchApp.fetch(webhookUrl, options)
}

function sendReport(report){
  const reportKeys =  Object.keys(report);
  let reportString = "\n";
  if(reportKeys.length > 0){
    for(let i = 0; i < reportKeys.length; i++){
      const dayString = reportKeys[i];
      let dayPrefixString = "*"+dayString+"*";
      let dayPayloadString = '';
      if(report[dayString].length === 0){
        reportString += dayPrefixString + '\n' + " Nothing To Report " + '\n';
      }
      else{
        for(let y = 0; y < report[dayString].length; y++){
          const alData = report[dayString][y];
          dayPrefixString += '\n';
          
          const newName = alData.name.padEnd(50-alData.name.length, ' ');
          dayPrefixString += "          "+newName+alData.value;
        }
        reportString += dayPrefixString + '\n';
      }
    }
  }
  postToSlack(reportString);
}

function getLeaveContext(value){
  let message = "";
  switch(value){
    case "AL":
      message += "Annual Leave (Enjoy)";
      break;
    case "Requested":
      message += "Requested - Not Approved <@manager>";
      break;
  }
  if(message.length === 0){ // << messy hack
    if(value.indexOf('OOO') > -1){
      message += "Out of the Office: "+ value.split('OOO:')[1];
    }
    else if(value !== ''){
      message += "Note: "+ value;
    }
  }
  if(message.length === 0){
    return null;
  }
  return message;
}