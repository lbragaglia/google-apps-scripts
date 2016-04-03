var ss = SpreadsheetApp.getActiveSpreadsheet();
//if (ss.getSheetByName("impostazioni").getRange("B1").getValue()!=""){
//  var TZ = CalendarApp.openByName(ss.getSheetByName("impostazioni").getRange("B1").getValue()).getTimeZone();
//}
var conf = {
  calendario: 'B1',
  notifiche: 'B2'
};

function _getConf(prop) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('impostazioni').getRange(conf[prop]).getValue();
}
function _getCalendario() {
  return _getConf('calendario');
}
function onOpen() {
  if (_getCalendario() == "") {
    ss.toast("Click on setup to start", "Notifiche", 10);
  }
  var menuEntries = [ {name: "Aggiorna calendario", functionName: "importIntoCalendar"} ];
  ss.addMenu("Calendario", menuEntries);
}


function importIntoCalendar(){
  var calendario = _getCalendario();
  if (!calendario) {
    return;
  }
  var dataSheet = ss.getSheetByName("tavoli");
  var dataRange = dataSheet.getRange(2, 1, dataSheet.getMaxRows(), dataSheet.getMaxColumns());

  var templateSheet = ss.getSheetByName("Templates");
  var calendarName = templateSheet.getRange("E1").getValue();
  var siteUrl = templateSheet.getRange("E2").getValue();
  
  var cal =  CalendarApp.getCalendarsByName(calendario)[0];
  var eventTitleTemplate = templateSheet.getRange("E3").getValue();
  var descriptionTemplate = templateSheet.getRange("E4").getValue();  
  var siteTemplate = templateSheet.getRange("E5").getValue(); 

  // Create one JavaScript object per row of data.
  objects = getRowsData(dataSheet, dataRange);

  // For every row object, create a personalized email from a template and send
  // it to the appropriate person.
  for (var i = 0; i < objects.length; ++i) {
      // Get a row object
      var rowData = objects[i];
      if (rowData.eventId && rowData.eventTitle && rowData.action == "Y" ){
        var eventTitle = fillInTemplateFromObject(eventTitleTemplate, rowData);
        var description = fillInTemplateFromObject(descriptionTemplate, rowData);
        var siteText = fillInTemplateFromObject(siteTemplate, rowData);
        // add to calendar bit
        if(rowData.endDate == "All-day"){
           var eventid = cal.createAllDayEvent(eventTitle, rowData.startDate, rowData.endDate, {location:rowData.location, description: description}).getId();
        }
        else{
          var eventid = cal.createEvent(eventTitle, rowData.startDate, rowData.endDate, {location:rowData.location, description: description}).getId();
        }
        var events = cal.getEvents(rowData.startDate, rowData.endDate);
        var event = getEvent(events, eventid);
        
      
        dataSheet.getRange(i+2, 1, 1, 1).setValue("");
        dataSheet.getRange(i+2, 2, 1, 1).setValue("Added "+ Utilities.formatDate(new Date(), TZ, "dd MMM yy HH:mm")).setBackgroundRGB(221, 221, 221);
        dataSheet.getRange(i+2,1,1,dataSheet.getMaxColumns()).setBackgroundRGB(221, 221, 221);
        // Make sure the cell is updated right away in case  the script is interrupted
        SpreadsheetApp.flush();
      }
  }  
  ss.toast("People can now register to those events", "Events imported");
}


// Replaces markers in a template string with values define in a JavaScript data object.
// Arguments:
//   - template: string containing markers, for instance ${"Column name"}
//   - data: JavaScript object with values to that will replace markers. For instance
//           data.columnName will replace marker ${"Column name"}
// Returns a string without markers. If no data is found to replace a marker, it is
// simply removed.
function fillInTemplateFromObject(template, data) {
  var email = template;
  // Search for all the variables to be replaced, for instance ${"Column name"}
  var templateVars = template.match(/\$\{\"[^\"]+\"\}/g);

  // Replace variables from the template with the actual values from the data object.
  // If no value is available, replace with the empty string.
  for (var i = 0; i < templateVars.length; ++i) {
    // normalizeHeader ignores ${"} so we can call it directly here.
    var variableData = isDate(data[normalizeHeader(templateVars[i])]);
    email = email.replace(templateVars[i], variableData || "");
  }

  return email;
}

// Test if value is a date and if so format 
function isDate(sDate) {
  var scratch = new Date(sDate);
  if (scratch.toString() == "NaN" || scratch.toString() == "Invalid Date") {
    return sDate;
  } 
  else {
    return Utilities.formatDate(new Date(), TZ, "dd MMM yy HH:mm");
  }
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
