var ss = SpreadsheetApp.getActiveSpreadsheet();
var conf = {
  calendario: 'B1',
  notifiche: 'B2',
  titolo: 'B4',
  descrizione: 'B5'
};

function _getConf(prop) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('impostazioni').getRange(conf[prop]).getValue();
}
function _getNomeCalendario() {
  return _getConf('calendario');
}
function _getCalendar() {
  return CalendarApp.getCalendarById(_getNomeCalendario());
}

function onOpen() {
  if (_getNomeCalendario() == "") {
    ss.toast("Click on setup to start", "Notifiche", 10);
  }
  var menuEntries = [ {name: "Aggiorna calendario", functionName: "importIntoCalendar"} ];
  ss.addMenu("Calendario", menuEntries);
}

function importIntoCalendar(){
  var calendario = _getNomeCalendario();
  if (!calendario) {
    return;
  }
  var dataSheet = ss.getSheetByName("tavoli");
  var dataRange = dataSheet.getRange(2, 1, dataSheet.getMaxRows(), dataSheet.getMaxColumns());

  var cal =  _getCalendar();
  var TZ = cal.getTimeZone();
  var eventTitleTemplate = _getConf('titolo');
  var descriptionTemplate = _getConf('descrizione');

  var banchetti = getRowsData(dataSheet, dataRange);
  for (var i = 0; i < banchetti.length; ++i) {
    var banchetto = banchetti[i];
    if (!banchetto.confermato){
      continue;
    }
    var start = _toDate(banchetto.data, banchetto.da); 
    var end = _toDate(banchetto.data, banchetto.a);

    _geocode(dataSheet, i+2);
    
    var event;
    if (banchetto.idCalendario) {
      event = _getEvent(cal, start, end, banchetto.idCalendario);
      if (event) {
        event.deleteEvent();
      }
    }
    
    Logger.log('Crea evento #' + i + ' ...');
    var eventTitle = fillInTemplateFromObject(eventTitleTemplate, banchetto);
    var description = fillInTemplateFromObject(descriptionTemplate, banchetto);
    var eventId = cal.createEvent(eventTitle, start, end, {location:banchetto.luogo, description: description}).getId();
    Logger.log('Evento #' + i + ' creato con id [' + eventId + ']');
    
    event = _getEvent(cal, start, end, eventId);
    if (!event) {
      Browser.msgBox("Evento non trovato - function getEvents");
    }
    
    dataSheet.getRange(i+2, 1, 1, 1).setValue(eventId);
    //"Added "+ Utilities.formatDate(new Date(), TZ, "dd MMM yy HH:mm")
    dataSheet.getRange(i+2, 2, 1, 1).setValue(new Date()).setBackgroundRGB(221, 221, 221);
    dataSheet.getRange(i+2,1,1,dataSheet.getMaxColumns()).setBackgroundRGB(221, 221, 221);
    // Make sure the cell is updated right away in case  the script is interrupted
    SpreadsheetApp.flush();
  }
  ss.toast("People can now register to those events", "Events imported");
}

function _geocode(cells, row) {
  var addressColumn = 5;
  var addressRow = row;
  
  var latColumn = addressColumn + 1;
  var lngColumn = addressColumn + 2;
  
  var geocoder = Maps.newGeocoder().setRegion('it');
  var location;
  
  var address = cells.getRange(addressRow, addressColumn).getValue();
  Logger.log(address);
  // Geocode the address and plug the lat, lng pair into the
  // 2nd and 3rd elements of the current range row.
  location = geocoder.geocode(address);
  Logger.log(location);
  // Only change cells if geocoder seems to have gotten a
  // valid response.
  if (location.status == 'OK') {
    lat = location["results"][0]["geometry"]["location"]["lat"];
    lng = location["results"][0]["geometry"]["location"]["lng"];
    
    cells.getRange(addressRow, latColumn).setValue(lat);
    cells.getRange(addressRow, lngColumn).setValue(lng);
  }
}

function _toDate(data, ora) {
  var datetime;
  if (data instanceof Date) {
    datetime = new Date(data.getTime());
    datetime.setHours(_toHours(ora));
    datetime.setMinutes(_toMinutes(ora));
  } else {
    var parts = data.split('/');
    new Date(parts[2], parts[1], parts[0], _toHours(ora), _toMinutes(ora));
  }
  return datetime;
}
function _toHours(ora) {
  var hours = ora;
  if (String(ora).indexOf(':') !== -1) {
    hours = String(ora).split(':')[0];
  }
  return hours;
}
function _toMinutes(ora) {
  var minutes = 0;
  if (String(ora).indexOf(':') !== -1) {
    minutes = String(ora).split(':')[1];
  }
  return minutes;
}

function _getEvent(cal, from, to, eventId) {
  var events = cal.getEvents(from, to);
  return getEvent(events, eventId);
}

// http://code.google.com/p/google-apps-script-issues/issues/detail?id=264#c35
function getEvent(events, eventID) {
  for (var i in events) {
    if (events[i].getId() == eventID) {
      var event = events[i];
      return event;
    }
  }
  //Browser.msgBox("Event not found - function getEvents");
  //return event; //last one!?
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

  if (!templateVars) {
    return email;
  }
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
