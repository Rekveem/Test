function getJSONArr(mainSheet) {
  var sheetData = mainSheet.getSheetByName('Лист1');
  var lastRow = sheetData.getLastRow();
  var lastColumn = sheetData.getLastColumn();
  var arrayData = sheetData.getSheetValues(1,1,lastRow,lastColumn);
  var headers  = arrayData[0];
  var ignoreHeaders = ["Date", "GachaBalanceType", "RouletteBalanceType"];
  var result = {};
  
  for (let headerIndex = 0; headerIndex<lastColumn; headerIndex++) {
    var header = headers[headerIndex];
    if (ignoreHeaders.includes(header) || header.length<=0 || header == null || header == "") {
      continue
    } 

   var geoResult = [];
   for (let rowIndex = 1; rowIndex<lastRow; rowIndex++) {
    var cellValues = arrayData[rowIndex][headerIndex];
    if (cellValues == null || cellValues.length<=0 || cellValues == ""){ 
      continue
    }
    var cellValuesArray = cellValues.split(",");
    var count = cellValuesArray.length;
    for (let i = 0; i < count; i++) {
      var cellValue = cellValuesArray[i];
      var date = Utilities.formatDate(arrayData[rowIndex][headers.indexOf("Date")], "GMT+3", "MM/dd/yyyy");
      var offerObject = geoResult.findLast(offerObjectFind=>offerObjectFind['balance'] == cellValue)
      if (offerObject == null || (rowIndex>1 && !arrayData[rowIndex-1][headerIndex].includes(cellValue) )) {
        offerObject = {};
        offerObject['start_date'] = date;
        offerObject['balance'] = cellValue;

  
        geoResult.push(offerObject);
      }
      else {
        offerObject['end_date'] = date;
      }
    }
   }
   result[header + "_Balance"] = geoResult;

  }
 Logger.log(JSON.stringify(result));
}

// Функция для обработки запроса doGet
function doGet(e) {
    try {
        // Открываем таблицу по её ID
        const mainSheet = SpreadsheetApp.openById(String(e.parameter.id));
        // Получаем JSON данные и возвращаем их
        return ContentService.createTextOutput(getJSONArr(mainSheet)).setMimeType(ContentService.MimeType.JSON);
    } catch (error) {
        console.error('Error:', error);
        return ContentService.createTextOutput(JSON.stringify({ error: 'An error occurred.' })).setMimeType(ContentService.MimeType.JSON);
    }
}

// Пример вызова функции doGetTest
function doGetTest() {
    try {
        // Открываем тестовую таблицу по её ID
        const mainSheet = SpreadsheetApp.openById("1ExICg-I77Ib7urQ_smxtOnma-833I1De-Eu_GcZCy68");
        // Получаем JSON данные и возвращаем их
        return ContentService.createTextOutput(getJSONArr(mainSheet)).setMimeType(ContentService.MimeType.JSON);
    } catch (error) {
        console.error('Error:', error);
        return ContentService.createTextOutput(JSON.stringify({ error: 'An error occurred.' })).setMimeType(ContentService.MimeType.JSON);
    }
}