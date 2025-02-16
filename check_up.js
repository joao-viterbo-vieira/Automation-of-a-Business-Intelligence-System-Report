function check_And_Update_Data() {
    var sheetName = "xxxxx"; // Substituir pelo nome da folha
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    var lastRow = sheet.getLastRow();
    var numRowsToCheck = 10000;

    // Ajustar para a quantidade total de linhas se lastRow for menor que numRowsToCheck
    if (lastRow < numRowsToCheck) {
        numRowsToCheck = lastRow;
    }

    var todayMinus2 = new Date();
    todayMinus2.setDate(todayMinus2.getDate() - 2);
    var dataRange = sheet.getRange(lastRow - numRowsToCheck + 1, 5, numRowsToCheck, 1);
    var dataValues = dataRange.getValues();
    var isDateInLast20Rows = false;

    for (var i = 0; i < numRowsToCheck; i++) {
        var rowDate = dataValues[i][0];
        if (rowDate.toDateString() === todayMinus2.toDateString()) {
            isDateInLast20Rows = true;
            break;
        }
    }

    if (!isDateInLast20Rows) {
        update_Sheet_With_CSV();
    }
}