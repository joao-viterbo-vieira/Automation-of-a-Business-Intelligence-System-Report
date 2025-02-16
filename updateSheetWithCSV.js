function update_Sheet_With_CSV() {
    var spreadSheetId = 'xxxxxxxx';
    var sheetName = 'xxxxxxxx'; // Caso o Sheetname troque de nome é necessário atualizar aqui

    // Obtém a data de hoje e calcula o dia anterior
    var today = new Date();
    var yesterday = new Date(today);
    yesterday.setDate(yesterday.getDate() - 2);

    // Formata as datas no formato "yyyy/M/d"
    var formattedToday = Utilities.formatDate(today, "GMT", "yyyy/M/d");
    var formattedYesterday = Utilities.formatDate(yesterday, "GMT", "yyyy/M/d");

    duplicar_Folha_E_Inserir_Data('Template', yesterday, spreadSheetId);
    Logger.log('Data inicio: ' + get_Latest_Date_Above_LastRow());
    Logger.log('Data fim: ' + formattedYesterday);

    // Recupera a API Key e constrói a URL da API
    var Key_API = PropertiesService.getScriptProperties().getProperty('Key_API');
    var apiUrl =
        'https://xxxxxxxxxxxxxxxxxxx.com/api/xxxxxxxx.asp?key=' + Key_API +
        '&reportname=Member&reportformat=csv&reportmember=all&startdate=' +
        get_Latest_Date_Above_LastRow() + '&enddate=' + formattedYesterday;

    // Busca e processa o CSV
    var response = UrlFetchApp.fetch(apiUrl);
    var csvData = response.getContentText();
    var sheet = SpreadsheetApp.openById(spreadSheetId).getSheetByName(sheetName);

    var csv = Utilities.newBlob(csvData).getDataAsString();
    var parsedData = Utilities.parseCsv(csv);

    // Remove as linhas de cabeçalho (6 linhas) e a última linha desnecessária
    for (var j = 0; j < 6; j++) {
        parsedData.shift();
    }
    parsedData.pop();

    var numRows = parsedData.length;
    var lastRow = sheet.getRange("A" + sheet.getLastRow()).getRow();

    // Arrays para armazenar os dados a inserir
    var valuesToFill = [];
    var valuesToFillE = [];
    var valuesToFillO = [];
    var valuesToFillF = [];
    var valuesToFillI = [];
    var valuesToFillJ = [];

    // Processa cada linha do CSV
    for (var i = 0; i < numRows; i++) {
        var row = parsedData[i];

        var dateValue = contar_Dias_Desde_1900(row[3]);
        valuesToFill.push([dateValue + "|" + row[9]]);

        var formattedDate = convert_To_DDMMYYYY(row[3]);
        valuesToFillE.push([formattedDate]);

        valuesToFillO.push([row[15]]);
        valuesToFillF.push([row[20]]);
        valuesToFillI.push([row[9]]);
        valuesToFillJ.push([row[10]]);
    }

    // Define os intervalos para inserir os dados
    var rangeColumnA = sheet.getRange(lastRow + 1, 1, numRows, 1);
    var rangeColumnE = sheet.getRange(lastRow + 1, 5, numRows, 1); // Coluna E
    var rangeColumnO = sheet.getRange(lastRow + 1, 15, numRows, 1); // Coluna O
    var rangeColumnF = sheet.getRange(lastRow + 1, 20, numRows, 1); // Coluna F
    var rangeColumnI = sheet.getRange(lastRow + 1, 9, numRows, 1);  // Coluna I
    var rangeColumnJ = sheet.getRange(lastRow + 1, 10, numRows, 1); // Coluna J

    // Insere os dados nas colunas correspondentes
    rangeColumnA.setValues(valuesToFill);
    rangeColumnE.setValues(valuesToFillE);
    rangeColumnO.setValues(valuesToFillO);
    rangeColumnF.setValues(valuesToFillF);
    rangeColumnI.setValues(valuesToFillI);
    rangeColumnJ.setValues(valuesToFillJ);

    Logger.log('Ultima linha: ' + lastRow + " + " + numRows);
}
