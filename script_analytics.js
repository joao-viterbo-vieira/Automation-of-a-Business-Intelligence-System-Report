function duplicar_Folha_E_Inserir_Data(SheetName, date, spreadSheetId) {
    // Define the name of the new Sheet based on the date
    var nomeNovaFolha = getNomeMesProximoAno(date);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    // Check if a Sheet with the same name already exists
    var existingSheet = ss.getSheetByName(nomeNovaFolha);
    if (existingSheet) {
        Logger.log("Sheet with name '" + nomeNovaFolha + "' already exists. Skipping...");
        return;
    }
    var folhaOrigem = SpreadsheetApp.openById(spreadSheetId).getSheetByName(SheetName);
    var novaFolha = folhaOrigem.copyTo(ss);
    novaFolha.setName(nomeNovaFolha);
    var primeiroDiaProximoMes = new Date(date.getFullYear(), date.getMonth(), 1 + 1);
    var dataFormatada = Utilities.formatDate(primeiroDiaProximoMes, 'GMT', 'd/MM/yyyy');
    novaFolha.getRange('M2').setValue(dataFormatada);
}

function get_Nome_Mes_Proximo_Ano(date) {
    var proximoMes = new Date(date.getFullYear(), date.getMonth());
    var nomeMeses = ['JANEIRO', 'FEVEREIRO', 'MARÇO', 'ABRIL', 'MAIO', 'JUNHO', 'JULHO',
        'AGOSTO', 'SETEMBRO', 'OUTUBRO', 'NOVEMBRO', 'DEZEMBRO'
    ];
    var nomeMes = nomeMeses[proximoMes.getMonth()];
    return "CALENDÁRIO " + nomeMes + ' ' + proximoMes.getFullYear();
}

function contar_Dias_Desde_1900(data) {
    var dataBase = new Date("1900-01-01");
    var dataFornecida = new Date(data);
    var diferencaEmMilissegundos = dataFornecida.getTime() - dataBase.getTime();
    // Converte a diferença em dias
    var diasDesde1900 = Math.floor(diferencaEmMilissegundos / (1000 * 60 * 60 * 24) + 3);
    return diasDesde1900;
}

function convert_To_DDMMYYYY(dateString) {
    // Verifica se a string está no formato mm/dd/yyyy
    var dateRegex = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/;
    var match = dateString.match(dateRegex);
    if (!match) {
        throw new Error("A data fornecida não está no formato mm/dd/yyyy.");
    }
    var month = Number(match[1]);
    var day = Number(match[2]);
    var year = Number(match[3]);
    // Formata a data para dd/mm/yyyy
    var formattedDate = Utilities.formatString('%02d/%02d/%04d', day, month, year);
    return formattedDate;
}

function get_Latest_Date_Above_LastRow() {
    var spreadSheetId = 'xxxxxxxx';
    var SheetName = 'xxxxxxxx';
    var dateColumnIndex = 5; // Column E
    var valuesAboveLastRow = 500;
    var Sheet = SpreadsheetApp.openById(spreadSheetId).getSheetByName(SheetName);
    var lastRow = Sheet.getLastRow();
    var range = Sheet.getRange(lastRow - valuesAboveLastRow + 1, dateColumnIndex, valuesAboveLastRow, 1);
    var values = range.getValues();
    var latestDate = new Date(0);
    for (var i = 0; i < valuesAboveLastRow; i++) {
        var cellValue = values[i][0];
        if (cellValue instanceof Date && cellValue > latestDate) {
            latestDate = cellValue;
        }
    }
    var nextDay = new Date(latestDate);
    nextDay.setDate(nextDay.getDate() + 1);
    var todayMinus2 = new Date();
    todayMinus2.setDate(todayMinus2.getDate() - 2);
    if (nextDay > todayMinus2) {
        return Utilities.formatDate(todayMinus2, "GMT", "yyyy/M/d");
    } else {
        nextDay.setDate(nextDay.getDate() + 1);
        return Utilities.formatDate(nextDay, "GMT", "yyyy/M/d");
    }
}