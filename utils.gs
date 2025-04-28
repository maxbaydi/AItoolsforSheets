// Вспомогательные функции
function logMessage(message, isError = false) {
    if (!message || typeof message !== 'string') {
        throw new Error('Сообщение для логирования должно быть непустой строкой');
    }
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = spreadsheet.getSheetByName(LOG_SHEET_NAME);
    
    // Если лист логов не существует, создаем его и сразу скрываем
    if (!logSheet) {
        logSheet = spreadsheet.insertSheet(LOG_SHEET_NAME);
        logSheet.hideSheet(); // Скрываем лист
    }
    
    if (logSheet.getLastRow() === 0) {
        logSheet.appendRow(["Время", "Сообщение", "Ошибка"]);
        logSheet.setFrozenRows(1);
    }
    logSheet.appendRow([new Date(), message, isError]);
}

function calculateMD5(input) {
    if (!input) {
        throw new Error('Входные данные для расчета MD5 не могут быть пустыми');
    }
    const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, input);
    let txtHash = '';
    for (let i = 0; i < rawHash.length; i++) {
        let hashVal = rawHash[i];
        if (hashVal < 0) {
            hashVal += 256;
        }
        if (hashVal.toString(16).length === 1) {
            txtHash += '0';
        }
        txtHash += hashVal.toString(16);
    }
    return txtHash;
}

/**
 * Возвращает целевые заголовки активного листа из указанной строки (для вызова из клиента)
 * @param {number} headerRow
 * @returns {string[]}
 */
function getTargetHeadersFromServer(headerRow) {
    try {
        if (typeof headerRow !== 'number' || headerRow < 1) {
            throw new Error('Номер строки заголовков должен быть положительным числом');
        }
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
        const lastColumn = sheet.getLastColumn();
        if (lastColumn === 0) return [];
        const headers = sheet.getRange(headerRow, 1, 1, lastColumn).getValues()[0];
        return headers.map(String);
    } catch (e) {
        logMessage(`Ошибка при получении заголовков: ${e.message}`, true);
        return [];
    }
}

/**
 * Проверяет наличие и доступность сервисов Google Apps Script
 * @returns {object} Объект с информацией о доступности сервисов
 */
function checkAvailableServices() {
    const services = {
        "DriveApp": typeof DriveApp !== 'undefined',
        "Drive": typeof Drive !== 'undefined',
        "SpreadsheetApp": typeof SpreadsheetApp !== 'undefined',
        "DocumentApp": typeof DocumentApp !== 'undefined',
        "CacheService": typeof CacheService !== 'undefined',
        "Utilities": typeof Utilities !== 'undefined'
    };
    
    // Дополнительная проверка Drive API
    let driveApiVersion = "Недоступно";
    try {
        if (services.Drive) {
            driveApiVersion = "Доступно";
            // Попытка получить информацию о версии API или выполнить базовую операцию
            const rootFolder = DriveApp.getRootFolder();
            driveApiVersion += `, ID корневой папки: ${rootFolder.getId()}`;
        }
    } catch (e) {
        driveApiVersion = `Ошибка: ${e.toString()}`;
    }
    
    services["Drive API Version"] = driveApiVersion;
    
    return services;
}
