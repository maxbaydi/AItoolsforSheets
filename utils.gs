/**
 * Записывает сообщение в лог с отметкой времени
 * @param {string} message - сообщение для записи в лог
 * @param {boolean} isError - флаг ошибки (по умолчанию false)
 * @param {number} level - уровень детализации лога (1-5, где 5 - максимальная детализация)
 */
function logMessage(message, isError = false, level = 3) {
  try {
    const timestamp = new Date().toLocaleString('ru-RU');
    const logType = isError ? '[ОШИБКА]' : '[ИНФО]';
    const fullMessage = `${timestamp} ${logType} ${message}`;
    
    // Получаем текущие настройки логирования
    const properties = PropertiesService.getScriptProperties();
    const logLevel = parseInt(properties.getProperty('LOG_LEVEL') || '3');
    
    // Записываем сообщение только если его уровень не превышает настроенный уровень логирования
    if (level <= logLevel) {
      console.log(fullMessage);
      
      // Также записываем в таблицу логов, если включено
      const sheet = SpreadsheetApp.getActiveSpreadsheet();
      if (sheet) {
        // Используем оригинальное имя листа и делаем его скрытым
        let logsSheet = sheet.getSheetByName('__AI_LOGS__');
        if (!logsSheet) {
          // Создаем скрытый лист для логов
          logsSheet = sheet.insertSheet('__AI_LOGS__');
          logsSheet.hideSheet();
        }
        logsSheet.appendRow([timestamp, message, isError]);
      }
    }
  } catch (e) {
    // Если возникла ошибка при логировании, выводим в консоль
    console.log(`Ошибка логирования: ${e.toString()}`);
  }
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

/**
 * Грубо оценивает количество токенов в тексте
 * Оценка основана на предположении, что 1 токен ≈ 4 символа для латиницы и 
 * примерно 2-3 символа для кириллицы
 * @param {string} text - Текст для оценки
 * @returns {number} - Приблизительное количество токенов
 */
function estimateTokenCount(text) {
  if (!text) return 0;
  
  // Разделяем текст на части с латиницей и кириллицей
  const latinChars = (text.match(/[a-zA-Z0-9.,;:?!()\[\]{}"'`~@#$%^&*_+=<>|\\/-]/g) || []).length;
  const cyrillicChars = (text.match(/[а-яА-ЯёЁ]/g) || []).length;
  const otherChars = text.length - latinChars - cyrillicChars;
  
  // Используем разные коэффициенты для разных алфавитов
  // Для латиницы примерно 4 символа на токен
  // Для кириллицы примерно 2-3 символа на токен (берем 2.5)
  // Для других символов примерно 3 символа на токен
  const latinTokens = latinChars / 4.0;
  const cyrillicTokens = cyrillicChars / 2.5;
  const otherTokens = otherChars / 3.0;
  
  // Учитываем пробелы и переносы строк
  const whitespaceChars = (text.match(/\s+/g) || []).join('').length;
  const whitespaceTokens = whitespaceChars / 5.0; // Примерно 5 пробелов на токен
  
  // Получаем общее количество токенов и округляем до целого
  const totalTokens = Math.ceil(latinTokens + cyrillicTokens + otherTokens + whitespaceTokens);
  
  // Для больших текстов добавляем небольшой запас на сложность контекста
  const complexityFactor = Math.min(1.1, 1 + (text.length / 100000) * 0.1);
  
  return Math.ceil(totalTokens * complexityFactor);
}
