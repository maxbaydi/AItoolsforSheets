/**
 * @fileoverview Содержит функции для обработки файлов.
 * Сюда следует перенести соответствующую функциональность из Code.gs.
 */

/**
 * Обрабатывает загруженный файл, преобразует его во временную Google Таблицу
 * и возвращает ID таблицы, имена листов (для табличных файлов) и тип файла.
 * @param {object} fileData Объект с данными файла {name, type, data (base64), pageRange?}.
 * @returns {{ sheetNames: string[], tempSheetId: string, fileType: string } | { error: string }}
 */
function processUploadedFile(fileData) {
  let tempSheetId;
  let fileId; // ID временного файла (если создавался)
  let tempDocId; // ID временного Google Doc (если создавался)

  try {
    // Проверяем наличие обязательных полей
    if (!fileData || !fileData.name || !fileData.type || !fileData.data) {
      throw new Error("Отсутствуют обязательные параметры файла (name, type, data)");
    }
    
    const mimeType = fileData.type;
    // Определяем тип файла по MIME-type и расширению
    const fileExtension = fileData.name.split('.').pop().toLowerCase();
    const isTextFile = ['doc', 'docx', 'txt'].includes(fileExtension) || mimeType === 'text/plain' || mimeType === 'application/msword' || mimeType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
    // Определяем, является ли файл табличным (XLSX, XLS, ODS, CSV)
    const isSpreadsheetFile = ['xlsx', 'xls', 'ods', 'csv'].includes(fileExtension) ||
                              mimeType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
                              mimeType === 'application/vnd.ms-excel' ||
                              mimeType === 'application/vnd.oasis.opendocument.spreadsheet' ||
                              mimeType === 'text/csv';

    logMessage(`processUploadedFile: Начало обработки файла "${fileData.name}", тип: ${mimeType}, расширение: ${fileExtension}, текстовый: ${isTextFile}, табличный: ${isSpreadsheetFile}`);

    // Проверяем допустимый тип файла
    if (!isTextFile && !isSpreadsheetFile) {
      throw new Error(`Неподдерживаемый тип файла: ${fileData.name} (MIME: ${mimeType}, Ext: ${fileExtension})`);
    }
    
    // Проверяем размер данных
    const dataLength = fileData.data.length;
    if (dataLength > 10 * 1024 * 1024) { // ~10 МБ в base64 (примерно)
      logMessage(`processUploadedFile: Предупреждение - большой размер данных (${Math.round(dataLength/1024/1024 * 100) / 100} MB)`, true);
    }
    
    try {
      // Безопасное создание Blob с обработкой ошибок
      logMessage(`processUploadedFile: Попытка декодирования base64 данных и создания Blob...`);
      
      // Проверяем валидность base64
      if (!isValidBase64(fileData.data)) {
        throw new Error("Невалидные данные Base64");
      }
      
      // Декодируем и создаем blob по частям для больших файлов
      let decodedData;
      try {
        decodedData = Utilities.base64Decode(fileData.data);
        logMessage(`processUploadedFile: Данные успешно декодированы из Base64, размер: ${decodedData.length} байт`);
      } catch (decodeError) {
        logMessage(`processUploadedFile: Ошибка при декодировании Base64: ${decodeError}`, true);
        throw new Error(`Ошибка декодирования данных: ${decodeError}`);
      }
      
      let blob;
      try {
        blob = Utilities.newBlob(decodedData, mimeType, fileData.name);
        logMessage(`processUploadedFile: Blob успешно создан, размер: ${blob.getBytes().length} байт`);
      } catch (blobError) {
        logMessage(`processUploadedFile: Ошибка при создании Blob: ${blobError}`, true);
        throw new Error(`Ошибка создания Blob: ${blobError}`);
      }
      
      const folder = DriveApp.getRootFolder(); // Можно заменить на временную папку

      if (isSpreadsheetFile) {
        // --- Обработка табличных файлов: прямое конвертирование Blob в Google Sheets ---
        logMessage(`processUploadedFile: Прямое конвертирование Blob в Google Таблицу для ${fileData.name}`);
        try {
          const baseName = fileData.name.replace(/\.[^/.]+$/, '');
          const resource = { title: baseName + '_converted', mimeType: MimeType.GOOGLE_SHEETS };
          const sheetFile = Drive.Files.insert(resource, blob, { convert: true });
          tempSheetId = sheetFile.id;
          logMessage(`processUploadedFile: Временная Google Таблица создана: ${tempSheetId}`);
        } catch (e) {
          logMessage(`processUploadedFile: Ошибка прямой конвертации Blob: ${e}`, true);
          throw new Error(`Не удалось конвертировать табличный файл: ${e}`);
        }

      } else if (isTextFile) {
        // --- Обработка текстовых файлов ---
        if (fileExtension === 'txt' || mimeType === 'text/plain') {
          logMessage(`processUploadedFile: Обработка как TXT файл.`);
          // TXT файл преобразуем сразу в таблицу с одной ячейкой
          // pageRange для TXT не имеет смысла, но функция его принимает
          tempSheetId = createTempSpreadsheetFromBlob(blob, fileData.name, fileData.pageRange);
          logMessage(`processUploadedFile: Временная Google Таблица из TXT создана: ${tempSheetId}`);
        } else { // doc, docx
          logMessage(`processUploadedFile: Обработка как DOC/DOCX файл.`);
          const tempFile = folder.createFile(blob);
          fileId = tempFile.getId();
          logMessage(`processUploadedFile: Временный DOC/DOCX файл создан: ${fileId}`);
          tempDocId = convertFileToGoogleDoc(tempFile, folder);
          logMessage(`processUploadedFile: Временный Google Doc создан: ${tempDocId}`);
          tempSheetId = convertGoogleDocToSpreadsheet(tempDocId, fileData.pageRange); // Передаем pageRange
          logMessage(`processUploadedFile: Временная Google Таблица из DOC/DOCX создана: ${tempSheetId}`);
          // Удаляем временные файлы DOC/DOCX и Google Doc
          try { DriveApp.getFileById(fileId).setTrashed(true); logMessage(`processUploadedFile: Временный файл ${fileId} удален.`); } catch(e) { logMessage(`Не удалось удалить временный файл ${fileId}: ${e}`, true); }
          try { DriveApp.getFileById(tempDocId).setTrashed(true); logMessage(`processUploadedFile: Временный Google Doc ${tempDocId} удален.`); } catch(e) { logMessage(`Не удалось удалить временный Google Doc ${tempDocId}: ${e}`, true); }
          fileId = null; // Сбрасываем
          tempDocId = null; // Сбрасываем
        }
      }

      // Проверка, что tempSheetId был создан
      if (!tempSheetId) {
        throw new Error('Не удалось создать временную Google Таблицу.');
      }

      // Получаем имена листов (только для табличных файлов)
      let sheetNames = [];
      if (isSpreadsheetFile) { // Используем isSpreadsheetFile вместо !isTextFile для ясности
          try {
             const tempSpreadsheet = SpreadsheetApp.openById(tempSheetId);
             // Иногда сразу после создания таблица может быть недоступна, добавим паузу
             Utilities.sleep(1000);
             sheetNames = tempSpreadsheet.getSheets().map(sheet => sheet.getName());
             logMessage(`processUploadedFile: Получены имена листов: ${JSON.stringify(sheetNames)}`);
             // Если лист всего один и называется "Лист1", вернем пустое имя, чтобы не смущать пользователя
             if (sheetNames.length === 1 && sheetNames[0] === "Лист1") {
               // Не меняем sheetNames, пусть будет "Лист1", т.к. его надо будет передать
               logMessage(`processUploadedFile: Обнаружен один лист "Лист1".`);
             }
          } catch (e) {
             logMessage(`Не удалось получить имена листов для ${tempSheetId}: ${e}. Возвращаем ["Лист1"]`, true);
             sheetNames = ["Лист1"]; // Возвращаем дефолтное имя в случае ошибки
          }
      } else {
          // Для текстовых файлов имя листа не нужно, возвращаем пустой массив
          sheetNames = [];
          logMessage(`processUploadedFile: Файл текстовый, имена листов не извлекались.`);
      }

      logMessage(`processUploadedFile: Завершено. Возвращаем: sheetNames=${JSON.stringify(sheetNames)}, tempSheetId=${tempSheetId}, fileType=${isTextFile ? 'text' : 'table'}`);
      // Возвращаем ID временной таблицы, имена листов и тип файла
      return { sheetNames: sheetNames, tempSheetId: tempSheetId, fileType: isTextFile ? 'text' : 'table' };

    } catch (innerError) {
      logMessage(`processUploadedFile: Внутренняя ошибка при обработке файла: ${innerError}`, true);
      throw new Error(`Внутренняя ошибка при обработке файла: ${innerError}`);
    }

  } catch (error) {
    logMessage(`Ошибка в processUploadedFile: ${error.toString()} ${error.stack}`, true);
    // Попытаемся удалить временные файлы, если они были созданы
    if (fileId) try { DriveApp.getFileById(fileId).setTrashed(true); logMessage(`Удален временный файл ${fileId} после ошибки.`); } catch(e) {}
    if (tempDocId) try { DriveApp.getFileById(tempDocId).setTrashed(true); logMessage(`Удален временный GDoc ${tempDocId} после ошибки.`); } catch(e) {}
    // tempSheetId удаляется в analyzeAndInsertExtractedData
    return { error: `Ошибка при обработке файла: ${error.toString()}` };
  }
}


/**
 * Создает временную таблицу из Blob (для CSV, TXT, ODS).
 * @param {GoogleAppsScript.Base.Blob} blob
 * @param {string} fileName
 * @param {string} [pageRange] Диапазон страниц (для текстовых файлов, если применимо).
 * @returns {string} ID созданной временной таблицы.
 */
 function createTempSpreadsheetFromBlob(blob, fileName, pageRange) {
    const folder = DriveApp.getRootFolder();
    const tempSpreadsheet = SpreadsheetApp.create("Temp Sheet " + fileName + Date.now());
    const sheet = tempSpreadsheet.getActiveSheet();

    if (blob.getContentType() === 'text/csv') {
        const data = Utilities.parseCsv(blob.getDataAsString());
        sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    } else if (blob.getContentType() === 'text/plain') {
      let text = blob.getDataAsString();
      // Разбиваем текст на строки, чтобы избежать слишком большого значения в одной ячейке
      const lines = text.split(/\r?\n/);
      const values = lines.map(line => [line]);
      sheet.getRange(1, 1, values.length, 1).setValues(values);

    } else if (blob.getContentType() === 'application/vnd.oasis.opendocument.spreadsheet') {
        const file = folder.createFile(blob);
         const tempSheet = Drive.Files.copy({
            mimeType: MimeType.GOOGLE_SHEETS,
             parents: [{ id: folder.getId() }],
              title: file.getName() + "_temp"
            }, file.getId());

        DriveApp.getFileById(file.getId()).setTrashed(true); // Удаляем
        return tempSheet.getId();
    }

    return tempSpreadsheet.getId();
}

/**
 * Вспомогательная функция для преобразования файла в Google Таблицу.
 * Использует улучшенный метод с проверками и обработкой ошибок для решения проблемы PERMISSION_DENIED.
 * Адаптирована для работы с Drive API v3.
 * @param {GoogleAppsScript.Drive.File} file
 * @param {GoogleAppsScript.Drive.Folder} folder
 * @returns {string} ID созданной Google Таблицы.
 */
function convertFileToSpreadsheet(file, folder) {
    // Упрощенный и надёжный метод конвертации
    const fileName = file.getName();
    const ext = fileName.split('.').pop().toLowerCase();
    logMessage(`convertFileToSpreadsheet: Выбрано расширение ${ext}`);
    if (['xlsx', 'xls', 'ods'].includes(ext)) {
        try {
            logMessage(`convertFileToSpreadsheet: Конвертация бинарного файла через Drive.Files.insert: ${fileName}`);
            const baseName = fileName.replace(/\.[^/.]+$/, '');
            const blob = file.getBlob();
            const resource = { title: baseName + '_converted', parents: [{ id: folder.getId() }] };
            const sheetFile = Drive.Files.insert(resource, blob, { convert: true });
            logMessage(`convertFileToSpreadsheet: Успешно создан Google Sheet ID=${sheetFile.id}`);
            return sheetFile.id;
        } catch (e) {
            logMessage(`convertFileToSpreadsheet: Ошибка конвертации бинарного файла: ${e}`, true);
            throw new Error(`Не удалось конвертировать файл ${fileName}: ${e}`);
        }
    }
    // Обработка CSV и TXT
    if (ext === 'csv' || file.getMimeType() === 'text/plain') {
        logMessage(`convertFileToSpreadsheet: Обработка CSV/TXT файла ${fileName}`);
        const temp = SpreadsheetApp.create(`Temp_${fileName}_${Date.now()}`);
        const data = Utilities.parseCsv(file.getBlob().getDataAsString());
        temp.getActiveSheet().getRange(1,1,data.length,data[0].length).setValues(data);
        return temp.getId();
    }
    throw new Error(`Конвертация для расширения .${ext} не поддерживается`);
}

/**
 * Преобразует файл (предположительно DOCX) в Google Doc.
 * Обновлено для работы с Drive API v3.
 * @param {GoogleAppsScript.Drive.File} file
 * @param {GoogleAppsScript.Drive.Folder} folder
 * @returns {string} ID созданного Google Doc.
 */
function convertFileToGoogleDoc(file, folder) {
    try {
        // Логируем для отладки
        logMessage(`convertFileToGoogleDoc: Начало конвертации файла ${file.getName()}, ID: ${file.getId()}`);
        
        // Конвертация DOC/DOCX в Google Doc через Drive API v3
        const originalFileId = file.getId();
        const baseName = file.getName().replace(/\.[^/.]+$/, '');
        
        // Создаем метаданные для нового файла в формате Drive API v3
        const resource = {
            name: baseName + '_temp_doc',
            mimeType: MimeType.GOOGLE_DOCS,
            parents: [folder.getId()]
        };
        
        // Копируем файл в формат Google Doc с использованием API v3
        logMessage(`convertFileToGoogleDoc: Вызываем Drive.Files.copy с API v3`);
        
        const copied = Drive.Files.copy(resource, originalFileId);
        const docId = copied.id;
        
        // Даем время на завершение конвертации
        Utilities.sleep(2000);
        
        // Удаляем исходный временный файл через API v3
        try {
            // Пробуем через Drive API v3
            Drive.Files.update({trashed: true}, originalFileId);
            logMessage(`convertFileToGoogleDoc: Исходный файл помечен на удаление через Drive API v3`);
        } catch (e) {
            // Запасной вариант через DriveApp
            try {
                DriveApp.getFileById(originalFileId).setTrashed(true);
                logMessage(`convertFileToGoogleDoc: Исходный файл помечен на удаление через DriveApp`);
            } catch (driveError) {
                logMessage(`convertFileToGoogleDoc: Не удалось удалить исходный DOC/DOCX ${originalFileId}: ${driveError}`, true);
            }
        }
        
        logMessage(`convertFileToGoogleDoc: Успешно создан Google Doc с ID: ${docId}`);
        return docId;
    } catch (error) {
        logMessage(`Ошибка в convertFileToGoogleDoc: ${error}`, true);
        throw new Error(`Не удалось конвертировать файл в Google Doc: ${error.message || error}`);
    }
}

/**
 * Преобразует Google Doc в Google Sheet, извлекая текст (с учетом pageRange).
 * @param {string} docId ID Google Doc.
 * @param {string} [pageRange] Диапазон страниц для извлечения (например, "1,3-5").
 * @returns {string} ID созданной Google Таблицы.
 */
function convertGoogleDocToSpreadsheet(docId, pageRange) {
    // Открываем Google Doc и ждём завершения конвертации
    const doc = DocumentApp.openById(docId);
    Utilities.sleep(2000);
    
    // Создаем временную таблицу
    const tempSpreadsheet = SpreadsheetApp.create("Temp Sheet from Doc " + Date.now());
    const sheet = tempSpreadsheet.getActiveSheet();
    
    try {
        logMessage(`convertGoogleDocToSpreadsheet: Начинаем извлечение текста из документа ${docId}`);
        
        // Логика извлечения текста с учетом pageRange
        if (pageRange) {
            const pages = pageRange.split(','); // Разбираем строку диапазона
            const body = doc.getBody();
            // Получаем все дочерние элементы через getNumChildren/getChild
            const elements = [];
            const count = body.getNumChildren();
            
            // Извлекаем текст по частям для больших документов
            let rowIndex = 1;
            let currentPage = 0;
            let pageBreakFound = false;
            
            for (let i = 0; i < count; i++) {
                try {
                    const element = body.getChild(i);
                    let elementType = element.getType();
                    let isPageBreak = (elementType === DocumentApp.ElementType.PAGE_BREAK);
                    
                    // Увеличиваем счетчик страниц при разрыве страницы или в начале нового абзаца после разрыва
                    if (isPageBreak || (pageBreakFound && elementType === DocumentApp.ElementType.PARAGRAPH)) {
                        currentPage++;
                        pageBreakFound = isPageBreak;
                    } else if (i === 0) {
                        currentPage = 1; // Начинаем с первой страницы
                    }
                    
                    // Проверяем, входит ли текущая страница в заданный диапазон
                    let include = false;
                    for (const page of pages) {
                        const range = page.trim().split('-');
                        if (range.length === 1) { // Одиночная страница "1"
                            if (parseInt(range[0]) === currentPage) {
                                include = true;
                                break;
                            }
                        } else if (range.length === 2) { // Диапазон "2-4"
                            const start = parseInt(range[0]);
                            const end = parseInt(range[1]);
                            if (!isNaN(start) && !isNaN(end) && currentPage >= start && currentPage <= end) {
                                include = true;
                                break;
                            }
                        }
                    }
                    
                    // Если страница входит в диапазон и это не сам разрыв страницы, добавляем текст
                    if (include && !isPageBreak) {
                        if (elementType === DocumentApp.ElementType.PARAGRAPH) {
                            const paragraphText = element.asParagraph().getText();
                            // Записываем текст абзаца в отдельную ячейку
                            if (paragraphText.trim()) {
                                sheet.getRange(rowIndex++, 1).setValue(paragraphText);
                            }
                        } else if (elementType === DocumentApp.ElementType.TABLE) {
                            const table = element.asTable();
                            for (let r = 0; r < table.getNumRows(); r++) {
                                let rowText = [];
                                for (let c = 0; c < table.getRow(r).getNumCells(); c++) {
                                    rowText.push(table.getCell(r, c).getText());
                                }
                                // Записываем строку таблицы в отдельную ячейку
                                if (rowText.join('|').trim()) {
                                    sheet.getRange(rowIndex++, 1).setValue(rowText.join('|'));
                                }
                            }
                        } else if (elementType === DocumentApp.ElementType.LIST_ITEM) {
                            const listItemText = '* ' + element.asListItem().getText();
                            // Записываем элемент списка в отдельную ячейку
                            if (listItemText.trim()) {
                                sheet.getRange(rowIndex++, 1).setValue(listItemText);
                            }
                        }
                    }
                } catch (elementError) {
                    logMessage(`Ошибка при обработке элемента ${i}: ${elementError}`, true);
                    // Продолжаем обработку, пропускаем проблемный элемент
                }
            }
            
            logMessage(`convertGoogleDocToSpreadsheet: Текст извлечен, создано ${rowIndex-1} строк`);
        } else {
            // Если pageRange не указан, извлекаем весь текст блоками
            const body = doc.getBody();
            const text = body.getText();
            
            // Разбиваем весь текст на строки и вставляем по одной в таблицу
            const lines = text.split(/\r?\n/);
            let rowIndex = 1;
            
            for (let i = 0; i < lines.length; i++) {
                if (lines[i].trim()) {
                    sheet.getRange(rowIndex++, 1).setValue(lines[i]);
                }
            }
            
            logMessage(`convertGoogleDocToSpreadsheet: Весь текст извлечен, создано ${rowIndex-1} строк`);
        }
    } catch (e) {
        logMessage(`Ошибка при конвертации документа в таблицу: ${e}`, true);
        // Записываем сообщение об ошибке в первую ячейку
        sheet.getRange(1, 1).setValue("Ошибка при извлечении текста: " + e.toString());
    }

    return tempSpreadsheet.getId();
}


/**
 * Получает данные и заголовки из указанного листа временной таблицы.
 * Если sheetName не указан (для текстовых), берет первый лист.
 * @param {string} sheetId ID временной Google Таблицы.
 * @param {string} [sheetName] Имя листа (для табличных файлов).
 * @returns {{ sheetData: Array<Array<string>>, headers: string[] }}
 */
function getSheetDataAndHeaders(sheetId, sheetName) {
    if (!sheetId) {
      throw new Error("getSheetDataAndHeaders: Не передан ID временной таблицы (sheetId).");
    }
    let tempSpreadsheet;
    try {
        tempSpreadsheet = SpreadsheetApp.openById(sheetId);
    } catch (e) {
         throw new Error(`getSheetDataAndHeaders: Не удалось открыть временную таблицу по ID: ${sheetId}. Ошибка: ${e}`);
    }

    let sourceSheet;
    if (sheetName) {
        // Ищем лист по имени для табличных файлов
         sourceSheet = tempSpreadsheet.getSheetByName(sheetName);
         if (!sourceSheet) {
            // Если лист с таким именем не найден, пробуем взять первый
            logMessage(`getSheetDataAndHeaders: Лист с именем "${sheetName}" не найден, пробуем взять первый лист...`, false);
            sourceSheet = tempSpreadsheet.getSheets()[0];
             if (!sourceSheet) {
                throw new Error(`Лист с именем "${sheetName}" не найден и не удалось получить первый лист во временной таблице (ID: ${sheetId}).`);
             }
             logMessage(`getSheetDataAndHeaders: Используется первый лист "${sourceSheet.getName()}".`, false);
         }
    } else {
        // Если sheetName не указан (текстовый файл или ошибка получения имени), берем первый лист
        sourceSheet = tempSpreadsheet.getSheets()[0];
         if (!sourceSheet) {
            throw new Error(`getSheetDataAndHeaders: Не удалось получить первый лист из временной таблицы (ID: ${sheetId}).`);
         }
         logMessage(`getSheetDataAndHeaders: Имя листа не указано, используется первый лист "${sourceSheet.getName()}".`, false);
    }

    const sheetData = sourceSheet.getDataRange().getValues();
    logMessage(`getSheetDataAndHeaders: Получено ${sheetData.length} строк данных из листа "${sourceSheet.getName()}".`);


    if (!sheetData || sheetData.length === 0) {
        logMessage("getSheetDataAndHeaders: Лист пуст, возвращаем пустые данные и заголовки.", false);
        return { sheetData: [], headers: [] };
    }

    // Определяем заголовки ТОЛЬКО если это был табличный файл (sheetName был передан ИЛИ sheetData выглядит как таблица - более 1 столбца в первой строке)
    let headers = [];
    // Добавим проверку - если sheetName был или если первая строка содержит больше одной ячейки (похоже на заголовки таблицы)
    if (sheetName || (sheetData[0] && sheetData[0].length > 1)) {
        headers = sheetData[0].map(String);
        logMessage(`getSheetDataAndHeaders: Извлечены заголовки (для табличного файла): ${JSON.stringify(headers)}`);
    } else {
        logMessage(`getSheetDataAndHeaders: Заголовки не извлекались (предположительно текстовый файл).`);
    }

    return { sheetData, headers };
}

/**
 * Основная функция для анализа данных из временной таблицы и вставки в целевой лист.
 * @param {object} fileData Объект с данными файла {name, type, data, sheetName?, headerRow, selectedHeaders?, aiInstructions?, pageRange?}.
 * @returns {string | { error: string }} Сообщение об успехе или объект с ошибкой.
 */
function analyzeAndInsertExtractedData(fileData) {
    let tempSheetId;
    try {
        const processResult = processUploadedFile(fileData);
        if (processResult.error || !processResult.tempSheetId) {
            logMessage(`Ошибка на этапе processUploadedFile: ${processResult.error || 'tempSheetId не получен'}`, true);
            // Возвращаем объект ошибки, чтобы он был обработан в analyzeDataById
            return { error: `Ошибка при начальной обработке файла: ${processResult.error || 'Не удалось получить ID временной таблицы.'}` };
        }
        tempSheetId = processResult.tempSheetId;
        const isTextFile = processResult.fileType === 'text';

        logMessage(`Начало analyzeAndInsertExtractedData. sheetId: ${tempSheetId}, sheetName: ${fileData.sheetName}, headerRow: ${fileData.headerRow}, selectedHeaders: ${JSON.stringify(fileData.selectedHeaders)}, aiInstructions: ${fileData.aiInstructions}, pageRange: ${fileData.pageRange}, isTextFile: ${isTextFile}`);
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        const targetSheet = spreadsheet.getActiveSheet();

        if (!fileData.headerRow || isNaN(parseInt(fileData.headerRow)) || parseInt(fileData.headerRow) < 1) {
            throw new Error("Не указан или указан неверный номер строки с заголовками в целевом листе.");
        }

        const { sheetData, headers } = getSheetDataAndHeaders(tempSheetId, fileData.sheetName);
        const targetHeaders = getTargetHeaders(targetSheet, fileData.headerRow);

        if (!targetHeaders || targetHeaders.length === 0) {
            throw new Error("Не удалось получить заголовки целевого листа. Убедитесь, что указан правильный номер строки с заголовками и что строка не пустая.");
        }

        const selectedHeaders = fileData.selectedHeaders && fileData.selectedHeaders.length > 0
            ? fileData.selectedHeaders
            : targetHeaders;

        const prompt = isTextFile
          ? buildTextPrompt(sheetData, selectedHeaders, fileData.aiInstructions)
          : buildTablePrompt(sheetData, headers, selectedHeaders, fileData.aiInstructions);

        logMessage(`analyzeAndInsertExtractedData prompt: ${prompt}`);
        // Используем настройки для AI запроса
        const settings = getSettings();
        const aiResponse = openRouterRequest(
            prompt,
            settings.model,
            settings.temperature,
            settings.retryAttempts,
            settings.maxTokens
        );

        const parsedData = parseCsvAnswer(aiResponse, '|');

        insertDataIntoSheet(parsedData, targetSheet, selectedHeaders, targetHeaders);

        // Сообщение об успехе возвращается после попытки удаления в finally
        return "Данные успешно извлечены и вставлены.";

    } catch (error) {
        logMessage(`Ошибка в analyzeAndInsertExtractedData: ${error.toString()} ${error.stack}`, true);
        // Возвращаем объект ошибки
        return { error: `Ошибка: ${error.toString()}` };
    } finally {
        // Этот блок выполнится всегда: и после успешного try, и после catch
        if (tempSheetId) {
          try {
            DriveApp.getFileById(tempSheetId).setTrashed(true);
             logMessage(`Временная таблица ${tempSheetId} помещена в корзину (из finally).`);
          } catch(e) {
            logMessage(`Не удалось поместить в корзину временную таблицу ${tempSheetId} (из finally): ${e}`, true);
          }
        }
    }
}

/**
 * Формирует УПРОЩЕННЫЙ промпт для извлечения данных из табличных файлов.
 */
function buildTablePrompt(sheetData, sourceHeaders, targetHeaders, aiInstructions) {
    // Собираем ВСЕ строки данных (пропуская заголовок)
    let dataString = "";
    for (let i = 1; i < sheetData.length; i++) {
        if (sheetData[i].some(cell => String(cell).trim() !== '')) {
            dataString += sheetData[i].join('|') + ';';
        }
    }
    if (dataString.endsWith(';')) {
        dataString = dataString.slice(0, -1);
    }

    let prompt = `ЗАДАЧА: Извлечь данные из предоставленного CSV-фрагмента ("ДАННЫЕ ИЗ ФАЙЛА") в соответствии с "ЦЕЛЕВЫМИ ЗАГОЛОВКАМИ". Применить "ДОПОЛНИТЕЛЬНЫЕ ИНСТРУКЦИИ" при извлечении. Вернуть результат СТРОГО в формате CSV ('|' - столбцы, ';' - строки), БЕЗ ЗАГОЛОВКОВ.

ЦЕЛЕВЫЕ ЗАГОЛОВКИ (извлечь данные для них):
${targetHeaders.join(', ')}

ЗАГОЛОВКИ ИСХОДНОГО ФАЙЛА (для контекста): ${sourceHeaders.join(", ")}

ДАННЫЕ ИЗ ФАЙЛА (CSV, '|' - столбцы, ';' - строки):
${dataString}
`;

    if (aiInstructions) {
        prompt += `\nДОПОЛНИТЕЛЬНЫЕ ИНСТРУКЦИИ (применить при извлечении и форматировании):\n${aiInstructions}\n`;
    } else {
        prompt += `\nДОПОЛНИТЕЛЬНЫЕ ИНСТРУКЦИИ: Нет.\n`;
    }

    prompt += `\nПРАВИЛА ФОРМАТИРОВАНИЯ ОТВЕТА:
* Вернуть ТОЛЬКО CSV данные.
* Разделитель столбцов: '|'.
* Разделитель строк: ';'.
* Если данных для целевого столбца нет, оставить поле ПУСТЫМ.
* Если в извлекаемых данных встречается символ ';', ЗАМЕНИТЬ его на '.' (точка).
* НЕ включать заголовки столбцов в ответ.
* НЕ добавлять никаких пояснений или другого текста.`;

    return prompt;
}

/**
 * Формирует УПРОЩЕННЫЙ промпт для извлечения данных из текстовых файлов.
 */
function buildTextPrompt(sheetData, targetHeaders, aiInstructions) {
    let fullText = "";
    if (sheetData && sheetData.length > 0) {
        fullText = sheetData.map(row => row.join(" ")).join("\n");
    }

    let prompt = `ЗАДАЧА: Извлечь данные из предоставленного ТЕКСТА ("ТЕКСТ ФАЙЛА") в соответствии с "ЦЕЛЕВЫМИ ЗАГОЛОВКАМИ". Применить "ДОПОЛНИТЕЛЬНЫЕ ИНСТРУКЦИИ" при извлечении. Вернуть результат СТРОГО в формате CSV ('|' - столбцы, ';' - строки), БЕЗ ЗАГОЛОВКОВ.

ЦЕЛЕВЫЕ ЗАГОЛОВКИ (извлечь данные для них):
${targetHeaders.join(', ')}

ТЕКСТ ФАЙЛА:
${fullText}
`;

    if (aiInstructions) {
        prompt += `\nДОПОЛНИТЕЛЬНЫЕ ИНСТРУКЦИИ (применить при извлечении и форматировании):\n${aiInstructions}\n`;
    } else {
        prompt += `\nДОПОЛНИТЕЛЬНЫЕ ИНСТРУКЦИИ: Нет.\n`;
    }

    // Обрабатывать весь текст целиком, страницы уже учтены на уровне скрипта
    prompt += `\nОбрабатывать весь текст целиком.\n`;

    prompt += `\nПРАВИЛА ФОРМАТИРОВАНИЯ ОТВЕТА:
* Вернуть ТОЛЬКО CSV данные.
* Разделитель столбцов: '|'.
* Разделитель строк: ';'.
* Если данных для целевого столбца нет, оставить поле ПУСТЫМ.
* Если в извлекаемых данных встречается символ ';', ЗАМЕНИТЬ его на '.' (точка).
* НЕ включать заголовки столбцов в ответ.
* НЕ добавлять никаких пояснений или другого текста.`;

    return prompt;
}

/**
 * Обрабатывает ответ AI и возвращает массив строковых массивов.
 */
function parseCsvAnswer(aiResponse, columnDelimiter) {
    const rowDelimiter = ';';
    if (!aiResponse?.choices?.[0]?.message) {
        throw new Error("Неожиданный ответ от OpenRouter: " + JSON.stringify(aiResponse));
    }
    let answer = aiResponse.choices[0].message.content.trim();
    answer = answer.replace(/^```csv\s*/i, '').replace(/```\s*$/i, '').trim();
    answer = answer.replace(/[\r\n]+/g, rowDelimiter);
    answer = answer.replace(/^;+|;+$|;;+/g, ';');
    logMessage(`parseCsvAnswer, answer after cleanup: ${answer}`);
    const result = [];
    if (answer) {
        try {
            const rows = answer.split(rowDelimiter);
            for (const rowString of rows) {
                if (rowString.trim() === '') continue;
                const parsedRow = Utilities.parseCsv(rowString, columnDelimiter.charCodeAt(0))[0];
                if (parsedRow) result.push(parsedRow.map(cell => cell.trim()));
            }
        } catch (error) {
            logMessage(`Ошибка парсинга CSV: ${error.message}. Ответ AI: "${answer}"`, true);
            return [];
        }
    }
    logMessage(`parseCsvAnswer, final result: ${JSON.stringify(result)}`);
    return result;
}

/**
 * Вставляет данные в целевой лист после analyzeAndInsertExtractedData.
 */
function insertDataIntoSheet(data, targetSheet, selectedHeaders, targetHeaders) {
    if (!data || data.length === 0 || !targetSheet || !selectedHeaders?.length || !targetHeaders?.length) return;
    const startRow = targetSheet.getLastRow() + 1;
    const numRows = data.length;
    const targetHeaderIndexMap = {};
    targetHeaders.forEach((h,i) => { if (h) targetHeaderIndexMap[h.toLowerCase()] = i; });
    let maxCol = -1;
    selectedHeaders.forEach(sh => { const idx = targetHeaderIndexMap[sh.toLowerCase()]; if (idx >= 0) maxCol = Math.max(maxCol, idx); });
    if (maxCol < 0) return;
    const numCols = maxCol + 1;
    const outputData = Array.from({ length: numRows }, () => Array(numCols).fill(""));
    for (let i=0; i<numRows; i++) {
        const row = Array.isArray(data[i]) ? data[i] : [];
        for (let j=0; j<selectedHeaders.length; j++) {
            const idx = targetHeaderIndexMap[selectedHeaders[j].toLowerCase()];
            if (idx >= 0 && j < row.length) outputData[i][idx] = row[j];
        }
    }
    try {
        targetSheet.getRange(startRow, 1, numRows, numCols).setValues(outputData);
    } catch(e) {
        logMessage(`Ошибка при записи данных: ${e}`, true);
    }
}

/**
 * Возвращает массив заголовков из целевого листа по номеру строки.
 */
function getTargetHeaders(targetSheet, headerRow) {
    if (!targetSheet || headerRow < 1) throw new Error("Неверный лист или строка заголовков.");
    const lastCol = targetSheet.getLastColumn();
    if (lastCol === 0) return [];
    const headers = targetSheet.getRange(headerRow,1,1,lastCol).getValues()[0];
    if (!headers.some(h => String(h).trim())) throw new Error("Строка заголовков пуста.");
    return headers.map(String);
}

/**
 * Обрабатывает данные из клиента и запускает processUploadedFile и analyzeAndInsertExtractedData 
 * в два шага, чтобы избежать ошибки PERMISSION_DENIED при передаче больших данных.
 * @param {object} fileData Объект с данными файла.
 * @returns {string} Сообщение об успехе или ошибке.
 */
function processFileAndAnalyze(fileData) {
  try {
    // Сохраняем fileData во временное хранилище, чтобы не передавать большие данные дважды
    const fileDataId = Utilities.getUuid();
    CacheService.getScriptCache().put(
      "fileDataCache_" + fileDataId, 
      JSON.stringify(fileData),
      3600 // 1 час хранения
    );
    
    return {
      success: true,
      fileDataId: fileDataId,
      message: "Файл успешно обработан. Теперь можно извлекать данные."
    };
  } catch (error) {
    return {
      success: false,
      error: `Ошибка при обработке файла: ${error.toString()}`
    };
  }
}

/**
 * Второй шаг - анализ данных по ID из кэша
 */
function analyzeDataById(fileDataId, headerRow, selectedHeaders, aiInstructions, pageRange) {
  try {
    // Получаем данные из кэша
    const cachedData = CacheService.getScriptCache().get("fileDataCache_" + fileDataId);
    if (!cachedData) {
      return {
        success: false,
        error: "Данные файла не найдены в кэше. Пожалуйста, загрузите файл заново."
      };
    }
    
    const fileData = JSON.parse(cachedData);
    
    // Добавляем параметры анализа
    fileData.headerRow = headerRow;
    fileData.selectedHeaders = selectedHeaders;
    fileData.aiInstructions = aiInstructions;
    fileData.pageRange = pageRange;
    
    // Вызываем функцию анализа
    const result = analyzeAndInsertExtractedData(fileData);
    
    // Очищаем кэш
    CacheService.getScriptCache().remove("fileDataCache_" + fileDataId);
    
    if (typeof result === 'string') {
      return {
        success: true,
        message: result
      };
    } else if (result.error) {
      return {
        success: false,
        error: result.error
      };
    } else {
      return {
        success: true,
        message: "Данные успешно извлечены и вставлены."
      };
    }
  } catch (error) {
    return {
      success: false,
      error: `Ошибка при анализе данных: ${error.toString()}`
    };
  }
}

/**
 * Проверяет, является ли строка валидной Base64
 * @param {string} str Строка для проверки
 * @returns {boolean} Результат проверки
 */
function isValidBase64(str) {
  if (!str || typeof str !== 'string') return false;
  
  // Регулярное выражение для проверки Base64 строки (может содержать только A-Z, a-z, 0-9, +, /, =)
  const base64Regex = /^[A-Za-z0-9+/=]+$/;
  
  // Проверяем базовый формат
  if (!base64Regex.test(str)) {
    logMessage("Base64 содержит недопустимые символы", true);
    return false;
  }
  
  // Проверяем правильную длину (кратность 4) и корректное размещение символов = (только в конце)
  const paddingCount = (str.match(/=/g) || []).length;
  if (paddingCount > 2) {
    logMessage("Неверное количество символов = в Base64", true);
    return false;
  }
  
  if (paddingCount > 0 && str.indexOf('=') !== str.length - paddingCount) {
    logMessage("Символы = находятся не в конце строки Base64", true);
    return false;
  }

  // Базовая проверка пройдена
  return true;
}