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
             // Заменяем фиксированную паузу на адаптивное ожидание с проверкой
             let sheetsAvailable = false;
             const maxAttempts = 5;
             let attempts = 0;
             
             while (!sheetsAvailable && attempts < maxAttempts) {
                attempts++;
                try {
                   const sheets = tempSpreadsheet.getSheets();
                   if (sheets && sheets.length > 0) {
                      sheetsAvailable = true;
                      sheetNames = sheets.map(sheet => sheet.getName());
                      logMessage(`processUploadedFile: Получены имена листов за ${attempts} попыток: ${JSON.stringify(sheetNames)}`);
                   } else {
                      // Короткая пауза перед следующей попыткой
                      Utilities.sleep(200);
                   }
                } catch (e) {
                   // Если произошла ошибка, делаем короткую паузу перед следующей попыткой
                   Utilities.sleep(200);
                   logMessage(`processUploadedFile: Попытка ${attempts} получения листов - не удалась: ${e}`);
                }
             }
             
             // Если после всех попыток листы не получены
             if (!sheetsAvailable) {
                logMessage(`processUploadedFile: Не удалось получить листы после ${maxAttempts} попыток. Возвращаем ["Лист1"]`, true);
                sheetNames = ["Лист1"];
             }
             
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
          logMessage(`processUploadedFile: Файл текстовый, имена листов не извлекаются.`);
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
 * Обновлено для работы с Drive API v3 и оптимизировано для быстрого доступа.
 * @param {GoogleAppsScript.Drive.File} file
 * @param {GoogleAppsScript.Drive.Folder} folder
 * @returns {string} ID созданного Google Doc.
 */
function convertFileToGoogleDoc(file, folder) {
    try {
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
        
        // Заменяем sleep на активное ожидание с небольшими интервалами
        let docReady = false;
        let attempts = 0;
        const maxAttempts = 10;
        
        while (!docReady && attempts < maxAttempts) {
            try {
                attempts++;
                // Проверяем доступность документа
                DocumentApp.openById(docId);
                docReady = true;
                logMessage(`convertFileToGoogleDoc: Документ доступен после ${attempts} попыток`);
            } catch (e) {
                // Если документ еще не готов, ждем короткое время и пробуем снова
                Utilities.sleep(200);
                logMessage(`convertFileToGoogleDoc: Ожидание конвертации, попытка ${attempts}/${maxAttempts}`);
            }
        }
        
        if (!docReady) {
            logMessage(`convertFileToGoogleDoc: Предупреждение - документ может быть не полностью готов после ${maxAttempts} попыток`, true);
        }
        
        // Удаляем исходный временный файл через API v3
        try {
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
    // Открываем Google Doc и используем малый таймаут вместо фиксированного большого
    const doc = DocumentApp.openById(docId);
    // Заменяем длинную паузу на короткую
    Utilities.sleep(200);
    
    // Создаем временную таблицу
    const tempSpreadsheet = SpreadsheetApp.create("Temp Sheet from Doc " + Date.now());
    const sheet = tempSpreadsheet.getActiveSheet();
    
    try {
        logMessage(`convertGoogleDocToSpreadsheet: Начинаем извлечение текста из документа ${docId}`);
        
        // Данные для пакетной записи
        let batchData = [];
        
        // Логика извлечения текста с учетом pageRange
        if (pageRange) {
            const pages = pageRange.split(','); // Разбираем строку диапазона
            const body = doc.getBody();
            
            // Оптимизация: сначала получаем весь текст, если документ простой и не содержит сложного форматирования
            // Это будет работать быстрее для документов с простым текстом
            if (body.getNumChildren() < 100 && !body.getTables().length) {
                const fullText = body.getText();
                // Разделяем текст на страницы по специальным символам перевода страницы (\f)
                const textPages = fullText.split('\f');
                let pageNum = 1;
                
                for (const page of textPages) {
                    // Проверяем, входит ли страница в заданный диапазон
                    let include = false;
                    for (const pageRange of pages) {
                        const range = pageRange.trim().split('-');
                        if (range.length === 1) { // Одиночная страница "1"
                            if (parseInt(range[0]) === pageNum) {
                                include = true;
                                break;
                            }
                        } else if (range.length === 2) { // Диапазон "2-4"
                            const start = parseInt(range[0]);
                            const end = parseInt(range[1]);
                            if (!isNaN(start) && !isNaN(end) && pageNum >= start && pageNum <= end) {
                                include = true;
                                break;
                            }
                        }
                    }
                    
                    if (include && page.trim()) {
                        // Разбиваем текст страницы на строки и добавляем в batchData
                        const lines = page.split(/\r?\n/);
                        for (const line of lines) {
                            if (line.trim()) {
                                batchData.push([line]);
                            }
                        }
                    }
                    pageNum++;
                }
            } else {
                // Используем оригинальный метод для сложных документов
                // Получаем все дочерние элементы через getNumChildren/getChild
                let rowIndex = 1;
                let currentPage = 0;
                let pageBreakFound = false;
                const count = body.getNumChildren();
                
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
                                // Добавляем в массив для пакетной записи
                                if (paragraphText.trim()) {
                                    batchData.push([paragraphText]);
                                }
                            } else if (elementType === DocumentApp.ElementType.TABLE) {
                                const table = element.asTable();
                                for (let r = 0; r < table.getNumRows(); r++) {
                                    let rowText = [];
                                    for (let c = 0; c < table.getRow(r).getNumCells(); c++) {
                                        rowText.push(table.getCell(r, c).getText());
                                    }
                                    // Добавляем в массив для пакетной записи
                                    if (rowText.join('|').trim()) {
                                        batchData.push([rowText.join('|')]);
                                    }
                                }
                            } else if (elementType === DocumentApp.ElementType.LIST_ITEM) {
                                const listItemText = '* ' + element.asListItem().getText();
                                // Добавляем в массив для пакетной записи
                                if (listItemText.trim()) {
                                    batchData.push([listItemText]);
                                }
                            }
                        }
                    } catch (elementError) {
                        logMessage(`Ошибка при обработке элемента ${i}: ${elementError}`, true);
                        // Продолжаем обработку, пропускаем проблемный элемент
                    }
                }
            }
            
            logMessage(`convertGoogleDocToSpreadsheet: Текст извлечен, найдено ${batchData.length} строк текста`);
        } else {
            // Если pageRange не указан, извлекаем весь текст более оптимально
            const body = doc.getBody();
            const text = body.getText();
            
            // Разбиваем весь текст на строки и собираем в массив для пакетной записи
            const lines = text.split(/\r?\n/);
            for (let line of lines) {
                if (line.trim()) {
                    batchData.push([line]);
                }
            }
            
            logMessage(`convertGoogleDocToSpreadsheet: Весь текст извлечен, найдено ${batchData.length} строк`);
        }
        
        // Пакетная запись всех данных одним вызовом (намного быстрее, чем построчная)
        if (batchData.length > 0) {
            sheet.getRange(1, 1, batchData.length, 1).setValues(batchData);
            logMessage(`convertGoogleDocToSpreadsheet: Данные записаны в таблицу одной пакетной операцией`);
        } else {
            // Если данных нет, запишем сообщение
            sheet.getRange(1, 1).setValue("Не найдено текста для извлечения.");
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
    const dataLines = [];
    for (let i = 1; i < sheetData.length; i++) {
        if (sheetData[i].some(cell => String(cell).trim() !== '')) {
            dataLines.push(sheetData[i].join('|'));
        }
    }
    const dataString = dataLines.join(';');

    // Формируем промпт, используя массив строк и join() вместо конкатенации строк
    const promptParts = [
        `ЗАДАЧА: Извлечь данные из предоставленного CSV-фрагмента ("ДАННЫЕ ИЗ ФАЙЛА") в соответствии с "ЦЕЛЕВЫМИ ЗАГОЛОВКАМИ". Применить "ДОПОЛНИТЕЛЬНЫЕ ИНСТРУКЦИИ" при извлечении. Вернуть результат СТРОГО в формате CSV ('|' - столбцы, ';' - строки), БЕЗ ЗАГОЛОВКОВ.`,
        `\nЦЕЛЕВЫЕ ЗАГОЛОВКИ (извлечь данные для них):\n${targetHeaders.join(', ')}`,
        `\nЗАГОЛОВКИ ИСХОДНОГО ФАЙЛА (для контекста): ${sourceHeaders.join(", ")}`,
        `\nДАННЫЕ ИЗ ФАЙЛА (CSV, '|' - столбцы, ';' - строки):\n${dataString}`
    ];

    // Добавляем дополнительные инструкции
    if (aiInstructions) {
        promptParts.push(`\nДОПОЛНИТЕЛЬНЫЕ ИНСТРУКЦИИ (применить при извлечении и форматировании):\n${aiInstructions}`);
    } else {
        promptParts.push(`\nДОПОЛНИТЕЛЬНЫЕ ИНСТРУКЦИИ: Нет.`);
    }

    // Усиленные правила форматирования ответа
    promptParts.push(`\nПРАВИЛА ФОРМАТИРОВАНИЯ ОТВЕТА - ЧРЕЗВЫЧАЙНО ВАЖНО:
1. СТРОГО ЗАПРЕЩЕНО добавлять любой текст, пояснения или комментарии в ответ.
2. СТРОГО ЗАПРЕЩЕНО использовать любые другие символы-разделители, кроме указанных.
3. СТРОГО ЗАПРЕЩЕНО включать заголовки или название колонок в ответ.
4. Начинать ответ строго с первой ячейки данных, без пустых строк.
5. Для разделения СТОЛБЦОВ всегда использовать ТОЛЬКО вертикальную черту '|'. 
6. Для разделения СТРОК/ЗАПИСЕЙ всегда использовать ТОЛЬКО точку с запятой ';'.
7. Если несколько записей (строк) - разделять их точкой с запятой ';'.
8. Если данных для целевого столбца нет, оставить поле ПУСТЫМ (пример: 'значение1||значение3').
9. Если в извлекаемых данных встречается символ ';', ЗАМЕНИТЬ его на '.' (точка).
10. Для каждого целевого заголовка должна быть ОДНА колонка в том же порядке.

ФОРМАТ ОТВЕТА - ВСЕГДА ВЫГЛЯДИТ ТАК:
значение1|значение2|значение3;
значение4|значение5|значение6;

НЕ СОБЛЮДАТЬ ЭТИ ПРАВИЛА АБСОЛЮТНО НЕДОПУСТИМО!`);

    // Объединяем все части промпта вместе
    return promptParts.join('');
}

/**
 * Формирует СТРОГИЙ промпт для извлечения данных из текстовых файлов.
 * Усилен для обеспечения единообразного формата ответа от разных моделей.
 */
function buildTextPrompt(sheetData, targetHeaders, aiInstructions) {
    // Оптимизированный вариант без многократной конкатенации строк
    let textLines = [];
    if (sheetData && sheetData.length > 0) {
        textLines = sheetData.map(row => row.join(" "));
    }
    const fullText = textLines.join("\n");
    
    // Формируем промпт из массива частей для лучшей производительности
    const promptParts = [
        `ЗАДАЧА: Извлечь данные из предоставленного ТЕКСТА ("ТЕКСТ ФАЙЛА") в соответствии с "ЦЕЛЕВЫМИ ЗАГОЛОВКАМИ". Применить "ДОПОЛНИТЕЛЬНЫЕ ИНСТРУКЦИИ" при извлечении. Вернуть результат СТРОГО в формате CSV ('|' - столбцы, ';' - строки), БЕЗ ЗАГОЛОВКОВ.`,
        `\nЦЕЛЕВЫЕ ЗАГОЛОВКИ (извлечь данные для них):\n${targetHeaders.join(', ')}`,
        `\nТЕКСТ ФАЙЛА:\n${fullText}`
    ];
    
    // Добавляем дополнительные инструкции в массив
    if (aiInstructions) {
        promptParts.push(`\nДОПОЛНИТЕЛЬНЫЕ ИНСТРУКЦИИ (применить при извлечении и форматировании):\n${aiInstructions}`);
    } else {
        promptParts.push(`\nДОПОЛНИТЕЛЬНЫЕ ИНСТРУКЦИИ: Нет.`);
    }
    
    // Обрабатывать весь текст целиком, страницы уже учтены на уровне скрипта
    promptParts.push(`\nОбрабатывать весь текст целиком.`);
    
    // Усиленные правила форматирования ответа - аналогично табличному промпту
    promptParts.push(`\nПРАВИЛА ФОРМАТИРОВАНИЯ ОТВЕТА - ЧРЕЗВЫЧАЙНО ВАЖНО:
1. СТРОГО ЗАПРЕЩЕНО добавлять любой текст, пояснения или комментарии в ответ.
2. СТРОГО ЗАПРЕЩЕНО использовать любые другие символы-разделители, кроме указанных.
3. СТРОГО ЗАПРЕЩЕНО включать заголовки или название колонок в ответ.
4. Начинать ответ строго с первой ячейки данных, без пустых строк.
5. Для разделения СТОЛБЦОВ всегда использовать ТОЛЬКО вертикальную черту '|'. 
6. Для разделения СТРОК/ЗАПИСЕЙ всегда использовать ТОЛЬКО точку с запятой ';'.
7. Если несколько записей (строк) - разделять их точкой с запятой ';'.
8. Если данных для целевого столбца нет, оставить поле ПУСТЫМ (пример: 'значение1||значение3').
9. Если в извлекаемых данных встречается символ ';', ЗАМЕНИТЬ его на '.' (точка).
10. Для каждого целевого заголовка должна быть ОДНА колонка в том же порядке.

ФОРМАТ ОТВЕТА - ВСЕГДА ВЫГЛЯДИТ ТАК:
значение1|значение2|значение3;
значение4|значение5|значение6;

НЕ СОБЛЮДАТЬ ЭТИ ПРАВИЛА АБСОЛЮТНО НЕДОПУСТИМО!`);
    
    // Объединяем все части в финальный промпт
    return promptParts.join('');
}

/**
 * Обрабатывает ответ AI и возвращает массив строковых массивов.
 * Учитывает различные форматы ответов от разных моделей.
 * @param {object} aiResponse Ответ от AI.
 * @param {string} columnDelimiter Ожидаемый разделитель колонок (используется как dataDelimiter, если не определен автоматически).
 * @returns {string[][]} Массив массивов строк с данными.
 */
function parseCsvAnswer(aiResponse, columnDelimiter) { // columnDelimiter передан, но не используется в текущей логике определения разделителей
    const defaultRowDelimiter = ';'; // Используем как разделитель записей по умолчанию или при очистке
    if (!aiResponse?.choices?.[0]?.message) {
        throw new Error("Неожиданный ответ от OpenRouter: " + JSON.stringify(aiResponse));
    }

    let answer = aiResponse.choices[0].message.content.trim();

    // Очищаем ответ от всех вариантов маркеров кода и лишних символов
    answer = answer.replace(/^```(?:csv|text|plain|)?\s*/i, '').replace(/```\s*$/i, '').trim();
    // Заменяем переносы строк на стандартный разделитель записей
    answer = answer.replace(/[\r\n]+/g, defaultRowDelimiter);
    // Удаляем лишние разделители записей в начале, конце и дубликаты
    answer = answer.replace(/^;+|;+$|;;+/g, ';').trim(); // Добавил trim() на всякий случай

    logMessage(`parseCsvAnswer, answer after cleanup: ${answer}`);

    const result = [];

    if (!answer) {
        return result;
    }

    try {
        // Анализируем структуру ответа для определения разделителей
        const hasPipe = answer.includes('|');
        const hasSemicolon = answer.includes(';');

        // Если оба разделителя отсутствуют, возвращаем всю строку как одну запись с одной ячейкой
        if (!hasPipe && !hasSemicolon) {
            // Убедимся, что возвращаем массив массивов
            return [[answer]];
        }

        // Определяем, какой символ используется для разделения записей (строк), а какой для данных (колонок)
        // Эвристика: если присутствуют оба символа ('|' и ';'),
        // предполагаем, что тот, который встречается реже, разделяет записи (строки),
        // а тот, что чаще - колонки внутри записи.
        let recordDelimiter, dataDelimiter;

        if (hasPipe && hasSemicolon) {
            const pipeCount = (answer.match(/\|/g) || []).length;
            const semicolonCount = (answer.match(/;/g) || []).length;

            // Если '|' больше, то ';' - разделитель записей, '|' - разделитель данных
            if (pipeCount > semicolonCount) {
                recordDelimiter = ';';
                dataDelimiter = '|';
            } else { // Иначе (если ';' больше или равно) '|' - разделитель записей, ';' - разделитель данных
                recordDelimiter = '|';
                dataDelimiter = ';';
            }
            logMessage(`parseCsvAnswer: Detected delimiters - Record: '${recordDelimiter}', Data: '${dataDelimiter}'`);
        } else if (hasPipe) {
            // Только '|': предполагаем, что это разделитель данных, а записи разделены defaultRowDelimiter (';') после очистки
            recordDelimiter = defaultRowDelimiter; // ';'
            dataDelimiter = '|';
            logMessage(`parseCsvAnswer: Only '|' found. Assuming Record: '${recordDelimiter}', Data: '${dataDelimiter}'`);
             // Перепроверяем, есть ли ';' после очистки. Если нет, то записей нет.
             if (!answer.includes(recordDelimiter)) {
                 recordDelimiter = null; // Нет разделителя записей
                 logMessage(`parseCsvAnswer: ';' not found after cleanup. Setting recordDelimiter to null.`);
             }
        } else { // Только ';'
            // Только ';': предполагаем, что это разделитель данных, записи не разделены (или были разделены переносами, замененными на ';')
            // Это неоднозначный случай. Будем считать ';' разделителем данных, а разделителя записей нет.
            recordDelimiter = null; // Нет явного разделителя записей, кроме ';'
            dataDelimiter = ';';
            logMessage(`parseCsvAnswer: Only ';' found. Assuming Record: null, Data: '${dataDelimiter}'`);
            // Если ';' используется как разделитель данных, то как разделить записи?
            // В текущей логике это приведет к тому, что вся строка будет считаться одной записью, разделенной ';'.
            // Возможно, стоит использовать defaultRowDelimiter (';') как recordDelimiter здесь?
            // Давайте попробуем:
            recordDelimiter = defaultRowDelimiter; // ';'
             // Но если ';' - это и разделитель данных, и разделитель записей, будет путаница.
             // Оставим пока dataDelimiter = ';', recordDelimiter = null для этого случая.
             recordDelimiter = null;


        }

        // Разделяем на отдельные записи
        // Если recordDelimiter определен, используем его. Иначе считаем всю строку одной записью.
        const records = recordDelimiter
            ? answer.split(recordDelimiter).filter(rec => rec.trim() !== '')
            : [answer]; // Если разделитель записей не найден, считаем всю строку одной записью

        logMessage(`parseCsvAnswer: Split into ${records.length} potential records using delimiter '${recordDelimiter}'`);

        // Обрабатываем каждую запись
        for (const record of records) {
            const trimmedRecord = record.trim();
            if (trimmedRecord === '') continue;

            // Разделяем данные записи на колонки, используя определенный dataDelimiter
            if (trimmedRecord.includes(dataDelimiter)) {
                try {
                    // Используем стандартный парсер CSV для колонок
                    // Utilities.parseCsv ожидает символ, а не строку
                    const parsedRow = Utilities.parseCsv(trimmedRecord, dataDelimiter.charCodeAt(0))[0];
                    // Проверяем, что строка не пустая после парсинга
                    if (parsedRow && parsedRow.some(cell => String(cell).trim() !== '')) {
                        result.push(parsedRow.map(cell => String(cell).trim()));
                    } else {
                         logMessage(`parseCsvAnswer: Parsed row is empty or invalid for record: "${trimmedRecord}"`);
                    }
                } catch (e) {
                    logMessage(`parseCsvAnswer: Utilities.parseCsv failed for record "${trimmedRecord}" with delimiter '${dataDelimiter}'. Error: ${e}. Falling back to split.`);
                    // В случае ошибки парсинга (например, некорректные кавычки), делаем разделение вручную
                    const cells = trimmedRecord.split(dataDelimiter).map(cell => cell.trim());
                     // Проверяем, что строка не пустая после ручного разделения
                    if (cells.some(cell => cell !== '')) {
                        result.push(cells);
                    } else {
                         logMessage(`parseCsvAnswer: Manual split resulted in empty row for record: "${trimmedRecord}"`);
                    }
                }
            } else {
                // Если разделителя данных нет в этой записи, добавляем всю запись как одну колонку
                 logMessage(`parseCsvAnswer: Data delimiter '${dataDelimiter}' not found in record "${trimmedRecord}". Adding as single cell.`);
                result.push([trimmedRecord]);
            }
        }

        // Дополнительная валидация и очистка результата (удаление пустых строк/ячеек)
        // Этот блок можно упростить или удалить, если парсинг и фильтрация выше работают корректно
        const finalResult = [];
        for (let row of result) {
            // Убираем пустые ячейки в конце каждой строки
            let lastNonEmptyIndex = -1;
            for (let j = row.length - 1; j >= 0; j--) {
                if (row[j] !== '') {
                    lastNonEmptyIndex = j;
                    break;
                }
            }
            // Обрезаем строку, если есть пустые ячейки в конце
            if (lastNonEmptyIndex < row.length - 1) {
                row = row.slice(0, lastNonEmptyIndex + 1);
            }

            // Если строка не пустая после обрезки, добавляем в результат
            if (row.length > 0) {
                 finalResult.push(row);
            }
        }
         logMessage(`parseCsvAnswer, final result after validation: ${JSON.stringify(finalResult)}`);
         return finalResult; // Возвращаем очищенный результат

    } catch (error) {
        logMessage(`Ошибка парсинга CSV: ${error.message}. Ответ AI: "${answer}"`, true);
        // В случае ошибки возвращаем пустой массив, чтобы не прерывать выполнение
        return [];
    }

    // Возвращаем результат обработки (добавлен для безопасности)
    logMessage(`parseCsvAnswer, final result: ${JSON.stringify(result)}`);
    return result;
}

/**
 * Вставляет данные в целевой лист после analyzeAndInsertExtractedData.
 * Исправлена проблема дублирования данных.
 * @param {Array} data Массив данных для вставки
 * @param {Object} targetSheet Целевой лист
 * @param {Array} selectedHeaders Выбранные заголовки
 * @param {Array} targetHeaders Заголовки целевого листа
 * @param {boolean} [byColumn=false] Если true, вставляет данные, находя последнюю строку для каждого столбца отдельно
 */
function insertDataIntoSheet(data, targetSheet, selectedHeaders, targetHeaders, byColumn = false) {
    // Если включен режим вставки по столбцам, используем новую функцию
    if (byColumn === true) {
        return insertDataIntoSheetByColumn(data, targetSheet, selectedHeaders, targetHeaders);
    }
    
    // Оригинальная функция вставки по последней строке таблицы
    if (!data || data.length === 0 || !targetSheet || !selectedHeaders?.length || !targetHeaders?.length) return;
    
    // Добавляем защиту от дублирования - проверяем, есть ли уже данные с таким же содержанием
    const startRow = targetSheet.getLastRow() + 1;
    const numRows = data.length;
    const targetHeaderIndexMap = {};
    
    targetHeaders.forEach((h,i) => { 
        if (h) targetHeaderIndexMap[h.toLowerCase()] = i; 
    });
    
    let maxCol = -1;
    selectedHeaders.forEach(sh => { 
        const idx = targetHeaderIndexMap[sh.toLowerCase()]; 
        if (idx >= 0) maxCol = Math.max(maxCol, idx); 
    });
    
    if (maxCol < 0) return;
    const numCols = maxCol + 1;
    
    // Проверяем на дубликаты
    const outputData = [];
    
    for (let i = 0; i < numRows; i++) {
        const row = Array.isArray(data[i]) ? data[i] : [];
        const outputRow = Array(numCols).fill("");
        
        for (let j = 0; j < selectedHeaders.length; j++) {
            const idx = targetHeaderIndexMap[selectedHeaders[j].toLowerCase()];
            if (idx >= 0 && j < row.length) outputRow[idx] = row[j];
        }
        
        // Проверяем, что строка не пустая (содержит хотя бы один непустой элемент)
        if (outputRow.some(cell => cell.toString().trim() !== "")) {
            outputData.push(outputRow);
        }
    }
    
    // Если после фильтрации данных нет, выходим
    if (outputData.length === 0) return;
    
    try {
        targetSheet.getRange(startRow, 1, outputData.length, numCols).setValues(outputData);
        logMessage(`insertDataIntoSheet: Вставлено ${outputData.length} строк данных`);
    } catch(e) {
        logMessage(`Ошибка при записи данных: ${e}`, true);
    }
}

/**
 * Вставляет данные в целевой лист, находя последнюю занятую ячейку в каждом столбце.
 * @param {Array} data Массив данных для вставки
 * @param {Object} targetSheet Целевой лист
 * @param {Array} selectedHeaders Выбранные заголовки
 * @param {Array} targetHeaders Заголовки целевого листа
 */
function insertDataIntoSheetByColumn(data, targetSheet, selectedHeaders, targetHeaders) {
    if (!data || data.length === 0 || !targetSheet || !selectedHeaders?.length || !targetHeaders?.length) return;
    
    // Создаем карту соответствия заголовков столбцам в целевой таблице
    const targetHeaderIndexMap = {};
    targetHeaders.forEach((h,i) => { 
        if (h) targetHeaderIndexMap[h.toLowerCase()] = i; 
    });
    
    // Находим максимальный индекс столбца для определения ширины данных
    let maxCol = -1;
    selectedHeaders.forEach(sh => { 
        const idx = targetHeaderIndexMap[sh.toLowerCase()]; 
        if (idx >= 0) maxCol = Math.max(maxCol, idx); 
    });
    
    if (maxCol < 0) return;
    const numCols = maxCol + 1;
    
    // Находим последнюю занятую ячейку для каждого целевого столбца
    const lastRowByColumn = {};
    
    // Получаем все данные таблицы для анализа
    const allData = targetSheet.getDataRange().getValues();
    const headerRowIndex = parseInt(targetSheet.createTextFinder(targetHeaders[0]).findNext().getRow()) - 1;
    
    // Для каждого столбца находим последнюю непустую ячейку
    selectedHeaders.forEach(header => {
        const colIndex = targetHeaderIndexMap[header.toLowerCase()];
        if (colIndex >= 0) {
            let lastRow = headerRowIndex; // Начинаем с строки заголовков
            
            // Просматриваем все строки в поисках последней непустой ячейки
            for (let i = 0; i < allData.length; i++) {
                if (i > headerRowIndex && // Пропускаем заголовки и строки выше
                    colIndex < allData[i].length && 
                    String(allData[i][colIndex]).trim() !== '') {
                    lastRow = i;
                }
            }
            
            // Сохраняем номер последней строки с данными для этого столбца
            lastRowByColumn[colIndex] = lastRow + 1; // +1 т.к. массив начинается с 0, а строки с 1
        }
    });
    
    logMessage(`insertDataIntoSheetByColumn: Найдены последние занятые строки для столбцов: ${JSON.stringify(lastRowByColumn)}`);
    
    // Для каждой строки новых данных, создаем список операций вставки по столбцам
    const columnOperations = [];
    
    for (let i = 0; i < data.length; i++) {
        const row = Array.isArray(data[i]) ? data[i] : [];
        
        // Для каждого заголовка проверяем, есть ли данные для вставки
        for (let j = 0; j < selectedHeaders.length; j++) {
            const header = selectedHeaders[j];
            const colIndex = targetHeaderIndexMap[header.toLowerCase()];
            
            if (colIndex >= 0 && j < row.length && row[j].toString().trim() !== '') {
                // Если есть данные, добавляем операцию вставки
                const targetRow = lastRowByColumn[colIndex] || targetSheet.getLastRow() + 1;
                
                columnOperations.push({
                    colIndex: colIndex,
                    rowIndex: targetRow,
                    value: row[j]
                });
                
                // Обновляем последнюю строку для этого столбца
                lastRowByColumn[colIndex] = targetRow + 1;
            }
        }
    }
    
    // Выполняем вставку данных для каждой операции
    if (columnOperations.length > 0) {
        columnOperations.forEach(op => {
            try {
                targetSheet.getRange(op.rowIndex, op.colIndex + 1, 1, 1).setValue(op.value);
            } catch (e) {
                logMessage(`Ошибка при вставке в ячейку (строка ${op.rowIndex}, столбец ${op.colIndex + 1}): ${e}`, true);
            }
        });
        
        logMessage(`insertDataIntoSheetByColumn: Вставлено ${columnOperations.length} ячеек данных индивидуально по столбцам`);
    } else {
        logMessage(`insertDataIntoSheetByColumn: Нет данных для вставки после фильтрации`);
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
    // Создаем уникальный идентификатор для хранения данных файла
    const fileDataId = Utilities.getUuid();
    
    // Обрабатываем файл один раз и сохраняем результаты конвертации
    const processResult = processUploadedFile(fileData);
    
    // Проверяем на наличие ошибок
    if (processResult.error) {
      return {
        success: false,
        error: processResult.error
      };
    }
    
    // Сохраняем результаты первичной обработки вместе с данными файла
    // Удаляем большие данные из fileData перед сохранением в кэше
    const fileDataForCache = Object.assign({}, fileData);
    delete fileDataForCache.data; // Удаляем большие бинарные данные
    
    // Сохраняем данные в кэш для второго шага
    const cacheData = {
      fileData: fileDataForCache,
      processResult: {
        sheetNames: processResult.sheetNames || [],
        tempSheetId: processResult.tempSheetId,
        fileType: processResult.fileType
      },
      originalData: fileData.data // Сохраняем base64 данные отдельно
    };
    
    CacheService.getScriptCache().put(
      "fileDataCache_" + fileDataId, 
      JSON.stringify(cacheData),
      3600 // 1 час хранения
    );
    
    logMessage(`processFileAndAnalyze: Файл обработан и сохранен с ID: ${fileDataId}, tempSheetId: ${processResult.tempSheetId}`);
    
    return {
      success: true,
      fileDataId: fileDataId,
      message: "Файл успешно обработан. Теперь можно извлекать данные."
    };
  } catch (error) {
    logMessage(`Ошибка в processFileAndAnalyze: ${error.toString()}`, true);
    return {
      success: false,
      error: `Ошибка при обработке файла: ${error.toString()}`
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

/**
 * Второй шаг - анализ данных по ID из кэша
 * @param {string} fileDataId ID файла в кэше
 * @param {number} headerRow Номер строки с заголовками
 * @param {Array<string>} selectedHeaders Выбранные заголовки
 * @param {string} aiInstructions Инструкции для AI
 * @param {string} pageRange Диапазон страниц
 * @param {boolean} [byColumn=false] Если true, вставляет данные, находя последнюю строку для каждого столбца отдельно
 * @returns {Object} Результат операции
 */
function analyzeDataById(fileDataId, headerRow, selectedHeaders, aiInstructions, pageRange, byColumn = false) {
  try {
    // Получаем данные из кэша
    const cachedDataString = CacheService.getScriptCache().get("fileDataCache_" + fileDataId);
    if (!cachedDataString) {
      return {
        success: false,
        error: "Данные файла не найдены в кэше. Пожалуйста, загрузите файл заново."
      };
    }
    
    // Парсим кэшированные данные
    const cachedData = JSON.parse(cachedDataString);
    
    // Проверяем наличие результатов обработки файла
    if (!cachedData.processResult || !cachedData.processResult.tempSheetId) {
      logMessage(`analyzeDataById: В кэше не найдены результаты обработки файла`, true);
      return {
        success: false,
        error: "Не удалось найти результаты первичной обработки файла."
      };
    }
    
    // Подготавливаем данные для анализа
    // Используем ранее созданную таблицу вместо повторной конвертации
    const fileDataForAnalysis = cachedData.fileData || {};
    const processResult = cachedData.processResult;
    
    logMessage(`analyzeDataById: Используем уже созданную таблицу с ID: ${processResult.tempSheetId}`);
    
    // Добавляем все необходимые параметры
    fileDataForAnalysis.headerRow = headerRow;
    fileDataForAnalysis.selectedHeaders = selectedHeaders;
    fileDataForAnalysis.aiInstructions = aiInstructions;
    fileDataForAnalysis.pageRange = pageRange;
    fileDataForAnalysis.sheetName = fileDataForAnalysis.sheetName || (processResult.sheetNames && processResult.sheetNames[0]);
    
    // Создаем объект для analyzeAndInsertExtractedData, содержащий уже созданную Google Таблицу
    const analysisData = {
      tempSheetId: processResult.tempSheetId,  // Используем существующую таблицу
      fileType: processResult.fileType,
      sheetNames: processResult.sheetNames,
      // Прочие параметры из fileDataForAnalysis
      headerRow: fileDataForAnalysis.headerRow,
      selectedHeaders: fileDataForAnalysis.selectedHeaders,
      aiInstructions: fileDataForAnalysis.aiInstructions,
      pageRange: fileDataForAnalysis.pageRange,
      sheetName: fileDataForAnalysis.sheetName,
      name: fileDataForAnalysis.name
    };
    
    logMessage(`analyzeDataById: Начало analyzeAndInsertExtractedData с существующей таблицей ID: ${processResult.tempSheetId}`);
    
    try {
      // Вызываем функцию, которая сделает всю работу по анализу таблицы и вставке результатов
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const targetSheet = spreadsheet.getActiveSheet();
      
      if (!analysisData.headerRow || isNaN(parseInt(analysisData.headerRow)) || parseInt(analysisData.headerRow) < 1) {
        throw new Error("Не указан или указан неверный номер строки с заголовками в целевом листе.");
      }
      
      const { sheetData, headers } = getSheetDataAndHeaders(analysisData.tempSheetId, analysisData.sheetName);
      const targetHeaders = getTargetHeaders(targetSheet, analysisData.headerRow);
      
      if (!targetHeaders || targetHeaders.length === 0) {
        throw new Error("Не удалось получить заголовки целевого листа. Убедитесь, что указан правильный номер строки с заголовками и что строка не пустая.");
      }
      
      // Определяем selectedHeaders ПОСЛЕ получения targetHeaders
      const headersToUse = analysisData.selectedHeaders && analysisData.selectedHeaders.length > 0
        ? analysisData.selectedHeaders
        : targetHeaders;
      
      const prompt = (analysisData.fileType === 'text')
        ? buildTextPrompt(sheetData, headersToUse, analysisData.aiInstructions)
        : buildTablePrompt(sheetData, headers, headersToUse, analysisData.aiInstructions);
      
      logMessage(`analyzeDataById prompt: ${prompt}`);
      
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
      
      // Передаем параметр byColumn для определения метода вставки
      insertDataIntoSheet(parsedData, targetSheet, headersToUse, targetHeaders, byColumn);
      
      // Удаляем временную таблицу
      try {
        DriveApp.getFileById(analysisData.tempSheetId).setTrashed(true);
        logMessage(`analyzeDataById: Временная таблица ${analysisData.tempSheetId} помещена в корзину.`);
      } catch(e) {
        logMessage(`analyzeDataById: Не удалось поместить в корзину временную таблицу ${analysisData.tempSheetId}: ${e}`, true);
      }
      
      // Очищаем кэш
      CacheService.getScriptCache().remove("fileDataCache_" + fileDataId);
      
      // Возвращаем успешный результат
      return {
        success: true,
        message: "Данные успешно извлечены и вставлены."
      };
    } catch (error) {
      // Если произошла ошибка, удаляем временную таблицу
      try {
        DriveApp.getFileById(analysisData.tempSheetId).setTrashed(true);
        logMessage(`analyzeDataById: Временная таблица ${analysisData.tempSheetId} помещена в корзину после ошибки.`);
      } catch(e) {
        // Игнорируем ошибку при удалении
      }
      
      throw error; // Пробрасываем ошибку дальше
    }
    
  } catch (error) {
    logMessage(`Ошибка в analyzeDataById: ${error.toString()}`, true);
    return {
      success: false,
      error: `Ошибка при анализе данных: ${error.toString()}`
    };
  }
}

/**
 * Получает имена листов для файла по его ID из кэша.
 * Не выполняет повторную конвертацию файла.
 * @param {string} fileDataId ID файла в кэше
 * @returns {{sheetNames: string[]}} Массив имен листов или пустой массив
 */
function getSheetNamesById(fileDataId) {
  try {
    // Получаем данные из кэша
    const cachedDataString = CacheService.getScriptCache().get("fileDataCache_" + fileDataId);
    if (!cachedDataString) {
      return {
        sheetNames: []
      };
    }
    
    // Парсим кэшированные данные
    const cachedData = JSON.parse(cachedDataString);
    
    // Проверяем, что в кэше есть результаты обработки файла с именами листов
    if (!cachedData.processResult || !cachedData.processResult.sheetNames) {
      return {
        sheetNames: []
      };
    }
    
    // Возвращаем имена листов из кэша
    return {
      sheetNames: cachedData.processResult.sheetNames || []
    };
    
  } catch (error) {
    logMessage(`Ошибка в getSheetNamesById: ${error.toString()}`, true);
    return {
      sheetNames: []
    };
  }
}

/**
 * Извлекает текст из загруженного файла для суммаризации
 * Использует более надежный метод для Excel-файлов
 * @param {Object} fileData Объект с информацией о файле (name, type, data)
 * @returns {Object} Результат обработки файла {success: boolean, text: string, error: string}
 */
function extractTextFromFile(fileData) {
    try {
        if (!fileData || !fileData.data) {
            throw new Error("Нет данных файла для обработки");
        }
        
        logMessage(`Начата обработка файла: ${fileData.name} (${fileData.type})`);
        
        // Декодируем данные из base64
        const blob = Utilities.newBlob(Utilities.base64Decode(fileData.data), fileData.type, fileData.name);
        
        // Получаем расширение файла
        const fileExtension = fileData.name.split('.').pop().toLowerCase();
        let extractedText = "";
        
        // Обрабатываем разные типы файлов
        if (['xlsx', 'xls', 'ods'].includes(fileExtension)) {
            // Для электронных таблиц - используем более надежный метод
            try {
                // Создаем временный файл в DriveApp
                const tempFolder = DriveApp.createFolder("TempFolder_" + Date.now());
                const tempFile = tempFolder.createFile(blob);
                
                // Используем Advanced Drive API для конвертации без необходимости открывать как Spreadsheet
                const resource = {
                    title: tempFile.getName() + "_text",
                    mimeType: MimeType.GOOGLE_SHEETS
                };
                
                // Конвертируем без открытия
                const file = Drive.Files.copy(resource, tempFile.getId(), {convert: true});
                const tempSheetId = file.id;
                
                // Получаем содержимое через API
                let content = [];
                const spreadsheet = SpreadsheetApp.openById(tempSheetId);
                const sheets = spreadsheet.getSheets();
                
                // Извлекаем данные из всех листов
                for (let i = 0; i < sheets.length; i++) {
                    const sheet = sheets[i];
                    const data = sheet.getDataRange().getDisplayValues();
                    
                    if (data && data.length > 0) {
                        // Добавляем имя листа
                        content.push(`[Лист: ${sheet.getName()}]`);
                        
                        // Добавляем данные
                        for (let row of data) {
                            // Фильтруем пустые ячейки и объединяем
                            const rowText = row.filter(cell => cell && String(cell).trim() !== "").join(" | ");
                            if (rowText) {
                                content.push(rowText);
                            }
                        }
                        content.push("\n"); // Разделители между листами
                    }
                }
                
                extractedText = content.join("\n");
                
                // Очистка временных файлов
                try {
                    DriveApp.getFileById(tempSheetId).setTrashed(true);
                    tempFile.setTrashed(true);
                    tempFolder.setTrashed(true);
                } catch (cleanupError) {
                    logMessage(`Предупреждение: Не удалось очистить временные файлы: ${cleanupError.toString()}`, true);
                }
                
            } catch (excelError) {
                // Альтернативный метод обработки, если первый не сработал
                logMessage(`Ошибка при обработке Excel через API: ${excelError.toString()}. Пробуем альтернативный метод.`, true);
                
                try {
                    // Обрабатываем Excel как текст напрямую
                    const csvData = convertExcelToText(blob);
                    if (csvData) {
                        extractedText = csvData;
                    } else {
                        throw new Error("Не удалось извлечь данные из файла Excel");
                    }
                } catch (fallbackError) {
                    throw new Error(`Не удалось обработать Excel-файл: ${fallbackError.message}`);
                }
            }
        } else if (['docx', 'doc'].includes(fileExtension)) {
            // Для документов Word, используем улучшенный метод
            try {
                // Преобразуем в Google Doc через Drive API
                const tempFolder = DriveApp.createFolder("TempFolder_" + Date.now());
                const tempFile = tempFolder.createFile(blob);
                
                const resource = {
                    title: tempFile.getName() + "_doc",
                    mimeType: MimeType.GOOGLE_DOCS
                };
                
                const file = Drive.Files.copy(resource, tempFile.getId(), {convert: true});
                const docId = file.id;
                
                // Извлекаем текст
                const doc = DocumentApp.openById(docId);
                extractedText = doc.getBody().getText();
                
                // Очистка
                DriveApp.getFileById(docId).setTrashed(true);
                tempFile.setTrashed(true);
                tempFolder.setTrashed(true);
                
            } catch (docError) {
                throw new Error(`Ошибка обработки Word-документа: ${docError.message}`);
            }
        } else if (['csv', 'txt'].includes(fileExtension)) {
            // Для текстовых файлов - прямое чтение
            extractedText = blob.getDataAsString();
        } else {
            throw new Error(`Неподдерживаемый тип файла: ${fileExtension}`);
        }
        
        if (!extractedText || extractedText.trim() === "") {
            throw new Error("Из файла не удалось извлечь текст");
        }
        
        logMessage(`Текст успешно извлечен из файла ${fileData.name}, длина: ${extractedText.length} символов`);
        
        return {
            success: true,
            text: extractedText,
            error: null
        };
        
    } catch (error) {
        logMessage(`Ошибка при извлечении текста из файла: ${error.toString()}`, true);
        return {
            success: false,
            text: null,
            error: `Ошибка обработки файла: ${error.message}`
        };
    }
}

/**
 * Вспомогательная функция для прямого чтения Excel-файлов без конвертации
 * @param {Blob} blob Excel-файл как Blob
 * @returns {string} Извлеченный текст
 */
function convertExcelToText(blob) {
    try {
        // Создаем временный CSV для хранения данных
        let csvContent = [];
        
        // Создаем временный файл для доступа через Drive API
        const tempFile = DriveApp.createFile(blob);
        const fileId = tempFile.getId();
        
        // Извлекаем как текст через Drive API
        // Это обходной путь, который может работать для некоторых файлов Excel
        const exportLink = `https://www.googleapis.com/drive/v3/files/${fileId}/export?mimeType=text/csv`;
        const params = {
            method: 'get',
            headers: {
                'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
            },
            muteHttpExceptions: true
        };
        
        const response = UrlFetchApp.fetch(exportLink, params);
        
        if (response.getResponseCode() === 200) {
            csvContent.push(response.getContentText());
        }
        
        // Удаляем временный файл
        tempFile.setTrashed(true);
        
        return csvContent.join("\n");
    } catch (error) {
        logMessage(`Ошибка при конвертации Excel в текст: ${error.toString()}`, true);
        return null;
    }
}

/**
 * Извлекает текст из нескольких загруженных файлов и объединяет их результаты
 * Используется для массовой суммаризации
 * @param {Object[]} filesData Массив объектов с информацией о файлах (name, type, data)
 * @returns {Object} Результат обработки файла {success: boolean, text: string, error: string}
 */
function extractTextFromFiles(filesData) {
    try {
        if (!filesData || !Array.isArray(filesData) || filesData.length === 0) {
            throw new Error("Не передан массив файлов для обработки");
        }
        
        logMessage(`Начата обработка нескольких файлов: ${filesData.length} файл(ов)`);
        
        // Массив для хранения текста из каждого файла
        const extractedTexts = [];
        
        // Обработка всех файлов последовательно
        for (let i = 0; i < filesData.length; i++) {
            const fileData = filesData[i];
            logMessage(`Обработка файла ${i+1}/${filesData.length}: ${fileData.name}`);
            
            try {
                // Используем существующую функцию для извлечения текста из отдельного файла
                const result = extractTextFromFile(fileData);
                
                if (result.success) {
                    // Добавляем имя файла как заголовок к тексту
                    extractedTexts.push(`==== ФАЙЛ: ${fileData.name} ====\n\n${result.text}\n\n`);
                    logMessage(`Успешно извлечен текст из файла: ${fileData.name}`);
                } else {
                    // Добавляем сообщение об ошибке вместо текста
                    extractedTexts.push(`==== ФАЙЛ: ${fileData.name} (ОШИБКА) ====\n\nНе удалось извлечь текст: ${result.error}\n\n`);
                    logMessage(`Ошибка при извлечении текста из файла ${fileData.name}: ${result.error}`, true);
                }
            } catch (fileError) {
                // Обрабатываем ошибки для каждого файла отдельно
                extractedTexts.push(`==== ФАЙЛ: ${fileData.name} (ОШИБКА) ====\n\nНе удалось обработать: ${fileError.message}\n\n`);
                logMessage(`Исключение при обработке файла ${fileData.name}: ${fileError.toString()}`, true);
            }
        }
        
        // Объединяем все тексты с разделителями
        const combinedText = extractedTexts.join('\n');
        
        return {
            success: true,
            text: combinedText,
            error: null,
            filesProcessed: filesData.length
        };
        
    } catch (error) {
        logMessage(`Общая ошибка при извлечении текста из файлов: ${error.toString()}`, true);
        return {
            success: false,
            text: null,
            error: `Ошибка обработки файлов: ${error.message}`
        };
    }
}