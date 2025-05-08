// Функции для извлечения данных
function extractData(sourceRangeStr, headersStr, outputFormat, insertMode) {
    try {
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = spreadsheet.getActiveSheet();

        let sourceRange;
        try {
            sourceRange = sheet.getRange(sourceRangeStr);
        } catch (e) {
            throw new Error("Неверный формат диапазона с данными: " + e.message);
        }

        const sourceValues = sourceRange.getValues();
        const headers = headersStr.split(',').map(h => h.trim()).filter(h => h !== "");
        if (headers.length === 0) {
            throw new Error("Укажите заголовки данных.");
        }

        const activeCell = sheet.getActiveCell();
        if (!activeCell) {
            throw new Error("Выберите ячейку для вставки.");
        }

        const prompt = buildExtractDataPrompt(sourceValues, headers, 'extract');
        logMessage(`extractData prompt:\n${prompt}`);

        // Получаем настройки из свойств скрипта
        const settings = getSettings();
        const model = settings.model;
        const temperature = settings.temperature;
        const maxTokens = settings.maxTokens;

        try {
            const aiResponse = openRouterRequest(prompt, model, temperature, 3, maxTokens);
            const parsedData = getExtractedData(aiResponse, outputFormat, headers);
            insertExtractedData(activeCell, parsedData, headers);
            return "Данные извлечены.";
        } catch (error) {
            logMessage(`Ошибка в extractData: ${error.toString()}`, true);
            throw error;
        }
    } catch (error) {
        logMessage(`Ошибка в extractData: ${error.toString()}`, true);
        throw new Error('Ошибка извлечения: ' + error.message);
    }
}

function buildExtractDataPrompt(sourceValues, headers, mode) {
    let prompt = "";

    if (mode === 'extract') {
        prompt = `Извлеки из текста следующие данные: ${headers.join(', ')}. Верни данные в формате CSV.\n`;
        prompt += "Текст:\n";

        sourceValues.forEach((row, rowIndex) => {
            row.forEach((value, colIndex) => {
                prompt += `Ячейка ${rowIndex + 1},${colIndex + 1}: ${value || ""}\n`;
            });
        });

        prompt += "Правила форматирования ответа:\n";
        prompt += "1. Используй разделитель столбцов: '|'.\n";
        prompt += "2. Используй разделитель строк: ';'.\n";
        prompt += "3. ВАЖНО! Каждая отдельная сущность (например, компания, человек, продукт и т.д.) должна быть в отдельной строке.\n";
        prompt += "4. ВАЖНО! Если для одного поля у одной сущности есть несколько значений (например, несколько телефонов одной компании), объедини их внутри поля, используя символ '~' как разделитель.\n";
        prompt += "5. Если значение содержит '|', ';' или '~', замени эти символы на '\\n'.\n";
        prompt += "6. Не добавляй заголовки столбцов в ответ.\n";
        prompt += "7. Если для какого-то заголовка данных нет, верни пустую строку.\n";
        prompt += "8. Не добавляй никакой дополнительный текст или пояснения.\n";
        prompt += "9. Для каждой сущности (каждой строки) строго используй количество полей равное количеству заголовков (${headers.length}).\n";
        prompt += "10. Не добавляй лишний символ разделителя '|' в конце строки.\n";
        prompt += "11. Используй только символ ';' в качестве разделителя между разными сущностями, не используй перенос строки.\n";
        prompt += "12. Пример правильного ответа для двух компаний, где у одной два телефона, а у другой два адреса:\n";
        prompt += "Компания A|Адрес компании A|Телефон1~Телефон2|Email компании A;Компания B|Адрес1~Адрес2|Телефон компании B|Email компании B\n";
        return prompt;
    }

    throw new Error("Неверный режим для buildExtractDataPrompt: " + mode);
}

function getExtractedData(aiResponse, outputFormat, headers) {
    if (!aiResponse?.choices?.[0]?.message) {
        throw new Error("Неожиданный ответ от OpenRouter");
    }

    let answer = aiResponse.choices[0].message.content.trim();
    logMessage(`getExtractedData answer:\n${answer}`);

    answer = answer.replace(/^```csv\s*/i, '').replace(/```\s*$/i, '').trim();

    if (outputFormat === 'csv') {
        // Создаем массив объектов, где каждый объект представляет одну сущность
        const entities = [];
        
        // Изменяем разделитель строк с ';' на '\n' или оба варианта
        const rows = answer.split(/[\n;]/).map(row => row.trim()).filter(row => row !== '');

        rows.forEach(row => {
            const values = row.split('|').map(v => v.trim());
            const entity = {};
            
            values.forEach((value, index) => {
                if (index < headers.length) {
                    const header = headers[index];
                    
                    // Заменяем специальную последовательность \n на реальный перенос строки
                    let processedValue = value.replace(/\\n/g, '\n');
                    
                    // Обрабатываем множественные значения, разделенные символом '~'
                    if (processedValue.includes('~')) {
                        // Разделяем на массив значений
                        entity[header] = processedValue.split('~').map(v => v.trim()).filter(v => v !== '');
                    } else {
                        entity[header] = [processedValue];
                    }
                }
            });
            
            entities.push(entity);
        });
        
        // Преобразуем массив объектов обратно в формат dataByHeader для совместимости
        const dataByHeader = {};
        headers.forEach(header => dataByHeader[header] = []);
        
        entities.forEach(entity => {
            headers.forEach(header => {
                if (entity[header]) {
                    dataByHeader[header].push(entity[header].join('\n'));
                } else {
                    dataByHeader[header].push('');
                }
            });
        });
        
        return dataByHeader;
    }

    throw new Error("Неподдерживаемый формат вывода в getExtractedData: " + outputFormat);
}

function insertExtractedData(startCell, dataByHeader, headers) {
    const sheet = startCell.getSheet();
    const startRow = startCell.getRow();
    const startCol = startCell.getColumn();

    // Определяем максимальное количество строк для вставки
    let maxRows = 0;
    // Используем headers для итерации, так как dataByHeader может не содержать все ключи, если данных не было
    headers.forEach(header => {
        // Проверяем наличие ключа и что это массив
        if (dataByHeader[header] && Array.isArray(dataByHeader[header])) {
            maxRows = Math.max(maxRows, dataByHeader[header].length);
        }
    });

    if (maxRows === 0) {
        logMessage("insertExtractedData: Нет данных для вставки (maxRows = 0).");
        return; // Нет данных для вставки
    }

    // Определяем количество столбцов по количеству заголовков
    const numCols = headers.length;
    // Создаем пустой массив нужного размера
    const outputData = Array(maxRows).fill(null).map(() => Array(numCols).fill(""));

    // Заполняем массив данными
    for (let i = 0; i < headers.length; i++) {
        const header = headers[i]; // Используем оригинальный заголовок
        const colIndex = i; // Индекс столбца соответствует порядку в headers
        // Получаем данные по оригинальному заголовку
        const values = dataByHeader[header] || []; // Используем header, а не headerLower

        for (let j = 0; j < maxRows; j++) {
            // Проверяем, есть ли значение для этой строки в массиве values
            if (j < values.length && values[j] !== undefined && values[j] !== null) {
                 // Убедимся, что записываем строку
                outputData[j][colIndex] = String(values[j]);
            }
            // Если значения нет или оно undefined/null, оставляем пустую строку "" (уже установлено при создании outputData)
        }
    }

    // Вставляем данные одним вызовом
    if (outputData.length > 0) {
        logMessage(`insertExtractedData: Вставка ${outputData.length} строк и ${numCols} столбцов в диапазон ${sheet.getName()}!${startCell.getA1Notation()}`);
        sheet.getRange(startRow, startCol, outputData.length, numCols).setValues(outputData);
         logMessage(`insertExtractedData: Вставка завершена.`);
    } else {
         logMessage("insertExtractedData: Массив outputData пуст, вставка не выполнена.");
    }
}

/**
 * Возвращает текст выделенного диапазона как одну строку (для клиента).
 * Используется для передачи текста в AI-сайдбары.
 */
function getSelectedRangeText() {
    const range = SpreadsheetApp.getActiveRange();
    if (!range) {
        return null;
    }
    const values = range.getValues();
    // Объединяем все значения в одну строку, убираем пустые и схлопываем массивы
    return values.flat().filter(v => v !== "" && v != null).join(" ");
}

/**
 * Возвращает A1 нотацию выделенного диапазона (для клиента).
 */
function getSelectedRangeA1Notation() {
    const range = SpreadsheetApp.getActiveRange();
    if (!range) {
        return null;
    }
    return range.getA1Notation();
}

/**
 * Суммаризирует предоставленный текст, сохраняя базовое форматирование
 * @param {string} text Текст для суммаризации
 * @param {number} temperature Температура для генерации (опционально)
 * @returns {string} Суммаризированный текст с сохраненным форматированием
 */
function summarizeText(text, temperature) {
    try {
        if (!text || text.trim() === "") {
            throw new Error("Нет текста для суммаризации");
        }

        // Получаем настройки из свойств скрипта
        const settings = getSettings();
        const model = settings.model;
        const temp = temperature !== undefined && temperature !== null ? temperature : settings.temperature;
        const maxTokens = settings.maxTokens;
        const maxRetries = settings.retryAttempts;

        // Строим промпт для подсчета токенов
        const prompt = buildSummarizePromptWithFormatting(text);
        
        // Оцениваем количество токенов в тексте и промпте
        const promptTokens = estimateTokenCount(prompt);
        const textTokens = estimateTokenCount(text);
        
        // Проверяем, не превышает ли количество токенов максимально допустимое значение
        if (promptTokens > maxTokens) {
            throw new Error(`В полученном тексте слишком много символов (${textTokens} токенов в тексте + ${promptTokens - textTokens} токенов в промпте = ${promptTokens} токенов), чем максимально возможное (${maxTokens} токенов)`);
        }
        
        logMessage(`summarizeText prompt: ${prompt.substring(0, 200)}... [Примерно ${promptTokens} токенов]`);

        // Делаем запрос к API
        const aiResponse = openRouterRequest(prompt, model, temp, maxRetries, maxTokens);
        
        // Извлекаем текст ответа из JSON объекта
        if (!aiResponse || !aiResponse.choices || !aiResponse.choices[0] || !aiResponse.choices[0].message) {
            throw new Error("Неожиданный формат ответа от API");
        }
        
        // Берем содержимое ответа
        let summarizedText = aiResponse.choices[0].message.content.trim();
        
        // Очищаем текст только от технических элементов, сохраняя форматирование
        
        // Удаляем блоки кода в тройных обратных кавычках
        summarizedText = summarizedText.replace(/```[\s\S]*?```/g, '');
        
        // Удаляем обратные кавычки для кода, но сохраняем их содержимое
        summarizedText = summarizedText.replace(/`([^`]*)`/g, '$1');
        
        // Преобразуем заголовки в жирный текст (вместо удаления)
        summarizedText = summarizedText.replace(/^\s*#{1,6}\s+(.*)$/gm, '**$1**');
        
        // Удаляем HTML-теги
        summarizedText = summarizedText.replace(/<[^>]*>/g, '');
        
        // Сохраняем жирный текст, курсив, маркированные и нумерованные списки
        // (удаляем соответствующие строки из предыдущей версии функции)
        
        // Очищаем лишние пробелы, но сохраняем структуру текста
        summarizedText = summarizedText.replace(/[ \t]+/g, ' ').trim();
        
        // Нормализуем, но сохраняем двойные переносы строк для абзацев
        summarizedText = summarizedText.replace(/\n{4,}/g, '\n\n\n');
        
        const outputTokens = estimateTokenCount(summarizedText);
        logMessage(`summarizeText: Текст подготовлен с сохранением форматирования, итоговый размер: ${summarizedText.length} символов (примерно ${outputTokens} токенов)`);
        
        return summarizedText;
    } catch (error) {
        logMessage(`Ошибка в summarizeText: ${error.toString()}`, true);
        throw new Error('Ошибка суммаризации: ' + error.message);
    }
}

/**
 * Строит промпт для суммаризации текста с сохранением форматирования
 * @param {string} text Текст для суммаризации
 * @returns {string} Промпт для суммаризации
 */
function buildSummarizePromptWithFormatting(text) {
    let prompt = "Суммаризируй следующий текст, сохраняя все ключевые моменты и сокращая избыточную информацию. ";
    prompt += "Суммаризация должна быть содержательной, логичной и связной. ";
    prompt += "Старайся использовать до 30% от длины исходного текста. ";
    prompt += "\nВАЖНО: Используй форматирование markdown в своем ответе: ";
    prompt += "**жирный шрифт** для важных утверждений, ";
    prompt += "*курсив* для выделения ключевых терминов, ";
    prompt += "- маркированные списки для перечислений, ";
    prompt += "1. нумерованные списки для последовательных шагов или приоритетов, ";
    prompt += "и структурируй текст с помощью абзацев, разделенных пустой строкой.\n\n";
    prompt += "Текст для суммаризации:\n\n";
    prompt += text;
    return prompt;
}

/**
 * Вставляет текст в активную ячейку
 * @param {string} text Текст для вставки
 */
function insertTextIntoActiveCell(text) {
    try {
        const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        const activeSheet = activeSpreadsheet.getActiveSheet();
        const activeCell = activeSheet.getActiveCell();
        
        if (!activeCell) {
            throw new Error("Не выбрана активная ячейка");
        }
        
        activeCell.setValue(text);
        return "Текст вставлен в активную ячейку";
    } catch (error) {
        logMessage(`Ошибка в insertTextIntoActiveCell: ${error.toString()}`, true);
        throw new Error("Ошибка при вставке текста: " + error.message);
    }
}

/**
 * Вставляет форматированный текст в активную ячейку, преобразуя разметку markdown
 * @param {string} markdownText Текст с разметкой markdown
 * @returns {string} Сообщение о результате
 */
function insertFormattedTextIntoActiveCell(markdownText) {
    try {
        const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        const activeSheet = activeSpreadsheet.getActiveSheet();
        const activeCell = activeSheet.getActiveCell();
        
        if (!activeCell) {
            throw new Error("Не выбрана активная ячейка");
        }
        
        // Создаем Rich Text объект
        const richTextBuilder = SpreadsheetApp.newRichTextValue().setText(markdownText);
        
        // Обрабатываем жирный текст
        const boldPattern = /\*\*([^*]+)\*\*/g;
        let boldMatch;
        while ((boldMatch = boldPattern.exec(markdownText)) !== null) {
            const startIndex = boldMatch.index;
            const endIndex = boldMatch.index + boldMatch[0].length;
            // Получаем текст без символов разметки
            const originalText = boldMatch[1];
            // Удаляем маркеры жирного текста
            const plainText = markdownText.substring(0, startIndex) + 
                             originalText + 
                             markdownText.substring(endIndex);
            
            // Применяем форматирование для этого участка текста
            richTextBuilder.setText(plainText)
                          .setTextStyle(startIndex, startIndex + originalText.length, 
                                       SpreadsheetApp.newTextStyle()
                                       .setBold(true)
                                       .build());
            
            // Обновляем markdownText для последующих итераций
            markdownText = plainText;
            // Сбрасываем индекс поиска
            boldPattern.lastIndex = 0;
        }
        
        // Обрабатываем курсив
        const italicPattern = /\*([^*]+)\*/g;
        let italicMatch;
        while ((italicMatch = italicPattern.exec(markdownText)) !== null) {
            const startIndex = italicMatch.index;
            const endIndex = italicMatch.index + italicMatch[0].length;
            const originalText = italicMatch[1];
            const plainText = markdownText.substring(0, startIndex) + 
                             originalText + 
                             markdownText.substring(endIndex);
            
            richTextBuilder.setText(plainText)
                          .setTextStyle(startIndex, startIndex + originalText.length, 
                                       SpreadsheetApp.newTextStyle()
                                       .setItalic(true)
                                       .build());
            
            markdownText = plainText;
            italicPattern.lastIndex = 0;
        }
        
        // Устанавливаем значение ячейки как Rich Text
        activeCell.setRichTextValue(richTextBuilder.build());
        
        return "Форматированный текст вставлен в активную ячейку";
    } catch (error) {
        logMessage(`Ошибка в insertFormattedTextIntoActiveCell: ${error.toString()}`, true);
        throw new Error("Ошибка при вставке форматированного текста: " + error.message);
    }
}