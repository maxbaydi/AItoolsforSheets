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