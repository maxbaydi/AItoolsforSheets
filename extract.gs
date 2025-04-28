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
            insertExtractedData(activeCell, parsedData, headers, insertMode);
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
        prompt += "3. Если значение содержит '|' или ';', замени эти символы на '\\n'.\n";
        prompt += "4. Не добавляй заголовки столбцов в ответ.\n";
        prompt += "5. Если для какого-то заголовка данных нет, верни пустую строку.\n";
        prompt += "6. Не добавляй никакой дополнительный текст или пояснения.\n";
        prompt += "7. Если данных одного типа несколько, извлеки их все и раздели символом '\\n'.\n";
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
        const dataByHeader = {};
        headers.forEach(header => dataByHeader[header] = []);

        const rows = answer.split(';').map(row => row.trim()).filter(row => row !== '');

        rows.forEach(row => {
            const values = row.split('|').map(v => v.trim());

            values.forEach((value, index) => {
                if (index < headers.length) {
                    // Заменяем специальную последовательность \n на реальный перенос строки
                    const processedValue = value.replace(/\\n/g, '\n');
                    
                    // Проверяем, содержит ли значение переносы строк
                    if (processedValue.includes('\n')) {
                        // Разделяем значение по переносам строк и добавляем каждую часть как отдельный элемент
                        const multipleValues = processedValue.split('\n').map(v => v.trim()).filter(v => v !== '');
                        multipleValues.forEach(v => dataByHeader[headers[index]].push(v));
                    } else {
                        // Если нет переносов строк, добавляем как обычно
                        dataByHeader[headers[index]].push(processedValue);
                    }
                }
            });
        });

        return dataByHeader;
    }

    throw new Error("Неподдерживаемый формат вывода в getExtractedData: " + outputFormat);
}

function insertExtractedData(startCell, dataByHeader, headers, insertMode) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const startRow = startCell.getRow();
    const startCol = startCell.getColumn();

    if (insertMode === 'row') {
        // Режим "В одну строку" - каждое значение в своей колонке в одной строке
        for (let i = 0; i < headers.length; i++) {
            const header = headers[i];
            const cell = startCell.offset(0, i);
            cell.setValue(dataByHeader[header].join('\n'));
            cell.setWrap(true);
        }
    } else {
        // Режим "По строкам"
        // Проверяем, есть ли заголовки с несколькими значениями
        let hasMultipleValues = false;
        let maxDataCount = 0;
        
        for (const header of headers) {
            if (dataByHeader[header].length > 1) {
                hasMultipleValues = true;
            }
            maxDataCount = Math.max(maxDataCount, dataByHeader[header].length);
        }
        
        // Если у всех заголовков только одно значение, вставляем как в режиме "В одну строку"
        if (!hasMultipleValues) {
            for (let i = 0; i < headers.length; i++) {
                const header = headers[i];
                const cell = startCell.offset(0, i);
                cell.setValue(dataByHeader[header][0] || '');
            }
            return;
        }
        
        // Если есть заголовки с несколькими значениями, создаем таблицу без заголовков
        // Вставляем только данные
        for (let i = 0; i < maxDataCount; i++) {
            for (let j = 0; j < headers.length; j++) {
                const header = headers[j];
                const value = dataByHeader[header][i] || '';
                startCell.offset(i, j).setValue(value);
            }
        }
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