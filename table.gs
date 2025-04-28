// Функции для работы с таблицами
function createTable(query, tableStyle, keywords, includeExamples) {
    try {
        if (!query || query.length < 10) {
            throw new Error('Запрос слишком короткий. Опишите подробнее, какую таблицу вы хотите создать.');
        }

        const sheet = SpreadsheetApp.getActiveSheet();
        const activeCell = sheet.getActiveCell();
        if (!activeCell) {
            throw new Error("Выберите ячейку, с которой начать создание таблицы.");
        }

        const settings = getSettings();
        const model = settings.model;
        const temperature = settings.temperature;
        const maxRetries = settings.retryAttempts;
        const maxTokens = settings.maxTokens;

        let prompt = buildCreateTablePrompt(query, tableStyle, keywords, includeExamples);
        logMessage(`createTable prompt:\n${prompt}`);

        const response = openRouterRequest(prompt, model, temperature, maxRetries, maxTokens);

        if (!response || !response.choices || !response.choices[0] || !response.choices[0].message) {
            throw new Error('Не удалось получить корректный ответ от AI. Попробуйте изменить запрос.');
        }

        let responseText = response.choices[0].message.content.trim();
        responseText = responseText
            .replace(/^```(?:csv)?\s*/i, '')
            .replace(/```\s*$/i, '')
            .replace(/<[^>]*>/g, '')
            .trim();

        logMessage(`createTable raw response:\n${responseText}`);

        // Поддержка CSV с собственным или стандартным разделителем
        let rows;
        if (responseText.includes(COLUMN_DELIMITER) && !responseText.includes(',')) {
            rows = Utilities.parseCsv(responseText, COLUMN_DELIMITER);
        } else {
            rows = Utilities.parseCsv(responseText);
        }

        if (!rows || rows.length === 0) {
            throw new Error('Ответ от AI пуст или некорректен после парсинга.');
        }

        const numRows = rows.length;
        const numCols = rows[0].length;
        sheet.getRange(activeCell.getRow(), activeCell.getColumn(), numRows, numCols).setValues(rows);

        return `Таблица (${numRows}x${numCols}) успешно создана.`;
    } catch (error) {
        logMessage(`Ошибка в createTable: ${error.toString()}`, true);
        throw new Error('Ошибка при создании таблицы: ' + error.message);
    }
}

function buildCreateTablePrompt(query, tableStyle, keywords, includeExamples) {
    let prompt = `Создай таблицу на основе следующего запроса: ${query}\n\n`;

    if (tableStyle === 'short') {
        prompt += "Таблица должна быть краткой, содержать только основные доступные тебе данные.\n";
    } else if (tableStyle === 'detailed') {
        prompt += "Таблица должна быть подробной, содержать максимум доступной тебе информации.\n";
    } else if (tableStyle === 'withoutHeaders') {
        prompt += "Таблица НЕ ДОЛЖНА содержать заголовки столбцов.\n";
    } else {
        prompt += "Создай обычную таблицу (средней детализации) с заголовками столбцов.\n";
    }

    if (keywords) {
        prompt += `Обязательно включи в таблицу информацию по следующим ключевым словам: ${keywords}.\n`;
    }

    if (includeExamples) {
        prompt += "\nПримеры запросов и таблиц:\n";
        prompt += "Запрос: Сравни iPhone 15 Pro и Samsung Galaxy S24 Ultra. Стиль: Подробный.\n";
        prompt += "Таблица:\nМодель¦Процессор¦Камера (Мп)¦Экран (дюймы)¦Цена (руб.)¦Время работы (часы)¦Вес (г)¦Защита\n";
        prompt += "iPhone 15 Pro¦A17 Bionic¦48+12+12¦6.1¦129990¦23¦187¦IP68\n";
        prompt += "Samsung Galaxy S24 Ultra¦Snapdragon 8 Gen 3¦200+12+10+10¦6.8¦139990¦25¦233¦IP68\n";
    }

    prompt += "\nВАЖНО: Ответ должен быть ТОЛЬКО в формате CSV!\n";
    return prompt;
}

/**
 * Объединяет ячейки в выделенном диапазоне построчно, разделяя значения в строке пробелами.
 */
function combineCellsByRows() {
    const range = SpreadsheetApp.getActiveRange();
    if (!range) {
        SpreadsheetApp.getUi().alert('Ошибка', 'Выделите диапазон ячеек для объединения.', SpreadsheetApp.getUi().ButtonSet.OK);
        return;
    }

    const values = range.getValues();
    const numRows = range.getNumRows();
    const numCols = range.getNumColumns();

    for (let i = 0; i < numRows; i++) {
        let combinedString = "";
        for (let j = 0; j < numCols; j++) {
            const cellValue = values[i][j];
            if (cellValue !== "") {
                combinedString += cellValue + " ";
            }
        }
        combinedString = combinedString.trim();
        range.getCell(i + 1, 1).setValue(combinedString);
        for (let j = 1; j < numCols; j++) {
            range.getCell(i + 1, j + 1).clearContent();
        }
    }
}

/**
 * Объединяет ячейки в выделенном диапазоне, разделяя значения пробелами.
 */
function combineCellsWithSpace() {
    combineCells(SpreadsheetApp.getActiveRange(), ' ');
}

/**
 * Объединяет ячейки в выделенном диапазоне, разделяя значения переносами строк.
 */
function combineCellsWithNewline() {
    combineCells(SpreadsheetApp.getActiveRange(), '\n', true);
}

/**
 * Объединяет ячейки в указанном диапазоне.
 * @param {GoogleAppsScript.Spreadsheet.Range} range Диапазон ячеек.
 * @param {string} separator Разделитель.
 * @param {boolean} [wrap=false] Нужно ли делать перенос текста в первой ячейке.
 */
function combineCells(range, separator, wrap = false) {
    if (!range) {
        SpreadsheetApp.getUi().alert('Ошибка', 'Выделите диапазон ячеек для объединения.', SpreadsheetApp.getUi().ButtonSet.OK);
        return;
    }

    const values = range.getValues();
    let combinedString = "";

    for (let i = 0; i < values.length; i++) {
        for (let j = 0; j < values[i].length; j++) {
            const cellValue = values[i][j];
            if (cellValue !== "") {
                combinedString += cellValue + separator;
            }
        }
    }

    combinedString = combinedString.trim();

    const firstCell = range.getCell(1, 1);
    firstCell.setValue(combinedString);

    if (wrap) {
        firstCell.setWrap(true);
    }

    const numRows = range.getNumRows();
    const numCols = range.getNumColumns();
    if (numRows > 1 || numCols > 1) {
        for (let i = 0; i < numRows; i++) {
            for (let j = 0; j < numCols; j++) {
                if (i !== 0 || j !== 0) {
                    range.getCell(i + 1, j + 1).clearContent();
                }
            }
        }
    }
}

/**
 * Заполняет пустые ячейки в выделенном диапазоне.
 */
function fillCells() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getActiveSheet();
    const range = sheet.getActiveRange();

    if (!range) {
        SpreadsheetApp.getUi().alert('Ошибка', 'Выделите диапазон ячеек для заполнения.', SpreadsheetApp.getUi().ButtonSet.OK);
        return;
    }

    const values = range.getValues();
    const numRows = values.length;
    const numCols = values[0].length;

    // 1. Список координат пустых ячеек
    const emptyCellsCoordinates = [];
    for (let i = 0; i < numRows; i++) {
        for (let j = 0; j < numCols; j++) {
            if (values[i][j] === "") {
                emptyCellsCoordinates.push(`${i + 1},${j + 1}`);
            }
        }
    }

    // Проверка на наличие пустых ячеек
    if (emptyCellsCoordinates.length === 0) {
        SpreadsheetApp.getUi().alert('Нет пустых ячеек', 'В выделенном диапазоне нет пустых ячеек.', SpreadsheetApp.getUi().ButtonSet.OK);
        return;
    }

    // 2. Формируем CSV таблицы
    let csvData = "";
    for (let i = 0; i < numRows; i++) {
        let rowValues = [];
        for (let j = 0; j < numCols; j++) {
            rowValues.push(values[i][j] ? values[i][j] : "");
        }
        csvData += rowValues.join(COLUMN_DELIMITER) + (i < numRows - 1 ? ROW_DELIMITER : "");
    }

    // 3. Формируем промпт для полной таблицы в формате CSV
    const prompt = `Заполни пустые ячейки в следующей таблице и верни результат в формате CSV той же структуры без пояснений:` +
                   `\n${csvData}`;

    logMessage(`fillCells prompt CSV:\n${prompt}`);

    const settings = getSettings();
    const model = settings.model;
    const temperature = settings.temperature;
    const retries = settings.retryAttempts;
    const maxTokens = settings.maxTokens;
    let filledTable = [];

    let attempt = 0;
    let success = false;
    let aiResponse;

    while (attempt < retries && !success) {
        attempt++;
        logMessage(`Попытка ${attempt} из ${retries}`);
        try {
            aiResponse = openRouterRequest(prompt, model, temperature, retries, maxTokens);
            if (!aiResponse?.choices?.[0]?.message) {
                throw new Error("Ошибка: Неожиданный ответ от OpenRouter в fillCells");
            }
            // Обрабатываем полный CSV ответа
            let csvResponse = aiResponse.choices[0].message.content.trim()
                .replace(/^```csv\s*/i, '')
                .replace(/```\s*$/i, '')
                .trim();
            logMessage(`fillCells CSV response (попытка ${attempt}):\n${csvResponse}`);
            filledTable = Utilities.parseCsv(csvResponse, COLUMN_DELIMITER);
            if (filledTable.length !== numRows || filledTable[0].length !== numCols) {
                throw new Error(`Неверный формат таблицы: ожидается ${numRows}x${numCols}, получено ${filledTable.length}x${filledTable[0]?.length}`);
            }
            success = true;
        } catch (error) {
            logMessage(`Ошибка в fillCells (попытка ${attempt}): ${error.toString()}`, true);
            if (attempt === retries) {
                throw new Error(`Не удалось получить корректный ответ от AI после ${retries} попыток.`);
            }
        }
    }

    // Проверка успеха
    if (!success) {
        throw new Error(`Не удалось получить и обработать ответ от AI после ${retries} попыток.`);
    }

    if (success && aiResponse) {
        const promptPrefix = prompt.substring(0, prompt.indexOf('\n'));
        const effectiveModel = SCRIPT_PROPERTIES.getProperty(MODEL_SETTING_KEY) || model || DEFAULT_MODEL;
        const effectiveTemperature = parseFloat(SCRIPT_PROPERTIES.getProperty(TEMPERATURE_SETTING_KEY)) || temperature || DEFAULT_TEMPERATURE;
        const effectiveMaxTokens = parseInt(SCRIPT_PROPERTIES.getProperty(MAX_TOKENS_SETTING_KEY), 10) ||  DEFAULT_MAX_TOKENS;

        const promptHash = calculateMD5(promptPrefix + effectiveModel + effectiveTemperature + effectiveMaxTokens);
        const cacheKey = `promptHash:${promptHash}`;
        CACHE.put(cacheKey, JSON.stringify(aiResponse), 21600);
        logMessage(`fillCells: Ответ AI сохранен в кэше для ключа: ${cacheKey}`);
    }

    // 5. Заполняем пустые ячейки из заполненной таблицы
    for (let i = 0; i < numRows; i++) {
        for (let j = 0; j < numCols; j++) {
            if (values[i][j] === "" && filledTable[i][j] !== "") {
                range.getCell(i + 1, j + 1).setValue(filledTable[i][j]);
            }
        }
    }

    SpreadsheetApp.getUi().alert('Ячейки заполнены', 'Пустые ячейки в выделенном диапазоне были заполнены.', SpreadsheetApp.getUi().ButtonSet.OK);
}