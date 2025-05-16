// Функции для работы с таблицами
function createTable(query, tableStyle, keywords, includeExamples, includeFormatting = false) {
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

        let prompt = buildCreateTablePrompt(query, tableStyle, keywords, includeExamples, includeFormatting);
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

        // Проверяем, включает ли ответ инструкции по форматированию
        let formattingInstructions = null;
        if (includeFormatting) {
            const formattingResult = parseFormattingInstructions(responseText);
            responseText = formattingResult.data;
            formattingInstructions = formattingResult.formatting;
        }

        // Улучшенная логика для парсинга CSV - всегда используем COLUMN_DELIMITER, если он есть
        let rows;
        if (responseText.includes(COLUMN_DELIMITER)) {
            // Используем COLUMN_DELIMITER, даже если есть запятые
            rows = Utilities.parseCsv(responseText, COLUMN_DELIMITER);
        } else {
            // Если специальный разделитель не найден, используем стандартный запятую
            rows = Utilities.parseCsv(responseText);
        }

        if (!rows || rows.length === 0) {
            throw new Error('Ответ от AI пуст или некорректен после парсинга.');
        }

        const numRows = rows.length;
        const numCols = rows[0].length;
        const tableRange = sheet.getRange(activeCell.getRow(), activeCell.getColumn(), numRows, numCols);
        tableRange.setValues(rows);
        
        // Применяем форматирование, если оно было запрошено и получено
        if (includeFormatting && formattingInstructions) {
            applyTableFormatting(tableRange, formattingInstructions, tableStyle);
        }

        return `Таблица (${numRows}x${numCols}) успешно создана${includeFormatting ? " с форматированием" : ""}.`;
    } catch (error) {
        logMessage(`Ошибка в createTable: ${error.toString()}`, true);
        throw new Error('Ошибка при создании таблицы: ' + error.message);
    }
}

function buildCreateTablePrompt(query, tableStyle, keywords, includeExamples, includeFormatting = false) {
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
    prompt += "ОЧЕНЬ ВАЖНО: Используй символ '¦' (вертикальная черта) как разделитель столбцов, а НЕ запятую!\n";
    prompt += "Если в данных есть запятые, оставляй их как есть - не используй запятые в качестве разделителя столбцов.\n";
    
    if (includeFormatting) {
        prompt += "\nДОБАВЬ ИНСТРУКЦИИ ПО ФОРМАТИРОВАНИЮ:\n";
        prompt += "После данных таблицы добавь блок с инструкциями по форматированию, начав с новой строки '@FORMATTING:'\n";
        prompt += "Для каждой строки или ячейки можно указать форматирование в формате 'тип:строка:столбец:инструкция'\n";
        
        prompt += "\nОЧЕНЬ ВАЖНО: СДЕЛАЙ ФОРМАТИРОВАНИЕ КРЕАТИВНЫМ И УНИКАЛЬНЫМ!\n";
        prompt += "- Используй разные цвета для разных строк и столбцов, а не один и тот же цвет для всех\n";
        prompt += "- Выделяй важные данные яркими цветами и жирным шрифтом\n";
        prompt += "- Используй разные типы форматирования для разных типов данных\n";
        prompt += "- Не форматируй все строки одинаково\n";
        prompt += "- Для цветов используй разные оттенки, не только базовые цвета\n";
        
        prompt += "\nПримеры креативного форматирования таблицы:\n";
        prompt += "@FORMATTING:\n";
        prompt += "header:1:all:bold,background:#4285F4,color:#FFFFFF,fontSize:14,alignH:center\n";
        prompt += "column:all:1:bold,background:#E8F0FE,width:120\n";
        prompt += "row:2:all:background:#F1F8E9\n";
        prompt += "row:3:all:background:#FFF3E0\n";
        prompt += "row:4:all:background:#E3F2FD\n";
        prompt += "row:5:all:background:#F3E5F5\n";
        prompt += "cell:2:3:bold,color:#0F9D58\n";
        prompt += "cell:3:4:bold,color:#DB4437\n";
        prompt += "cell:4:2:color:#4285F4,italic\n";
        prompt += "cell:3:2:bold,color:#F4B400\n";
        
        prompt += "\nДоступные инструкции форматирования:\n";
        prompt += "- bold - жирный шрифт\n";
        prompt += "- italic - курсив\n";
        prompt += "- color:#RRGGBB - цвет текста (например, color:#FF0000 для красного)\n";
        prompt += "- background:#RRGGBB - цвет фона (например, background:#E6F2FF для светло-голубого)\n";
        prompt += "- wrap - перенос по словам\n";
        prompt += "- alignH:[left|center|right] - горизонтальное выравнивание\n";
        prompt += "- alignV:[top|middle|bottom] - вертикальное выравнивание\n";
        prompt += "- fontSize:[8-24] - размер шрифта\n";
        prompt += "- width:[50-250] - ширина столбца в пикселях\n\n";
        prompt += "Обязательно укажи width для всех столбцов, чтобы избежать слишком широких столбцов.\n";
        prompt += "Для столбцов с короткими данными (числа, годы, имена) используй меньшие значения width (50-100).\n";
        prompt += "Для столбцов с длинным текстом используй wrap и width:150-200.\n\n";
        prompt += "Выбери подходящее форматирование исходя из содержания и типа таблицы. Будь креативным.\n";
    }
    
    return prompt;
}

/**
 * Парсит инструкции форматирования из ответа AI.
 * @param {string} responseText - Текст ответа от AI
 * @returns {Object} Объект с данными таблицы и инструкциями форматирования
 */
function parseFormattingInstructions(responseText) {
    const result = {
        data: responseText,
        formatting: []
    };
    
    // Проверяем наличие блока форматирования
    const formattingMarker = '@FORMATTING:';
    const formattingIndex = responseText.indexOf(formattingMarker);
    
    if (formattingIndex === -1) {
        return result; // Нет инструкций форматирования
    }
    
    // Разделяем данные и инструкции форматирования
    result.data = responseText.substring(0, formattingIndex).trim();
    const formattingText = responseText.substring(formattingIndex + formattingMarker.length).trim();
    
    // Парсим инструкции форматирования
    const formattingLines = formattingText.split('\n');
    for (const line of formattingLines) {
        const trimmedLine = line.trim();
        if (!trimmedLine) continue;
        
        // Парсим инструкцию форматирования
        // Формат: тип:строка:столбец:инструкция
        const parts = trimmedLine.split(':');
        if (parts.length < 4) continue;
        
        const type = parts[0].trim();
        const row = parts[1].trim();
        const column = parts[2].trim();
        const instructions = parts[3].trim();
        
        result.formatting.push({
            type,
            row,
            column,
            instructions
        });
    }
    
    return result;
}

/**
 * Применяет инструкции форматирования к диапазону.
 * @param {GoogleAppsScript.Spreadsheet.Range} range - Диапазон ячеек
 * @param {string} instructionsText - Текст инструкций форматирования
 */
function applyFormatting(range, instructionsText) {
    if (!range || !instructionsText) return;
    
    const instructions = instructionsText.split(',');
    for (const instruction of instructions) {
        const trimmed = instruction.trim();
        
        if (trimmed === 'bold') {
            range.setFontWeight('bold');
        } else if (trimmed === 'italic') {
            range.setFontStyle('italic');
        } else if (trimmed === 'wrap') {
            range.setWrap(true);
        } else if (trimmed.startsWith('color:')) {
            const color = trimmed.substring(6).trim();
            if (color.match(/^#[0-9A-Fa-f]{6}$/)) {
                range.setFontColor(color);
            }
        } else if (trimmed.startsWith('background:')) {
            const color = trimmed.substring(11).trim();
            if (color.match(/^#[0-9A-Fa-f]{6}$/)) {
                range.setBackground(color);
            }
        } else if (trimmed.startsWith('alignH:')) {
            const alignment = trimmed.substring(7).trim().toLowerCase();
            if (alignment === 'left') {
                range.setHorizontalAlignment('left');
            } else if (alignment === 'center') {
                range.setHorizontalAlignment('center');
            } else if (alignment === 'right') {
                range.setHorizontalAlignment('right');
            }
        } else if (trimmed.startsWith('alignV:')) {
            const alignment = trimmed.substring(7).trim().toLowerCase();
            if (alignment === 'top') {
                range.setVerticalAlignment('top');
            } else if (alignment === 'middle') {
                range.setVerticalAlignment('middle');
            } else if (alignment === 'bottom') {
                range.setVerticalAlignment('bottom');
            }
        } else if (trimmed.startsWith('fontSize:')) {
            const size = parseInt(trimmed.substring(9).trim());
            if (!isNaN(size) && size >= 8 && size <= 24) {
                range.setFontSize(size);
            }
        }
        // Обработку width выполняем отдельно в функции applyTableFormatting
        // для получения доступа к sheet и точного управления шириной столбцов
    }
}

/**
 * Применяет инструкции форматирования к диапазону.
 * @param {GoogleAppsScript.Spreadsheet.Range} range - Диапазон ячеек таблицы
 * @param {Array} formattingInstructions - Массив инструкций форматирования
 * @param {string} tableStyle - Стиль таблицы (detailed, short, normal, withoutHeaders)
 */
function applyTableFormatting(range, formattingInstructions, tableStyle) {
    if (!range || !formattingInstructions || formattingInstructions.length === 0) {
        return;
    }
    
    const numRows = range.getNumRows();
    const numCols = range.getNumColumns();
    const sheet = range.getSheet();
    
    // Применяем базовое форматирование для всей таблицы
    range.setBorder(true, true, true, true, true, true, '#D3D3D3', SpreadsheetApp.BorderStyle.SOLID);
    
    // Настройка переноса слов для длинных текстов
    range.setWrap(true);
    
    // Создаем карту инструкций ширины столбцов
    const columnWidths = {};
    const columnsWithWidth = {};
    
    // Сначала собираем все инструкции по ширине столбцов
    for (const instruction of formattingInstructions) {
        if (instruction.type === 'column' && instruction.instructions.includes('width:')) {
            // Извлекаем значение width из инструкций
            const widthMatch = instruction.instructions.match(/width:(\d+)/);
            if (widthMatch && widthMatch.length > 1) {
                const width = parseInt(widthMatch[1]);
                
                if (instruction.column === 'all') {
                    // Устанавливаем ширину для всех столбцов
                    for (let col = 1; col <= numCols; col++) {
                        columnWidths[col] = width;
                        columnsWithWidth[col] = true;
                    }
                } else {
                    const col = parseInt(instruction.column);
                    if (!isNaN(col) && col >= 1 && col <= numCols) {
                        columnWidths[col] = width;
                        columnsWithWidth[col] = true;
                    }
                }
            }
        }
    }
    
    // Если есть ячейки или строки с указанием wrap, обрабатываем их
    for (const instruction of formattingInstructions) {
        if (instruction.instructions.includes('wrap')) {
            // Этот код применяется отдельно, чтобы не перезаписать другие форматирования
            if (instruction.type === 'cell') {
                const row = parseInt(instruction.row);
                const col = parseInt(instruction.column);
                if (!isNaN(row) && !isNaN(col) && row >= 1 && row <= numRows && col >= 1 && col <= numCols) {
                    range.getCell(row, col).setWrap(true);
                }
            } else if (instruction.type === 'row' && instruction.column === 'all') {
                const row = parseInt(instruction.row);
                if (!isNaN(row) && row >= 1 && row <= numRows) {
                    range.offset(row - 1, 0, 1, numCols).setWrap(true);
                }
            } else if (instruction.type === 'column' && instruction.row === 'all') {
                const col = parseInt(instruction.column);
                if (!isNaN(col) && col >= 1 && col <= numCols) {
                    range.offset(0, col - 1, numRows, 1).setWrap(true);
                }
            }
        }
    }
    
    // Если стиль не withoutHeaders и больше 1 строки, форматируем заголовки
    if (tableStyle !== 'withoutHeaders' && numRows > 1) {
        const headerRange = range.offset(0, 0, 1, numCols);
        headerRange.setFontWeight('bold');
        headerRange.setBackground('#E6F2FF');
        headerRange.setHorizontalAlignment('center');
        headerRange.setVerticalAlignment('middle');
    }
    
    // Если стиль detailed, добавляем чередующиеся цвета строк
    if (tableStyle === 'detailed' && numRows > 2) {
        for (let row = 2; row <= numRows; row++) {
            if (row % 2 === 0) {
                range.offset(row - 1, 0, 1, numCols).setBackground('#F8F9FA');
            }
        }
    }
    
    // Применяем пользовательские инструкции форматирования
    for (const instruction of formattingInstructions) {
        try {
            applyFormatInstruction(range, instruction, numRows, numCols);
        } catch (error) {
            logMessage(`Ошибка при применении форматирования: ${error.toString()}`, true);
        }
    }
    
    // Задаем ширину столбцов в последнюю очередь
    for (let col = 1; col <= numCols; col++) {
        const actualColumn = range.getCell(1, col).getColumn();
        if (columnsWithWidth[col]) {
            // Если у столбца задана ширина через инструкции, используем ее
            sheet.setColumnWidth(actualColumn, columnWidths[col]);
        } else {
            // Иначе используем автоподбор но с разумным ограничением
            sheet.autoResizeColumn(actualColumn);
            const currentWidth = sheet.getColumnWidth(actualColumn);
            
            // Определяем максимальную ширину на основе типа данных в столбце
            let isNumericColumn = true;
            let maxTextLength = 0;
            
            // Проверяем содержимое столбца
            for (let row = 1; row <= numRows; row++) {
                const cellValue = range.getCell(row, col).getValue();
                
                if (cellValue === null || cellValue === "") continue;
                
                if (typeof cellValue === 'string') {
                    const textLength = cellValue.length;
                    maxTextLength = Math.max(maxTextLength, textLength);
                    
                    // Если строка не число, это текстовый столбец
                    if (isNumericColumn && isNaN(parseFloat(cellValue))) {
                        isNumericColumn = false;
                    }
                }
            }
            
            // Ограничиваем ширину столбца в зависимости от типа данных
            let maxWidth;
            if (isNumericColumn) {
                // Числовым столбцам нужно меньше места
                maxWidth = Math.min(currentWidth, 100);
            } else if (maxTextLength < 20) {
                // Короткие текстовые столбцы
                maxWidth = Math.min(currentWidth, 150);
            } else {
                // Длинные текстовые столбцы - ограничиваем шириной и добавляем перенос
                maxWidth = Math.min(currentWidth, 200);
                range.offset(0, col - 1, numRows, 1).setWrap(true);
            }
            
            // Устанавливаем ограниченную ширину
            sheet.setColumnWidth(actualColumn, maxWidth);
        }
    }
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

    // 1. Список координат и маркеров пустых ячеек
    const emptyCellsToFill = []; // Сохраняем объекты {row: i, col: j, originalIndex: k}
    let emptyCellCounter = 0;
    for (let i = 0; i < numRows; i++) {
        for (let j = 0; j < numCols; j++) {
            if (values[i][j] === "") {
                emptyCellsToFill.push({ row: i, col: j, originalIndex: emptyCellCounter++ });
            }
        }
    }

    if (emptyCellsToFill.length === 0) {
        SpreadsheetApp.getUi().alert('Нет пустых ячеек', 'В выделенном диапазоне нет пустых ячеек.', SpreadsheetApp.getUi().ButtonSet.OK);
        return;
    }

    // 2. Создаём визуальное представление таблицы в виде Markdown-таблицы
    // и пронумеруем пустые ячейки для однозначного порядка
    let markdownTable = "";
    let emptyCellCounterVisual = 1; // Переименовано для избежания конфликта
    
    // Создаём массив для хранения соответствия между пронумерованными пустыми ячейками и их координатами
    const emptyCellsMappings = [];
    
    // Создаём заголовок таблицы с номерами столбцов
    markdownTable += "| |";
    for (let j = 0; j < numCols; j++) {
        markdownTable += ` Колонка ${j+1} |`;
    }
    markdownTable += "\n|--|";
    
    // Добавляем разделитель заголовка
    for (let j = 0; j < numCols; j++) {
        markdownTable += "--|";
    }
    markdownTable += "\n";
    
    // Заполняем таблицу данными с нумерацией пустых ячеек
    for (let i = 0; i < numRows; i++) {
        markdownTable += `| Строка ${i+1} |`;
        for (let j = 0; j < numCols; j++) {
            if (values[i][j] === "") {
                // Пустая ячейка - нумеруем
                markdownTable += ` [ПУСТО${emptyCellCounterVisual}] |`;
                emptyCellsMappings.push({
                    label: `ПУСТО${emptyCellCounterVisual}`,
                    row: i,
                    col: j
                });
                emptyCellCounterVisual++;
            } else {
                // Не пустая ячейка - просто значение
                markdownTable += ` ${values[i][j]} |`;
            }
        }
        markdownTable += "\n";
    }
    
    // Создаём детальные инструкции по каждой пустой ячейке
    let emptyDetailsText = "";
    for (let i = 0; i < emptyCellsMappings.length; i++) {
        const cell = emptyCellsMappings[i];
        emptyDetailsText += `${i+1}. [${cell.label}]: Ячейка на пересечении строки ${cell.row+1}, колонки ${cell.col+1}\n`;
    }
    
    const numberOfEmptyCells = emptyCellsMappings.length;
    
    // 3. Формируем промпт с визуальной таблицей и чёткими инструкциями
    const prompt = `Таблица:
${markdownTable}
Требуется заполнить ${numberOfEmptyCells} пустых ячеек:
${emptyDetailsText}
Верните ТОЛЬКО значения для пустых ячеек через разделитель '¦' в порядке их нумерации от [ПУСТО1] до [ПУСТО${numberOfEmptyCells}].

Формат ответа - одна строка без пробелов вокруг разделителя:
значение1¦значение2¦...¦значение${numberOfEmptyCells}

Если для какой-то ячейки невозможно определить значение, используйте 'n/a'.
Не добавляйте никаких комментариев или пояснений перед или после значений.`;

    logMessage(`fillCells prompt:\n${prompt}`);

    const settings = getSettings();
    const model = settings.model;
    // Устанавливаем меньшее значение temperature для более детерминированных ответов
    const temperature = settings.temperature; // Используем fixed temperature вместо settings.temperature
    const retries = settings.retryAttempts;
    const maxTokens = settings.maxTokens; // Убедимся, что maxTokens достаточен для ответа

    let attempt = 0;
    let success = false;
    let aiResponseText;

    while (attempt < retries && !success) {
        attempt++;
        logMessage(`Попытка ${attempt} из ${retries}`);
        try {
            const aiFullResponse = openRouterRequest(prompt, model, temperature, retries, maxTokens);
            if (!aiFullResponse?.choices?.[0]?.message?.content) {
                throw new Error("Ошибка: Неожиданный или пустой ответ от OpenRouter в fillCells");
            }
            aiResponseText = aiFullResponse.choices[0].message.content.trim();
            logMessage(`fillCells AI response (попытка ${attempt}):\n${aiResponseText}`);

            // Очищаем ответ от лишних кавычек и пробелов
            const cleanedResponse = aiResponseText.replace(/^['"`]|['"`]$/g, '').trim();
            const filledValuesRaw = cleanedResponse.split('¦');

            if (filledValuesRaw.length !== numberOfEmptyCells) {
                throw new Error(`Неверное количество значений в ответе: ожидается ${numberOfEmptyCells}, получено ${filledValuesRaw.length}. Ответ: "${aiResponseText}"`);
            }
            
            const filledValues = filledValuesRaw.map(v => {
                const trimmedValue = v.trim();
                return (trimmedValue === "[Н/Д]" || trimmedValue === "n/a") ? "" : trimmedValue;
            });
            
            // 5. Заполняем пустые ячейки полученными значениями в соответствии с маппингом
            for (let i = 0; i < numberOfEmptyCells; i++) {
                const cellInfo = emptyCellsMappings[i];
                // Проверяем, что значение не undefined
                const valueToSet = filledValues[i] !== undefined ? filledValues[i] : ""; 
                range.getCell(cellInfo.row + 1, cellInfo.col + 1).setValue(valueToSet);
            }
            success = true;

            // Кэширование (если нужно, можно адаптировать или оставить как есть)
            const promptPrefix = `fillEmptyCells_v2_context:${tableContextString.substring(0,100)}`; // Пример префикса для кэша
            const effectiveModel = SCRIPT_PROPERTIES.getProperty(MODEL_SETTING_KEY) || model || DEFAULT_MODEL;
            const effectiveTemperature = parseFloat(SCRIPT_PROPERTIES.getProperty(TEMPERATURE_SETTING_KEY)) || temperature || DEFAULT_TEMPERATURE;
            const effectiveMaxTokens = parseInt(SCRIPT_PROPERTIES.getProperty(MAX_TOKENS_SETTING_KEY), 10) || DEFAULT_MAX_TOKENS;
            const promptHash = calculateMD5(promptPrefix + effectiveModel + effectiveTemperature + effectiveMaxTokens + emptyCellsToFill.map(c => `${c.row},${c.col}`).join(';'));
            const cacheKey = `promptHash:${promptHash}`;
            CACHE.put(cacheKey, JSON.stringify({responseText: aiResponseText}), 21600); // Кэшируем только текст ответа
            logMessage(`fillCells: Ответ AI сохранен в кэше для ключа: ${cacheKey}`);
        } catch (error) {
            logMessage(`Ошибка в fillCells (попытка ${attempt}): ${error.toString()}`, true);
            if (attempt === retries) {
                 SpreadsheetApp.getUi().alert('Ошибка', `Не удалось получить корректный ответ от AI после ${retries} попыток. Последняя ошибка: ${error.message}`);
                return; // Выходим из функции, если все попытки неудачны
            }
        }
    }

    if (!success) {
        // Это сообщение не должно появиться, если return выше сработал
        SpreadsheetApp.getUi().alert('Ошибка', `Не удалось получить и обработать ответ от AI после ${retries} попыток.`);
        return;
    }

    SpreadsheetApp.getUi().alert('Ячейки заполнены', 'Пустые ячейки в выделенном диапазоне были заполнены.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Применяет отдельную инструкцию форматирования.
 * @param {GoogleAppsScript.Spreadsheet.Range} range - Диапазон ячеек таблицы
 * @param {Object} instruction - Инструкция форматирования
 * @param {number} numRows - Количество строк в таблице
 * @param {number} numCols - Количество столбцов в таблице
 */
function applyFormatInstruction(range, instruction, numRows, numCols) {
    // Получаем диапазон ячеек, к которым применить форматирование
    let targetRange;
    
    if (instruction.type === 'header' && instruction.row === '1') {
        // Форматирование строки заголовка
        const row = 1;
        if (instruction.column === 'all') {
            targetRange = range.offset(0, 0, 1, numCols);
        } else {
            const col = parseInt(instruction.column);
            if (isNaN(col) || col < 1 || col > numCols) return;
            targetRange = range.getCell(1, col);
        }
    } else if (instruction.type === 'row') {
        // Форматирование строки
        let row;
        if (instruction.row === 'all') {
            // Применяем ко всем строкам по одной
            for (let r = 1; r <= numRows; r++) {
                const rowRange = range.offset(r - 1, 0, 1, numCols);
                applyFormatting(rowRange, instruction.instructions);
            }
            return;
        } else {
            row = parseInt(instruction.row);
            if (isNaN(row) || row < 1 || row > numRows) return;
        }
        
        if (instruction.column === 'all') {
            targetRange = range.offset(row - 1, 0, 1, numCols);
        } else {
            const col = parseInt(instruction.column);
            if (isNaN(col) || col < 1 || col > numCols) return;
            targetRange = range.getCell(row, col);
        }
    } else if (instruction.type === 'column') {
        // Форматирование столбца
        let col;
        if (instruction.column === 'all') {
            // Применяем ко всем столбцам по одному
            for (let c = 1; c <= numCols; c++) {
                const colRange = range.offset(0, c - 1, numRows, 1);
                applyFormatting(colRange, instruction.instructions);
            }
            return;
        } else {
            col = parseInt(instruction.column);
            if (isNaN(col) || col < 1 || col > numCols) return;
        }
        
        if (instruction.row === 'all') {
            targetRange = range.offset(0, col - 1, numRows, 1);
        } else {
            const row = parseInt(instruction.row);
            if (isNaN(row) || row < 1 || row > numRows) return;
            targetRange = range.getCell(row, col);
        }
    } else if (instruction.type === 'cell') {
        // Форматирование отдельной ячейки
        const row = parseInt(instruction.row);
        const col = parseInt(instruction.column);
        if (isNaN(row) || isNaN(col) || row < 1 || row > numRows || col < 1 || col > numCols) return;
        
        targetRange = range.getCell(row, col);
    } else {
        // Неизвестный тип инструкции
        return;
    }
    
    // Применяем форматирование к целевому диапазону
    if (targetRange) {
        applyFormatting(targetRange, instruction.instructions);
    }
}