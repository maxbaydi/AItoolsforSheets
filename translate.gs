// Константы

// Функции для перевода
function translateToRussian() {
    translateRange('русский', SpreadsheetApp.getActiveRange().getA1Notation());
}

function translateToEnglish() {
    translateRange('английский', SpreadsheetApp.getActiveRange().getA1Notation());
}

function translateToChinese() {
    translateRange('китайский', SpreadsheetApp.getActiveRange().getA1Notation());
}

function translateToSpanish() {
    translateRange('испанский', SpreadsheetApp.getActiveRange().getA1Notation());
}

function translateToFrench() {
    translateRange('французский', SpreadsheetApp.getActiveRange().getA1Notation());
}

/**
 * Функция-обертка для перевода выбранного диапазона на указанный язык
 * @param {string} language - Язык, на который нужно перевести
 * @param {string} rangeStr - A1-нотация диапазона ячеек
 * @returns {string} - Сообщение о результате перевода
 */
function translateRange(language, rangeStr) {
    try {
        // Используем существующую функцию translateRangeWithModel с настройками по умолчанию
        return translateRangeWithModel(language, rangeStr);
    } catch (error) {
        logMessage(`Ошибка в translateRange: ${error.toString()}`, true);
        throw new Error('Ошибка перевода: ' + error.message);
    }
}

// Быстрый перевод через Google Translate (машинный перевод)
function quickTranslateWithGoogle(sourceLang, targetLang) {
  // Преобразуем названия языков в коды
  var sourceCode = getLanguageCode(sourceLang);
  var targetCode = getLanguageCode(targetLang);
  
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getActiveRange();
  var values = range.getValues();
  var formulas = [];
  // Выбираем разделитель аргументов формулы: запятая в англоязычных локалях, иначе точка с запятой
  var locale = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetLocale();
  var sep = locale && locale.toLowerCase().startsWith('en') ? ',' : ';';

  function sanitizeText(text) {
    if (typeof text !== 'string') return '';
    // Удаляем управляющие символы, кроме \n, экранируем кавычки (дублируем их), заменяем переносы строк на пробел
    let sanitized = text.replace(/[\u0000-\u001F\u007F-\u009F]/g, ' ')
      .replace(/"/g, '""')
      .replace(/\r?\n|\r/g, ' ')
      .replace(/[""«»]/g, '"');
    return sanitized;
  }

  for (var i = 0; i < values.length; i++) {
    formulas[i] = [];
    for (var j = 0; j < values[i].length; j++) {
      var cellValue = values[i][j];
      if (typeof cellValue === 'string' && cellValue.trim() !== '') {
        var safeValue = sanitizeText(cellValue);
        formulas[i][j] = '=GOOGLETRANSLATE("' + safeValue + '"' + sep + '"' + sourceCode + '"' + sep + '"' + targetCode + '")';
      } else {
        formulas[i][j] = '';
      }
    }
  }
  range.setFormulas(formulas);
  SpreadsheetApp.flush();
  var translations = range.getValues();
  var finalValues = [];
  for (var i = 0; i < values.length; i++) {
    finalValues[i] = [];
    for (var j = 0; j < values[i].length; j++) {
      // сохраняем переведённый текст для строковых значений, иначе оригинал
      if (typeof values[i][j] === 'string' && values[i][j].trim() !== '') {
        finalValues[i][j] = translations[i][j];
      } else {
        finalValues[i][j] = values[i][j];
      }
    }
  }
  range.setValues(finalValues);
  
  // Возвращаем информацию о выполненном переводе
  return `Перевод выполнен (${sourceLang === 'auto' ? 'Автоопределение' : sourceLang} → ${targetLang})`;
}

function quickTranslateToRussian() {
  quickTranslateWithGoogle('auto', 'ru');
}
function quickTranslateToEnglish() {
  quickTranslateWithGoogle('auto', 'en');
}
function quickTranslateToChinese() {
  quickTranslateWithGoogle('auto', 'zh-CN');
}
function quickTranslateToSpanish() {
  quickTranslateWithGoogle('auto', 'es');
}
function quickTranslateToFrench() {
  quickTranslateWithGoogle('auto', 'fr');
}

// Функция для перевода с выбором модели
function translateRangeWithModel(language, rangeStr, temperature, model) {
    try {
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = spreadsheet.getActiveSheet();
        const range = sheet.getRange(rangeStr);
        const values = range.getValues();
        
        const settings = getSettings();
        const temp = temperature !== undefined ? temperature : settings.temperature;
        const maxRetries = settings.retryAttempts;
        const maxTokens = settings.maxTokens;
        // Используем переданную модель или модель по умолчанию из настроек
        const modelToUse = model || settings.model;
        
        // Проверяем, не является ли диапазон пустым
        if (values.length === 0 || values.every(row => row.every(cell => !cell || cell.toString().trim() === ''))) {
            return "Нет данных для перевода";
        }
        
        // Новый подход: оптимизация обработки диапазона
        const allTranslated = translateOptimized(values, language, modelToUse, temp, maxRetries, maxTokens);
        
        // Вставляем переведенные тексты
        insertOptimizedTranslations(range, allTranslated, values);
        
        return "Перевод выполнен";
    } catch (error) {
        logMessage(`Ошибка в translateRangeWithModel: ${error.toString()}`, true);
        throw new Error('Ошибка перевода: ' + error.message);
    }
}

/**
 * Оптимизированная функция для перевода диапазона данных с дедупликацией
 */
function translateOptimized(values, language, model, temperature, maxRetries, maxTokens) {
    // Анализируем структуру данных для оптимального разделения
    const structuredData = analyzeDataStructure(values);
    
    // Выполняем дедупликацию текстов перед отправкой на перевод
    const { uniqueItems, duplicateMapping } = deduplicateTexts(structuredData);
    
    logMessage(`Всего элементов: ${structuredData.length}, уникальных элементов: ${uniqueItems.length}, повторений: ${structuredData.length - uniqueItems.length}`);
    
    // Кэш переводов для оптимизации повторяющихся значений
    const translationCache = {};
    const result = new Array(structuredData.length);
    
    // Определяем оптимальный размер чанка в зависимости от объема текста
    const chunkOptimizer = new ChunkSizeOptimizer(uniqueItems, maxTokens);
    const chunks = chunkOptimizer.createOptimizedChunks();
    
    logMessage(`Созданы оптимизированные чанки: ${chunks.length} чанков`);
    
    // Обрабатываем каждый чанк
    for (let i = 0; i < chunks.length; i++) {
        const chunk = chunks[i];
        logMessage(`Обработка чанка ${i+1}/${chunks.length}, ${chunk.items.length} элементов`);
        
        // Проверяем кэш для исключения повторного перевода идентичных текстов
        const cacheableItems = [];
        const nonCacheableItems = [];
        let chunkItemsToTranslate = [];
        
        // Разделяем элементы чанка на кэшируемые и не кэшируемые
        chunk.items.forEach(item => {
            const text = item.text.trim();
            // Если текст длинный или содержит структурированные данные, не кэшируем
            const isCacheable = text.length < 500 && !containsStructuredData(text);
            
            if (isCacheable) {
                cacheableItems.push({
                    index: item.index,
                    text: text,
                    isCached: translationCache[text] !== undefined
                });
                
                if (translationCache[text] === undefined) {
                    chunkItemsToTranslate.push(item);
                }
            } else {
                nonCacheableItems.push(item);
                chunkItemsToTranslate.push(item);
            }
        });
        
        // Если все элементы уже в кэше, пропускаем запрос к API
        if (chunkItemsToTranslate.length === 0) {
            logMessage(`Все элементы чанка найдены в кэше, пропускаем API-запрос`);
            
            // Восстанавливаем переводы из кэша
            cacheableItems.forEach(item => {
                result[item.index] = translationCache[item.text];
            });
            
            continue;
        }
        
        // Формируем промпт только для элементов, которые нужно перевести
        const chunkPrompt = buildOptimizedPrompt(language, chunkItemsToTranslate, i, chunks.length);
        logMessage(`Размер промпта: ${chunkPrompt.length} символов`);
        
        // Получаем перевод для чанка
        let translations = [];
        let attempt = 0;
        
        while (attempt < maxRetries && translations.length === 0) {
            attempt++;
            try {
                const response = openRouterRequest(chunkPrompt, model, temperature, 1, maxTokens);
                const answer = extractTranslationsFromResponse(response, chunkItemsToTranslate.length);
                translations = answer;
                
                // Если получили неправильное количество переводов, пробуем еще раз
                if (translations.length !== chunkItemsToTranslate.length) {
                    logMessage(`Получено ${translations.length} переводов, ожидалось ${chunkItemsToTranslate.length}. Повторная попытка...`, true);
                    translations = [];
                    continue;
                }
            } catch (error) {
                logMessage(`Ошибка при переводе чанка (попытка ${attempt}): ${error.message}`, true);
                if (attempt >= maxRetries) {
                    throw new Error(`Не удалось получить перевод после ${maxRetries} попыток: ${error.message}`);
                }
            }
        }
        
        if (translations.length === 0) {
            throw new Error(`Не удалось получить перевод чанка ${i+1} после ${maxRetries} попыток`);
        }
        
        // Обновляем кэш и результаты
        chunkItemsToTranslate.forEach((item, idx) => {
            const text = item.text.trim();
            const translation = translations[idx];
            
            result[item.index] = translation;
            
            // Кэшируем только небольшие тексты без форматирования
            if (text.length < 500 && !containsStructuredData(text)) {
                translationCache[text] = translation;
            }
        });
        
        // Добавляем кэшированные элементы в результаты
        cacheableItems.filter(item => item.isCached).forEach(item => {
            result[item.index] = translationCache[item.text];
        });
        
        // Добавляем задержку между чанками, чтобы избежать ограничений API
        if (i < chunks.length - 1) {
            Utilities.sleep(500);
        }
    }
    
    // Применяем переводы к дубликатам
    for (const [sourceIndex, duplicateIndices] of Object.entries(duplicateMapping)) {
        const translatedText = result[parseInt(sourceIndex)];
        for (const duplicateIndex of duplicateIndices) {
            result[duplicateIndex] = translatedText;
        }
    }
    
    return result.filter(item => item !== undefined);
}

/**
 * Выполняет дедупликацию текстов, выявляет уникальные тексты и создает отображение дубликатов
 * @param {Array} dataItems - Массив элементов данных
 * @returns {Object} Объект с уникальными элементами и отображением дубликатов
 */
function deduplicateTexts(dataItems) {
    const textMap = new Map(); // Для отслеживания уже встреченных текстов
    const uniqueItems = []; // Список уникальных элементов
    const duplicateMapping = {}; // Отображение исходных элементов на их дубликаты
    
    // Первый проход: выявляем уникальные элементы и строим отображение дубликатов
    for (let i = 0; i < dataItems.length; i++) {
        const item = dataItems[i];
        const normalizedText = normalizeTextForDeduplication(item.text);
        
        // Если это большой или структурированный текст, не дедуплицируем его
        if (item.length > 3000 || containsStructuredData(item.text)) {
            uniqueItems.push(item);
            continue;
        }
        
        if (textMap.has(normalizedText)) {
            // Это дубликат, получаем индекс исходного элемента
            const sourceIndex = textMap.get(normalizedText);
            
            // Добавляем текущий индекс в список дубликатов для исходного элемента
            if (!duplicateMapping[sourceIndex]) {
                duplicateMapping[sourceIndex] = [];
            }
            duplicateMapping[sourceIndex].push(item.index);
        } else {
            // Это уникальный элемент
            textMap.set(normalizedText, item.index);
            uniqueItems.push(item);
        }
    }
    
    // Сортируем уникальные элементы для оптимального разделения на чанки
    uniqueItems.sort((a, b) => {
        // Сначала по длине (от длинных к коротким)
        const lengthDiff = b.length - a.length;
        if (Math.abs(lengthDiff) > 500) return lengthDiff;
        
        // Затем по сложности
        return b.complexity - a.complexity;
    });
    
    return { uniqueItems, duplicateMapping };
}

/**
 * Нормализует текст для дедупликации, чтобы выявить текстовые дубликаты
 * @param {string} text - Исходный текст
 * @returns {string} Нормализованный текст
 */
function normalizeTextForDeduplication(text) {
    // Убираем лишние пробелы и приводим к нижнему регистру
    return text.trim()
               .toLowerCase()
               .replace(/\s+/g, ' ') // Заменяем множественные пробелы на один
               .replace(/\n+/g, '\n') // Заменяем множественные переносы строк на один
               .replace(/\t+/g, '\t'); // Заменяем множественные табуляции на одну
}

/**
 * Анализирует структуру данных для оптимального разделения на чанки
 */
function analyzeDataStructure(values) {
    const flatData = [];
    let index = 0;
    
    // Проходим по всем строкам и столбцам
    for (let i = 0; i < values.length; i++) {
        for (let j = 0; j < values[i].length; j++) {
            const cellValue = values[i][j];
            if (cellValue !== null && cellValue !== undefined && String(cellValue).trim() !== '') {
                const text = String(cellValue);
                flatData.push({
                    index: index++,
                    row: i,
                    col: j,
                    text: text,
                    length: text.length,
                    isMultiline: text.includes('\n'),
                    lineCount: text.split('\n').length,
                    complexity: assessTextComplexity(text)
                });
            }
        }
    }
    
    return flatData;
}

/**
 * Создает оптимизированный промпт для перевода чанка
 */
function buildOptimizedPrompt(language, items, chunkIndex, totalChunks) {
    // Более строгий промпт для обеспечения правильного формата ответа
    let prompt = `Переведи следующие ${items.length} тексты на ${language} язык.\n\n`;
    
    // Подробные инструкции для модели
    prompt += `ВАЖНЫЕ ИНСТРУКЦИИ:\n`;
    prompt += `1. НЕ переводи имена собственные, цифры, коды, email адреса.\n`;
    prompt += `2. Сохраняй форматирование и структуру текста.\n`;
    prompt += `3. ОБЯЗАТЕЛЬНО используй символ '${COLUMN_DELIMITER}' как разделитель между переводами.\n`;
    prompt += `4. Верни ровно ${items.length} переводов, разделенных символом '${COLUMN_DELIMITER}'.\n`;
    prompt += `5. Не добавляй номера, пояснения или дополнительный текст.\n`;
    prompt += `6. Каждый перевод должен идти строго в том же порядке, что и исходные тексты.\n`;
    prompt += `7. Маркер <n> в тексте замени на перенос строки в переводе.\n`;
    prompt += `8. Убедись, что разделитель '${COLUMN_DELIMITER}' используется только между переводами.\n\n`;
    
    // Добавляем более наглядный пример с реальным разделителем
    prompt += `ФОРМАТ ОТВЕТА:\n`;
    if (items.length === 1) {
        prompt += `перевод текста 1\n\n`;
    } else if (items.length === 2) {
        prompt += `перевод текста 1${COLUMN_DELIMITER}перевод текста 2\n\n`;
    } else {
        prompt += `перевод текста 1${COLUMN_DELIMITER}перевод текста 2${COLUMN_DELIMITER}перевод текста 3${COLUMN_DELIMITER}...\n\n`;
    }
    
    // Добавляем исходные тексты
    prompt += `ТЕКСТЫ ДЛЯ ПЕРЕВОДА:\n`;
    
    items.forEach((item, index) => {
        prompt += `ТЕКСТ ${index + 1}:\n${item.text.replace(/\n/g, " <n> ")}\n\n`;
    });
    
    // Дополнительное напоминание в конце с указанием конкретного количества переводов
    prompt += `ВАЖНО: Верни ${items.length} переводов`;
    if (items.length > 1) {
        prompt += `, разделённых символом '${COLUMN_DELIMITER}'`;
    }
    prompt += `. Не добавляй номера текстов или другие пояснения. Просто переводы, `;
    if (items.length > 1) {
        prompt += `разделенные символом '${COLUMN_DELIMITER}'. `;
    }
    prompt += `Маркер <n> замени на перенос строки.`;
    
    return prompt;
}

/**
 * Извлекает переводы из ответа модели
 */
function extractTranslationsFromResponse(response, expectedCount) {
    if (!response?.choices?.[0]?.message) {
        throw new Error("Неожиданный формат ответа от модели");
    }
    
    const answer = response.choices[0].message.content.trim();
    logMessage(`Получен ответ от модели: ${answer.substring(0, 100)}...`, false, 3);
    logMessage(`Полный ответ модели: ${answer}`, false, 5);
    
    // Убираем лишние разделители в начале и конце
    const cleanAnswer = answer.replace(/^¦+/, '').replace(/¦+$/, '');
    
    // Разделяем по основному разделителю
    let translations = cleanAnswer.split(COLUMN_DELIMITER);
    
    // Проверяем количество переводов
    if (translations.length !== expectedCount) {
        logMessage(`Предупреждение: получено ${translations.length} переводов, ожидалось ${expectedCount}`, true, 1);
        logMessage(`Разделитель '${COLUMN_DELIMITER}' найден ${cleanAnswer.split(COLUMN_DELIMITER).length - 1} раз`, false, 2);
        
        // Пытаемся восстановить структуру, если модель не использовала разделители
        if (translations.length === 1 && expectedCount > 1) {
            logMessage(`Пытаемся восстановить структуру перевода из текста ответа...`, false, 2);
            
            // Пробуем найти маркеры "ТЕКСТ N:" или похожие паттерны
            const textMarkers = cleanAnswer.match(/ТЕКСТ\s+\d+:|TEXT\s+\d+:|Текст\s+\d+:|текст\s+\d+:/g);
            
            if (textMarkers && textMarkers.length >= expectedCount) {
                logMessage(`Найдены текстовые маркеры: ${textMarkers.join(', ')}`, false, 2);
                
                const parts = [];
                const indices = [];
                
                // Находим индексы всех маркеров
                textMarkers.forEach(marker => {
                    const index = cleanAnswer.indexOf(marker);
                    if (index !== -1) {
                        indices.push(index);
                    }
                });
                
                // Сортируем индексы
                indices.sort((a, b) => a - b);
                
                // Разбиваем текст по индексам
                for (let i = 0; i < indices.length; i++) {
                    const start = indices[i];
                    const end = i < indices.length - 1 ? indices[i + 1] : cleanAnswer.length;
                    
                    // Находим конец маркера
                    const markerEnd = cleanAnswer.indexOf(':', start) + 1;
                    // Извлекаем текст без маркера
                    const part = cleanAnswer.substring(markerEnd, end).trim();
                    parts.push(part);
                    logMessage(`Извлечена часть ${i+1}: ${part.substring(0, 30)}...`, false, 4);
                }
                
                if (parts.length === expectedCount) {
                    logMessage(`Успешно восстановлено ${parts.length} частей текста по маркерам`, false, 2);
                    return parts.map(part => part.replace(/<n>/g, '\n'));
                } else {
                    logMessage(`Не удалось восстановить ожидаемое количество частей: получено ${parts.length}, ожидалось ${expectedCount}`, true, 2);
                }
            } else {
                logMessage(`Не найдены подходящие текстовые маркеры или их недостаточно`, false, 2);
            }
            
            // Пробуем разбить по числовым маркерам
            const numberMarkers = cleanAnswer.match(/^\d+\.\s/gm);
            if (numberMarkers && numberMarkers.length >= expectedCount) {
                logMessage(`Найдены числовые маркеры: ${numberMarkers.join(', ')}`, false, 2);
                
                const parts = cleanAnswer.split(/^\d+\.\s/m).filter(part => part.trim() !== '');
                if (parts.length === expectedCount) {
                    logMessage(`Успешно восстановлено ${parts.length} частей текста по числовым маркерам`, false, 2);
                    return parts.map(part => part.trim().replace(/<n>/g, '\n'));
                } else {
                    logMessage(`Не удалось восстановить ожидаемое количество частей по числовым маркерам: получено ${parts.length}, ожидалось ${expectedCount}`, false, 2);
                }
            } else {
                logMessage(`Не найдены подходящие числовые маркеры или их недостаточно`, false, 2);
            }
            
            // Пробуем разбить по двойным переносам строк
            if (cleanAnswer.includes('\n\n')) {
                logMessage(`Пытаемся разбить по двойным переносам строк`, false, 2);
                
                const parts = cleanAnswer.split(/\n\n+/).filter(part => part.trim() !== '');
                if (parts.length === expectedCount) {
                    logMessage(`Успешно восстановлено ${parts.length} частей текста по двойным переносам строк`, false, 2);
                    return parts.map(part => part.trim().replace(/<n>/g, '\n'));
                } else {
                    logMessage(`Не удалось восстановить ожидаемое количество частей по двойным переносам: получено ${parts.length}, ожидалось ${expectedCount}`, false, 2);
                }
            } else {
                logMessage(`Не найдены двойные переносы строк в ответе`, false, 2);
            }

            // В крайнем случае, пробуем просто разбить на равные части
            if (expectedCount > 1) {
                logMessage(`Попытка разбить ответ на ${expectedCount} равных частей`, false, 2);
                const roughLength = Math.floor(cleanAnswer.length / expectedCount);
                const parts = [];
                
                for (let i = 0; i < expectedCount; i++) {
                    const start = i * roughLength;
                    const end = (i === expectedCount - 1) ? cleanAnswer.length : (i + 1) * roughLength;
                    parts.push(cleanAnswer.substring(start, end).trim());
                }
                
                logMessage(`Разбито на ${parts.length} частей равной длины (примерно ${roughLength} символов каждая)`, false, 2);
                return parts.map(part => part.replace(/<n>/g, '\n'));
            }
        }
    }
    
    // Приводим все переводы к стандартному виду и заменяем маркеры <n> на переносы строк
    translations = translations.map(t => t.trim().replace(/<n>/g, '\n'));
    
    // Если всё ещё не совпадает количество, дополняем или обрезаем
    while (translations.length < expectedCount) {
        translations.push("");
        logMessage(`Добавлен пустой перевод для соответствия ожидаемому количеству`, false, 2);
    }
    
    if (translations.length > expectedCount) {
        translations = translations.slice(0, expectedCount);
        logMessage(`Переводы обрезаны до ожидаемого количества ${expectedCount}`, false, 2);
    }
    
    return translations;
}

/**
 * Оценивает сложность текста для оптимизации размера чанка
 */
function assessTextComplexity(text) {
    let complexity = 1.0;
    
    // Длинные тексты сложнее
    if (text.length > 1000) complexity *= 1.5;
    if (text.length > 3000) complexity *= 1.5;
    
    // Тексты с форматированием сложнее
    if (text.includes('\n')) complexity *= 1.2;
    if (text.split('\n').length > 10) complexity *= 1.3;
    
    // Тексты с техническими терминами сложнее
    if (/[A-Z]{2,}/.test(text)) complexity *= 1.1;
    if (/\d{5,}/.test(text)) complexity *= 1.1;
    
    // Тексты с таблицами или структурированными данными сложнее
    if (containsStructuredData(text)) complexity *= 1.5;
    
    return complexity;
}

/**
 * Определяет, содержит ли текст структурированные данные
 */
function containsStructuredData(text) {
    // Простые эвристики для определения таблиц и структурированных данных
    const hasTable = text.split('\n').some(line => line.includes('\t') || (line.match(/\s{2,}/g)?.length > 2));
    const hasForm = text.split('\n').filter(line => line.includes(':') || line.match(/^[\w\s]+\s{2,}/)).length > 3;
    
    return hasTable || hasForm;
}

/**
 * Оптимизатор размера чанков на основе оценки сложности текста
 */
class ChunkSizeOptimizer {
    constructor(dataItems, maxTokens = 4000) {
        this.dataItems = dataItems;
        this.maxTokens = maxTokens;
        this.baseChunkSize = 20; // Увеличено с 10 до 20
        this.maxChunkSize = 100;  // Увеличено с 50 до 100
    }
    
    // Примерная оценка количества токенов в тексте
    estimateTokens(text) {
        const words = text.split(/\s+/).length;
        // ~1.5 токена на слово - это примерное соотношение
        return Math.ceil(words * 1.5); 
    }
    
    // Создает оптимизированные чанки
    createOptimizedChunks() {
        const chunks = [];
        let currentChunk = { items: [], tokenEstimate: 0, complexitySum: 0 };
        
        // Константы для промпта
        const promptOverheadTokens = 200; // Примерная оценка для заголовка промпта
        
        // Сортируем элементы - длинные и сложные тексты идут отдельно
        const sortedItems = [...this.dataItems].sort((a, b) => {
            // Очень длинные тексты всегда идут вначале отдельно
            if (a.length > 2000 && b.length <= 2000) return -1;
            if (b.length > 2000 && a.length <= 2000) return 1;
            
            // Затем идут тексты со сложным форматированием
            if (a.isMultiline && !b.isMultiline) return -1;
            if (b.isMultiline && !a.isMultiline) return 1;
            
            // Остальные сортируем по размеру
            return b.length - a.length;
        });
        
        for (const item of sortedItems) {
            const itemTokens = this.estimateTokens(item.text);
            
            // Определяем, нужно ли создать новый чанк
            const wouldExceedTokenLimit = 
                (currentChunk.tokenEstimate + itemTokens + promptOverheadTokens > this.maxTokens) || 
                (currentChunk.items.length >= this.getOptimalChunkSize(item.complexity));
            
            // Очень большие тексты всегда идут в отдельный чанк
            const isVeryLarge = itemTokens > this.maxTokens * 0.7;
            
            if ((wouldExceedTokenLimit || isVeryLarge) && currentChunk.items.length > 0) {
                chunks.push(currentChunk);
                currentChunk = { items: [], tokenEstimate: 0, complexitySum: 0 };
            }
            
            // Если текст слишком большой для одного запроса, разбиваем его
            if (itemTokens > this.maxTokens - promptOverheadTokens) {
                const splitItems = this.splitLargeItem(item);
                
                for (const splitItem of splitItems) {
                    const splitItemTokens = this.estimateTokens(splitItem.text);
                    
                    if (currentChunk.items.length > 0 && 
                        currentChunk.tokenEstimate + splitItemTokens + promptOverheadTokens > this.maxTokens) {
                        chunks.push(currentChunk);
                        currentChunk = { items: [], tokenEstimate: 0, complexitySum: 0 };
                    }
                    
                    currentChunk.items.push(splitItem);
                    currentChunk.tokenEstimate += splitItemTokens;
                    currentChunk.complexitySum += splitItem.complexity;
                }
            } else {
                currentChunk.items.push(item);
                currentChunk.tokenEstimate += itemTokens;
                currentChunk.complexitySum += item.complexity;
            }
        }
        
        // Добавляем последний чанк, если в нем есть элементы
        if (currentChunk.items.length > 0) {
            chunks.push(currentChunk);
        }
        
        return chunks;
    }
    
    // Определяет оптимальный размер чанка в зависимости от сложности текста
    getOptimalChunkSize(complexity) {
        // Для более сложных текстов уменьшаем размер чанка
        return Math.min(
            this.maxChunkSize,
            Math.ceil(this.baseChunkSize / Math.sqrt(complexity))
        );
    }
    
    // Разбивает большой элемент на несколько меньших
    splitLargeItem(item) {
        const text = item.text;
        
        // Если текст многострочный, разбиваем по логическим блокам
        if (item.isMultiline) {
            const lines = text.split('\n');
            const paragraphs = [];
            let currentParagraph = [];
            
            for (const line of lines) {
                currentParagraph.push(line);
                
                // Пустая строка обычно означает конец параграфа
                if (line.trim() === '') {
                    if (currentParagraph.length > 0) {
                        paragraphs.push(currentParagraph.join('\n'));
                        currentParagraph = [];
                    }
                }
            }
            
            // Добавляем последний параграф, если он не пустой
            if (currentParagraph.length > 0) {
                paragraphs.push(currentParagraph.join('\n'));
            }
            
            // Если у нас получились параграфы, создаем элементы на их основе
            if (paragraphs.length > 1) {
                return paragraphs.map((paragraph, idx) => ({
                    index: item.index,
                    text: paragraph,
                    isPartOfLarger: true,
                    partIndex: idx,
                    totalParts: paragraphs.length,
                    complexity: item.complexity * 0.8, // Уменьшаем сложность для частей
                    originalItem: item
                }));
            }
        }
        
        // Если не получилось разбить по параграфам, делим по размеру
        const chunkCount = Math.ceil(this.estimateTokens(text) / (this.maxTokens * 0.7));
        const chunks = [];
        
        // Примерное количество символов в каждом чанке
        const chunkSize = Math.ceil(text.length / chunkCount);
        
        for (let i = 0; i < chunkCount; i++) {
            const start = i * chunkSize;
            const end = Math.min((i + 1) * chunkSize, text.length);
            
            // Немного корректируем границы, чтобы не разрывать слова
            let adjustedStart = start;
            let adjustedEnd = end;
            
            if (i > 0 && start < text.length) {
                // Ищем начало слова
                while (adjustedStart < text.length && !/\s/.test(text[adjustedStart - 1])) {
                    adjustedStart++;
                }
            }
            
            if (i < chunkCount - 1 && end < text.length) {
                // Ищем конец слова
                while (adjustedEnd > 0 && !/\s/.test(text[adjustedEnd])) {
                    adjustedEnd--;
                }
                // Если не нашли пробел, просто используем оригинальную границу
                if (adjustedEnd <= adjustedStart) {
                    adjustedEnd = end;
                }
            }
            
            chunks.push({
                index: item.index,
                text: text.substring(adjustedStart, adjustedEnd),
                isPartOfLarger: true,
                partIndex: i,
                totalParts: chunkCount,
                complexity: item.complexity * 0.8,
                originalItem: item
            });
        }
        
        return chunks;
    }
}

/**
 * Вставляет оптимизированные переводы в диапазон
 */
function insertOptimizedTranslations(range, translations, originalValues) {
    // Получаем текущие значения, чтобы не перезаписывать ячейки, которые не нужно было переводить
    const currentValues = range.getValues();
    let translationIndex = 0;

    for (let i = 0; i < originalValues.length; i++) {
        for (let j = 0; j < originalValues[i].length; j++) {
            // Проверяем, была ли эта ячейка в исходных данных для перевода
            if (originalValues[i][j] !== null &&
                originalValues[i][j] !== undefined &&
                String(originalValues[i][j]).trim() !== '') {

                if (translationIndex < translations.length) {
                    // Обновляем значение в массиве currentValues
                    currentValues[i][j] = translations[translationIndex];
                }
                translationIndex++;
            }
        }
    }
    // Записываем обновленный массив одним вызовом
    range.setValues(currentValues);
}

function isTranslationSameAsOriginal(translatedTexts, originalValues) {
    let allMatch = true;
    let translatedIndex = 0;

    for (let i = 0; i < originalValues.length; i++) {
        for (let j = 0; j < originalValues[i].length; j++) {
            const original = (originalValues[i][j] === null || originalValues[i][j] === undefined)
                ? ""
                : String(originalValues[i][j]).trim();

            const translated = (translatedTexts[translatedIndex] === undefined)
                ? ""
                : translatedTexts[translatedIndex].trim();

            if (original.toLowerCase() !== translated.toLowerCase()) {
                allMatch = false;
            }

            translatedIndex++;
        }
    }
    return allMatch;
}

// Функция для преобразования названий языков в коды ISO
function getLanguageCode(languageName) {
  if (!languageName) return 'auto';
  
  // Если передан уже код языка, возвращаем его
  if (languageName.length <= 5 && /^[a-z]{2}(-[A-Z]{2})?$/.test(languageName)) {
    return languageName;
  }
  
  // Нормализуем название языка (убираем лишние пробелы, приводим к нижнему регистру)
  const normalizedName = languageName.toString().toLowerCase().trim();
  
  // Словарь соответствия названий языков и их кодов
  const languageMap = {
    // Русские названия
    'русский': 'ru',
    'английский': 'en',
    'китайский': 'zh-CN',
    'испанский': 'es',
    'французский': 'fr',
    'немецкий': 'de',
    'итальянский': 'it',
    'португальский': 'pt',
    'арабский': 'ar',
    'японский': 'ja',
    'корейский': 'ko',
    'турецкий': 'tr',
    'голландский': 'nl',
    'нидерландский': 'nl',
    'греческий': 'el',
    'польский': 'pl',
    'украинский': 'uk',
    'белорусский': 'be',
    'финский': 'fi',
    'шведский': 'sv',
    'датский': 'da',
    'норвежский': 'no',
    'чешский': 'cs',
    'венгерский': 'hu',
    'вьетнамский': 'vi',
    'тайский': 'th',
    'иврит': 'he',
    'хинди': 'hi',
    'бенгальский': 'bn',
    'тамильский': 'ta',
    
    // Английские названия
    'russian': 'ru',
    'english': 'en',
    'chinese': 'zh-CN',
    'spanish': 'es',
    'french': 'fr',
    'german': 'de',
    'italian': 'it',
    'portuguese': 'pt',
    'arabic': 'ar',
    'japanese': 'ja',
    'korean': 'ko',
    'turkish': 'tr',
    'dutch': 'nl',
    'greek': 'el',
    'polish': 'pl',
    'ukrainian': 'uk',
    'belarusian': 'be',
    'finnish': 'fi',
    'swedish': 'sv',
    'danish': 'da',
    'norwegian': 'no',
    'czech': 'cs',
    'hungarian': 'hu',
    'vietnamese': 'vi',
    'thai': 'th',
    'hebrew': 'he',
    'hindi': 'hi',
    'bengali': 'bn',
    'tamil': 'ta',
    
    // Особые случаи
    'авто': 'auto',
    'auto': 'auto',
    'автоматически': 'auto',
    'определить автоматически': 'auto'
  };
  
  // Возвращаем код языка или 'auto' если не найден
  return languageMap[normalizedName] || 'auto';
}