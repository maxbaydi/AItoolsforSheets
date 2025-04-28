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

// Быстрый перевод через Google Translate (машинный перевод)
function quickTranslateWithGoogle(sourceLang, targetLang) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getActiveRange();
  var values = range.getValues();
  var formulas = [];
  // Выбираем разделитель аргументов формулы: запятая в англоязычных локалях, иначе точка с запятой
  var locale = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetLocale();
  var sep = locale && locale.toLowerCase().startsWith('en') ? ',' : ';';

  function sanitizeText(text) {
    if (typeof text !== 'string') return '';
    // Удаляем управляющие символы, кроме \n, экранируем кавычки (дублируем их), заменяем переносы строк на пробелы
    let sanitized = text.replace(/[\u0000-\u001F\u007F-\u009F]/g, ' ')
      .replace(/"/g, '""')
      .replace(/\r?\n|\r/g, ' ')
      .replace(/[“”«»]/g, '"');
    return sanitized;
  }

  for (var i = 0; i < values.length; i++) {
    formulas[i] = [];
    for (var j = 0; j < values[i].length; j++) {
      var cellValue = values[i][j];
      if (typeof cellValue === 'string' && cellValue.trim() !== '') {
        var safeValue = sanitizeText(cellValue);
        formulas[i][j] = '=GOOGLETRANSLATE("' + safeValue + '"' + sep + '"' + sourceLang + '"' + sep + '"' + targetLang + '")';
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

function translateRange(language, rangeStr, temperature) {
    try {
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = spreadsheet.getActiveSheet();
        const range = sheet.getRange(rangeStr);
        const values = range.getValues();
        
        const settings = getSettings();
        const model = settings.model;
        const temp = temperature !== undefined ? temperature : settings.temperature;
        const maxRetries = settings.retryAttempts;
        const maxTokens = settings.maxTokens;
        const chunkSize = settings.chunkSize || 50; // максимальное количество фрагментов за один запрос

        // Флэттеним все значения и готовим чанки
        const flatOriginal = values.reduce((acc, row) => acc.concat(row), []);
        let allTranslated = [];
        for (let start = 0; start < flatOriginal.length; start += chunkSize) {
            const chunk = flatOriginal.slice(start, start + chunkSize);
            // представляем чанк как 2D-массив для buildTranslatePrompt
            const pseudoValues = chunk.map(item => [item]);
            const prompt = buildTranslatePrompt(language, pseudoValues);
            logMessage(`translateRange chunk prompt: ${prompt}`);
            const translatedChunk = getTranslatedTexts(prompt, model, temp, maxRetries, maxTokens, pseudoValues);
            allTranslated = allTranslated.concat(translatedChunk);
        }

        // Вставляем все переведенные фрагменты по оригинальной структуре
        insertTranslatedTexts(range, allTranslated, values);
        return "Перевод выполнен";
    } catch (error) {
        logMessage(`Ошибка в translateRange: ${error.toString()}`, true);
        throw new Error('Ошибка перевода: ' + error.message);
    }
}

function buildTranslatePrompt(language, values) {
    let prompt = `Ты - бот-переводчик. Твоя задача - сделать перевод текста на ${language} язык. `;
    prompt += `Не переводи имена собственные.`;
    prompt += `Верни ТОЛЬКО переведенные фрагменты, разделенные символом '${COLUMN_DELIMITER}', без пояснений, без примеров, без дополнительных фраз. `;
    prompt += `Сохраняй порядок фрагментов. Если фрагмент пустой - верни пустую строку для него.`;

    for (let i = 0; i < values.length; i++) {
        for (let j = 0; j < values[i].length; j++) {
            prompt += `\nФрагмент ${i * values[i].length + j + 1}: ${values[i][j]}`;
        }
    }
    return prompt;
}

function getTranslatedTexts(prompt, model, temperature, maxRetries, maxTokens, originalValues) {
    let translatedTexts = [];
    let attempt = 0;

    while (attempt < maxRetries && translatedTexts.length === 0) {
        attempt++;
        logMessage(`Попытка перевода №${attempt}`);

        const jsonResponse = openRouterRequest(prompt, model, temperature, maxRetries, maxTokens);
        if (!jsonResponse?.choices?.[0]?.message) {
            const errorMessage = "Ошибка: Неожиданный ответ от OpenRouter при переводе.";
            logMessage(errorMessage, true);
            throw new Error(errorMessage);
        }

        let answer = jsonResponse.choices[0].message.content.trim();
        answer = answer.replace(/^¦+/, '').replace(/¦+$/, '');
        logMessage(`translateRange answer: ${answer}`);
        translatedTexts = answer.split(COLUMN_DELIMITER).map(t => t.trim());

        if (isTranslationSameAsOriginal(translatedTexts, originalValues)) {
            logMessage(`Перевод совпадает с оригиналом (попытка ${attempt}). Повторяем запрос...`);
            translatedTexts = [];
        }
    }

    if (translatedTexts.length === 0) {
        throw new Error("Не удалось получить перевод после нескольких попыток.");
    }

    return translatedTexts;
}

function insertTranslatedTexts(range, translatedTexts, originalValues) {
    let textIndex = 0;
    for (let i = 0; i < originalValues.length; i++) {
        for (let j = 0; j < originalValues[i].length; j++) {
            if (translatedTexts[textIndex] !== undefined) {
                range.getCell(i + 1, j + 1).setValue(translatedTexts[textIndex]);
            }
            textIndex++;
        }
    }
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