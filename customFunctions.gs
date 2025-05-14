/**
 * @OnlyCurrentDoc
 *
 * Генерирует SEO-описание для товара с использованием AI.
 *
 * @param {string|string[][]} inputRange Диапазон ячеек или одна ячейка с данными о товаре.
 * @param {number} [formatType=1] Тип форматирования: 1 - без разметки, 2 - HTML, 3 - Markdown.
 * @return {string} Сгенерированное описание товара.
 * @customfunction
 */
function AIDISC(inputRange, formatType) {
  try {
    let name = "";
    let specs = "";
    // const hints = ""; // Пока не используется, можно будет добавить позже

    if (typeof inputRange === 'string') {
      // Если передана одна ячейка как строка
      name = inputRange;
    } else if (Array.isArray(inputRange)) {
      // Если передан диапазон
      if (inputRange.length > 0 && Array.isArray(inputRange[0])) {
        // Обрабатываем только первую строку диапазона, если их несколько
        const firstRow = inputRange[0];
        if (firstRow.length > 0) {
          name = String(firstRow[0]);
        }
        if (firstRow.length > 1) {
          specs = firstRow.slice(1).filter(String).join(", "); // Объединяем остальные ячейки как характеристики
        }
      } else if (inputRange.length > 0) {
         // Если это одномерный массив (например, результат другой функции)
         name = String(inputRange[0]);
         if (inputRange.length > 1) {
           specs = inputRange.slice(1).filter(String).join(", ");
         }
      }
    } else {
      return "Ошибка: Неверный формат входных данных. Ожидается ячейка или диапазон.";
    }

    if (!name) {
      return "Ошибка: Название товара не указано.";
    }

    let formattingInstructions = "Не используй HTML и Markdown разметку, такую как звездочки (**) для выделения текста. Не добавляй никаких ссылок или упоминаний о других товарах. Не используй эмодзи.";
    const userFormatType = parseInt(formatType, 10);

    if (userFormatType === 2) {
      formattingInstructions = "Используй HTML-теги для форматирования текста (например, <br> для переноса строки, <b> для жирного шрифта, <i> для курсива, <ul> и <li> для списков). Не обрамляй ответ в общие HTML теги, такие как <html> или <body>. Верни только фрагмент HTML, представляющий описание. Не используй эмодзи.";
    } else if (userFormatType === 3) {
      formattingInstructions = "Используй Markdown разметку для форматирования текста (например, переносы строк, **жирный шрифт**, *курсив*, - для списков). Не используй эмодзи.";
    }

    const promptTemplate = `Сгенерируйте SEO-описание для товара "{name}" со следующими характеристиками:
{specs}

Особенности для выделения:
{hints}

Формат:
- Не более 750 слов
- Маркированный список преимуществ
- Технические детали
- Делай краткий вывод в конце
- Используй ключевые слова, связанные с товаром

${formattingInstructions}

Важно: Верни только сгенерированное описание товара без каких-либо дополнительных фраз или пояснений.`;

    let prompt = promptTemplate.replace("{name}", name);
    prompt = prompt.replace("{specs}", specs || "нет дополнительных характеристик");
    prompt = prompt.replace("{hints}", ""); // hints пока не используем

    logMessage("AIDISC: Сформирован промпт: " + prompt);

    const settings = getSettings();
    const model = settings.model;
    const temperature = settings.temperature;
    const maxTokens = settings.maxTokens || 2000; // Установим лимит для описания, если не задан глобально
    const retries = settings.retryAttempts || 3;

    // Кэширование
    const cacheKey = "AIDISC_" + calculateMD5(prompt + model + temperature + maxTokens + (userFormatType || 1));
    const cachedResult = CACHE.get(cacheKey);
    if (cachedResult) {
      logMessage("AIDISC: Результат из кэша для ключа " + cacheKey);
      return cachedResult;
    }

    logMessage("AIDISC: Запрос к API. Модель: " + model + ", Температура: " + temperature + ", Макс.токены: " + maxTokens);
    const aiResponse = openRouterRequest(prompt, model, temperature, retries, maxTokens);

    if (aiResponse && aiResponse.choices && aiResponse.choices[0] && aiResponse.choices[0].message && aiResponse.choices[0].message.content) {
      const description = aiResponse.choices[0].message.content.trim();
      CACHE.put(cacheKey, description, 21600); // Кэшируем на 6 часов
      logMessage("AIDISC: Ответ от AI: " + description);
      return description;
    } else {
      logMessage("AIDISC: Ошибка от API или неверный формат ответа: " + JSON.stringify(aiResponse), true);
      return "Ошибка: Не удалось получить описание от AI.";
    }

  } catch (e) {
    logMessage("AIDISC: КРИТИЧЕСКАЯ ОШИБКА: " + e.toString() + " " + e.stack, true);
    return "Ошибка: " + e.message;
  }
}

/**
 * @OnlyCurrentDoc
 *
 * Отправляет простой запрос к AI.
 *
 * @param {string} promptTemplate Строка запроса. Может содержать плейсхолдер {data} для вставки данных из диапазона.
 * @param {string|string[][]} [dataRange] Диапазон ячеек или одна ячейка, данные из которой будут вставлены вместо {data} в запросе.
 * @param {number} [temperatureValue] Температура для генерации ответа (например, 0.7). Если не указана, используется значение из настроек.
 * @return {string} Ответ от AI.
 * @customfunction
 */
function AIQ(promptTemplate, dataRange, temperatureValue) {
  try {
    let processedPrompt = String(promptTemplate);
    let dataText = "";

    if (dataRange !== undefined && dataRange !== null && dataRange !== "") {
      if (typeof dataRange === 'string') {
        dataText = dataRange;
      } else if (Array.isArray(dataRange)) {
        if (dataRange.length > 0 && Array.isArray(dataRange[0])) { // 2D array
          dataText = dataRange.map(row => row.filter(String).join(" ")).filter(s => s.length > 0).join("\n");
        } else if (dataRange.length > 0) { // 1D array
          dataText = dataRange.filter(String).join(" ");
        }
      } else {
        dataText = String(dataRange);
      }
      processedPrompt = processedPrompt.replace(/{data}/g, dataText);
    } else {
      // Если dataRange не предоставлен, но {data} есть в шаблоне, удалить его или заменить на пустую строку
      processedPrompt = processedPrompt.replace(/{data}/g, "");
    }

    // Базовые инструкции для AI
    const baseInstructions = "Ответь только на поставленный вопрос или выполни инструкцию. Не добавляй никаких вступлений, объяснений, извинений или дополнительного текста, кроме самого ответа. Не используй Markdown или HTML разметку.";
    const finalPrompt = processedPrompt + "\n\n" + baseInstructions;

    logMessage("AIQ: Сформирован промпт: " + finalPrompt);

    const settings = getSettings();
    const model = settings.model; // Можно выбрать другую модель по умолчанию для простых запросов
    const temperature = (typeof temperatureValue === 'number' && !isNaN(temperatureValue)) ? temperatureValue : settings.temperature;
    const maxTokens = settings.maxTokensShortAnswer || 500; // Лимит для коротких ответов
    const retries = settings.retryAttempts || 3;

    // Кэширование
    const cacheKey = "AIQ_" + calculateMD5(finalPrompt + model + temperature + maxTokens);
    const cachedResult = CACHE.get(cacheKey);
    if (cachedResult) {
      logMessage("AIQ: Результат из кэша для ключа " + cacheKey);
      return cachedResult;
    }

    logMessage("AIQ: Запрос к API. Модель: " + model + ", Температура: " + temperature + ", Макс.токены: " + maxTokens);
    const aiResponse = openRouterRequest(finalPrompt, model, temperature, retries, maxTokens);

    if (aiResponse && aiResponse.choices && aiResponse.choices[0] && aiResponse.choices[0].message && aiResponse.choices[0].message.content) {
      const result = aiResponse.choices[0].message.content.trim();
      CACHE.put(cacheKey, result, 21600); // Кэшируем на 6 часов
      logMessage("AIQ: Ответ от AI: " + result);
      return result;
    } else {
      logMessage("AIQ: Ошибка от API или неверный формат ответа: " + JSON.stringify(aiResponse), true);
      return "Ошибка: Не удалось получить ответ от AI.";
    }

  } catch (e) {
    logMessage("AIQ: КРИТИЧЕСКАЯ ОШИБКА: " + e.toString() + " " + e.stack, true);
    return "Ошибка: " + e.message;
  }
}
