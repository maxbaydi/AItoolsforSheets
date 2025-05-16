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
    let dataText = "";

    if (typeof inputRange === 'string') {
      dataText = inputRange.trim();
    } else if (Array.isArray(inputRange)) {
      if (inputRange.length > 0 && Array.isArray(inputRange[0])) { // 2D array
        dataText = inputRange
          .map(row => row.map(cell => String(cell).trim()).filter(String).join(" "))
          .filter(s => s.length > 0)
          .join("\n");
      } else if (inputRange.length > 0) { // 1D array (вероятно, одна строка или столбец)
        dataText = inputRange.map(cell => String(cell).trim()).filter(String).join(" ");
      }
    } else {
      return "Ошибка: Неверный формат входных данных. Ожидается ячейка или диапазон.";
    }

    if (!dataText) {
      return "Ошибка: Входные данные для описания товара не указаны или пусты.";
    }

    let formattingInstructions = "Не используй HTML и Markdown разметку, такую как звездочки (**) для выделения текста. Не добавляй никаких ссылок или упоминаний о других товарах. Не используй эмодзи.";
    const userFormatType = parseInt(formatType, 10);

    if (userFormatType === 2) {
      formattingInstructions = "Используй HTML-теги для форматирования текста (например, <br> для переноса строки, <b> для жирного шрифта, <i> для курсива, <ul> и <li> для списков). Не обрамляй ответ в общие HTML теги, такие как <html> или <body>. Верни только фрагмент HTML, представляющий описание. Не используй эмодзи.";
    } else if (userFormatType === 3) {
      formattingInstructions = "Используй Markdown разметку для форматирования текста (например, переносы строк, **жирный шрифт**, *курсив*, - для списков). Не используй эмодзи.";
    }

    const promptCoreInstruction = "Сгенерируй SEO-описание для товара на основе следующих данных. Постарайся определить название товара и его ключевые характеристики из предоставленного текста:";
    
    const promptFormatAndOutputRequirements = `
Формат описания:
1. Начни с краткого описания товара, включая его название и основные преимущества, особенности и уникальные selling points.
2. Затем добавь подробное описание, включая:
    - Преимущества товара.
    - Технические характеристики.
3. Используй маркированные списки для перечисления характеристик и преимуществ, если это возможно.
4. Убедись, что описание товара соответствует требованиям поисковых систем и содержит все необходимые ключевые слова для SEO.

${formattingInstructions}

Важно: Верни только сгенерированное описание товара без каких-либо дополнительных фраз, вступлений или пояснений. Не повторяй исходные данные в начале ответа.`;

    const finalPrompt = promptCoreInstruction + "\n\n" + dataText + "\n\n" + promptFormatAndOutputRequirements;

    logMessage("AIDISC: Сформирован промпт: " + finalPrompt);

    const settings = getSettings();
    const model = settings.model;
    const temperature = settings.temperature;
    const maxTokens = settings.maxTokens || 2000; // Установим лимит для описания, если не задан глобально
    const retries = settings.retryAttempts || 3;

    // Кэширование
    const cacheKey = "AIDISC_" + calculateMD5(finalPrompt + model + temperature + maxTokens + (userFormatType || 1));
    const cachedResult = CACHE.get(cacheKey);
    if (cachedResult) {
      logMessage("AIDISC: Результат из кэша для ключа " + cacheKey);
      return cachedResult;
    }

    logMessage("AIDISC: Запрос к API. Модель: " + model + ", Температура: " + temperature + ", Макс.токены: " + maxTokens);
    const aiResponse = openRouterRequest(finalPrompt, model, temperature, retries, maxTokens);

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
 * Использует AI для генерации текста, обобщения информации, классификации данных или анализа тональности текста.
 * Поведение функции стремится соответствовать документированной функции =AI() от Google, но работает как стандартная пользовательская функция, возвращающая результат напрямую.
 *
 * @param {string} prompt Запрос, описывающий желаемое действие (например, "Суммаризируй текст" или "Какая тональность у этого отзыва?"). Если указан dataRange, этот запрос будет применен к данным из dataRange.
 * @param {string|string[][]} [dataRange] Необязательный диапазон ячеек или одна ячейка, данные из которой будут обработаны согласно инструкции в prompt.
 * @param {number} [temperatureValue] Необязательная температура для генерации ответа (например, 0.7). Если не указана, используется значение из настроек.
 * @return {string} Ответ от AI.
 * @customfunction
 */
function AIQ(prompt, dataRange, temperatureValue) {
  try {
    let userQueryPart = String(prompt);
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
      // Формируем запрос: инструкция, затем данные
      userQueryPart = userQueryPart + "\n\n" + dataText;
    } 

    // Базовые инструкции для AI
    const baseInstructions = "Ответь только на поставленный вопрос или выполни инструкцию. Не добавляй никаких вступлений, объяснений, пояснений, извинений или дополнительного текста, кроме самого ответа. В ответе только результат.";
    const finalPrompt = userQueryPart + "\n\n" + baseInstructions;

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
