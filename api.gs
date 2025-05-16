/**
 * Выполняет запрос к API (OpenRouter или VseGPT в зависимости от настроек)
 * @param {string} prompt - Текст запроса
 * @param {string} model - Модель для использования
 * @param {number} temperature - Параметр температуры (0-2)
 * @param {number} [retries=3] - Количество попыток при ошибках
 * @param {number} maxTokens - Максимальное количество токенов
 * @returns {Object} Ответ API
 * @throws {Error} Если запрос не удался после всех попыток
 */
function openRouterRequest(prompt, model, temperature, retries = 3, maxTokens) {
    // Валидация входных параметров
    if (!prompt || typeof prompt !== 'string') {
        throw new Error('Текст запроса должен быть непустой строкой');
    }
    if (!model || typeof model !== 'string') {
        throw new Error('Модель должна быть непустой строкой');
    }
    if (typeof temperature !== 'number' || temperature < 0 || temperature > 2) {
        throw new Error('Температура должна быть числом от 0 до 2');
    }
    
    if (maxTokens !== undefined && (typeof maxTokens !== 'number' || maxTokens <= 0)) {
        throw new Error('Максимальное количество токенов должно быть положительным числом, если указано.');
    }

    const activeApiInfo = getActiveApiKeyInfo(); 
    if (!activeApiInfo || !activeApiInfo.apiKey || !activeApiInfo.apiUrl) {
        throw new Error("Ключ API или URL для выбранного провайдера не настроены. Пожалуйста, проверьте настройки.");
    }

    const headers = {
        'Authorization': `Bearer ${activeApiInfo.apiKey}`,
        'Content-Type': 'application/json',
    };

    if (activeApiInfo.apiType === 'openrouter') {
        headers['HTTP-Referer'] = ScriptApp.getService().getUrl(); 
    }    // Если активный API - VseGPT, используем либо переданную модель, либо модель из настроек VseGPT
    let modelToUse = model;
    if (activeApiInfo.apiType === 'vsegpt' && (!model || model === DEFAULT_MODEL)) {
        const vsegptModel = SCRIPT_PROPERTIES.getProperty(VSEGPT_MODEL_SETTING_KEY);
        if (vsegptModel) {
            modelToUse = vsegptModel;
            logMessage(`Использование модели VseGPT из настроек: ${modelToUse}`);
        }
    }

    const data = {
        'model': modelToUse,
        'messages': [{ 'role': 'user', 'content': prompt }],
        'temperature': temperature,
    };

    if (maxTokens !== undefined) {
        data['max_tokens'] = maxTokens;
    }

    const options = {
        'method': 'post',
        'headers': headers,
        'payload': JSON.stringify(data),
        'muteHttpExceptions': true,
    };

    let attempt = 0;
    while (attempt < retries) {
        attempt++;
        try {
            const response = UrlFetchApp.fetch(activeApiInfo.apiUrl, options);
            const statusCode = response.getResponseCode();
            const responseBody = response.getContentText(); 
            
            logMessage(`${activeApiInfo.apiType.toUpperCase()} запрос: модель ${model}, попытка ${attempt}/${retries}, статус: ${statusCode}, URL: ${activeApiInfo.apiUrl}`);
            if (statusCode !== 200) {
                 logMessage(`Тело ответа ${activeApiInfo.apiType.toUpperCase()} при ошибке: ${responseBody}`, true);
            }

            if (statusCode === 200) {
                return JSON.parse(responseBody);
            } else if (statusCode === 401) {
                throw new Error(`Неверный API-ключ для ${activeApiInfo.apiType.toUpperCase()}. Проверьте ключ в настройках.`);
            } else if (statusCode === 429) {
                throw new Error(`Превышен лимит запросов или недостаточно средств для ${activeApiInfo.apiType.toUpperCase()}. Тело ответа: ${responseBody}`);
            } else if (statusCode === 400 && activeApiInfo.apiType === 'vsegpt') {
                 throw new Error(`Ошибка запроса к VseGPT (400): ${responseBody}. Проверьте параметры запроса или настройки аккаунта VseGPT.`);
            } else {
                throw new Error(`Ошибка API ${activeApiInfo.apiType.toUpperCase()}: ${statusCode} - ${responseBody}`);
            }
        } catch (error) {
            if (!(error.message.startsWith('Неверный API-ключ') || error.message.startsWith('Превышен лимит запросов') || error.message.startsWith('Ошибка запроса к VseGPT') || error.message.startsWith('Ошибка API'))) {
                logMessage(`Исключение при запросе к ${activeApiInfo.apiType.toUpperCase()}: ${error.message}`, true);
            }
            if (attempt === retries) {
                throw error; 
            }
            Utilities.sleep(1000 * Math.pow(2, attempt -1)); 
        }
    }
    throw new Error(`Не удалось выполнить запрос к ${activeApiInfo.apiType.toUpperCase()} после ${retries} попыток.`);
}
