/**
 * Выполняет запрос к OpenRouter API
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
    if (typeof maxTokens !== 'number' || maxTokens <= 0) {
        throw new Error('Максимальное количество токенов должно быть положительным числом');
    }

    const apiKey = getApiKey();
    if (!apiKey) {
        throw new Error("API-ключ OpenRouter не установлен. Пожалуйста, установите ключ в настройках.");
    }

    const headers = {
        'Authorization': `Bearer ${apiKey}`,
        'Content-Type': 'application/json',
        'HTTP-Referer': ScriptApp.getService().getUrl(),
    };

    const data = {
        'model': model,
        'messages': [{ 'role': 'user', 'content': prompt }],
        'temperature': temperature,
        'max_tokens': maxTokens,
    };

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
            const response = UrlFetchApp.fetch(OPENROUTER_API_URL, options);
            const statusCode = response.getResponseCode();
            
            // Логирование запроса
            logMessage(`OpenRouter запрос: ${model}, попытка ${attempt}/${retries}, статус: ${statusCode}`);
            
            if (statusCode === 200) {
                return JSON.parse(response.getContentText());
            } else if (statusCode === 401) {
                throw new Error('Неверный API-ключ');
            } else if (statusCode === 429) {
                throw new Error('Превышен лимит запросов');
            } else {
                throw new Error(`Ошибка API: ${statusCode} - ${response.getContentText()}`);
            }
        } catch (error) {
            logMessage(`Ошибка OpenRouter: ${error.message}`, true);
            if (attempt === retries) {
                throw error;
            }
        }
    }
}
