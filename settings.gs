// Функции для работы с настройками

// Предполагается, что эти константы определены где-то (например, в constants.gs)
const HINTS_ENABLED_SETTING = 'HINTS_ENABLED_SETTING';
const FREE_ONLY_SETTING = 'FREE_ONLY_SETTING';

/**
 * Сохраняет API-ключ, температуру, модель, макс. токены и другие настройки в Script Properties.
 * @param {string} apiKey API-ключ OpenRouter.
 * @param {string} temperature Температура генерации (от 0 до 1).
 * @param {string} model Идентификатор модели.
 * @param {string} maxTokens Максимальное количество токенов.
 * @param {boolean} [enableHints] Включить подсказки (может быть undefined).
 * @param {boolean} [freeOnly] Показывать только бесплатные модели (может быть undefined).
 * @returns {string} Сообщение об успехе или ошибке.
 */
function saveApiKeyAndTemperature(apiKey, temperature, model, maxTokens, enableHints, freeOnly) {
  try {
    if (!apiKey) {
      SCRIPT_PROPERTIES.deleteProperty('OPENROUTER_API_KEY');
      SCRIPT_PROPERTIES.deleteProperty(TEMPERATURE_SETTING_KEY);
      SCRIPT_PROPERTIES.deleteProperty(MODEL_SETTING_KEY);
      SCRIPT_PROPERTIES.deleteProperty(MAX_TOKENS_SETTING_KEY);
      SCRIPT_PROPERTIES.deleteProperty(HINTS_ENABLED_SETTING);
      SCRIPT_PROPERTIES.deleteProperty(FREE_ONLY_SETTING);
      return 'Ключ API и все связанные настройки удалены.';
    }

    SCRIPT_PROPERTIES.setProperty('OPENROUTER_API_KEY', apiKey);
    console.log("API ключ сохранен:", apiKey ? "***" : null);

    if (temperature) {
      const tempValue = parseFloat(temperature);
      if (isNaN(tempValue) || tempValue < 0 || tempValue > 1) {
        return 'Ошибка: Температура должна быть числом от 0 до 1.';
      }
      SCRIPT_PROPERTIES.setProperty(TEMPERATURE_SETTING_KEY, tempValue.toString());
      console.log("Температура сохранена:", tempValue);
    } else {
      SCRIPT_PROPERTIES.deleteProperty(TEMPERATURE_SETTING_KEY);
      console.log("Температура удалена");
    }

    if (model) {
      SCRIPT_PROPERTIES.setProperty(MODEL_SETTING_KEY, model);
      console.log("Модель сохранена:", model);
    } else {
      SCRIPT_PROPERTIES.deleteProperty(MODEL_SETTING_KEY);
      console.log("Модель удалена");
    }

    if (maxTokens) {
      const maxTokensValue = parseInt(maxTokens, 10);
      if (isNaN(maxTokensValue) || maxTokensValue <= 0) {
        return 'Ошибка: Макс. токенов должно быть положительным числом.';
      }
      SCRIPT_PROPERTIES.setProperty(MAX_TOKENS_SETTING_KEY, maxTokensValue.toString());
      console.log("Макс. токенов сохранено:", maxTokensValue);
    } else {
      SCRIPT_PROPERTIES.deleteProperty(MAX_TOKENS_SETTING_KEY);
      console.log("Макс. токенов удалено");
    }

    // Сохраняем новые настройки с проверкой на undefined
    const hintsValue = (typeof enableHints === 'boolean') ? enableHints : true; // По умолчанию true
    SCRIPT_PROPERTIES.setProperty(HINTS_ENABLED_SETTING, hintsValue.toString());
    console.log("Подсказки включены:", hintsValue);

    const freeOnlyValue = (typeof freeOnly === 'boolean') ? freeOnly : false; // По умолчанию false
    SCRIPT_PROPERTIES.setProperty(FREE_ONLY_SETTING, freeOnlyValue.toString());
    console.log("Только бесплатные модели:", freeOnlyValue);

    // Проверяем, что настройки действительно сохранились
    const savedTemp = SCRIPT_PROPERTIES.getProperty(TEMPERATURE_SETTING_KEY);
    const savedModel = SCRIPT_PROPERTIES.getProperty(MODEL_SETTING_KEY);
    const savedMaxTokens = SCRIPT_PROPERTIES.getProperty(MAX_TOKENS_SETTING_KEY);
    const savedHints = SCRIPT_PROPERTIES.getProperty(HINTS_ENABLED_SETTING);
    const savedFreeOnly = SCRIPT_PROPERTIES.getProperty(FREE_ONLY_SETTING);

    console.log("Проверка сохраненных настроек - Температура:", savedTemp);
    console.log("Проверка сохраненных настроек - Модель:", savedModel);
    console.log("Проверка сохраненных настроек - Макс. токенов:", savedMaxTokens);
    console.log("Проверка сохраненных настроек - Подсказки:", savedHints);
    console.log("Проверка сохраненных настроек - Бесплатные модели:", savedFreeOnly);

    return `API-ключ и настройки успешно сохранены (температура: ${temperature || 'по умолчанию'}, модель: ${model || 'по умолчанию'}, макс. токенов: ${maxTokens || 'по умолчанию'}, подсказки: ${hintsValue}, бесплатные: ${freeOnlyValue})`;
  } catch (error) {
    console.error("Ошибка при сохранении настроек:", error);
    logMessage(`Ошибка при сохранении настроек: ${error.toString()}`, true);
    return `Ошибка при сохранении настроек: ${error.toString()}`;
  }
}

/**
 * Удаляет API-ключ и все настройки
 * @returns {string} Сообщение об успехе или ошибке.
 */
function deleteApiAndTemperature() {
  try {
    SCRIPT_PROPERTIES.deleteProperty('OPENROUTER_API_KEY');
    SCRIPT_PROPERTIES.deleteProperty(TEMPERATURE_SETTING_KEY);
    SCRIPT_PROPERTIES.deleteProperty(MODEL_SETTING_KEY);
    SCRIPT_PROPERTIES.deleteProperty(MAX_TOKENS_SETTING_KEY);
    SCRIPT_PROPERTIES.deleteProperty(HINTS_ENABLED_SETTING);
    SCRIPT_PROPERTIES.deleteProperty(FREE_ONLY_SETTING);

    // Проверяем, что настройки действительно удалены
    const apiKey = SCRIPT_PROPERTIES.getProperty('OPENROUTER_API_KEY');
    const temp = SCRIPT_PROPERTIES.getProperty(TEMPERATURE_SETTING_KEY);
    const model = SCRIPT_PROPERTIES.getProperty(MODEL_SETTING_KEY);
    const maxTokens = SCRIPT_PROPERTIES.getProperty(MAX_TOKENS_SETTING_KEY);
    const hints = SCRIPT_PROPERTIES.getProperty(HINTS_ENABLED_SETTING);
    const freeOnly = SCRIPT_PROPERTIES.getProperty(FREE_ONLY_SETTING);

    console.log("Проверка удаления - API ключ:", apiKey ? "существует" : "удален");
    console.log("Проверка удаления - Температура:", temp ? "существует" : "удалена");
    console.log("Проверка удаления - Модель:", model ? "существует" : "удалена");
    console.log("Проверка удаления - Макс. токенов:", maxTokens ? "существует" : "удалены");
    console.log("Проверка удаления - Подсказки:", hints ? "существует" : "удалены");
    console.log("Проверка удаления - Бесплатные модели:", freeOnly ? "существует" : "удалены");

    return 'Ключ API и все настройки успешно удалены.';
  } catch (error) {
    console.error("Ошибка при удалении настроек:", error);
    logMessage(`Ошибка при удалении настроек: ${error.toString()}`, true);
    return `Ошибка при удалении настроек: ${error.toString()}`;
  }
}

function getApiKey() {
    return SCRIPT_PROPERTIES.getProperty('OPENROUTER_API_KEY');
}

/**
 * Получает температуру из настроек (для клиента)
 * @returns {string} Значение температуры или null
 */
function getTemperatureFromClient() {
  return SCRIPT_PROPERTIES.getProperty(TEMPERATURE_SETTING_KEY) || null;
}

/**
 * Получает модель из настроек (для клиента)
 * @returns {string} Идентификатор модели или null
 */
function getModelFromClient() {
  return SCRIPT_PROPERTIES.getProperty(MODEL_SETTING_KEY) || null;
}

/**
 * Получает максимальное количество токенов из настроек (для клиента)
 * @returns {string} Значение макс. токенов или null
 */
function getMaxTokensFromClient() {
  return SCRIPT_PROPERTIES.getProperty(MAX_TOKENS_SETTING_KEY) || null;
}

/**
 * Получает API-ключ из настроек (для клиента)
 * @returns {string} API-ключ или null
 */
function getApiKeyFromClient() {
  return SCRIPT_PROPERTIES.getProperty('OPENROUTER_API_KEY') || null;
}

/**
 * Получает настройку включения подсказок
 * @returns {boolean} true если подсказки включены
 */
function getHintsEnabledSetting() {
  const setting = SCRIPT_PROPERTIES.getProperty(HINTS_ENABLED_SETTING);
  return setting === null || setting === 'true'; 
}

/**
 * Получает настройку отображения только бесплатных моделей
 * @returns {boolean} true если показывать только бесплатные модели
 */
function getFreeOnlySetting() {
  const setting = SCRIPT_PROPERTIES.getProperty(FREE_ONLY_SETTING);
  return setting === 'true';
}

/**
 * Сохраняет настройку отображения только бесплатных моделей
 * @param {boolean} freeOnly Показывать только бесплатные модели
 */
function saveFreeOnlySetting(freeOnly) {
  SCRIPT_PROPERTIES.setProperty(FREE_ONLY_SETTING, freeOnly.toString());
}

/**
 * Сохраняет выбранную модель
 * @param {string} model Идентификатор модели
 */
function saveSelectedModel(model) {
  if (model) {
    SCRIPT_PROPERTIES.setProperty(MODEL_SETTING_KEY, model);
  } else {
    SCRIPT_PROPERTIES.deleteProperty(MODEL_SETTING_KEY);
  }
}

/**
 * Получает список доступных моделей для клиента
 * @param {boolean} freeOnly Показывать только бесплатные модели
 * @returns {Array<Object>} Массив объектов с моделями {id: string, name: string}
 */
function getModelsListFromClient(freeOnly) {
  const cacheKey = 'MODELS_LIST_CACHE';
  let models;
  const cached = CACHE.get(cacheKey);
  if (cached) {
    try {
      models = JSON.parse(cached);
    } catch (e) {
      models = [];
    }
  } else {
    const apiKey = getApiKey();
    if (!apiKey) throw new Error('API-ключ не установлен');
    const url = OPENROUTER_API_URL.replace('/chat/completions', '/models');
    const options = {
      method: 'get',
      headers: { Authorization: 'Bearer ' + apiKey },
      muteHttpExceptions: true,
    };
    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() !== 200) {
      throw new Error('Не удалось получить список моделей: ' + response.getContentText());
    }
    const json = JSON.parse(response.getContentText());
    const list = json.data || json.models || [];
    models = list.map(m => ({ id: m.id, name: m.id }));
    CACHE.put(cacheKey, JSON.stringify(models), 3600);
  }
  if (freeOnly) {
    return models.filter(m => m.id.includes(':free') || m.name.toLowerCase().includes('free'));
  }
  return models;
}