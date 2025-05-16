// Функции для работы с настройками

// Предполагается, что эти константы определены где-то (например, в constants.gs)
const HINTS_ENABLED_SETTING = 'HINTS_ENABLED_SETTING';
const FREE_ONLY_SETTING = 'FREE_ONLY_SETTING';
const OPENROUTER_API_KEY_PROPERTY = 'OPENROUTER_API_KEY';
const VSEGPT_API_KEY_PROPERTY = 'VSEGPT_API_KEY';
const SELECTED_API_PROPERTY = 'SELECTED_API_PROVIDER';

/**
 * Сохраняет API-ключ, температуру, модель, макс. токены и другие настройки в Script Properties.
 * @param {string} apiKey API-ключ OpenRouter.
 * @param {string} temperature Температура генерации (от 0 до 1).
 * @param {string} model Идентификатор модели.
 * @param {string} maxTokens Максимальное количество токенов.
 * @param {boolean} [enableHints] Включить подсказки (может быть undefined).
 * @param {boolean} [freeOnly] Показывать только бесплатные модели (может быть undefined).
 * @param {string} [vseGptKey] API-ключ VseGPT (может быть undefined).
 * @param {string} [selectedApi] Выбранный API-провайдер ('openrouter' или 'vsegpt', может быть undefined).
 * @returns {string} Сообщение об успехе или ошибке.
 */
function saveApiKeyAndTemperature(apiKey, temperature, model, maxTokens, enableHints, freeOnly, vseGptKey, selectedApi) {
  try {
    if (!apiKey) {
      SCRIPT_PROPERTIES.deleteProperty(OPENROUTER_API_KEY_PROPERTY);
      SCRIPT_PROPERTIES.deleteProperty(TEMPERATURE_SETTING_KEY);
      SCRIPT_PROPERTIES.deleteProperty(MODEL_SETTING_KEY);
      SCRIPT_PROPERTIES.deleteProperty(VSEGPT_MODEL_SETTING_KEY); // Удаляем модель и для VseGPT
      SCRIPT_PROPERTIES.deleteProperty(MAX_TOKENS_SETTING_KEY);
      SCRIPT_PROPERTIES.deleteProperty(HINTS_ENABLED_SETTING);
      SCRIPT_PROPERTIES.deleteProperty(FREE_ONLY_SETTING);
      // Также удаляем выбранный API, если основной ключ удаляется
      SCRIPT_PROPERTIES.deleteProperty(SELECTED_API_PROPERTY);
      PropertiesService.getUserProperties().deleteProperty(SELECTED_API_PROPERTY);
      return 'Ключ API и все связанные настройки удалены.';
    }

    SCRIPT_PROPERTIES.setProperty(OPENROUTER_API_KEY_PROPERTY, apiKey);
    console.log("API ключ сохранен:", apiKey ? "***" : null);

    if (temperature) {
      const tempValue = parseFloat(temperature);
      if (isNaN(tempValue) || tempValue < 0 || tempValue > 1) {
        return 'Ошибка: Температура должна быть числом от 0 до 1.';
      }
      SCRIPT_PROPERTIES.setProperty(TEMPERATURE_SETTING_KEY, tempValue.toString());
      console.log("Температура сохранена:", tempValue);
    } else if (temperature === '' || temperature === null) { // Явное удаление, если передана пустая строка или null
      SCRIPT_PROPERTIES.deleteProperty(TEMPERATURE_SETTING_KEY);
      console.log("Температура удалена");
    }

    if (maxTokens) {
      const maxTokensValue = parseInt(maxTokens, 10);
      if (isNaN(maxTokensValue) || maxTokensValue <= 0) {
        return 'Ошибка: Макс. токенов должно быть положительным числом.';
      }
      SCRIPT_PROPERTIES.setProperty(MAX_TOKENS_SETTING_KEY, maxTokensValue.toString());
      console.log("Макс. токенов сохранено:", maxTokensValue);
    } else if (maxTokens === '' || maxTokens === null) { // Явное удаление
      SCRIPT_PROPERTIES.deleteProperty(MAX_TOKENS_SETTING_KEY);
      console.log("Макс. токенов удалено");
    }

    const hintsValue = (typeof enableHints === 'boolean') ? enableHints : (SCRIPT_PROPERTIES.getProperty(HINTS_ENABLED_SETTING) !== 'false'); // Сохраняем текущее, если не передано
    SCRIPT_PROPERTIES.setProperty(HINTS_ENABLED_SETTING, hintsValue.toString());
    console.log("Подсказки включены:", hintsValue);

    const freeOnlyValue = (typeof freeOnly === 'boolean') ? freeOnly : (SCRIPT_PROPERTIES.getProperty(FREE_ONLY_SETTING) === 'true'); // Сохраняем текущее, если не передано
    SCRIPT_PROPERTIES.setProperty(FREE_ONLY_SETTING, freeOnlyValue.toString());
    console.log("Только бесплатные модели:", freeOnlyValue);
    
    const userProperties = PropertiesService.getUserProperties();
    if (vseGptKey !== undefined) {
      if (vseGptKey) {
        SCRIPT_PROPERTIES.setProperty(VSEGPT_API_KEY_PROPERTY, vseGptKey);
        console.log("VseGPT API ключ сохранен: ***");
      } else {
        SCRIPT_PROPERTIES.deleteProperty(VSEGPT_API_KEY_PROPERTY);
        console.log("VseGPT API ключ удален");
      }
    }

    let effectiveSelectedApi = SCRIPT_PROPERTIES.getProperty(SELECTED_API_PROPERTY) || 'openrouter';

    if (selectedApi !== undefined) {
      if (selectedApi && (selectedApi === 'openrouter' || selectedApi === 'vsegpt')) {
        SCRIPT_PROPERTIES.setProperty(SELECTED_API_PROPERTY, selectedApi);
        userProperties.setProperty(SELECTED_API_PROPERTY, selectedApi);
        effectiveSelectedApi = selectedApi;
        console.log("Выбранный API сохранен:", effectiveSelectedApi);
      } else if (selectedApi === null || selectedApi === '') { // Сброс на дефолт
        SCRIPT_PROPERTIES.setProperty(SELECTED_API_PROPERTY, 'openrouter');
        userProperties.setProperty(SELECTED_API_PROPERTY, 'openrouter');
        effectiveSelectedApi = 'openrouter';
        console.log("Выбранный API установлен на значение по умолчанию: openrouter");
      }
    }

    if (model !== undefined) { // Только если параметр model был передан
      if (model) { // Если модель не пустая
        if (effectiveSelectedApi === 'vsegpt') {
          SCRIPT_PROPERTIES.setProperty(VSEGPT_MODEL_SETTING_KEY, model);
          console.log("Модель для VseGPT сохранена:", model);
        } else { // 'openrouter'
          SCRIPT_PROPERTIES.setProperty(MODEL_SETTING_KEY, model);
          console.log("Модель для OpenRouter сохранена:", model);
        }
      } else { // Если model пустая строка или null - удаляем для текущего API
        if (effectiveSelectedApi === 'vsegpt') {
          SCRIPT_PROPERTIES.deleteProperty(VSEGPT_MODEL_SETTING_KEY);
          console.log("Модель для VseGPT удалена");
        } else { // 'openrouter'
          SCRIPT_PROPERTIES.deleteProperty(MODEL_SETTING_KEY);
          console.log("Модель для OpenRouter удалена");
        }
      }
    }
    
    // Проверка сохраненных настроек
    const savedTemp = SCRIPT_PROPERTIES.getProperty(TEMPERATURE_SETTING_KEY);
    const savedOpenRouterModel = SCRIPT_PROPERTIES.getProperty(MODEL_SETTING_KEY);
    const savedVseGptModel = SCRIPT_PROPERTIES.getProperty(VSEGPT_MODEL_SETTING_KEY);
    const savedMaxTokens = SCRIPT_PROPERTIES.getProperty(MAX_TOKENS_SETTING_KEY);
    const savedHints = SCRIPT_PROPERTIES.getProperty(HINTS_ENABLED_SETTING);
    const savedFreeOnly = SCRIPT_PROPERTIES.getProperty(FREE_ONLY_SETTING);
    const finalSelectedApi = SCRIPT_PROPERTIES.getProperty(SELECTED_API_PROPERTY) || 'openrouter';

    console.log("Проверка сохраненных настроек - Температура:", savedTemp);
    console.log("Проверка сохраненных настроек - Модель OpenRouter:", savedOpenRouterModel);
    console.log("Проверка сохраненных настроек - Модель VseGPT:", savedVseGptModel);
    console.log("Проверка сохраненных настроек - Макс. токенов:", savedMaxTokens);
    console.log("Проверка сохраненных настроек - Подсказки:", savedHints);
    console.log("Проверка сохраненных настроек - Бесплатные модели:", savedFreeOnly);
    console.log("Проверка сохраненных настроек - Выбранный API:", finalSelectedApi);

    // Формируем сообщение об успехе
    let modelMessage = 'по умолчанию';
    if (finalSelectedApi === 'vsegpt') {
      modelMessage = savedVseGptModel || 'не указана';
    } else {
      modelMessage = savedOpenRouterModel || 'не указана';
    }

    return `API-ключ и настройки успешно сохранены (API: ${finalSelectedApi}, температура: ${savedTemp || 'по умолчанию'}, модель: ${modelMessage}, макс. токенов: ${savedMaxTokens || 'по умолчанию'}, подсказки: ${hintsValue}, бесплатные: ${freeOnlyValue})`;
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
    SCRIPT_PROPERTIES.deleteProperty(OPENROUTER_API_KEY_PROPERTY);
    SCRIPT_PROPERTIES.deleteProperty(TEMPERATURE_SETTING_KEY);
    SCRIPT_PROPERTIES.deleteProperty(MODEL_SETTING_KEY); // Модель для OpenRouter
    SCRIPT_PROPERTIES.deleteProperty(MAX_TOKENS_SETTING_KEY);
    SCRIPT_PROPERTIES.deleteProperty(HINTS_ENABLED_SETTING);
    SCRIPT_PROPERTIES.deleteProperty(FREE_ONLY_SETTING);
    
    // Удаление настроек VseGPT API, модели и выбранного провайдера
    const userProperties = PropertiesService.getUserProperties();
    SCRIPT_PROPERTIES.deleteProperty(VSEGPT_API_KEY_PROPERTY);
    SCRIPT_PROPERTIES.deleteProperty(VSEGPT_MODEL_SETTING_KEY); // Модель для VseGPT
    
    // Удаляем выбранный API из обоих хранилищ
    userProperties.deleteProperty(SELECTED_API_PROPERTY);
    SCRIPT_PROPERTIES.deleteProperty(SELECTED_API_PROPERTY);

    // Проверяем, что настройки действительно удалены
    const apiKey = SCRIPT_PROPERTIES.getProperty(OPENROUTER_API_KEY_PROPERTY);
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

/**
 * Сохраняет API-ключи и выбранный API.
 * @param {string} openRouterApiKey API-ключ для OpenRouter.
 * @param {string} vseGptApiKey API-ключ для VseGPT.
 * @param {string} selectedApi Выбранный API ('openrouter' или 'vsegpt').
 */
function saveApiKeys(openRouterApiKey, vseGptApiKey, selectedApi) {
  const userProperties = PropertiesService.getUserProperties();
  if (openRouterApiKey) {
    SCRIPT_PROPERTIES.setProperty(OPENROUTER_API_KEY_PROPERTY, openRouterApiKey);
  } else {
    SCRIPT_PROPERTIES.deleteProperty(OPENROUTER_API_KEY_PROPERTY);
  }
  if (vseGptApiKey) {
    SCRIPT_PROPERTIES.setProperty(VSEGPT_API_KEY_PROPERTY, vseGptApiKey);
  } else {
    SCRIPT_PROPERTIES.deleteProperty(VSEGPT_API_KEY_PROPERTY);
  }
  if (selectedApi) {
    // Сохраняем в оба хранилища для согласованности
    SCRIPT_PROPERTIES.setProperty(SELECTED_API_PROPERTY, selectedApi);
    userProperties.setProperty(SELECTED_API_PROPERTY, selectedApi);
  } else {
    // По умолчанию openrouter
    SCRIPT_PROPERTIES.setProperty(SELECTED_API_PROPERTY, 'openrouter');
    userProperties.setProperty(SELECTED_API_PROPERTY, 'openrouter');
  }
  
  // Очищаем кэш моделей при изменении API, чтобы избежать проблем
  clearModelsCache('openrouter');
  clearModelsCache('vsegpt');
  
  logMessage(`Настройки API сохранены. Выбранный API: ${selectedApi}`);
}

/**
 * Загружает сохраненные API-ключи и выбранный API.
 * @returns {object} Объект с ключами openRouterApiKey, vseGptApiKey и selectedApi.
 */
function loadApiKeys() {
  // Получаем информацию из getActiveApiKeyInfo для корректной синхронизации API между хранилищами
  try {
    const apiInfo = getActiveApiKeyInfo();
    return {
      openRouterApiKey: SCRIPT_PROPERTIES.getProperty(OPENROUTER_API_KEY_PROPERTY),
      vseGptApiKey: SCRIPT_PROPERTIES.getProperty(VSEGPT_API_KEY_PROPERTY),
      selectedApi: apiInfo.apiType
    };
  } catch (e) {
    // В случае ошибки (например, API ключ не установлен) используем значения по умолчанию
    return {
      openRouterApiKey: SCRIPT_PROPERTIES.getProperty(OPENROUTER_API_KEY_PROPERTY),
      vseGptApiKey: SCRIPT_PROPERTIES.getProperty(VSEGPT_API_KEY_PROPERTY),
      selectedApi: 'openrouter' // Значение по умолчанию
    };
  }
}

/**
 * Получает информацию о текущем активном API (ключ, тип, URL).
 * @returns {object|null} Объект с apiKey, apiType, apiUrl или null, если не настроено.
 * @throws {Error} Если ключ для выбранного API не установлен.
 */
function getActiveApiKeyInfo() {
  // Проверяем сначала SCRIPT_PROPERTIES, затем userProperties для обратной совместимости
  const userProperties = PropertiesService.getUserProperties();
  const selectedApiFromScript = SCRIPT_PROPERTIES.getProperty(SELECTED_API_PROPERTY);
  const selectedApiFromUser = userProperties.getProperty(SELECTED_API_PROPERTY);
  
  // Используем значение из SCRIPT_PROPERTIES если оно есть, иначе из userProperties
  const selectedApi = selectedApiFromScript || selectedApiFromUser || 'openrouter';
  
  // Для обеспечения согласованности, синхронизируем значения между хранилищами
  if (selectedApiFromScript !== selectedApiFromUser) {
    if (selectedApiFromScript) {
      userProperties.setProperty(SELECTED_API_PROPERTY, selectedApiFromScript);
    } else if (selectedApiFromUser) {
      SCRIPT_PROPERTIES.setProperty(SELECTED_API_PROPERTY, selectedApiFromUser);
    }
    logMessage(`Синхронизирован выбранный API между хранилищами: ${selectedApi}`);
  }
  
  let apiKey;
  let apiUrl;

  if (selectedApi === 'openrouter') {
    apiKey = SCRIPT_PROPERTIES.getProperty(OPENROUTER_API_KEY_PROPERTY);
    apiUrl = OPENROUTER_API_URL;
    if (!apiKey) {
      throw new Error("API-ключ OpenRouter не установлен. Пожалуйста, установите ключ в настройках.");
    }
  } else if (selectedApi === 'vsegpt') {
    apiKey = SCRIPT_PROPERTIES.getProperty(VSEGPT_API_KEY_PROPERTY);
    apiUrl = VSEGPT_API_URL; // Убедитесь, что VSEGPT_API_URL определен в config.gs
    if (!apiKey) {
      throw new Error("API-ключ VseGPT не установлен. Пожалуйста, установите ключ в настройках.");
    }
  } else {
    throw new Error("Неизвестный тип API выбран в настройках.");
  }

  return {
    apiKey: apiKey,
    apiType: selectedApi,
    apiUrl: apiUrl
  };
}

/**
 * @deprecated Используйте getActiveApiKeyInfo
 * Получает API-ключ OpenRouter.
 * @returns {string|null} API-ключ или null, если не установлен.
 */
function getApiKey() {
  const userProperties = PropertiesService.getUserProperties();
  return userProperties.getProperty(OPENROUTER_API_KEY_PROPERTY);
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
  return SCRIPT_PROPERTIES.getProperty(OPENROUTER_API_KEY_PROPERTY) || null;
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
  try {
    const apiInfo = getActiveApiKeyInfo();
    const currentSelectedApi = apiInfo.apiType;

    if (model) {
      if (currentSelectedApi === 'vsegpt') {
        SCRIPT_PROPERTIES.setProperty(VSEGPT_MODEL_SETTING_KEY, model);
        console.log("Модель для VseGPT сохранена (через saveSelectedModel):", model);
      } else { // 'openrouter' или по умолчанию
        SCRIPT_PROPERTIES.setProperty(MODEL_SETTING_KEY, model);
        console.log("Модель для OpenRouter сохранена (через saveSelectedModel):", model);
      }
    } else {
      // Если модель пустая (null или пустая строка), удаляем ее для текущего провайдера
      if (currentSelectedApi === 'vsegpt') {
        SCRIPT_PROPERTIES.deleteProperty(VSEGPT_MODEL_SETTING_KEY);
        console.log("Модель для VseGPT удалена (через saveSelectedModel)");
      } else { // 'openrouter' или по умолчанию
        SCRIPT_PROPERTIES.deleteProperty(MODEL_SETTING_KEY);
        console.log("Модель для OpenRouter удалена (через saveSelectedModel)");
      }
    }
  } catch (error) {
    console.error("Ошибка при сохранении выбранной модели:", error);
    logMessage(`Ошибка при сохранении выбранной модели: ${error.toString()}`, true);
    // Можно выбросить ошибку дальше или вернуть false/статус
  }
}

/**
 * Получает список доступных моделей для клиента
 * @param {boolean} freeOnly Показывать только бесплатные модели
 * @param {boolean} [skipCache=false] Пропустить кэш и загрузить заново
 * @returns {Array<Object>} Массив объектов с моделями {id: string, name: string}
 */
function getModelsListFromClient(apiProviderToList, freeOnly, skipCache = false) {
  try {
    // const apiInfo = getActiveApiKeyInfo(); // Больше не используем для определения API здесь
    // const selectedApi = apiInfo.apiType; // Используем apiProviderToList
    const selectedApi = apiProviderToList;

    logMessage(`Запрос списка моделей для ${selectedApi} (передан клиентом), skipCache: ${skipCache}, freeOnly: ${freeOnly}.`);

    const cacheKey = 'MODELS_LIST_CACHE_' + selectedApi + (freeOnly ? '_free' : '_all');
    let models;
    
    // logMessage(`Запрос списка моделей для ${selectedApi}, skipCache: ${skipCache}, freeOnly: ${freeOnly}. CacheKey: ${cacheKey}`); // Дублирующее сообщение удалено
    
    if (skipCache) {
      CACHE.remove(cacheKey);
      logMessage(`Принудительная очистка кэша для ${cacheKey}`);
    }
    
    const userProperties = PropertiesService.getUserProperties();
    // Логика с lastApiUsedForCacheProperty должна учитывать, что selectedApi теперь передается
    // и может отличаться от глобально сохраненного. Эта логика кэширования корректна.
    const lastApiUsedForCacheProperty = 'LAST_API_USED_FOR_MODELS_CACHE_V2'; 
    const lastApiAndFreeOnly = userProperties.getProperty(lastApiUsedForCacheProperty);
    const currentApiAndFreeOnly = `${selectedApi}_${freeOnly}`;

    if (lastApiAndFreeOnly && lastApiAndFreeOnly !== currentApiAndFreeOnly) {
      CACHE.remove(cacheKey); 
      logMessage(`Обнаружена смена API/freeOnly (для листинга) с ${lastApiAndFreeOnly} на ${currentApiAndFreeOnly}. Кэш для ${cacheKey} будет обновлен.`);
    }
    userProperties.setProperty(lastApiUsedForCacheProperty, currentApiAndFreeOnly);
    
    const cached = !skipCache ? CACHE.get(cacheKey) : null;
    if (cached) {
      try {
        models = JSON.parse(cached);
        logMessage(`Загружены модели из кэша (${cacheKey}), найдено: ${models.length}`);
      } catch (e) {
        models = [];
        logMessage(`Ошибка при разборе кэша моделей (${cacheKey}): ${e}`, true);
      }
    } else {
      logMessage(`Кэш (${cacheKey}) не найден или пропущен. Загрузка с API ${selectedApi}...`);
      
      let apiKeyToUse;
      if (selectedApi === 'openrouter') {
        apiKeyToUse = SCRIPT_PROPERTIES.getProperty(OPENROUTER_API_KEY_PROPERTY);
        if (!apiKeyToUse) {
          logMessage('API-ключ OpenRouter не установлен. Невозможно загрузить модели.', true);
          throw new Error("API-ключ OpenRouter не установлен. Пожалуйста, установите ключ в настройках, чтобы загрузить список моделей.");
        }
        logMessage(`Fetching OpenRouter models. Key: Exists`, false); // Ключ проверен выше

        const url = OPENROUTER_API_URL.replace('/chat/completions', '/models');
        logMessage(`Fetching from URL: ${url}`, false);
        const options = {
          method: 'get',
          headers: { Authorization: 'Bearer ' + apiKeyToUse },
          muteHttpExceptions: true,
        };
        
        let response;
        try {
          response = UrlFetchApp.fetch(url, options);
        } catch (e) {
          logMessage(`UrlFetchApp.fetch FAILED for OpenRouter models: ${e.toString()}`, true);
          CACHE.remove(cacheKey); // Clear cache on fetch error
          throw new Error(`Ошибка сети при запросе моделей OpenRouter: ${e.message}`);
        }
        
        const responseCode = response.getResponseCode();
        const responseText = response.getContentText();
        logMessage(`OpenRouter /models response code: ${responseCode}. Length: ${responseText.length}`, false);

        if (responseCode !== 200) {
          logMessage(`Ошибка загрузки моделей с OpenRouter API: ${responseCode} - ${responseText.substring(0, 500)}`, true);
          CACHE.remove(cacheKey); 
          throw new Error(`Не удалось получить список моделей OpenRouter (${responseCode}): ${responseText.substring(0, 200)}`);
        }
        
        let json;
        try {
          json = JSON.parse(responseText);
          logMessage(`OpenRouter /models response JSON parsed. Preview: ${responseText.substring(0, 200)}...`, false);
        } catch (e) {
          logMessage(`Ошибка при разборе ответа от OpenRouter API: ${e.toString()}. Response: ${responseText.substring(0,500)}`, true);
          CACHE.remove(cacheKey);
          throw new Error('Невозможно разобрать ответ от OpenRouter API: ' + e.toString());
        }
        
        logMessage(`Parsed JSON keys: ${JSON.stringify(Object.keys(json))}`, false);

        const list = json.data || []; 
        if (!json.data) {
            logMessage(json.models ? 'json.data missing, json.models was present but not used by default.' : 'json.data missing.', true);
        }

        if (!Array.isArray(list)) {
            logMessage(`Expected 'list' (from json.data) to be an array, but got: ${typeof list}. Value: ${JSON.stringify(list).substring(0,200)}`, true);
            models = [];
        } else {
            logMessage(`Raw list from API (length ${list.length}): ${JSON.stringify(list.slice(0, 1))}... (first item shown if exists)`, false);
            models = list.map(m => {
              if (typeof m !== 'object' || m === null || !m.id) {
                logMessage(`Skipping invalid model entry: ${JSON.stringify(m)}`, true);
                return null; 
              }
              if (freeOnly) {
                const pricing = m.pricing || {}; 
                const promptPrice = typeof pricing.prompt === 'string' ? parseFloat(pricing.prompt) : (typeof pricing.prompt === 'number' ? pricing.prompt : -1);
                const completionPrice = typeof pricing.completion === 'string' ? parseFloat(pricing.completion) : (typeof pricing.completion === 'number' ? pricing.completion : -1);
                if (!(promptPrice === 0 && completionPrice === 0)) {
                  return null; 
                }
              }
              return { id: m.id, name: m.name || m.id }; 
            }).filter(m => m !== null);
        }
        
        if (models.length === 0) {
          logMessage(`OpenRouter: After processing and filtering (freeOnly=${freeOnly}), list resulted in zero valid models. Original list length: ${list.length}`, true);
        } else {
          logMessage(`Загружено и обработано ${models.length} моделей из OpenRouter API (freeOnly=${freeOnly}).`);
        }

      } else if (selectedApi === 'vsegpt') {
        // apiKeyToUse = SCRIPT_PROPERTIES.getProperty(VSEGPT_API_KEY_PROPERTY); // Ключ не нужен для списка моделей VseGPT
        models = []; 
        logMessage('Для VseGPT список моделей не загружается (согласно прямому указанию API), предполагается ручной ввод.', false);
      } else {
        CACHE.remove(cacheKey); 
        throw new Error('Неизвестный провайдер API для списка моделей: ' + selectedApi);
      }
      
      if (models && models.length > 0) {
        CACHE.put(cacheKey, JSON.stringify(models), 3600); 
        logMessage(`Сохранено ${models.length} моделей в кэше (${cacheKey})`);
      } else {
        logMessage(`Нет моделей для сохранения в кэше (${cacheKey})`, models ? false : true); // Log as error if models is null/undefined
        // If models is an empty array [], it's not an error, just no models to cache.
        if (!models) models = []; // Ensure models is at least an empty array if it became null
      }
    }
    return models;
  } catch (error) {
    logMessage(`КРИТИЧЕСКАЯ ОШИБКА в getModelsListFromClient: ${error.toString()} ${error.stack ? error.stack : ''}`, true);
    return [];
  }
}

/**
 * Загружает все настройки для диалогового окна ApiKeyDialog
 * @returns {Object} Объект со всеми настройками
 */
function loadSettingsForDialog() {
  try {
    const userProperties = PropertiesService.getUserProperties();
    // Загрузка ключей API и выбранного API
    const apiKey = SCRIPT_PROPERTIES.getProperty(OPENROUTER_API_KEY_PROPERTY) || '';
    const vseGptApiKey = SCRIPT_PROPERTIES.getProperty(VSEGPT_API_KEY_PROPERTY) || '';
    // Определяем выбранный API, отдавая приоритет SCRIPT_PROPERTIES, затем userProperties, и по умолчанию 'openrouter'
    const selectedApi = SCRIPT_PROPERTIES.getProperty(SELECTED_API_PROPERTY) || userProperties.getProperty(SELECTED_API_PROPERTY) || 'openrouter';
    
    // Загрузка других настроек
    const temperature = SCRIPT_PROPERTIES.getProperty(TEMPERATURE_SETTING_KEY) || '0.7';
    const maxTokens = SCRIPT_PROPERTIES.getProperty(MAX_TOKENS_SETTING_KEY) || '1000';
    const enableHints = SCRIPT_PROPERTIES.getProperty(HINTS_ENABLED_SETTING) !== 'false'; // true по умолчанию
    const freeOnly = SCRIPT_PROPERTIES.getProperty(FREE_ONLY_SETTING) === 'true'; // false по умолчанию

    // Загружаем сохраненные модели для каждого провайдера
    const openRouterModel = SCRIPT_PROPERTIES.getProperty(MODEL_SETTING_KEY) || '';
    const vseGptModel = SCRIPT_PROPERTIES.getProperty(VSEGPT_MODEL_SETTING_KEY) || '';

    // Определяем активную модель на основе текущего selectedApi
    let activeModel;
    if (selectedApi === 'vsegpt') {
      activeModel = vseGptModel;
    } else { // 'openrouter' или по умолчанию
      activeModel = openRouterModel;
    }
    
    return {
      apiKey: apiKey,
      vseGptApiKey: vseGptApiKey,
      selectedApi: selectedApi,
      temperature: temperature,
      model: activeModel, // Модель для текущего активного API
      openRouterModel: openRouterModel, // Сохраненная модель для OpenRouter
      vseGptModel: vseGptModel,       // Сохраненная модель для VseGPT
      maxTokens: maxTokens,
      enableHints: enableHints,
      freeOnly: freeOnly
    };
  } catch (error) {
    console.error("Ошибка при загрузке настроек:", error);
    logMessage(`Ошибка при загрузке настроек: ${error.toString()}`, true);
    throw new Error(`Не удалось загрузить настройки: ${error.message || error}`);
  }
}

/**
 * Очищает кэш моделей для указанного API
 * @param {string} apiType - Тип API ('openrouter' или 'vsegpt')
 * @returns {boolean} - true если кэш успешно очищен
 */
function clearModelsCache(apiType) {
  const cacheKey = 'MODELS_LIST_CACHE_' + apiType;
  CACHE.remove(cacheKey);
  
  // Также очищаем кэш "последнего API" для гарантии перезагрузки
  const userProperties = PropertiesService.getUserProperties();
  userProperties.deleteProperty('LAST_USED_API');
  
  logMessage(`Кэш моделей для ${apiType} очищен.`);
  return true;
}