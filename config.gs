// Константы и настройки
const DEFAULT_MODEL = "deepseek/deepseek-r1:free";
const DEFAULT_MAX_TOKENS = 8192;
const CACHE_SHEET_NAME = "__AI_CACHE__";
const LOG_SHEET_NAME = "__AI_LOGS__";
const OPENROUTER_API_URL = 'https://openrouter.ai/api/v1/chat/completions';
const COLUMN_DELIMITER = '¦';
const ROW_DELIMITER = ';';
const SCRIPT_PROPERTIES = PropertiesService.getScriptProperties();
const DEFAULT_TEMPERATURE = 0.5;
const TEMPERATURE_SETTING_KEY = 'TRANSLATE_TEMPERATURE';
const MODEL_SETTING_KEY = 'MODEL_SETTING';
const MAX_TOKENS_SETTING_KEY = 'MAX_TOKENS_SETTING';
const CACHE = CacheService.getScriptCache();
const ORIGINAL_TEXT_ATTRIBUTE = 'data-original-text';

/**
 * Возвращает объект с настройками модели, температуры и макс. токенов из ScriptProperties.
 */
function getSettings() {
    return {
        model: SCRIPT_PROPERTIES.getProperty(MODEL_SETTING_KEY) || DEFAULT_MODEL,
        temperature: parseFloat(SCRIPT_PROPERTIES.getProperty(TEMPERATURE_SETTING_KEY)) || DEFAULT_TEMPERATURE,
        maxTokens: parseInt(SCRIPT_PROPERTIES.getProperty(MAX_TOKENS_SETTING_KEY), 10) || DEFAULT_MAX_TOKENS,
        retryAttempts: 3
    };
}

// Сохранение настроек
function saveSettings(settings) {
    try {
        if (settings.model) {
            SCRIPT_PROPERTIES.setProperty(MODEL_SETTING_KEY, settings.model);
        }
        if (settings.temperature !== undefined) {
            SCRIPT_PROPERTIES.setProperty(TEMPERATURE_SETTING_KEY, settings.temperature.toString());
        }
        if (settings.maxTokens) {
            SCRIPT_PROPERTIES.setProperty(MAX_TOKENS_SETTING_KEY, settings.maxTokens.toString());
        }
        return "Настройки успешно сохранены";
    } catch (error) {
        logMessage(`Ошибка при сохранении настроек: ${error.toString()}`, true);
        throw new Error(`Не удалось сохранить настройки: ${error.message || error}`);
    }
}