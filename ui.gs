// UI функции
function onOpen() {
    const ui = SpreadsheetApp.getUi();

    // Меню для работы с данными
    const dataMenu = ui.createMenu('Работа с данными')
        .addItem('Универсальный загрузчик', 'showUniversalUploader')
        .addItem('Извлечь данные', 'showExtractDataDialog') 
        .addItem('Суммировать данные', 'showSummarizeSidebar');

    // Меню для объединения ячеек
    const combineMenu = ui.createMenu('Объединить ячейки')
        .addItem('В одну ячейку с пробелами', 'combineCellsWithSpace')
        .addItem('В одну ячейку с переносами', 'combineCellsWithNewline')
        .addItem('Построчно с пробелами', 'combineCellsByRows');

    // Меню для перевода
    const translateMenu = ui.createMenu('Перевод')
        .addItem('На русский (ИИ)', 'translateToRussian')
        .addItem('На английский (ИИ)', 'translateToEnglish')
        .addItem('На китайский (ИИ)', 'translateToChinese')
        .addItem('На испанский (ИИ)', 'translateToSpanish')
        .addItem('На французский (ИИ)', 'translateToFrench')
        .addSeparator()
        .addItem('На русский (Google)', 'quickTranslateToRussian')
        .addItem('На английский (Google)', 'quickTranslateToEnglish')
        .addItem('На китайский (Google)', 'quickTranslateToChinese')
        .addItem('На испанский (Google)', 'quickTranslateToSpanish')
        .addItem('На французский (Google)', 'quickTranslateToFrench')
        .addSeparator()
        .addItem('Настройки перевода', 'showTranslateDialog');

    // Меню для AI функций
    const aiToolsMenu = ui.createMenu('Инструменты')
        .addItem('Заполнить ячейки', 'fillCells')
        .addItem('Создать таблицу', 'showCreateTableSidebar')
        .addItem('Генерировать текст', 'showGenerateTextSidebar');
        
    // Новое меню для работы с изображениями
    const imageMenu = ui.createMenu('Изображения')
        .addItem('🔍 Просмотреть изображение', 'viewLastImage')
        .addItem('📋 Вставить изображение', 'showImagePasteDialog');

    // Главное меню
    ui.createMenu('AI Ассистент')
        .addSubMenu(dataMenu)
        .addSubMenu(translateMenu)
        .addSubMenu(combineMenu)
        .addSubMenu(aiToolsMenu)
        .addSeparator()
        .addItem('⚙️ Настройки API', 'showApiKeyDialog')
        .addToUi();
    
    // Добавляем отдельное меню для изображений
    imageMenu.addToUi();
    
    // Вызываем setupImageViewer() отдельно от создания меню
    try {
        setupImageViewer();
    } catch (e) {
        console.error("Ошибка при настройке просмотрщика изображений:", e);
    }
}

function showUniversalUploader() {
    const html = HtmlService.createHtmlOutputFromFile('UniversalUploader')
        .setWidth(300)
        .setHeight(400);
    SpreadsheetApp.getUi().showSidebar(html);
}

// Показывает сайдбар для извлечения данных
function showExtractDataDialog() {
    const html = HtmlService.createHtmlOutputFromFile('ExtractDataDialog')
        .setWidth(300)
        .setHeight(400);
    SpreadsheetApp.getUi().showSidebar(html);
}

// Показывает сайдбар для создания таблицы
function showCreateTableSidebar() {
    const html = HtmlService.createHtmlOutputFromFile('CreateTableSidebar')
        .setWidth(400)
        .setHeight(300);
    SpreadsheetApp.getUi().showSidebar(html);
}

// Показывает сайдбар для генерации текста
function showGenerateTextSidebar() {
    const html = HtmlService.createHtmlOutputFromFile('GenerateTextSidebar')
        .setTitle('Генерация текста')
        .setWidth(300);
    SpreadsheetApp.getUi().showSidebar(html);
}

// Показывает сайдбар для суммаризации текста
function showSummarizeSidebar() {
    const html = HtmlService.createHtmlOutputFromFile('SummarizeSidebar')
        .setTitle('Суммировать текст')
        .setWidth(300);
    SpreadsheetApp.getUi().showSidebar(html);
}

// Показывает сайдбар для настроек перевода
function showTranslateDialog() {
    const html = HtmlService.createHtmlOutputFromFile('TranslateDialog')
        .setTitle('Перевести')
        .setWidth(300);
    SpreadsheetApp.getUi().showSidebar(html);
}

// Показывает сайдбар для ввода API ключа
function showApiKeyDialog() {
    const ui = SpreadsheetApp.getUi();
    const html = HtmlService.createHtmlOutputFromFile('ApiKeyDialog')
        .setWidth(300)
        .setHeight(350);
    ui.showSidebar(html);
}

// Функция доступная для google.script.run: возвращает заголовки указанной строки
function getTargetHeadersFromServer(headerRow) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const lastCol = sheet.getLastColumn();
    if (lastCol === 0) return [];
    return sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0].map(String);
  } catch (e) {
    return [];
  }
}

// Показывает диалог вставки изображения
function showImagePasteDialog() {
  const html = HtmlService.createHtmlOutputFromFile('ImagePasteDialog')
      .setWidth(800)
      .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Вставить изображение');
}