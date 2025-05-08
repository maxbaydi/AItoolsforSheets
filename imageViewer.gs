/**
 * Функции для работы с просмотром изображений в Google Таблицах
 */

// Активация просмотрщика изображений - происходит автоматически при открытии
function setupImageViewer() {
  try {
    // Удаляем существующий триггер, если он есть
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'onSelectionChange') {
        ScriptApp.deleteTrigger(trigger);
      }
    });

    // Создаем новый триггер вручную с использованием правильного API
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ScriptApp.newTrigger('onSelectionChange')
      .forSpreadsheet(ss)
      .onEdit()  // Используем onEdit вместо onSelectionChange
      .create();

    console.log("Просмотрщик изображений активирован");
  } catch (e) {
    console.error('Ошибка при настройке просмотрщика изображений: ' + e.toString());
  }
}

// Проверка, является ли объект изображением
function isImageObject(cell) {
  try {
    const images = cell.getSheet().getImages();
    for (const image of images) {
      const range = image.getAnchorCell();
      if (range.getRow() === cell.getRow() && range.getColumn() === cell.getColumn()) {
        return image.getUrl();
      }
    }
    return null;
  } catch (e) {
    console.error('Ошибка при проверке объекта изображения:', e);
    return null;
  }
}

// Обработчик изменения выбора ячейки
function onSelectionChange(e) {
  try {
    const sheet = e.range.getSheet();
    const cell = e.range;
    let imageFound = false;

    // Проверяем только одиночные выделенные ячейки
    if (cell.getNumRows() !== 1 || cell.getNumColumns() !== 1) {
      return;
    }

    // Проверяем, содержит ли ячейка изображение или формулу IMAGE/HYPERLINK
    const formula = cell.getFormula();
    if (formula) {
      // Проверка на IMAGE или HYPERLINK с изображением
      if (formula.toUpperCase().indexOf('=IMAGE') === 0 || 
          (formula.toUpperCase().indexOf('=HYPERLINK') === 0 && isImageUrl(formula))) {
        // Извлекаем URL изображения
        const imageUrl = extractImageUrl(formula);
        if (imageUrl) {
          showImageViewer(imageUrl);
          imageFound = true;
        }
      }
    } else {
      // Проверяем, есть ли в ячейке URL изображения
      const value = cell.getValue();
      if (typeof value === 'string' && isImageUrl(value)) {
        showImageViewer(value);
        imageFound = true;
      } else {
        // Проверяем, является ли объект изображением
        const imageUrl = isImageObject(cell);
        if (imageUrl) {
          showImageViewer(imageUrl);
          imageFound = true;
        }
      }
    }

    // Если в ячейке не найдено изображения, очищаем сохраненный URL и убираем заметку
    if (!imageFound) {
      clearImageData();
      
      // Если у текущей ячейки есть примечание об изображении, удаляем его
      const note = cell.getNote();
      if (note && note.includes('Обнаружено изображение:')) {
        cell.clearNote();
      }
    }
  } catch (e) {
    console.error('Ошибка при обработке выбора ячейки:', e);
  }
}

// Функция для очистки данных об изображении
function clearImageData() {
  try {
    PropertiesService.getUserProperties().deleteProperty('LAST_IMAGE_URL');
    console.log('Данные об изображении очищены');
  } catch (e) {
    console.error('Ошибка при очистке данных изображения:', e);
  }
}

// Проверка, является ли URL ссылкой на изображение
function isImageUrl(url) {
  if (typeof url !== 'string') return false;

  url = url.toLowerCase().trim();
  return url.endsWith('.jpg') || 
         url.endsWith('.jpeg') || 
         url.endsWith('.png') || 
         url.endsWith('.gif') || 
         url.endsWith('.bmp') || 
         url.endsWith('.webp') || 
         url.endsWith('.svg');
}

// Извлечение URL изображения из формулы
function extractImageUrl(formula) {
  try {
    // Извлекаем URL из формулы IMAGE
    if (formula.toUpperCase().indexOf('=IMAGE') === 0) {
      const match = formula.match(/=IMAGE\("([^"]+)"/i) || formula.match(/=IMAGE\('([^']+)'/i);
      return match ? match[1] : null;
    }

    // Извлекаем URL из формулы HYPERLINK
    if (formula.toUpperCase().indexOf('=HYPERLINK') === 0) {
      const match = formula.match(/=HYPERLINK\("([^"]+)"/i) || formula.match(/=HYPERLINK\('([^']+)'/i);
      return match ? match[1] : null;
    }

    return null;
  } catch (e) {
    console.error('Ошибка при извлечении URL изображения:', e);
    return null;
  }
}

// Показать просмотрщик изображений с указанным URL
function showImageViewer(imageUrl) {
  try {
    // Вместо показа модального окна, сохраняем URL для последующего просмотра
    PropertiesService.getUserProperties().setProperty('LAST_IMAGE_URL', imageUrl);
    
    // Логируем для отладки
    console.log('Изображение обнаружено и сохранено для просмотра: ' + imageUrl);
    
    // Создаем заметку на ячейке с изображением для информирования пользователя
    // Убираем отображение URL в заметке, оставляем только уведомление
    const activeCell = SpreadsheetApp.getActiveSpreadsheet().getActiveCell();
    activeCell.setNote('Обнаружено изображение!\n\nДля просмотра выберите: Изображения → 🔍 Просмотреть изображение');
  } catch (e) {
    console.error('Ошибка при обработке изображения:', e);
  }
}

// Новая функция для просмотра последнего изображения через меню
function viewLastImage() {
  try {
    // Проверяем наличие изображения в текущей активной ячейке
    const activeCell = SpreadsheetApp.getActiveSpreadsheet().getActiveCell();
    let imageUrl = null;
    let currentCellHasImage = false;
    
    // Проверяем формулы в активной ячейке
    const formula = activeCell.getFormula();
    if (formula) {
      if (formula.toUpperCase().indexOf('=IMAGE') === 0 || 
          (formula.toUpperCase().indexOf('=HYPERLINK') === 0 && isImageUrl(formula))) {
        imageUrl = extractImageUrl(formula);
        currentCellHasImage = !!imageUrl;
      }
    } else {
      // Проверяем текстовое содержимое
      const value = activeCell.getValue();
      if (typeof value === 'string' && isImageUrl(value)) {
        imageUrl = value;
        currentCellHasImage = true;
      } else {
        // Проверяем наличие объекта изображения
        imageUrl = isImageObject(activeCell);
        currentCellHasImage = !!imageUrl;
      }
    }
    
    // Если в текущей ячейке нет изображения, проверяем, сохранено ли предыдущее изображение
    if (!currentCellHasImage) {
      const note = activeCell.getNote();
      // Если в ячейке нет примечания об изображении, удаляем сохраненный URL
      if (!note || !note.includes('Обнаружено изображение')) {
        clearImageData();
        SpreadsheetApp.getUi().alert('В текущей ячейке нет изображения для просмотра');
        return;
      }
      
      // Используем сохраненный URL только если в текущей ячейке есть примечание об изображении
      imageUrl = PropertiesService.getUserProperties().getProperty('LAST_IMAGE_URL');
    } else {
      // Обновляем сохраненный URL, если нашли изображение в текущей ячейке
      PropertiesService.getUserProperties().setProperty('LAST_IMAGE_URL', imageUrl);
    }
    
    if (!imageUrl) {
      SpreadsheetApp.getUi().alert('Нет изображения для просмотра');
      return;
    }
    
    // Проверяем, является ли URL ссылкой на Google Drive и при необходимости преобразуем ее
    if (imageUrl.includes('drive.google.com') && imageUrl.includes('id=')) {
      const fileIdMatch = imageUrl.match(/id=([^&]+)/);
      if (fileIdMatch && fileIdMatch[1]) {
        const fileId = fileIdMatch[1];
        // Используем вложенный iframe для показа изображения в превью режиме
        // Это обычно работает лучше для Google Drive
        const html = HtmlService.createHtmlOutput(
          `<iframe src="https://drive.google.com/file/d/${fileId}/preview" width="100%" height="600" frameborder="0"></iframe>`
        )
        .setWidth(900)  // Увеличиваем ширину с 800 до 900
        .setHeight(700)  // Увеличиваем высоту с 650 до 700
        .setTitle('Просмотр изображения');
        
        SpreadsheetApp.getUi().showModalDialog(html, 'Просмотр изображения');
        return;
      }
    }
    
    // Если это не Google Drive или не удалось получить ID, используем обычный просмотрщик
    const template = HtmlService.createTemplateFromFile('ImageViewerModal');
    template.imageUrl = imageUrl;
    
    const html = template.evaluate()
                      .setWidth(900)  // Увеличиваем ширину с 800 до 900
                      .setHeight(700)  // Увеличиваем высоту с 600 до 700
                      .setTitle('Просмотр изображения');
    
    // Показываем в диалоге
    SpreadsheetApp.getUi().showModalDialog(html, 'Просмотр изображения');
  } catch (e) {
    console.error('Ошибка при показе изображения:', e);
    SpreadsheetApp.getUi().alert('Ошибка при показе изображения: ' + e.toString());
  }
}

// Показать диалоговое окно для вставки изображения из буфера обмена или файла
function showImagePasteDialog() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('ImagePasteDialog')
      .setWidth(900)  // Увеличиваем ширину с 500 до 900 для лучшего отображения
      .setHeight(800)  // Увеличиваем высоту с 600 до 800 для лучшего соотношения
      .setTitle('Вставить изображение');
    
    SpreadsheetApp.getUi().showModalDialog(html, 'Вставить изображение');
  } catch (e) {
    console.error('Ошибка при открытии диалога вставки изображений:', e);
    SpreadsheetApp.getUi().alert('Ошибка при открытии диалога вставки изображений: ' + e.toString());
  }
}

// Вставка изображения в выбранную ячейку (вызывается из диалога)
function insertImageToActiveCell(imageData, imageName, imageType) {
  try {
    // Проверяем, что есть активная ячейка
    const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const activeRange = SpreadsheetApp.getActiveRange();
    
    if (!activeRange || activeRange.getNumRows() !== 1 || activeRange.getNumColumns() !== 1) {
      return { 
        success: false, 
        error: 'Выберите одну ячейку для вставки изображения' 
      };
    }
    
    // Загружаем изображение на Google Drive
    const imageUrl = uploadImageToDrive(imageData, imageName, imageType);
    
    if (!imageUrl) {
      return { 
        success: false, 
        error: 'Не удалось загрузить изображение' 
      };
    }
    
    // Вставляем формулу IMAGE с URL в ячейку
    const cell = activeRange;
    cell.setFormula(`=IMAGE("${imageUrl}")`);
    
    // Сохраняем URL изображения для возможности просмотра
    PropertiesService.getUserProperties().setProperty('LAST_IMAGE_URL', imageUrl);
    
    return { success: true, imageUrl: imageUrl };
  } catch (e) {
    console.error('Ошибка при вставке изображения:', e);
    return { 
      success: false, 
      error: `Ошибка при вставке изображения: ${e.toString()}` 
    };
  }
}

// Загрузка изображения на Google Drive в специальную папку
function uploadImageToDrive(imageData, imageName, imageType) {
  try {
    // Создаем имя с датой и временем, чтобы избежать дубликатов
    const timestamp = new Date().toISOString().replace(/[^0-9]/g, '');
    const fileName = `image_${timestamp}_${imageName || 'image.png'}`;
    
    // Очищаем base64 заголовок, если он есть
    const base64Data = imageData.split(',')[1] || imageData;
    
    // Преобразуем base64 в blob
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), imageType || 'image/png', fileName);
    
    // Получаем или создаем папку для изображений
    const folder = getOrCreateImageFolder();
    
    if (!folder) {
      throw new Error('Не удалось создать папку для изображений');
    }
    
    // Сохраняем изображение в папку
    const file = folder.createFile(blob);
    
    // Настраиваем доступ к файлу (только для чтения для всех, у кого есть ссылка)
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // Получаем ID файла
    const fileId = file.getId();
    
    // Создаем URL в формате, правильном для формулы IMAGE()
    // Формат: https://drive.google.com/uc?export=view&id=FILE_ID
    const imageUrl = `https://drive.google.com/uc?export=view&id=${fileId}`;
    
    return imageUrl;
  } catch (e) {
    console.error('Ошибка при загрузке изображения на Google Drive:', e);
    return null;
  }
}

// Получение или создание папки для изображений
function getOrCreateImageFolder() {
  try {
    const folderName = 'SheetImages'; // Более короткое и информативное название для папки
    
    // Ищем существующую папку
    let folderIterator = DriveApp.getFoldersByName(folderName);
    
    if (folderIterator.hasNext()) {
      return folderIterator.next();
    }
    
    // Если папки нет, создаем новую в корневом каталоге
    return DriveApp.createFolder(folderName);
  } catch (e) {
    console.error('Ошибка при поиске/создании папки для изображений:', e);
    return null;
  }
}