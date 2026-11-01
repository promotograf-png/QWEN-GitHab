function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = data.sheetName || 'Лист1';
    const title = data.title || '';
    const subtitle = data.subtitle || '';

    // Ищем существующий лист с таким именем или создаём новый
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    } else {
      sheet.clear();
    }

    const headers = data.headers;
    const rows = data.rows;

    // Строка 1: заголовок "График работы видеоинженеров"
    if (title) {
      const titleRange = sheet.getRange(1, 1);
      titleRange.setValue(title);
      titleRange.setFontWeight('bold');
      titleRange.setFontSize(14);
    }

    // Строка 2: подзаголовок "Октябрь 2026"
    if (subtitle) {
      const subtitleRange = sheet.getRange(2, 1);
      subtitleRange.setValue(subtitle);
      subtitleRange.setFontSize(12);
    }

    // Строка 3: заголовки колонок
    const headerRow = title && subtitle ? 4 : (subtitle ? 3 : (title ? 3 : 1));

    const headerRange = sheet.getRange(headerRow, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setFontWeight('bold');

    // Строки данных
    for (let i = 0; i < rows.length; i++) {
      const rowData = rows[i];
      const values = rowData.map(cell => cell.text);
      const colors = rowData.map(cell => cell.bg);
      const rowRange = sheet.getRange(headerRow + 1 + i, 1, 1, rowData.length);
      rowRange.setValues([values]);
      rowRange.setBackgrounds([colors]);
    }

    return ContentService.createTextOutput(JSON.stringify({status: 'success'}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({status: 'error', message: error.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
