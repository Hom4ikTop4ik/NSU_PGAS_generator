// ⚙️ КОНФИГУРАЦИЯ, ЗАМЕНИТЕ НА СВОИ ID
// Настройки для генерации по шаблону
const TEMPLATE_ID = ''; // ID Google Docs шаблон ПГАС
const FOLDER_ID = ''; // ID папки для сохранения одиночек (можно оставить пустым, тогда будет в корне сохранять)

// Настройки для объединения заполненных в единый файл (папки могут быть одинаковыми или разными, как вам удобно)
// const SOURCE_FOLDER_ID = ''; // ID папки, где лежит куча файлов одиночек (вы можете хоть третью папку создать, откуда именно склейка будет, вдруг вам не все подряд ПГАС надо в сборник)
const SOURCE_FOLDER_ID = FOLDER_ID; // ❗️ верхнее ЛИБО просто оставьте так ❗️
const MERGED_FILE_NAME = 'Сборник ПГАС'; // Имя итогового файла + дата и время в конце
const OUTPUT_FOLDER_ID = '' // ID папки куда сохранить сборник ПГАС (настоятельно рекомендую отдельную папку, иначе может возникнуть склейка сборника, а то и нескольких)


function onFormSubmit(e) {
  try {
    const itemResponses = e.response.getItemResponses();
    
    // === 1. ОБЩАЯ ИНФОРМАЦИЯ ===
    const fio = itemResponses[0]?.getResponse() || '';
    const department = itemResponses[1]?.getResponse() || '';
    const group = itemResponses[2]?.getResponse() || '';
    
    // === 2. СБОР МЕРОПРИЯТИЙ (5 блоков по 4 поля) ===
    // Порядок полей в форме:
    // 0: Название мероприятия
    // 1: Дата участия
    // 2: Статус участника
    // 3: Ссылка на публикацию
    const events = [];
    for (let i = 0; i < 5; i++) {
      const baseIndex = 3 + (i * 4);
      
      const name = itemResponses[baseIndex]?.getResponse() || '';
      const date = itemResponses[baseIndex + 1]?.getResponse() || '';
      const status = itemResponses[baseIndex + 2]?.getResponse() || '';
      const link = itemResponses[baseIndex + 3]?.getResponse() || '';
      
      if (name && name.trim() !== '') {
        events.push({ name, date, status, link });
      }
    }
    
    if (events.length === 0) {
      console.log('Нет мероприятий для добавления');
      return;
    }
    
    // === 3. СОЗДАНИЕ КОПИИ ДОКУМЕНТА ===
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm-ss');
    const fileName = `${fio}_Представление_${timestamp}`;
    
    let folder = DriveApp.getRootFolder();
    if (FOLDER_ID && FOLDER_ID !== '') {
      folder = DriveApp.getFolderById(FOLDER_ID);
    }
    
    const template = DriveApp.getFileById(TEMPLATE_ID);
    const newFile = template.makeCopy(fileName, folder);
    const doc = DocumentApp.openById(newFile.getId());
    const body = doc.getBody();
    
    // === 4. ЗАМЕНА ОСНОВНЫХ ПОЛЕЙ ===
    const currentDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy');
    body.replaceText('{{Doc_date}}', currentDate);
    body.replaceText('{{Department}}', department);
    body.replaceText('{{Group}}', group);
    
    // ФИО — жирным шрифтом
    const fioPlaceholder = body.findText('{{FIO}}');
    if (fioPlaceholder) {
      const text = fioPlaceholder.getElement().asText();
      const start = fioPlaceholder.getStartOffset();
      const end = fioPlaceholder.getEndOffsetInclusive();
      text.deleteText(start, end);
      text.insertText(start, fio);
      text.setBold(start, start + fio.length - 1, true);
    }
    
    // === 5. ЗАПОЛНЕНИЕ ТАБЛИЦЫ ===
    const tables = body.getTables();
    if (tables.length === 0) {
      console.error('В шаблоне нет таблицы!');
      doc.saveAndClose();
      return;
    }
    
    const table = tables[0];
    const totalRows = table.getNumRows();
    
    // Заполняем строки (строка 0 — заголовок, строки 1-5 — пустые заготовки)
    for (let i = 0; i < events.length; i++) {
      const rowIndex = i + 1; // начинаем с первой строки после заголовка
      
      if (rowIndex < totalRows) {
        const row = table.getRow(rowIndex);
        const event = events[i];
        
        // Порядок колонок в вашем шаблоне:
        // 0: №
        // 1: Наименование мероприятия
        // 2: Дата проведения
        // 3: Ссылка на публикацию
        // 4: Статус участника
        
        row.getCell(0).setText((i + 1).toString());           // №
        row.getCell(1).setText(event.name || '');              // Наименование
        row.getCell(2).setText(event.date || '—');             // Дата проведения
        
        // Ссылка (колонка 3)
        const linkCell = row.getCell(3);
        if (event.link && event.link.trim() !== '') {
          linkCell.setText(event.link);
          if (event.link.startsWith('http')) {
            linkCell.setLinkUrl(event.link);
          }
        } else {
          linkCell.setText('');
        }
        
        row.getCell(4).setText(event.status || '—');           // Статус участника
      }
    }
    
    // Удаляем лишние пустые строки (оставляем только заголовок + заполненные)
    const rowsToKeep = events.length + 1;
    for (let i = totalRows - 1; i >= rowsToKeep; i--) {
      table.removeRow(i);
    }
    
    doc.saveAndClose();
    
    console.log(`✅ Документ создан: ${fileName}`);
    console.log(`📊 Мероприятий: ${events.length}`);
    console.log(`🔗 Ссылка: ${newFile.getUrl()}`);
    
  } catch (error) {
    console.error('❌ Ошибка: ' + error.toString());
    console.error('Stack: ' + error.stack);
  }
}

// Вспомогательная функция для жирного форматирования
function makeBold(body, placeholder, replacementText) {
  const element = body.findText(placeholder);
  if (element) {
    const text = element.getElement().asText();
    const start = element.getStartOffset();
    const end = element.getEndOffsetInclusive();
    text.deleteText(start, end);
    text.insertText(start, replacementText);
    text.setBold(start, start + replacementText.length - 1, true);
  }
}


// // // // // // // // // // // // // // // // // // // // // // // // //
// Объединение файлов отдельных людей в единый файл                      //
// (Дарья Александровна просит отправлять единый файл со всеми ребятами) //
// // // // // // // // // // // // // // // // // // // // // // // // //

// ===== ГЛАВНАЯ ФУНКЦИЯ ОБЪЕДИНЕНИЯ =====
function mergeAllDocuments() {
  try {
    // 1. Получаем все документы из папки
    const folder = DriveApp.getFolderById(SOURCE_FOLDER_ID);
    const files = folder.getFilesByType(MimeType.GOOGLE_DOCS);
    
    // 2. Сортируем файлы по дате создания (старые сверху — порядок отправки форм)
    const docList = [];
    while (files.hasNext()) {
      const file = files.next();
      docList.push({
        id: file.getId(),
        name: file.getName(),
        dateCreated: file.getDateCreated()
      });
    }
    // Сортируем по дате создания (от старых к новым)
    docList.sort((a, b) => a.dateCreated - b.dateCreated);
    
    if (docList.length === 0) {
      Logger.log('Нет документов для объединения');
      return;
    }
    
    // 3. Создаём итоговый документ
    
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm-ss');
    const mergedDoc = DocumentApp.create(`${MERGED_FILE_NAME} ${timestamp}`);
    const mergedBody = mergedDoc.getBody();
    
    // 4. Добавляем содержимое каждого документа
    for (let i = 0; i < docList.length; i++) {
      const docInfo = docList[i];
      const sourceDoc = DocumentApp.openById(docInfo.id);
      const sourceBody = sourceDoc.getBody();
      
      // Копируем ВСЁ содержимое исходного документа
      copyDocumentContent(sourceBody, mergedBody);
      
      // Добавляем разрыв страницы после каждого документа (кроме последнего)
      if (i < docList.length - 1) {
        mergedBody.appendPageBreak();
      }
    }
    
    mergedDoc.saveAndClose();

    
    let output_folder = DriveApp.getRootFolder();
    if (OUTPUT_FOLDER_ID && OUTPUT_FOLDER_ID !== '') {
      output_folder = DriveApp.getFolderById(OUTPUT_FOLDER_ID);
      
      // если не корень, то можно переместить
      const file_id = mergedDoc.getId();
      const file = DriveApp.getFileById(file_id);

      file.moveTo(output_folder);
    }

    // mergeDoc и file — ссылка одна и та же
    // однако file будет обращаться к Google Drive: "эй, дай URL на этот файл",
    // а mergeDoc быстрее ведь имеет URL уже с момента создания
    Logger.log(`✅ Итоговый документ создан: ${mergedDoc.getUrl()}`);
    return mergedDoc.getUrl();
    
  } catch (error) {
    console.error('Ошибка при объединении: ' + error.toString());
    throw error;
  }
}

// ===== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ =====

// Копирование содержимого документа с сохранением форматирования
function copyDocumentContent(sourceBody, targetBody) {
  const numChildren = sourceBody.getNumChildren();
  
  for (let i = 0; i < numChildren; i++) {
    const element = sourceBody.getChild(i);
    const elementType = element.getType();
    
    if (elementType === DocumentApp.ElementType.TABLE) {
      const table = element.asTable();
      const copiedTable = targetBody.appendTable(table.copy());
      preserveTableFormatting(table, copiedTable);
    } 
    else if (elementType === DocumentApp.ElementType.PARAGRAPH) {
      const paragraph = element.asParagraph();
      const copiedParagraph = targetBody.appendParagraph(paragraph.copy());
      // Просто копируем выравнивание (самое важное)
      try {
        const alignment = paragraph.getAlignment();
        if (alignment) copiedParagraph.setAlignment(alignment);
      } catch(e) { /* игнорируем */ }
    }
    else if (elementType === DocumentApp.ElementType.LIST_ITEM) {
      const listItem = element.asListItem();
      targetBody.appendListItem(listItem.copy());
    }
    else {
      // Остальные элементы просто копируем
      targetBody.appendParagraph(element.copy());
    }
  }
}

// Сохранение форматирования таблицы
function preserveTableFormatting(sourceTable, targetTable) {
  // Копируем ширину колонок
  const numCols = sourceTable.getRow(0).getNumCells();
  for (let col = 0; col < numCols; col++) {
    const width = sourceTable.getColumnWidth(col);
    if (width) targetTable.setColumnWidth(col, width);
  }
  
  // Копируем стиль ячеек (фон, выравнивание) для каждой строки и колонки
  for (let row = 0; row < sourceTable.getNumRows(); row++) {
    const sourceRow = sourceTable.getRow(row);
    const targetRow = targetTable.getRow(row);
    
    for (let col = 0; col < sourceRow.getNumCells(); col++) {
      const sourceCell = sourceRow.getCell(col);
      const targetCell = targetRow.getCell(col);
      
      // Копируем фон ячейки
      const bgColor = sourceCell.getBackgroundColor();
      if (bgColor) targetCell.setBackgroundColor(bgColor);
      
      // Копируем вертикальное выравнивание
      const vAlign = sourceCell.getVerticalAlignment();
      if (vAlign) targetCell.setVerticalAlignment(vAlign);
      
      // Копируем ширину ячейки
      const cellWidth = sourceCell.getWidth();
      if (cellWidth) targetCell.setWidth(cellWidth);
    }
  }
}

// Создание стиля для разделителя
function createSeparatorStyle() {
  const style = {};
  style[DocumentApp.Attribute.FONT_SIZE] = 8;
  style[DocumentApp.Attribute.FOREGROUND_COLOR] = '#999999';
  style[DocumentApp.Attribute.BOLD] = false;
  return style;
}

// Извлечение номера студента из имени файла
function extractStudentNumber(filename) {
  // Имя файла имеет формат: "Иванов Иван_Представление_2026-01-15_14-30-00"
  const match = filename.match(/^([^_]+)/);
  if (match) {
    return match[1]; // Фамилия Имя
  }
  return filename.substring(0, 30);
}
