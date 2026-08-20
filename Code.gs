const SPREADSHEET_ID = '1RB1WcfHQwfJubjb6Y2XUfTZXOTI2-SHCI09jsyUHk0w';

function doGet(e) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const menuSheet = ss.getSheetByName('menu');
  const data = menuSheet.getDataRange().getValues();
  const [header, ...rows] = data;

  const menu = rows.map(row => {
    const obj = {};
    header.forEach((key, i) => {
      obj[String(key).trim()] = row[i];
    });
    return obj;
  });

  return ContentService
    .createTextOutput(JSON.stringify({ menu }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('order');

  try {
    if (!sheet) {
      // Частая причина "заказы не пишутся в таблицу" — опечатка в названии листа
      // или лист называется иначе (например "Order" с большой буквы, "orders" и т.п.)
      throw new Error('Лист с именем "order" не найден в таблице. Проверьте название вкладки.');
    }

    const data = JSON.parse(e.postData.contents);
    const name = data.name || '';
    const phone = data.phone || '';
    const address = data.address || '';
    const order = data.items || '';
    const payment = data.payment || '';
    const total = data.total || '';

    sheet.appendRow([
      new Date(),
      name,
      phone,
      address,
      order,
      payment,
      total
    ]);

    // Уведомление в Telegram — токен и chat_id хранятся только здесь, на сервере
    // (Script Properties), в код сайта они никогда не попадают.
    // Если уведомление не отправится (например, токен ещё не настроен) —
    // заказ всё равно уже записан в таблицу строкой выше, это не сломает запись.
    if (data.text) {
      try {
        sendToTelegram_(data.text);
      } catch (telegramErr) {
        console.error('Telegram notify failed: ' + telegramErr.message);
      }
    }

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Отправляет текст заказа в Telegram. Токен и chat_id читаются из
 * Script Properties (Настройки проекта → Свойства скрипта) — это
 * единственное место, где они хранятся. В коде сайта их нет.
 */
function sendToTelegram_(text) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('TELEGRAM_BOT_TOKEN');
  const chatId = props.getProperty('TELEGRAM_CHAT_ID');

  if (!token || !chatId) {
    throw new Error('TELEGRAM_BOT_TOKEN / TELEGRAM_CHAT_ID не заданы в Script Properties');
  }

  const url = 'https://api.telegram.org/bot' + token + '/sendMessage';
  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ chat_id: chatId, text: text }),
    muteHttpExceptions: true
  });

  const result = JSON.parse(response.getContentText());
  if (!result.ok) {
    throw new Error('Telegram API error: ' + response.getContentText());
  }
}
