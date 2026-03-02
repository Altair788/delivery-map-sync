/**
 * Delivery Map Sync — синхронизация Google Sheets с картой
 * Версия: 1.0
 * Последнее обновление: 2026-03-01
 */

const CONFIG = {
  SHEET_NAME: 'Общая',
  ADDRESS_COL: 1,
  STATUS_HEADERS: [
    'Ожидаем товар',
    'Готов к поставке',
    'Доставка своими',
    'Сторонняя доставка',
    'Забрать документы'
  ],
  KML_FOLDER_NAME: 'MapSync',
  KML_FILE_NAME: 'delivery_map.kml',
  GEOCODE_CACHE_SHEET: 'GeoCache',
  ENABLE_LIVE_GEOCODING: false,
};


const NOMINATIM_LIMIT = {
  lastRequestTime: 0,
  minInterval: 2500
};

function waitForNominatimRateLimit() {
  const now = Date.now();
  const elapsed = now - NOMINATIM_LIMIT.lastRequestTime;
  const waitTime = Math.max(0, NOMINATIM_LIMIT.minInterval - elapsed);
  if (waitTime > 0) Utilities.sleep(waitTime);
  NOMINATIM_LIMIT.lastRequestTime = Date.now();
}

// Цвета маркеров для KML (формат: aabbggrr)
const STATUS_COLORS = {
  'Ожидаем товар': 'ff0000ff',
  'Готов к поставке': 'ff00ff00',
  'Доставка своими': 'ffffa500',
  'Сторонняя доставка': 'ff800080',
  'Забрать документы': 'ff00ffff'
};

/**
 * Главная функция синхронизации
 */
function syncDeliveryMap() {
  Logger.log('Запуск синхронизации: ' + new Date());

  try {
    const points = readSheetData();
    Logger.log('Прочитано строк: ' + points.length);

    if (points.length === 0) {
      Logger.log('Нет данных для обработки.');
      return;
    }

    const geocoded = geocodePoints(points);
    const success = geocoded.filter(p => p.coords).length;
    Logger.log('Геокодировано: ' + success + '/' + geocoded.length);

    const grouped = groupByStatus(geocoded);
    const kmlContent = generateKML(grouped);
    const fileUrl = saveKMLToFile(kmlContent);

    Logger.log('Синхронизация завершена. KML: ' + fileUrl);

  } catch (error) {
    Logger.log('Ошибка: ' + error.toString());
    MailApp.sendEmail(Session.getActiveUser().getEmail(), 'Ошибка синхронизации карты', error.toString());
    throw error;
  }
}

/**
 * Читает данные из Google Sheets
 */
function readSheetData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    const sheets = ss.getSheets();
    throw new Error('Вкладка "' + CONFIG.SHEET_NAME + '" не найдена! Доступные: ' + sheets.map(s => s.getName()).join(', '));
  }

  const data = sheet.getDataRange().getValues();

  if (data.length <= 1) {
    Logger.log('В таблице нет данных');
    return [];
  }

  const headers = data[0];
  const points = [];

  const cleanHeaders = headers.map(h => {
    if (!h) return '';
    return h.toString().trim().replace(/\s+/g, ' ');
  });

  Logger.log('Заголовки: ' + cleanHeaders.join(' | '));

  const statusIndices = {};
  for (const status of CONFIG.STATUS_HEADERS) {
    let foundIdx = cleanHeaders.indexOf(status);

    if (foundIdx < 0) {
      for (let i = 0; i < cleanHeaders.length; i++) {
        if (cleanHeaders[i].includes(status) || status.includes(cleanHeaders[i])) {
          foundIdx = i;
          break;
        }
      }
    }

    if (foundIdx >= 0) {
      statusIndices[status] = foundIdx;
      Logger.log('Статус "' + status + '" -> колонка ' + String.fromCharCode(65 + foundIdx) + ' (No' + (foundIdx + 1) + ')');
    }
  }

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const address = row[CONFIG.ADDRESS_COL - 1]?.toString().trim();

    if (!address || address === '' || address === 'null') continue;

    let status = null;
    for (const [statusName, colIndex] of Object.entries(statusIndices)) {
      const value = row[colIndex];
      const cleanValue = value ? value.toString().trim() : '';

      if (cleanValue !== '' && cleanValue !== '-' && cleanValue !== 'null' && value !== null) {
        status = statusName;
        Logger.log('Строка ' + (i + 1) + ': статус "' + status + '" активен (значение: ' + cleanValue + ')');
        break;
      }
    }

    if (status) {
      points.push({
        address: address,
        status: status,
        description: row.slice(0, headers.length).map((v, idx) => {
          const headerName = cleanHeaders[idx] || 'Кол. ' + (idx + 1);
          return headerName + ': ' + (v !== null && v !== '' ? v : '-');
        }).join('\n')
      });
    } else {
      Logger.log('Строка ' + (i + 1) + ' пропущена: все статусы "-" (адрес: ' + address + ')');
    }
  }

  return points;
}

/**
 * Геокодирует адреса через Nominatim (OpenStreetMap)
 */
function geocodePoints(points) {
  const cache = getGeoCache();
  const result = [];
  const errorLog = [];
  const DEFAULT_COORDS = { lat: 59.9343, lng: 30.3351 };

  const normalize = (s) => {
    if (!s) return '';
    return s.toString()
      .toLowerCase()
      .replace(/\s+/g, ' ')
      .replace(/\s*-\s*/g, '-')
      .trim();
  };

  for (const point of points) {
    const address = point.address;
    const normAddr = normalize(address);

    if (cache[normAddr]) {
      point.coords = cache[normAddr];
      Logger.log('Кэш: ' + address);
      result.push(point);
      continue;
    }

    let coords = null;
    let geocodingError = null;

    for (let attempt = 1; attempt <= 3; attempt++) {
      try {
        // Экспоненциальная задержка: 2.5с → 4.5с → 8.1с
        const delay = 2500 * Math.pow(1.8, attempt - 1);
        Utilities.sleep(delay);

        const query = address + ', Россия';
        const encodedQuery = encodeURIComponent(query);

        const url = 'https://nominatim.openstreetmap.org/search?format=json&q=' + encodedQuery + '&limit=1&addressdetails=0';

        waitForNominatimRateLimit();
        const response = UrlFetchApp.fetch(url, {
          'muteHttpExceptions': true,
          'headers': {
            'User-Agent': 'DeliveryMapSync/1.0 (mazohinav@gmail.com)',
            'Accept': 'application/json'
          }
        });

        const responseCode = response.getResponseCode();

        if (responseCode === 403 || responseCode === 429) {
          const backoff = 5000 * attempt;
          Logger.log('Nominatim блокирует (HTTP ' + responseCode + '), пауза 3 сек (попытка ' + attempt + '/3)');
          Utilities.sleep(backoff);
          continue;
        }

        if (responseCode !== 200) {
          Logger.log('HTTP ' + responseCode + ' (попытка ' + attempt + '/3)');
          if (attempt < 3) continue;
          geocodingError = 'HTTP ' + responseCode;
          break;
        }

        const contentText = response.getContentText().trim();

        if (contentText.startsWith('<') || contentText.startsWith('!')) {
          Logger.log('Nominatim вернул HTML (попытка ' + attempt + '/3)');
          if (attempt < 3) continue;
          geocodingError = 'HTML response';
          break;
        }

        let data;
        try {
          data = JSON.parse(contentText);
        } catch (e) {
          Logger.log('Ошибка парсинга JSON: ' + e.message + ' (попытка ' + attempt + '/3)');
          if (attempt < 3) continue;
          geocodingError = 'JSON parse error';
          break;
        }

        if (!Array.isArray(data) || data.length === 0 || !data[0].lat || !data[0].lon) {
          Logger.log('Не найдено в Nominatim: ' + address);
          geocodingError = 'Not found in Nominatim';
          break;
        }

        coords = { lat: parseFloat(data[0].lat), lng: parseFloat(data[0].lon) };
        Logger.log('Nominatim: ' + address + ' -> ' + coords.lat + ',' + coords.lng);
        break;

      } catch (e) {
        Logger.log('Ошибка сети (попытка ' + attempt + '/3): ' + address + ' - ' + e.message);
        if (attempt === 3) {
          geocodingError = 'Network error: ' + e.message;
        }
      }
    }

    // Обработка результата геокодинга
    if (coords) {
      point.coords = coords;
      point.geocoded = true;
      cache[normAddr] = coords;  // Сохраняем в кэш
      result.push(point);
    } else {
      // ❌ Адрес не геокодирован
      point.coords = null;
      point.geocoded = false;
      point.geocodingError = geocodingError;

      // Логируем ошибку для последующего анализа
      errorLog.push({
        address: address,
        status: point.status,
        error: geocodingError,
        timestamp: new Date()
      });

      Logger.log('⚠️ Не геокодирован: ' + address + ' (' + geocodingError + ')');

      // Варианты поведения:

      // ВАРИАНТ А: Не добавлять точку в результат вообще (строгий режим)
      // result.push(point);  // ← закомментируйте, если хотите скрывать такие точки

      // ВАРИАНТ Б: Добавить точку с null-координатами (фронтенд решит, как показать)
      //result.push(point);
    }
  }

  // Сохраняем кэш
  saveGeoCache(cache);

  // Логируем ошибки в отдельный лист
  // В консоли карты после загрузки:
  // console.log('Отрисовано:', window.lastMapData?.filter(p => p.coords)?.length);
  // console.log('Пропущено (нет координат):', window.lastMapData?.filter(p => !p.coords)?.length);
  if (errorLog.length > 0) {
    logGeocodingErrors(errorLog);
  }

  Logger.log('Геокодировано: ' + result.filter(p => p.geocoded).length + '/' + result.length);
  return result;
}


/**
 * Логирует ошибки геокодинга в отдельный лист
 */
function logGeocodingErrors(errors) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('GeocodingErrors');

  // Создаём лист, если нет
  if (!sheet) {
    sheet = ss.insertSheet('GeocodingErrors');
    sheet.appendRow(['Timestamp', 'Address', 'Status', 'Error']);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
    sheet.freezeRows(1);
  }

  // Добавляем новые ошибки
  const rows = errors.map(e => [
    e.timestamp,
    e.address,
    e.status,
    e.error
  ]);

  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
    Logger.log('📝 Записано ошибок в GeocodingErrors: ' + rows.length);
  }
}



/**
 * Загружает кэш геокодов из листа GeoCache
 */
function getGeoCache() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.GEOCODE_CACHE_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.GEOCODE_CACHE_SHEET);
    sheet.appendRow(['Адрес', 'lat', 'lng', 'Обновлено']);
    sheet.hideSheet();
    return {};
  }

  const data = sheet.getDataRange().getValues();
  const cache = {};

  for (let i = 1; i < data.length; i++) {
    const [address, lat, lng] = data[i];
    if (address && lat && lng) {
      cache[address.toString()] = { lat: lat, lng: lng };
    }
  }

  return cache;
}

/**
 * Сохраняет кэш геокодов в лист GeoCache
 */
function saveGeoCache(cache) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.GEOCODE_CACHE_SHEET);

  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 4).clearContent();
  }

  const rows = Object.entries(cache).map(([address, coords]) => [
    address,
    coords.lat,
    coords.lng,
    new Date()
  ]);

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 4).setValues(rows);
  }
}

/**
 * Группирует точки по статусам
 */
function groupByStatus(points) {
  const grouped = {};

  for (const status of CONFIG.STATUS_HEADERS) {
    grouped[status] = [];
  }

  for (const point of points) {
    if (point.coords && grouped[point.status]) {
      grouped[point.status].push(point);
    }
  }

  return grouped;
}

/**
 * Генерирует KML-файл из сгруппированных данных
 */
function generateKML(grouped) {
  const kmlHeader = '<?xml version="1.0" encoding="UTF-8"?>\n' +
    '<kml xmlns="http://www.opengis.net/kml/2.2">\n' +
    '<Document>\n' +
    '  <name>Delivery Map</name>\n' +
    '  <description>Автообновляемая карта доставки</description>\n';

  const kmlFooter = '</Document>\n</kml>';

  let layers = '';

  for (const [status, points] of Object.entries(grouped)) {
    if (points.length === 0) continue;

    const color = STATUS_COLORS[status] || 'ffffffff';

    layers += '\n  <Folder>\n' +
      '    <name>' + status + '</name>\n' +
      '    <Style id="style_' + status.replace(/\s+/g, '_') + '">\n' +
      '      <IconStyle>\n' +
      '        <color>' + color + '</color>\n' +
      '        <Icon><href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href></Icon>\n' +
      '      </IconStyle>\n' +
      '    </Style>\n';

    for (const point of points) {
      const safeDesc = point.description?.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;') || '';

      layers += '\n    <Placemark>\n' +
        '      <name>' + point.address + '</name>\n' +
        '      <description><![CDATA[' + safeDesc + ']]></description>\n' +
        '      <styleUrl>#style_' + status.replace(/\s+/g, '_') + '</styleUrl>\n' +
        '      <Point>\n' +
        '        <coordinates>' + point.coords.lng + ',' + point.coords.lat + ',0</coordinates>\n' +
        '      </Point>\n' +
        '    </Placemark>';
    }

    layers += '\n  </Folder>';
  }

  return kmlHeader + layers + kmlFooter;
}

/**
 * Сохраняет KML-файл на Google Диск
 */
function saveKMLToFile(kmlContent) {
  const folders = DriveApp.getFoldersByName(CONFIG.KML_FOLDER_NAME);
  let folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(CONFIG.KML_FOLDER_NAME);

  const files = folder.getFilesByName(CONFIG.KML_FILE_NAME);
  let file;

  if (files.hasNext()) {
    file = files.next();
    file.setContent(kmlContent);
    Logger.log('Обновлён файл: ' + file.getName());
  } else {
    file = folder.createFile(CONFIG.KML_FILE_NAME, kmlContent, 'application/vnd.google-earth.kml+xml');
    Logger.log('Создан файл: ' + file.getName());
  }

  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

/**
 * Создаёт триггер на запуск каждые 5 минут
 */
function create5MinTrigger() {
  // Удаляем старые триггеры syncDeliveryMap, чтобы не дублировать
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'syncDeliveryMap') {
      ScriptApp.deleteTrigger(t);
      Logger.log('🗑️ Удалён старый триггер');
    }
  });

  // Создаём новый: каждые 5 минут
  ScriptApp.newTrigger('syncDeliveryMap')
    .timeBased()
    .everyMinutes(5)
    .create();

  Logger.log('✅ Триггер создан: запуск каждые 5 минут');
  Logger.log('🔍 Проверить: Редактор скриптов → Триггеры (иконка ⏰ слева)');
}


/**
 * Web App: отдаёт данные для карты в формате JSON
 * Исправления: обработка всех статусов в строке, корректный URL геокодинга
 */
function getMapData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Общая');

    if (!sheet) {
      Logger.log('Лист "Общая" не найден');
      return [];
    }

    const data = sheet.getDataRange().getValues();

    if (data.length <= 1) {
      Logger.log('В таблице нет данных');
      return [];
    }

    const headers = data[0];

    const addressIdx = 0;
    const statusCols = {
      'Ожидаем товар': 2,
      'Готов к поставке': 3,
      'Доставка своими': 5,
      'Сторонняя доставка': 6,
      'Забрать документы': 8
    };

    const normalize = (s) => {
      if (!s) return '';
      return s.toString()
        .toLowerCase()
        .replace(/\s+/g, ' ')
        .replace(/\s*-\s*/g, '-')
        .trim();
    };


    const isAddress = (val) => {
      if (!val) return false;
      const v = val.toString().toLowerCase().trim();

      // 1. Типы улиц (полные + сокращённые) — работает для любого города РФ
      const streetTypes = /(проспект|просп\.|пр-кт|улица|ул\.|набережная|наб\.|переулок|пер\.|площадь|пл\.|шоссе|ш\.|бульвар|б-р|проезд|тупик|линия|аллея|квартал|мкр\.?|село|деревня|поселок|п\.?|станция|ст\.|дорога|дор\.)/i;

      // 2. Паттерн "название + номер дома" — универсален для любого города
      // Примеры: "Галстяна 1", "Ленина 30", "Тверская 15к2", "Москва 5"
      const hasNameAndNumber = /[а-яё]+\s+\d+/.test(v);

      // 3. Дополнительная проверка: отсечь ложные срабатывания ("офис 12", "этаж 3")
      const falsePositives = /^(офис|комната|этаж|подъезд|корпус|строение|литер|помещение|пом\.?)\s*\d+$/i;
      if (falsePositives.test(v)) return false;

      // 4. Минимальная длина для паттерна "название + номер" (защита от "а 1")
      const isShort = v.length < 5;

      return streetTypes.test(v) || (hasNameAndNumber && !isShort);
    };
    const cache = {};
    const cacheSheet = ss.getSheetByName('GeoCache');
    if (cacheSheet) {
      const cacheData = cacheSheet.getDataRange().getValues();
      for (let i = 1; i < cacheData.length; i++) {
        const addr = cacheData[i][0];
        const lat = cacheData[i][1];
        const lng = cacheData[i][2];
        if (addr && lat && lng) {
          const normKey = normalize(addr);
          cache[normKey] = { lat: parseFloat(lat), lng: parseFloat(lng) };
        }
      }
      Logger.log('Кэш загружен: ' + Object.keys(cache).length + ' координат');
    }

    const result = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const baseAddress = row[addressIdx]?.toString().trim();

      if (!baseAddress || baseAddress === '' || baseAddress === '-' ||
          baseAddress.toLowerCase().includes('http') ||
          baseAddress.toLowerCase().includes('адрес')) {
        continue;
      }

      // Обрабатываем ВСЕ активные статусы в строке (без break)
      for (const [statusName, colIdx] of Object.entries(statusCols)) {
        const val = row[colIdx];
        const cleanVal = val ? val.toString().trim() : '';

        // Если ячейка статуса заполнена — создаём точку
        if (cleanVal && cleanVal !== '-' && cleanVal !== '' && cleanVal !== 'null') {

          // Динамический адрес: если в ячейке похоже на адрес — используем его
          const pointAddress = isAddress(cleanVal) ? cleanVal : baseAddress;

          // Поиск координат в кэше
          const normAddr = normalize(pointAddress);
          const cached = cache[normAddr];
          let coords = cached;

          // Геокодинг "на лету" если не в кэше
          if (CONFIG.ENABLE_LIVE_GEOCODING && !coords) {
            try {
              const query = encodeURIComponent(pointAddress + ', Россия');
              const url = 'https://nominatim.openstreetmap.org/search?format=json&q=' + query + '&limit=1';

              waitForNominatimRateLimit();
              const resp = UrlFetchApp.fetch(url, {
                headers: { 'User-Agent': 'DeliveryMapSync/1.0 (mazohinav@gmail.com)' },
                muteHttpExceptions: true,
                timeout: 5000
              });

              if (resp.getResponseCode() === 200) {
                const json = JSON.parse(resp.getContentText());
                if (json[0] && json[0].lat && json[0].lon) {
                  coords = {
                    lat: parseFloat(json[0].lat),
                    lng: parseFloat(json[0].lon)
                  };
                  cache[normAddr] = coords;
                  saveGeoCache(cache);
                  Logger.log('Геокодирован на лету: ' + pointAddress);
                }
              }
            } catch(e) {
              Logger.log('Не удалось геокодировать на лету: ' + pointAddress + ' - ' + e.message);
            }
          }

          coords = coords || { lat: 55.7558, lng: 37.6173 };

          // Описание
          const description = row.slice(0, 5).map((v, idx) => {
            const headerName = headers[idx] ? headers[idx].toString().trim() : 'Кол.' + (idx + 1);
            return headerName + ': ' + (v !== null && v !== '' ? v : '-');
          }).join('<br>');

          // Добавляем точку в результат
          result.push({
            address: pointAddress,
            status: statusName,
            lat: parseFloat(coords.lat),
            lng: parseFloat(coords.lng),
            description: description
          });

          Logger.log('Строка ' + (i + 1) + ': ' + statusName + ' -> ' + pointAddress);
        }
      }
    }

    Logger.log('Возвращаем точек: ' + result.length);
    Logger.log('Примеры координат:');
    result.slice(0, 5).forEach(p => {
      const isDefault = (p.lat === 59.9343 && p.lng === 30.3351);
      Logger.log('   ' + p.address + ': ' + p.lat + ', ' + p.lng + ' ' + (isDefault ? 'ДЕФОЛТ' : 'ОК'));
    });

    return result;

  } catch (e) {
    Logger.log('Ошибка в getMapData: ' + e.toString() + '\n' + e.stack);
    return [];
  }
}

/**
 * Точка входа для Web App
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('MapPage')
    .setTitle('Карта Доставки')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Обновляет кэш геокодов для новых адресов
 */
function refreshGeoCache() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const cache = getGeoCache();
  let updated = 0;

  for (let i = 1; i < data.length; i++) {
    const address = data[i][0]?.toString().trim();
    if (!address || cache[address]) continue;

    Utilities.sleep(1100);
    try {
      const query = encodeURIComponent(address + ', Россия');
      const url = 'https://nominatim.openstreetmap.org/search?format=json&q=' + query + '&limit=1&addressdetails=0';

      waitForNominatimRateLimit();
      const resp = UrlFetchApp.fetch(url, {
        headers: { 'User-Agent': 'DeliveryMapSync/1.0' },
        muteHttpExceptions: true
      });
      const json = JSON.parse(resp.getContentText());
      if (json[0]) {
        cache[address] = { lat: parseFloat(json[0].lat), lng: parseFloat(json[0].lon) };
        updated++;
      }
    } catch (e) {
      // Игнорируем ошибки для отдельных адресов
    }
  }

  saveGeoCache(cache);
  SpreadsheetApp.getUi().alert('Обновлено координат: ' + updated);
}

/**
 * Диагностика: проверяет структуру таблицы
 */
function debugTableStructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('Таблица: ' + ss.getName());
  Logger.log('ID: ' + ss.getId());

  const sheets = ss.getSheets();
  Logger.log('Листы (' + sheets.length + '):');
  sheets.forEach(s => Logger.log('   - ' + s.getName() + ' (строк: ' + s.getLastRow() + ')'));

  const sheet = ss.getSheetByName('Общая');
  if (!sheet) {
    Logger.log('Лист "Общая" не найден!');
    return;
  }

  const data = sheet.getDataRange().getValues();
  Logger.log('Всего строк: ' + data.length);
  Logger.log('Всего колонок: ' + data[0].length);

  for (let i = 0; i < Math.min(5, data.length); i++) {
    Logger.log('Строка ' + (i + 1) + ': ' + data[i].slice(0, 5).join(' | '));
  }

  const headers = data[0];
  Logger.log('Поиск колонок:');
  Logger.log('   "Адрес поставки": #' + (headers.findIndex(h => h?.toString().includes('Адрес')) + 1));
  Logger.log('   "Ожидаем товар": #' + (headers.findIndex(h => h?.toString().includes('Ожидаем')) + 1));
  Logger.log('   "Готов к поставке": #' + (headers.findIndex(h => h?.toString().includes('Готов')) + 1));
  Logger.log('   "Доставка своими": #' + (headers.findIndex(h => h?.toString().includes('своими')) + 1));
  Logger.log('   "Сторонняя доставка": #' + (headers.findIndex(h => h?.toString().includes('Сторонняя')) + 1));
  Logger.log('   "Забрать документы": #' + (headers.findIndex(h => h?.toString().includes('Забрать')) + 1));

  let validRows = 0;
  for (let i = 1; i < data.length; i++) {
    const addr = data[i][0]?.toString().trim();
    if (addr && addr !== '-' && !addr.includes('http') && !addr.includes('Адрес')) {
      validRows++;
    }
  }
  Logger.log('Валидных строк с адресами: ' + validRows);
}

/**
 * Диагностика: информация о таблице и кэше
 */
function debugSpreadsheetInfo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('Таблица: ' + ss.getName());
  Logger.log('ID: ' + ss.getId());
  Logger.log('URL: ' + ss.getUrl());

  const cacheSheet = ss.getSheetByName('GeoCache');
  if (cacheSheet) {
    const data = cacheSheet.getDataRange().getValues();
    Logger.log('GeoCache строк: ' + (data.length - 1));
    for (let i = 1; i <= Math.min(3, data.length - 1); i++) {
      Logger.log('   ' + data[i][0] + ' -> ' + data[i][1] + ', ' + data[i][2]);
    }
  } else {
    Logger.log('Лист GeoCache не найден!');
  }
}