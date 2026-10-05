/**
 * Mini App «Заказ услуг» — серверная часть на Google Apps Script.
 * Принимает фото и заказы из Telegram Mini App, проверяет подпись Telegram,
 * пишет заказ одной строкой на лист «Заказы», складывает фото в папку заказа на Google Диске
 * и дублирует заказ в группу вместе с фото (оригиналами, файлами).
 * Прайс берётся из листа «Прайс» — цены меняются прямо в таблице.
 */

// Секреты НЕ хранятся в коде (репозиторий публичный). Они лежат в свойствах скрипта:
// Apps Script → ⚙️ Настройки проекта → Свойства скрипта: BOT_TOKEN, CHAT_ID, APP_LINK
const SECRETS = PropertiesService.getScriptProperties();

const CONFIG = {
  BOT_TOKEN: SECRETS.getProperty('BOT_TOKEN') || '',   // от @BotFather
  CHAT_ID: SECRETS.getProperty('CHAT_ID') || '',       // id супергруппы, начинается с -100
  APP_LINK: SECRETS.getProperty('APP_LINK') || '',     // ссылка на Mini App от BotFather (для postOrderButton)
  TIMEZONE: 'Europe/Minsk',
  CURRENCY: 'р',
  ORDERS_SHEET: 'Заказы',
  PRICE_SHEET: 'Прайс',
  PHOTO_FOLDER: 'Фото заказов Mini App', // папка на Google Диске
  MAX_PHOTO_MB: 20,                    // максимальный размер одного фото
  MAX_AGE_SEC: 24 * 60 * 60            // сколько живёт сессия Mini App
};

const APP_VERSION = 7;

// Лист «Заказы»: эти колонки, затем по колонке на каждую услугу (в ячейке — количество), в конце сумма заказа
// Данные для табличек (ФИО и даты) — отдельные колонки сразу после агента
const MEMORIAL_HEADERS = ['ФИО усопшего', 'Дата рождения', 'Дата смерти'];
const BASE_HEADERS = ['Дата', 'Время', '№ заказа', 'Username', 'Фамилия Имя агента'].concat(MEMORIAL_HEADERS);
const TOTAL_HEADER = 'Сумма заказа';

/** Для каких услуг нужны ФИО усопшего и даты: все таблички (сублимационная и с гравировкой, любые размеры). */
function isMemorialService_(name) {
  return String(name).toLowerCase().indexOf('табличк') >= 0;
}

/** К каким услугам можно прикрепить фото: вся печать фото (с рамкой и без) и сублимационная табличка. */
function isPhotoService_(name) {
  const n = String(name).trim().toLowerCase();
  return n.indexOf('фото') === 0 || n.indexOf('сублимац') >= 0;
}

const DEFAULT_PRICES = [
  ['Фото в рамке', 'Фото 20×30 в рамке', 15],
  ['Фото в рамке', 'Фото 18×24 в рамке', 13.5],
  ['Фото в рамке', 'Фото 15×21 в рамке', 12.5],
  ['Таблички', 'Табличка сублимационная', 27],
  ['Таблички', 'Табличка с лазерной гравировкой 13×18', 37],
  ['Таблички', 'Табличка с лазерной гравировкой 18×24', 54],
  ['Таблички', 'Табличка с лазерной гравировкой 13×18 (золото)', 47],
  ['Таблички', 'Табличка с лазерной гравировкой 18×24 (золото)', 70],
  ['Фото без рамки', 'Фото без рамки 15×21', 1.4],
  ['Фото без рамки', 'Фото без рамки 18×24', 1.7],
  ['Фото без рамки', 'Фото без рамки 20×30', 2.1],
  ['Обработка', 'Ламинирование фото', 3.2],
  ['Обработка', 'Ретушь 1-й категории', 10],
  ['Обработка', 'Ретушь 2-й категории', 20]
];

/* ---------- Разовая настройка: запусти из редактора (и после обновления кода тоже) ---------- */

function setup() {
  const ss = SpreadsheetApp.getActive();

  const prices = ss.getSheetByName(CONFIG.PRICE_SHEET) || ss.insertSheet(CONFIG.PRICE_SHEET);
  if (prices.getLastRow() === 0) {
    const head = ['ID', 'Категория', 'Услуга', 'Цена', 'Показывать'];
    const rows = DEFAULT_PRICES.map((p, i) => [i + 1, p[0], p[1], p[2], true]);
    prices.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    prices.getRange(2, 1, rows.length, head.length).setValues(rows);
    prices.getRange(2, 5, rows.length, 1).insertCheckboxes();
    prices.getRange(2, 4, rows.length, 1).setNumberFormat('0.00');
    prices.setFrozenRows(1);
    prices.autoResizeColumns(1, head.length);
  }

  const orders = ss.getSheetByName(CONFIG.ORDERS_SHEET) || ss.insertSheet(CONFIG.ORDERS_SHEET);
  if (orders.getLastRow() === 0) {
    const head = BASE_HEADERS.concat([TOTAL_HEADER]);
    orders.getRange(1, 1, 1, head.length).setValues([head])
      .setFontWeight('bold').setWrap(true).setVerticalAlignment('middle');
    orders.setFrozenRows(1);
    orders.setFrozenColumns(BASE_HEADERS.length);
    orders.setRowHeight(1, 60);
  }
  CacheService.getScriptCache().remove('catalog');
  ensureServiceColumns_(orders, getCatalog_().map(it => it.name));

  const root = photoRoot_(); // заодно запрашивает доступ к Google Диску
  incoming_();
  Logger.log('Готово. Листы на месте, папка для фото: %s', root.getUrl());
}

/**
 * Следит, чтобы у каждой услуги была своя колонка (перед «Сумма заказа»).
 * Новая услуга в прайсе — новая колонка появится сама при первом заказе или при запуске setup().
 */
function ensureServiceColumns_(sheet, names) {
  const headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0].map(String);

  // Базовые колонки, которых ещё нет (например, ФИО и даты на старом листе), вставляем на свои места
  let added = false;
  BASE_HEADERS.forEach((h, k) => {
    if (headers.indexOf(h) >= 0) return;
    const after = k === 0 ? 0 : headers.indexOf(BASE_HEADERS[k - 1]) + 1;
    sheet.insertColumnAfter(Math.max(after, 1));
    sheet.getRange(1, after + 1).setValue(h);
    sheet.setColumnWidth(after + 1, h === 'ФИО усопшего' ? 200 : 110);
    headers.splice(after, 0, h);
    added = true;
  });
  if (added) sheet.setFrozenColumns(BASE_HEADERS.length);

  let totalIdx = headers.indexOf(TOTAL_HEADER);
  if (totalIdx < 0) {
    throw new Error('На листе «' + CONFIG.ORDERS_SHEET + '» нет колонки «' + TOTAL_HEADER + '». Удалите лист и запустите setup().');
  }
  names.forEach(name => {
    if (headers.indexOf(name) >= 0) return;
    sheet.insertColumnBefore(totalIdx + 1);
    sheet.getRange(1, totalIdx + 1).setValue(name);
    sheet.setColumnWidth(totalIdx + 1, 120);
    headers.splice(totalIdx, 0, name);
    totalIdx++;
  });
  return headers;
}

/** Показывает в журнале id групп, где бот видел сообщения. */
function findGroupId() {
  checkToken_();
  const res = UrlFetchApp.fetch(tgUrl_('getUpdates'), { muteHttpExceptions: true });
  const data = JSON.parse(res.getContentText());
  if (!data.ok) throw new Error('Telegram: ' + data.description);
  const chats = {};
  (data.result || []).forEach(u => {
    const m = u.message || u.my_chat_member || u.edited_message || u.channel_post;
    if (m && m.chat && m.chat.type !== 'private') chats[m.chat.id] = m.chat.title;
  });
  const ids = Object.keys(chats);
  if (!ids.length) {
    Logger.log('Групп не найдено. Удали бота из группы и добавь снова (или напиши в группе /start@имя_бота), затем запусти функцию ещё раз.');
    return;
  }
  ids.forEach(id => Logger.log('%s → %s', chats[id], id));
}

/** Проверка связи: отправляет тестовое сообщение в группу. */
function testMessage() {
  sendToGroup_('✅ Бот подключён. Сюда будут приходить заказы.');
}

/* ---------- Веб-приложение ---------- */

function doGet() {
  try {
    return json_({ ok: true, version: APP_VERSION, items: getCatalog_(), currency: CONFIG.CURRENCY, maxPhotoMb: CONFIG.MAX_PHOTO_MB });
  } catch (err) {
    return json_({ ok: false, error: err.message });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const user = verifyInitData_(body.initData);
    if (body.action === 'upload') return json_(handleUpload_(body, user));
    return json_(handleOrder_(body, user));
  } catch (err) {
    return json_({ ok: false, error: err.message });
  }
}

/** Приём одного фото: кладём во «входящие» и помечаем, кто загрузил. */
function handleUpload_(body, user) {
  const mime = String(body.mime || '');
  if (mime.indexOf('image/') !== 0) throw new Error('Можно прикреплять только фото.');
  const bytes = Utilities.base64Decode(String(body.data || ''));
  if (!bytes.length) throw new Error('Файл пустой.');
  if (bytes.length > CONFIG.MAX_PHOTO_MB * 1024 * 1024) throw new Error('Фото больше ' + CONFIG.MAX_PHOTO_MB + ' МБ.');
  const name = String(body.name || 'photo.jpg').replace(/[\\/]/g, '_').slice(0, 120);
  const file = incoming_().createFile(Utilities.newBlob(bytes, mime, name));
  file.setDescription('tg:' + user.id);
  return { ok: true, fileId: file.getId() };
}

function handleOrder_(body, user) {
  // Защита от двойного заказа: повтор с тем же ключом вернёт уже оформленный заказ
  const cache = CacheService.getScriptCache();
  const key = body.orderKey ? 'order:' + user.id + ':' + String(body.orderKey).slice(0, 64) : null;
  if (key) {
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const seen = cache.get(key);
      if (seen === 'pending') throw new Error('Этот заказ ещё оформляется. Подождите минуту и проверьте группу.');
      if (seen) return JSON.parse(seen);
      cache.put(key, 'pending', 600);
    } finally {
      lock.releaseLock();
    }
  }
  try {
    const result = createOrder_(body, user);
    if (key) cache.put(key, JSON.stringify(result), 6 * 60 * 60);
    return result;
  } catch (err) {
    if (key) cache.remove(key);
    throw err;
  }
}

function createOrder_(body, user) {
  const agent = String(body.agent || '').replace(/\s+/g, ' ').trim().slice(0, 100);
  if (!agent) throw new Error('Укажите фамилию и имя агента.');

  // Цены берём только из таблицы — цену из приложения не принимаем
  const byId = {};
  getCatalog_().forEach(it => (byId[it.id] = it));
  const lines = (body.items || []).map(x => {
    const it = byId[String(x.id)];
    const qty = Math.floor(Number(x.qty));
    if (!it) throw new Error('Одной из услуг больше нет в прайсе. Откройте приложение заново.');
    if (!(qty > 0 && qty <= 10000)) throw new Error('Неверное количество для «' + it.name + '».');
    // Фото по позициям: photos[i] — фото для i-й штуки (или пусто)
    const photos = (Array.isArray(x.photos) ? x.photos : []).slice(0, qty).map(v => (v ? String(v) : null));
    if (photos.some(Boolean) && !it.photo) throw new Error('К услуге «' + it.name + '» фото не прикрепляются.');

    // Для табличек — ФИО и даты на каждую штуку
    let details = [];
    if (it.memorial) {
      const src = Array.isArray(x.details) ? x.details : [];
      for (let i = 0; i < qty; i++) {
        const d = src[i] || {};
        const fio = String(d.fio || '').replace(/\s+/g, ' ').trim().slice(0, 150);
        const born = String(d.born || '').trim(), died = String(d.died || '').trim();
        const label = '«' + it.name + '»' + (qty > 1 ? ', табличка ' + (i + 1) : '');
        if (!fio) throw new Error('Укажите ФИО усопшего: ' + label + '.');
        const b = parseDate_(born), dd = parseDate_(died);
        if (!b) throw new Error('Неверная дата рождения: ' + label + '. Формат дд.мм.гггг.');
        if (!dd) throw new Error('Неверная дата смерти: ' + label + '. Формат дд.мм.гггг.');
        if (dd < b) throw new Error('Дата смерти раньше даты рождения: ' + label + '.');
        details.push({ fio: fio, born: born, died: died });
      }
    }
    return { name: it.name, qty: qty, price: it.price, sum: round2_(qty * it.price), photos: photos, details: details };
  });
  if (!lines.length) throw new Error('Корзина пустая.');

  // Проверяем фото: загружены этим пользователем и ещё не ушли в другой заказ
  const files = {};
  const allIds = [].concat.apply([], lines.map(l => l.photos.filter(Boolean)));
  if (allIds.length) {
    if (new Set(allIds).size !== allIds.length) throw new Error('Одно и то же фото прикреплено дважды.');
    const incId = incoming_().getId();
    allIds.forEach(id => {
      let f;
      try { f = DriveApp.getFileById(id); } catch (e) { throw new Error('Фото не найдено на сервере. Прикрепите его заново.'); }
      const parents = f.getParents();
      if (f.getDescription() !== 'tg:' + user.id || !parents.hasNext() || parents.next().getId() !== incId) {
        throw new Error('Одно из фото уже использовано в другом заказе. Прикрепите его заново.');
      }
      files[id] = f;
    });
  }

  const username = user.username
    ? '@' + user.username
    : [user.first_name, user.last_name].filter(Boolean).join(' ') + ' (id ' + user.id + ')';

  const now = new Date();
  const date = Utilities.formatDate(now, CONFIG.TIMEZONE, 'dd.MM.yyyy');
  const time = Utilities.formatDate(now, CONFIG.TIMEZONE, 'HH:mm');
  const total = round2_(lines.reduce((s, l) => s + l.sum, 0));

  let orderNo, sheet, rowIdx, headers;
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const props = PropertiesService.getScriptProperties();
    orderNo = Number(props.getProperty('ORDER_NO') || 0) + 1;
    props.setProperty('ORDER_NO', String(orderNo));

    sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.ORDERS_SHEET);
    if (!sheet) throw new Error('Нет листа «' + CONFIG.ORDERS_SHEET + '». Запустите setup().');

    // Один заказ = одна строка: количество в колонке своей услуги, в конце сумма
    headers = ensureServiceColumns_(sheet, lines.map(l => l.name));
    const row = headers.map(() => '');
    const put = (h, v) => { row[headers.indexOf(h)] = v; };
    put('Дата', date);
    put('Время', time);
    put('№ заказа', orderNo);
    put('Username', username);
    put('Фамилия Имя агента', agent);
    const plates = [].concat.apply([], lines.map(l => l.details));
    if (plates.length) {
      const col = f => plates.map((d, i) => (plates.length > 1 ? (i + 1) + ') ' : '') + d[f]).join('\n');
      put('ФИО усопшего', col('fio'));
      put('Дата рождения', col('born'));
      put('Дата смерти', col('died'));
    }
    lines.forEach(l => {
      const i = headers.indexOf(l.name);
      row[i] = (Number(row[i]) || 0) + l.qty;
    });
    put(TOTAL_HEADER, total);

    rowIdx = sheet.getLastRow() + 1;
    sheet.getRange(rowIdx, 1, 1, 2).setNumberFormat('@');   // дата и время как текст
    const memCol = headers.indexOf(MEMORIAL_HEADERS[0]) + 1;
    sheet.getRange(rowIdx, memCol, 1, MEMORIAL_HEADERS.length).setNumberFormat('@').setWrap(true);
    sheet.getRange(rowIdx, 1, 1, row.length).setValues([row]);
  } finally {
    lock.releaseLock();
  }

  // Фото: папка заказа, имена файлов по услугам, ссылка на папку — в ячейке «№ заказа»
  if (allIds.length) {
    const folder = photoRoot_().createFolder('Заказ №' + orderNo + ' — ' + date + ' — ' + agent);
    lines.forEach(l => l.photos.forEach((id, i) => {
      if (!id) return;
      const f = files[id];
      const who = l.details[i] ? ' — ' + l.details[i].fio : '';
      const title = l.name + ' — ' + (i + 1) + ' из ' + l.qty + who;
      f.setName(title.replace(/[\\/:*?"<>|]/g, ' ') + ext_(f.getName()));
      f.setDescription('Заказ №' + orderNo + ', ' + title);
      f.moveTo(folder);
    }));
    const link = SpreadsheetApp.newRichTextValue().setText(String(orderNo)).setLinkUrl(folder.getUrl()).build();
    sheet.getRange(rowIdx, headers.indexOf('№ заказа') + 1).setRichTextValue(link);
  }

  let warning = null;
  try {
    const msg = sendToGroup_(buildMessage_(orderNo, date, time, username, agent, lines, total));
    if (allIds.length) sendPhotos_(lines, files, orderNo, msg.message_id);
  } catch (err) {
    console.error(err);
    warning = 'Заказ сохранён, но в группу отправилось не всё: ' + err.message;
  }

  return { ok: true, orderNo: orderNo, total: total, photos: allIds.length, warning: warning };
}

/* ---------- Ручная правка заказа в таблице ---------- */

/** Пересчёт суммы заказа при ручном изменении количества на листе «Заказы». */
function onEdit(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() === CONFIG.PRICE_SHEET) { CacheService.getScriptCache().remove('catalog'); return; }
  if (sheet.getName() !== CONFIG.ORDERS_SHEET) return;

  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const totalCol = headers.indexOf(TOTAL_HEADER) + 1;
  const firstSvc = BASE_HEADERS.length + 1;
  if (!totalCol) return;

  const r0 = Math.max(2, e.range.getRow());
  const r1 = e.range.getRow() + e.range.getNumRows() - 1;
  const c0 = e.range.getColumn(), c1 = c0 + e.range.getNumColumns() - 1;
  if (r1 < 2 || c1 < firstSvc || c0 >= totalCol) return;

  const prices = priceMap_();
  const missing = new Set();
  const rows = sheet.getRange(r0, 1, r1 - r0 + 1, lastCol).getValues();
  const totals = rows.map(row => {
    if (row[0] === '' && row[2] === '') return [row[totalCol - 1]];
    let sum = 0;
    for (let c = firstSvc; c < totalCol; c++) {
      const qty = Number(String(row[c - 1]).replace(',', '.')) || 0;
      if (!qty) continue;
      const name = headers[c - 1];
      if (prices[name] === undefined) { missing.add(name); continue; }
      sum += qty * prices[name];
    }
    return [round2_(sum)];
  });
  sheet.getRange(r0, totalCol, totals.length, 1).setValues(totals);

  if (missing.size) {
    SpreadsheetApp.getActive().toast('Нет цены в прайсе для: ' + Array.from(missing).join(', ') + '. Эти позиции не вошли в сумму.', 'Проверь сумму', 8);
  }
}

/** Цены всех услуг из «Прайса», включая скрытые. */
function priceMap_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.PRICE_SHEET);
  const map = {};
  if (!sheet || sheet.getLastRow() < 2) return map;
  sheet.getRange(2, 3, sheet.getLastRow() - 1, 2).getValues().forEach(([name, price]) => {
    const p = Number(String(price).replace(',', '.'));
    if (String(name).trim() && isFinite(p)) map[String(name).trim()] = p;
  });
  return map;
}

/* ---------- Кнопка заказа в группе ---------- */

/** Отправляет в группу сообщение с кнопкой заказа и закрепляет его. */
function postOrderButton() {
  const msg = tgJson_('sendMessage', {
    chat_id: CONFIG.CHAT_ID,
    text: 'Оформить заказ на услуги 👇',
    reply_markup: { inline_keyboard: [[{ text: '🛒 Оформить заказ', url: CONFIG.APP_LINK }]] }
  });
  try {
    tgJson_('pinChatMessage', { chat_id: CONFIG.CHAT_ID, message_id: msg.message_id, disable_notification: true });
    Logger.log('Сообщение отправлено и закреплено.');
  } catch (err) {
    Logger.log('Отправлено, но не закреплено: %s. Закрепи вручную.', err.message);
  }
}

/* ---------- Вспомогательное ---------- */

/** Прайс из кэша (мгновенно); кэш сбрасывается при любой правке листа «Прайс». */
function getCatalog_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('catalog');
  if (hit) return JSON.parse(hit);
  const list = readCatalog_();
  cache.put('catalog', JSON.stringify(list), 6 * 60 * 60);
  return list;
}

function readCatalog_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.PRICE_SHEET);
  if (!sheet) throw new Error('Нет листа «' + CONFIG.PRICE_SHEET + '». Запустите setup().');
  const last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, 5).getValues()
    .filter(r => r[0] !== '' && String(r[2]).trim() && r[4] !== false)
    .map(r => ({
      id: String(r[0]),
      category: String(r[1] || 'Услуги').trim(),
      name: String(r[2]).trim(),
      price: Number(String(r[3]).replace(',', '.')),
      photo: isPhotoService_(r[2]),
      memorial: isMemorialService_(r[2])
    }))
    .filter(it => isFinite(it.price) && it.price >= 0);
}

function verifyInitData_(initData) {
  checkToken_();
  if (!initData) throw new Error('Откройте приложение через ссылку в Telegram.');

  const params = {};
  String(initData).split('&').forEach(pair => {
    const i = pair.indexOf('=');
    if (i > 0) params[pair.slice(0, i)] = decodeURIComponent(pair.slice(i + 1).replace(/\+/g, ' '));
  });
  const hash = params.hash;
  delete params.hash;

  const checkString = Object.keys(params).sort().map(k => k + '=' + params[k]).join('\n');
  const secret = Utilities.computeHmacSha256Signature(
    Utilities.newBlob(CONFIG.BOT_TOKEN).getBytes(),
    Utilities.newBlob('WebAppData').getBytes()
  );
  const sign = Utilities.computeHmacSha256Signature(Utilities.newBlob(checkString).getBytes(), secret)
    .map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');

  if (!hash || sign !== hash) throw new Error('Не удалось подтвердить вход через Telegram. Откройте приложение заново.');
  if (Date.now() / 1000 - Number(params.auth_date) > CONFIG.MAX_AGE_SEC) {
    throw new Error('Сессия устарела. Закройте и откройте приложение заново.');
  }
  const user = JSON.parse(params.user || 'null');
  if (!user || !user.id) throw new Error('Telegram не передал данные пользователя.');
  return user;
}

function buildMessage_(orderNo, date, time, username, agent, lines, total) {
  const out = [
    '<b>🧾 Новый заказ №' + orderNo + '</b>',
    'Дата: ' + date,
    'Время: ' + time,
    'Пользователь: ' + esc_(username),
    'Агент: ' + esc_(agent)
  ];
  lines.forEach(l => {
    out.push('', 'Услуга: ' + esc_(l.name), 'Количество: ' + l.qty, 'Цена: ' + money_(l.price));
    if (l.qty > 1) out.push('Сумма: ' + money_(l.sum));
    const n = l.photos.filter(Boolean).length;
    if (n) out.push('Фото: ' + n + ' шт. (ниже)');
    l.details.forEach((d, i) => {
      if (l.details.length > 1) out.push('<b>Табличка ' + (i + 1) + ':</b>');
      out.push('ФИО усопшего: ' + esc_(d.fio), 'Дата рождения: ' + d.born, 'Дата смерти: ' + d.died);
    });
  });
  out.push('', '<b>Итого: ' + money_(total) + '</b>');
  return out.join('\n');
}

/** Фото уходят в группу файлами (без сжатия), ответом на сообщение заказа, с подписью услуги. */
function sendPhotos_(lines, files, orderNo, replyTo) {
  const items = [];
  lines.forEach(l => l.photos.forEach((id, i) => {
    if (!id) return;
    items.push({
      file: files[id],
      caption: 'Заказ №' + orderNo + '\n' + l.name + ' — ' + (i + 1) + ' из ' + l.qty + (l.details[i] ? '\n' + l.details[i].fio : ''),
      fname: 'zakaz' + orderNo + '_' + (items.length + 1) + ext_(files[id].getName())
    });
  }));

  const LIMIT = 45 * 1024 * 1024;
  let batch = [], size = 0;
  const flush = () => { if (batch.length) sendDocs_(batch, replyTo); batch = []; size = 0; };
  items.forEach(it => {
    const s = it.file.getSize();
    if (batch.length === 10 || (batch.length && size + s > LIMIT)) flush();
    batch.push(it);
    size += s;
  });
  flush();
}

function sendDocs_(batch, replyTo) {
  const reply = JSON.stringify({ message_id: replyTo, allow_sending_without_reply: true });
  const blob = it => it.file.getBlob().setName(it.fname);
  if (batch.length === 1) {
    tgMultipart_('sendDocument', { chat_id: String(CONFIG.CHAT_ID), caption: batch[0].caption, reply_parameters: reply, document: blob(batch[0]) });
    return;
  }
  const payload = {
    chat_id: String(CONFIG.CHAT_ID),
    reply_parameters: reply,
    media: JSON.stringify(batch.map((it, i) => ({ type: 'document', media: 'attach://f' + i, caption: it.caption })))
  };
  batch.forEach((it, i) => (payload['f' + i] = blob(it)));
  tgMultipart_('sendMediaGroup', payload);
}

function sendToGroup_(text) {
  return tgJson_('sendMessage', { chat_id: CONFIG.CHAT_ID, text: text, parse_mode: 'HTML', disable_web_page_preview: true });
}

function tgJson_(method, data) {
  checkToken_();
  const res = UrlFetchApp.fetch(tgUrl_(method), {
    method: 'post', contentType: 'application/json', payload: JSON.stringify(data), muteHttpExceptions: true
  });
  const d = JSON.parse(res.getContentText());
  if (!d.ok) throw new Error('Telegram: ' + d.description);
  return d.result;
}

function tgMultipart_(method, payload) {
  checkToken_();
  const res = UrlFetchApp.fetch(tgUrl_(method), { method: 'post', payload: payload, muteHttpExceptions: true });
  const d = JSON.parse(res.getContentText());
  if (!d.ok) throw new Error('Telegram: ' + d.description);
  return d.result;
}

function photoRoot_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('PHOTO_ROOT');
  if (id) {
    try { const f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) {}
  }
  const folder = DriveApp.createFolder(CONFIG.PHOTO_FOLDER);
  props.setProperty('PHOTO_ROOT', folder.getId());
  return folder;
}

function incoming_() {
  const root = photoRoot_();
  const it = root.getFoldersByName('_входящие');
  return it.hasNext() ? it.next() : root.createFolder('_входящие');
}

/** Удаляет из «_входящих» фото старше 2 дней (загрузили, но заказ не оформили). */
function cleanupIncoming() {
  const border = Date.now() - 2 * 24 * 60 * 60 * 1000;
  const files = incoming_().getFiles();
  let n = 0;
  while (files.hasNext()) {
    const f = files.next();
    if (f.getDateCreated().getTime() < border) { f.setTrashed(true); n++; }
  }
  Logger.log('Убрано в корзину: %s', n);
}

function checkToken_() {
  if (!CONFIG.BOT_TOKEN) throw new Error('Нет BOT_TOKEN в свойствах скрипта (⚙️ Настройки проекта → Свойства скрипта).');
  if (!CONFIG.CHAT_ID) throw new Error('Нет CHAT_ID в свойствах скрипта (⚙️ Настройки проекта → Свойства скрипта).');
}

/** 'дд.мм.гггг' → Date или null. */
function parseDate_(s) {
  const m = String(s).match(/^(\d{2})\.(\d{2})\.(\d{4})$/);
  if (!m) return null;
  const d = new Date(+m[3], +m[2] - 1, +m[1]);
  if (d.getFullYear() !== +m[3] || d.getMonth() !== +m[2] - 1 || d.getDate() !== +m[1]) return null;
  if (+m[3] < 1800 || d.getTime() > Date.now() + 86400000) return null;
  return d;
}

function ext_(name) { const m = String(name).match(/\.[a-z0-9]{2,5}$/i); return m ? m[0].toLowerCase() : '.jpg'; }
function tgUrl_(method) { return 'https://api.telegram.org/bot' + CONFIG.BOT_TOKEN + '/' + method; }
function round2_(n) { return Math.round(n * 100) / 100; }
function num_(n) { return round2_(n).toFixed(2).replace(/\.00$/, '').replace('.', ','); }
function money_(n) { return num_(n) + ' ' + CONFIG.CURRENCY; }
function esc_(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
function json_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
