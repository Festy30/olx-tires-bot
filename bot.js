require('dotenv').config();
const { Telegraf, Markup } = require('telegraf');
const XLSX = require('xlsx');
const axios = require('axios');
const AdmZip = require('adm-zip');

if (!process.env.BOT_TOKEN) {
  console.error('Не задано BOT_TOKEN у .env');
  process.exit(1);
}

const bot = new Telegraf(process.env.BOT_TOKEN);

const groupSessions = {};
const mediaGroupBuffers = {};

const inlineMenu = Markup.inlineKeyboard([
  [
    Markup.button.callback('📥 Експорт в Excel', 'action_export'),
    Markup.button.callback('🗑 Очистити список', 'action_clear')
  ]
]);

// Повідомлення, старші за цю дату, ігноруються
const FILTER_DATE = new Date('2026-09-30T00:00:00Z');

// Налаштування CSV для імпорту в OLX
const CSV_DELIMITER = ',';  // якщо OLX вимагатиме крапку з комою — змініть на ';'
const CSV_BOM = true;       // UTF-8 з BOM; якщо OLX не приймає — поставте false

// Telegram дозволяє боту надсилати файли до 50 МБ — беремо з запасом
const MAX_ZIP_BYTES = 45 * 1024 * 1024;

// Роздільник імен файлів у колонці images.
// За документацією OLX Pro імена фото розділяються вертикальною рискою «|»
// (кома й крапка з комою не підходять). Пробілів у роздільнику бути не повинно.
const IMAGES_DELIMITER = '|';

// Коди індексу швидкості OLX (літера з шини -> код у CSV).
// Літери, яких тут немає, залишаються порожніми, щоб імпорт не відхилявся.
const SPEED_INDEX_CODES = {
  K: 'k', L: 'l', M: 'm', N: 'n', P: 'p', Q: 'q', R: 'r', S: 's', T: 't', U: 'u',
  H: 'v_210',
  V: 'v_240',
  W: 'w',
  Y: 'y'
};

// Номери телефонів для кожного оголошення (колонка phone в CSV)
const PHONE_NUMBERS = '050-057-40-10, 050-86-05-888';

function getSession(chatId) {
  if (!groupSessions[chatId]) {
    groupSessions[chatId] = { items: [], messageIds: [] };
  }
  return groupSessions[chatId];
}

// Одне «вікно» зі списком усіх доданих товарів і кнопками вивантаження.
// При кожному новому товарі старе повідомлення видаляється, а нове з оновленим списком
// з'являється внизу чату. Виклики ставляться в чергу, щоб при кількох одночасних
// постах не лишалось зайвих повідомлень.
const SUMMARY_MAX_LINES = 15;

// Через скільки мілісекунд видаляти файли експорту (CSV і ZIP) з чату
const EXPORT_DELETE_DELAY_MS = 10 * 1000;

function showSummary(ctx, chatId) {
  const session = getSession(chatId);

  session.summaryQueue = (session.summaryQueue || Promise.resolve())
    .then(async () => {
      if (session.summaryMsgId) {
        try {
          await ctx.telegram.deleteMessage(chatId, session.summaryMsgId);
        } catch (err) {}
        session.summaryMsgId = null;
      }

      const items = session.items;
      if (items.length === 0) return;

      const shown = items.slice(-SUMMARY_MAX_LINES);
      const hidden = items.length - shown.length;
      const lines = shown.map((item, i) => {
        const photos = item._photos ? item._photos.length : 0;
        return `${hidden + i + 1}. ${item.title} — ${item.price_value} грн` +
          (photos ? ` (🖼 ${photos})` : '');
      });

      const text =
        `Додано товарів: ${items.length}\n\n` +
        (hidden > 0 ? `… і ще ${hidden} раніше\n` : '') +
        lines.join('\n');

      const sent = await ctx.reply(text, inlineMenu);
      session.summaryMsgId = sent.message_id;
      console.log(`Додано товар #${items.length}: ${items[items.length - 1].title}`);
    })
    .catch(err => console.error('Помилка оновлення списку:', err));

  return session.summaryQueue;
}

async function sendTempMessage(ctx, text, delayMs = 5000) {
  try {
    const sentMsg = await ctx.reply(text);
    setTimeout(async () => {
      try {
        await ctx.deleteMessage(sentMsg.message_id);
      } catch (err) {}
    }, delayMs);
  } catch (err) {
    console.error('Помилка відправки/видалення:', err);
  }
}

// Завантажує фото з Telegram у пам'ять (Buffer). При помилці повертає null.
async function downloadTelegramPhoto(fileId) {
  try {
    const fileLink = await bot.telegram.getFileLink(fileId);
    const fileUrl = typeof fileLink === 'string' ? fileLink : fileLink.href;

    const response = await axios.get(fileUrl, {
      responseType: 'arraybuffer',
      timeout: 20000
    });
    return Buffer.from(response.data);
  } catch (err) {
    console.error('Помилка завантаження фото з Telegram:', err?.response?.data || err.message);
    return null;
  }
}

// external_id використовується в іменах файлів, тому він має бути унікальним
function ensureUniqueExternalId(session, item) {
  const used = new Set(session.items.map(i => i.external_id));
  while (used.has(item.external_id)) {
    item.external_id = `TIRE-${Math.floor(10000 + Math.random() * 90000)}`;
  }
}

// Присвоює фото імена файлів, записує їх у поле images (для CSV)
// і зберігає самі файли в _photos (для ZIP)
function attachPhotos(item, buffers) {
  const photos = buffers
    .filter(Boolean)
    .map((buffer, i) => ({ name: `${item.external_id}_${i + 1}.jpg`, buffer }));
  item._photos = photos;
  item.images = photos.map(p => p.name).join(IMAGES_DELIMITER);
  return photos.length;
}

async function clearCollectedMessages(ctx, chatId) {
  const session = groupSessions[chatId];
  if (!session || !session.messageIds || session.messageIds.length === 0) return;

  const ids = session.messageIds;
  session.messageIds = [];

  // deleteMessages приймає не більше 100 id за раз
  for (let i = 0; i < ids.length; i += 100) {
    const chunk = ids.slice(i, i + 100);
    let bulkOk = false;

    try {
      if (ctx.telegram.deleteMessages) {
        await ctx.telegram.deleteMessages(chatId, chunk);
        bulkOk = true;
      }
    } catch (err) {
      // переходимо до видалення по одному
    }

    if (!bulkOk) {
      for (const msgId of chunk) {
        try {
          await ctx.telegram.deleteMessage(chatId, msgId);
        } catch (e) {}
      }
    }
  }
}

bot.start((ctx) => {
  getSession(ctx.chat.id);

  ctx.reply(
    'Вітаю! Я збираю пости з шинами, завантажую фото та формую CSV і ZIP-архів для імпорту в OLX.\n\n' +
    'Використовуйте кнопки нижче для керування:',
    inlineMenu
  );
});

const DESCRIPTION_FOOTER = `Оплата відбувається на р/р фоп при умові попереднього внесення авансу у розмірі 250 грн за одне колесо.

відправка щодня в день замовлення не пізніше 17:30 або ж наступного дня і тільки новою поштою, так як склад шин знаходиться за містом 15 км від центру міста Луцьк.

роки виготовлення а саме DOT усіх шин рідний тобто оригінальний а не перебитий. це ми гарантуємо.

займаємось саме стоковими шинами переважно тому звертайте увагу на рік виготовлення. шини зберігались до купівлі в сухому приміщення з відповідним рівнем вологості без доступу прямого сонячного світла UV.

телефонуйте або пишіть нам на вайбер 050-057-40-10. з радістю підберемо потрібні для вас шини чи диски. опрацьовуємо індивідуальні замовлення теж при можливості і наявності шин.

Наша адреса: м. Луцьк, вул. Набережна, 69 (вишківське перехрестя)

Продаж і заміна шин за пільговою ціною, послуги шиномонтажу, рихтування дисків, ремонт бічних пошкоджень (в сезон обмежена послуга)

Запис на шиномонтаж на Набережну по телефону 066-087-80-88.

Склад-магазин, шиномонтаж і гуртовий відділ відправок шин і дисків відбувається виключно Новою Поштою і знаходиться за адресою:

Луцький район, с. Піддубці по трасі на Рівне на відстані 7км від Луцька (Епіцентр,кільце).
Зліва над дорогою великий біг-борд в кінці села навпроти відділення Нової Пошти.

Тел. 050-057-40-10, VIber, Whatsup, Telegram.

Звичайний телефон на шиномонтаж і запис на нього 050-86-05-888.`;

function parseTirePost(text) {
  if (!text) return null;

  const sizeMatch = text.match(/(\d{3})\/(\d{2})\s*R?(\d{2})/i);
  const priceMatch = text.match(/Ціна\s*—?\s*(\d+)/i) || text.match(/(\d+)\s*грн/i);
  const qtyMatch = text.match(/в\s*наявності\s*(\d+)\s*шт/i) || text.match(/(\d+)\s*шт/i);
  const yearMatch = text.match(/(?:Рік|DOT)?\s*:?\s*(20\d{2})/i);
  const loadSpeedMatch = text.match(/\b(\d{2,3}[A-Z])\b/);

  if (!sizeMatch && !priceMatch) return null;

  const lines = text.split('\n').map(l => l.trim()).filter(Boolean);
  const firstLine = lines[0] || '';

  const width = sizeMatch ? sizeMatch[1] : '';
  const height = sizeMatch ? sizeMatch[2] : '';
  const diameter = sizeMatch ? sizeMatch[3] : '';
  const sizeStr = sizeMatch ? `${width}/${height}R${diameter}` : '';

  // Якщо сезон у пості не вказано — вважаємо шини зимовими
  let season = 'Зимові';
  let seasonTitle = 'зимові';
  if (/всесезон|all[\s-]?season/i.test(text)) {
    season = 'Всесезонні';
    seasonTitle = 'всесезонні';
  } else if (/літні|літо|summer/i.test(text)) {
    season = 'Літні';
    seasonTitle = 'літні';
  }

  let state = 'Б/в';
  let stateTitle = 'б/в';
  if (/нові|новые|новий/i.test(text)) {
    state = 'Нове';
    stateTitle = 'нові';
  }

  const price = priceMatch ? priceMatch[1] : '';
  // Кількість в оголошенні: більше 4 → 4 шт.; 2 або 3 → 2 шт.; 1 → 1 шт.
  const totalQty = (qtyMatch ? parseInt(qtyMatch[1], 10) : 2) || 2;
  const displayQty = totalQty >= 4 ? 4 : (totalQty >= 2 ? 2 : 1);
  const quantity = String(displayQty);
  // Код для поля quantity в OLX
  const quantityCode = displayQty === 4 ? 'set_of_4' : (displayQty === 2 ? 'pair' : '1_tire_unpaired');
  const year = yearMatch ? yearMatch[1] : '';

  let rawBrandModel = firstLine;
  if (sizeMatch) {
    rawBrandModel = rawBrandModel.replace(sizeMatch[0], '');
  }
  if (loadSpeedMatch) {
    rawBrandModel = rawBrandModel.replace(loadSpeedMatch[0], '');
  }
  rawBrandModel = rawBrandModel
    .replace(/\bXL\b/gi, '')
    .replace(/[()]/g, ' ')   // дужки навколо індексу (99H) не потрібні
    .replace(/\s+/g, ' ')
    .trim();

  const brandListMatch = firstLine.match(/Hankook|Michelin|Continental|Bridgestone|Pirelli|Nokian|Goodyear|Dunlop|Yokohama|Kumho|Toyo|Sava|Lassa|Barum|Nexen|Tigar|Gislaved/i);
  const manufacturer = brandListMatch ? brandListMatch[0] : (rawBrandModel.split(' ')[0] || '');

  const isXl = /\bXL\b|Extra Load|посилені|усиленные/i.test(text);
  const xlText = isXl ? 'XL' : '';

  const loadSpeedFull = loadSpeedMatch ? loadSpeedMatch[1] : '';
  const speedIndex = loadSpeedMatch ? loadSpeedMatch[1].slice(-1) : '';

  const titleParts = [
    `[${quantity}шт]`,
    stateTitle,
    seasonTitle,
    'шини',
    sizeStr,
    rawBrandModel,
    xlText,
    loadSpeedFull
  ].filter(Boolean);

  const formattedTitle = titleParts.join(' ').replace(/\s+/g, ' ').trim();

  // Короткий заголовок для першого рядка опису: розмір, бренд/модель, XL, індекс
  const descTitle = [sizeStr, rawBrandModel, xlText, loadSpeedFull]
    .filter(Boolean)
    .join(' ')
    .replace(/\s+/g, ' ')
    .trim();

  const stateText = state === 'Нове' ? 'Нові' : 'Б/в';
  const madeInMatch = text.match(/Made\s+in\s+([A-Za-z]+(?:\s[A-Za-z]+)?)/i);

  const descriptionLines = [
    descTitle,
    '',
    `${stateText} ${seasonTitle} шини`,
    year ? `Рік: ${year}` : null,
    `В наявності ${quantity} шт.`,
    'Гарантія на шиномонтаж.',
    '',
    price ? `Ціна — ${price} грн за штуку` : null,
    madeInMatch ? `Made in ${madeInMatch[1].trim()}` : null,
    '',
    DESCRIPTION_FOOTER
  ].filter(line => line !== null);

  const formattedDescription = descriptionLines.join('\n');

  const externalId = `TIRE-${Math.floor(10000 + Math.random() * 90000)}`;
  const uniqueKey = `${sizeStr}_${firstLine}_${price}_${season}`.toLowerCase().replace(/\s+/g, '');

  return {
    id: '',
    title: formattedTitle,
    description: formattedDescription,
    external_id: externalId,
    location_city: '72',
    location_district: '',
    phone: PHONE_NUMBERS,
    images: '',
    price_value: price,
    price_currency: 'UAH',
    state: state === 'Нове' ? 'new' : 'used',
    season: season === 'Літні' ? '1' : (season === 'Зимові' ? '2' : '3'),
    width: width,
    height: height,
    diameter30: diameter,
    function: '1',
    quantity: quantityCode,
    tires_manufacturer: brandListMatch ? brandListMatch[0].toLowerCase() : '',
    speed_index: SPEED_INDEX_CODES[speedIndex.toUpperCase()] || '',
    tires_special: isXl ? 'Посилені (XL)' : '',
    year_made: year,
    _uniqueKey: uniqueKey
  };
}

bot.action('action_export', async (ctx) => {
  try {
    await ctx.answerCbQuery();
  } catch (err) {}

  const chatId = ctx.chat?.id;
  const session = groupSessions[chatId] || { items: [], messageIds: [] };
  const items = session.items || [];
  console.log(`Натиснуто «Експорт»: товарів у списку ${items.length}`);

  if (items.length === 0) {
    return sendTempMessage(ctx, '❌ Немає збережених товарів для експорту.');
  }

  // _uniqueKey і _photos — службові поля, у CSV вони не потрібні
  const exportData = items.map(({ _uniqueKey, _photos, ...item }) => item);

  const worksheet = XLSX.utils.json_to_sheet(exportData);
  const csv = XLSX.utils.sheet_to_csv(worksheet, { FS: CSV_DELIMITER, RS: '\n' });

  // Явно кодуємо в UTF-8 (з BOM або без, залежно від CSV_BOM)
  const csvBuffer = Buffer.from((CSV_BOM ? '\uFEFF' : '') + csv, 'utf8');

  // ZIP з фото; імена файлів збігаються з тими, що записані в колонці images
  const allPhotos = items.flatMap(item => item._photos || []);
  let zipBuffer = null;
  if (allPhotos.length > 0) {
    const zip = new AdmZip();
    for (const photo of allPhotos) {
      zip.addFile(photo.name, photo.buffer);
    }
    zipBuffer = zip.toBuffer();

    if (zipBuffer.length > MAX_ZIP_BYTES) {
      const mb = (zipBuffer.length / 1024 / 1024).toFixed(1);
      return ctx.reply(
        `❌ ZIP з фото завеликий (${mb} МБ), Telegram не дозволяє боту надсилати файли понад 50 МБ.\n` +
        'Експортуйте товари меншими партіями: список не очищено.'
      );
    }
  }

  const stamp = Date.now();
  const sentMessageIds = [];
  console.log(`Експорт: товарів ${items.length}, фото ${allPhotos.length}, ZIP ${zipBuffer ? (zipBuffer.length / 1024 / 1024).toFixed(1) : 0} МБ`);

  // Спочатку надсилаємо файли — і лише після успіху чистимо чат і список
  try {
    const csvMsg = await ctx.replyWithDocument({
      source: csvBuffer,
      filename: `tires_export_${stamp}.csv`
    }, {
      caption: `📊 CSV-файл у UTF-8 (товарів: ${exportData.length}).` +
        (zipBuffer ? '' : '\n\n🧹 Оголошення та список очищено!')
    });
    sentMessageIds.push(csvMsg.message_id);

    if (zipBuffer) {
      const zipMsg = await ctx.replyWithDocument({
        source: zipBuffer,
        filename: `tires_photos_${stamp}.zip`
      }, {
        caption: `🖼 ZIP з фото (файлів: ${allPhotos.length}). Завантажте його в OLX у поле «Фото» разом із CSV.\n\n🧹 Оголошення та список очищено!`
      });
      sentMessageIds.push(zipMsg.message_id);
    }
  } catch (err) {
    console.error('Помилка відправки файлів експорту:', err);
    try {
      await ctx.reply(
        `❌ Не вдалося надіслати файли: ${err?.description || err?.message || 'невідома помилка'}\n\n` +
        'Список товарів збережено. Спробуйте експорт ще раз.'
      );
    } catch (e) {}
    return;
  }

  console.log('Файли експорту надіслано');

  // Файли з експортом видаляються з чату через EXPORT_DELETE_DELAY_MS
  setTimeout(async () => {
    for (const msgId of sentMessageIds) {
      try {
        await ctx.telegram.deleteMessage(chatId, msgId);
      } catch (err) {}
    }
  }, EXPORT_DELETE_DELAY_MS);

  try {
    await ctx.deleteMessage();
  } catch (err) {}

  await clearCollectedMessages(ctx, chatId);

  groupSessions[chatId] = { items: [], messageIds: [] };
});

bot.action('action_clear', async (ctx) => {
  try {
    await ctx.answerCbQuery();
  } catch (err) {}

  const chatId = ctx.chat?.id;

  try {
    await ctx.deleteMessage();
  } catch (err) {}

  await clearCollectedMessages(ctx, chatId);

  groupSessions[chatId] = { items: [], messageIds: [] };

  await sendTempMessage(ctx, '🧹 Список товарів та оголошення очищено!');
});

bot.on(['message', 'channel_post'], async (ctx) => {
  const msg = ctx.message || ctx.channel_post;
  if (!msg) return;

  const chatId = ctx.chat.id;
  const session = getSession(chatId);

  if (msg.message_id) {
    if (msg.media_group_id) {
      for (let i = -9; i <= 0; i++) {
        const targetId = msg.message_id + i;
        if (targetId > 0 && !session.messageIds.includes(targetId)) {
          session.messageIds.push(targetId);
        }
      }
    } else if (!session.messageIds.includes(msg.message_id)) {
      session.messageIds.push(msg.message_id);
    }
  }

  let photoFileId = null;
  if (msg.photo && msg.photo.length > 0) {
    photoFileId = msg.photo[msg.photo.length - 1].file_id;
  }

  const text = msg.text || msg.caption || '';
  const msgDate = new Date((msg.date || 0) * 1000);

  if (msgDate < FILTER_DATE) {
    console.log(`Пропущено повідомлення ${msg.message_id}: старше за FILTER_DATE (${msgDate.toISOString()})`);
    return;
  }

  // ---- Альбом (кілька фото) ----
  if (msg.media_group_id) {
    const mgId = msg.media_group_id;

    if (!mediaGroupBuffers[mgId]) {
      mediaGroupBuffers[mgId] = { downloads: [], text: '', timer: null };
    }
    const buffer = mediaGroupBuffers[mgId];

    if (text) buffer.text = text;

    // Запускаємо завантаження одразу, а в таймері чекаємо на всі проміси
    if (photoFileId) {
      buffer.downloads.push(downloadTelegramPhoto(photoFileId));
    }

    if (buffer.timer) clearTimeout(buffer.timer);

    buffer.timer = setTimeout(async () => {
      try {
        const groupData = mediaGroupBuffers[mgId];
        delete mediaGroupBuffers[mgId];

        if (!groupData || !groupData.text) return;

        const parsedData = parseTirePost(groupData.text);
        if (!parsedData) {
          console.log('Пропущено альбом: текст не схожий на пост із шиною');
          return;
        }

        const currentSession = getSession(chatId);

        const isDuplicate = currentSession.items.some(
          item => item._uniqueKey === parsedData._uniqueKey
        );

        if (isDuplicate) {
          await sendTempMessage(
            ctx,
            `⚠️ Повторний пост (дублікат проігноровано):\n${parsedData.title} — ${parsedData.price_value} грн`
          );
          return;
        }

        ensureUniqueExternalId(currentSession, parsedData);
        const buffers = await Promise.all(groupData.downloads);
        const photosCount = attachPhotos(parsedData, buffers);
        currentSession.items.push(parsedData);

        await showSummary(ctx, chatId);
      } catch (err) {
        console.error('Помилка обробки media group:', err);
      }
    }, 2500);

    return;
  }

  // ---- Звичайний пост ----
  if (!text || text.startsWith('/')) return;

  const parsedData = parseTirePost(text);
  if (!parsedData) {
    console.log('Пропущено повідомлення: текст не схожий на пост із шиною');
    return;
  }

  const isDuplicate = session.items.some(
    item => item._uniqueKey === parsedData._uniqueKey
  );

  if (isDuplicate) {
    await sendTempMessage(
      ctx,
      `⚠️ Повторний пост (дублікат проігноровано):\n${parsedData.title} — ${parsedData.price_value} грн`
    );
    return;
  }

  ensureUniqueExternalId(session, parsedData);

  if (photoFileId) {
    const buffer = await downloadTelegramPhoto(photoFileId);
    attachPhotos(parsedData, [buffer]);
  }

  session.items.push(parsedData);

  await showSummary(ctx, chatId);
});

// ---- Глобальна обробка помилок, щоб бот не падав ----
bot.catch((err, ctx) => {
  console.error(`Помилка в обробнику (${ctx?.updateType}):`, err);
});

process.on('unhandledRejection', (err) => console.error('unhandledRejection:', err));
process.on('uncaughtException', (err) => console.error('uncaughtException:', err));

bot.launch({ dropPendingUpdates: true }, () => {
  console.log('Бот запущений та готовий до роботи!');
}).catch((err) => {
  console.error('Не вдалося запустити бота:', err);
  process.exit(1);
});

process.once('SIGINT', () => bot.stop('SIGINT'));
process.once('SIGTERM', () => bot.stop('SIGTERM'));
