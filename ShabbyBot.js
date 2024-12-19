const TelegramBot = require("node-telegram-bot-api");
const axios = require("axios");
const XLSX = require("xlsx");
const fs = require("fs");
const filePath = "./.xlsx";
const path = require("path"); // Для удобной работы с путями
const express = require("express");
const CSV_URL =
  "https://docs.google.com/spreadsheets/d/1ObGU1tzWNYOYE-F_wLlvMflVzw-aG3YL/export?format=csv"; // Ссылка на CSV-файл
const FILE_PATH = path.resolve(__dirname, "local_data.txt"); // Локальный файл для хранения данных
const url =
  "https://docs.google.com/spreadsheets/d/1tr1CAnNUpXX9YvaJmkUakn1qj2jNpQLY/export?format=xlsx&gid=133878880";

const bot = new TelegramBot("7569058734:AAHzloldM_k7clVkRXik-7lWsXs5uRAP_oc", {
  polling: true,
  baseApiUrl: "https://api.telegram.org",
});
const userMessages = {}; // Объект для хранения времени сообщений пользователей
const userFrozen = {}; // Объект для отслеживания "замороженных" пользователей
const SPAM_LIMIT = 2; // Лимит сообщений
const TIME_LIMIT = 2 * 1000; // Время в миллисекундах (1 секунд)
const FREEZE_TIME = 4 * 1000; // Время заморозки в миллисекундах (4 секунд)
const userGroups = {}; // Объект для хранения групп пользователей

// ================================= Статичесие кнопки =================================

// prettier-ignore
const Staticoptions = {
  reply_markup: {
    keyboard: [["Чи на часі Зара пара?", "Тест Всіх пар дня тижня"],],
  },
};
const StopButton = {
  reply_markup: {
    keyboard: [["Стоп!"]],
  },
};

const Day_coptions = {
  reply_markup: {
    inline_keyboard: [
      [{ text: "Понеділок", callback_data: 1 }],
      [{ text: "Вівторок", callback_data: 2 }],
      [{ text: "Середа", callback_data: 3 }],
      [{ text: "Четвер", callback_data: 4 }],
      [{ text: "П'ятниця", callback_data: 5 }],
      [{ text: "Субота", callback_data: 6 }],
      [{ text: "Неділля", callback_data: 0 }],
    ],
  },
};

const groupOptions = {
  reply_markup: {
    // prettier-ignore

    inline_keyboard: [
      [{ text: "КС11", callback_data: `КС11` },{ text: "КС12", callback_data: `КС12` },{ text: "КС13", callback_data: `КС13` }, { text: "КС14", callback_data: `КС14` }],
      [{ text: "КБ11", callback_data: `КБ11` },{ text: "КБ12", callback_data: `КБ12` }],
      [{ text: "КУ11", callback_data: `КУ11` }],
      [{ text: "КІ11", callback_data: `КІ11` }],
    ],
  },
};
// ================================= Статичесие кнопки ==================================
// Вебсервер для UptimeRobot
const app = express();

app.get("/", (req, res) => {
  res.send("Бот працює!");
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Сервер запущено на порту ${PORT}`);
});

//================================== Обработка-Команд ===================================
bot.onText(/\/start/, async (msg) => {
  return bot.sendMessage(msg.chat.id, "Привіт! ", Staticoptions);
});
bot.onText(/\/info/, (msg) => {
  return bot.sendMessage(
    msg.chat.id,
    "Я Раб Ленивого. По приказу говорю какая сейчас пара в гуппе КУ-11 Основываясь на информации что дана нам в Таблице на гугл диске "
  );
});

//================================== Обработка-Команд ===================================

bot.on("message", async (msg) => {
  // console.log(msg);
  const userId = msg.from.id;
  const chatId = msg.chat.id;
  const text = msg.text;
  const date = msg.date;
  const datenow = new Date(date * 1000); // Перетворюємо timestamp в об'єкт Date
  const date_Root = new Date(); // для работы со временем
  const dayOfWeek = datenow.getDay(); // время в днях ( дата )
  const timeNOW_minutes = datenow.getMinutes(); // время в минутах
  const timeNOW_hours = datenow.getHours(); // Время в часах
  const totalMinutesNow = timeNOW_hours * 60 + timeNOW_minutes;
  // Если пользователь "заморожен", не реагируем на его сообщение
  if (await userFrozen[userId]) {
    return;
  }
  if (await checkSpam(userId)) {
    // Если спам, отправляем сообщение и замораживаем пользователя
    await bot.sendMessage(chatId, `Досить спамити козаче..`);
    await freezeUser(userId, chatId);
    await wait(3000);
    return unfreezeUser(userId, chatId);
  }
  console.log(
    `timeNOW_hours :${timeNOW_hours} timeNOW_minutes :${timeNOW_minutes}`
  );

  await sendMessageWithTyping(chatId, text);

  // Если у пользователя нет сохраненной группы, запрашиваем её
  await takeGroup(userId, userGroups, chatId);

  // Добавляем время нового сообщения
  if (await !userMessages[userId]) {
    userMessages[userId] = [];
  }
  userMessages[userId].push(Date.now());

  //================================== функции статических кнопок ==================================
  if (text === "Чи на часі Зара пара?") {
    await HowPair(chatId, dayOfWeek, timeNOW_hours, timeNOW_minutes, userId);
  }
  if (text === "Тест Всіх пар дня тижня") {
    // freezeUser(userId, chatId);
    await HowPairTEST(dayOfWeek, chatId, userId);
  }
  if (text === "Стоп!") {
    unfreezeUser(userId, chatId);
    return bot.sendMessage(chatId, `Ви розморожені.`);
  }
});
//================================== функции статических кнопок ==================================

//============================================= Функции ==========================================

function readRangeWithCoordinates(
  filePath,
  startRow,
  endRow,
  startCol,
  endCol
) {
  const workbook = XLSX.readFile(filePath); // Читаем файл Excel
  const sheet = workbook.Sheets[workbook.SheetNames[0]]; // Получаем первый лист
  const merges = sheet["!merges"] || []; // Список объединённых диапазонов
  let result = []; // Результат
  let processedCells = new Set(); // Для отслеживания уже обработанных ячеек

  // Функция для проверки, находится ли ячейка в объединённом диапазоне
  function findMergeRange(row, col) {
    for (const merge of merges) {
      const { s, e } = merge; // s - начальная ячейка, e - конечная ячейка
      if (
        row >= s.r + 1 &&
        row <= e.r + 1 &&
        col >= s.c + 1 &&
        col <= e.c + 1
      ) {
        return merge;
      }
    }
    return null;
  }

  for (let row = startRow; row <= endRow; row++) {
    for (let col = startCol; col <= endCol; col++) {
      const cellKey = `${row}-${col}`;
      if (processedCells.has(cellKey)) {
        // Пропускаем ячейки, которые уже обработаны
        continue;
      }

      const mergeRange = findMergeRange(row, col);

      if (mergeRange) {
        const { s, e } = mergeRange;
        const mergeStartCell = XLSX.utils.encode_cell({ r: s.r, c: s.c });
        const mergeValue = sheet[mergeStartCell]
          ? sheet[mergeStartCell].v
          : null;

        // Добавляем объединённый диапазон только один раз
        result.push({
          rowStart: startRow,
          rowEnd: endRow,
          colStart: startCol,
          colEnd: endCol,
          cellAddress: mergeStartCell,
          value: mergeValue,
          merged: true,
        });

        // Помечаем все ячейки из этого диапазона как обработанные
        for (let mergeRow = s.r + 1; mergeRow <= e.r + 1; mergeRow++) {
          for (let mergeCol = s.c + 1; mergeCol <= e.c + 1; mergeCol++) {
            processedCells.add(`${mergeRow}-${mergeCol}`);
          }
        }
      } else {
        // Если ячейка не объединена
        const cellAddress = XLSX.utils.encode_cell({ r: row - 1, c: col - 1 });
        const cellValue = sheet[cellAddress] ? sheet[cellAddress].v : null;

        result.push({
          rowStart: startRow,
          rowEnd: endRow,
          colStart: startCol,
          colEnd: endCol,
          cellAddress: cellAddress,
          value: cellValue,
          merged: false,
        });

        processedCells.add(cellKey);
      }
    }
  }

  return result;
}

// Функция поиска ячейки с конкретным словом (например, "Понеділок")
function findCellWithWord(
  data,
  word,
  startRow = 0,
  endRow = data.length,
  startCol = 0,
  endCol = data.length
) {
  word += "";
  for (let col = startCol; col < endCol; col++) {
    for (let row = startRow; row < endRow; row++) {
      if (
        data[row][col] &&
        data[row][col]
          .toString()
          .trim()
          .toLowerCase()
          .includes(word.toLowerCase())
      ) {
        return { row, col, word };
      }
    }
  }
  return null; // Возвращает null, если слово не найдено в указанном диапазоне
}

function findMergedRange(sheet, cellAddress) {
  // Преобразуем адрес ячейки в строку и столбец
  function parseCellAddress(address) {
    const match = address.match(/^([A-Z]+)(\d+)$/);
    if (!match) {
      throw new Error(`Invalid cell address: ${address}`);
    }
    const col = letterToColumn(match[1]); // Преобразуем букву столбца в индекс
    const row = parseInt(match[2], 10); // Преобразуем строку в число
    return { row, col };
  }

  // Преобразуем буквы столбца в индекс (например, A -> 0, B -> 1)
  function letterToColumn(letter) {
    let col = 0;
    for (let i = 0; i < letter.length; i++) {
      col = col * 26 + (letter.charCodeAt(i) - 65 + 1);
    }
    return col - 1;
  }

  // Преобразуем индекс столбца в буквы (например, 0 -> A)
  function columnToLetter(col) {
    let letter = "";
    while (col >= 0) {
      letter = String.fromCharCode((col % 26) + 65) + letter;
      col = Math.floor(col / 26) - 1;
    }
    return letter;
  }

  // Проверяем, входит ли ячейка в диапазон
  function isCellInRange(cell, range) {
    return (
      cell.row >= range.s.r &&
      cell.row <= range.e.r &&
      cell.col >= range.s.c &&
      cell.col <= range.e.c
    );
  }

  // Разбираем адрес целевой ячейки
  const targetCell = parseCellAddress(cellAddress);

  // Перебираем все объединенные диапазоны в листе
  const merges = sheet["!merges"] || [];
  for (const range of merges) {
    // Если целевая ячейка находится в диапазоне
    if (isCellInRange(targetCell, range)) {
      // Преобразуем диапазон в Excel-формат
      const mergedRange = {
        rowStart: range.s.r + 1,
        rowEnd: range.e.r + 1,
        colStart: range.s.c,
        colEnd: range.e.c,
      };
      return mergedRange;
    }
  }

  // Если ячейка не найдена в объединенных диапазонах
  return null;
}

// Расписание пар ( ближайшая пара)
function getNextPair(timeNOW_hours, timeNOW_minutes) {
  if (timeNOW_minutes <= 0 && timeNOW_hours <= 0) {
    timeNOW_minutes++;
  }
  const totalMinutesNow = timeNOW_hours * 60 + timeNOW_minutes;
  let otvet;

  // Расписание пар
  const pairs = [
    { start: 510, end: 590 }, // 1 пара: 8:30 - 9:50
    { start: 610, end: 690 }, // 2 пара: 10:10 - 11:30
    { start: 720, end: 800 }, // 3 пара: 12:00 - 13:20
    { start: 820, end: 900 }, // 4 пара: 13:40 - 15:00
    { start: 920, end: 1000 }, // 5 пара: 15:20 - 16:40
    { start: 1020, end: 1100 }, // 6 пара: 17:00 - 18:20
    { start: 1120, end: 1200 }, // 7 пара: 18:40 - 20:00
    { start: 1120, end: 1200 }, // 8 Опциональна
  ];

  // Проверяем, если текущее время меньше первой пары
  if (totalMinutesNow < pairs[0].start && totalMinutesNow > 0) {
    otvet = `Час Для першої пари не пробив.`;
    console.log(`Пишется що час для першої пари не пробив `);
    return { getpair1: 1, otvet: otvet }; //`Следующая пара: 1 (завтра в 08:30)`;
  }

  // Проверяем, если текущее время больше последней пары
  if (totalMinutesNow > pairs[pairs.length - 1].end) {
    console.log(`Усі Пари скінчились і вертаєтся значення -1`);
    otvet = `На сьогодні Пари вже, не на часі. `;
    return { getpair1: -1, otvet: otvet }; //`Все пары закончились. Следующая пара: 1 (завтра в 08:30)`;
  }

  // Находим следующую пару
  for (let i = 0; i < pairs.length; i++) {
    if (i >= 7) {
      otvet = `зараз час ${i} пари але її не знайденно в списку! Відпочиваймо.`;
      return { getpair1: i, otvet: otvet };
    }
    if (totalMinutesNow < pairs[i].start) {
      if (totalMinutesNow > pairs[i - 1].end) {
        otvet = `Зараз перерва, наспуна пара : ${i + 1}`;
        console.log(
          `Пишется що зараз перерва, вертаєтся значення наступної пари`
        );

        return { getpair1: i + 1, otvet: otvet };
      }

      console.log(
        `пишется що зараз йде такато пара, вертаєтся значення цієї пари`
      );
      otvet = `Зараз йде: ${i} пара`;

      return { getpair1: i, otvet: otvet };
    }
  }

  // Если все пары закончились, возвращаем сообщение про завтрашнюю пару
  return 1; //`Все пары закончились. Следующая пара: 1 (завтра в 08:30)`;
}

function convertCellToExcelAddress(cell) {
  // Преобразуем индекс столбца (например, 0) в букву (A, B, C...)
  const columnToLetter = (col) => {
    let letter = "";
    while (col >= 0) {
      letter = String.fromCharCode((col % 26) + 65) + letter;
      col = Math.floor(col / 26) - 1;
    }
    return letter;
  };

  // Получаем столбец в виде буквы
  const colLetter = columnToLetter(cell.col);

  // Возвращаем адрес ячейки в формате "A130"
  return `${colLetter}${cell.row}`;
}

// функція 'Яка зараз пара?'

async function Pair_isit(
  chatId,
  dayOfWeek,
  data,
  timeNOW_hours,
  timeNOW_minutes,
  sheet,
  days,
  userId
) {
  const groupsPromise = takeGroup(userId, userGroups, chatId);
  const groups = await groupsPromise; // Используем await для получения значения из Promise
  console.log(groups); // Теперь groups содержит значение 'КУ11'
  let { getpair1 } = getNextPair(timeNOW_hours, timeNOW_minutes);
  let { otvet } = getNextPair(timeNOW_hours, timeNOW_minutes);
  let pair;
  let J = 0;
  let group = findCellWithWord(data, groups);
  console.log(group);
  const excelRangegroup = convertCellToExcelAddress(group);
  const mergedRangegroup = findMergedRange(sheet, excelRangegroup);
  day = findCellWithWord(data, days + "");
  console.log("день:", days);
  const excelRangeday = convertCellToExcelAddress(day);
  const mergedRange = findMergedRange(sheet, excelRangeday);
  console.log(mergedRange);
  console.log(
    `${mergedRange.rowStart} это столбики нашего слова \n${mergedRange.rowEnd} это рядочки нашего сглова\n `
  );
  J += mergedRange.rowEnd;
  pair = findCellWithWord(
    data,
    getpair1 + "",
    (startRow = mergedRange.rowStart - 1)
  );

  console.log(`Pair: `, pair);

  if (pair === null) {
    if (getpair1 != -1) {
      otvet += `\nПари не буде.`;
    }
    await bot.sendMessage(
      chatId,
      `${days + ""}\n\t\t${timeNOW_hours}:${timeNOW_minutes} `
    );

    return bot.sendMessage(chatId, `${otvet}`);
  }
  const excelRange_pair = convertCellToExcelAddress(pair);
  console.log("Номер нашей пары:", getNextPair(timeNOW_hours, timeNOW_minutes));
  const mergedRange_pair = findMergedRange(sheet, excelRange_pair);
  console.log(`Это ячейка пары: `, excelRange_pair, mergedRange_pair);
  console.log(`Это ячейка группы:`, excelRangegroup, group.col);
  console.log(
    "Диапазон строк номера Пары которые мы читаем:",
    mergedRange_pair.rowStart,
    mergedRange_pair.rowEnd
  );
  let nextpair = readRangeWithCoordinates(
    filePath,
    mergedRange_pair.rowStart,
    mergedRange_pair.rowEnd,
    group.col + 1,
    group.col + 1
  );
  console.log(
    `(Количество столбцов с содержимым)values.length:`,
    nextpair.length
  );

  let values = nextpair
    .map((item) => item.value)
    .filter((value) => value !== null); // Убираем null
  const formattedText = values.join("\n"); // Соединяем через перенос строки

  // Читаем диапазон ячеек
  const readRange = readRangeWithCoordinates(
    filePath,
    mergedRange_pair.rowStart,
    mergedRange_pair.rowEnd,
    group.col + 1,
    group.col + 1
  );

  console.log("Столбец группы:", group.col + 1 + "", `\n ${readRange}`);
  console.log(`Текст пары номер ${getpair1}: ${formattedText}`);
  await bot.sendMessage(
    chatId,
    `${days + ""}\n\t\t${timeNOW_hours}:${timeNOW_minutes} \n ${group.word}`
  );
  if (values.length === 0) {
    return bot.sendMessage(chatId, `${otvet} \nПари не буде.`);
  } else {
    await bot.sendMessage(chatId, `${otvet}, `);
    return bot.sendMessage(chatId, `${formattedText}`);
  }
}

// функція кнопки " яка зараз пара "
async function Pair_is(
  chatId,
  dayOfWeek,
  data,
  timeNOW_hours,
  timeNOW_minutes,
  sheet,
  userId
) {
  switch (dayOfWeek) {
    case 1:
      Pair_isit(
        chatId,
        dayOfWeek,
        data,
        timeNOW_hours,
        timeNOW_minutes,
        sheet,
        "Понеділок",
        userId
      );
      break;
    case 2:
      Pair_isit(
        chatId,
        dayOfWeek,
        data,
        timeNOW_hours,
        timeNOW_minutes,
        sheet,
        "Вівторок",
        userId
      );
      break;
    case 3:
      Pair_isit(
        chatId,
        dayOfWeek,
        data,
        timeNOW_hours,
        timeNOW_minutes,
        sheet,
        "Середа",
        userId
      );
      break;
    case 4:
      Pair_isit(
        chatId,
        dayOfWeek,
        data,
        timeNOW_hours,
        timeNOW_minutes,
        sheet,
        "Четвер",
        userId
      );
      break;
    case 5:
      Pair_isit(
        chatId,
        dayOfWeek,
        data,
        timeNOW_hours,
        timeNOW_minutes,
        sheet,
        "П'ятниця",
        userId
      );
      break;
    default:
      console.log(dayOfWeek);
      bot.sendMessage(chatId, `скоріше за все зараз віхідні, енжой.`);
  }

  return;
}

// Функція для створення затримки
function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

//================================== функции статических кнопок ==================================

async function HowPair(
  chatId,
  dayOfWeek,
  timeNOW_hours,
  timeNOW_minutes,
  userId
) {
  try {
    // Загружаем Excel-файл по ссылке
    const response = await axios({
      method: "GET",
      responseType: "arraybuffer",
      url: "https://docs.google.com/spreadsheets/d/1tr1CAnNUpXX9YvaJmkUakn1qj2jNpQLY/export?format=xlsx&gid=133878880",
    });

    fetchCSVData()
      .then((reserch) => {
        if (reserch) {
          console.log("Данные не изменились");
        } else {
          console.log("Данные изменились, обновляем файл...");
          checkAndDownloadFile(url, filePath);
        }
      })
      .catch((error) => console.error("Ошибка: чтения либо загрузки", error));
    const workbook = XLSX.readFile(filePath);
    // Дальнейшая обработка файла
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    // Выводим данные в лог с проверкой объединенных ячеек
    console.log("Данные из таблицы (учитывая объединения):");
    console.log(
      `timeNOW_hours :${timeNOW_hours} timeNOW_minutes :${timeNOW_minutes}`
    );

    dayOfWeek = 1;
    await Pair_is(
      chatId,
      dayOfWeek,
      data,
      timeNOW_hours,
      timeNOW_minutes,
      sheet,
      userId
    );

    // Получаем данные об объединении ячеек
    const merges = sheet["!merges"] || [];

    // Логируем строки с указанием объединенных ячеек
    data.forEach((row, rowIndex) => {
      const rowLog = row
        .map((cell, colIndex) => {
          const merged = isMergedCell(rowIndex, colIndex) ? "(Merged)" : "";
          return `${cell || ""} ${merged}`.trim(); // Помечаем объединенные ячейки
        })
        .join(" | ");
      // console.log(rowLog); // вывести как видит код наш документ в консоль
    });

    // Функция для проверки, находится ли ячейка в диапазоне объединения
    function isMergedCell(row, col) {
      for (const merge of merges) {
        const { s, e } = merge; // s - начало объединения, e - конец объединения
        if (row >= s.r && row <= e.r && col >= s.c && col <= e.c) {
          return true; // Ячейка объединена
        }
      }
      return false;
    }
    // Удаляем временный файл
  } catch (error) {
    console.error("Ошибка при скачивании или обработке файла:", error);
    bot.sendMessage(
      chatId,
      " Леле! Через клятого розробника все знову пішло по пізді, повідомте йому цю гарну новину!"
    );
  }
}

// функція кнопки "Тест Всіх пар дня тижня"
async function HowPairTEST(dayOfWeek, chatId, userId, text) {
  try {
    // Загружаем Excel-файл по ссылке
    const response = await axios({
      url: "https://docs.google.com/spreadsheets/d/1tr1CAnNUpXX9YvaJmkUakn1qj2jNpQLY/export?format=xlsx&gid=133878880",
      method: "GET",
      responseType: "arraybuffer",
    });
    const fileUrl =
      "https://docs.google.com/spreadsheets/d/1tr1CAnNUpXX9YvaJmkUakn1qj2jNpQLY/export?format=xlsx&gid=133878880"; // Замените на вашу ссылку
    const filePath = "./file.xlsx"; // Замените на ваш путь к файлу

    await fetchCSVData()
      .then((reserch) => {
        if (reserch) {
          console.log("Данные не изменились");
        } else {
          console.log("Данные изменились, обновляем файл...");
          checkAndDownloadFile(url, filePath);
        }
      })
      .catch((error) => console.error("Ошибка: чтения либо загрузки", error));
    const workbook = XLSX.readFile(filePath);

    // Дальнейшая обработка файла
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    // Выводим данные в лог с проверкой объединенных ячеек
    console.log("Данные из таблицы (учитывая объединения):");
    bot.sendMessage(chatId, `Який день тижня хочите протестити?`, Day_coptions);

    dayOfWeek = await waitForButtonPressINT(bot, chatId);
    console.log(dayOfWeek);
    if (dayOfWeek === 6) {
      return bot.sendMessage(
        chatId,
        `Функція визначення пар в суботу ще не написана..`
      );
    } else if (dayOfWeek === 0) {
      return bot.sendMessage(chatId, `По неділям ми не вчимось!`);
    }
    bot.sendMessage(chatId, `ПЇХАЛИ`, StopButton);

    for (let h = 0; h < 23; h++) {
      for (let m = 0; m < 60; m += 10) {
        if (text === "Стоп!") {
          break;
        }
        console.log(`h = ${h}\nm = ${m}\n`);

        await Pair_is(chatId, dayOfWeek, data, h, m, sheet, userId);
        await delay(1000);
      }
    }
    Staticoptions;
    // Получаем данные об объединении ячеек
    const merges = sheet["!merges"] || [];

    // Логируем строки с указанием объединенных ячеек
    data.forEach((row, rowIndex) => {
      const rowLog = row
        .map((cell, colIndex) => {
          const merged = isMergedCell(rowIndex, colIndex) ? "(Merged)" : "";
          return `${cell || ""} ${merged}`.trim(); // Помечаем объединенные ячейки
        })
        .join(" | ");
      // console.log(rowLog); // вывести как видит код наш документ в консоль
    });

    // Функция для проверки, находится ли ячейка в диапазоне объединения
    function isMergedCell(row, col) {
      for (const merge of merges) {
        const { s, e } = merge; // s - начало объединения, e - конец объединения
        if (row >= s.r && row <= e.r && col >= s.c && col <= e.c) {
          return true; // Ячейка объединена
        }
      }
      return false;
    }
    // Удаляем временный файл
    return fs.unlinkSync(filePath);
  } catch (error) {
    console.error("Ошибка при скачивании или обработке файла:", error);
    bot.sendMessage(
      chatId,
      " Леле! Через клятого розробника все знову пішло по пізді, повідомте йому цю гарну новину!"
    );
  }
}

// Функция для ожидания нажатия кнопки
function waitForButtonPress(bot, chatId) {
  return new Promise((resolve, reject) => {
    bot.once("callback_query", (query) => {
      if (
        query.message &&
        query.message.chat &&
        query.message.chat.id === chatId
      ) {
        if (query.data) {
          resolve(query.data); // Возвращаем данные кнопки
        } else {
          reject(new Error("No data in callback query"));
        }
      } else {
        reject(new Error("Callback query from different chat"));
      }
    });
  });
}

function waitForButtonPressINT(bot, chatId) {
  return new Promise((resolve) => {
    bot.once("callback_query", (query) => {
      if (query.message.chat.id === chatId) {
        resolve(parseInt(query.data, 10)); // Преобразуем данные кнопки в целое число
      }
    });
  });
}

// функція антимпам
function checkSpam(userId) {
  const currentTime = Date.now();

  // Очистка старых сообщений
  if (userMessages[userId]) {
    userMessages[userId] = userMessages[userId].filter(
      (timestamp) => currentTime - timestamp < TIME_LIMIT
    );
  }

  // Проверка на спам
  if (userMessages[userId] && userMessages[userId].length >= SPAM_LIMIT) {
    return true;
  }

  return false;
}

// Функція заморозки користувача
function freezeUser(userId, chatId) {
  userFrozen[userId] = Date.now();

  console.log(`Користувач ${userId} заморожений`);
}

// Функція для зняття заморозки користувача вручну
function unfreezeUser(userId) {
  if (userFrozen[userId]) {
    delete userFrozen[userId];
    console.log(`Користувач ${userId} розморожений вручну`);
  } else {
    console.log(`Користувач ${userId} не був заморожений`);
  }
}
async function takeGroup(userId, userGroups, chatId) {
  if (!userGroups[userId]) {
    freezeUser(userId, chatId);
    bot.sendMessage(chatId, "Будьласка обери свою групу:", groupOptions);
    const groups = await waitForButtonPress(bot, chatId);
    userGroups[userId] = groups; // Сохраняем группу пользователя
    await bot.sendMessage(chatId, `Ваша группа сохранена: ${groups}`);
    unfreezeUser(userId);

    return groups;
  } else {
    return userGroups[userId];
  }
}

// функция таймера
function wait(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

// Функция для отправки сообщения и удаления "Генерую Відповідь..."
async function sendMessageWithTyping(chatId, text, userId) {
  if (userFrozen[userId]) {
    return;
  }
  if (await checkSpam(userId)) {
    // Если спам, отправляем сообщение и замораживаем пользователя
    await bot.sendMessage(chatId, `Досить спамити козаче..`);
    await freezeUser(userId, chatId);
    await wait(5000);
    return unfreezeUser(userId, chatId);
  }
  if (userFrozen[userId]) {
    return;
  }
  // Отправляем сообщение "Генерую Відповідь..."
  const typingMessage = await bot.sendMessage(chatId, "Генерую Відповідь...");
  await freezeUser(userId, chatId);
  await wait(2000);
  // Удаляем сообщение "Генерую Відповідь..."
  await bot.deleteMessage(chatId, typingMessage.message_id);
  return unfreezeUser(userId, chatId);
}

async function fetchCSVData() {
  try {
    const response = await axios.get(CSV_URL, { responseType: "arraybuffer" }); // Загружаем CSV как байты
    const csvData = response.data;
    handleLocalFile(csvData);
  } catch (error) {
    console.error("Ошибка при получении данных:", error);
  }
}

async function handleLocalFile(csvData) {
  if (fs.existsSync(FILE_PATH)) {
    // Файл существует, читаем его байты
    const localData = fs.readFileSync(FILE_PATH);

    if (!Buffer.compare(localData, csvData)) {
      console.log("Данные совпадают. Файл не обновлялся.");
      return true; // Данные не изменились
    } else {
      console.log("Данные отличаются. Перезаписываем файл...");
      fs.writeFileSync(FILE_PATH, csvData);
      return false; // Данные изменились
    }
  } else {
    // Файла нет, создаем новый
    console.log("Файл не найден. Создаем новый файл...");
    fs.writeFileSync(FILE_PATH, csvData);
    return false; // Файл был создан
  }
}

async function checkAndDownloadFile(url, filePath) {
  // Проверка существования локального файла
  const reserch = fetchCSVData();
  if (!fs.existsSync(filePath) || !reserch) {
    console.log("Файл не найден или данные изменились, скачиваем...");
    const response = await axios.get(url, { responseType: "arraybuffer" });
    fs.writeFileSync(filePath, response.data);
    console.log("Файл успешно скачан.");
  } else {
    console.log("Файл уже существует и актуален.");
  }

  // Чтение файла
  const workbook = XLSX.readFile(filePath);
  return workbook;
}
