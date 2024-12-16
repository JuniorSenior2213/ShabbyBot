const TelegramBot = require("node-telegram-bot-api");
const axios = require("axios");
const XLSX = require("xlsx");
const fs = require("fs");
const filePath = "./.xlsx";
const express = require("express");
const bot = new TelegramBot("7569058734:AAHzloldM_k7clVkRXik-7lWsXs5uRAP_oc", {
  polling: true,
  baseApiUrl: "https://api.telegram.org",
});

// ================================= Статичесие кнопки =================================
const Staticoptions = {
  reply_markup: {
    keyboard: [["Чи на часі Зара пара?"], ["Тест Всіх пар дня тижня"]],
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
bot.onText(/\/start/, (msg) => {
  return bot.sendMessage(msg.chat.id, `Привет`, Staticoptions);
});
bot.onText(/\/info/, (msg) => {
  return bot.sendMessage(
    msg.chat.id,
    "Я Раб Ленивого. По приказу говорю какая сейчас пара в гуппе КУ-11 Основываясь на информации что дана нам в Таблице на гугл диске "
  );
});
//================================== Обработка-Команд ===================================

bot.on("message", async (msg) => {
  const chatId = msg.chat.id;
  const text = msg.text;
  const date = msg.date;
  const datenow = new Date(date * 1000); // Перетворюємо timestamp в об'єкт Date
  const date_Root = new Date(); // для работы со временем
  const dayOfWeek = datenow.getDay(); // время в днях ( дата )
  const timeNOW_minutes = datenow.getMinutes(); // время в минутах
  const timeNOW_hours = datenow.getHours(); // Время в часах
  const totalMinutesNow = timeNOW_hours * 60 + timeNOW_minutes;
  console.log(
    `timeNOW_hours :${timeNOW_hours} timeNOW_minutes :${timeNOW_minutes}`
  );

  //================================== функции статических кнопок ==================================
  if (text === "Чи на часі Зара пара?") {
    HowPair(chatId, dayOfWeek, timeNOW_hours, timeNOW_minutes);
  }
  if (text === "Тест Всіх пар дня тижня") {
    HowPairTEST(dayOfWeek, chatId);
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
        return { row, col };
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
  days
) {
  let { getpair1 } = getNextPair(timeNOW_hours, timeNOW_minutes);
  let { otvet } = getNextPair(timeNOW_hours, timeNOW_minutes);
  let pair;
  let J = 0;
  let group = findCellWithWord(data, "КУ11");
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
  console.log(`Это ячейка группы:`, excelRangegroup, group);
  console.log(
    "Диапазон строк номера Пары которые мы читаем:",
    mergedRange_pair.rowStart,
    mergedRange_pair.rowEnd
  );
  let nextpair = readRangeWithCoordinates(
    filePath,
    mergedRange_pair.rowStart,
    mergedRange_pair.rowEnd,
    7,
    7
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
    7,
    7
  );

  console.log("Столбец группы:", 7, `\n ${readRange}`);
  console.log(`Текст пары номер ${getpair1}: ${formattedText}`);
  await bot.sendMessage(
    chatId,
    `${days + ""}\n\t\t${timeNOW_hours}:${timeNOW_minutes} `
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
  sheet
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
        "Понеділок"
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
        "Вівторок"
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
        "Середа"
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
        "Четвер"
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
        "П'ятниця"
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

async function HowPair(
  chatId,
  dayOfWeek,

  timeNOW_hours,
  timeNOW_minutes
) {
  try {
    // Загружаем Excel-файл по ссылке
    const response = await axios({
      url: "https://docs.google.com/spreadsheets/d/1ObGU1tzWNYOYE-F_wLlvMflVzw-aG3YL/export?format=xlsx&gid=961061976",
      method: "GET",
      responseType: "arraybuffer",
    });
    // Сохраняем файл временно
    fs.writeFileSync(filePath, response.data);

    // Читаем файл
    const workbook = XLSX.readFile(filePath);

    // Первый лист
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // Преобразуем в массив массивов
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    // Выводим данные в лог с проверкой объединенных ячеек
    console.log("Данные из таблицы (учитывая объединения):");
    console.log(
      `timeNOW_hours :${timeNOW_hours} timeNOW_minutes :${timeNOW_minutes}`
    );
    await Pair_is(
      chatId,
      dayOfWeek,
      data,
      timeNOW_hours,
      timeNOW_minutes,
      sheet
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
    return fs.unlinkSync(filePath);
  } catch (error) {
    console.error("Ошибка при скачивании или обработке файла:", error);
    bot.sendMessage(
      chatId,
      " Леле! Через клятого розробника все знову пішло по пізді, повідомте йому цю гарну новину!"
    );
  }
}

// функція кнопки "Тест Всіх пар дня тижня"
async function HowPairTEST(
  dayOfWeek,

  chatId
) {
  //================================== функции статических кнопок ==================================

  try {
    // Загружаем Excel-файл по ссылке
    const response = await axios({
      url: "https://docs.google.com/spreadsheets/d/1tr1CAnNUpXX9YvaJmkUakn1qj2jNpQLY/export?format=xlsx&gid=133878880",
      method: "GET",
      responseType: "arraybuffer",
    });
    // Сохраняем файл временно
    fs.writeFileSync(filePath, response.data);

    // Читаем файл
    const workbook = XLSX.readFile(filePath);

    // Первый лист
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // Преобразуем в массив массивов
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    // Выводим данные в лог с проверкой объединенных ячеек
    console.log("Данные из таблицы (учитывая объединения):");
    bot.sendMessage(chatId, `Який день тижня хочите протестити?`, Day_coptions);

    dayOfWeek = await waitForButtonPress(bot);
    bot.sendMessage(chatId, `ПЇХАЛИ`);

    for (let h = 0; h < 23; h++) {
      for (let m = 0; m < 60; m += 10) {
        console.log(`h = ${h}\nm = ${m}\n`);

        await Pair_is(chatId, dayOfWeek, data, h, m, sheet);
        await delay(1000);
      }
    }

    // Pair_is(chatId, dayOfWeek, data, timeNOW_hours1, timeNOW_minutes1, sheet);

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

const waitForButtonPress = (bot) => {
  return new Promise((resolve) => {
    bot.once("callback_query", (callbackQuery) => {
      dayOfWeek = parseInt(callbackQuery.data, 10);

      // Повідомляємо Promise, що кнопка натиснута
      resolve(dayOfWeek);
    });
  });
};
//============================================= Функции ==============================================
