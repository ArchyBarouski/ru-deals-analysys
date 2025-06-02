// === ГЛАВНАЯ ФУНКЦИЯ ===
function prepareDealsReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("Nemind_export");
  const markupSheet = ss.getSheetByName("Разметка источников");
  const reportSheet = ss.getSheetByName("Отчёт по сделкам") || ss.insertSheet("Отчёт по сделкам");
  const dataSheet = ss.getSheetByName("Data");
  reportSheet.clear();

  const dataMap = {};
  if (dataSheet) {
    const dataValues = dataSheet.getRange("A2:C" + dataSheet.getLastRow()).getValues();
    dataValues.forEach(row => {
      const key = String(row[0] || "").replace(/[^\d]/g, "").trim();
      const val = String(row[2] || "").trim();
      if (key) dataMap[key] = val;
    });
  }

  function normalizeSource(rawSource, description, dataMap) {
    let source = String(rawSource || "").trim();
    if (/Оффлайн\s*\/\s*Offline/i.test(source)) return "Оффлайн / Offline";
    if (/Существующий клиент/i.test(source)) return "Существующий клиент";
    const phoneFromDesc = String(description || "").match(/\d{7,}/g);
    if (phoneFromDesc) {
      for (const phone of phoneFromDesc) {
        const normalized = phone.replace(/[^\d]/g, "");
        if (dataMap[normalized]) {
          source += ` (${dataMap[normalized]})`;
          break;
        }
      }
    }
    return source;
  }

  const markupData = markupSheet.getRange(2, 1, markupSheet.getLastRow() - 1, 5).getValues();
  const groupMap = {};
  markupData.forEach(row => {
    const key = row.slice(0, 4).map(String).map(v => v.trim()).join("|");
    groupMap[key] = row[4]?.trim() || "Прочие";
  });

  const sourceData = sourceSheet.getDataRange().getValues();
  const header = sourceData[0];
  const monthCol = header.indexOf("Дата создания");
  const stageGroupCol = header.indexOf("Группа стадий");
  const sourceCols = [
    header.indexOf("Источник"),
    header.indexOf("UTM Source"),
    header.indexOf("UTM Medium"),
    header.indexOf("UTM Campaign")
  ];
  const descCol = header.indexOf("Описание источника");

  const resultMap = new Map();
  const allMonthsSet = new Set();

  for (let i = 1; i < sourceData.length; i++) {
    const row = sourceData[i];
    const rawDate = row[monthCol];
    const stage = row[stageGroupCol]?.trim();
    const parsed = typeof rawDate === 'string' && rawDate.match(/(\d{2})[./-](\d{2})[./-](\d{4})/);
    if (!parsed) continue;
    const [_, day, monthNum, year] = parsed;
    const monthKey = `${year}-${monthNum}`;

    const rawSource = String(row[sourceCols[0]] || '').trim();
    const description = String(row[descCol] || '').trim();
    const normalizedSource = normalizeSource(rawSource, description, dataMap);
    const utmSource = String(row[sourceCols[1]] || '').trim();
    const utmMedium = String(row[sourceCols[2]] || '').trim();
    const utmCampaign = String(row[sourceCols[3]] || '').trim();

function normalizeKeyParts(...parts) {
  return parts.map(v => (v || "").toString().trim());
}

const [s1, s2, s3, s4] = normalizeKeyParts(normalizedSource, utmSource, utmMedium, utmCampaign);

const sourceKey = `${s1}|${s2}|${s3}|${s4}`;
const partialKey = `${s1}|${s2}||`;
const fallbackKey = `${s1}|||`;

// Логирование Reffection и всех совпадений
if (s1.toLowerCase().includes("reffection")) {
  Logger.log("🔍 Проверка Reffection:");
  Logger.log("  sourceKey:    " + sourceKey);
  Logger.log("  partialKey:   " + partialKey);
  Logger.log("  fallbackKey:  " + fallbackKey);
}

const group =
  groupMap[sourceKey] ||
  groupMap[partialKey] ||
  groupMap[fallbackKey] ||
  "Прочие";


    if (!resultMap.has(group)) resultMap.set(group, new Map());
    const groupData = resultMap.get(group);
    if (!groupData.has(monthKey)) groupData.set(monthKey, { total: 0, success: 0, open: 0, fail: 0 });

    const stats = groupData.get(monthKey);
    stats.total++;
    if (/успешно/i.test(stage)) stats.success++;
    else if (/открыт/i.test(stage)) stats.open++;
    else if (/провален/i.test(stage)) stats.fail++;

    allMonthsSet.add(monthKey);
  }

  const sortedMonths = Array.from(allMonthsSet).sort();
  const headerRow = ["Источник", ...sortedMonths.map(m => {
    const [y, mo] = m.split("-");
    return Utilities.formatDate(new Date(`${y}-${mo}-01`), "ru", "LLL.yy");
  }), "Весь период"]

let rawOutput = [];
const totalsByMonth = sortedMonths.map(() => ({ total: 0, success: 0, open: 0, fail: 0 }));
let grand = { total: 0, success: 0, open: 0, fail: 0 };

Logger.log("=== КЛЮЧИ В resultMap ===");
for (const key of resultMap.keys()) {
  Logger.log("→ " + key);
}
for (const [group, monthMap] of resultMap.entries()) {
  const row = [
    SpreadsheetApp.newRichTextValue()
      .setText(group)
      .setTextStyle(SpreadsheetApp.newTextStyle().setBold(true).build())
      .build()
  ];
  const full = { total: 0, success: 0, open: 0, fail: 0 };

  sortedMonths.forEach((month, i) => {
    const { total = 0, success = 0, open = 0, fail = 0 } = monthMap.get(month) || {};
    totalsByMonth[i].total += total;
    totalsByMonth[i].success += success;
    totalsByMonth[i].open += open;
    totalsByMonth[i].fail += fail;
    full.total += total;
    full.success += success;
    full.open += open;
    full.fail += fail;

    if (total === 0) {
      row.push(SpreadsheetApp.newRichTextValue().setText("").build());
    } else {
      const s = Math.round((success / total) * 100);
      const o = Math.round((open / total) * 100);
      const f = Math.round((fail / total) * 100);
      const text = `${total}\n% ${s}-${o}-${f}\nшт ${success}-${open}-${fail}`;
      const rich = SpreadsheetApp.newRichTextValue().setText(text);
      safeSetTextStyle(rich, text, String(total).length);
      row.push(rich.build());
    }
  });

  const { total, success, open, fail } = full;
  if (total > 0) {
    const s = Math.round((success / total) * 100);
    const o = Math.round((open / total) * 100);
    const f = Math.round((fail / total) * 100);
    const text = `${total}\n% ${s}-${o}-${f}\nшт ${success}-${open}-${fail}`;
    const rich = SpreadsheetApp.newRichTextValue().setText(text);
    safeSetTextStyle(rich, text, String(total).length);
    row.push(rich.build());

    grand.total += total;
    grand.success += success;
    grand.open += open;
    grand.fail += fail;
  } else {
    row.push(SpreadsheetApp.newRichTextValue().setText("").build());
  }

  // ✅ Добавляем в любом случае — даже если total = 0
  rawOutput.push({ group, richRow: row, total });
}


const analyticGroups = [
  "Yandex", "Organic", "Offline", "Бывший клиент", "Maps (yandex+2gis)", "Прямой переход", "Дизайнер", "Vk.com", "Email", "Referral", "Google", "Instagram", "Facebook", "Личный контакт","Холодная база","Reffection", "Марквиз", 
];

const existingGroupsNormalized = new Set(
  rawOutput.map(r => r.group.trim().toLowerCase())
);

for (const g of analyticGroups) {
  if (!existingGroupsNormalized.has(g.trim().toLowerCase())) {
    const emptyRow = [
      SpreadsheetApp.newRichTextValue()
        .setText(g)
        .setTextStyle(SpreadsheetApp.newTextStyle().setBold(true).build())
        .build()
    ];
    for (let i = 0; i < sortedMonths.length + 1; i++) {
      emptyRow.push(SpreadsheetApp.newRichTextValue().setText("").build());
    }
    rawOutput.push({ group: g, richRow: emptyRow, total: 0 });
  }
}


rawOutput = rawOutput.filter(r => analyticGroups.includes(r.group));
rawOutput.sort((a, b) => {
  const aIndex = analyticGroups.indexOf(a.group);
  const bIndex = analyticGroups.indexOf(b.group);
  if (aIndex !== -1 && bIndex !== -1) return aIndex - bIndex;
  if (aIndex !== -1) return -1;
  if (bIndex !== -1) return 1;
  return b.total - a.total;
});

rawOutput.forEach((entry, i) => {
  // ✅ Исправляем возможные несоответствия по длине строки
  while (entry.richRow.length < headerRow.length) {
    entry.richRow.push(SpreadsheetApp.newRichTextValue().setText("").build());
  }
  if (entry.richRow.length > headerRow.length) {
    entry.richRow = entry.richRow.slice(0, headerRow.length);
  }

  // ✅ Логируем, чтобы убедиться, что вставка сработала
  Logger.log(`⬇️ Вставляем строку: ${entry.group}, ячеек: ${entry.richRow.length}`);

  reportSheet.getRange(i + 3, 1, 1, headerRow.length).setRichTextValues([entry.richRow]);
});


const totalRowIndex = rawOutput.length + 1;
const totalRow = [];
totalRow[0] = SpreadsheetApp.newRichTextValue().setText("Итого").setTextStyle(SpreadsheetApp.newTextStyle().setBold(true).setFontSize(9).build()).build();

totalsByMonth.forEach(({ total, success, open, fail }, i) => {
  if (total === 0) {
    totalRow[i + 1] = SpreadsheetApp.newRichTextValue().setText("").build();
  } else {
    const s = Math.round((success / total) * 100);
    const o = Math.round((open / total) * 100);
    const f = Math.round((fail / total) * 100);
    const text = `${total}\n% ${s}-${o}-${f}\nшт ${success}-${open}-${fail}`;
    const rich = SpreadsheetApp.newRichTextValue().setText(text);
    safeSetTextStyle(rich, text, String(total).length);
    totalRow[i + 1] = rich.build();
  }
});

// Добавляем финальный столбец "Весь период"
if (grand.total > 0) {
  const s = Math.round((grand.success / grand.total) * 100);
  const o = Math.round((grand.open / grand.total) * 100);
  const f = Math.round((grand.fail / grand.total) * 100);
  const text = `${grand.total}\n% ${s}-${o}-${f}\nшт ${grand.success}-${grand.open}-${grand.fail}`;
  const rich = SpreadsheetApp.newRichTextValue().setText(text);
  safeSetTextStyle(rich, text, String(grand.total).length);
  totalRow.push(rich.build());
} else {
  totalRow.push(SpreadsheetApp.newRichTextValue().setText("").build());
}

const totalRowData = totalRow.slice(1);
reportSheet.getRange(totalRowIndex, 2, 1, totalRowData.length)
  .setRichTextValues([totalRowData]);


// Вставка и стилизация
styleFinalTable(reportSheet, headerRow, totalRowIndex);
addAnalyticsToGroups(reportSheet, analyticGroups, sortedMonths, resultMap, totalsByMonth, dataSheet);

function safeSetTextStyle(richTextBuilder, text, boldEnd) {
  const fullLen = text.length;
  const boldLen = Math.min(boldEnd, fullLen);
  const normalStart = boldLen;
  const boldStyle = SpreadsheetApp.newTextStyle().setBold(true).setFontSize(9).build();
  const normalStyle = SpreadsheetApp.newTextStyle().setFontSize(9).build();
  try {
    if (boldLen > 0) richTextBuilder.setTextStyle(0, boldLen, boldStyle);
    if (normalStart < fullLen) richTextBuilder.setTextStyle(normalStart, fullLen, normalStyle);
  } catch (e) {
    richTextBuilder.setTextStyle(0, fullLen, normalStyle);
  }
}

function addAnalyticsToGroups(reportSheet, analyticGroups, sortedMonths, resultMap, totalsByMonth, dataSheet) {
  const dataValues = dataSheet.getRange(2, 7, dataSheet.getLastRow() - 1, 3).getValues();
  const groupNamesInSheet = reportSheet
    .getRange(3, 1, reportSheet.getLastRow() - 2, 1)
    .getRichTextValues()
    .map(row => row[0].getText().trim());

  for (const groupName of analyticGroups) {
    const groupIndex = groupNamesInSheet.findIndex(text => text === groupName);
    if (groupIndex === -1) continue;

    const groupStatsMap = resultMap.get(groupName);
    if (!groupStatsMap) continue;

    sortedMonths.forEach((monthKey, i) => {
      const stats = groupStatsMap.get(monthKey) || {};
      const total = stats.total || 0;
      const success = stats.success || 0;
      const allMonthSuccess = totalsByMonth[i].success || 0;

      const cost = dataValues
        .filter(row => {
          const type = (row[0] || "").toString().trim();
          const dataMonthKey = convertMonthNameToKey((row[1] || "").toString().trim());
          return type === groupName && dataMonthKey === monthKey;
        })
        .reduce((sum, row) => sum + (parseFloat(row[2]) || 0), 0);

      let extraText = "";
      if (cost > 0) extraText += `\nРасход: $${cost.toFixed(0)}`;
      if (cost > 0 && success > 0) extraText += `\nЦена успех: $${(cost / success).toFixed(0)}`;
      if (success > 0 && allMonthSuccess > 0) {
        const share = Math.round((success / allMonthSuccess) * 100);
        extraText += `\nДоля успех: ${share}%`;
      }

      if (!extraText.trim()) return;

      const cell = reportSheet.getRange(groupIndex + 3, i + 2);
      const currentValue = cell.getRichTextValue();
      const currentText = currentValue?.getText() || "";
      const newText = currentText + extraText;

      const totalMatch = newText.match(/^(\d+)/);
      const boldEnd = totalMatch ? totalMatch[1].length : 0;

      const richVal = SpreadsheetApp.newRichTextValue().setText(newText);
      safeSetTextStyle(richVal, newText, boldEnd);
      cell.setRichTextValue(richVal.build());

      if (/Доля успех: \d+%/.test(newText)) {
        const bg = getCellBackgroundBySuccessRate(newText);
        cell.setBackground(bg);
      }
    });
  }
}

function convertMonthNameToKey(text) {
  const monthMap = {
    Jan: "01", Feb: "02", Mar: "03", Apr: "04", May: "05", Jun: "06",
    Jul: "07", Aug: "08", Sep: "09", Oct: "10", Nov: "11", Dec: "12"
  };
  const match = text.match(/(\w{3})\.(\d{2})/);
  if (!match) return null;
  const [_, mmm, yy] = match;
  const mm = monthMap[mmm];
  const yyyy = "20" + yy;
  return `${yyyy}-${mm}`;
}

function getCellBackgroundBySuccessRate(text) {
  if (!text || !text.trim()) return "#F2F2F2";
  const match = text.match(/Доля успех: (\d+)%/);
  if (!match) return "#F2F2F2";
  const percent = parseInt(match[1]);
  if (percent > 20) return "#C6EFCE";
  if (percent > 10) return "#E2EFDA";
  if (percent > 0) return "#FFF2CC";
  return "#F2F2F2";
}


/// === ГЛАВНАЯ ФУНКЦИЯ ===
// (весь код предыдущий остаётся без изменений)

function styleFinalTable(reportSheet, headerRow, totalRowIndex) {
  const mergeCols = headerRow.length;
  const totalRow = reportSheet.getRange(totalRowIndex, 2, 1, headerRow.slice(1).length).getRichTextValues();
  reportSheet.insertRowBefore(1);
  reportSheet.getRange(1, 2, 1, headerRow.slice(1).length).setValues([headerRow.slice(1)]);

 // 🧹 Очистка
reportSheet.clear();

// 🟦 Заголовок в A1:A2
const now = new Date();
const formattedDate = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd.MM.yy HH:mm:ss");
const fullHeader = `Итого\n(успех, в работе, провал)\nДата обновления: ${formattedDate}`;
const headerBuilder = SpreadsheetApp.newRichTextValue().setText(fullHeader);

headerBuilder.setTextStyle(0, fullHeader.length,
  SpreadsheetApp.newTextStyle().setFontSize(10).setBold(true).setForegroundColor("white").build());
headerBuilder.setTextStyle(fullHeader.indexOf("успех"), fullHeader.indexOf("успех") + 6,
  SpreadsheetApp.newTextStyle().setFontSize(10).setBold(true).setForegroundColor("#00B050").build());
headerBuilder.setTextStyle(fullHeader.indexOf("в работе"), fullHeader.indexOf("в работе") + 9,
  SpreadsheetApp.newTextStyle().setFontSize(10).setBold(true).setForegroundColor("#FFC000").build());
headerBuilder.setTextStyle(fullHeader.indexOf("провал"), fullHeader.indexOf("провал") + 6,
  SpreadsheetApp.newTextStyle().setFontSize(10).setBold(true).setForegroundColor("#FF6F91").build());
headerBuilder.setTextStyle(fullHeader.indexOf("Дата обновления:"), fullHeader.length,
  SpreadsheetApp.newTextStyle().setFontSize(9).setForegroundColor("white").build());

reportSheet.getRange("A1:A2").merge().setRichTextValue(headerBuilder.build())
  .setBackground("#305496")
  .setHorizontalAlignment("center")
  .setVerticalAlignment("middle")
  .setWrap(true);

// 🟨 Заголовки месяцев (строка 1)
const boldHeaderRow = headerRow.slice(1).map(text =>
  SpreadsheetApp.newRichTextValue()
    .setText(text)
    .setTextStyle(0, text.length, SpreadsheetApp.newTextStyle().setBold(true).build())
    .build()
);
reportSheet.getRange(1, 2, 1, boldHeaderRow.length).setRichTextValues([boldHeaderRow]);

Logger.log("Тип totalRow:", Array.isArray(totalRow));
Logger.log("Вложенный массив? →", Array.isArray(totalRow[0]));
Logger.log("Тип первой ячейки:", typeof totalRow[0]);

// 🟩 Строка "Итого" (строка 3)
if (Array.isArray(totalRow[0])) {
  reportSheet.getRange(2, 2, 1, totalRow[0].length).setRichTextValues(totalRow);
} else {
  reportSheet.getRange(2, 2, 1, totalRow.length).setRichTextValues([totalRow]);
}



// 📊 Основная таблица с 3-й строки
rawOutput.forEach((entry, i) => {
  while (entry.richRow.length < headerRow.length) {
    entry.richRow.push(SpreadsheetApp.newRichTextValue().setText("").build());
  }
  if (entry.richRow.length > headerRow.length) {
    entry.richRow = entry.richRow.slice(0, headerRow.length);
  }

  reportSheet.getRange(i + 3, 1, 1, headerRow.length)
    .setRichTextValues([entry.richRow]);
});


  const monthLabels = headerRow.slice(1);
  for (let i = 0; i < monthLabels.length; i++) {
    const [mon] = monthLabels[i].split(".");
    const monthNum = {
      Jan: 1, Feb: 2, Mar: 3,
      Apr: 4, May: 5, Jun: 6,
      Jul: 7, Aug: 8, Sep: 9,
      Oct: 10, Nov: 11, Dec: 12
    }[mon];

    let bg = "#FFFFFF", fg = "#000000";
    if ([1, 2, 3].includes(monthNum)) {
      bg = "#D9D2E9"; fg = "#5C3399"; // Q1 — сиреневый
    } else if ([4, 5, 6].includes(monthNum)) {
      bg = "#F4CCCC"; fg = "#990000"; // Q2 — розовый
    } else if ([7, 8, 9].includes(monthNum)) {
      bg = "#CFE2F3"; fg = "#0B5394"; // Q3 — голубой
    } else if ([10, 11, 12].includes(monthNum)) {
      bg = "#D9EAD3"; fg = "#274E13"; // Q4 — зелёный
    }

    const col = i + 2;
    const monthText = monthLabels[i];
    const richMonth = SpreadsheetApp.newRichTextValue()
      .setText(monthText)
      .setTextStyle(0, monthText.length, SpreadsheetApp.newTextStyle().setBold(true).build())
      .build();

// 🟦 Новый цвет для Q4
if ([10, 11, 12].includes(monthNum)) {
  bg = "#EAD1DC"; // яркий розово-сиреневый
  fg = "#783F61"; // насыщенный тёмно-сиреневый
}

// Жирная граница между кварталами
const isQuarterEnd = [3, 6, 9, 12].includes(monthNum);
const borderStyle = isQuarterEnd ? SpreadsheetApp.BorderStyle.THICK : SpreadsheetApp.BorderStyle.SOLID;

reportSheet.getRange(1, col)
  .setRichTextValue(richMonth)
  .setBackground(bg).setFontColor(fg)
  .setBorder(true, true, true, isQuarterEnd, null, null, "black", isQuarterEnd ? SpreadsheetApp.BorderStyle.THICK : SpreadsheetApp.BorderStyle.SOLID);

reportSheet.getRange(2, col)
  .setBackground(bg).setFontColor(fg)
  .setBorder(true, true, true, isQuarterEnd, null, null, "black", isQuarterEnd ? SpreadsheetApp.BorderStyle.THICK : SpreadsheetApp.BorderStyle.SOLID);

  }

  // Последний столбец — "Весь период" — всегда серый полностью
  const lastCol = headerRow.length;
  const lastColRange = reportSheet.getRange(1, lastCol, reportSheet.getLastRow(), 1);
  lastColRange.setBackground("#DDDDDD").setFontColor("black");

  // Автоширина и автовысота + ширина первой колонки
  reportSheet.setColumnWidth(1, 188);
  for (let i = 1; i <= reportSheet.getLastRow(); i++) {
    reportSheet.setRowHeight(i, 1); // авто
  }

  // Окраска колонки "Источник"
  const groupRange = reportSheet.getRange(3, 1, reportSheet.getLastRow() - 2, 1);
  groupRange.setBackground("#DAE3F3")
            .setFontWeight("bold")
            .setVerticalAlignment("middle");

  // === Границы всех ячеек + выравнивание по верхнему краю ===
  const allRange = reportSheet.getRange(1, 1, reportSheet.getLastRow(), reportSheet.getLastColumn());
  allRange.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
  allRange.setVerticalAlignment("top");
}
}