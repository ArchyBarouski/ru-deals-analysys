// Google Apps Script для Google Sheets
// Скрипт для подготовки разметки и отчета по сделкам на основе выгрузки Nemind

function normalizeSource(rawSource, description, dataMap) {
  let source = String(rawSource || "").trim();
  if (/Оффлайн\s*\/\s*Offline/i.test(source)) {
    return "Оффлайн / Offline";
  } else if (/Существующий клиент/i.test(source)) {
    return "Существующий клиент";
  }
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

function generateDealsReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("Nemind_export");
  const dataSheet = ss.getSheetByName("Data");
  const dataMap = {};

  if (dataSheet) {
    const dataValues = dataSheet.getRange("A2:C" + dataSheet.getLastRow()).getValues();
    dataValues.forEach(row => {
      const key = String(row[0] || "").replace(/[^\d]/g, "").trim();
      const val = String(row[2] || "").trim();
      if (key) dataMap[key] = val;
    });
  }

  const markupSheetName = "Разметка источников";
  let markupSheet = ss.getSheetByName(markupSheetName);
  const previousGroups = {};

  if (markupSheet) {
    const previousData = markupSheet.getRange(2, 1, markupSheet.getLastRow() - 1, 5).getValues();
    previousData.forEach(row => {
      const key = row.slice(0, 4).join("|");
      previousGroups[key] = row[4];
    });
    markupSheet.clear();
  } else {
    markupSheet = ss.insertSheet(markupSheetName);
  }

  const rulesSheetName = "Правила трансформации";
  let rulesSheet = ss.getSheetByName(rulesSheetName);
  if (!rulesSheet) {
    rulesSheet = ss.insertSheet(rulesSheetName);
    rulesSheet.getRange(1, 1, 1, 4).setValues([
      ["Дата", "Поле", "Условие / Значение", "Описание преобразования"]
    ]);
  }

  const rulesHeaderRange = rulesSheet.getRange("1:1");
  rulesHeaderRange.setFontWeight("bold");
  rulesSheet.setFrozenRows(1);
  rulesSheet.autoResizeColumns(1, 4);

  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

  function upsertRule(field, condition, description) {
    const rulesData = rulesSheet.getRange(2, 1, rulesSheet.getLastRow() - 1, 4).getValues();
    for (let i = 0; i < rulesData.length; i++) {
      if (rulesData[i][1] === field && rulesData[i][2] === condition && rulesData[i][3] === description) {
        rulesSheet.getRange(i + 2, 1).setValue(today);
        return;
      }
    }
    rulesSheet.appendRow([today, field, condition, description]);
  }

  upsertRule("Описание источника", "если содержит номер телефона — перенести в 'Источник', иначе удалить", "Очистка всех значений столбца B, перенос номера телефона в столбец A");
  upsertRule("Источник", "если содержит 'Оффлайн / Offline'", "Приведение всех значений к 'Оффлайн / Offline'");
  upsertRule("Источник", "если содержит 'Существующий клиент'", "Приведение всех значений к 'Существующий клиент'");
  upsertRule("Источник", "если содержит номер телефона и он есть в Data", "Заменить только номер телефона внутри источника на соответствующий источник из Data");

  const fullSourceData = sourceSheet.getDataRange().getValues();
  const results = [];

  for (let i = 1; i < fullSourceData.length; i++) {
    const row = fullSourceData[i];
    const normalized = normalizeSource(row[9], row[10], dataMap);
    results.push([
      normalized,
      String(row[25] || "").trim(),
      String(row[26] || "").trim(),
      String(row[27] || "").trim()
    ]);
  }

  const seen = new Set();
  const uniqueResults = results.filter(r => {
    const key = r.join("|");
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });

  const enrichedResults = uniqueResults.map(row => {
    const key = row.join("|");
    const group = previousGroups[key] || "";
    return [...row, group];
  });

  const headers = ["Источник", "UTM Source", "UTM Medium", "UTM Campaign", "Группа"];
  markupSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  markupSheet.getRange(2, 1, enrichedResults.length, headers.length).setValues(enrichedResults);
  markupSheet.getRange("1:1").setFontWeight("bold");
  markupSheet.setFrozenRows(1);
  markupSheet.autoResizeColumns(1, headers.length);

  const sourceMap = {};
  const uniqueSourceValues = [...new Set(enrichedResults.map(row => row[0]))];
  const pastelColors = [
    "#fff2cc", "#d9ead3", "#fce5cd", "#e2efda", "#d0e0e3", "#f4cccc", "#ead1dc", "#cfe2f3",
    "#fde9d9", "#d9d2e9", "#e6b8af", "#f9cb9c", "#f6f5a4", "#c9daf8", "#d4e9e2"
  ];
  let colorIndex = 0;
  uniqueSourceValues.forEach(src => {
    sourceMap[src] = pastelColors[colorIndex % pastelColors.length];
    colorIndex++;
  });

  for (let i = 0; i < enrichedResults.length; i++) {
    const sourceVal = enrichedResults[i][0];
    const bgColor = sourceMap[sourceVal];
    markupSheet.getRange(i + 2, 1, 1, headers.length).setBackgrounds([
      Array(headers.length).fill(bgColor)
    ]);
  }

  const validationRange = dataSheet.getRange("E2:E" + dataSheet.getLastRow());
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(validationRange, true)
    .setAllowInvalid(true)
    .build();
  markupSheet.getRange(2, 5, enrichedResults.length).setDataValidation(rule);

  markupSheet.sort(1);
  SpreadsheetApp.flush();
}
