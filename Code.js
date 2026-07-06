// ─── Menu ────────────────────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('AdMaven')
    .addItem('Query Builder', 'showAdMavenSidebar')
    .addToUi();
}

function showAdMavenSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('AdMaven Query Builder');
  SpreadsheetApp.getUi().showSidebar(html);
}

// ─── Sidebar server-side functions ───────────────────────────────────────────

function getAdMavenOptions(type) {
  const config = getConfig();
  const admaven = new AdMaven(config.apiKey);
  return admaven.getOptions(type);
}

function getAdMavenCampaigns() {
  const config = getConfig();
  const admaven = new AdMaven(config.apiKey);
  return admaven.getAllCampaigns('name,id');
}

function runCustomQuery(params) {
  const config = getConfig();
  const admaven = new AdMaven(config.apiKey);

  let data = admaven.getReport(
    params.fromDate,
    params.toDate,
    params.groupBy,
    params.columns,
    params.filters || {}
  );

  // Compute CPA when both cost and conversions are present in columns
  const cols = params.columns || [];
  const hasCost = cols.includes('cost');
  const hasConversions = cols.includes('conversions');
  const hasRedirects = cols.includes('redirects');
  const computeCpa = hasCost && hasConversions;
  const computeCpm = hasCost && hasRedirects;
  if (computeCpa || computeCpm) {
    data = data.map(function(record) {
      const cost = parseFloat(record['cost'] || 0);
      const extra = {};
      if (computeCpa) {
        const conversions = parseFloat(record['conversions'] || 0);
        extra.cpa = conversions > 0 ? cost / conversions : cost;
      }
      if (computeCpm) {
        const redirects = parseFloat(record['redirects'] || 0);
        extra.cpm = redirects > 0 ? (cost / redirects) * 1000 : 0;
      }
      return Object.assign({}, record, extra);
    });
  }

  // Apply result filters (post-processing, server-side)
  if (params.resultFilters && params.resultFilters.length > 0) {
    data = data.filter(function(record) {
      return params.resultFilters.every(function(f) {
        if (!f.field || f.value === '' || f.value === null || f.value === undefined) return true;
        const val = parseFloat(record[f.field]);
        const threshold = parseFloat(f.value);
        if (isNaN(val) || isNaN(threshold)) return true;
        switch (f.op) {
          case '>':  return val > threshold;
          case '<':  return val < threshold;
          case '>=': return val >= threshold;
          case '<=': return val <= threshold;
          case '=':  return val === threshold;
          case '!=': return val !== threshold;
          default:   return true;
        }
      });
    });
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  // Insert computed columns: CPA after 'cost', CPM after 'cpa' (or after 'cost' if no CPA)
  let outputColumns = cols.slice();
  if (computeCpa) {
    const costIdx = outputColumns.indexOf('cost');
    if (costIdx >= 0) {
      outputColumns.splice(costIdx + 1, 0, 'cpa');
    } else {
      outputColumns.push('cpa');
    }
  }
  if (computeCpm) {
    const cpaIdx = outputColumns.indexOf('cpa');
    if (cpaIdx >= 0) {
      outputColumns.splice(cpaIdx + 1, 0, 'cpm');
    } else {
      const costIdx = outputColumns.indexOf('cost');
      if (costIdx >= 0) {
        outputColumns.splice(costIdx + 1, 0, 'cpm');
      } else {
        outputColumns.push('cpm');
      }
    }
  }

  writeCustomQueryToSheet(sheet, data, params.groupBy, outputColumns);

  return { rowCount: data.length, sheetName: sheet.getName() };
}

function writeCustomQueryToSheet(sheet, rows, groupBy, columns) {
  // Build ordered header list: groupBy fields first, then columns
  // (deduplicate in case campaign_id appears in both)
  const seen = new Set();
  const headers = [];
  groupBy.concat(columns).forEach(function(key) {
    if (!seen.has(key)) {
      seen.add(key);
      headers.push(key);
    }
  });

  // Clear all existing content
  sheet.clearContents();

  // Write headers to row 1
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Write data rows
  if (rows.length > 0) {
    const dataRows = rows.map(function(record) {
      return headers.map(function(key) {
        const val = record[key];
        return val !== undefined && val !== null ? val : '';
      });
    });
    sheet.getRange(2, 1, dataRows.length, headers.length).setValues(dataRows);
  }

  // Update timestamp
  sheet.getRange(1, headers.length + 2).setValue(`LAST UPDATE: ${formatDateTime(new Date())}`);
}

// ─── Config ──────────────────────────────────────────────────────────────────

function getConfig() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CONFIG");
  return {
    enable: sheet.getRange("B1").getValue().toString().trim().toLowerCase(),
    apiKey: sheet.getRange("B2").getValue().toString().trim(),
    today: sheet.getRange("B3").getValue(), // Date object or timestamp from spreadsheet
    groupBy: sheet
      .getRange("B4")
      .getValue()
      .toString()
      .trim()
      .split(",")
      .map((s) => s.trim()),
    columns: sheet
      .getRange("B5")
      .getValue()
      .toString()
      .trim()
      .split(",")
      .map((s) => s.trim()),
  };
}

function formatDate(date) {
  const d = new Date(date);
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}

function getDateRange(sheetName, today) {
  const d = new Date(today);
  // Normalize to midnight
  d.setHours(0, 0, 0, 0);

  switch (sheetName) {
    case "TODAY": {
      const from = new Date(d);
      return { fromDate: formatDate(from), toDate: formatDate(from) };
    }
    case "YESTERDAY": {
      const from = new Date(d);
      from.setDate(from.getDate() - 1);
      return { fromDate: formatDate(from), toDate: formatDate(from) };
    }
    case "H-2":
    case "H-3":
    case "H-4": {
      const daysBack = parseInt(sheetName.split("-")[1], 10);
      const day = new Date(d);
      day.setDate(day.getDate() - daysBack);
      return { fromDate: formatDate(day), toDate: formatDate(day) };
    }
    case "LAST7DAY": {
      const from = new Date(d);
      from.setDate(from.getDate() - 6);
      return { fromDate: formatDate(from), toDate: formatDate(d) };
    }
    case "LAST30DAY": {
      const from = new Date(d);
      from.setDate(from.getDate() - 29);
      return { fromDate: formatDate(from), toDate: formatDate(d) };
    }
    case "THISMONTH": {
      const from = new Date(d.getFullYear(), d.getMonth(), 1);
      return { fromDate: formatDate(from), toDate: formatDate(d) };
    }
    case "LASTMONTH": {
      const from = new Date(d.getFullYear(), d.getMonth() - 1, 1);
      const to = new Date(d.getFullYear(), d.getMonth(), 0);
      return { fromDate: formatDate(from), toDate: formatDate(to) };
    }
    default:
      throw new Error(`Unknown sheet name: ${sheetName}`);
  }
}

// ─── Row builder ─────────────────────────────────────────────────────────────

// Fixed sheet column order: A=ID, B=NAME, C=GEO, D=REDIRECT, E=CONV, F=COST, G=CPA
function buildRow(record, nameMap) {
  const id = record["campaign_id"] || "";
  const cost = parseFloat(record["cost"] || 0);
  const conversions = parseFloat(record["conversions"] || 0);
  const cpa = conversions > 0 ? cost / conversions : cost;

  return [
    id,
    (nameMap && nameMap[id]) || "",
    record["country_code"] || "",
    record["redirects"] || 0,
    conversions,
    cost,
    cpa,
  ];
}

function formatDateTime(date) {
  const d = new Date(date);
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  const hh = String(d.getHours()).padStart(2, "0");
  const ii = String(d.getMinutes()).padStart(2, "0");
  const ss = String(d.getSeconds()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd} ${hh}:${ii}:${ss}`;
}

// ─── Sheet writer ─────────────────────────────────────────────────────────────

function writeToSheet(sheetName, rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" not found`);
  }

  // Clear only A2:G (columns 1-7), leave H untouched except H1
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 7).clearContent();
  }

  // Write data rows starting from row 2, columns A-G only
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 7).setValues(rows);
  }

  // Update H1 with last update timestamp
  sheet.getRange("J1").setValue(`LAST UPDATE: ${formatDateTime(new Date())}`);
}

// ─── Update functions (schedule-triggered) ───────────────────────────────────

function updateReport(sheetName) {
  const config = getConfig();
  if (config.enable !== "y") return;

  const { fromDate, toDate } = getDateRange(sheetName, config.today);
  const admaven = new AdMaven(config.apiKey);
  const data = admaven.getReport(
    fromDate,
    toDate,
    config.groupBy,
    config.columns,
  );

  // Fetch campaign data for all unique IDs in the report, build id→name map
  const uniqueIds = [
    ...new Set(data.map((r) => r["campaign_id"]).filter(Boolean)),
  ];
  const campaigns = admaven.getCampaigns(uniqueIds, "name,id");
  const nameMap = campaigns.reduce((map, c) => {
    map[c.id] = c.name || "";
    return map;
  }, {});

  const rows = data.map((record) => buildRow(record, nameMap));

  writeToSheet(sheetName, rows);
}

function updateToday() {
  updateReport("TODAY");
}
function updateYesterday() {
  updateReport("YESTERDAY");
}
function updateH2() {
  updateReport("H-2");
}
function updateH3() {
  updateReport("H-3");
}
function updateH4() {
  updateReport("H-4");
}
function updateLast7Day() {
  updateReport("LAST7DAY");
}
function updateLast30Day() {
  updateReport("LAST30DAY");
}
function updateThisMonth() {
  updateReport("THISMONTH");
}
function updateLastMonth() {
  updateReport("LASTMONTH");
}

function updateDaily() {
  const config = getConfig();
  if (config.enable !== "y") return;

  const today = new Date(config.today);
  today.setHours(0, 0, 0, 0);
  const from = new Date(today);
  from.setDate(from.getDate() - 29);

  const admaven = new AdMaven(config.apiKey);
  const data = admaven.getReport(
    formatDate(from),
    formatDate(today),
    ["report_date"],
    ["redirects", "conversions", "cost"],
  );

  const rows = data
    .map((record) => {
      const cost = parseFloat(record["cost"] || 0);
      const conversions = parseFloat(record["conversions"] || 0);
      const cpa = conversions > 0 ? cost / conversions : cost;
      return [
        record["report_date"] || "",
        record["redirects"] || 0,
        conversions,
        cost,
        cpa,
      ];
    })
    .sort((a, b) => (a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0));

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("DAILY");
  if (!sheet) throw new Error('Sheet "DAILY" not found');

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 5).clearContent();
  }
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 5).setValues(rows);
  }

  sheet.getRange("G1").setValue(`LAST UPDATE: ${formatDateTime(new Date())}`);
}

function updateAllReports() {
  const config = getConfig();
  if (config.enable !== "y") return;

  const sheets = [
    "TODAY",
    "YESTERDAY",
    "H-2",
    "H-3",
    "H-4",
    "LAST7DAY",
    "LAST30DAY",
    "THISMONTH",
    "LASTMONTH",
  ];
  sheets.forEach((name) => updateReport(name));
}
