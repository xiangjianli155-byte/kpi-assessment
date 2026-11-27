import * as XLSX from "https://cdn.sheetjs.com/xlsx-0.20.3/package/xlsx.mjs";

const TOPICS = ["INV", "EO"];
const PANELS = ["INV", "EO", "AGING"];
const COLORS = { red: "#c63d49", green: "#18804f", gray: "#8a98ac", blue: "#4777b9", ink: "#263851" };
const PLANNER_NAMES = {
  "li chao": "Li Chao", "she xiaolei": "She xiaolei", "xin mei": "Xin mei",
  "jessie jin": "Jessie Jin", "zhang mingshuai": "Zhang Mingshuai",
  "li xiang": "Li Xiang", lixiang: "Li Xiang"
};
const state = { records: { INV: [], EO: [], AGING: [] }, grain: {}, start: {}, end: {}, planner: {}, lower: {}, charts: {}, trendCharts: {} };

function esc(value) {
  return String(value == null ? "" : value).replace(/[&<>"']/g, function (char) {
    return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[char];
  });
}
function clean(value) {
  if (value == null || value === "") return "";
  const text = String(value).trim();
  return /^(nan|none|nat|#n\/a|#value!)$/i.test(text) ? "" : text;
}
function normalizePlanner(value) {
  const text = clean(value);
  if (!text) return "未分配";
  if (text.indexOf("/") >= 0) return "多人负责";
  return PLANNER_NAMES[text.toLowerCase()] || text;
}
function normalizeWHS(value) {
  const text = clean(value);
  if (!text) return "";
  const number = Number(text);
  return Number.isFinite(number) && Number.isInteger(number) ? String(number) : text;
}
function toNumber(value) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (value == null || String(value).trim() === "") return null;
  let text = String(value).trim().replace(/[,$￥¥\s]/g, "");
  const negative = /^\(.*\)$/.test(text);
  text = text.replace(/[()]/g, "");
  const number = Number(text);
  return Number.isFinite(number) ? (negative ? -number : number) : null;
}
function isoWeekMonday(year, week) {
  const jan4 = new Date(Date.UTC(year, 0, 4));
  const weekday = jan4.getUTCDay() || 7;
  const monday = new Date(jan4.getTime() - (weekday - 1) * 86400000 + (week - 1) * 7 * 86400000);
  return monday.toISOString().slice(0, 10);
}
function parsePeriod(value, fallbackYear) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.getFullYear() + "-" + String(value.getMonth() + 1).padStart(2, "0") + "-" + String(value.getDate()).padStart(2, "0");
  }
  if (typeof value === "number" && value > 20000 && value < 80000) {
    const date = new Date(Date.UTC(1899, 11, 30) + Math.round(value) * 86400000);
    return date.toISOString().slice(0, 10);
  }
  const text = clean(value);
  if (!text) return null;
  let match = text.match(/(20\d{2})\s*[-_/]?\s*[Ww](\d{1,2})\b/);
  if (match) {
    const week = Number(match[2]);
    return week >= 1 && week <= 53 ? isoWeekMonday(Number(match[1]), week) : null;
  }
  match = text.match(/(20\d{2})\s*[-/.年]\s*(\d{1,2})\s*[-/.月]\s*(\d{1,2})/);
  if (match) return validDate(Number(match[1]), Number(match[2]), Number(match[3]));
  match = text.match(/(20\d{2})\s*[-/.年]\s*(\d{1,2})(?:月)?\s*$/);
  if (match) return monthEnd(Number(match[1]), Number(match[2]));
  match = text.match(/(?<!\d)(\d{1,2})\s*月/);
  if (match) return monthEnd(fallbackYear, Number(match[1]));
  match = text.match(/Q\s*([1-4])/i);
  if (match) return monthEnd(fallbackYear, Number(match[1]) * 3);
  const parsed = new Date(text);
  return Number.isNaN(parsed.getTime()) ? null : parsed.toISOString().slice(0, 10);
}
function validDate(year, month, day) {
  const date = new Date(Date.UTC(year, month - 1, day));
  return date.getUTCFullYear() === year && date.getUTCMonth() === month - 1 && date.getUTCDate() === day
    ? date.toISOString().slice(0, 10) : null;
}
function monthEnd(year, month) {
  if (month < 1 || month > 12) return null;
  return new Date(Date.UTC(year, month, 0)).toISOString().slice(0, 10);
}
function headerLabel(value) { return clean(value).toLowerCase(); }
function findColumn(headers, kind) {
  const labels = headers.map(headerLabel);
  if (kind === "planner") return labels.findIndex(function (label) { return label === "planner"; }) >= 0
    ? labels.lastIndexOf("planner") : labels.findIndex(function (label) { return label.indexOf("planner") >= 0; });
  if (kind === "whs") return labels.findIndex(function (label) { return label === "whs"; });
  if (kind === "target") {
    const exact = labels.map(function (label, index) { return label.indexOf("target") === 0 || label.indexOf("standard") === 0 || label.indexOf("目标") >= 0 || label.indexOf("标准") >= 0 ? index : -1; }).filter(function (index) { return index >= 0; });
    return exact.length ? exact[exact.length - 1] : -1;
  }
  if (kind === "actual") return labels.findIndex(function (label) { return ["actual", "act", "实际", "实际值", "实绩"].indexOf(label) >= 0 || label.indexOf("actual") >= 0 || label.indexOf("实际") >= 0; });
  if (kind === "date") return labels.findIndex(function (label) { return ["date", "week", "period", "time", "日期", "时间", "周次", "月份"].some(function (word) { return label.indexOf(word) >= 0; }); });
  return -1;
}
function headerRowFor(rows) {
  let best = 0, scoreBest = -1;
  for (let index = 0; index < Math.min(15, rows.length); index += 1) {
    const labels = (rows[index] || []).map(headerLabel);
    const score = (labels.some(function (x) { return x.indexOf("whs") >= 0; }) ? 3 : 0)
      + (labels.some(function (x) { return x.indexOf("planner") >= 0; }) ? 3 : 0)
      + (labels.some(function (x) { return x.indexOf("target") >= 0 || x.indexOf("目标") >= 0; }) ? 2 : 0)
      + (labels.some(function (x) { return x.indexOf("actual") >= 0 || x === "act" || x.indexOf("实际") >= 0; }) ? 1 : 0);
    if (score > scoreBest) { best = index; scoreBest = score; }
  }
  return scoreBest >= 6 ? best : -1;
}
function workbookYear(filename) {
  const match = filename.match(/\b(20\d{2})\b/);
  return match ? Number(match[1]) : new Date().getFullYear();
}
function parseSheet(rows, sheetName, filename, topic) {
  const headerIndex = headerRowFor(rows);
  if (headerIndex < 0) throw new Error("找不到 Planner 和 WHS 表头");
  const headers = rows[headerIndex] || [];
  const body = rows.slice(headerIndex + 1), stopAt = body.findIndex(function (row) { return /^(total|whs)$/i.test(clean((row || [])[findColumn(headers, "whs")])); });
  const dataRows = stopAt >= 0 ? body.slice(0, stopAt) : body;
  const plannerCol = findColumn(headers, "planner"), whsCol = findColumn(headers, "whs");
  const targetCol = findColumn(headers, "target"), dateCol = findColumn(headers, "date"), actualCol = findColumn(headers, "actual");
  if (plannerCol < 0 || whsCol < 0) throw new Error("找不到 Planner 或 WHS 列");
  const year = workbookYear(filename), dateColumns = [];
  for (let column = 0; column < headers.length; column += 1) {
    const period = isActualHeader(headers[column]) ? parsePeriod(headers[column], year) : null;
    if (period && dataRows.some(function (row) { return toNumber((row || [])[column]) != null; })) dateColumns.push({ column: column, period: period });
  }
  if (!dateColumns.length && headerIndex > 0) {
    const parent = rows[headerIndex - 1] || [];
    headers.forEach(function (label, column) {
      if (["act", "actual", "实际", "实际值", "实绩"].indexOf(headerLabel(label)) >= 0) {
        const period = parsePeriod(parent[column], year);
        if (period && dataRows.some(function (row) { return toNumber((row || [])[column]) != null; })) dateColumns.push({ column: column, period: period });
      }
    });
  }
  const records = [];
  dataRows.forEach(function (row, rowOffset) {
    row = row || [];
    const whs = normalizeWHS(row[whsCol]);
    if (!whs) return;
    const planner = normalizePlanner(row[plannerCol]);
    const target = targetCol >= 0 ? toNumber(row[targetCol]) : null;
    if (dateColumns.length) {
      dateColumns.forEach(function (info) {
        const actual = toNumber(row[info.column]);
        if (actual == null && target == null) return;
        records.push({ topic: topic, planner: planner, whs: whs, date: info.period, actual: actual, target: target, order: rowOffset * 1000 + info.column, source: sheetName });
      });
    } else if (actualCol >= 0) {
      const period = dateCol >= 0 ? parsePeriod(row[dateCol], year) : parsePeriod(filename, year);
      const actual = toNumber(row[actualCol]);
      if (!period || (actual == null && target == null)) return;
      records.push({ topic: topic, planner: planner, whs: whs, date: period, actual: actual, target: target, order: rowOffset, source: sheetName });
    }
  });
  if (!records.length) throw new Error("未识别到带日期、实际值或目标值的记录");
  return records;
}
function isActualHeader(value) {
  if (value instanceof Date) return true;
  return !/vs|gap|target|planner|aging|增|减|差|标准|目标/i.test(clean(value));
}
function parseAging(rows, fallbackDate, year) {
  const headerIndex = headerRowFor(rows);
  if (headerIndex < 0) return [];
  const headers = rows[headerIndex] || [], parent = headerIndex > 0 ? rows[headerIndex - 1] || [] : [];
  const plannerCol = findColumn(headers, "planner"), whsCol = findColumn(headers, "whs");
  if (plannerCol < 0 || whsCol < 0) return [];
  const body = rows.slice(headerIndex + 1);
  const stopAt = body.findIndex(function (row) { return /^(total|whs)$/i.test(clean((row || [])[whsCol])); });
  const dataRows = stopAt >= 0 ? body.slice(0, stopAt) : body;
  const columns = [];
  headers.forEach(function (label, column) {
    if (/aging\s*[>＞]\s*12/i.test(clean(label))) columns.push({ column: column, date: parsePeriod(parent[column], year) || fallbackDate });
  });
  const records = [];
  dataRows.forEach(function (row, offset) {
    row = row || [];
    const whs = normalizeWHS(row[whsCol]);
    if (!whs) return;
    const planner = normalizePlanner(row[plannerCol]);
    columns.forEach(function (info) {
      const actual = toNumber(row[info.column]);
      if (actual == null || !info.date) return;
      records.push({ topic: "AGING", planner: planner, whs: whs, date: info.date, actual: actual, target: null, order: offset * 1000 + info.column, source: "Aging >12" });
    });
  });
  return records;
}
function median(values) {
  const sorted = values.slice().sort(function (a, b) { return a - b; });
  if (!sorted.length) return 0;
  const middle = Math.floor(sorted.length / 2);
  return sorted.length % 2 ? sorted[middle] : (sorted[middle - 1] + sorted[middle]) / 2;
}
function latestUnique(records, keys) {
  const map = new Map();
  records.slice().sort(function (a, b) { return a.date.localeCompare(b.date) || a.order - b.order; }).forEach(function (row) {
    map.set(keys.map(function (key) { return row[key]; }).join("\u0000"), row);
  });
  return Array.from(map.values());
}
function money(value, signed) {
  if (value == null || !Number.isFinite(Number(value))) return "—";
  const amount = Number(value), abs = Math.abs(amount), rounded = Math.round(abs);
  const prefix = rounded === 0 ? "" : signed && amount > 0 ? "+" : amount < 0 ? "-" : "";
  const text = abs >= 1000 ? Math.round(abs / 1000).toLocaleString("en-US") + "K" : rounded.toLocaleString("en-US");
  return prefix + "$" + text;
}
function status(actual, target, lower) {
  if (actual == null || target == null) return "待更新";
  return lower ? (actual <= target ? "已达标" : "未达标") : (actual >= target ? "已达标" : "未达标");
}
function grainName(grain) { return grain === "year" ? "按年" : grain === "month" ? "按月" : "按周"; }
const PLANNER_COLORS = ["#3d74b8", "#1f9d6b", "#e8963a", "#d6455d", "#7a5fc0", "#1fa3b8", "#8a98ac"];
const PLANNER_COLOR_MAP = { "li chao": "#3d74b8", "she xiaolei": "#1f9d6b", "xin mei": "#e8963a", "jessie jin": "#d6455d", "zhang mingshuai": "#7a5fc0", "li xiang": "#1fa3b8", "多人负责": "#8a98ac", "未分配": "#b9c3d0" };
function plannerColor(name, index) { return PLANNER_COLOR_MAP[String(name).toLowerCase()] || PLANNER_COLORS[index % PLANNER_COLORS.length]; }
function hexA(hex, alpha) {
  const value = parseInt(hex.slice(1), 16);
  return "rgba(" + (value >> 16) + "," + ((value >> 8) & 255) + "," + (value & 255) + "," + alpha + ")";
}
function axisLabels(periods, grain) {
  const years = new Set(periods.map(function (period) { return period.slice(0, 4); }));
  return periods.map(function (period) {
    if (grain === "year") return period.slice(0, 4);
    if (grain === "month") return years.size > 1 ? period.slice(0, 7) : Number(period.slice(5, 7)) + "月";
    return years.size > 1 ? period : period.slice(5);
  });
}
function dateLabel(date, grain) {
  return grain === "year" ? date.slice(0, 4) : grain === "month" ? date.slice(0, 7) : date;
}
function defaultRange(records, grain) {
  const latest = records.map(function (row) { return row.date; }).sort().slice(-1)[0];
  if (!latest) return ["", ""];
  const date = new Date(latest + "T00:00:00Z");
  if (grain === "year") return [latest.slice(0, 4) + "-01-01", latest.slice(0, 4) + "-12-31"];
  if (grain === "month") return [latest.slice(0, 7) + "-01", monthEnd(date.getUTCFullYear(), date.getUTCMonth() + 1)];
  const weekday = date.getUTCDay() || 7, monday = new Date(date.getTime() - (weekday - 1) * 86400000);
  const sunday = new Date(monday.getTime() + 6 * 86400000);
  return [monday.toISOString().slice(0, 10), sunday.toISOString().slice(0, 10)];
}
function plannerOptions(topic) {
  return Array.from(new Set(state.records[topic].map(function (row) { return row.planner; }))).sort(function (a, b) { return a.localeCompare(b); });
}
function renderFilters(topic) {
  const root = document.getElementById("filters-" + topic);
  const planners = plannerOptions(topic), selected = state.planner[topic] || "全部 Planner";
  root.innerHTML = '<div class="field"><label for="planner-' + topic + '">Planner</label><select id="planner-' + topic + '"><option>全部 Planner</option>' + planners.map(function (name) { return '<option' + (selected === name ? " selected" : "") + '>' + esc(name) + "</option>"; }).join("") + '</select></div>'
    + '<div class="field"><label for="grain-' + topic + '">时间粒度</label><select id="grain-' + topic + '"><option value="week"' + (state.grain[topic] === "week" ? " selected" : "") + '>按周</option><option value="month"' + (state.grain[topic] === "month" ? " selected" : "") + '>按月</option><option value="year"' + (state.grain[topic] === "year" ? " selected" : "") + '>按年</option></select></div>'
    + '<div class="field date-field"><label>日期范围</label><div class="date-range"><input id="start-' + topic + '" type="date" value="' + esc(state.start[topic] || "") + '" aria-label="开始日期"><span>至</span><input id="end-' + topic + '" type="date" value="' + esc(state.end[topic] || "") + '" aria-label="结束日期"></div></div>'
    + '<div class="field"><label for="rule-' + topic + '">达标规则</label><select id="rule-' + topic + '"><option value="lower"' + (state.lower[topic] !== false ? " selected" : "") + '>实际值 ≤ 标准值</option><option value="higher"' + (state.lower[topic] === false ? " selected" : "") + '>实际值 ≥ 标准值</option></select></div>'
    + '<button class="button button-outline reset-filter" id="reset-' + topic + '" type="button">重置</button>';
  document.getElementById("planner-" + topic).addEventListener("change", function (event) { state.planner[topic] = event.target.value; renderTopic(topic); });
  document.getElementById("grain-" + topic).addEventListener("change", function (event) {
    state.grain[topic] = event.target.value;
    const range = defaultRange(state.records[topic], state.grain[topic]); state.start[topic] = range[0]; state.end[topic] = range[1]; renderTopic(topic);
  });
  ["start", "end"].forEach(function (key) { document.getElementById(key + "-" + topic).addEventListener("change", function (event) { state[key][topic] = event.target.value; renderTopic(topic); }); });
  document.getElementById("rule-" + topic).addEventListener("change", function (event) { state.lower[topic] = event.target.value === "lower"; renderTopic(topic); });
  document.getElementById("reset-" + topic).addEventListener("click", function () {
    state.planner[topic] = "全部 Planner"; state.grain[topic] = "week"; state.lower[topic] = true;
    const range = defaultRange(state.records[topic], "week"); state.start[topic] = range[0]; state.end[topic] = range[1]; renderTopic(topic);
  });
}
function statusClass(value) { return value === "已达标" ? "pass" : value === "未达标" ? "fail" : "wait"; }
function summaryMarkup(rows, lower) {
  const actuals = rows.map(function (row) { return row.actual; }).filter(function (value) { return value != null; });
  const targets = rows.map(function (row) { return row.target; }).filter(function (value) { return value != null; });
  const actual = actuals.length ? actuals.reduce(function (a, b) { return a + b; }, 0) : null;
  const target = targets.length ? targets.reduce(function (a, b) { return a + b; }, 0) : null;
  const gap = actual == null || target == null ? null : actual - target;
  const result = status(actual, target, lower);
  const cards = [
    ["实际值", money(actual), "所选范围内 WHS 最新记录"],
    ["标准值", money(target), "按 WHS 标准汇总"],
    ["差额", money(gap, true), "实际值 − 标准值"],
    ["考核结果", '<span class="status-chip ' + statusClass(result) + '">' + result + '</span>', "按当前达标规则"]
  ];
  return cards.map(function (item) { return '<article class="summary-card"><div class="summary-label">' + item[0] + '</div><div class="summary-value">' + item[1] + '</div><div class="summary-note">' + item[2] + '</div></article>'; }).join("");
}
function detailMarkup(rows, lower) {
  const sorted = rows.slice().sort(function (a, b) { return a.planner.localeCompare(b.planner) || a.whs.localeCompare(b.whs, undefined, { numeric: true }); });
  if (!sorted.length) return '<tbody><tr><td class="empty-table" colspan="7">所选范围没有 WHS 明细。</td></tr></tbody>';
  return '<thead><tr><th>Planner</th><th>WHS</th><th>期间</th><th class="money">实际值（USD）</th><th class="money">标准值（USD）</th><th class="money">与标准差额（USD）</th><th>状态</th></tr></thead><tbody>' + sorted.map(function (row) {
    const result = status(row.actual, row.target, lower), gap = row.actual == null || row.target == null ? null : row.actual - row.target;
    return '<tr><td>' + esc(row.planner) + '</td><td>' + esc(row.whs) + '</td><td>' + esc(row.date) + '</td><td class="money">' + money(row.actual) + '</td><td class="money">' + money(row.target) + '</td><td class="money ' + (result === "已达标" ? "gap-pass" : result === "未达标" ? "gap-fail" : "") + '">' + money(gap, true) + '</td><td class="state-cell"><span class="table-status ' + statusClass(result) + '">' + result + '</span></td></tr>';
  }).join("") + '</tbody>';
}
function groupedByPlanner(rows, lower) {
  const map = new Map();
  rows.forEach(function (row) {
    if (!map.has(row.planner)) map.set(row.planner, { planner: row.planner, actual: 0, target: 0, hasActual: false, hasTarget: false });
    const group = map.get(row.planner);
    if (row.actual != null) { group.actual += row.actual; group.hasActual = true; }
    if (row.target != null) { group.target += row.target; group.hasTarget = true; }
  });
  return Array.from(map.values()).map(function (row) {
    row.actual = row.hasActual ? row.actual : null; row.target = row.hasTarget ? row.target : null;
    row.status = status(row.actual, row.target, lower); return row;
  }).sort(function (a, b) { return a.planner.localeCompare(b.planner); });
}
function chartRows(rows, grain) {
  const map = new Map();
  rows.forEach(function (row) {
    const period = grain === "year" ? row.date.slice(0, 4) + "-01-01" : grain === "month" ? row.date.slice(0, 7) + "-01" : mondayOf(row.date);
    const key = [row.planner, row.whs, period].join("\u0000");
    const prior = map.get(key);
    if (!prior || row.date.localeCompare(prior.date) >= 0) map.set(key, Object.assign({}, row, { period: period }));
  });
  const groups = new Map();
  Array.from(map.values()).forEach(function (row) {
    const key = [row.planner, row.period].join("\u0000");
    if (!groups.has(key)) groups.set(key, { planner: row.planner, period: row.period, actual: 0, target: 0, hasActual: false, hasTarget: false });
    const group = groups.get(key);
    if (row.actual != null) { group.actual += row.actual; group.hasActual = true; }
    if (row.target != null) { group.target += row.target; group.hasTarget = true; }
  });
  return Array.from(groups.values()).map(function (row) {
    row.actual = row.hasActual ? row.actual : null; row.target = row.hasTarget ? row.target : null; return row;
  }).sort(function (a, b) { return a.period.localeCompare(b.period) || a.planner.localeCompare(b.planner); });
}
function mondayOf(day) {
  const date = new Date(day + "T00:00:00Z"), weekday = date.getUTCDay() || 7;
  return new Date(date.getTime() - (weekday - 1) * 86400000).toISOString().slice(0, 10);
}
function currencyAxis(value) { return money(value); }
function chartOptions(horizontal) {
  return {
    responsive: true, maintainAspectRatio: false, animation: false,
    plugins: {
      legend: { position: "top", align: "end", labels: { usePointStyle: true, boxWidth: 8, color: COLORS.ink, font: { size: 11 } } },
      tooltip: { callbacks: { label: function (context) { return context.dataset.label + ": " + money(context.parsed[horizontal ? "x" : "y"]); } } },
      datalabels: { display: function (context) { return context.dataset.label !== "标准值（USD）"; }, anchor: horizontal ? "end" : "end", align: horizontal ? "right" : "top", color: COLORS.ink, clamp: true, font: { size: 9, weight: "600" }, formatter: function (value) { return value == null ? "" : money(value); } }
    },
    scales: horizontal ? {
      x: { beginAtZero: true, grid: { color: "#edf1f5" }, ticks: { color: "#7a8799", callback: currencyAxis } },
      y: { grid: { display: false }, ticks: { color: COLORS.ink } }
    } : {
      x: { grid: { display: false }, ticks: { color: "#7a8799", maxRotation: 0, autoSkip: true } },
      y: { beginAtZero: true, grid: { color: "#edf1f5" }, ticks: { color: "#7a8799", callback: currencyAxis } }
    }
  };
}
function drawChart(topic, selectedRows, lower) {
  if (state.charts[topic]) { state.charts[topic].destroy(); state.charts[topic] = null; }
  const empty = document.getElementById("chart-empty-" + topic), canvas = document.getElementById("chart-" + topic);
  const planner = state.planner[topic] || "全部 Planner", grain = state.grain[topic] || "week";
  const allPlanner = planner === "全部 Planner";
  const chartData = allPlanner ? groupedByPlanner(selectedRows, lower) : chartRows(isWeeklyUnavailable(topic) ? [] : state.records[topic], grain).filter(function (row) { return row.planner === planner; });
  if (!chartData.length) { canvas.parentElement.hidden = true; empty.hidden = false; return; }
  canvas.parentElement.hidden = false; empty.hidden = true;
  let config;
  if (allPlanner) {
    const labels = chartData.map(function (row) { return row.planner; });
    config = { type: "bar", data: { labels: labels, datasets: [
      { label: "标准值（USD）", data: chartData.map(function (row) { return row.target; }), backgroundColor: COLORS.gray, borderRadius: 4, datalabels: { display: true, color: "#fff", formatter: function (value) { return value == null ? "" : money(value); } } },
      { label: "实际值（USD）", data: chartData.map(function (row) { return row.actual; }), backgroundColor: chartData.map(function (row) { return row.status === "未达标" ? COLORS.red : row.status === "已达标" ? COLORS.green : COLORS.blue; }), borderRadius: 4 }
    ] }, options: Object.assign(chartOptions(true), { indexAxis: "y", plugins: Object.assign(chartOptions(true).plugins, { legend: { position: "top", align: "end" } }) }) };
    document.getElementById("chart-title-" + topic).textContent = "各 Planner 实际值与标准值";
  } else {
    config = singleLineConfig(chartData, grain, lower);
    document.getElementById("chart-title-" + topic).textContent = planner + " 实际值与标准值趋势（" + grainName(grain) + "）";
  }
  if (window.ChartDataLabels && window.Chart) window.Chart.register(window.ChartDataLabels);
  state.charts[topic] = new window.Chart(canvas.getContext("2d"), config);
}
function isWeeklyUnavailable(topic) {
  const dates = Array.from(new Set(state.records[topic].map(function (row) { return row.date; }))).sort();
  const gaps = dates.slice(1).map(function (date, index) { return (new Date(date + "T00:00:00Z") - new Date(dates[index] + "T00:00:00Z")) / 86400000; });
  return (state.grain[topic] || "week") === "week" && dates.length > 1 && median(gaps) > 14;
}
function setChartDefaults() {
  if (!window.Chart) return;
  if (window.ChartDataLabels) window.Chart.register(window.ChartDataLabels);
  window.Chart.defaults.font.family = getComputedStyle(document.body).fontFamily || "sans-serif";
  window.Chart.defaults.color = "#7a8799";
}
const crosshairPlugin = {
  id: "crosshair",
  afterDatasetsDraw: function (chart) {
    const active = chart.tooltip && chart.tooltip.getActiveElements ? chart.tooltip.getActiveElements() : [];
    if (!active.length) return;
    const x = active[0].element.x, area = chart.chartArea, ctx = chart.ctx;
    ctx.save(); ctx.beginPath(); ctx.moveTo(x, area.top); ctx.lineTo(x, area.bottom);
    ctx.lineWidth = 1; ctx.strokeStyle = "rgba(38,56,81,.18)"; ctx.setLineDash([4, 4]); ctx.stroke(); ctx.restore();
  }
};
const endLabelsPlugin = {
  id: "endLabels",
  afterDatasetsDraw: function (chart, args, options) {
    if (!options || !options.enabled) return;
    const ctx = chart.ctx, family = window.Chart.defaults.font.family, items = [];
    chart.data.datasets.forEach(function (dataset, index) {
      if (!chart.isDatasetVisible(index)) return;
      let last = -1; dataset.data.forEach(function (value, position) { if (value != null) last = position; });
      const point = chart.getDatasetMeta(index).data[last];
      if (last < 0 || !point) return;
      items.push({ x: point.x, y: point.y, ly: point.y, color: dataset.borderColor, name: dataset.label, value: money(dataset.data[last]) });
    });
    items.sort(function (a, b) { return a.y - b.y; });
    for (let i = 1; i < items.length; i += 1) if (items[i].ly - items[i - 1].ly < 16) items[i].ly = items[i - 1].ly + 16;
    if (items.length) { const overflow = items[items.length - 1].ly - (chart.chartArea.bottom - 2); if (overflow > 0) items.forEach(function (item) { item.ly -= overflow; }); }
    ctx.save(); ctx.textBaseline = "middle"; ctx.textAlign = "left";
    items.forEach(function (item) {
      const tx = item.x + 13;
      if (Math.abs(item.ly - item.y) > 2) { ctx.beginPath(); ctx.moveTo(item.x + 6, item.y); ctx.lineTo(tx - 3, item.ly); ctx.globalAlpha = 0.5; ctx.lineWidth = 1; ctx.strokeStyle = item.color; ctx.stroke(); ctx.globalAlpha = 1; }
      ctx.font = "700 12px " + family; ctx.fillStyle = item.color; ctx.fillText(item.name, tx, item.ly);
      const width = ctx.measureText(item.name).width;
      ctx.font = "600 12px " + family; ctx.fillStyle = "#5b6b80"; ctx.fillText("  " + item.value, tx + width, item.ly);
    });
    ctx.restore();
  }
};
function lineScales() {
  return {
    x: { offset: true, grid: { display: false }, border: { display: false }, ticks: { color: "#7a8799", maxRotation: 0, autoSkip: true, font: { size: 12 } } },
    y: { beginAtZero: true, border: { display: false }, grid: { color: "#eef2f6" }, ticks: { color: "#7a8799", maxTicksLimit: 5, padding: 8, callback: currencyAxis, font: { size: 12 } } }
  };
}
function tooltipStyle(callbacks) {
  return { backgroundColor: "#ffffff", titleColor: COLORS.ink, bodyColor: "#42536a", borderColor: "#dbe3ec", borderWidth: 1, padding: 12, cornerRadius: 10, boxPadding: 5, usePointStyle: true,
    titleFont: { weight: "700", size: 12 }, bodyFont: { size: 12 }, itemSort: function (a, b) { return (b.parsed.y || 0) - (a.parsed.y || 0); }, callbacks: callbacks };
}
function gradientFill(color) {
  return function (context) {
    const area = context.chart.chartArea;
    if (!area) return hexA(color, 0.1);
    const gradient = context.chart.ctx.createLinearGradient(0, area.top, 0, area.bottom);
    gradient.addColorStop(0, hexA(color, 0.22)); gradient.addColorStop(1, hexA(color, 0));
    return gradient;
  };
}
function renderLegend(container, chart) {
  container.innerHTML = "";
  const sets = chart.data.datasets;
  if (!container || sets.length < 2) { container.hidden = true; return; }
  container.hidden = false;
  sets.forEach(function (dataset, index) {
    const chip = document.createElement("button"); chip.type = "button"; chip.className = "legend-chip";
    chip.innerHTML = '<i style="background:' + dataset.borderColor + '"></i><span>' + esc(dataset.label) + "</span>";
    chip.classList.toggle("off", !chart.isDatasetVisible(index));
    chip.addEventListener("click", function () {
      chart.setDatasetVisibility(index, !chart.isDatasetVisible(index));
      chip.classList.toggle("off", !chart.isDatasetVisible(index)); chart.update();
    });
    container.appendChild(chip);
  });
  const hint = document.createElement("span"); hint.className = "legend-hint"; hint.textContent = "点击名称可隐藏/显示"; container.appendChild(hint);
}
function changeText(now, prev) {
  if (now == null || prev == null) return "";
  const diff = now - prev, sign = diff > 0 ? "+" : diff < 0 ? "-" : "";
  return (diff > 0 ? "▲ " : diff < 0 ? "▼ " : "") + money(Math.abs(diff)) + (prev ? "（" + sign + Math.abs(diff / prev * 100).toFixed(1) + "%）" : "");
}
function plannerLineConfig(rows, grain, opts) {
  setChartDefaults();
  const periods = Array.from(new Set(rows.map(function (row) { return row.period; }))).sort();
  const planners = Array.from(new Set(rows.map(function (row) { return row.planner; }))).sort(function (a, b) { return a.localeCompare(b); });
  const lookup = new Map(rows.map(function (row) { return [row.planner + "\u0000" + row.period, row]; }));
  const single = planners.length === 1, manyPoints = periods.length > 12;
  const series = planners.map(function (name, index) {
    return { name: name, color: plannerColor(name, index), data: periods.map(function (period) { const row = lookup.get(name + "\u0000" + period); return row ? row.actual : null; }) };
  });
  if (opts.includeTotal && planners.length > 1) {
    series.push({ name: "合计", color: "#263851", dashed: true, hidden: true, data: periods.map(function (period) {
      return rows.filter(function (row) { return row.period === period; }).reduce(function (sum, row) { return sum + (row.actual || 0); }, 0);
    }) });
  }
  const measure = document.createElement("canvas").getContext("2d"); measure.font = "700 12px " + window.Chart.defaults.font.family;
  let widest = 0;
  const datasets = series.map(function (item) {
    const color = item.color, data = item.data;
    let last = -1; data.forEach(function (value, position) { if (value != null) last = position; });
    if (last >= 0 && !item.hidden) widest = Math.max(widest, measure.measureText(item.name + "  " + money(data[last])).width);
    const dataset = {
      label: item.name, data: data, borderColor: color, backgroundColor: single ? gradientFill(color) : color, fill: single, borderWidth: 2.5, spanGaps: true,
      cubicInterpolationMode: "monotone", pointHoverRadius: 7, pointBorderColor: color, pointBorderWidth: 2, hidden: !!item.hidden, borderDash: item.dashed ? [6, 5] : [],
      pointRadius: function (context) { return periods.length === 1 || context.dataIndex === last ? 5 : manyPoints ? 0 : 3; },
      pointBackgroundColor: function (context) { return context.dataIndex === last ? color : "#fff"; }
    };
    if (opts.labelAll) dataset.datalabels = { display: function (context) { return context.dataset.data[context.dataIndex] != null; }, align: "top", anchor: "end", offset: 4, color: "#42536a", font: { size: 11, weight: "700" }, formatter: function (value) { return money(value); } };
    return dataset;
  });
  const pad = opts.labelAll ? 16 : Math.round(Math.min(Math.max(widest, 60) + 40, (opts.width || 800) * 0.5));
  const fullLabels = periods.map(function (period) { return dateLabel(period, grain); });
  return {
    type: "line", plugins: [crosshairPlugin, endLabelsPlugin],
    data: { labels: axisLabels(periods, grain), datasets: datasets },
    options: {
      responsive: true, maintainAspectRatio: false, animation: false,
      layout: { padding: { top: opts.labelAll ? 22 : 12, right: pad } },
      interaction: { mode: "index", intersect: false },
      plugins: {
        legend: { display: false },
        datalabels: opts.labelAll ? {} : { display: false },
        endLabels: { enabled: !opts.labelAll },
        tooltip: tooltipStyle({
          title: function (items) { return items.length ? fullLabels[items[0].dataIndex] : ""; },
          label: function (context) { return " " + context.dataset.label + "  " + money(context.parsed.y); },
          afterLabel: function (context) {
            const lines = [], row = lookup.get(context.dataset.label + "\u0000" + periods[context.dataIndex]);
            if (opts.showTarget && row && row.target != null && row.actual != null) lines.push("   标准 " + money(row.target) + " · 差额 " + money(row.actual - row.target, true) + " · " + status(row.actual, row.target, opts.lower));
            if (opts.showChange) {
              const data = context.dataset.data; let prev = null;
              for (let i = context.dataIndex - 1; i >= 0; i -= 1) if (data[i] != null) { prev = data[i]; break; }
              const text = changeText(data[context.dataIndex], prev);
              if (text) lines.push("   较上期 " + text);
            }
            return lines;
          }
        })
      },
      scales: lineScales()
    }
  };
}
function singleLineConfig(chartData, grain, lower) {
  setChartDefaults();
  const periods = chartData.map(function (row) { return row.period; });
  const fullLabels = periods.map(function (period) { return dateLabel(period, grain); });
  const pointColors = chartData.map(function (row) { const result = status(row.actual, row.target, lower); return result === "未达标" ? COLORS.red : result === "已达标" ? COLORS.green : COLORS.blue; });
  const lastIndex = chartData.length - 1, labelEvery = chartData.length <= 10;
  return {
    type: "line", plugins: [crosshairPlugin],
    data: { labels: axisLabels(periods, grain), datasets: [
      { label: "标准值", data: chartData.map(function (row) { return row.target; }), borderColor: "#9aa7b8", backgroundColor: "#9aa7b8", borderDash: [6, 5], borderWidth: 2, pointRadius: 0, pointHoverRadius: 4, cubicInterpolationMode: "monotone", order: 2,
        datalabels: { display: function (context) { return context.dataIndex === lastIndex && context.dataset.data[lastIndex] != null; }, align: "bottom", anchor: "end", offset: 6, color: "#7a8799", font: { size: 11, weight: "600" }, formatter: function (value) { return "标准 " + money(value); } } },
      { label: "实际值", data: chartData.map(function (row) { return row.actual; }), borderColor: COLORS.blue, backgroundColor: gradientFill(COLORS.blue), fill: true, borderWidth: 3, spanGaps: true, cubicInterpolationMode: "monotone", order: 1,
        pointBackgroundColor: pointColors, pointBorderColor: "#fff", pointBorderWidth: 2, pointRadius: 5, pointHoverRadius: 7,
        datalabels: { display: function (context) { return context.dataset.data[context.dataIndex] != null && (labelEvery || context.dataIndex === lastIndex); }, align: "top", anchor: "end", offset: 4, color: COLORS.ink, font: { size: 11, weight: "700" }, formatter: function (value) { return money(value); } } }
    ] },
    options: {
      responsive: true, maintainAspectRatio: false, animation: false,
      layout: { padding: { top: 24, right: 16 } },
      interaction: { mode: "index", intersect: false },
      plugins: {
        legend: { position: "top", align: "start", labels: { usePointStyle: true, boxWidth: 8, boxHeight: 8, padding: 16, color: COLORS.ink, font: { size: 12, weight: "600" } } },
        tooltip: tooltipStyle({
          title: function (items) { return items.length ? fullLabels[items[0].dataIndex] : ""; },
          label: function (context) { return " " + context.dataset.label + "  " + money(context.parsed.y); },
          afterBody: function (items) {
            const row = items.length ? chartData[items[0].dataIndex] : null;
            return !row || row.target == null || row.actual == null ? "" : "差额 " + money(row.actual - row.target, true) + " · " + status(row.actual, row.target, lower);
          }
        })
      },
      scales: lineScales()
    }
  };
}
function drawTrend(topic, allPlanner) {
  const panel = document.getElementById("trend-panel-" + topic);
  if (!panel) return;
  const detail = document.getElementById("detail-panel-" + topic), legend = document.getElementById("legend-" + topic);
  panel.hidden = !allPlanner;
  if (detail) detail.hidden = allPlanner;
  if (legend) legend.hidden = true;
  if (state.trendCharts[topic]) { state.trendCharts[topic].destroy(); state.trendCharts[topic] = null; }
  if (!allPlanner) return;
  const grain = state.grain[topic] || "week", lower = state.lower[topic] !== false;
  const canvas = document.getElementById("trend-" + topic), empty = document.getElementById("trend-empty-" + topic);
  const unavailable = isWeeklyUnavailable(topic);
  const rows = unavailable ? [] : chartRows(state.records[topic], grain);
  document.getElementById("trend-title-" + topic).textContent = "各 Planner 实际值趋势（" + grainName(grain) + "）";
  if (!rows.length) {
    canvas.parentElement.hidden = true; empty.hidden = false;
    empty.textContent = unavailable ? "当前工作簿记录的是月度数据，请切换为按月或按年查看。" : "当前没有可绘制趋势的数据。";
    return;
  }
  canvas.parentElement.hidden = false; empty.hidden = true;
  state.trendCharts[topic] = new window.Chart(canvas.getContext("2d"), plannerLineConfig(rows, grain, { lower: lower, showTarget: true, labelAll: false, width: canvas.parentElement.clientWidth }));
  if (legend) renderLegend(legend, state.trendCharts[topic]);
}
function fullRange(records) {
  const dates = records.map(function (row) { return row.date; }).sort();
  return dates.length ? [dates[0], dates[dates.length - 1]] : ["", ""];
}
function renderAgingFilters() {
  const root = document.getElementById("filters-AGING");
  const planners = plannerOptions("AGING"), selected = state.planner.AGING || "全部 Planner", grain = state.grain.AGING || "month";
  const bounds = fullRange(state.records.AGING);
  root.innerHTML = '<div class="field"><label for="planner-AGING">Planner</label><select id="planner-AGING"><option>全部 Planner</option>' + planners.map(function (name) { return '<option' + (selected === name ? " selected" : "") + '>' + esc(name) + "</option>"; }).join("") + '</select></div>'
    + '<div class="field"><label for="grain-AGING">时间粒度</label><select id="grain-AGING">' + [["week", "按周"], ["month", "按月"], ["year", "按年"]].map(function (item) { return '<option value="' + item[0] + '"' + (grain === item[0] ? " selected" : "") + ">" + item[1] + "</option>"; }).join("") + '</select></div>'
    + '<div class="field date-field"><label>日期范围</label><div class="date-range"><input id="start-AGING" type="date" value="' + esc(state.start.AGING || "") + '" aria-label="开始日期"><span>至</span><input id="end-AGING" type="date" value="' + esc(state.end.AGING || "") + '" aria-label="结束日期"></div></div>'
    + '<button class="button button-outline reset-filter" id="reset-AGING" type="button">重置</button>';
  document.getElementById("planner-AGING").addEventListener("change", function (event) { state.planner.AGING = event.target.value; renderAging(); });
  document.getElementById("grain-AGING").addEventListener("change", function (event) { state.grain.AGING = event.target.value; renderAging(); });
  ["start", "end"].forEach(function (key) { document.getElementById(key + "-AGING").addEventListener("change", function (event) { state[key].AGING = event.target.value; renderAging(); }); });
  document.getElementById("reset-AGING").addEventListener("click", function () { state.planner.AGING = "全部 Planner"; state.grain.AGING = "month"; state.start.AGING = bounds[0]; state.end.AGING = bounds[1]; renderAging(); });
}
function agingSummary(rows, grain) {
  const totals = new Map();
  rows.forEach(function (row) { totals.set(row.period, (totals.get(row.period) || 0) + (row.actual || 0)); });
  const periods = Array.from(totals.keys()).sort(), values = periods.map(function (period) { return totals.get(period); });
  const n = periods.length, latest = n ? values[n - 1] : null;
  let peak = -1; values.forEach(function (value, index) { if (peak < 0 || value > values[peak]) peak = index; });
  function label(index) { return index >= 0 && index < n ? dateLabel(periods[index], grain) : ""; }
  function delta(from) {
    if (latest == null || from == null) return "—";
    const diff = latest - from, pct = from ? Math.abs(diff / from * 100).toFixed(1) + "%" : "";
    return '<span class="' + (diff > 0 ? "delta-up" : diff < 0 ? "delta-down" : "") + '">' + (diff > 0 ? "▲ " : diff < 0 ? "▼ " : "") + money(Math.abs(diff)) + (pct ? "<small>" + (diff > 0 ? "+" : diff < 0 ? "-" : "") + pct + "</small>" : "") + "</span>";
  }
  const more = n > 1;
  const cards = [
    ["最新 Aging >12", money(latest), n ? "所选范围最新一期 · " + label(n - 1) : "所选范围没有数据"],
    ["较上期", delta(more ? values[n - 2] : null), more ? "对比 " + label(n - 2) + " · 下降为绿色" : "需要至少 2 期数据"],
    ["较期初", delta(more ? values[0] : null), more ? "对比 " + label(0) + " · 下降为绿色" : "需要至少 2 期数据"],
    ["期间峰值", money(peak >= 0 ? values[peak] : null), peak >= 0 ? "出现在 " + label(peak) : "—"]
  ];
  const text = n ? "最新 " + label(n - 1) + "：" + money(latest) + (more ? "　较上期 " + (changeText(latest, values[n - 2]) || "—") : "") : "";
  return { html: cards.map(function (item) { return '<article class="summary-card"><div class="summary-label">' + item[0] + '</div><div class="summary-value">' + item[1] + '</div><div class="summary-note">' + item[2] + "</div></article>"; }).join(""), text: text };
}
function drawAging() {
  if (state.charts.AGING) { state.charts.AGING.destroy(); state.charts.AGING = null; }
  const planner = state.planner.AGING || "全部 Planner", grain = state.grain.AGING || "month", all = planner === "全部 Planner";
  const canvas = document.getElementById("chart-AGING"), empty = document.getElementById("chart-empty-AGING"), legend = document.getElementById("legend-AGING");
  legend.hidden = true;
  let start = state.start.AGING || "", end = state.end.AGING || "";
  if (start && end && start > end) { const swap = start; start = end; end = swap; }
  const unavailable = isWeeklyUnavailable("AGING");
  const inRange = state.records.AGING.filter(function (row) { return (!start || row.date >= start) && (!end || row.date <= end); });
  let rows = unavailable ? [] : chartRows(inRange, grain);
  if (!all) rows = rows.filter(function (row) { return row.planner === planner; });
  const summary = agingSummary(rows, grain);
  document.getElementById("summary-AGING").innerHTML = summary.html;
  state.agingText = summary.text;
  document.getElementById("chart-title-AGING").textContent = (all ? "各 Planner" : planner) + " Aging >12 个月趋势（" + grainName(grain) + "）";
  if (!rows.length) {
    canvas.parentElement.hidden = true; empty.hidden = false;
    empty.textContent = unavailable ? "当前数据是月度记录，请切换为按月或按年查看。" : "所选日期范围内没有 Aging >12 数据，请调整日期范围。";
    return;
  }
  canvas.parentElement.hidden = false; empty.hidden = true;
  state.charts.AGING = new window.Chart(canvas.getContext("2d"), plannerLineConfig(rows, grain, { lower: true, showTarget: false, showChange: true, includeTotal: all, labelAll: !all, width: canvas.parentElement.clientWidth }));
  renderLegend(legend, state.charts.AGING);
}
function renderAging() {
  if (!state.records.AGING.length) return;
  renderAgingFilters();
  const periods = Array.from(new Set(state.records.AGING.map(function (row) { return row.date; }))).sort();
  const note = document.getElementById("aging-note");
  note.hidden = periods.length > 1;
  if (periods.length === 1) note.querySelector("p").textContent = "目前 Aging >12 只有 " + periods[0] + " 这一期，所以每个 Planner 只显示一个点。以后每月在 INV 工作表新增一列“Aging >12个月”，并在该列第 2 行填写日期，曲线就会自动连成线。";
  drawAging();
}
function renderTopic(topic) {
  if (!state.records[topic].length) return;
  renderFilters(topic);
  const all = state.records[topic], start = state.start[topic], end = state.end[topic], planner = state.planner[topic] || "全部 Planner";
  let filtered = all.filter(function (row) { return (!start || row.date >= start) && (!end || row.date <= end); });
  if (planner !== "全部 Planner") filtered = filtered.filter(function (row) { return row.planner === planner; });
  const dates = Array.from(new Set(all.map(function (row) { return row.date; }))).sort();
  const gaps = dates.slice(1).map(function (date, index) { return (new Date(date + "T00:00:00Z") - new Date(dates[index] + "T00:00:00Z")) / 86400000; });
  const weeklyUnavailable = state.grain[topic] === "week" && dates.length > 1 && median(gaps) > 14;
  if (weeklyUnavailable) filtered = [];
  const latestRows = latestUnique(filtered, ["planner", "whs"]);
  document.getElementById("summary-" + topic).innerHTML = summaryMarkup(latestRows, state.lower[topic] !== false);
  document.getElementById("detail-" + topic).innerHTML = detailMarkup(latestRows, state.lower[topic] !== false);
  document.getElementById("detail-count-" + topic).textContent = latestRows.length + " 条记录";
  const notice = document.getElementById("time-note-" + topic);
  if (notice) { notice.hidden = !weeklyUnavailable; }
  const chartEmpty = document.getElementById("chart-empty-" + topic);
  if (weeklyUnavailable) chartEmpty.textContent = "当前工作簿记录的是月度数据，请切换为按月查看。";
  else chartEmpty.textContent = "当前筛选范围内没有可比较的数据。";
  drawChart(topic, filtered, state.lower[topic] !== false);
  drawTrend(topic, planner === "全部 Planner");
}
function showToast(message) {
  const toast = document.createElement("div"); toast.className = "toast"; toast.textContent = message;
  document.body.appendChild(toast); window.setTimeout(function () { toast.remove(); }, 3200);
}
function activateTopic(topic) {
  document.querySelectorAll(".topic-tab").forEach(function (tab) { tab.classList.toggle("active", tab.dataset.topic === topic); });
  PANELS.forEach(function (item) {
    const panel = document.getElementById("topic-" + item), active = item === topic;
    panel.classList.toggle("active", active); panel.hidden = !active;
    if (active) window.setTimeout(function () { if (state.charts[item]) state.charts[item].resize(); if (state.trendCharts[item]) state.trendCharts[item].resize(); }, 0);
  });
  document.body.dataset.activeTopic = topic;
}
async function exportAgingPpt(deck) {
  const planner = state.planner.AGING || "全部 Planner", grain = state.grain.AGING || "month";
  const canvas = document.getElementById("chart-AGING");
  const slide = deck.addSlide(); slide.background = { color: "FFFFFF" };
  slide.addText("Movex KPI · Aging >12 个月", { x: 0.55, y: 0.3, w: 12.1, h: 0.45, fontSize: 22, bold: true, color: "18263B" });
  slide.addText("Planner：" + planner + "    粒度：" + grainName(grain) + "    范围：" + (state.start.AGING || "—") + " 至 " + (state.end.AGING || "—") + (state.agingText ? "    " + state.agingText : ""), { x: 0.6, y: 0.85, w: 12.1, h: 0.3, fontSize: 11, color: "77849A" });
  if (canvas && state.charts.AGING && canvas.parentElement.hidden === false) {
    const ratio = canvas.height / canvas.width; let w = 12.1, h = w * ratio;
    if (h > 5.4) { h = 5.4; w = h / ratio; }
    slide.addImage({ data: canvas.toDataURL("image/png", 1), x: 0.6 + (12.1 - w) / 2, y: 1.3, w: w, h: h });
  }
  slide.addText("金额单位：USD", { x: 10.5, y: 6.85, w: 2, h: 0.2, fontSize: 8, color: "77849A", align: "right" });
  await deck.writeFile({ fileName: "KPI_AGING_" + grain + ".pptx" });
}
async function exportPpt(topic) {
  const Pptx = window.PptxGenJS || window.pptxgen;
  if (!Pptx) { showToast("PowerPoint 导出组件加载失败，请刷新后重试。"); return; }
  const deck = new Pptx(); deck.layout = "LAYOUT_WIDE"; deck.author = "KPI 考核"; deck.subject = topic + " 专题 KPI"; deck.title = "KPI 考核";
  deck.theme = { headFontFace: "Aptos Display", bodyFontFace: "Aptos", lang: "zh-CN" };
  if (topic === "AGING") { await exportAgingPpt(deck); return; }
  const planner = state.planner[topic] || "全部 Planner", grain = state.grain[topic] || "week";
  const start = state.start[topic] || "", end = state.end[topic] || "";
  const filtered = state.records[topic].filter(function (row) { return (!start || row.date >= start) && (!end || row.date <= end) && (planner === "全部 Planner" || row.planner === planner); });
  const rows = latestUnique(filtered, ["planner", "whs"]), lower = state.lower[topic] !== false;
  const actual = rows.reduce(function (sum, row) { return sum + (row.actual == null ? 0 : row.actual); }, 0);
  const target = rows.reduce(function (sum, row) { return sum + (row.target == null ? 0 : row.target); }, 0);
  const result = status(rows.some(function (r) { return r.actual != null; }) ? actual : null, rows.some(function (r) { return r.target != null; }) ? target : null, lower);
  const slide = deck.addSlide(); slide.background = { color: "F4F7FB" };
  slide.addText("KPI 考核 · " + (topic === "INV" ? "On Hand INV" : "EO") + " 专题", { x: 0.55, y: 0.3, w: 12.1, h: 0.45, fontSize: 23, bold: true, color: "18263B" });
  slide.addText("Planner：" + planner + "    时间：" + start + " 至 " + end + "    粒度：" + grainName(grain), { x: 0.6, y: 0.85, w: 12, h: 0.3, fontSize: 11, color: "77849A" });
  const cards = [["实际值", money(rows.some(function (r) { return r.actual != null; }) ? actual : null)], ["标准值", money(rows.some(function (r) { return r.target != null; }) ? target : null)], ["差额", money(rows.length ? actual - target : null, true)], ["考核结果", result]];
  cards.forEach(function (item, index) {
    const x = 0.6 + index * 3.1; slide.addShape(deck.ShapeType.roundRect, { x: x, y: 1.25, w: 2.85, h: 0.78, rectRadius: 0.08, line: { color: "E1E7EF" }, fill: { color: index === 3 ? (result === "已达标" ? "E9F6EE" : result === "未达标" ? "FFF0F0" : "FFF6E6") : "FFFFFF" } });
    slide.addText(item[0], { x: x + 0.14, y: 1.36, w: 2.6, h: 0.2, fontSize: 10, color: "77849A" });
    slide.addText(item[1], { x: x + 0.14, y: 1.63, w: 2.6, h: 0.25, fontSize: 15, bold: true, color: "263851" });
  });
  const chartCanvas = document.getElementById("chart-" + topic);
  if (chartCanvas && chartCanvas.parentElement.hidden === false) slide.addImage({ data: chartCanvas.toDataURL("image/png", 1), x: 0.65, y: 2.2, w: 12, h: 3.6, transparency: 0 });
  slide.addText("金额单位：USD", { x: 10.5, y: 5.85, w: 2, h: 0.2, fontSize: 8, color: "77849A", align: "right" });
  const trendCanvas = document.getElementById("trend-" + topic);
  const showTrend = !!(trendCanvas && planner === "全部 Planner" && state.trendCharts[topic]);
  if (showTrend) {
    const trendSlide = deck.addSlide(); trendSlide.background = { color: "FFFFFF" };
    trendSlide.addText("各 Planner 实际值趋势（" + grainName(grain) + "）", { x: 0.55, y: 0.3, w: 12, h: 0.45, fontSize: 20, bold: true, color: "18263B" });
    const ratio = trendCanvas.height / trendCanvas.width; let w = 12.1, h = w * ratio;
    if (h > 5.4) { h = 5.4; w = h / ratio; }
    trendSlide.addImage({ data: trendCanvas.toDataURL("image/png", 1), x: 0.6 + (12.1 - w) / 2, y: 1.1, w: w, h: h });
    trendSlide.addText("金额单位：USD", { x: 10.5, y: 6.7, w: 2, h: 0.2, fontSize: 8, color: "77849A", align: "right" });
  }
  const chunks = [];
  for (let i = 0; i < rows.length; i += 16) chunks.push(rows.slice(i, i + 16));
  if (!chunks.length) chunks.push([]);
  (showTrend ? [] : chunks).forEach(function (chunk, page) {
    const detail = deck.addSlide(); detail.background = { color: "FFFFFF" };
    detail.addText("Planner × WHS 达标明细" + (chunks.length > 1 ? " · " + (page + 1) : ""), { x: 0.55, y: 0.3, w: 12, h: 0.45, fontSize: 20, bold: true, color: "18263B" });
    const data = [["Planner", "WHS", "期间", "实际值（USD）", "标准值（USD）", "差额（USD）", "状态"]];
    chunk.forEach(function (row) { const s = status(row.actual, row.target, lower); data.push([row.planner, row.whs, row.date, money(row.actual), money(row.target), money(row.actual == null || row.target == null ? null : row.actual - row.target, true), s]); });
    if (!chunk.length) data.push(["—", "—", "—", "—", "—", "—", "无记录"]);
    detail.addTable(data, { x: 0.55, y: 1.0, w: 12.25, h: 5.7, border: { type: "solid", color: "E5EAF1", pt: 0.5 }, fill: "FFFFFF", color: "42536A", fontSize: 9, margin: 0.08, rowH: 0.3, colW: [1.85, 1.1, 1.6, 1.9, 1.9, 1.95, 1.75],
      autoFit: false, bold: false, valign: "mid", breakLine: false, paraSpaceAfterPt: 0,
      rowOptions: { fill: "FFFFFF" },
      borderColor: "E5EAF1",
      /* header formatting is applied uniformly by pptxgenjs */
    });
  });
  const filename = "KPI_" + topic + "_" + (start || "all") + "_" + (end || "all") + ".pptx";
  await deck.writeFile({ fileName: filename });
}
function setupExports() {
  document.querySelectorAll("[data-export-pdf]").forEach(function (button) { button.addEventListener("click", function () { activateTopic(button.dataset.exportPdf); window.print(); }); });
  document.querySelectorAll("[data-export-ppt]").forEach(function (button) { button.addEventListener("click", function () { exportPpt(button.dataset.exportPpt).catch(function (error) { showToast("PowerPoint 导出失败：" + error.message); }); }); });
}
function setLoadError(message) {
  const notice = document.getElementById("load-error"); notice.hidden = false;
  notice.querySelector("strong").textContent = "KPI 数据暂时无法读取";
  notice.querySelector("p").textContent = message;
  const statusBadge = document.getElementById("data-status"); statusBadge.classList.add("error");
  statusBadge.lastChild.textContent = " 数据未就绪";
}
async function loadWorkbook() {
  const response = await fetch(new URL("../data/current_kpi.xlsx", import.meta.url), { cache: "no-store" });
  if (!response.ok) throw new Error("请先将 KPI Excel 放入 GitHub 仓库的 data 文件夹，并命名为 current_kpi.xlsx。");
  const workbook = XLSX.read(await response.arrayBuffer(), { type: "array", cellDates: true });
  if (!workbook.SheetNames.length) throw new Error("工作簿里没有可读取的工作表。");
  TOPICS.forEach(function (topic) {
    const sheetName = workbook.SheetNames.find(function (name) { return topic === "INV" ? name.toUpperCase().indexOf("INV") >= 0 : name.toUpperCase().indexOf("EO") >= 0 || name.toUpperCase().indexOf("E&O") >= 0; });
    if (!sheetName) return;
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, raw: true, defval: null, blankrows: false });
    state.records[topic] = parseSheet(rows, sheetName, "current_kpi.xlsx", topic);
    if (topic === "INV") {
      try {
        const latestInv = state.records.INV.map(function (row) { return row.date; }).sort().slice(-1)[0];
        state.records.AGING = parseAging(rows, latestInv, workbookYear("current_kpi.xlsx"));
      } catch (error) { console.error(error); state.records.AGING = []; }
    }
  });
  if (!state.records.INV.length && !state.records.EO.length) throw new Error("未找到名称包含 INV 或 EO 的工作表。");
  TOPICS.forEach(function (topic) {
    const records = state.records[topic]; state.grain[topic] = "week"; state.planner[topic] = "全部 Planner"; state.lower[topic] = true;
    if (records.length) { const range = defaultRange(records, "week"); state.start[topic] = range[0]; state.end[topic] = range[1]; }
    else { document.querySelector('[data-topic="' + topic + '"]').disabled = true; document.getElementById("topic-" + topic).hidden = true; }
  });
  state.grain.AGING = "month"; state.planner.AGING = "全部 Planner";
  { const span = fullRange(state.records.AGING); state.start.AGING = span[0]; state.end.AGING = span[1]; }
  if (!state.records.AGING.length) { document.querySelector('[data-topic="AGING"]').disabled = true; document.getElementById("topic-AGING").hidden = true; }
  const latest = state.records.INV.concat(state.records.EO).map(function (row) { return row.date; }).sort().slice(-1)[0];
  document.getElementById("source-label").textContent = latest ? "数据更新至 " + latest : "KPI 数据已加载";
  const statusBadge = document.getElementById("data-status"); statusBadge.classList.add("ready"); statusBadge.lastChild.textContent = " 数据已更新";
  document.getElementById("load-error").hidden = true;
  document.querySelectorAll(".topic-tab").forEach(function (tab) { tab.addEventListener("click", function () { activateTopic(tab.dataset.topic); }); });
  TOPICS.forEach(function (topic) { if (state.records[topic].length) renderTopic(topic); });
  renderAging();
  activateTopic(state.records.INV.length ? "INV" : "EO");
}
setupExports();
loadWorkbook().catch(function (error) {
  setLoadError(error.message || "读取 Excel 失败，请检查文件格式和工作表。");
  console.error(error);
});

