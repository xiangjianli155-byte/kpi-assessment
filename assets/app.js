import * as XLSX from "https://cdn.sheetjs.com/xlsx-0.20.3/package/xlsx.mjs";

const TOPICS = ["INV", "EO"];
const COLORS = { red: "#c63d49", green: "#18804f", gray: "#8a98ac", blue: "#4777b9", ink: "#263851" };
const PLANNER_NAMES = {
  "li chao": "Li Chao", "she xiaolei": "She xiaolei", "xin mei": "Xin mei",
  "jessie jin": "Jessie Jin", "zhang mingshuai": "Zhang Mingshuai",
  "li xiang": "Li Xiang", lixiang: "Li Xiang"
};
const state = { records: { INV: [], EO: [] }, grain: {}, start: {}, end: {}, planner: {}, lower: {}, charts: {} };

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
    const exact = labels.map(function (label, index) { return /^(target|target value|目标|目标值|标准|标准值)$/.test(label) || label.indexOf("目标") >= 0 ? index : -1; }).filter(function (index) { return index >= 0; });
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
  const plannerCol = findColumn(headers, "planner"), whsCol = findColumn(headers, "whs");
  const targetCol = findColumn(headers, "target"), dateCol = findColumn(headers, "date"), actualCol = findColumn(headers, "actual");
  if (plannerCol < 0 || whsCol < 0) throw new Error("找不到 Planner 或 WHS 列");
  const year = workbookYear(filename), dateColumns = [];
  for (let column = 0; column < headers.length; column += 1) {
    const period = parsePeriod(headers[column], year);
    if (period && rows.slice(headerIndex + 1).some(function (row) { return toNumber((row || [])[column]) != null; })) dateColumns.push({ column: column, period: period });
  }
  if (!dateColumns.length && headerIndex > 0) {
    const parent = rows[headerIndex - 1] || [];
    headers.forEach(function (label, column) {
      if (["act", "actual", "实际", "实际值", "实绩"].indexOf(headerLabel(label)) >= 0) {
        const period = parsePeriod(parent[column], year);
        if (period && rows.slice(headerIndex + 1).some(function (row) { return toNumber((row || [])[column]) != null; })) dateColumns.push({ column: column, period: period });
      }
    });
  }
  const records = [];
  rows.slice(headerIndex + 1).forEach(function (row, rowOffset) {
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
  const amount = Number(value), prefix = signed && amount > 0 ? "+" : amount < 0 ? "-" : "";
  return prefix + "$" + Math.abs(amount).toLocaleString("en-US", { maximumFractionDigits: 2 });
}
function status(actual, target, lower) {
  if (actual == null || target == null) return "待更新";
  return lower ? (actual <= target ? "已达标" : "未达标") : (actual >= target ? "已达标" : "未达标");
}
function dateLabel(date, grain) {
  return grain === "month" ? date.slice(0, 7) : date;
}
function defaultRange(records, grain) {
  const latest = records.map(function (row) { return row.date; }).sort().slice(-1)[0];
  if (!latest) return ["", ""];
  const date = new Date(latest + "T00:00:00Z");
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
    + '<div class="field"><label for="grain-' + topic + '">时间粒度</label><select id="grain-' + topic + '"><option value="week"' + (state.grain[topic] === "week" ? " selected" : "") + '>按周</option><option value="month"' + (state.grain[topic] === "month" ? " selected" : "") + '>按月</option></select></div>'
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
    const period = grain === "month" ? row.date.slice(0, 7) + "-01" : mondayOf(row.date);
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
function currencyAxis(value) { return "$" + Number(value).toLocaleString("en-US", { maximumFractionDigits: 0 }); }
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
  const chartData = allPlanner ? groupedByPlanner(selectedRows, lower) : chartRows(selectedRows, grain).filter(function (row) { return row.planner === planner; });
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
    const labels = chartData.map(function (row) { return dateLabel(row.period, grain); });
    const line = grain === "week" && chartData.length > 1;
    const pointColors = chartData.map(function (row) { const result = status(row.actual, row.target, lower); return result === "未达标" ? COLORS.red : result === "已达标" ? COLORS.green : COLORS.blue; });
    if (line) {
      config = { type: "line", data: { labels: labels, datasets: [
        { label: "标准值（USD）", data: chartData.map(function (row) { return row.target; }), borderColor: COLORS.gray, borderDash: [6, 5], borderWidth: 2, pointRadius: 2, datalabels: { display: true, align: "bottom", color: COLORS.gray, formatter: function (value) { return value == null ? "" : money(value); } }, tension: 0.2 },
        { label: "实际值（USD）", data: chartData.map(function (row) { return row.actual; }), borderColor: COLORS.blue, backgroundColor: pointColors, pointBackgroundColor: pointColors, pointBorderColor: "#fff", pointBorderWidth: 2, pointRadius: 5, borderWidth: 3, spanGaps: true, tension: 0.2 }
      ] }, options: chartOptions(false) };
    } else {
      config = { type: "bar", data: { labels: labels, datasets: [
        { label: "标准值（USD）", data: chartData.map(function (row) { return row.target; }), backgroundColor: COLORS.gray, borderRadius: 4, datalabels: { display: true, color: "#fff", formatter: function (value) { return value == null ? "" : money(value); } } },
        { label: "实际值（USD）", data: chartData.map(function (row) { return row.actual; }), backgroundColor: pointColors, borderRadius: 4 }
      ] }, options: Object.assign(chartOptions(false), { datasets: { bar: { maxBarThickness: 44 } } }) };
    }
    document.getElementById("chart-title-" + topic).textContent = planner + " 实际值与标准值（" + (grain === "month" ? "按月" : "按周") + "）";
  }
  if (window.ChartDataLabels && window.Chart) window.Chart.register(window.ChartDataLabels);
  state.charts[topic] = new window.Chart(canvas.getContext("2d"), config);
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
}
function showToast(message) {
  const toast = document.createElement("div"); toast.className = "toast"; toast.textContent = message;
  document.body.appendChild(toast); window.setTimeout(function () { toast.remove(); }, 3200);
}
function activateTopic(topic) {
  document.querySelectorAll(".topic-tab").forEach(function (tab) { tab.classList.toggle("active", tab.dataset.topic === topic); });
  TOPICS.forEach(function (item) {
    const panel = document.getElementById("topic-" + item), active = item === topic;
    panel.classList.toggle("active", active); panel.hidden = !active;
    if (active) window.setTimeout(function () { if (state.charts[item]) state.charts[item].resize(); }, 0);
  });
  document.body.dataset.activeTopic = topic;
}
async function exportPpt(topic) {
  const Pptx = window.PptxGenJS || window.pptxgen;
  if (!Pptx) { showToast("PowerPoint 导出组件加载失败，请刷新后重试。"); return; }
  const deck = new Pptx(); deck.layout = "LAYOUT_WIDE"; deck.author = "KPI 考核"; deck.subject = topic + " 专题 KPI"; deck.title = "KPI 考核";
  deck.theme = { headFontFace: "Aptos Display", bodyFontFace: "Aptos", lang: "zh-CN" };
  const planner = state.planner[topic] || "全部 Planner", grain = state.grain[topic] || "week";
  const start = state.start[topic] || "", end = state.end[topic] || "";
  const filtered = state.records[topic].filter(function (row) { return (!start || row.date >= start) && (!end || row.date <= end) && (planner === "全部 Planner" || row.planner === planner); });
  const rows = latestUnique(filtered, ["planner", "whs"]), lower = state.lower[topic] !== false;
  const actual = rows.reduce(function (sum, row) { return sum + (row.actual == null ? 0 : row.actual); }, 0);
  const target = rows.reduce(function (sum, row) { return sum + (row.target == null ? 0 : row.target); }, 0);
  const result = status(rows.some(function (r) { return r.actual != null; }) ? actual : null, rows.some(function (r) { return r.target != null; }) ? target : null, lower);
  const slide = deck.addSlide(); slide.background = { color: "F4F7FB" };
  slide.addText("KPI 考核 · " + (topic === "INV" ? "On Hand INV" : "EO") + " 专题", { x: 0.55, y: 0.3, w: 12.1, h: 0.45, fontSize: 23, bold: true, color: "18263B" });
  slide.addText("Planner：" + planner + "    时间：" + start + " 至 " + end + "    粒度：" + (grain === "week" ? "按周" : "按月"), { x: 0.6, y: 0.85, w: 12, h: 0.3, fontSize: 11, color: "77849A" });
  const cards = [["实际值", money(rows.some(function (r) { return r.actual != null; }) ? actual : null)], ["标准值", money(rows.some(function (r) { return r.target != null; }) ? target : null)], ["差额", money(rows.length ? actual - target : null, true)], ["考核结果", result]];
  cards.forEach(function (item, index) {
    const x = 0.6 + index * 3.1; slide.addShape(deck.ShapeType.roundRect, { x: x, y: 1.25, w: 2.85, h: 0.78, rectRadius: 0.08, line: { color: "E1E7EF" }, fill: { color: index === 3 ? (result === "已达标" ? "E9F6EE" : result === "未达标" ? "FFF0F0" : "FFF6E6") : "FFFFFF" } });
    slide.addText(item[0], { x: x + 0.14, y: 1.36, w: 2.6, h: 0.2, fontSize: 10, color: "77849A" });
    slide.addText(item[1], { x: x + 0.14, y: 1.63, w: 2.6, h: 0.25, fontSize: 15, bold: true, color: "263851" });
  });
  const chartCanvas = document.getElementById("chart-" + topic);
  if (chartCanvas && chartCanvas.parentElement.hidden === false) slide.addImage({ data: chartCanvas.toDataURL("image/png", 1), x: 0.65, y: 2.2, w: 12, h: 3.6, transparency: 0 });
  slide.addText("金额单位：USD", { x: 10.5, y: 5.85, w: 2, h: 0.2, fontSize: 8, color: "77849A", align: "right" });
  const chunks = [];
  for (let i = 0; i < rows.length; i += 16) chunks.push(rows.slice(i, i + 16));
  if (!chunks.length) chunks.push([]);
  chunks.forEach(function (chunk, page) {
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
  });
  if (!state.records.INV.length && !state.records.EO.length) throw new Error("未找到名称包含 INV 或 EO 的工作表。");
  TOPICS.forEach(function (topic) {
    const records = state.records[topic]; state.grain[topic] = "week"; state.planner[topic] = "全部 Planner"; state.lower[topic] = true;
    if (records.length) { const range = defaultRange(records, "week"); state.start[topic] = range[0]; state.end[topic] = range[1]; }
    else { document.querySelector('[data-topic="' + topic + '"]').disabled = true; document.getElementById("topic-" + topic).hidden = true; }
  });
  const latest = state.records.INV.concat(state.records.EO).map(function (row) { return row.date; }).sort().slice(-1)[0];
  document.getElementById("source-label").textContent = latest ? "数据更新至 " + latest : "KPI 数据已加载";
  const statusBadge = document.getElementById("data-status"); statusBadge.classList.add("ready"); statusBadge.lastChild.textContent = " 数据已更新";
  document.getElementById("load-error").hidden = true;
  document.querySelectorAll(".topic-tab").forEach(function (tab) { tab.addEventListener("click", function () { activateTopic(tab.dataset.topic); }); });
  TOPICS.forEach(function (topic) { if (state.records[topic].length) renderTopic(topic); });
  activateTopic(state.records.INV.length ? "INV" : "EO");
}
setupExports();
loadWorkbook().catch(function (error) {
  setLoadError(error.message || "读取 Excel 失败，请检查文件格式和工作表。");
  console.error(error);
});

