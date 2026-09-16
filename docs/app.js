const METRICS = ["Calories", "Protein", "Carbs", "Fat"];
const LABELS = { TRAINING: "訓練日", REST: "休息日" };
const state = { data: null, dayType: "TRAINING" };

const numberFormat = new Intl.NumberFormat("zh-TW", { maximumFractionDigits: 1 });

function formatNumber(value) {
  return value === null || value === undefined || Number.isNaN(Number(value))
    ? "—"
    : numberFormat.format(Number(value));
}

function formatDate(value) {
  if (!value) return "尚無日期";
  const [year, month, day] = value.split("-");
  return `${year}.${month}.${day}`;
}

function targetText(range) {
  if (!range) return "目標 ≥ —";
  return range.min === range.max
    ? `目標 ≥ ${formatNumber(range.min)}`
    : `目標 ${formatNumber(range.min)}–${formatNumber(range.max)}`;
}

function computeAdvice(actual, dayType, data) {
  const target = data.targetRanges[dayType];
  if (!target) return "請先選擇訓練日或休息日。";
  if (METRICS.some((metric) => actual?.[metric] === null || actual?.[metric] === undefined)) {
    return "請補齊 Calories、Protein、Carbs、Fat。";
  }
  const diff = Object.fromEntries(METRICS.map((metric) => [metric, target[metric].min - Number(actual[metric])]));
  if (diff.Protein > 0) return "優先補蛋白質。";
  if (diff.Carbs > 0) return "蛋白質已足夠，優先補碳水，例如白飯或香蕉。";
  if (diff.Fat > 0) return "優先補脂肪。";
  if (diff.Calories > 0) return "三大營養素已達最低，熱量尚差；可補少量主食。";
  return "已達每日最低目標；蛋白質足夠，不需補充乳清。";
}

function latestRecord() {
  return state.data?.days?.[0] || null;
}

function setStatus(text, isError = false) {
  const banner = document.getElementById("statusBanner");
  banner.classList.toggle("is-error", isError);
  document.getElementById("statusText").textContent = text;
}

function renderMetrics() {
  const latest = latestRecord();
  const ranges = state.data.targetRanges[state.dayType];
  METRICS.forEach((metric) => {
    const actual = latest?.actual?.[metric] ?? null;
    const target = ranges?.[metric]?.min ?? null;
    const diffNode = document.querySelector(`[data-diff="${metric}"]`);
    document.querySelector(`[data-value="${metric}"]`).textContent = formatNumber(actual);
    document.querySelector(`[data-target="${metric}"]`).textContent = targetText(ranges?.[metric]);
    document.querySelector(`[data-progress="${metric}"]`).style.width = actual !== null && target ? `${Math.min(100, Math.max(0, (actual / target) * 100))}%` : "0%";
    diffNode.classList.remove("is-good", "is-low");
    if (actual === null || target === null) {
      diffNode.textContent = "尚未記錄";
    } else if (target - actual > 0) {
      diffNode.textContent = `還差 ${formatNumber(target - actual)}`;
      diffNode.classList.add("is-low");
    } else {
      diffNode.textContent = `已達最低，多 ${formatNumber(actual - target)}`;
      diffNode.classList.add("is-good");
    }
  });
}

function renderInsight() {
  const latest = latestRecord();
  const typeText = LABELS[state.dayType];
  document.getElementById("latestRange").textContent = `${typeText}目標`;
  if (!latest) {
    document.getElementById("latestDate").textContent = "尚無日期";
    document.getElementById("latestDayType").textContent = "等待紀錄";
    document.getElementById("adviceText").textContent = "尚未有每日紀錄。完成今天的四項輸入後，這裡會顯示簡短建議。";
    return;
  }
  document.getElementById("latestDate").textContent = formatDate(latest.date);
  document.getElementById("latestDayType").textContent = LABELS[latest.dayType] || "未填日型";
  document.getElementById("adviceText").textContent = computeAdvice(latest.actual, state.dayType, state.data);
}

function renderMenuPlan() {
  const section = document.getElementById("menuSection");
  const menuPlan = state.data?.menuPlan;
  if (!section || !menuPlan) {
    if (section) section.hidden = true;
    return;
  }

  const selected = menuPlan[state.dayType === "REST" ? "rest" : "training"];
  if (!selected) {
    section.hidden = true;
    return;
  }

  section.hidden = false;
  document.getElementById("menuTitle").textContent = selected.title || `${LABELS[state.dayType]}菜單`;
  document.getElementById("menuTarget").textContent = `${LABELS[state.dayType]} ${targetText(state.data.targetRanges[state.dayType]?.Calories)}`;

  const menuBody = document.getElementById("menuBody");
  menuBody.replaceChildren();
  selected.rows.forEach((item) => {
    const row = document.createElement("tr");
    [item.time, item.meal, item.chicken, item.rice, item.riceCup, item.other, item.greens, item.note].forEach((value, index) => {
      row.append(cell(value || "—", index === 7 ? "menu-note" : ""));
    });
    menuBody.append(row);
  });

  const total = selected.total || {};
  document.getElementById("menuTotal").textContent = [
    "每日合計",
    `雞胸 ${total.chicken || "—"}`,
    `白飯 ${total.rice || "—"}`,
    `生米 ${total.riceCup || "—"}`,
    `其他 ${total.other || "—"}`,
    `青菜 ${total.greens || "—"}`,
  ].join("　");

  const weeklyBody = document.getElementById("weeklyBody");
  weeklyBody.replaceChildren();
  (menuPlan.weekly || []).forEach((item) => {
    const row = document.createElement("tr");
    row.append(cell(item.day || "—"));
    const typeCell = document.createElement("td");
    const type = document.createElement("span");
    type.className = `day-type${item.type === "休息日" ? " rest" : ""}`;
    type.textContent = item.type || "—";
    typeCell.append(type);
    row.append(typeCell);
    [item.chicken, item.rice, item.mackerel, item.eggs].forEach((value) => row.append(cell(value || "—")));
    weeklyBody.append(row);
  });

  const shoppingList = document.getElementById("shoppingList");
  shoppingList.replaceChildren();
  (menuPlan.shopping || []).forEach((item) => {
    const listItem = document.createElement("li");
    const name = document.createElement("span");
    name.textContent = item.item || "—";
    const quantity = document.createElement("strong");
    quantity.textContent = item.quantity || "—";
    listItem.append(name, quantity);
    shoppingList.append(listItem);
  });

  const replacementBody = document.getElementById("replacementBody");
  replacementBody.replaceChildren();
  (menuPlan.replacements || []).forEach((item) => {
    const row = document.createElement("tr");
    [item.item, item.replaces, item.portion, item.howToUse, item.protein, item.carbs, item.fat, item.note].forEach((value, index) => {
      row.append(cell(value || "—", index === 3 || index === 7 ? "menu-note" : ""));
    });
    replacementBody.append(row);
  });
}

function cell(text, className = "") {
  const node = document.createElement("td");
  node.textContent = text;
  if (className) node.className = className;
  return node;
}

function renderTable() {
  const records = state.data.days || [];
  const body = document.getElementById("recordsBody");
  body.replaceChildren();
  document.getElementById("recordCount").textContent = `${records.length} 天`;
  document.getElementById("tableEmpty").hidden = records.length > 0;
  records.slice(0, 10).forEach((record) => {
    const row = document.createElement("tr");
    const type = document.createElement("span");
    type.className = `day-type${record.dayType === "REST" ? " rest" : ""}`;
    type.textContent = LABELS[record.dayType] || "未填";
    row.append(cell(formatDate(record.date)));
    const typeCell = document.createElement("td");
    typeCell.append(type);
    row.append(typeCell);
    row.append(cell(formatNumber(record.actual.Calories)));
    row.append(cell(formatNumber(record.actual.Protein)));
    row.append(cell(formatNumber(record.actual.Carbs)));
    row.append(cell(formatNumber(record.actual.Fat)));
    row.append(cell(record.advice || "—", "table-advice"));
    body.append(row);
  });
}

function drawTrend() {
  const canvas = document.getElementById("trendChart");
  const empty = document.getElementById("chartEmpty");
  const records = (state.data.days || []).slice(0, 7).reverse();
  empty.hidden = records.length > 0;
  canvas.hidden = records.length === 0;
  if (!records.length) return;

  const rect = canvas.getBoundingClientRect();
  const ratio = window.devicePixelRatio || 1;
  canvas.width = Math.max(1, Math.floor(rect.width * ratio));
  canvas.height = Math.max(1, Math.floor(rect.height * ratio));
  const ctx = canvas.getContext("2d");
  ctx.scale(ratio, ratio);
  const width = rect.width;
  const height = rect.height;
  const pad = { top: 18, right: 17, bottom: 31, left: 45 };
  const values = records.flatMap((record) => [record.actual.Calories, record.target.Calories]).filter((value) => value !== null && value !== undefined);
  const minValue = Math.floor((Math.min(...values) - 100) / 100) * 100;
  const maxValue = Math.ceil((Math.max(...values) + 100) / 100) * 100;
  const x = (index) => pad.left + (index * (width - pad.left - pad.right)) / Math.max(1, records.length - 1);
  const y = (value) => pad.top + ((maxValue - value) * (height - pad.top - pad.bottom)) / Math.max(1, maxValue - minValue);

  ctx.clearRect(0, 0, width, height);
  ctx.font = "11px Inter, system-ui, sans-serif";
  ctx.lineWidth = 1;
  ctx.strokeStyle = "#e8efed";
  ctx.fillStyle = "#8b9b9a";
  for (let index = 0; index <= 3; index += 1) {
    const value = minValue + ((maxValue - minValue) * index) / 3;
    const yPosition = y(value);
    ctx.beginPath();
    ctx.moveTo(pad.left, yPosition);
    ctx.lineTo(width - pad.right, yPosition);
    ctx.stroke();
    ctx.fillText(formatNumber(Math.round(value)), 0, yPosition + 4);
  }

  ctx.setLineDash([5, 5]);
  ctx.strokeStyle = "#a9b7b8";
  ctx.beginPath();
  records.forEach((record, index) => {
    const pointX = x(index);
    const pointY = y(record.target.Calories);
    if (index === 0) ctx.moveTo(pointX, pointY); else ctx.lineTo(pointX, pointY);
  });
  ctx.stroke();
  ctx.setLineDash([]);

  ctx.strokeStyle = "#299b8a";
  ctx.lineWidth = 3;
  ctx.beginPath();
  records.forEach((record, index) => {
    const pointX = x(index);
    const pointY = y(record.actual.Calories);
    if (index === 0) ctx.moveTo(pointX, pointY); else ctx.lineTo(pointX, pointY);
  });
  ctx.stroke();
  records.forEach((record, index) => {
    ctx.beginPath();
    ctx.fillStyle = "#ffffff";
    ctx.arc(x(index), y(record.actual.Calories), 4.5, 0, Math.PI * 2);
    ctx.fill();
    ctx.lineWidth = 2.5;
    ctx.strokeStyle = "#299b8a";
    ctx.stroke();
    ctx.fillStyle = "#8b9b9a";
    ctx.font = "11px Inter, system-ui, sans-serif";
    ctx.textAlign = "center";
    ctx.fillText(record.date.slice(5).replace("-", "/"), x(index), height - 8);
  });
  ctx.textAlign = "left";
}

function render() {
  if (!state.data) return;
  const latest = latestRecord();
  document.querySelectorAll("[data-day-type]").forEach((button) => button.classList.toggle("is-active", button.dataset.dayType === state.dayType));
  document.getElementById("updatedAt").textContent = state.data.generatedAt ? formatDate(state.data.generatedAt.slice(0, 10)) : "—";
  setStatus(latest ? `已載入 ${state.data.days.length} 天，最新紀錄 ${formatDate(latest.date)}。` : "尚未有每日紀錄，先輸入今天的四項數值即可開始。");
  renderMetrics();
  renderInsight();
  renderMenuPlan();
  renderTable();
  drawTrend();
}

async function loadData() {
  try {
    const response = await fetch("./data.json", { cache: "no-store" });
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    state.data = await response.json();
    state.dayType = state.data.days?.[0]?.dayType || "TRAINING";
    render();
  } catch (error) {
    setStatus("讀取資料失敗，請確認 GitHub Pages 已包含 data.json。", true);
    document.getElementById("updatedAt").textContent = "讀取失敗";
    console.error(error);
  }
}

document.querySelectorAll("[data-day-type]").forEach((button) => {
  button.addEventListener("click", () => {
    if (!state.data) return;
    state.dayType = button.dataset.dayType;
    render();
  });
});
window.addEventListener("resize", () => { if (state.data) drawTrend(); });
loadData();
