const state = {
  workbook: null,
  sheetName: "",
  rows: [],
  columns: [],
  results: null,
};

const els = {
  fileInput: document.getElementById("fileInput"),
  dropZone: document.getElementById("dropZone"),
  fileStatus: document.getElementById("fileStatus"),
  sheetSelect: document.getElementById("sheetSelect"),
  sampleColumnSelect: document.getElementById("sampleColumnSelect"),
  objectiveMethod: document.getElementById("objectiveMethod"),
  combineMethod: document.getElementById("combineMethod"),
  alphaInput: document.getElementById("alphaInput"),
  methodHint: document.getElementById("methodHint"),
  indicatorBody: document.getElementById("indicatorBody"),
  addIndicatorBtn: document.getElementById("addIndicatorBtn"),
  calculateBtn: document.getElementById("calculateBtn"),
  exportExcelBtn: document.getElementById("exportExcelBtn"),
  exportImageBtn: document.getElementById("exportImageBtn"),
  messageBox: document.getElementById("messageBox"),
  weightTotal: document.getElementById("weightTotal"),
  rowCount: document.getElementById("rowCount"),
  weightsPreview: document.getElementById("weightsPreview"),
  resultPreview: document.getElementById("resultPreview"),
  weightChart: document.getElementById("weightChart"),
};

const defaultIndicators = [
  { name: "抗拉强度", column: "抗拉强度/MPa", direction: "max", weight: 60 },
  { name: "硬度", column: "硬度/HV", direction: "max", weight: 20 },
  { name: "延伸率", column: "延伸率/%", direction: "max", weight: 20 },
];

function init() {
  renderIndicators(defaultIndicators);
  bindEvents();
  refreshWeightTotal();
  drawEmptyChart();

  if (window.lucide) window.lucide.createIcons();
  if (!window.XLSX) {
    showMessage("Excel 解析库未加载，请确认网络可访问 CDN 后刷新页面。", "error");
  }
}

function bindEvents() {
  els.fileInput.addEventListener("change", (event) => {
    const [file] = event.target.files;
    if (file) handleFile(file);
  });

  ["dragenter", "dragover"].forEach((eventName) => {
    els.dropZone.addEventListener(eventName, (event) => {
      event.preventDefault();
      els.dropZone.classList.add("dragging");
    });
  });

  ["dragleave", "drop"].forEach((eventName) => {
    els.dropZone.addEventListener(eventName, (event) => {
      event.preventDefault();
      els.dropZone.classList.remove("dragging");
    });
  });

  els.dropZone.addEventListener("drop", (event) => {
    const [file] = event.dataTransfer.files;
    if (file) handleFile(file);
  });

  els.sheetSelect.addEventListener("change", () => loadSheet(els.sheetSelect.value));
  els.addIndicatorBtn.addEventListener("click", () => addIndicatorRow({ name: "", column: "", direction: "max", weight: 0 }));
  els.calculateBtn.addEventListener("click", calculate);
  els.exportExcelBtn.addEventListener("click", exportWorkbook);
  els.exportImageBtn.addEventListener("click", exportChartImage);
  els.alphaInput.addEventListener("input", invalidateResults);
  els.objectiveMethod.addEventListener("change", () => {
    updateMethodHint();
    invalidateResults();
  });
  els.combineMethod.addEventListener("change", () => {
    updateMethodHint();
    invalidateResults();
  });
  updateMethodHint();
}

async function handleFile(file) {
  if (!window.XLSX) {
    showMessage("Excel 解析库未加载，暂时无法读取文件。", "error");
    return;
  }

  const buffer = await file.arrayBuffer();
  state.workbook = XLSX.read(buffer, { type: "array" });
  state.sheetName = state.workbook.SheetNames[0] || "";

  els.sheetSelect.innerHTML = state.workbook.SheetNames.map((name) => `<option value="${escapeHtml(name)}">${escapeHtml(name)}</option>`).join("");
  els.sheetSelect.disabled = false;
  els.sheetSelect.value = state.sheetName;
  els.fileStatus.textContent = `已导入：${file.name}`;
  loadSheet(state.sheetName);
}

function loadSheet(sheetName) {
  state.sheetName = sheetName;
  const sheet = state.workbook.Sheets[sheetName];
  state.rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
  state.columns = state.rows.length > 0 ? Object.keys(state.rows[0]) : [];
  state.results = null;

  els.rowCount.textContent = String(state.rows.length);
  fillSelect(els.sampleColumnSelect, state.columns);
  els.sampleColumnSelect.disabled = state.columns.length === 0;
  chooseDefaultSampleColumn();
  syncIndicatorColumnOptions();

  els.calculateBtn.disabled = state.rows.length === 0 || state.columns.length === 0;
  invalidateResults();
  showMessage(state.rows.length ? "Excel 已读取，可以调整指标后计算。" : "当前工作表没有数据。", state.rows.length ? "success" : "warning");
}

function chooseDefaultSampleColumn() {
  const candidates = ["试样编号", "样品编号", "编号", "Sample", "sample"];
  const found = candidates.find((name) => state.columns.includes(name));
  if (found) els.sampleColumnSelect.value = found;
}

function fillSelect(select, options, selectedValue = "") {
  select.innerHTML = options.map((option) => `<option value="${escapeHtml(option)}">${escapeHtml(option)}</option>`).join("");
  if (selectedValue && options.includes(selectedValue)) select.value = selectedValue;
}

function renderIndicators(indicators) {
  els.indicatorBody.innerHTML = "";
  indicators.forEach((indicator) => addIndicatorRow(indicator));
}

function addIndicatorRow(indicator) {
  const row = document.createElement("tr");
  row.innerHTML = `
    <td><input class="name-input" type="text" value="${escapeAttr(indicator.name)}" placeholder="指标名称" /></td>
    <td><select class="column-select"></select></td>
    <td>
      <select class="direction-select">
        <option value="max">越大越好</option>
        <option value="min">越小越好</option>
      </select>
    </td>
    <td><input class="weight-input" type="number" min="0" step="0.01" value="${Number(indicator.weight) || 0}" /></td>
    <td>
      <button class="icon-only delete-row" type="button" title="删除指标" aria-label="删除指标">
        <i data-lucide="trash-2"></i>
      </button>
    </td>
  `;

  row.querySelector(".direction-select").value = indicator.direction || "max";
  els.indicatorBody.appendChild(row);

  const columnSelect = row.querySelector(".column-select");
  fillSelect(columnSelect, state.columns, guessColumn(indicator.column || indicator.name));
  if (!state.columns.length && indicator.column) {
    columnSelect.innerHTML = `<option value="${escapeHtml(indicator.column)}">${escapeHtml(indicator.column)}</option>`;
  }

  row.querySelector(".delete-row").addEventListener("click", () => {
    row.remove();
    refreshWeightTotal();
    invalidateResults();
  });

  row.querySelectorAll("input, select").forEach((control) => {
    control.addEventListener("input", () => {
      refreshWeightTotal();
      invalidateResults();
    });
    control.addEventListener("change", () => {
      refreshWeightTotal();
      invalidateResults();
    });
  });

  refreshWeightTotal();
  if (window.lucide) window.lucide.createIcons();
}

function syncIndicatorColumnOptions() {
  const rows = [...els.indicatorBody.querySelectorAll("tr")];
  rows.forEach((row) => {
    const name = row.querySelector(".name-input").value;
    const select = row.querySelector(".column-select");
    const current = select.value || guessColumn(name);
    fillSelect(select, state.columns, current);
  });
}

function guessColumn(value) {
  if (!value) return "";
  if (state.columns.includes(value)) return value;
  const normalized = normalizeText(value);
  return state.columns.find((column) => normalizeText(column).includes(normalized) || normalized.includes(normalizeText(column))) || "";
}

function invalidateResults() {
  state.results = null;
  els.exportExcelBtn.disabled = true;
  els.exportImageBtn.disabled = true;
  els.weightsPreview.className = "preview-empty";
  els.weightsPreview.textContent = "暂无计算结果";
  els.resultPreview.className = "preview-empty";
  els.resultPreview.textContent = "暂无计算结果";
  drawEmptyChart();
}

function refreshWeightTotal() {
  const total = collectIndicators({ allowEmptyColumn: true }).reduce((sum, item) => sum + item.weight, 0);
  els.weightTotal.textContent = `${formatNumber(total)}%`;
  els.weightTotal.style.color = Math.abs(total - 100) < 0.01 ? "var(--primary-dark)" : "var(--warning)";
}

function collectIndicators({ allowEmptyColumn = false } = {}) {
  return [...els.indicatorBody.querySelectorAll("tr")]
    .map((row) => ({
      name: row.querySelector(".name-input").value.trim(),
      column: row.querySelector(".column-select").value,
      direction: row.querySelector(".direction-select").value,
      weight: Number(row.querySelector(".weight-input").value) || 0,
    }))
    .filter((item) => item.name || item.column || allowEmptyColumn);
}

function calculate() {
  try {
    const indicators = collectIndicators();
    validateInputs(indicators);

    const scoreRows = normalizeIndicatorMatrix(state.rows, indicators);
    const subjectiveWeights = normalizeWeights(Object.fromEntries(indicators.map((item) => [item.name, item.weight])));
    const objective = els.objectiveMethod.value === "entropy" ? calculateEntropyWeights(scoreRows) : calculateCriticWeights(scoreRows);
    const alpha = clamp(Number(els.alphaInput.value) / 100, 0, 1);
    const finalWeights = combineWeights(subjectiveWeights, objective.weights, els.combineMethod.value, alpha);
    const topsis = weightedTopsis(scoreRows, finalWeights);

    const sampleCol = els.sampleColumnSelect.value;
    const resultRows = state.rows.map((row, index) => {
      const output = { [sampleCol]: row[sampleCol] };
      indicators.forEach((indicator) => {
        output[indicator.column] = row[indicator.column];
        output[`标准化_${indicator.name}`] = scoreRows[index][indicator.name];
      });
      return {
        ...output,
        "D_plus_到正理想解距离": topsis.rows[index].dPlus,
        "D_minus_到负理想解距离": topsis.rows[index].dMinus,
        "TOPSIS贴近度": topsis.rows[index].closeness,
        "TOPSIS评分_100分": topsis.rows[index].score,
      };
    }).sort((a, b) => b["TOPSIS贴近度"] - a["TOPSIS贴近度"]);

    resultRows.forEach((row, index) => {
      row["排名"] = index + 1;
    });

    const weightChartRows = indicators.map((indicator) => ({
      name: indicator.name,
      subjective: subjectiveWeights[indicator.name],
      objective: objective.weights[indicator.name],
      final: finalWeights[indicator.name],
    }));

    const weightsRows = indicators.map((indicator) => ({
      "指标名称": indicator.name,
      "Excel列名": indicator.column,
      "方向": indicator.direction === "max" ? "越大越好" : "越小越好",
      "主观权重": subjectiveWeights[indicator.name],
      "客观方法": els.objectiveMethod.value === "entropy" ? "熵权法" : "CRITIC法",
      "客观权重": objective.weights[indicator.name],
      "组合方法": getCombineMethodName(els.combineMethod.value),
      "组合权重": finalWeights[indicator.name],
      "组合权重_%": finalWeights[indicator.name] * 100,
    }));

    state.results = {
      resultRows,
      weightsRows,
      normalizedRows: scoreRows,
      objectiveDetailRows: objective.detailRows,
      corrRows: objective.corrRows || [],
      weightChartRows,
    };

    renderPreview(weightsRows, resultRows);
    drawWeightChart(weightChartRows);
    els.exportExcelBtn.disabled = false;
    els.exportImageBtn.disabled = false;
    showMessage("计算完成，可以分别导出 Excel 和图片。", "success");
  } catch (error) {
    state.results = null;
    els.exportExcelBtn.disabled = true;
    els.exportImageBtn.disabled = true;
    drawEmptyChart();
    showMessage(error.message, "error");
  }
}

function validateInputs(indicators) {
  if (!state.rows.length) throw new Error("请先导入包含数据的 Excel。");
  if (!els.sampleColumnSelect.value) throw new Error("请选择样品编号列。");
  if (indicators.length === 0) throw new Error("请至少添加一个指标。");

  const names = new Set();
  indicators.forEach((indicator, index) => {
    if (!indicator.name) throw new Error(`第 ${index + 1} 行缺少指标名称。`);
    if (names.has(indicator.name)) throw new Error(`指标名称重复：${indicator.name}`);
    names.add(indicator.name);
    if (!indicator.column) throw new Error(`指标 ${indicator.name} 未选择 Excel 列名。`);
    if (!state.columns.includes(indicator.column)) throw new Error(`Excel 中找不到列：${indicator.column}`);
    if (indicator.weight <= 0) throw new Error(`指标 ${indicator.name} 的权重必须大于 0。`);
  });
}

function normalizeIndicatorMatrix(rows, indicators) {
  return rows.map((_, rowIndex) => {
    const output = {};
    indicators.forEach((indicator) => {
      const values = rows.map((row) => toNumber(row[indicator.column], indicator.column));
      const min = Math.min(...values);
      const max = Math.max(...values);
      let value = 1;

      if (!nearlyEqual(max, min)) {
        const raw = values[rowIndex];
        value = indicator.direction === "max" ? (raw - min) / (max - min) : (max - raw) / (max - min);
      }

      output[indicator.name] = value;
    });
    return output;
  });
}

function calculateCriticWeights(scoreRows) {
  const names = Object.keys(scoreRows[0]);
  const std = Object.fromEntries(names.map((name) => [name, standardDeviation(scoreRows.map((row) => row[name]))]));
  const corr = {};

  names.forEach((a) => {
    corr[a] = {};
    names.forEach((b) => {
      corr[a][b] = correlation(scoreRows.map((row) => row[a]), scoreRows.map((row) => row[b]));
    });
  });

  const info = {};
  names.forEach((name) => {
    const conflict = names.filter((other) => other !== name).reduce((sum, other) => sum + (1 - corr[name][other]), 0);
    info[name] = std[name] * conflict;
  });

  const infoTotal = Object.values(info).reduce((sum, value) => sum + value, 0);
  const weights = {};
  names.forEach((name) => {
    weights[name] = nearlyEqual(infoTotal, 0) ? 1 / names.length : info[name] / infoTotal;
  });

  const detailRows = names.map((name) => ({
    "指标名称": name,
    "标准差": std[name],
    "冲突性": names.filter((other) => other !== name).reduce((sum, other) => sum + (1 - corr[name][other]), 0),
    "信息量Cj": info[name],
    "客观权重": weights[name],
  }));

  const corrRows = names.map((name) => {
    const row = { "指标名称": name };
    names.forEach((other) => {
      row[other] = corr[name][other];
    });
    return row;
  });

  return { weights, detailRows, corrRows };
}

function calculateEntropyWeights(scoreRows) {
  const names = Object.keys(scoreRows[0]);
  const sampleCount = scoreRows.length;
  const entropyFactor = sampleCount > 1 ? 1 / Math.log(sampleCount) : 0;
  const epsilon = 1e-12;
  const entropy = {};
  const divergence = {};

  names.forEach((name) => {
    const column = scoreRows.map((row) => row[name] + epsilon);
    const total = column.reduce((sum, value) => sum + value, 0);
    const proportions = column.map((value) => value / total);
    entropy[name] = -entropyFactor * proportions.reduce((sum, value) => sum + value * Math.log(value), 0);
    divergence[name] = 1 - entropy[name];
  });

  const divergenceTotal = Object.values(divergence).reduce((sum, value) => sum + value, 0);
  const weights = {};
  names.forEach((name) => {
    weights[name] = nearlyEqual(divergenceTotal, 0) ? 1 / names.length : divergence[name] / divergenceTotal;
  });

  const detailRows = names.map((name) => ({
    "指标名称": name,
    "信息熵Ej": entropy[name],
    "差异系数Dj": divergence[name],
    "客观权重": weights[name],
  }));

  return { weights, detailRows, corrRows: [] };
}

function combineWeights(subjective, objective, method, alpha) {
  const names = Object.keys(subjective);
  const combined = {};

  names.forEach((name) => {
    if (method === "additive") {
      combined[name] = alpha * subjective[name] + (1 - alpha) * objective[name];
    } else if (method === "product") {
      combined[name] = subjective[name] * objective[name];
    } else {
      combined[name] = Math.pow(subjective[name] + 1e-12, alpha) * Math.pow(objective[name] + 1e-12, 1 - alpha);
    }
  });

  return normalizeWeights(combined);
}

function weightedTopsis(scoreRows, weights) {
  const names = Object.keys(weights);
  const rows = scoreRows.map((row) => {
    const dPlus = Math.sqrt(names.reduce((sum, name) => sum + weights[name] * Math.pow(row[name] - 1, 2), 0));
    const dMinus = Math.sqrt(names.reduce((sum, name) => sum + weights[name] * Math.pow(row[name], 2), 0));
    const closeness = dMinus / (dPlus + dMinus);
    return { dPlus, dMinus, closeness, score: closeness * 100 };
  });

  return { rows };
}

function updateMethodHint() {
  const objectiveName = els.objectiveMethod.value === "entropy" ? "熵权法" : "CRITIC";
  els.methodHint.value = `${objectiveName} + ${getCombineMethodName(els.combineMethod.value)}`;
}

function getCombineMethodName(value) {
  if (value === "additive") return "加权加法集成法";
  if (value === "product") return "乘法集成法";
  return "指数加权乘法集成法";
}

function renderPreview(weightsRows, resultRows) {
  els.weightsPreview.className = "";
  els.resultPreview.className = "";
  els.weightsPreview.innerHTML = renderTable(weightsRows, ["指标名称", "主观权重", "客观权重", "组合权重_%"], 10);
  els.resultPreview.innerHTML = renderTable(resultRows, Object.keys(resultRows[0]).slice(0, 10), 8);
}

function renderTable(rows, columns, limit) {
  const visibleRows = rows.slice(0, limit);
  return `
    <div class="preview-table-wrap">
      <table>
        <thead>
          <tr>${columns.map((column) => `<th>${escapeHtml(column)}</th>`).join("")}</tr>
        </thead>
        <tbody>
          ${visibleRows
            .map(
              (row) => `
            <tr>
              ${columns.map((column) => `<td>${escapeHtml(formatCell(row[column]))}</td>`).join("")}
            </tr>
          `,
            )
            .join("")}
        </tbody>
      </table>
    </div>
  `;
}

function drawEmptyChart() {
  const ctx = els.weightChart.getContext("2d");
  const { width, height } = els.weightChart;
  ctx.clearRect(0, 0, width, height);
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, width, height);
  ctx.strokeStyle = "#dbe2de";
  ctx.lineWidth = 2;
  ctx.strokeRect(40, 40, width - 80, height - 80);
  ctx.fillStyle = "#68746f";
  ctx.font = '32px "Microsoft YaHei", sans-serif';
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  ctx.fillText("计算后生成权重变化图", width / 2, height / 2);
}

function drawWeightChart(rows) {
  const canvas = els.weightChart;
  const ctx = canvas.getContext("2d");
  const width = canvas.width;
  const height = canvas.height;
  const groups = [
    { key: "subjective", label: "主观权重", color: "#c8f08a" },
    { key: "objective", label: "客观权重", color: "#f2e88b" },
    { key: "final", label: "组合权重", color: "#a7d7ee" },
  ];
  const maxValue = Math.max(0.65, ...rows.flatMap((row) => groups.map((group) => row[group.key])));
  const chart = { left: 92, top: 72, right: width - 72, bottom: height - 132 };
  const chartWidth = chart.right - chart.left;
  const chartHeight = chart.bottom - chart.top;
  const groupGap = 76;
  const innerGap = 18;
  const groupWidth = (chartWidth - groupGap * (groups.length - 1)) / groups.length;
  const barWidth = Math.min(74, (groupWidth - innerGap * (rows.length - 1)) / rows.length);

  ctx.clearRect(0, 0, width, height);
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, width, height);

  ctx.strokeStyle = "#1f1f1f";
  ctx.lineWidth = 3;
  ctx.strokeRect(chart.left, chart.top, chartWidth, chartHeight);

  ctx.strokeStyle = "#e6e6e6";
  ctx.lineWidth = 1;
  ctx.fillStyle = "#1f1f1f";
  ctx.font = '26px "Times New Roman", "Microsoft YaHei", serif';
  ctx.textAlign = "right";
  ctx.textBaseline = "middle";
  for (let tick = 0; tick <= maxValue + 1e-9; tick += 0.1) {
    const y = chart.bottom - (tick / maxValue) * chartHeight;
    if (Math.round(tick * 10) % 2 === 0) {
      ctx.fillText(tick.toFixed(1), chart.left - 24, y);
      ctx.beginPath();
      ctx.moveTo(chart.left, y);
      ctx.lineTo(chart.right, y);
      ctx.stroke();
    }
  }

  groups.forEach((group, groupIndex) => {
    const groupStart = chart.left + groupIndex * (groupWidth + groupGap);
    const totalBarsWidth = rows.length * barWidth + (rows.length - 1) * innerGap;
    const firstX = groupStart + (groupWidth - totalBarsWidth) / 2;

    if (groupIndex > 0) {
      const lineX = groupStart - groupGap / 2;
      ctx.save();
      ctx.setLineDash([20, 18]);
      ctx.strokeStyle = "#111111";
      ctx.lineWidth = 3;
      ctx.beginPath();
      ctx.moveTo(lineX, chart.top + 12);
      ctx.lineTo(lineX, chart.bottom);
      ctx.stroke();
      ctx.restore();
    }

    rows.forEach((row, index) => {
      const value = row[group.key];
      const x = firstX + index * (barWidth + innerGap);
      const barHeight = (value / maxValue) * chartHeight;
      const y = chart.bottom - barHeight;

      ctx.fillStyle = group.color;
      ctx.strokeStyle = "#111111";
      ctx.lineWidth = 2;
      ctx.fillRect(x, y, barWidth, barHeight);
      ctx.strokeRect(x, y, barWidth, barHeight);

      ctx.fillStyle = "#111111";
      ctx.font = '25px "Times New Roman", "Microsoft YaHei", serif';
      ctx.textAlign = "center";
      ctx.textBaseline = "bottom";
      ctx.fillText(trimNumber(value), x + barWidth / 2, y - 6);

      ctx.font = '23px "Microsoft YaHei", sans-serif';
      ctx.textBaseline = "top";
      ctx.fillText(row.name, x + barWidth / 2, chart.bottom + 18);
    });

    ctx.fillStyle = "#111111";
    ctx.font = '34px "Microsoft YaHei", sans-serif';
    ctx.textAlign = "center";
    ctx.textBaseline = "top";
    ctx.fillText(group.label, groupStart + groupWidth / 2, chart.bottom + 62);
  });

  ctx.strokeStyle = "#111111";
  ctx.lineWidth = 3;
  ctx.beginPath();
  ctx.moveTo(chart.left, chart.bottom + 54);
  ctx.lineTo(chart.right, chart.bottom + 54);
  ctx.stroke();
}

async function exportWorkbook() {
  if (!state.results || !window.XLSX) return;

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(state.results.resultRows), "topsis_result");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(state.results.weightsRows), "weights");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(state.results.normalizedRows), "normalized");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(state.results.objectiveDetailRows), "objective_detail");
  if (state.results.corrRows.length) {
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(state.results.corrRows), "correlation_matrix");
  }

  const arrayBuffer = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
  await saveBlob(new Blob([arrayBuffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), "topsis_output.xlsx", [
    {
      description: "Excel 文件",
      accept: { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"] },
    },
  ]);
}

async function exportChartImage() {
  if (!state.results) return;
  const blob = await new Promise((resolve) => els.weightChart.toBlob(resolve, "image/png", 1));
  await saveBlob(blob, "weight_chart.png", [
    {
      description: "PNG 图片",
      accept: { "image/png": [".png"] },
    },
  ]);
}

async function saveBlob(blob, suggestedName, types) {
  if (window.showSaveFilePicker) {
    try {
      const handle = await window.showSaveFilePicker({ suggestedName, types });
      const writable = await handle.createWritable();
      await writable.write(blob);
      await writable.close();
      return;
    } catch (error) {
      if (error.name === "AbortError") return;
    }
  }

  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = suggestedName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function normalizeWeights(values) {
  const entries = Object.entries(values);
  const total = entries.reduce((sum, [, value]) => sum + Number(value), 0);
  if (total <= 0) throw new Error("权重之和必须大于 0。");
  return Object.fromEntries(entries.map(([key, value]) => [key, Number(value) / total]));
}

function standardDeviation(values) {
  const mean = values.reduce((sum, value) => sum + value, 0) / values.length;
  const variance = values.reduce((sum, value) => sum + Math.pow(value - mean, 2), 0) / values.length;
  return Math.sqrt(variance);
}

function correlation(a, b) {
  const meanA = a.reduce((sum, value) => sum + value, 0) / a.length;
  const meanB = b.reduce((sum, value) => sum + value, 0) / b.length;
  const numerator = a.reduce((sum, value, index) => sum + (value - meanA) * (b[index] - meanB), 0);
  const denomA = Math.sqrt(a.reduce((sum, value) => sum + Math.pow(value - meanA, 2), 0));
  const denomB = Math.sqrt(b.reduce((sum, value) => sum + Math.pow(value - meanB, 2), 0));
  const denominator = denomA * denomB;
  return nearlyEqual(denominator, 0) ? 0 : numerator / denominator;
}

function toNumber(value, column) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const numeric = Number(String(value).replace(/,/g, "").trim());
  if (!Number.isFinite(numeric)) throw new Error(`列 ${column} 中存在空值或非数值：${value}`);
  return numeric;
}

function clamp(value, min, max) {
  return Math.min(max, Math.max(min, value));
}

function nearlyEqual(a, b) {
  return Math.abs(a - b) < 1e-12;
}

function formatNumber(value) {
  return Number(value).toLocaleString("zh-CN", { maximumFractionDigits: 4 });
}

function trimNumber(value) {
  return Number(value).toFixed(3).replace(/0+$/, "").replace(/\.$/, "");
}

function formatCell(value) {
  return typeof value === "number" ? formatNumber(value) : value;
}

function normalizeText(value) {
  return String(value).replace(/\s+/g, "").toLowerCase();
}

function showMessage(message, type = "") {
  els.messageBox.textContent = message;
  els.messageBox.className = `message ${type}`.trim();
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function escapeAttr(value) {
  return escapeHtml(value);
}

init();
