import { Streamlit } from "streamlit-component-lib";
import "./style.css";

const root = document.getElementById("app");

let activeBrush = "F";
let isPainting = false;
let rows = [];
let days = [];
let initialRows = [];
let undoStack = [];
let currentVersion = null;

const brushItems = [
  { code: "F", label: "F", title: "Frühdienst" },
  { code: "S", label: "S", title: "Spätdienst" },
  { code: "N", label: "N", title: "Nachtdienst" },
  { code: "Frei", label: "Frei", title: "Arbeitsfrei" },
  { code: "", label: "⌫", title: "Radierer" },
];

function copyRows(value) {
  return JSON.parse(JSON.stringify(value || []));
}

function normalizeCode(value) {
  const code = String(value || "").trim();
  return code.toUpperCase() === "FREI" ? "Frei" : code.toUpperCase();
}

function codeClass(code) {
  const value = normalizeCode(code);
  if (value === "F") return "code-f";
  if (value === "S") return "code-s";
  if (value === "N") return "code-n";
  if (value === "U") return "code-u";
  if (value === "W") return "code-w";
  if (value === "FB") return "code-fb";
  if (value === "K") return "code-k";
  if (value === "Frei") return "code-frei";
  return "code-empty";
}

function saveUndo() {
  undoStack.push(copyRows(rows));
  if (undoStack.length > 50) undoStack.shift();
}

function paintCell(rowIndex, dayKey, cell) {
  const row = rows[rowIndex];
  if (!row || row.locked?.[dayKey]) return;
  const next = normalizeCode(activeBrush);
  if (normalizeCode(row.cells?.[dayKey]) === next) return;
  row.cells[dayKey] = next;
  cell.textContent = next;
  cell.className = `np-cell ${codeClass(next)}`;
  cell.dataset.locked = "false";
  updateDirtyState();
}

function updateDirtyState() {
  const dirty = JSON.stringify(rows) !== JSON.stringify(initialRows);
  const saveButton = document.getElementById("save-plan");
  const resetButton = document.getElementById("reset-plan");
  if (saveButton) saveButton.disabled = !dirty;
  if (resetButton) resetButton.disabled = !dirty;
}

function renderToolbar() {
  const toolbar = document.createElement("div");
  toolbar.className = "np-toolbar";

  const hint = document.createElement("div");
  hint.className = "np-toolbar-title";
  hint.textContent = "Dienst auswählen";
  toolbar.appendChild(hint);

  brushItems.forEach((item) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = `np-brush ${codeClass(item.code)}`;
    if (item.code === activeBrush) button.classList.add("active");
    button.textContent = item.label;
    button.title = item.title;
    button.addEventListener("click", () => {
      activeBrush = item.code;
      document.querySelectorAll(".np-brush").forEach((element) => {
        element.classList.remove("active");
      });
      button.classList.add("active");
    });
    toolbar.appendChild(button);
  });

  const spacer = document.createElement("div");
  spacer.className = "np-spacer";
  toolbar.appendChild(spacer);

  const undoButton = document.createElement("button");
  undoButton.type = "button";
  undoButton.className = "np-action secondary";
  undoButton.textContent = "↶ Rückgängig";
  undoButton.addEventListener("click", () => {
    const previous = undoStack.pop();
    if (previous) {
      rows = previous;
      render();
    }
  });
  toolbar.appendChild(undoButton);

  const resetButton = document.createElement("button");
  resetButton.type = "button";
  resetButton.id = "reset-plan";
  resetButton.className = "np-action secondary";
  resetButton.textContent = "Zurücksetzen";
  resetButton.addEventListener("click", () => {
    saveUndo();
    rows = copyRows(initialRows);
    render();
  });
  toolbar.appendChild(resetButton);
  return toolbar;
}

function renderTable() {
  const wrapper = document.createElement("div");
  wrapper.className = "np-grid-wrap";

  const table = document.createElement("table");
  table.className = "np-grid";
  const head = document.createElement("thead");
  const headerRow = document.createElement("tr");

  const nameHeader = document.createElement("th");
  nameHeader.className = "np-name-cell np-name-header";
  nameHeader.textContent = "Mitarbeitende";
  headerRow.appendChild(nameHeader);

  days.forEach((day) => {
    const th = document.createElement("th");
    th.className = day.weekend ? "np-day weekend" : "np-day";
    th.title = day.full_label || day.key;
    th.innerHTML = `<span>${day.day}</span><small>${day.weekday}</small>`;
    headerRow.appendChild(th);
  });
  head.appendChild(headerRow);
  table.appendChild(head);

  const body = document.createElement("tbody");
  rows.forEach((row, rowIndex) => {
    const tr = document.createElement("tr");
    const nameCell = document.createElement("th");
    nameCell.className = "np-name-cell";
    nameCell.textContent = row.name;
    nameCell.title = row.name;
    tr.appendChild(nameCell);

    days.forEach((day) => {
      const value = normalizeCode(row.cells?.[day.key]);
      const locked = Boolean(row.locked?.[day.key]);
      const td = document.createElement("td");
      td.className = `np-cell ${codeClass(value)}${locked ? " locked" : ""}`;
      td.textContent = value;
      td.dataset.locked = String(locked);
      td.title = locked
        ? `${row.name} · ${day.full_label}: geschützte Abwesenheit`
        : `${row.name} · ${day.full_label}`;
      td.addEventListener("pointerdown", (event) => {
        if (locked) return;
        event.preventDefault();
        saveUndo();
        isPainting = true;
        paintCell(rowIndex, day.key, td);
      });
      td.addEventListener("pointerenter", (event) => {
        if (!isPainting || locked) return;
        event.preventDefault();
        paintCell(rowIndex, day.key, td);
      });
      tr.appendChild(td);
    });
    body.appendChild(tr);
  });
  table.appendChild(body);
  wrapper.appendChild(table);
  return wrapper;
}

function renderFooter() {
  const footer = document.createElement("div");
  footer.className = "np-footer";

  const explanation = document.createElement("div");
  explanation.className = "np-explanation";
  explanation.textContent = "Klicken oder mit gedrückter Maustaste über mehrere Felder ziehen.";
  footer.appendChild(explanation);

  const saveButton = document.createElement("button");
  saveButton.type = "button";
  saveButton.id = "save-plan";
  saveButton.className = "np-save";
  saveButton.textContent = "Dienste übernehmen";
  saveButton.addEventListener("click", () => {
    const token = `${Date.now()}-${Math.random().toString(16).slice(2)}`;
    Streamlit.setComponentValue({ token, rows: copyRows(rows) });
  });
  footer.appendChild(saveButton);
  return footer;
}

function render() {
  root.innerHTML = "";
  root.appendChild(renderToolbar());
  root.appendChild(renderTable());
  root.appendChild(renderFooter());
  updateDirtyState();
  requestAnimationFrame(() => Streamlit.setFrameHeight(document.body.scrollHeight + 4));
}

function onRender(event) {
  const args = event.detail.args || {};
  const incomingVersion = String(args.plan_version ?? "");
  days = Array.isArray(args.days) ? args.days : [];
  if (currentVersion !== incomingVersion) {
    rows = copyRows(args.rows);
    initialRows = copyRows(args.rows);
    undoStack = [];
    currentVersion = incomingVersion;
  }
  render();
}

document.addEventListener("pointerup", () => {
  isPainting = false;
});
document.addEventListener("pointercancel", () => {
  isPainting = false;
});

Streamlit.events.addEventListener(Streamlit.RENDER_EVENT, onRender);
Streamlit.setComponentReady();
Streamlit.setFrameHeight(560);
