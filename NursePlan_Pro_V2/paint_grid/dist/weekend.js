function markWeekendColumns() {
  const table = document.querySelector(".np-grid");
  if (!table) return;

  const weekendColumnIndexes = Array.from(table.querySelectorAll("thead th"))
    .map((cell, index) => (cell.classList.contains("weekend") ? index : -1))
    .filter((index) => index >= 0);

  table.querySelectorAll("tbody tr").forEach((row) => {
    weekendColumnIndexes.forEach((index) => {
      const cell = row.children[index];
      if (cell) cell.classList.add("weekend");
    });
  });
}

const app = document.getElementById("app");
if (app) {
  new MutationObserver(markWeekendColumns).observe(app, {
    childList: true,
    subtree: true,
  });
}

markWeekendColumns();
