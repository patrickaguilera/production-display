
/* =========================================================
   CONFIGURATION (EDIT ONLY THIS SECTION IF NEEDED)
========================================================= */

// XLSX file
const xlsxFileUrl = "data/latest.xlsx";

// ORIGINAL XLSX COLUMN MAP (DO NOT USE VISUAL INDEXES)
const COL = {
  STATUS: 2,        // REQD DATE (date-based color)
  PART: 3,          // Part Number
  QTY: 4,           // Right aligned / filter column
  DESCRIPTION: 9,
  COMMENTS: 10
};

// Columns to remove completely from dataset
const hiddenColumnIndexes = [2, 4, 7, 8];

// Column widths (optional visual tuning)
const columnWidths = {
  0: "13%",
  1: "10%",
  2: "15%",
  3: "25%",
  4: "10%"
};

// Duty state
let dutyMode = "medium";

// Description mode toggle
let descriptionMode = "description";

// Full dataset cache
let allRows = [];


/* =========================================================
   DUTY CLASSIFICATION
========================================================= */

function getDutyType(partNumber) {
  if (!partNumber) return "light";

  let str = String(partNumber).trim().toUpperCase();

  // Ignore NS- prefix
  if (str.startsWith("NS-")) {
    str = str.substring(3);
  }

  if (!str.includes("-")) return "light";

  const first = str.split("-")[0];

  if (
    first.includes("M") ||
    first.includes("CA") ||
    first.includes("HRNS")
  ) {
    return "medium";
  }

  return "light";
}


/* =========================================================
   TABLE BUILDER
========================================================= */

function buildTable(rows) {
  if (!rows.length) return "<p>No data</p>";

  const colCount = rows[0].length;

  let html = "<table><thead><tr>";

  // HEADER
  for (let i = 0; i < colCount; i++) {
    let label = rows[0][i];

    // Replace Description header with toggle behavior
    if (i === 5) {
      label = `
        <span class="desc-toggle" onclick="toggleDescriptionMode()">
          ${descriptionMode === "description" ? "Description ▼" : "Comments ▼"}
        </span>
      `;
    }

    html += `<th style="width:${columnWidths[i] || "auto"}">${label}</th>`;
  }

  html += "</tr></thead><tbody>";

  // BODY
  for (let r = 1; r < rows.length; r++) {
    html += "<tr>";

    for (let c = 0; c < colCount; c++) {
      let val = rows[r][c] ?? "";
      let style = "";
      let className = "";

      // RIGHT ALIGN QTY
      if (c === 4) className += "text-right";

      // STATUS DATE COLOR
      if (c === COL.STATUS) {
        const d = new Date(val);
        if (!isNaN(d)) {
          const today = new Date();
          today.setHours(0,0,0,0);
          d.setHours(0,0,0,0);

          className += d < today ? " status-past" : " status-ok";
        }
      }

      // DESCRIPTION / COMMENTS SWITCH
      if (c === 5) {
        val =
          descriptionMode === "description"
            ? rows[r].description
            : rows[r].comments;
      }

      // TOOLTIP IF TRUNCATED
      const title = `title="${String(val).replace(/"/g,'&quot;')}"`;

      html += `<td class="${className}" ${title}>${val}</td>`;
    }

    html += "</tr>";
  }

  html += "</tbody></table>";

  return html;
}


/* =========================================================
   FILTERING
========================================================= */

function filterByDuty(rows) {
  return rows.filter((r, i) => {
    if (i === 0) return true;
    return r.duty === dutyMode;
  });
}


/* =========================================================
   RENDER
========================================================= */

function render() {
  const filtered = filterByDuty(allRows);
  document.getElementById("output").innerHTML = buildTable(filtered);
}


/* =========================================================
   XLSX LOADER
========================================================= */

function loadData() {
  fetch(xlsxFileUrl)
    .then(r => r.arrayBuffer())
    .then(data => {
      const wb = XLSX.read(data, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];

      let raw = XLSX.utils.sheet_to_json(ws, { header: 1 });

      // Build enriched dataset BEFORE column removal
      allRows = raw.map(row => {
        return {
          original: row,
          duty: getDutyType(row[COL.PART]),
          qty: row[COL.QTY],
          status: row[COL.STATUS],
          part: row[COL.PART],
          description: row[COL.DESCRIPTION],
          comments: row[COL.COMMENTS]
        };
      });

      render();
    });
}


/* =========================================================
   TOGGLES
========================================================= */

document.getElementById("mediumBtn").onclick = () => {
  dutyMode = "medium";
  document.getElementById("mediumBtn").classList.add("active");
  document.getElementById("lightBtn").classList.remove("active");
  render();
};

document.getElementById("lightBtn").onclick = () => {
  dutyMode = "light";
  document.getElementById("lightBtn").classList.add("active");
  document.getElementById("mediumBtn").classList.remove("active");
  render();
};


/* Description / Comments toggle */
window.toggleDescriptionMode = function () {
  descriptionMode =
    descriptionMode === "description" ? "comments" : "description";
  render();
};


/* =========================================================
   AUTO REFRESH (8 AM DAILY)
========================================================= */

function scheduleRefresh(hour = 8) {
  const now = new Date();
  const next = new Date();

  next.setHours(hour, 0, 0, 0);
  if (now > next) next.setDate(next.getDate() + 1);

  setTimeout(() => {
    loadData();
    scheduleRefresh(hour);
  }, next - now);
}


/* =========================================================
   STARTUP
========================================================= */

loadData();
scheduleRefresh(8);
