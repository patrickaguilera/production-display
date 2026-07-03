/* =========================================================
   CONFIGURATION
========================================================= */

const xlsxFileUrl = "data/latest.xlsx";

/*
  ORIGINAL XLSX COLUMN INDEXES (DO NOT CHANGE UNLESS FILE CHANGES)
*/
const COL = {
  STATUS: 2,
  PART: 3,
  QTY: 4,
  DESCRIPTION: 9,
  COMMENTS: 10
};

/* Column layout in the rendered table */
const VIEW = {
  DESCRIPTION_COL: 5
};

/* Duty mode */
let dutyMode = "medium";

/* description/comments toggle */
let descriptionMode = "description";

/* full dataset */
let allRows = [];


/* =========================================================
   DUTY CLASSIFICATION
========================================================= */

function getDutyType(partNumber) {
  if (!partNumber) return "light";

  let str = String(partNumber).trim().toUpperCase();

  // Remove NS- prefix (applies to both types)
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
   RENDER TABLE
========================================================= */

function buildTable(rows) {
  if (!rows || rows.length === 0) {
    return "<p>No data available</p>";
  }

  const colCount = rows[0].length;

  let html = "<table><thead><tr>";

  // HEADER
  for (let c = 0; c < colCount; c++) {
    let header = rows[0][c];

    // Replace Description header with toggle
    if (c === VIEW.DESCRIPTION_COL) {
      header = `
        <span class="desc-toggle" onclick="toggleDescriptionMode()">
          ${descriptionMode === "description" ? "Description ▼" : "Comments ▼"}
        </span>
      `;
    }

    html += `<th>${header}</th>`;
  }

  html += "</tr></thead><tbody>";

  // BODY
  for (let r = 1; r < rows.length; r++) {
    html += "<tr>";

    for (let c = 0; c < colCount; c++) {
      let val = (rows[r] && rows[r][c] !== undefined) ? rows[r][c] : "";
      let className = "";

      /* RIGHT ALIGN QTY */
      if (c === 4) className += "text-right";

      /* REQD DATE COLOR */
      if (c === COL.STATUS) {
        const d = new Date(val);

        if (!isNaN(d)) {
          const today = new Date();
          today.setHours(0,0,0,0);
          d.setHours(0,0,0,0);

          className += d < today ? " status-past" : " status-ok";
        }
      }

      /* DESCRIPTION / COMMENTS TOGGLE */
      if (c === VIEW.DESCRIPTION_COL) {
        const desc = rows[r][COL.DESCRIPTION];
        const comm = rows[r][COL.COMMENTS];

        val = descriptionMode === "description" ? desc : comm;
      }

      /* TOOLTIP */
      const title = `title="${String(val).replace(/"/g, "&quot;")}"`;

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
  return rows.filter((row, i) => {
    if (i === 0) return true;
    const part = row[COL.PART] || "";
    return getDutyType(part) === dutyMode;
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
   LOAD XLSX
========================================================= */

function loadData() {
  fetch(xlsxFileUrl)
    .then(r => r.arrayBuffer())
    .then(data => {
      const wb = XLSX.read(data, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];

      const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });

      // KEEP ALL ROWS (no dangerous filtering yet)
      allRows = rows;

      render();
    })
    .catch(err => {
      console.error(err);
      document.getElementById("output").innerHTML =
        "<p>Error loading file</p>";
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
    descriptionMode === "description"
      ? "comments"
      : "description";

  render();
};


/* =========================================================
   AUTO REFRESH (8AM DAILY)
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
