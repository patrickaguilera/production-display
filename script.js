const hiddenColumnIndexes = [2,4,7,8,10];     // columns to hide
// New visible column indexes
const statusColumnIndex = 2;                // color "PAST DUE" or "OKAY"
const partColumnIndex = 3;                  // part number
const checkColumnIndex = 4;                 // remove rows where this column is empty
const xlsxFileUrl = "data/latest.xlsx";

// Manually define widths for each visible column
const columnWidths = {
  0: "13%",
  1: "10%",
  2: "15%",
  3: "25%",
  4: "8%"
};

function buildHtmlTable(rows) {
  if (rows.length === 0) return "<p>No data to display.</p>";

  const colCount = rows[0].length;

  /* ---------- BUILD TABLE HEADER ---------- */
  let thead = "<thead><tr>";
  for (let i = 0; i < colCount; i++) {
    const width = columnWidths[i] ? `style="width:${columnWidths[i]};"` : "";
    thead += `<th ${width}>${rows[0][i]}</th>`;
  }
  thead += "</tr></thead>";

  /* ---------- BUILD TABLE BODY ---------- */
  let tbody = "<tbody>";

  for (let r = 1; r < rows.length; r++) {
    tbody += "<tr>";

    for (let c = 0; c < colCount; c++) {
      let cellValue = rows[r][c] || "";
      let style = "";
  
      // Apply manual width
      if (columnWidths[c]) {
        style += `width:${columnWidths[c]};`;
      }
  
      // ===== STATUS COLUMN DATE COLORING =====
      if (c === statusColumnIndex) {
        const parsed = new Date(cellValue);
        if (!isNaN(parsed)) {
          const today = new Date();
          today.setHours(0,0,0,0);
          const cellDate = parsed;
          cellDate.setHours(0,0,0,0);

          if (cellDate < today) style += "color:red; font-weight:bold;";
            else style += "color:green; font-weight:bold;";
        }
      }
  
      // Right-align the CHECK column
      if (c === checkColumnIndex) {
        style += "text-align:right;";
      }
  
      // Add tooltip for truncated cells (or last column)
      let titleAttr = "";
      if (c === colCount - 1) {
        titleAttr = ` title="${cellValue}"`;
      }
  
      tbody += `<td style="${style}"${titleAttr}>${cellValue}</td>`;
    }  


    tbody += "</tr>";
  }

  tbody += "</tbody>";

  return `<table>${thead}${tbody}</table>`;
}


// Determine Medium vs Light Duty
function getDutyType(partNumber) {
    if (partNumber === undefined || partNumber === null) return "light";

    let str = String(partNumber).trim().toUpperCase();

    // Remove optional NS- prefix
    if (str.startsWith("NS-")) {
        str = str.substring(3);
    }

    // Must still contain a dash after removing NS-
    if (!str.includes("-")) return "light";

    const firstGroup = str.split("-")[0];

    // Medium Duty prefixes
    if (
        firstGroup.includes("M") ||
        firstGroup.includes("CA") ||
        firstGroup.includes("HRNS")
    ) {
        return "medium";
    }

    return "light";
}

function filterByDuty(rows, type){
  return rows.filter((row, idx) => {
    if (idx === 0) return true;
    return getDutyType(row[partColumnIndex]) === type;
  });
}

function renderFiltered(){
  const type = document.getElementById("orderTypeToggle").value;
  localStorage.setItem("orderTypeToggle", type);
  document.getElementById("output").innerHTML = buildHtmlTable(filterByDuty(allRows, type));
  addEllipsisTooltips();
}

// Load XLSX
function loadAndRenderTable(){
  fetch(xlsxFileUrl)
    .then(res => {
      if (!res.ok) throw new Error("File not found");
      return res.arrayBuffer();
    })
    .then(data => {
      const workbook = XLSX.read(data, { type: "array" });
      const sheet = workbook.SheetNames[0];
      let rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheet], { header: 1 });

      rows = rows.map(row => row.filter((_, idx) => !hiddenColumnIndexes.includes(idx)));

      rows = rows.filter(row => {
        const v = row[checkColumnIndex];
        return v !== undefined && v !== null && v !== "";
      });

      const colCount = rows[0].length;
      rows = rows.map(r => {
        while (r.length < colCount) r.push("");
        return r;
      });

      allRows = rows;

      const saved = localStorage.getItem("orderTypeToggle");
      if (saved) document.getElementById("orderTypeToggle").value = saved;

      renderFiltered();
    })
    .catch(err => {
      console.error(err);
      document.getElementById("output").innerHTML =
        `<p class="error-message">Error loading table.</p>`;
    });
}

loadAndRenderTable();
document.getElementById("orderTypeToggle").addEventListener("change", renderFiltered);

// Add tooltips to first column
function addEllipsisTooltips() {
    const cells = document.querySelectorAll("td");

    cells.forEach(cell => {
        const fullText = cell.textContent.trim();

        // Detect if the content is visually truncated
        const isTruncated =
            cell.scrollWidth > cell.clientWidth ||
            cell.scrollHeight > cell.clientHeight;

        if (isTruncated) {
            cell.setAttribute("title", fullText);
        } else {
            cell.removeAttribute("title");
        }
    });
}
  
// Auto-refresh daily @ 8 AM
function scheduleDailyRefresh(hour = 8, minute = 0){
  const now = new Date();
  let next = new Date();
  next.setHours(hour, minute, 0, 0);
  if (now > next) next.setDate(next.getDate() + 1);

  setTimeout(() => {
    loadAndRenderTable();
    scheduleDailyRefresh(hour, minute);
  }, next - now);
}
scheduleDailyRefresh();
