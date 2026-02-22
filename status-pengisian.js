(function () {
  "use strict";

  const MONTHS = window.NILAM_DATA.MONTHS.slice(0, 11);
  const CURRENT_YEAR = String(new Date().getFullYear());
  const statusEl = document.getElementById("statusInfo");
  const tbody = document.querySelector("#statusTable tbody");

  init();

  async function init() {
    try {
      const data = await window.NILAM_DATA.loadAllData();
      renderStatusTable(data);
      statusEl.textContent = `Status ${CURRENT_YEAR} dijana untuk ${data.studentsByClass.size} kelas.`;
    } catch (error) {
      console.error(error);
      statusEl.textContent = "Gagal memuatkan status pengisian.";
      statusEl.style.color = "#b00020";
      tbody.innerHTML = '<tr><td colspan="12" class="empty">Ralat memuatkan data.</td></tr>';
    }
  }

  function renderStatusTable(data) {
    const recordsForYear = data.records.filter((row) => String(row.tahun || "") === CURRENT_YEAR);
    const classes = [...new Set([
      ...data.studentsByClass.keys(),
      ...recordsForYear.map((row) => row.kelas),
    ])].sort((a, b) => a.localeCompare(b, "ms"));

    if (!classes.length) {
      tbody.innerHTML = '<tr><td colspan="12" class="empty">Tiada kelas dijumpai.</td></tr>';
      return;
    }

    const currentMonthIndex = new Date().getMonth();
    const rowsHtml = classes
      .map((kelas) => {
        const totalStudents = (data.studentsByClass.get(kelas) || []).length;
        const cells = MONTHS.map((month, monthIndex) =>
          renderStatusCell(recordsForYear, kelas, month, monthIndex, currentMonthIndex, totalStudents)
        ).join("");
        return `<tr><th scope="row">${escapeHtml(kelas)}</th>${cells}</tr>`;
      })
      .join("");

    tbody.innerHTML = rowsHtml;
  }

  function renderStatusCell(records, kelas, month, monthIndex, currentMonthIndex, totalStudents) {
    if (monthIndex > currentMonthIndex) {
      return '<td class="status-cell"><span class="light light-gray" aria-hidden="true"></span>-</td>';
    }

    if (!totalStudents) {
      return '<td class="status-cell"><span class="light light-gray" aria-hidden="true"></span>0%</td>';
    }

    const recordsForSlot = records.filter((row) => row.kelas === kelas && row.bulan === month);
    const filledStudents = new Set(
      recordsForSlot
        .filter((row) => window.NILAM_DATA.rowHasAnyInput(row))
        .map((row) => String(row.nama || "").trim().toLowerCase())
        .filter(Boolean)
    );

    const percent = Math.round((filledStudents.size / totalStudents) * 100);
    const isGreen = percent >= 50;
    const lightClass = isGreen ? "light-green" : "light-red";

    return `<td class="status-cell"><span class="light ${lightClass}" aria-hidden="true"></span>${percent}%</td>`;
  }

  function escapeHtml(value) {
    return String(value)
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#39;");
  }
})();
