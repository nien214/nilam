(function () {
  "use strict";

  const MONTHS = window.NILAM_DATA.MONTHS;
  const MATERIAL_KEYS = [
    ["bahan_digital", "Bahan Digital"],
    ["bahan_bukan_buku", "Bahan Bukan Buku"],
    ["fiksyen", "Fiksyen"],
    ["bukan_fiksyen", "Bukan Fiksyen"],
    ["bahasa_melayu", "Bahasa Melayu"],
    ["bahasa_inggeris", "Bahasa Inggeris"],
    ["lain_lain_bahasa", "Lain-lain Bahasa"],
  ];
  const PIE_COLORS = ["#1b9aaa", "#ef476f", "#ffd166", "#06d6a0", "#118ab2", "#ff8fab", "#a0c4ff"];
  const BAR_COLORS = ["#127475", "#ef476f", "#ffd166", "#06d6a0", "#118ab2", "#f97316", "#8b5cf6", "#0ea5e9", "#22c55e", "#e11d48"];

  const el = {
    tahun: document.getElementById("statTahun"),
    bulan: document.getElementById("statBulan"),
    kelas: document.getElementById("statKelas"),
    status: document.getElementById("statStatus"),
    pieTitle: document.getElementById("pieTitle"),
    barTitle: document.getElementById("barTitle"),
    summaryHeadRow: document.getElementById("summaryHeadRow"),
    summaryTbody: document.getElementById("summaryTbody"),
    pieWrap: document.getElementById("pieWrap"),
    barWrap: document.getElementById("barWrap"),
  };

  const state = {
    records: [],
    classes: [],
    selectedYear: String(new Date().getFullYear()),
    selectedMonth: "",
    selectedClass: "",
  };

  init();

  async function init() {
    try {
      const data = await window.NILAM_DATA.loadAllData();
      state.records = data.records;
      state.classes = [...new Set([...data.studentsByClass.keys(), ...data.records.map((row) => row.kelas)])].sort(
        (a, b) => a.localeCompare(b, "ms")
      );

      initDropdowns();
      renderAll();
      bindEvents();
    } catch (error) {
      console.error(error);
      el.status.textContent = "Gagal memuatkan statistik.";
      el.status.style.color = "#b00020";
    }
  }

  function initDropdowns() {
    if (el.tahun) {
      el.tahun.value = state.selectedYear;
    }

    const yearlyOption = document.createElement("option");
    yearlyOption.value = "__year__";
    yearlyOption.textContent = "Sepanjang Tahun";
    el.bulan.appendChild(yearlyOption);

    MONTHS.forEach((month) => {
      const option = document.createElement("option");
      option.value = month;
      option.textContent = month;
      el.bulan.appendChild(option);
    });

    el.bulan.value = "__year__";
    state.selectedMonth = "__year__";

    el.kelas.innerHTML = '<option value="">Pilih kelas</option>';
    const allOption = document.createElement("option");
    allOption.value = "__all__";
    allOption.textContent = "Semua Kelas";
    el.kelas.appendChild(allOption);

    state.classes.forEach((kelas) => {
      const option = document.createElement("option");
      option.value = kelas;
      option.textContent = kelas;
      el.kelas.appendChild(option);
    });

    el.kelas.value = "__all__";
    state.selectedClass = "__all__";
  }

  function bindEvents() {
    el.bulan.addEventListener("change", () => {
      state.selectedMonth = el.bulan.value;
      renderAll();
    });

    el.kelas.addEventListener("change", () => {
      state.selectedClass = el.kelas.value;
      renderAll();
    });
  }

  function renderAll() {
    if (!state.selectedMonth || !state.selectedClass) {
      el.status.textContent = "Sila pilih bulan dan kelas.";
      renderEmptySummary(true);
      el.pieWrap.innerHTML = '<p class="empty">Tiada data.</p>';
      el.barWrap.innerHTML = '<p class="empty">Tiada data.</p>';
      return;
    }

    const isYearMode = state.selectedMonth === "__year__";
    const isAllClasses = state.selectedClass === "__all__";
    const periodText = isYearMode ? "Sepanjang Tahun" : state.selectedMonth;
    const classText = isAllClasses ? "Semua Kelas" : state.selectedClass;
    updateChartTitles(classText, periodText);

    const periodRecords = state.records.filter((row) => {
      if (String(row.tahun || "") !== state.selectedYear) {
        return false;
      }
      if (isYearMode) {
        return true;
      }
      return row.bulan === state.selectedMonth;
    });

    renderSummaryTable(periodRecords, isAllClasses, state.selectedClass);

    const filtered = isAllClasses
      ? periodRecords
      : periodRecords.filter((row) => row.kelas === state.selectedClass);

    if (!filtered.length) {
      el.status.textContent = `Tiada rekod untuk ${state.selectedYear} ${periodText} ${classText}.`;
      el.pieWrap.innerHTML = '<p class="empty">Tiada data jenis bahan.</p>';
      el.barWrap.innerHTML = '<p class="empty">Tiada data jumlah bacaan.</p>';
      return;
    }

    el.status.textContent = `${filtered.length} rekod dianalisis untuk ${state.selectedYear} ${periodText} ${classText}.`;
    renderPie(filtered);
    renderBar(filtered);
  }

  function updateChartTitles(classText, periodText) {
    if (el.pieTitle) {
      el.pieTitle.textContent = `Jenis Bahan Dibaca (${classText} / ${periodText})`;
    }
    if (el.barTitle) {
      el.barTitle.textContent = `Jumlah Bacaan Murid (${classText} / ${periodText})`;
    }
  }

  function renderSummaryTable(periodRecords, isAllClasses, selectedClass) {
    if (!el.summaryTbody || !el.summaryHeadRow) {
      return;
    }

    if (!periodRecords.length) {
      renderEmptySummary(isAllClasses);
      return;
    }

    if (isAllClasses) {
      renderSummaryAllClasses(periodRecords);
      return;
    }

    renderSummaryOneClass(periodRecords, selectedClass);
  }

  function renderEmptySummary(isAllClasses) {
    setSummaryHeaders(isAllClasses);
    if (el.summaryTbody) {
      el.summaryTbody.innerHTML = '<tr><td colspan="10" class="empty">Tiada data ringkasan.</td></tr>';
    }
  }

  function setSummaryHeaders(isAllClasses) {
    if (!el.summaryHeadRow) {
      return;
    }
    if (isAllClasses) {
      el.summaryHeadRow.innerHTML = `
        <th>Bil</th>
        <th>Kelas</th>
        <th>Bahan Digital</th>
        <th>Bahan Bukan Buku</th>
        <th>Fiksyen</th>
        <th>Bukan Fiksyen</th>
        <th>Bahasa Melayu</th>
        <th>Bahasa Inggeris</th>
        <th>Lain-lain Bahasa</th>
        <th>JUMLAH BACAAN</th>
      `;
      return;
    }

    el.summaryHeadRow.innerHTML = `
      <th>Bil</th>
      <th>Nama Murid</th>
      <th>Bahan Digital</th>
      <th>Bahan Bukan Buku</th>
      <th>Fiksyen</th>
      <th>Bukan Fiksyen</th>
      <th>Bahasa Melayu</th>
      <th>Bahasa Inggeris</th>
      <th>Lain-lain Bahasa</th>
      <th>JUMLAH BACAAN</th>
    `;
  }

  function renderSummaryAllClasses(periodRecords) {
    setSummaryHeaders(true);

    const byClass = new Map();
    periodRecords.forEach((row) => {
      const kelas = String(row.kelas || "").trim();
      if (!kelas) {
        return;
      }
      if (!byClass.has(kelas)) {
        byClass.set(kelas, {
          kelas,
          bahan_digital: 0,
          bahan_bukan_buku: 0,
          fiksyen: 0,
          bukan_fiksyen: 0,
          bahasa_melayu: 0,
          bahasa_inggeris: 0,
          lain_lain_bahasa: 0,
          jumlah_bacaan: 0,
        });
      }

      const slot = byClass.get(kelas);
      slot.bahan_digital += Number(row.bahan_digital || 0);
      slot.bahan_bukan_buku += Number(row.bahan_bukan_buku || 0);
      slot.fiksyen += Number(row.fiksyen || 0);
      slot.bukan_fiksyen += Number(row.bukan_fiksyen || 0);
      slot.bahasa_melayu += Number(row.bahasa_melayu || 0);
      slot.bahasa_inggeris += Number(row.bahasa_inggeris || 0);
      slot.lain_lain_bahasa += Number(row.lain_lain_bahasa || 0);
      slot.jumlah_bacaan += Number(row.jumlah_aktiviti || 0);
    });

    const rows = [...byClass.values()].sort((a, b) => a.kelas.localeCompare(b.kelas, "ms"));
    const total = {
      bahan_digital: 0,
      bahan_bukan_buku: 0,
      fiksyen: 0,
      bukan_fiksyen: 0,
      bahasa_melayu: 0,
      bahasa_inggeris: 0,
      lain_lain_bahasa: 0,
      jumlah_bacaan: 0,
    };
    rows.forEach((row) => {
      total.bahan_digital += row.bahan_digital;
      total.bahan_bukan_buku += row.bahan_bukan_buku;
      total.fiksyen += row.fiksyen;
      total.bukan_fiksyen += row.bukan_fiksyen;
      total.bahasa_melayu += row.bahasa_melayu;
      total.bahasa_inggeris += row.bahasa_inggeris;
      total.lain_lain_bahasa += row.lain_lain_bahasa;
      total.jumlah_bacaan += row.jumlah_bacaan;
    });

    el.summaryTbody.innerHTML =
      rows
      .map(
        (row, index) => `
      <tr>
        <td>${index + 1}</td>
        <td>${escapeHtml(row.kelas)}</td>
        <td>${row.bahan_digital}</td>
        <td>${row.bahan_bukan_buku}</td>
        <td>${row.fiksyen}</td>
        <td>${row.bukan_fiksyen}</td>
        <td>${row.bahasa_melayu}</td>
        <td>${row.bahasa_inggeris}</td>
        <td>${row.lain_lain_bahasa}</td>
        <td>${row.jumlah_bacaan}</td>
      </tr>`
      )
      .join("") +
      `
      <tr>
        <td><strong></strong></td>
        <td><strong>JUMLAH</strong></td>
        <td><strong>${total.bahan_digital}</strong></td>
        <td><strong>${total.bahan_bukan_buku}</strong></td>
        <td><strong>${total.fiksyen}</strong></td>
        <td><strong>${total.bukan_fiksyen}</strong></td>
        <td><strong>${total.bahasa_melayu}</strong></td>
        <td><strong>${total.bahasa_inggeris}</strong></td>
        <td><strong>${total.lain_lain_bahasa}</strong></td>
        <td><strong>${total.jumlah_bacaan}</strong></td>
      </tr>`;
  }

  function renderSummaryOneClass(periodRecords, selectedClass) {
    setSummaryHeaders(false);

    const byStudent = new Map();
    periodRecords
      .filter((row) => row.kelas === selectedClass)
      .forEach((row) => {
        const key = String(row.no_kad_pengenalan || row.nama || "").trim();
        if (!byStudent.has(key)) {
          byStudent.set(key, {
            nama: String(row.nama || "").trim(),
            bahan_digital: 0,
            bahan_bukan_buku: 0,
            fiksyen: 0,
            bukan_fiksyen: 0,
            bahasa_melayu: 0,
            bahasa_inggeris: 0,
            lain_lain_bahasa: 0,
            jumlah_bacaan: 0,
          });
        }

        const slot = byStudent.get(key);
        slot.bahan_digital += Number(row.bahan_digital || 0);
        slot.bahan_bukan_buku += Number(row.bahan_bukan_buku || 0);
        slot.fiksyen += Number(row.fiksyen || 0);
        slot.bukan_fiksyen += Number(row.bukan_fiksyen || 0);
        slot.bahasa_melayu += Number(row.bahasa_melayu || 0);
        slot.bahasa_inggeris += Number(row.bahasa_inggeris || 0);
        slot.lain_lain_bahasa += Number(row.lain_lain_bahasa || 0);
        slot.jumlah_bacaan += Number(row.jumlah_aktiviti || 0);
      });

    const rows = [...byStudent.values()].sort((a, b) => a.nama.localeCompare(b.nama, "ms"));
    if (!rows.length) {
      el.summaryTbody.innerHTML = '<tr><td colspan="10" class="empty">Tiada data ringkasan.</td></tr>';
      return;
    }

    const total = {
      bahan_digital: 0,
      bahan_bukan_buku: 0,
      fiksyen: 0,
      bukan_fiksyen: 0,
      bahasa_melayu: 0,
      bahasa_inggeris: 0,
      lain_lain_bahasa: 0,
      jumlah_bacaan: 0,
    };
    rows.forEach((row) => {
      total.bahan_digital += row.bahan_digital;
      total.bahan_bukan_buku += row.bahan_bukan_buku;
      total.fiksyen += row.fiksyen;
      total.bukan_fiksyen += row.bukan_fiksyen;
      total.bahasa_melayu += row.bahasa_melayu;
      total.bahasa_inggeris += row.bahasa_inggeris;
      total.lain_lain_bahasa += row.lain_lain_bahasa;
      total.jumlah_bacaan += row.jumlah_bacaan;
    });

    el.summaryTbody.innerHTML =
      rows
      .map(
        (row, index) => `
      <tr>
        <td>${index + 1}</td>
        <td>${escapeHtml(row.nama)}</td>
        <td>${row.bahan_digital}</td>
        <td>${row.bahan_bukan_buku}</td>
        <td>${row.fiksyen}</td>
        <td>${row.bukan_fiksyen}</td>
        <td>${row.bahasa_melayu}</td>
        <td>${row.bahasa_inggeris}</td>
        <td>${row.lain_lain_bahasa}</td>
        <td>${row.jumlah_bacaan}</td>
      </tr>`
      )
      .join("") +
      `
      <tr>
        <td><strong></strong></td>
        <td><strong>JUMLAH</strong></td>
        <td><strong>${total.bahan_digital}</strong></td>
        <td><strong>${total.bahan_bukan_buku}</strong></td>
        <td><strong>${total.fiksyen}</strong></td>
        <td><strong>${total.bukan_fiksyen}</strong></td>
        <td><strong>${total.bahasa_melayu}</strong></td>
        <td><strong>${total.bahasa_inggeris}</strong></td>
        <td><strong>${total.lain_lain_bahasa}</strong></td>
        <td><strong>${total.jumlah_bacaan}</strong></td>
      </tr>`;
  }

  function renderPie(records) {
    const totals = MATERIAL_KEYS.map(([key, label], index) => ({
      key,
      label,
      value: records.reduce((sum, row) => sum + Number(row[key] || 0), 0),
      color: PIE_COLORS[index % PIE_COLORS.length],
    }));

    const grandTotal = totals.reduce((sum, row) => sum + row.value, 0);
    if (!grandTotal) {
      el.pieWrap.innerHTML = '<p class="empty">Semua nilai jenis bahan adalah 0.</p>';
      return;
    }

    let startAngle = -Math.PI / 2;
    const cx = 130;
    const cy = 130;
    const radius = 90;

    const slices = totals
      .filter((item) => item.value > 0)
      .map((item) => {
        const angle = (item.value / grandTotal) * Math.PI * 2;
        const endAngle = startAngle + angle;
        const x1 = cx + radius * Math.cos(startAngle);
        const y1 = cy + radius * Math.sin(startAngle);
        const x2 = cx + radius * Math.cos(endAngle);
        const y2 = cy + radius * Math.sin(endAngle);
        const largeArc = angle > Math.PI ? 1 : 0;

        const path = `M ${cx} ${cy} L ${x1} ${y1} A ${radius} ${radius} 0 ${largeArc} 1 ${x2} ${y2} Z`;
        startAngle = endAngle;
        return `<path d="${path}" fill="${item.color}"></path>`;
      })
      .join("");

    const legend = totals
      .map((item) => {
        const percent = grandTotal ? Math.round((item.value / grandTotal) * 100) : 0;
        return `<li><i style="background:${item.color}"></i>${item.label}: ${item.value} (${percent}%)</li>`;
      })
      .join("");

    el.pieWrap.innerHTML = `
      <div class="chart-layout">
        <svg viewBox="0 0 260 260" class="pie-svg" aria-label="Carta pai jenis bahan">
          ${slices}
        </svg>
        <ul class="chart-legend">${legend}</ul>
      </div>
    `;
  }

  function renderBar(records) {
    const totalsByStudent = new Map();
    records.forEach((row) => {
      const key = String(row.no_kad_pengenalan || row.nama || "").trim();
      const nama = String(row.nama || "").trim();
      const jumlah = Number.isFinite(Number(row.jumlah_aktiviti))
        ? Number(row.jumlah_aktiviti)
        : Number(row.bahasa_melayu || 0) + Number(row.bahasa_inggeris || 0) + Number(row.lain_lain_bahasa || 0);

      if (!totalsByStudent.has(key)) {
        totalsByStudent.set(key, { nama, jumlah: 0 });
      }
      const current = totalsByStudent.get(key);
      current.jumlah += Math.max(0, Math.trunc(jumlah || 0));
    });

    const rows = [...totalsByStudent.values()]
      .map((row) => ({
        nama: String(row.nama || "").trim(),
        jumlah: Math.max(0, Math.trunc(Number(row.jumlah) || 0)),
      }))
      .sort((a, b) => b.jumlah - a.jumlah)
      .slice(0, 10);

    const maxValue = Math.max(...rows.map((r) => r.jumlah), 1);
    const maxNameLen = Math.max(...rows.map((r) => r.nama.length), 1);
    const chartHeight = Math.max(260, 150 + maxNameLen * 7);
    const barWidth = 22;
    const gap = 10;
    const left = 40;
    const right = 20;
    const top = 20;
    const bottom = Math.max(120, 24 + maxNameLen * 7);
    const plotHeight = chartHeight - top - bottom;
    const plotWidth = rows.length * (barWidth + gap);
    const width = Math.max(620, left + plotWidth + right);

    const bars = rows
      .map((row, i) => {
        const h = (row.jumlah / maxValue) * plotHeight;
        const x = left + i * (barWidth + gap);
        const y = top + (plotHeight - h);
        const fullName = escapeHtml(row.nama);
        const barColor = BAR_COLORS[i % BAR_COLORS.length];
        const labelY = top + plotHeight + 16;
        return `
          <rect x="${x}" y="${y}" width="${barWidth}" height="${h}" fill="${barColor}"></rect>
          <text x="${x + barWidth / 2}" y="${top + plotHeight + 12}" text-anchor="middle" class="axis-label">${row.jumlah}</text>
          <text x="${x + barWidth / 2}" y="${labelY}" text-anchor="start" class="axis-label" transform="rotate(90 ${x + barWidth / 2} ${labelY})">${fullName}</text>
        `;
      })
      .join("");

    el.barWrap.innerHTML = `
      <div class="bar-scroll">
        <svg viewBox="0 0 ${width} ${chartHeight}" class="bar-svg" aria-label="Carta bar jumlah bacaan">
          <line x1="${left}" y1="${top + plotHeight}" x2="${width - right}" y2="${top + plotHeight}" stroke="#9bb7bf" />
          ${bars}
        </svg>
      </div>
    `;
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
