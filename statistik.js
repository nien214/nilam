(function () {
  "use strict";

  const MONTHS = window.NILAM_DATA.MONTHS;
  const MATERIAL_DONUT_KEYS = [
    ["bahan_digital", "Bahan Digital"],
    ["bahan_bukan_buku", "Bahan Bukan Buku"],
    ["bukan_fiksyen", "Buku Bukan Fiksyen"],
    ["fiksyen", "Buku Fiksyen"],
  ];
  const LANGUAGE_KEYS = [
    ["bahasa_melayu", "Bahasa Melayu"],
    ["bahasa_inggeris", "Bahasa Inggeris"],
    ["lain_lain_bahasa", "Lain-lain Bahasa"],
  ];
  const TINGKATAN_ORDER = ["PER", "1", "2", "3", "4", "5"];
  const PIE_COLORS = ["#1b9aaa", "#ef476f", "#ffd166", "#06d6a0", "#118ab2", "#ff8fab", "#a0c4ff"];
  const BAR_COLORS = ["#127475", "#ef476f", "#ffd166", "#06d6a0", "#118ab2", "#f97316", "#8b5cf6", "#0ea5e9", "#22c55e", "#e11d48"];

  const el = {
    tahun: document.getElementById("statTahun"),
    bulan: document.getElementById("statBulan"),
    tingkatan: document.getElementById("statTingkatan"),
    kelas: document.getElementById("statKelas"),
    includeAinsCheckbox: document.getElementById("statIncludeAinsJumlah"),
    status: document.getElementById("statStatus"),
    pieTitle: document.getElementById("pieTitle"),
    languageTitle: document.getElementById("languageTitle"),
    tingkatanTitle: document.getElementById("tingkatanTitle"),
    barTitle: document.getElementById("barTitle"),
    starDistTitle: document.getElementById("starDistTitle"),
    heatmapTitle: document.getElementById("heatmapTitle"),
    summaryHeadRow: document.getElementById("summaryHeadRow"),
    summaryTbody: document.getElementById("summaryTbody"),
    pieWrap: document.getElementById("pieWrap"),
    languageWrap: document.getElementById("languageWrap"),
    tingkatanWrap: document.getElementById("tingkatanWrap"),
    barWrap: document.getElementById("barWrap"),
    starDistWrap: document.getElementById("starDistWrap"),
    heatmapWrap: document.getElementById("heatmapWrap"),
  };

  const state = {
    records: [],
    students: [],
    allowedStudentKeysByClass: new Map(),
    classByNoKad: new Map(),
    nameByNoKad: new Map(),
    classesByName: new Map(),
    allClasses: [],
    classes: [],
    selectedYear: String(new Date().getFullYear()),
    selectedMonth: "",
    selectedTingkatan: "__all__",
    selectedClass: "",
    includeAinsInJumlah: true,
  };

  init();

  async function init() {
    try {
      const data = await window.NILAM_DATA.loadAllData();
      state.records = data.records;
      state.students = Array.isArray(data.students) ? data.students : [];
      state.allowedStudentKeysByClass = buildAllowedStudentKeysByClass(state.students);
      state.classByNoKad = buildClassByNoKad(state.students);
      state.nameByNoKad = buildNameByNoKad(state.students);
      state.classesByName = buildClassesByName(state.students);
      state.allClasses = [...data.studentsByClass.keys()].sort(
        (a, b) => a.localeCompare(b, "ms")
      );
      state.classes = [...state.allClasses];

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
    yearlyOption.textContent = getSehinggaLabel(state.records, state.selectedYear);
    el.bulan.appendChild(yearlyOption);

    MONTHS.forEach((month) => {
      const option = document.createElement("option");
      option.value = month;
      option.textContent = month;
      el.bulan.appendChild(option);
    });

    el.bulan.value = "__year__";
    state.selectedMonth = "__year__";

    el.tingkatan.innerHTML = "";
    const allTingkatanOption = document.createElement("option");
    allTingkatanOption.value = "__all__";
    allTingkatanOption.textContent = "Semua Tingkatan";
    el.tingkatan.appendChild(allTingkatanOption);
    ["PER", "1", "2", "3", "4", "5"].forEach((tingkatan) => {
      const option = document.createElement("option");
      option.value = tingkatan;
      option.textContent = tingkatan;
      el.tingkatan.appendChild(option);
    });
    el.tingkatan.value = "__all__";
    state.selectedTingkatan = "__all__";

    initClassDropdown();
  }

  function initClassDropdown() {
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

    el.tingkatan.addEventListener("change", () => {
      state.selectedTingkatan = el.tingkatan.value;
      state.classes = filterClassesByTingkatan(state.allClasses, state.selectedTingkatan);
      initClassDropdown();
      renderAll();
    });

    el.kelas.addEventListener("change", () => {
      state.selectedClass = el.kelas.value;
      renderAll();
    });
    if (el.includeAinsCheckbox) {
      el.includeAinsCheckbox.checked = true;
      el.includeAinsCheckbox.addEventListener("change", () => {
        state.includeAinsInJumlah = Boolean(el.includeAinsCheckbox.checked);
        renderAll();
      });
    }
  }

  function renderAll() {
    if (!state.selectedMonth || !state.selectedClass) {
      el.status.textContent = "Sila pilih bulan dan kelas.";
      renderEmptySummary(true);
      el.pieWrap.innerHTML = '<p class="empty">Tiada data.</p>';
      el.languageWrap.innerHTML = '<p class="empty">Tiada data.</p>';
      el.tingkatanWrap.innerHTML = '<p class="empty">Tiada data.</p>';
      el.barWrap.innerHTML = '<p class="empty">Tiada data.</p>';
      el.starDistWrap.innerHTML = '<p class="empty">Tiada data.</p>';
      el.heatmapWrap.innerHTML = '<p class="empty">Tiada data.</p>';
      return;
    }

    const isYearMode = state.selectedMonth === "__year__";
    const isAllClasses = state.selectedClass === "__all__";
    const periodText = isYearMode
      ? getSehinggaLabel(state.records, state.selectedYear)
      : state.selectedMonth;
    const tingkatanText = state.selectedTingkatan === "__all__" ? "Semua Tingkatan" : `Tingkatan ${state.selectedTingkatan}`;
    const classText = isAllClasses ? "Semua Kelas" : state.selectedClass;
    const recentMonthText = getMostRecentMonthText(state.records, state.selectedYear, state.selectedMonth);
    updateChartTitles(classText, periodText, tingkatanText, recentMonthText);

    const yearlyRecords = state.records
      .map(resolveRecordClassByMaster)
      .filter((row) => {
        return String(row.tahun || "") === state.selectedYear;
      });

    const yearlyMasterFiltered = filterRecordsByMasterList(yearlyRecords);
    const yearlyTingkatanRecords = state.selectedTingkatan === "__all__"
      ? yearlyMasterFiltered
      : yearlyMasterFiltered.filter((row) => classToTingkatan(row.kelas) === state.selectedTingkatan);

    const periodRecords = isYearMode
      ? yearlyTingkatanRecords
      : yearlyTingkatanRecords.filter((row) => row.bulan === state.selectedMonth);

    renderSummaryTable(periodRecords, isAllClasses, state.selectedClass, yearlyTingkatanRecords);

    const filtered = isAllClasses
      ? periodRecords
      : periodRecords.filter((row) => row.kelas === state.selectedClass);

    if (!filtered.length) {
      el.status.textContent = `Tiada rekod untuk ${state.selectedYear} ${periodText} ${tingkatanText} ${classText}.`;
      el.pieWrap.innerHTML = '<p class="empty">Tiada data jenis bahan.</p>';
      el.languageWrap.innerHTML = '<p class="empty">Tiada data bahasa.</p>';
      el.tingkatanWrap.innerHTML = '<p class="empty">Tiada data tingkatan.</p>';
      el.barWrap.innerHTML = '<p class="empty">Tiada data jumlah bacaan.</p>';
      renderStarDistribution(periodRecords, state.selectedClass, state.selectedTingkatan);
      el.heatmapWrap.innerHTML = '<p class="empty">Tiada data kelas.</p>';
      return;
    }

    el.status.textContent = `${filtered.length} rekod dianalisis untuk ${state.selectedYear} ${periodText} ${tingkatanText} ${classText}.`;
    renderPie(filtered);
    renderLanguageBars(filtered);
    renderTingkatanBars(periodRecords);
    renderBar(filtered);
    renderStarDistribution(periodRecords, state.selectedClass, state.selectedTingkatan);
    renderClassHeatmap(filtered);
  }

  function updateChartTitles(classText, periodText, tingkatanText, recentMonthText) {
    if (el.pieTitle) {
      el.pieTitle.textContent = `Analisis Bacaan Mengikut Jenis Bahan (${classText} / ${periodText})`;
    }
    if (el.languageTitle) {
      el.languageTitle.textContent = `Pilihan Bahasa (${classText} / ${periodText})`;
    }
    if (el.tingkatanTitle) {
      el.tingkatanTitle.textContent = `Analisis Jumlah Bacaan Mengikut Tingkatan (${periodText})`;
    }
    if (el.barTitle) {
      el.barTitle.textContent = `Jumlah Bacaan Murid Tertinggi sehingga ${recentMonthText} (${classText})`;
    }
    if (el.starDistTitle) {
      const starScopeText = classText === "Semua Kelas" ? "Semua Murid" : `Murid ${classText}`;
      el.starDistTitle.textContent = `Taburan Bintang 1-5 (${starScopeText} / ${periodText})`;
    }
    if (el.heatmapTitle) {
      el.heatmapTitle.textContent = `Analisis Jumlah Bacaan Mengikut Kelas (${tingkatanText} / ${periodText})`;
    }
  }

  function renderSummaryTable(periodRecords, isAllClasses, selectedClass, yearlyRecords) {
    if (!el.summaryTbody || !el.summaryHeadRow) {
      return;
    }

    if (!periodRecords.length) {
      renderEmptySummary(isAllClasses);
      return;
    }

    if (isAllClasses) {
      renderSummaryAllClasses(periodRecords, yearlyRecords);
      return;
    }

    renderSummaryOneClass(periodRecords, selectedClass, yearlyRecords);
  }

  function renderEmptySummary(isAllClasses) {
    setSummaryHeaders(isAllClasses);
    if (el.summaryTbody) {
      el.summaryTbody.innerHTML = '<tr><td colspan="12" class="empty">Tiada data ringkasan.</td></tr>';
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
        <th>AINS (Sepanjang Tahun)</th>
        <th>JUMLAH BACAAN</th>
        <th>Bintang</th>
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
      <th>AINS (Sepanjang Tahun)</th>
      <th>JUMLAH BACAAN</th>
      <th>Bintang</th>
    `;
  }

  function renderSummaryAllClasses(periodRecords, yearlyRecords) {
    setSummaryHeaders(true);

    const byClass = new Map();
    const yearAinsByClass = new Map();
    const maxAinsByStudent = new Map();
    (Array.isArray(yearlyRecords) ? yearlyRecords : []).forEach((row) => {
      const kelas = String(row.kelas || "").trim();
      if (!kelas) {
        return;
      }
      const noKad = normalizeKeyText(row.no_kad_pengenalan);
      const nama = normalizeKeyText(row.nama);
      const studentKey = noKad ? `ic:${noKad}` : nama ? `nm:${nama}|k:${kelas.toLowerCase()}` : "";
      if (!studentKey) {
        return;
      }
      const current = maxAinsByStudent.get(studentKey);
      const nextAins = Math.max(0, Math.trunc(Number(row.ains || 0)));
      if (!current || nextAins > current.ains) {
        maxAinsByStudent.set(studentKey, { kelas, ains: nextAins });
      }
    });
    maxAinsByStudent.forEach((entry) => {
      const kelas = String(entry.kelas || "").trim();
      if (!kelas) {
        return;
      }
      yearAinsByClass.set(kelas, (yearAinsByClass.get(kelas) || 0) + Number(entry.ains || 0));
    });

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
          ains_sepanjang_tahun: 0,
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
      slot.jumlah_bacaan += computeJumlahBacaan(row);
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
      ains_sepanjang_tahun: 0,
      jumlah_bacaan: 0,
    };
    rows.forEach((row) => {
      row.ains_sepanjang_tahun = yearAinsByClass.get(row.kelas) || 0;
      total.bahan_digital += row.bahan_digital;
      total.bahan_bukan_buku += row.bahan_bukan_buku;
      total.fiksyen += row.fiksyen;
      total.bukan_fiksyen += row.bukan_fiksyen;
      total.bahasa_melayu += row.bahasa_melayu;
      total.bahasa_inggeris += row.bahasa_inggeris;
      total.lain_lain_bahasa += row.lain_lain_bahasa;
      total.ains_sepanjang_tahun += row.ains_sepanjang_tahun;
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
        <td>${row.ains_sepanjang_tahun}</td>
        <td>${row.jumlah_bacaan}</td>
        <td>${renderBintang(row.jumlah_bacaan)}</td>
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
        <td><strong>${total.ains_sepanjang_tahun}</strong></td>
        <td><strong>${total.jumlah_bacaan}</strong></td>
        <td><strong>${renderBintang(total.jumlah_bacaan)}</strong></td>
      </tr>`;
  }

  function renderSummaryOneClass(periodRecords, selectedClass, yearlyRecords) {
    setSummaryHeaders(false);

    const byStudent = new Map();
    const byName = new Map();

    // Always seed from current master list so every active student appears, even with 0 record.
    state.students
      .filter((row) => String(row.kelas || "").trim() === selectedClass)
      .forEach((row) => {
        const noKad = String(row.no_kad_pengenalan || "").trim();
        const nama = String(row.nama || "").trim();
        if (!nama) {
          return;
        }
        const key = noKad || `nm:${nama.toLowerCase()}`;
        if (!byStudent.has(key)) {
          byStudent.set(key, {
            nama,
            no_kad_pengenalan: noKad,
            bahan_digital: 0,
            bahan_bukan_buku: 0,
            fiksyen: 0,
            bukan_fiksyen: 0,
            bahasa_melayu: 0,
            bahasa_inggeris: 0,
            lain_lain_bahasa: 0,
            ains_sepanjang_tahun: 0,
            jumlah_bacaan: 0,
          });
          if (!byName.has(nama.toLowerCase())) {
            byName.set(nama.toLowerCase(), key);
          }
        }
      });

    periodRecords
      .filter((row) => row.kelas === selectedClass)
      .forEach((row) => {
        const noKad = String(row.no_kad_pengenalan || "").trim();
        const nama = String(row.nama || "").trim();
        const existingKey =
          (noKad && byStudent.has(noKad) && noKad) ||
          (nama && byName.get(nama.toLowerCase())) ||
          "";
        const key = existingKey || noKad || `nm:${nama.toLowerCase()}`;

        if (!byStudent.has(key)) {
          byStudent.set(key, {
            nama: String(row.nama || "").trim(),
            no_kad_pengenalan: noKad,
            bahan_digital: 0,
            bahan_bukan_buku: 0,
            fiksyen: 0,
            bukan_fiksyen: 0,
            bahasa_melayu: 0,
            bahasa_inggeris: 0,
            lain_lain_bahasa: 0,
            ains_sepanjang_tahun: 0,
            jumlah_bacaan: 0,
          });
          if (nama && !byName.has(nama.toLowerCase())) {
            byName.set(nama.toLowerCase(), key);
          }
        }

        const slot = byStudent.get(key);
        slot.bahan_digital += Number(row.bahan_digital || 0);
        slot.bahan_bukan_buku += Number(row.bahan_bukan_buku || 0);
        slot.fiksyen += Number(row.fiksyen || 0);
        slot.bukan_fiksyen += Number(row.bukan_fiksyen || 0);
        slot.bahasa_melayu += Number(row.bahasa_melayu || 0);
        slot.bahasa_inggeris += Number(row.bahasa_inggeris || 0);
        slot.lain_lain_bahasa += Number(row.lain_lain_bahasa || 0);
        slot.jumlah_bacaan += computeJumlahBacaan(row);
      });

    (Array.isArray(yearlyRecords) ? yearlyRecords : [])
      .filter((row) => row.kelas === selectedClass)
      .forEach((row) => {
        const noKad = String(row.no_kad_pengenalan || "").trim();
        const nama = String(row.nama || "").trim();
        const existingKey =
          (noKad && byStudent.has(noKad) && noKad) ||
          (nama && byName.get(nama.toLowerCase())) ||
          "";
        const key = existingKey || noKad || `nm:${nama.toLowerCase()}`;
        const slot = byStudent.get(key);
        if (!slot) {
          return;
        }
        slot.ains_sepanjang_tahun = Math.max(
          slot.ains_sepanjang_tahun,
          Math.max(0, Math.trunc(Number(row.ains || 0)))
        );
      });

    const rows = [...byStudent.values()].sort((a, b) => a.nama.localeCompare(b.nama, "ms"));
    if (!rows.length) {
      el.summaryTbody.innerHTML = '<tr><td colspan="12" class="empty">Tiada data ringkasan.</td></tr>';
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
      ains_sepanjang_tahun: 0,
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
      total.ains_sepanjang_tahun += row.ains_sepanjang_tahun;
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
        <td>${row.ains_sepanjang_tahun}</td>
        <td>${row.jumlah_bacaan}</td>
        <td>${renderBintang(row.jumlah_bacaan)}</td>
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
        <td><strong>${total.ains_sepanjang_tahun}</strong></td>
        <td><strong>${total.jumlah_bacaan}</strong></td>
        <td><strong>${renderBintang(total.jumlah_bacaan)}</strong></td>
      </tr>`;
  }

  function getBintangCount(jumlahBacaan) {
    const jumlah = Math.max(0, Math.trunc(Number(jumlahBacaan) || 0));
    if (jumlah >= 120 && jumlah <= 239) {
      return 1;
    }
    if (jumlah >= 240 && jumlah <= 359) {
      return 2;
    }
    if (jumlah >= 360 && jumlah <= 479) {
      return 3;
    }
    if (jumlah >= 480 && jumlah <= 599) {
      return 4;
    }
    if (jumlah >= 600) {
      return 5;
    }
    return 0;
  }

  function renderBintang(jumlahBacaan) {
    const count = getBintangCount(jumlahBacaan);
    return count > 0 ? "⭐".repeat(count) : "-";
  }

  function renderStarDistribution(periodRecords, selectedClass, selectedTingkatan) {
    if (!el.starDistWrap) {
      return;
    }

    const students = buildStudentTotalsForStarChart(periodRecords, selectedClass, selectedTingkatan);
    const totalStudents = students.length;
    if (!totalStudents) {
      el.starDistWrap.innerHTML = '<p class="empty">Tiada murid untuk kiraan bintang.</p>';
      return;
    }

    const buckets = [1, 2, 3, 4, 5].map((star) => ({
      star,
      count: 0,
      percent: 0,
    }));
    students.forEach((row) => {
      const star = getBintangCount(row.jumlah_bacaan);
      if (star >= 1 && star <= 5) {
        buckets[star - 1].count += 1;
      }
    });
    buckets.forEach((row) => {
      row.percent = totalStudents ? (row.count / totalStudents) * 100 : 0;
    });

    const maxCount = Math.max(...buckets.map((row) => row.count), 1);
    const width = 520;
    const height = 280;
    const left = 52;
    const right = 20;
    const top = 18;
    const bottom = 58;
    const plotWidth = width - left - right;
    const plotHeight = height - top - bottom;
    const barColors = ["#16a34a", "#0ea5e9", "#f59e0b", "#f97316", "#8b5cf6"];
    const barWidth = Math.floor((plotWidth - 50) / buckets.length);
    const gap = Math.max(
      8,
      Math.floor((plotWidth - barWidth * buckets.length) / (buckets.length - 1))
    );

    const bars = buckets
      .map((row, index) => {
        const h = maxCount ? (row.count / maxCount) * plotHeight : 0;
        const x = left + index * (barWidth + gap);
        const y = top + (plotHeight - h);
        const labelY = h > 0 ? y - 6 : top + plotHeight - 6;
        return `
          <rect x="${x}" y="${y}" width="${barWidth}" height="${h}" fill="${barColors[index]}" rx="4" ry="4"></rect>
          <text x="${x + barWidth / 2}" y="${Math.max(14, labelY)}" text-anchor="middle" class="axis-label">${row.count}</text>
          <text x="${x + barWidth / 2}" y="${top + plotHeight + 16}" text-anchor="middle" class="axis-label">${row.percent.toFixed(1)}%</text>
          <text x="${x + barWidth / 2}" y="${top + plotHeight + 34}" text-anchor="middle" class="axis-label">${row.star}⭐</text>
        `;
      })
      .join("");

    el.starDistWrap.innerHTML = `
      <p class="chart-note">Jumlah murid diambil kira: <strong>${totalStudents}</strong> (termasuk 0 bintang).</p>
      <div class="bar-scroll">
        <svg viewBox="0 0 ${width} ${height}" class="bar-svg" aria-label="Carta taburan bintang 1 hingga 5">
          <line x1="${left}" y1="${top + plotHeight}" x2="${width - right}" y2="${top + plotHeight}" stroke="#9bb7bf" />
          ${bars}
        </svg>
      </div>
    `;
  }

  function buildStudentTotalsForStarChart(periodRecords, selectedClass, selectedTingkatan) {
    const byStudent = new Map();
    const isAllClasses = selectedClass === "__all__";
    const isAllTingkatan = selectedTingkatan === "__all__";

    state.students.forEach((row) => {
      const kelas = String(row.kelas || "").trim();
      const nama = String(row.nama || "").trim();
      if (!kelas || !nama) {
        return;
      }
      if (!isAllTingkatan && classToTingkatan(kelas) !== selectedTingkatan) {
        return;
      }
      if (!isAllClasses && kelas !== selectedClass) {
        return;
      }
      const noKad = normalizeKeyText(row.no_kad_pengenalan);
      const key = noKad ? `ic:${noKad}` : `nm:${normalizeKeyText(nama)}|k:${kelas.toLowerCase()}`;
      if (!byStudent.has(key)) {
        byStudent.set(key, {
          jumlah_bacaan: 0,
        });
      }
    });

    (Array.isArray(periodRecords) ? periodRecords : []).forEach((row) => {
      const kelas = String(row.kelas || "").trim();
      if (!kelas) {
        return;
      }
      if (!isAllTingkatan && classToTingkatan(kelas) !== selectedTingkatan) {
        return;
      }
      if (!isAllClasses && kelas !== selectedClass) {
        return;
      }

      const noKad = normalizeKeyText(row.no_kad_pengenalan);
      const nama = normalizeKeyText(row.nama);
      const icKey = noKad ? `ic:${noKad}` : "";
      const nameKey = nama ? `nm:${nama}|k:${kelas.toLowerCase()}` : "";
      let key = "";
      if (icKey && byStudent.has(icKey)) {
        key = icKey;
      } else if (nameKey && byStudent.has(nameKey)) {
        key = nameKey;
      } else if (icKey) {
        key = icKey;
      } else if (nameKey) {
        key = nameKey;
      }
      if (!key) {
        return;
      }
      if (!byStudent.has(key)) {
        byStudent.set(key, { jumlah_bacaan: 0 });
      }
      const slot = byStudent.get(key);
      slot.jumlah_bacaan += Math.max(0, Math.trunc(computeJumlahBacaan(row)));
    });

    return [...byStudent.values()];
  }

  function computeJumlahBacaan(row) {
    const withoutAins =
      Number(row.bahan_digital || 0) +
      Number(row.bahan_bukan_buku || 0) +
      Number(row.fiksyen || 0) +
      Number(row.bukan_fiksyen || 0);
    if (!state.includeAinsInJumlah) {
      return withoutAins;
    }
    return withoutAins + Number(row.ains || 0);
  }

  function renderPie(records) {
    const totals = MATERIAL_DONUT_KEYS.map(([key, label], index) => ({
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
    const positiveItems = totals.filter((item) => item.value > 0);

    if (positiveItems.length === 1) {
      const only = positiveItems[0];
      const legendSingle = totals
        .map((item) => {
          const percent = grandTotal ? Math.round((item.value / grandTotal) * 100) : 0;
          return `<li><i style="background:${item.color}"></i>${item.label}: ${item.value} (${percent}%)</li>`;
        })
        .join("");

      el.pieWrap.innerHTML = `
        <div class="chart-layout">
          <svg viewBox="0 0 260 260" class="pie-svg" aria-label="Carta pai jenis bahan">
            <circle cx="${cx}" cy="${cy}" r="${radius}" fill="${only.color}"></circle>
            <circle cx="${cx}" cy="${cy}" r="38" fill="#fdfefe"></circle>
          </svg>
          <ul class="chart-legend">${legendSingle}</ul>
        </div>
      `;
      return;
    }

    const slices = positiveItems
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
          <circle cx="${cx}" cy="${cy}" r="38" fill="#fdfefe"></circle>
        </svg>
        <ul class="chart-legend">${legend}</ul>
      </div>
    `;
  }

  function renderLanguageBars(records) {
    const rows = LANGUAGE_KEYS.map(([key, label], index) => ({
      label,
      value: records.reduce((sum, row) => sum + Number(row[key] || 0), 0),
      color: BAR_COLORS[index % BAR_COLORS.length],
    }));
    const grandTotal = rows.reduce((sum, row) => sum + row.value, 0);
    if (!grandTotal) {
      el.languageWrap.innerHTML = '<p class="empty">Semua nilai bahasa adalah 0.</p>';
      return;
    }

    let startAngle = -Math.PI / 2;
    const cx = 130;
    const cy = 130;
    const radius = 90;
    const positiveItems = rows.filter((item) => item.value > 0);

    if (positiveItems.length === 1) {
      const only = positiveItems[0];
      const legendSingle = rows
        .map((item) => {
          const percent = grandTotal ? Math.round((item.value / grandTotal) * 100) : 0;
          return `<li><i style="background:${item.color}"></i>${item.label}: ${item.value} (${percent}%)</li>`;
        })
        .join("");

      el.languageWrap.innerHTML = `
        <div class="chart-layout">
          <svg viewBox="0 0 260 260" class="pie-svg" aria-label="Carta donut pilihan bahasa">
            <circle cx="${cx}" cy="${cy}" r="${radius}" fill="${only.color}"></circle>
            <circle cx="${cx}" cy="${cy}" r="38" fill="#fdfefe"></circle>
          </svg>
          <ul class="chart-legend">${legendSingle}</ul>
        </div>
      `;
      return;
    }

    const slices = positiveItems
      .map((row, i) => {
        const angle = (row.value / grandTotal) * Math.PI * 2;
        const endAngle = startAngle + angle;
        const x1 = cx + radius * Math.cos(startAngle);
        const y1 = cy + radius * Math.sin(startAngle);
        const x2 = cx + radius * Math.cos(endAngle);
        const y2 = cy + radius * Math.sin(endAngle);
        const largeArc = angle > Math.PI ? 1 : 0;
        const path = `M ${cx} ${cy} L ${x1} ${y1} A ${radius} ${radius} 0 ${largeArc} 1 ${x2} ${y2} Z`;
        startAngle = endAngle;
        return `
          <path d="${path}" fill="${row.color}"></path>
        `;
      })
      .join("");

    const legend = rows
      .map((item) => {
        const percent = grandTotal ? Math.round((item.value / grandTotal) * 100) : 0;
        return `<li><i style="background:${item.color}"></i>${item.label}: ${item.value} (${percent}%)</li>`;
      })
      .join("");

    el.languageWrap.innerHTML = `
      <div class="chart-layout">
        <svg viewBox="0 0 260 260" class="pie-svg" aria-label="Carta donut pilihan bahasa">
          ${slices}
          <circle cx="${cx}" cy="${cy}" r="38" fill="#fdfefe"></circle>
        </svg>
        <ul class="chart-legend">${legend}</ul>
      </div>
    `;
  }

  function renderTingkatanBars(records) {
    const totals = new Map(TINGKATAN_ORDER.map((code) => [code, 0]));
    records.forEach((row) => {
      const tingkatan = classToTingkatan(row.kelas);
      if (!totals.has(tingkatan)) {
        return;
      }
      totals.set(tingkatan, totals.get(tingkatan) + computeJumlahBacaan(row));
    });
    const rows = TINGKATAN_ORDER.map((code) => ({
      code,
      label: code === "PER" ? "PER" : `T${code}`,
      value: totals.get(code) || 0,
    }));
    const maxValue = Math.max(...rows.map((r) => r.value), 0);
    if (!maxValue) {
      el.tingkatanWrap.innerHTML = '<p class="empty">Tiada jumlah bacaan ikut tingkatan.</p>';
      return;
    }

    const width = 520;
    const height = 280;
    const left = 52;
    const right = 20;
    const top = 18;
    const bottom = 42;
    const plotWidth = width - left - right;
    const plotHeight = height - top - bottom;
    const barWidth = Math.floor((plotWidth - 30) / rows.length);
    const gap = Math.max(6, Math.floor((plotWidth - barWidth * rows.length) / (rows.length - 1)));

    const bars = rows
      .map((row, i) => {
        const h = maxValue ? (row.value / maxValue) * plotHeight : 0;
        const x = left + i * (barWidth + gap);
        const y = top + (plotHeight - h);
        return `
          <rect x="${x}" y="${y}" width="${barWidth}" height="${h}" fill="#8d28b8" rx="4" ry="4"></rect>
          <text x="${x + barWidth / 2}" y="${Math.max(14, y - 4)}" text-anchor="middle" class="axis-label">${row.value}</text>
          <text x="${x + barWidth / 2}" y="${top + plotHeight + 16}" text-anchor="middle" class="axis-label">${row.label}</text>
        `;
      })
      .join("");

    el.tingkatanWrap.innerHTML = `
      <div class="bar-scroll">
        <svg viewBox="0 0 ${width} ${height}" class="bar-svg" aria-label="Carta jumlah bacaan mengikut tingkatan">
          <line x1="${left}" y1="${top + plotHeight}" x2="${width - right}" y2="${top + plotHeight}" stroke="#9bb7bf" />
          ${bars}
        </svg>
      </div>
    `;
  }

  function renderBar(records) {
    const totalsByStudent = new Map();
    records.forEach((row) => {
      const key = String(row.no_kad_pengenalan || row.nama || "").trim();
      const nama = String(row.nama || "").trim();
      const kelas = String(row.kelas || "").trim();
      const jumlah = computeJumlahBacaan(row);

      if (!totalsByStudent.has(key)) {
        totalsByStudent.set(key, { nama, kelas, jumlah: 0 });
      }
      const current = totalsByStudent.get(key);
      if (!current.kelas && kelas) {
        current.kelas = kelas;
      }
      current.jumlah += Math.max(0, Math.trunc(jumlah || 0));
    });

    const rows = [...totalsByStudent.values()]
      .map((row) => ({
        nama: String(row.nama || "").trim(),
        kelas: String(row.kelas || "").trim(),
        jumlah: Math.max(0, Math.trunc(Number(row.jumlah) || 0)),
      }))
      .sort((a, b) => b.jumlah - a.jumlah)
      .slice(0, 10);

    if (!rows.length) {
      el.barWrap.innerHTML = '<p class="empty">Tiada data jumlah bacaan.</p>';
      return;
    }

    const maxValue = Math.max(...rows.map((r) => r.jumlah), 1);
    const maxNameLen = Math.max(...rows.map((r) => r.nama.length), 1);
    const rowHeight = 28;
    const left = Math.min(320, Math.max(160, 7 * maxNameLen));
    const right = 48;
    const top = 18;
    const bottom = 24;
    const width = 680;
    const plotWidth = width - left - right;
    const chartHeight = top + bottom + rows.length * rowHeight;

    const bars = rows
      .map((row, i) => {
        const y = top + i * rowHeight;
        const w = maxValue ? (row.jumlah / maxValue) * plotWidth : 0;
        const fullName = escapeHtml(row.kelas ? `${row.nama} (${row.kelas})` : row.nama);
        const barColor = BAR_COLORS[i % BAR_COLORS.length];
        return `
          <text x="${left - 8}" y="${y + 18}" text-anchor="end" class="axis-label axis-label-name">${fullName}</text>
          <rect x="${left}" y="${y + 4}" width="${w}" height="18" fill="${barColor}" rx="4" ry="4"></rect>
          <text x="${left + w + 6}" y="${y + 18}" text-anchor="start" class="axis-label">${row.jumlah}</text>
        `;
      })
      .join("");

    el.barWrap.innerHTML = `
      <div class="bar-scroll">
        <svg viewBox="0 0 ${width} ${chartHeight}" class="bar-svg" aria-label="Carta bar jumlah bacaan">
          <line x1="${left}" y1="${top - 2}" x2="${left}" y2="${chartHeight - bottom + 2}" stroke="#c2d2d8" />
          ${bars}
        </svg>
      </div>
    `;
  }

  function renderClassHeatmap(records) {
    const classes = [...new Set(records.map((row) => String(row.kelas || "").trim()).filter(Boolean))].sort((a, b) =>
      a.localeCompare(b, "ms")
    );
    if (!classes.length) {
      el.heatmapWrap.innerHTML = '<p class="empty">Tiada data kelas.</p>';
      return;
    }

    const rowCodes = TINGKATAN_ORDER.filter((code) =>
      records.some((row) => classToTingkatan(row.kelas) === code)
    );
    const matrix = new Map();
    rowCodes.forEach((code) => matrix.set(code, new Map(classes.map((kelas) => [kelas, 0]))));

    records.forEach((row) => {
      const code = classToTingkatan(row.kelas);
      const kelas = String(row.kelas || "").trim();
      if (!matrix.has(code) || !kelas) {
        return;
      }
      const rowMap = matrix.get(code);
      rowMap.set(kelas, (rowMap.get(kelas) || 0) + computeJumlahBacaan(row));
    });

    const maxValue = Math.max(
      ...rowCodes.flatMap((code) => classes.map((kelas) => matrix.get(code).get(kelas) || 0)),
      0
    );
    const classTotals = new Map(classes.map((kelas) => [kelas, 0]));

    const bodyRows = rowCodes
      .map((code) => {
        const rowMap = matrix.get(code);
        let rowTotal = 0;
        const tds = classes
          .map((kelas) => {
            const value = rowMap.get(kelas) || 0;
            rowTotal += value;
            classTotals.set(kelas, (classTotals.get(kelas) || 0) + value);
            const shade = maxValue ? 0.18 + (value / maxValue) * 0.52 : 0;
            const bg = value ? ` style="background: rgba(83, 151, 238, ${shade.toFixed(3)});"` : "";
            return `<td${bg}>${value}</td>`;
          })
          .join("");
        const label = code === "PER" ? "PER" : `T${code}`;
        return `<tr><td>${label}</td>${tds}<td><strong>${rowTotal}</strong></td></tr>`;
      })
      .join("");

    const allTotal = [...classTotals.values()].reduce((sum, value) => sum + value, 0);
    const footerCells = classes
      .map((kelas) => `<td><strong>${classTotals.get(kelas) || 0}</strong></td>`)
      .join("");

    el.heatmapWrap.innerHTML = `
      <table class="heatmap-table">
        <thead>
          <tr>
            <th>Tingkatan</th>
            ${classes.map((kelas) => `<th>${escapeHtml(kelas)}</th>`).join("")}
            <th>Grand Total</th>
          </tr>
        </thead>
        <tbody>
          ${bodyRows}
          <tr>
            <td><strong>JUMLAH</strong></td>
            ${footerCells}
            <td><strong>${allTotal}</strong></td>
          </tr>
        </tbody>
      </table>
    `;
  }

  function normalizeKeyText(value) {
    return String(value || "").trim().toLowerCase();
  }

  function buildAllowedStudentKeysByClass(students) {
    const map = new Map();
    (Array.isArray(students) ? students : []).forEach((row) => {
      const kelas = String(row.kelas || "").trim();
      const nama = normalizeKeyText(row.nama);
      const noKad = normalizeKeyText(row.no_kad_pengenalan);
      if (!kelas || !nama) {
        return;
      }
      if (!map.has(kelas)) {
        map.set(kelas, new Set());
      }
      const keys = map.get(kelas);
      keys.add(`nm:${nama}`);
      if (noKad) {
        keys.add(`ic:${noKad}`);
      }
    });
    return map;
  }

  function buildClassByNoKad(students) {
    const map = new Map();
    (Array.isArray(students) ? students : []).forEach((row) => {
      const noKad = normalizeKeyText(row.no_kad_pengenalan);
      const kelas = String(row.kelas || "").trim();
      if (noKad && kelas) {
        map.set(noKad, kelas);
      }
    });
    return map;
  }

  function buildNameByNoKad(students) {
    const map = new Map();
    (Array.isArray(students) ? students : []).forEach((row) => {
      const noKad = normalizeKeyText(row.no_kad_pengenalan);
      const nama = normalizeKeyText(row.nama);
      if (noKad && nama) {
        map.set(noKad, nama);
      }
    });
    return map;
  }

  function buildClassesByName(students) {
    const map = new Map();
    (Array.isArray(students) ? students : []).forEach((row) => {
      const nama = normalizeKeyText(row.nama);
      const kelas = String(row.kelas || "").trim();
      if (!nama || !kelas) {
        return;
      }
      if (!map.has(nama)) {
        map.set(nama, new Set());
      }
      map.get(nama).add(kelas);
    });
    return map;
  }

  function resolveRecordClassByMaster(row) {
    const noKad = normalizeKeyText(row?.no_kad_pengenalan);
    const nama = normalizeKeyText(row?.nama);
    const classes = state.classesByName.get(nama);

    if (noKad && state.classByNoKad.has(noKad)) {
      const kelas = state.classByNoKad.get(noKad);
      const ownerName = state.nameByNoKad.get(noKad);
      if (ownerName && ownerName === nama) {
        if (kelas && kelas !== row.kelas) {
          return { ...row, kelas };
        }
        return row;
      }
      // IC/name mismatch: prefer unique name-based class to avoid swapped display.
      if (classes && classes.size === 1) {
        const [kelasByName] = classes;
        if (kelasByName && kelasByName !== row.kelas) {
          return { ...row, kelas: kelasByName };
        }
        return row;
      }
      if (kelas && kelas !== row.kelas) {
        return { ...row, kelas };
      }
      return row;
    }

    if (classes && classes.size === 1) {
      const [kelas] = classes;
      if (kelas && kelas !== row.kelas) {
        return { ...row, kelas };
      }
    }

    return row;
  }

  function filterRecordsByMasterList(records) {
    return (Array.isArray(records) ? records : []).filter((row) => {
      const kelas = String(row.kelas || "").trim();
      const keys = state.allowedStudentKeysByClass.get(kelas);
      if (!keys || !keys.size) {
        return false;
      }
      const nama = normalizeKeyText(row.nama);
      const noKad = normalizeKeyText(row.no_kad_pengenalan);
      if (noKad && keys.has(`ic:${noKad}`)) {
        return true;
      }
      return nama ? keys.has(`nm:${nama}`) : false;
    });
  }

  function classToTingkatan(kelasValue) {
    const kelas = String(kelasValue || "").trim().toUpperCase();
    if (!kelas) {
      return "";
    }
    if (kelas.startsWith("PER")) {
      return "PER";
    }
    const match = kelas.match(/^([1-5])/);
    return match ? match[1] : "";
  }

  function filterClassesByTingkatan(classes, selectedTingkatan) {
    if (selectedTingkatan === "__all__") {
      return [...classes];
    }
    return classes.filter((kelas) => classToTingkatan(kelas) === selectedTingkatan);
  }

  function getMostRecentMonthText(records, year, selectedMonth) {
    if (selectedMonth && selectedMonth !== "__year__") {
      return selectedMonth;
    }
    const currentYear = String(new Date().getFullYear());
    if (String(year) === currentYear) {
      return MONTHS[new Date().getMonth()];
    }
    const monthIndexes = records
      .filter((row) => String(row.tahun || "") === String(year))
      .map((row) => MONTHS.indexOf(String(row.bulan || "")))
      .filter((index) => index >= 0);
    if (!monthIndexes.length) {
      return MONTHS[new Date().getMonth()];
    }
    return MONTHS[Math.max(...monthIndexes)];
  }

  function getSehinggaLabel(records, year) {
    return `Sehingga ${getMostRecentMonthText(records, year, "__year__")}`;
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
