(function () {
  "use strict";

  const CSV_FILE = "senarai_nama_murid_10FEB2026 - ALL.csv";
  const MONTHS = [
    "Januari",
    "Februari",
    "Mac",
    "April",
    "Mei",
    "Jun",
    "Julai",
    "Ogos",
    "September",
    "Oktober",
    "November",
    "Disember",
  ];

  const inputColumns = [
    "bahan_digital",
    "bahan_bukan_buku",
    "fiksyen",
    "bukan_fiksyen",
    "bahasa_melayu",
    "bahasa_inggeris",
    "lain_lain_bahasa",
  ];
  const NAMELIST_OVERRIDE_KEY = "nilam_students_override_v1";
  const TEACHER_NAMES_KEY = "nilam_teacher_names_v1";
  const GURU_TYPES = ["Nilam", "BM", "BI"];


  const CLOUD_ONLY_MODE = true;

  function localLength() {
    if (CLOUD_ONLY_MODE) {
      return 0;
    }
    return window.localStorage.length;
  }

  function localKey(index) {
    if (CLOUD_ONLY_MODE) {
      return null;
    }
    return window.localStorage.key(index);
  }

  function localGetItem(key) {
    if (CLOUD_ONLY_MODE) {
      return null;
    }
    return window.localStorage.getItem(key);
  }

  function localSetItem(key, value) {
    if (CLOUD_ONLY_MODE) {
      return;
    }
    window.localStorage.setItem(key, value);
  }

  function localRemoveItem(key) {
    if (CLOUD_ONLY_MODE) {
      return;
    }
    window.localStorage.removeItem(key);
  }
  const SYSTEM_TEACHER_NAMES = new Set(["ADMIN_IMPORT", "ADMIN_OVERRIDE", "TIDAK DIKETAHUI"]);

  const state = {
    rawStudents: [],
    selectedYear: String(new Date().getFullYear()),
    selectedClass: "",
    selectedMonth: "",
    selectedDate: "",
    selectedTeacherName: "",
    selectedGuruType: "Nilam",
    includeAinsInJumlah: true,
    teacherNames: [],
    prefillRequestSeq: 0,
  };

  const el = {
    namaPengisi: document.getElementById("namaPengisi"),
    namaGuruCadanganList: document.getElementById("namaGuruCadanganList"),
    guruJenis: document.getElementById("guruJenis"),
    tarikh: document.getElementById("tarikh"),
    tahun: document.getElementById("tahun"),
    kelas: document.getElementById("kelas"),
    includeAinsCheckbox: document.getElementById("includeAinsJumlah"),
    tbody: document.querySelector("#studentTable tbody"),
    status: document.getElementById("status"),
    saveAll: document.getElementById("saveAll"),
    saveAllBottom: document.getElementById("saveAllBottom"),
  };

  init();

  // Load students from localStorage override / bundled data immediately (no network).
  function getLocalStudentsFast() {
    if (CLOUD_ONLY_MODE) {
      return [];
    }
    const overrideRaw = localGetItem(NAMELIST_OVERRIDE_KEY);
    if (overrideRaw) {
      try {
        const parsed = JSON.parse(overrideRaw);
        if (Array.isArray(parsed) && parsed.length) {
          return normalizeStudents(parsed);
        }
      } catch (_) {}
    }
    const bundled = Array.isArray(window.NILAM_STUDENTS) ? window.NILAM_STUDENTS : [];
    return normalizeStudents(bundled);
  }

  async function init() {
    initYearField();
    initGuruDropdown();
    initTeacherDropdown();
    initDateField();
    fitEntryControls();
    bindEvents();
    refreshTeacherNamesFromSupabase().catch((error) => {
      console.error("Gagal memuatkan nama guru dari Supabase", error);
    });

    // Phase 1 — show local data instantly (no network wait).
    const localStudents = getLocalStudentsFast();
    if (localStudents.length) {
      hydrateStudents(localStudents);
      setStatus(`${localStudents.length} murid dipaparkan. Memuatkan dari cloud...`);
    }

    // Phase 2 — fetch Supabase in background and refresh silently.
    try {
      const students = await loadStudents();
      const prevClass = state.selectedClass;
      if (prevClass) {
        // User already selected a class — update student list without resetting the table.
        state.rawStudents = [...students].sort((a, b) => {
          const byKelas = a.kelas.localeCompare(b.kelas, "ms");
          return byKelas !== 0 ? byKelas : a.nama.localeCompare(b.nama, "ms");
        });
        el.kelas.innerHTML = '<option value="">Pilih kelas</option>';
        initClassDropdown(state.rawStudents);
        if (students.some((s) => s.kelas === prevClass)) {
          el.kelas.value = prevClass;
          state.selectedClass = prevClass;
        } else {
          state.selectedClass = "";
        }
      } else {
        hydrateStudents(students);
      }

      const syncResult = await autoSyncLocalRecordsToSupabase();
      if (syncResult.mode === "cloud_error") {
        setStatus(
          `Data murid berjaya dimuatkan: ${students.length} murid. Sync cloud gagal: ${String(
            syncResult.errorMessage || "semak config.js / Supabase"
          ).slice(0, 220)}`,
          true
        );
        return;
      }
      if (syncResult.syncedCount > 0) {
        setStatus(
          `Data murid berjaya dimuatkan: ${students.length} murid. ${syncResult.syncedCount} rekod tertunda telah disegerakkan ke cloud.`
        );
        return;
      }
      setStatus(`Data murid berjaya dimuatkan: ${students.length} murid.`);
    } catch (error) {
      setStatus(
        `Gagal memuatkan data murid dari cloud. ${String(error?.message || "Sila semak config.js dan sambungan Internet.")}`,
        true
      );
      console.error(error);
    }
  }

  function bindEvents() {
    if (el.namaPengisi) {
      el.namaPengisi.addEventListener("input", () => {
        state.selectedTeacherName = String(el.namaPengisi.value || "").trim();
        renderTeacherSuggestions(state.selectedTeacherName);
        fitTeacherNameInputWidth();
      });
      el.namaPengisi.addEventListener("change", async () => {
        state.selectedTeacherName = String(el.namaPengisi.value || "").trim();
        renderTeacherSuggestions(state.selectedTeacherName);
        fitTeacherNameInputWidth();
        if (state.selectedTeacherName) {
          rememberTeacherName(state.selectedTeacherName);
        }
        if (state.selectedClass) {
          await renderTableAndPrefill();
        }
      });
      el.namaPengisi.addEventListener("focus", () => {
        renderTeacherSuggestions(state.selectedTeacherName);
      });
      el.namaPengisi.addEventListener("blur", () => {
        window.setTimeout(() => {
          hideTeacherSuggestions();
        }, 120);
      });
    }

    if (el.namaGuruCadanganList) {
      el.namaGuruCadanganList.addEventListener("mousedown", (event) => {
        const target = event.target instanceof HTMLElement ? event.target.closest(".teacher-suggestion-item") : null;
        if (!target) {
          return;
        }
        event.preventDefault();
        chooseTeacherSuggestion(String(target.dataset.value || target.textContent || "").trim());
      });
    }

    document.addEventListener("click", (event) => {
      if (!event.target || !(event.target instanceof Element)) {
        hideTeacherSuggestions();
        return;
      }
      if (!event.target.closest(".teacher-name-field")) {
        hideTeacherSuggestions();
      }
    });

    if (el.guruJenis) {
      el.guruJenis.addEventListener("change", async () => {
        const chosen = String(el.guruJenis.value || "Nilam").trim();
        state.selectedGuruType = GURU_TYPES.includes(chosen) ? chosen : "Nilam";
        fitSelectWidth(el.guruJenis, 12, 14);
        applyGuruModeToVisibleRows();
        if (state.selectedClass) {
          await renderTableAndPrefill();
        }
      });
    }

    if (el.tarikh) {
      el.tarikh.addEventListener("change", async () => {
        const normalized = normalizeDateInput(el.tarikh.value);
        if (!normalized) {
          return;
        }
        fitTextInputWidth(el.tarikh, 12, 14);
        state.selectedDate = normalized;
        const pickedDate = new Date(`${normalized}T00:00:00`);
        state.selectedYear = String(pickedDate.getFullYear());
        state.selectedMonth = MONTHS[pickedDate.getMonth()];
        if (el.tahun) {
          el.tahun.value = state.selectedYear;
        }
        if (state.selectedClass) {
          await renderTableAndPrefill();
        }
      });
    }

    el.kelas.addEventListener("change", async () => {
      state.selectedClass = el.kelas.value;
      fitSelectWidth(el.kelas, 12, 20);
      await renderTableAndPrefill();
    });

    if (el.includeAinsCheckbox) {
      el.includeAinsCheckbox.checked = true;
      el.includeAinsCheckbox.addEventListener("change", () => {
        state.includeAinsInJumlah = Boolean(el.includeAinsCheckbox.checked);
        const config = window.NILAM_CONFIG || {};
        loadAndApplyTotals(state.selectedYear, config).catch(() => {});
      });
    }

    if (el.saveAll) {
      el.saveAll.addEventListener("click", saveAllRecords);
    }
    if (el.saveAllBottom) {
      el.saveAllBottom.addEventListener("click", saveAllRecords);
    }
  }

  function initYearField() {
    if (!el.tahun) {
      return;
    }
    el.tahun.value = state.selectedYear;
  }

  function initGuruDropdown() {
    if (!el.guruJenis) {
      return;
    }
    el.guruJenis.innerHTML = "";
    GURU_TYPES.forEach((guruType) => {
      const option = document.createElement("option");
      option.value = guruType;
      option.textContent = guruType;
      el.guruJenis.appendChild(option);
    });
    el.guruJenis.value = state.selectedGuruType;
    fitSelectWidth(el.guruJenis, 12, 14);
  }

  function initTeacherDropdown() {
    if (!el.namaPengisi) {
      return;
    }
    state.teacherNames = loadTeacherNames();
    saveTeacherNames();
    renderTeacherSuggestions(state.selectedTeacherName);
    el.namaPengisi.value = state.selectedTeacherName;
    fitTeacherNameInputWidth();
  }

  function initDateField() {
    const today = new Date();
    const isoDate = toIsoDate(today);
    state.selectedDate = isoDate;
    state.selectedYear = String(today.getFullYear());
    state.selectedMonth = MONTHS[today.getMonth()];
    if (el.tarikh) {
      el.tarikh.value = isoDate;
      fitTextInputWidth(el.tarikh, 12, 14);
    }
    if (el.tahun) {
      el.tahun.value = state.selectedYear;
      fitTextInputWidth(el.tahun, 8, 10);
    }
  }

  function initClassDropdown(students) {
    const classes = [...new Set(students.map((s) => s.kelas))].sort();
    classes.forEach((className) => {
      const option = document.createElement("option");
      option.value = className;
      option.textContent = className;
      el.kelas.appendChild(option);
    });
    fitSelectWidth(el.kelas, 12, 20);
  }

  function resetClassDropdown() {
    el.kelas.innerHTML = '<option value="">Pilih kelas</option>';
    state.selectedClass = "";
    fitSelectWidth(el.kelas, 12, 20);
  }

  function fitEntryControls() {
    fitTeacherNameInputWidth();
    fitSelectWidth(el.guruJenis, 12, 14);
    fitTextInputWidth(el.tarikh, 12, 14);
    fitTextInputWidth(el.tahun, 8, 10);
    fitSelectWidth(el.kelas, 12, 20);
  }

  function fitTeacherNameInputWidth() {
    if (!el.namaPengisi) {
      return;
    }
    const valueLen = String(el.namaPengisi.value || "").trim().length;
    const placeholderLen = String(el.namaPengisi.getAttribute("placeholder") || "").trim().length;
    const longestSuggestionLen = state.teacherNames.reduce((maxLen, name) => {
      const len = String(name || "").trim().length;
      return len > maxLen ? len : maxLen;
    }, 0);
    const targetLen = Math.max(valueLen, placeholderLen, longestSuggestionLen);
    const widthCh = Math.max(14, Math.min(42, targetLen + 2));
    el.namaPengisi.style.width = `${widthCh}ch`;
  }

  function fitTextInputWidth(inputEl, minCh, maxCh) {
    if (!inputEl) {
      return;
    }
    const raw = String(inputEl.value || "").trim();
    const placeholder = String(inputEl.getAttribute("placeholder") || "").trim();
    const source = raw || placeholder;
    const length = source.length ? source.length + 2 : minCh;
    const clamped = Math.max(minCh, Math.min(maxCh, length));
    inputEl.style.width = `${clamped}ch`;
  }

  function fitSelectWidth(selectEl, minCh, maxCh) {
    if (!selectEl) {
      return;
    }
    const option = selectEl.options && selectEl.selectedIndex >= 0
      ? selectEl.options[selectEl.selectedIndex]
      : null;
    const text = String((option && option.textContent) || selectEl.value || "").trim();
    const length = text.length ? text.length + 4 : minCh;
    const clamped = Math.max(minCh, Math.min(maxCh, length));
    selectEl.style.width = `${clamped}ch`;
  }

  function hydrateStudents(students) {
    state.rawStudents = [...students].sort((a, b) => {
      const byKelas = a.kelas.localeCompare(b.kelas, "ms");
      return byKelas !== 0 ? byKelas : a.nama.localeCompare(b.nama, "ms");
    });
    resetClassDropdown();
    initClassDropdown(state.rawStudents);
    el.tbody.innerHTML =
      '<tr><td colspan="14" class="empty">Pilih kelas untuk paparkan senarai murid.</td></tr>';
  }

  async function refreshStudentsPreserveSelectedClass() {
    const selectedBefore = state.selectedClass;
    if (!selectedBefore) {
      return;
    }
    try {
      const latestStudents = await loadStudents();
      state.rawStudents = latestStudents;
      resetClassDropdown();
      initClassDropdown(latestStudents);
      const hasClass = latestStudents.some((row) => row.kelas === selectedBefore);
      if (hasClass) {
        el.kelas.value = selectedBefore;
        state.selectedClass = selectedBefore;
      }
    } catch (error) {
      console.error("Gagal refresh senarai murid semasa pilih kelas", error);
    }
  }

  async function renderTableAndPrefill() {
    await refreshStudentsPreserveSelectedClass();
    renderTable();
    await prefillSavedValuesForSelection();
  }

  function renderTable() {
    if (!state.selectedClass) {
      el.tbody.innerHTML =
        '<tr><td colspan="14" class="empty">Pilih kelas untuk paparkan senarai murid.</td></tr>';
      return;
    }

    const students = state.rawStudents.filter((s) => s.kelas === state.selectedClass);
    if (!students.length) {
      el.tbody.innerHTML =
        '<tr><td colspan="14" class="empty">Tiada murid untuk kelas ini.</td></tr>';
      return;
    }

    const rowsHtml = students
      .map((student, index) => {
        const rowId = `${sanitizeId(state.selectedMonth)}_${sanitizeId(student.kelas)}_${sanitizeId(
          student.nama
        )}_${index + 1}`;
        return `
          <tr data-row-id="${rowId}" data-nama="${escapeHtml(student.nama)}" data-kelas="${escapeHtml(
          student.kelas
        )}" data-no-kad="${escapeHtml(student.no_kad_pengenalan || "")}" data-has-saved-session="0">
            <td>${index + 1}</td>
            <td><span class="cell-static cell-name">${escapeHtml(student.nama)}</span></td>
            <td><span class="cell-static">${escapeHtml(student.kelas)}</span></td>
            <td>${numericInput("bahan_digital")}</td>
            <td>${numericInput("bahan_bukan_buku")}</td>
            <td>${numericInput("fiksyen")}</td>
            <td>${numericInput("bukan_fiksyen")}</td>
            <td>${numericInput("bahasa_melayu")}</td>
            <td>${numericInput("bahasa_inggeris")}</td>
            <td>${numericInput("lain_lain_bahasa")}</td>
            <td><span class="cell-total" data-col="jumlah_aktiviti">0</span></td>
            <td><span class="cell-total" data-col="ains_sepanjang_tahun">—</span></td>
            <td><span class="cell-total" data-col="jumlah_tahun">—</span></td>
            <td><span class="cell-total" data-col="jumlah_all_time">—</span></td>
          </tr>
        `;
      })
      .join("");

    el.tbody.innerHTML = rowsHtml;
    attachRowInputHandlers();
    applyGuruModeToVisibleRows();
  }

  function numericInput(colName, readOnly = false) {
    const roAttr = readOnly ? " readonly" : "";
    return `<input type="number" min="0" max="999" step="1" value="0" data-col="${colName}"${roAttr}>`;
  }

  function attachRowInputHandlers() {
    const inputs = el.tbody.querySelectorAll('input[type="number"]');
    inputs.forEach((input) => {
      input.addEventListener("input", (event) => {
        if (event.target.readOnly) {
          return;
        }
        normalizeIntegerInput(event.target);
        updateJumlahAktiviti(event.target.closest("tr"));
      });
    });
  }

  function normalizeIntegerInput(input) {
    let value = input.value.replace(/[^\d]/g, "").slice(0, 3);
    if (value === "") {
      input.value = "0";
      return;
    }
    const number = Number(value);
    input.value = String(Math.max(0, Math.min(999, number)));
  }

  function updateJumlahAktiviti(row) {
    if (!row) {
      return;
    }
    applyGuruAutoLanguageForRow(row);
    const total = computeJumlahBacaanFromRecord({
      bahan_digital: getNumberFromCell(row, "bahan_digital"),
      bahan_bukan_buku: getNumberFromCell(row, "bahan_bukan_buku"),
      fiksyen: getNumberFromCell(row, "fiksyen"),
      bukan_fiksyen: getNumberFromCell(row, "bukan_fiksyen"),
      ains: 0,
    }, false);
    row.querySelector('[data-col="jumlah_aktiviti"]').textContent = String(total);
  }

  function computeAinsTotalFromRecord(row) {
    return clampNilamNumber(row?.ains);
  }

  function computeTotalsMap(records, includeAins) {
    const materialsByStudent = new Map();
    (Array.isArray(records) ? records : []).forEach((r) => {
      const noKad = String(r.no_kad_pengenalan || "").trim();
      if (!noKad) {
        return;
      }
      const materials = computeJumlahBacaanFromRecord(r, false);
      materialsByStudent.set(noKad, (materialsByStudent.get(noKad) || 0) + materials);
    });

    if (!includeAins) {
      return materialsByStudent;
    }

    const ainsAnnualByStudent = computeAnnualAinsTotalsMap(records);
    const merged = new Map(materialsByStudent);
    ainsAnnualByStudent.forEach((ains, noKad) => {
      merged.set(noKad, (merged.get(noKad) || 0) + ains);
    });
    return merged;
  }

  function computeAinsTotalsMap(records) {
    const map = new Map();
    (Array.isArray(records) ? records : []).forEach((r) => {
      const noKad = String(r.no_kad_pengenalan || "").trim();
      if (!noKad) {
        return;
      }
      const current = map.get(noKad) || 0;
      const next = computeAinsTotalFromRecord(r);
      if (next > current) {
        map.set(noKad, next);
      } else if (!map.has(noKad)) {
        map.set(noKad, current);
      }
    });
    return map;
  }

  function computeAnnualAinsTotalsMap(records) {
    const maxByStudentYear = new Map();
    (Array.isArray(records) ? records : []).forEach((r) => {
      const noKad = String(r.no_kad_pengenalan || "").trim();
      const tahun = String(r.tahun || "").trim();
      if (!noKad) {
        return;
      }
      const key = `${noKad}|${tahun}`;
      const next = computeAinsTotalFromRecord(r);
      const current = maxByStudentYear.get(key) || 0;
      if (next > current) {
        maxByStudentYear.set(key, next);
      } else if (!maxByStudentYear.has(key)) {
        maxByStudentYear.set(key, current);
      }
    });

    const sumByStudent = new Map();
    maxByStudentYear.forEach((ains, key) => {
      const sepIndex = key.indexOf("|");
      const noKad = sepIndex >= 0 ? key.slice(0, sepIndex) : key;
      sumByStudent.set(noKad, (sumByStudent.get(noKad) || 0) + ains);
    });
    return sumByStudent;
  }

  function recalculateVisibleJumlahAktiviti() {
    const rows = [...el.tbody.querySelectorAll("tr[data-row-id]")];
    rows.forEach((row) => updateJumlahAktiviti(row));
  }

  function getNumberFromCell(row, colName) {
    const input = row.querySelector(`input[data-col="${colName}"]`);
    if (input) {
      return Number(input.value) || 0;
    }
    const span = row.querySelector(`span[data-col="${colName}"]`);
    return Number(span?.textContent || 0) || 0;
  }

  function setNumberToCell(row, colName, value) {
    const input = row.querySelector(`input[data-col="${colName}"]`);
    const number = Number(value);
    const safe = Number.isFinite(number) ? String(Math.max(0, Math.min(999, Math.trunc(number)))) : "0";
    if (input) {
      input.value = safe;
      return;
    }
    const span = row.querySelector(`span[data-col="${colName}"]`);
    if (!span) {
      return;
    }
    span.textContent = safe;
  }

  function collectCurrentRows() {
    // Build a name→IC lookup from the full student list for rows missing an IC.
    const icByName = new Map(
      state.rawStudents
        .filter((s) => s.no_kad_pengenalan)
        .map((s) => [s.nama.trim().toLowerCase(), s.no_kad_pengenalan.trim()])
    );

    const rows = [...el.tbody.querySelectorAll("tr[data-row-id]")];
    const records = [];
    let bil = 1;
    rows.forEach((row) => {
      const nama = (row.dataset.nama || "").trim();
      const kelas = (row.dataset.kelas || "").trim();
      let noKad = (row.dataset.noKad || "").trim();

      // If row has no IC, try to find one by matching name.
      if (!noKad) {
        noKad = icByName.get(nama.toLowerCase()) || "";
      }

      // Skip rows that still have no IC — cannot save without FK reference.
      if (!noKad) {
        return;
      }

      const bahanDigital = getNumberFromCell(row, "bahan_digital");
      const bahanBukanBuku = getNumberFromCell(row, "bahan_bukan_buku");
      const fiksyen = getNumberFromCell(row, "fiksyen");
      const bukanFiksyen = getNumberFromCell(row, "bukan_fiksyen");
      const languageValues = resolveLanguageValuesFromGuruType(
        state.selectedGuruType,
        {
          bahan_digital: bahanDigital,
          bahan_bukan_buku: bahanBukanBuku,
          fiksyen,
          bukan_fiksyen: bukanFiksyen,
        },
        {
          bahasa_melayu: getNumberFromCell(row, "bahasa_melayu"),
          bahasa_inggeris: getNumberFromCell(row, "bahasa_inggeris"),
          lain_lain_bahasa: getNumberFromCell(row, "lain_lain_bahasa"),
        }
      );
      const jumlahAktiviti = Number(
        row.querySelector('[data-col="jumlah_aktiviti"]').textContent || "0"
      );
      const record = {
        no_kad_pengenalan: noKad,
        tahun: state.selectedYear,
        bulan: state.selectedMonth,
        tarikh: state.selectedDate,
        nama_pengisi: state.selectedTeacherName,
        guru: state.selectedGuruType,
        bil: bil++,
        nama,
        kelas,
        bahan_digital: bahanDigital,
        bahan_bukan_buku: bahanBukanBuku,
        fiksyen,
        bukan_fiksyen: bukanFiksyen,
        ains: 0,
        bahasa_melayu: languageValues.bahasa_melayu,
        bahasa_inggeris: languageValues.bahasa_inggeris,
        lain_lain_bahasa: languageValues.lain_lain_bahasa,
        jumlah_aktiviti: Number.isFinite(jumlahAktiviti) ? jumlahAktiviti : 0,
        updated_at_client: new Date().toISOString(),
      };

      for (const col of inputColumns) {
        if (!Number.isInteger(record[col]) || record[col] < 0 || record[col] > 999) {
          throw new Error(`Nilai tidak sah untuk ${col} pada murid ${record.nama}`);
        }
      }

      const hasAnyInput = inputColumns.some((col) => record[col] > 0);
      const hadSavedSession = row.dataset.hasSavedSession === "1";
      if (!hasAnyInput && !hadSavedSession) {
        return;
      }

      records.push(record);
    });
    return records;
  }

  function validateNilamBalanceBeforeSave() {
    const rows = [...el.tbody.querySelectorAll("tr[data-row-id]")];
    const mismatches = [];

    rows.forEach((row) => {
      const bahanJumlah =
        getNumberFromCell(row, "bahan_digital") +
        getNumberFromCell(row, "bahan_bukan_buku") +
        getNumberFromCell(row, "fiksyen") +
        getNumberFromCell(row, "bukan_fiksyen");
      const bahasaJumlah =
        getNumberFromCell(row, "bahasa_melayu") +
        getNumberFromCell(row, "bahasa_inggeris") +
        getNumberFromCell(row, "lain_lain_bahasa");

      if (bahanJumlah === 0 && bahasaJumlah === 0) {
        return;
      }
      if (bahanJumlah !== bahasaJumlah) {
        mismatches.push({
          nama: String(row.dataset.nama || "").trim() || "Tanpa Nama",
          bahanJumlah,
          bahasaJumlah,
        });
      }
    });

    if (!mismatches.length) {
      return { ok: true, message: "" };
    }

    const preview = mismatches
      .slice(0, 8)
      .map(
        (item, index) =>
          `${index + 1}. ${item.nama} (Bahan: ${item.bahanJumlah}, Bahasa: ${item.bahasaJumlah})`
      )
      .join("\n");
    const moreCount = mismatches.length > 8 ? `\n...dan ${mismatches.length - 8} lagi.` : "";

    const message =
      "Peringatan mesra untuk Guru Nilam:\n" +
      "Jumlah Bahan Bacaan mesti sama dengan jumlah Bahasa bagi setiap murid.\n\n" +
      "Sila semak murid berikut:\n" +
      `${preview}${moreCount}\n\n` +
      "Formula semakan:\n" +
      "Bahan Digital + Bahan Bukan Buku + Fiksyen + Bukan Fiksyen\n" +
      "mesti sama dengan\n" +
      "Bahasa Melayu + Bahasa Inggeris + Lain-lain Bahasa.";

    return { ok: false, message };
  }

  async function saveAllRecords() {
    state.selectedTeacherName = String(el.namaPengisi?.value || state.selectedTeacherName || "").trim();
    if (state.selectedTeacherName) {
      rememberTeacherName(state.selectedTeacherName);
    }
    state.selectedGuruType = normalizeGuruType(el.guruJenis?.value || state.selectedGuruType);
    state.selectedDate = normalizeDateInput(el.tarikh?.value) || state.selectedDate;
    if (state.selectedDate) {
      const pickedDate = new Date(`${state.selectedDate}T00:00:00`);
      state.selectedYear = String(pickedDate.getFullYear());
      state.selectedMonth = MONTHS[pickedDate.getMonth()];
      if (el.tahun) {
        el.tahun.value = state.selectedYear;
      }
    }

    if (!state.selectedTeacherName) {
      const message = "Sila isi Nama Guru dahulu.";
      setStatus(message, true);
      showPopupStatus(message, true);
      return;
    }
    if (!state.selectedGuruType) {
      const message = "Sila pilih jenis Guru dahulu.";
      setStatus(message, true);
      showPopupStatus(message, true);
      return;
    }
    if (!state.selectedDate) {
      const message = "Sila pilih Tarikh dahulu.";
      setStatus(message, true);
      showPopupStatus(message, true);
      return;
    }
    if (!state.selectedClass) {
      const message = "Sila pilih kelas dahulu.";
      setStatus(message, true);
      showPopupStatus(message, true);
      return;
    }

    if (state.selectedGuruType === "Nilam") {
      const check = validateNilamBalanceBeforeSave();
      if (!check.ok) {
        setStatus(
          "Simpanan ditangguhkan: jumlah Bahan Bacaan dan jumlah Bahasa tidak sepadan untuk sesetengah murid.",
          true
        );
        window.alert(check.message);
        return;
      }
    }

    let records = [];
    try {
      records = collectCurrentRows();
    } catch (error) {
      setStatus(error.message, true);
      showPopupStatus(error.message, true);
      return;
    }

    if (!records.length) {
      const message = "Tiada rekod untuk disimpan.";
      setStatus(message, true);
      showPopupStatus(message, true);
      return;
    }

    const config = window.NILAM_CONFIG || {};
    if (!config.supabaseUrl || !config.supabaseAnonKey) {
      const message = "Cloud-only mode aktif. Isi `config.js` Supabase untuk simpan rekod.";
      setStatus(message, true);
      showToast("Simpanan gagal (cloud only)", true);
      showPopupStatus(message, true);
      return;
    }

    const supabaseUrl = config.supabaseUrl.replace(/\/$/, "");
    const endpoint =
      `${supabaseUrl}/rest/v1/nilam_records?on_conflict=tahun,bulan,tarikh,no_kad_pengenalan,nama_pengisi,guru`;

    try {
      await ensureStudentsExistInSupabase(records, config);

      const response = await fetch(endpoint, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          apikey: config.supabaseAnonKey,
          Authorization: `Bearer ${config.supabaseAnonKey}`,
          Prefer: "resolution=merge-duplicates,return=representation",
        },
        body: JSON.stringify(records),
      });

      if (!response.ok) {
        const detail = await response.text();
        throw new Error(`Ralat simpanan Supabase (${response.status}): ${detail}`);
      }

      setStatus(`Berjaya simpan ${records.length} rekod ke Supabase.`);
      showToast("Berjaya disimpan");
      showPopupStatus("Berjaya disimpan", false);
      loadAndApplyTotals(state.selectedYear, config).catch(() => {});
    } catch (error) {
      console.error(error);
      const message = `Simpanan ke Supabase gagal (cloud-only, tiada simpanan offline). (${String(
        error.message || "ralat tidak diketahui"
      ).slice(0, 180)})`;
      setStatus(message, true);
      showToast("Simpanan gagal", true);
      showPopupStatus(message, true);
      loadAndApplyTotals(state.selectedYear, config).catch(() => {});
    }
  }

  function saveLocal(records) {
    const key = `nilam_records_${state.selectedYear}_${state.selectedMonth}_${state.selectedClass}`;
    const existing = parseStoredRecords(localGetItem(key));
    const mergedByKey = new Map();
    const put = (record) => {
      const normalized = normalizeRecordForCloud(record, parseLocalStorageRecordKey(key), 1);
      if (!normalized) {
        return;
      }
      mergedByKey.set(getSessionRecordKey(normalized), normalized);
    };
    existing.forEach(put);
    records.forEach(put);
    localSetItem(key, JSON.stringify([...mergedByKey.values()]));
  }

  async function prefillSavedValuesForSelection() {
    if (!state.selectedClass || !state.selectedMonth) {
      return;
    }

    const requestSeq = state.prefillRequestSeq + 1;
    state.prefillRequestSeq = requestSeq;
    const config = window.NILAM_CONFIG || {};

    try {
      const totals = await loadTotals(state.selectedYear, config);
      if (requestSeq !== state.prefillRequestSeq) {
        return;
      }
      applyTotalsToTable(totals.yearAinsTotals, totals.yearTotals, totals.allTimeTotals);

      if (!state.selectedTeacherName || !state.selectedGuruType || !state.selectedDate) {
        setStatus(
          "Sila isi Nama Guru, pilih Guru, Tarikh, dan Kelas. Data dalam jadual ialah untuk sesi baharu.",
          false
        );
        return;
      }

      const savedRecords = await loadSavedRecords(
        state.selectedYear,
        state.selectedMonth,
        state.selectedClass,
        state.selectedDate,
        state.selectedTeacherName,
        state.selectedGuruType
      );
      if (requestSeq !== state.prefillRequestSeq) {
        return;
      }

      if (savedRecords.length) {
        applySavedRecordsToTable(savedRecords);
        setPrefillLoadedStatus(
          state.selectedClass,
          state.selectedMonth,
          state.selectedYear,
          state.selectedDate,
          state.selectedTeacherName,
          state.selectedGuruType,
          savedRecords.length
        );
      } else {
        setStatus(
          `Sesi baharu untuk ${state.selectedTeacherName} (${state.selectedGuruType}) pada ${formatDisplayDate(
            state.selectedDate
          )}.`,
          false
        );
      }
    } catch (error) {
      if (requestSeq !== state.prefillRequestSeq) {
        return;
      }
      console.error(error);
      setStatus("Gagal memuatkan rekod terdahulu. Anda masih boleh isi dan simpan semula.", true);
    }
  }

  function applySavedRecordsToTable(savedRecords) {
    const rows = [...el.tbody.querySelectorAll("tr[data-row-id]")];
    if (!rows.length) {
      return;
    }

    const savedByNoKad = new Map(
      savedRecords.map((record) => [String(record.no_kad_pengenalan || "").trim(), record])
    );

    rows.forEach((row) => {
      const key = String(row.dataset.noKad || "").trim();
      const saved = savedByNoKad.get(key);
      if (!saved) {
        return;
      }

      inputColumns.forEach((colName) => {
        setNumberToCell(row, colName, saved[colName]);
      });
      row.dataset.hasSavedSession = "1";
      updateJumlahAktiviti(row);
    });
  }

  async function loadSavedRecords(year, month, kelas, tarikh, namaPengisi, guruType) {
    const config = window.NILAM_CONFIG || {};
    const localRecords = loadSavedRecordsFromLocal(year, month, kelas, tarikh, namaPengisi, guruType);
    if (config.supabaseUrl && config.supabaseAnonKey) {
      try {
        const supabaseRecords = await loadSavedRecordsFromSupabase(
          year,
          month,
          kelas,
          tarikh,
          namaPengisi,
          guruType,
          config
        );
        if (supabaseRecords.length || localRecords.length) {
          return mergeSavedRecords(supabaseRecords, localRecords);
        }
      } catch (error) {
        console.error(error);
      }
    }

    return localRecords;
  }

  function mergeSavedRecords(supabaseRecords, localRecords) {
    const mergedByKey = new Map();
    const putLocal = (record) => {
      const key = getSessionRecordKey(record);
      if (!key) {
        return;
      }
      const existing = mergedByKey.get(key);
      const nextTs = Date.parse(record.updated_at_client || "") || 0;
      if (!existing) {
        mergedByKey.set(key, { record, ts: nextTs, source: "local" });
        return;
      }
      if (existing.source !== "local") {
        return;
      }
      if (nextTs >= existing.ts) {
        mergedByKey.set(key, { record, ts: nextTs, source: "local" });
      }
    };
    const putCloud = (record) => {
      const key = getSessionRecordKey(record);
      if (!key) {
        return;
      }
      const nextTs = Date.parse(record.updated_at_client || "") || 0;
      const existing = mergedByKey.get(key);
      if (!existing || existing.source !== "cloud" || nextTs >= existing.ts) {
        // Supabase is authoritative for the same session key.
        mergedByKey.set(key, { record, ts: nextTs, source: "cloud" });
      }
    };

    (Array.isArray(localRecords) ? localRecords : []).forEach(putLocal);
    (Array.isArray(supabaseRecords) ? supabaseRecords : []).forEach(putCloud);
    return [...mergedByKey.values()].map((entry) => entry.record);
  }

  function loadSavedRecordsFromLocal(year, month, kelas, tarikh, namaPengisi, guruType) {
    const all = [];
    const yearPrefix = `nilam_records_${year}_${month}_`;
    const legacyPrefix = `nilam_records_${month}_`;

    for (let i = 0; i < localLength(); i += 1) {
      const key = localKey(i);
      if (!key) {
        continue;
      }
      if (key.startsWith(yearPrefix) || key.startsWith(legacyPrefix)) {
        const parsed = parseStoredRecords(localGetItem(key));
        if (parsed.length) {
          parsed.forEach((row, index) => {
            const normalized = normalizeRecordForCloud(row, parseLocalStorageRecordKey(key), index + 1);
            if (!normalized) {
              return;
            }
            all.push(normalized);
          });
        }
      }
    }

    return all.filter((row) => {
      return (
        String(row.tahun || "") === String(year) &&
        String(row.bulan || "") === String(month) &&
        String(row.kelas || "") === String(kelas || "") &&
        String(row.tarikh || "") === String(tarikh || "") &&
        normalizeNameKey(row.nama_pengisi) === normalizeNameKey(namaPengisi) &&
        normalizeGuruType(row.guru) === normalizeGuruType(guruType)
      );
    });
  }

  function parseStoredRecords(raw) {
    if (!raw) {
      return [];
    }
    try {
      const parsed = JSON.parse(raw);
      return Array.isArray(parsed) ? parsed : [];
    } catch (error) {
      console.error(error);
      return [];
    }
  }

  function loadLocalRecordsForYear(year) {
    const all = [];
    const prefix = `nilam_records_${year}_`;
    for (let i = 0; i < localLength(); i += 1) {
      const key = localKey(i);
      if (key && key.startsWith(prefix)) {
        const meta = parseLocalStorageRecordKey(key);
        parseStoredRecords(localGetItem(key)).forEach((row, index) => {
          const normalized = normalizeRecordForCloud(row, meta, index + 1);
          if (normalized) {
            all.push(normalized);
          }
        });
      }
    }
    return all;
  }

  function loadAllLocalRecords() {
    const all = [];
    const prefix = "nilam_records_";
    for (let i = 0; i < localLength(); i += 1) {
      const key = localKey(i);
      if (key && key.startsWith(prefix)) {
        const meta = parseLocalStorageRecordKey(key);
        parseStoredRecords(localGetItem(key)).forEach((row, index) => {
          const normalized = normalizeRecordForCloud(row, meta, index + 1);
          if (normalized) {
            all.push(normalized);
          }
        });
      }
    }
    return all;
  }

  function computeJumlahBacaanFromRecord(row, includeAins = false) {
    const totalWithoutAins = (
      clampNilamNumber(row?.bahan_digital) +
      clampNilamNumber(row?.bahan_bukan_buku) +
      clampNilamNumber(row?.fiksyen) +
      clampNilamNumber(row?.bukan_fiksyen)
    );
    if (!includeAins) {
      return totalWithoutAins;
    }
    return totalWithoutAins + clampNilamNumber(row?.ains);
  }

  async function fetchTotalsFromSupabase(year, config) {
    const supabaseUrl = config.supabaseUrl.replace(/\/$/, "");
    const params = new URLSearchParams({
      select: "no_kad_pengenalan,bahan_digital,bahan_bukan_buku,fiksyen,bukan_fiksyen,ains,jumlah_aktiviti",
      limit: "10000",
    });
    if (year) {
      params.set("tahun", `eq.${year}`);
    }
    const response = await fetch(`${supabaseUrl}/rest/v1/nilam_records?${params.toString()}`, {
      cache: "no-store",
      headers: {
        apikey: config.supabaseAnonKey,
        Authorization: `Bearer ${config.supabaseAnonKey}`,
      },
    });
    if (!response.ok) {
      throw new Error(`Ralat muat jumlah Supabase (${response.status})`);
    }
    return response.json();
  }

  async function loadTotals(year, config) {
    if (config.supabaseUrl && config.supabaseAnonKey) {
      try {
        const [yearRecords, allRecords] = await Promise.all([
          fetchTotalsFromSupabase(year, config),
          fetchTotalsFromSupabase(null, config),
        ]);
        return {
          yearAinsTotals: computeAinsTotalsMap(yearRecords),
          yearTotals: computeTotalsMap(yearRecords, state.includeAinsInJumlah),
          allTimeTotals: computeTotalsMap(allRecords, state.includeAinsInJumlah),
        };
      } catch (error) {
        console.error("Gagal muat jumlah dari Supabase", error);
        return {
          yearAinsTotals: new Map(),
          yearTotals: new Map(),
          allTimeTotals: new Map(),
        };
      }
    }
    if (CLOUD_ONLY_MODE) {
      return {
        yearAinsTotals: new Map(),
        yearTotals: new Map(),
        allTimeTotals: new Map(),
      };
    }
    const localYearRecords = loadLocalRecordsForYear(year);
    const localAllRecords = loadAllLocalRecords();
    return {
      yearAinsTotals: computeAinsTotalsMap(localYearRecords),
      yearTotals: computeTotalsMap(localYearRecords, state.includeAinsInJumlah),
      allTimeTotals: computeTotalsMap(localAllRecords, state.includeAinsInJumlah),
    };
  }

  function applyTotalsToTable(yearAinsTotals, yearTotals, allTimeTotals) {
    const rows = [...el.tbody.querySelectorAll("tr[data-row-id]")];
    rows.forEach((row) => {
      const noKad = String(row.dataset.noKad || "").trim();
      const ainsYearCell = row.querySelector('[data-col="ains_sepanjang_tahun"]');
      const yrCell = row.querySelector('[data-col="jumlah_tahun"]');
      const atCell = row.querySelector('[data-col="jumlah_all_time"]');
      if (ainsYearCell) {
        ainsYearCell.textContent = String(yearAinsTotals.get(noKad) || 0);
      }
      if (yrCell) {
        yrCell.textContent = String(yearTotals.get(noKad) || 0);
      }
      if (atCell) {
        atCell.textContent = String(allTimeTotals.get(noKad) || 0);
      }
    });
  }

  async function loadAndApplyTotals(year, config) {
    const totals = await loadTotals(year, config);
    applyTotalsToTable(totals.yearAinsTotals, totals.yearTotals, totals.allTimeTotals);
  }

  async function loadSavedRecordsFromSupabase(year, month, kelas, tarikh, namaPengisi, guruType, config) {
    const supabaseUrl = config.supabaseUrl.replace(/\/$/, "");
    const params = new URLSearchParams({
      select:
        "tahun,no_kad_pengenalan,nama,kelas,bulan,tarikh,nama_pengisi,guru,bahan_digital,bahan_bukan_buku,fiksyen,bukan_fiksyen,ains,bahasa_melayu,bahasa_inggeris,lain_lain_bahasa,jumlah_aktiviti,updated_at_client",
      tahun: `eq.${year}`,
      bulan: `eq.${month}`,
      kelas: `eq.${kelas}`,
      tarikh: `eq.${tarikh}`,
      nama_pengisi: `eq.${namaPengisi}`,
      guru: `eq.${guruType}`,
      order: "updated_at_client.desc",
    });
    const endpoint = `${supabaseUrl}/rest/v1/nilam_records?${params.toString()}`;

    const response = await fetch(endpoint, {
      method: "GET",
      cache: "no-store",
      headers: {
        apikey: config.supabaseAnonKey,
        Authorization: `Bearer ${config.supabaseAnonKey}`,
      },
    });
    if (!response.ok) {
      const detail = await response.text();
      throw new Error(`Ralat muat rekod Supabase (${response.status}): ${detail}`);
    }
    const records = await response.json();
    return Array.isArray(records)
      ? records.map((row, index) => normalizeRecordForCloud(row, { tahun: year, bulan: month, kelas }, index + 1)).filter(Boolean)
      : [];
  }

  async function loadStudents() {
    const config = window.NILAM_CONFIG || {};
    const selectedYear = String(new Date().getFullYear());
    if (!config.supabaseUrl || !config.supabaseAnonKey) {
      throw new Error("Cloud-only mode memerlukan konfigurasi Supabase dalam config.js.");
    }

    // Supabase is always authoritative — active=neq.false filtered server-side.
    // No-IC students are stored with synthetic NOIC_ IDs in Supabase.
    try {
      const supabaseStudents = await loadStudentsFromSupabase(config, selectedYear);
      const normalized = normalizeStudents(supabaseStudents);
      if (normalized.length) {
        return normalized;
      }
      if (CLOUD_ONLY_MODE) {
        return [];
      }
    } catch (error) {
      console.error("Gagal muat senarai murid dari Supabase", error);
      if (CLOUD_ONLY_MODE) {
        throw error;
      }
    }

    // Offline fallback: localStorage override, then bundled data.
    let localStudents = [];
    const overrideRaw = localGetItem(NAMELIST_OVERRIDE_KEY);
    if (overrideRaw) {
      try {
        const parsed = JSON.parse(overrideRaw);
        if (Array.isArray(parsed)) {
          localStudents = parsed;
        }
      } catch (error) {
        console.error("Gagal parse namelist override", error);
      }
    }
    if (!localStudents.length) {
      localStudents = Array.isArray(window.NILAM_STUDENTS) ? window.NILAM_STUDENTS : [];
    }
    const normalizedLocal = normalizeStudents(localStudents);
    if (normalizedLocal.length) {
      return normalizedLocal;
    }

    const response = await fetch(CSV_FILE);
    if (!response.ok) {
      throw new Error(`Gagal baca fail CSV: ${CSV_FILE}`);
    }
    const text = await response.text();
    const rows = parseCsv(text);
    const namaColumn = resolveColumnName(rows, ["nama murid", "nama"]);
    const kelasColumn = resolveColumnName(rows, ["kelas 2026", "kelas"]);
    const noKadColumn = resolveColumnName(rows, ["no. kad pengenalan", "no kad pengenalan"]);

    return rows
      .map((row) => ({
        nama: getCsvValueByColumn(row, namaColumn),
        kelas: getCsvValueByColumn(row, kelasColumn),
        no_kad_pengenalan: getCsvValueByColumn(row, noKadColumn),
      }))
      .filter((row) => row.nama && row.kelas)
      .map((row) => ({
        nama: String(row.nama || "").trim(),
        kelas: String(row.kelas || "").trim(),
        no_kad_pengenalan: String(row.no_kad_pengenalan || "").trim(),
      }));
  }

  async function loadStudentsFromSupabase(config, year) {
    if (!config.supabaseUrl || !config.supabaseAnonKey) {
      return [];
    }

    const safeYear = /^\d{4}$/.test(String(year || "")) ? String(year) : String(new Date().getFullYear());
    const kelasField = `kelas_${safeYear}`;
    const supabaseUrl = config.supabaseUrl.replace(/\/$/, "");
    const params = new URLSearchParams({
      select: `no_kad_pengenalan,nama_murid,jantina,email_google_classroom,${kelasField}`,
      active: "neq.false",
      order: "nama_murid.asc",
      limit: "50000",
    });
    const endpoint = `${supabaseUrl}/rest/v1/nilam_students?${params.toString()}`;

    const response = await fetch(endpoint, {
      method: "GET",
      cache: "no-store",
      headers: {
        apikey: config.supabaseAnonKey,
        Authorization: `Bearer ${config.supabaseAnonKey}`,
      },
    });
    if (!response.ok) {
      const detail = await response.text();
      throw new Error(`Ralat muat nilam_students (${response.status}): ${detail}`);
    }

    const rows = await response.json();
    if (!Array.isArray(rows)) {
      return [];
    }

    return rows.map((row) => ({
      nama: String(row.nama_murid || "").trim(),
      kelas: String(row[kelasField] || "").trim(),
      no_kad_pengenalan: String(row.no_kad_pengenalan || "").trim(),
      jantina: String(row.jantina || "").trim(),
      email_google_classroom: String(row.email_google_classroom || "").trim(),
    }));
  }

  function normalizeStudents(students) {
    if (!Array.isArray(students)) {
      return [];
    }
    return students
      .map((row) => ({
        nama: String(row.nama || "").trim(),
        kelas: String(row.kelas || "").trim(),
        no_kad_pengenalan: String(row.no_kad_pengenalan || "").trim(),
      }))
      .filter((row) => row.nama && row.kelas);
  }

  function mergeStudentsByNoKad(primary, fallback) {
    const byNoKad = new Map();
    const byNameClass = new Set();
    const merged = [];

    const addRow = (row) => {
      const noKad = String(row.no_kad_pengenalan || "").trim();
      const key = `${row.nama.toLowerCase()}|${row.kelas.toLowerCase()}`;

      if (byNameClass.has(key)) {
        return;
      }
      if (noKad && byNoKad.has(noKad)) {
        return;
      }

      if (noKad) {
        byNoKad.set(noKad, row);
      }
      byNameClass.add(key);
      merged.push(row);
    };

    // Primary source (Supabase) gets priority.
    primary.forEach((row) => addRow(row));
    // Fallback source only fills truly missing students.
    fallback.forEach((row) => addRow(row));

    return merged;
  }

  function parseCsv(content) {
    const lines = [];
    let current = "";
    let inQuotes = false;

    for (let i = 0; i < content.length; i += 1) {
      const char = content[i];
      if (char === '"') {
        if (inQuotes && content[i + 1] === '"') {
          current += '"';
          i += 1;
        } else {
          inQuotes = !inQuotes;
        }
      } else if (char === "\n" && !inQuotes) {
        lines.push(current);
        current = "";
      } else if (char !== "\r") {
        current += char;
      }
    }
    if (current) {
      lines.push(current);
    }
    if (!lines.length) {
      return [];
    }

    const delimiter = detectDelimiter(lines[0]);
    const headers = splitCsvLine(lines[0], delimiter).map((h) => h.trim());
    return lines.slice(1).map((line) => {
      const values = splitCsvLine(line, delimiter);
      const row = {};
      headers.forEach((header, index) => {
        row[header] = (values[index] || "").trim();
      });
      return row;
    });
  }

  function splitCsvLine(line, delimiter) {
    const values = [];
    let token = "";
    let inQuotes = false;

    for (let i = 0; i < line.length; i += 1) {
      const char = line[i];
      if (char === '"') {
        if (inQuotes && line[i + 1] === '"') {
          token += '"';
          i += 1;
        } else {
          inQuotes = !inQuotes;
        }
      } else if (char === delimiter && !inQuotes) {
        values.push(token);
        token = "";
      } else {
        token += char;
      }
    }
    values.push(token);
    return values;
  }

  function detectDelimiter(headerLine) {
    const commaCount = (headerLine.match(/,/g) || []).length;
    const semicolonCount = (headerLine.match(/;/g) || []).length;
    return semicolonCount > commaCount ? ";" : ",";
  }

  function resolveColumnName(rows, aliases) {
    if (!rows.length) {
      return "";
    }

    const firstRowKeys = Object.keys(rows[0]);
    const normalizedKeyMap = {};
    firstRowKeys.forEach((key) => {
      normalizedKeyMap[normalizeHeader(key)] = key;
    });

    for (const alias of aliases) {
      const match = normalizedKeyMap[normalizeHeader(alias)];
      if (match) {
        return match;
      }
    }

    return "";
  }

  function getCsvValueByColumn(row, columnName) {
    if (!columnName) {
      return "";
    }
    return String(row[columnName] || "").trim();
  }

  function normalizeHeader(value) {
    return String(value || "")
      .replace(/^\uFEFF/, "")
      .toLowerCase()
      .replace(/\s+/g, " ")
      .trim();
  }

  async function autoSyncLocalRecordsToSupabase() {
    const config = window.NILAM_CONFIG || {};
    if (!config.supabaseUrl || !config.supabaseAnonKey) {
      return {
        mode: "cloud_error",
        syncedCount: 0,
        errorMessage: "Cloud-only mode memerlukan konfigurasi Supabase.",
      };
    }

    const payload = collectLocalRecordsForSync();
    if (!payload.length) {
      return { mode: "cloud_ready", syncedCount: 0 };
    }

    const supabaseUrl = config.supabaseUrl.replace(/\/$/, "");
    const endpoint =
      `${supabaseUrl}/rest/v1/nilam_records?on_conflict=tahun,bulan,tarikh,no_kad_pengenalan,nama_pengisi,guru`;

    try {
      await ensureStudentsExistInSupabase(payload, config);

      const response = await fetch(endpoint, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          apikey: config.supabaseAnonKey,
          Authorization: `Bearer ${config.supabaseAnonKey}`,
          Prefer: "resolution=merge-duplicates,return=minimal",
        },
        body: JSON.stringify(payload),
      });
      if (!response.ok) {
        const detail = await response.text();
        throw new Error(`Ralat sync local->Supabase (${response.status}): ${detail}`);
      }
      return { mode: "cloud_ready", syncedCount: payload.length };
    } catch (error) {
      console.error(error);
      return {
        mode: "cloud_error",
        syncedCount: 0,
        errorMessage: String(error?.message || "ralat tidak diketahui"),
      };
    }
  }

  function collectLocalRecordsForSync() {
    const dedup = new Map();
    for (let i = 0; i < localLength(); i += 1) {
      const key = localKey(i);
      if (!key || !key.startsWith("nilam_records_")) {
        continue;
      }
      const meta = parseLocalStorageRecordKey(key);
      const parsed = parseStoredRecords(localGetItem(key));
      parsed.forEach((row, index) => {
        const normalized = normalizeRecordForCloud(row, meta, index + 1);
        if (!normalized) {
          return;
        }
        dedup.set(getSessionRecordKey(normalized), normalized);
      });
    }
    return [...dedup.values()];
  }

  function parseLocalStorageRecordKey(key) {
    const match = String(key || "").match(/^nilam_records_(\d{4})_(.+?)_(.+)$/);
    if (match) {
      return { tahun: match[1], bulan: match[2], kelas: match[3] };
    }
    const legacyMatch = String(key || "").match(/^nilam_records_(.+?)_(.+)$/);
    if (legacyMatch) {
      return { tahun: "", bulan: legacyMatch[1], kelas: legacyMatch[2] };
    }
    return { tahun: "", bulan: "", kelas: "" };
  }

  function normalizeRecordForCloud(row, meta, bilFallback) {
    const noKad = String(row?.no_kad_pengenalan || "").trim();
    const tahun = String(row?.tahun || meta.tahun || "").trim();
    const bulan = String(row?.bulan || meta.bulan || "").trim();
    const kelas = String(row?.kelas || meta.kelas || "").trim();
    const nama = String(row?.nama || "").trim();
    const tarikh = normalizeRecordDate(row?.tarikh, tahun, bulan);
    const namaPengisi = String(row?.nama_pengisi || "").trim() || "Tidak Diketahui";
    const guruType = normalizeGuruType(row?.guru);
    if (!noKad || !tahun || !bulan || !kelas || !nama) {
      return null;
    }

    const bahanDigital = clampNilamNumber(row?.bahan_digital);
    const bahanBukanBuku = clampNilamNumber(row?.bahan_bukan_buku);
    const fiksyen = clampNilamNumber(row?.fiksyen);
    const bukanFiksyen = clampNilamNumber(row?.bukan_fiksyen);
    const ains = clampNilamNumber(row?.ains);
    const languageValues = resolveLanguageValuesFromGuruType(
      guruType,
      {
        bahan_digital: bahanDigital,
        bahan_bukan_buku: bahanBukanBuku,
        fiksyen,
        bukan_fiksyen: bukanFiksyen,
      },
      {
        bahasa_melayu: clampNilamNumber(row?.bahasa_melayu),
        bahasa_inggeris: clampNilamNumber(row?.bahasa_inggeris),
        lain_lain_bahasa: clampNilamNumber(row?.lain_lain_bahasa),
      }
    );
    const jumlahAsal = Number(row?.jumlah_aktiviti);
    const jumlahAktiviti = Number.isFinite(jumlahAsal)
      ? clampNilamNumber(jumlahAsal, 4995)
      : bahanDigital + bahanBukanBuku + fiksyen + bukanFiksyen + ains;
    const bilFromRow = Number(row?.bil);
    const bil = Number.isFinite(bilFromRow) && bilFromRow > 0 ? Math.trunc(bilFromRow) : bilFallback;

    return {
      no_kad_pengenalan: noKad,
      tahun,
      bulan,
      tarikh,
      nama_pengisi: namaPengisi,
      guru: guruType,
      bil,
      nama,
      kelas,
      bahan_digital: bahanDigital,
      bahan_bukan_buku: bahanBukanBuku,
      fiksyen,
      bukan_fiksyen: bukanFiksyen,
      ains,
      bahasa_melayu: languageValues.bahasa_melayu,
      bahasa_inggeris: languageValues.bahasa_inggeris,
      lain_lain_bahasa: languageValues.lain_lain_bahasa,
      jumlah_aktiviti: jumlahAktiviti,
      updated_at_client: row?.updated_at_client || new Date().toISOString(),
    };
  }

  function clampNilamNumber(value, max = 999) {
    const number = Number(value);
    if (!Number.isFinite(number)) {
      return 0;
    }
    return Math.max(0, Math.min(max, Math.trunc(number)));
  }

  async function ensureStudentsExistInSupabase(records, config) {
    if (!Array.isArray(records) || !records.length) {
      return;
    }

    const byNoKad = new Map();
    records.forEach((row) => {
      const noKad = String(row?.no_kad_pengenalan || "").trim();
      const nama = String(row?.nama || "").trim();
      if (!noKad || !nama) {
        return;
      }
      byNoKad.set(noKad, {
        no_kad_pengenalan: noKad,
        nama_murid: nama,
      });
    });

    const payload = [...byNoKad.values()];
    if (!payload.length) {
      return;
    }

    const supabaseUrl = config.supabaseUrl.replace(/\/$/, "");
    const endpoint = `${supabaseUrl}/rest/v1/nilam_students?on_conflict=no_kad_pengenalan`;
    const response = await fetch(endpoint, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        apikey: config.supabaseAnonKey,
        Authorization: `Bearer ${config.supabaseAnonKey}`,
        Prefer: "resolution=merge-duplicates,return=minimal",
      },
      body: JSON.stringify(payload),
    });

    if (!response.ok) {
      const detail = await response.text();
      throw new Error(`Ralat sync nilam_students (${response.status}): ${detail}`);
    }
  }

  function setStatus(message, isError) {
    el.status.textContent = message;
    el.status.style.color = isError ? "#b00020" : "";
  }

  function showPopupStatus(message, isError) {
    window.alert(isError ? String(message || "Operasi gagal.") : "Berjaya disimpan");
  }

  let toastTimer = null;
  function showToast(message, isError) {
    const toast = document.getElementById("saveToast");
    if (!toast) {
      return;
    }
    if (toastTimer) {
      clearTimeout(toastTimer);
    }
    toast.textContent = message;
    toast.classList.remove("toast-hide", "toast-error");
    if (isError) {
      toast.classList.add("toast-error");
    }
    toast.hidden = false;
    toastTimer = setTimeout(() => {
      toast.classList.add("toast-hide");
      toastTimer = setTimeout(() => {
        toast.hidden = true;
        toast.classList.remove("toast-hide", "toast-error");
        toastTimer = null;
      }, 400);
    }, 2200);
  }

  function setStatusHtml(messageHtml, isError) {
    el.status.innerHTML = messageHtml;
    el.status.style.color = isError ? "#b00020" : "";
  }

  function setPrefillLoadedStatus(kelas, bulan, tahun, tarikh, namaPengisi, guruType, recordCount) {
    setStatusHtml(
      `Data terdahulu dimuatkan untuk <span class="status-emph-kelas">Kelas ${escapeHtml(
        kelas
      )}</span> pada bulan <span class="status-emph-bulan">${escapeHtml(
        bulan
      )}</span> tahun <span class="status-emph-tahun">${escapeHtml(
        tahun
      )}</span>, Tarikh ${escapeHtml(formatDisplayDate(tarikh))}, Nama ${escapeHtml(
        namaPengisi
      )}, Guru ${escapeHtml(guruType)} (${Number(recordCount) || 0} rekod).`
    );
  }

  function loadTeacherNames() {
    const configNames = Array.isArray(window.NILAM_CONFIG?.teacherNames)
      ? window.NILAM_CONFIG.teacherNames
      : [];
    const localRaw = localGetItem(TEACHER_NAMES_KEY);
    let localNames = [];
    if (localRaw) {
      try {
        const parsed = JSON.parse(localRaw);
        if (Array.isArray(parsed)) {
          localNames = parsed;
        }
      } catch (error) {
        console.error("Gagal parse senarai nama guru", error);
      }
    }
    const merged = new Set();
    [...configNames, ...localNames].forEach((value) => {
      const clean = String(value || "").trim();
      if (!clean || isSystemTeacherName(clean)) {
        return;
      }
      merged.add(clean);
    });
    loadAllLocalRecords().forEach((row) => {
      const clean = String(row?.nama_pengisi || "").trim();
      if (clean && !isSystemTeacherName(clean)) {
        merged.add(clean);
      }
    });
    return sanitizeTeacherNames([...merged]);
  }

  async function fetchTeacherNamesFromSupabase(config) {
    if (!config.supabaseUrl || !config.supabaseAnonKey) {
      return [];
    }
    const supabaseUrl = config.supabaseUrl.replace(/\/$/, "");
    const params = new URLSearchParams({
      select: "nama_guru",
      order: "nama_guru.asc",
      limit: "5000",
    });
    const endpoint = `${supabaseUrl}/rest/v1/nilam_teachers?${params.toString()}`;
    const response = await fetch(endpoint, {
      method: "GET",
      cache: "no-store",
      headers: {
        apikey: config.supabaseAnonKey,
        Authorization: `Bearer ${config.supabaseAnonKey}`,
      },
    });
    if (!response.ok) {
      const detail = await response.text();
      throw new Error(`Ralat muat nama guru Supabase (${response.status}): ${detail}`);
    }
    const rows = await response.json();
    if (!Array.isArray(rows)) {
      return [];
    }
    return rows
      .map((row) => String(row?.nama_guru || "").trim())
      .filter((name) => name && !isSystemTeacherName(name));
  }

  async function refreshTeacherNamesFromSupabase() {
    const config = window.NILAM_CONFIG || {};
    if (!config.supabaseUrl || !config.supabaseAnonKey) {
      return;
    }
    const cloudNames = await fetchTeacherNamesFromSupabase(config);
    if (!cloudNames.length) {
      return;
    }
    const merged = sanitizeTeacherNames([...state.teacherNames, ...cloudNames]);
    if (merged.length === state.teacherNames.length && merged.every((name, idx) => name === state.teacherNames[idx])) {
      return;
    }
    state.teacherNames = merged;
    saveTeacherNames();
    renderTeacherSuggestions(state.selectedTeacherName);
    fitTeacherNameInputWidth();
  }

  function saveTeacherNames() {
    state.teacherNames = sanitizeTeacherNames(state.teacherNames);
    localSetItem(TEACHER_NAMES_KEY, JSON.stringify(state.teacherNames));
  }

  function renderTeacherSuggestions(filterText) {
    if (!el.namaGuruCadanganList || !el.namaPengisi) {
      return;
    }
    const query = String(filterText || "").trim().toLowerCase();
    const safeNames = sanitizeTeacherNames(state.teacherNames);
    const names = query
      ? safeNames.filter((name) => String(name || "").toLowerCase().includes(query))
      : safeNames;
    const uniqueNames = [...new Set(names)].slice(0, 80);

    el.namaGuruCadanganList.innerHTML = "";
    uniqueNames.forEach((name) => {
      const item = document.createElement("li");
      item.className = "teacher-suggestion-item";
      item.dataset.value = name;
      item.textContent = name;
      item.setAttribute("role", "option");
      el.namaGuruCadanganList.appendChild(item);
    });

    const isFocused = document.activeElement === el.namaPengisi;
    const inputValue = String(el.namaPengisi.value || "").trim().toLowerCase();
    const exactMatch = inputValue && uniqueNames.some((name) => name.toLowerCase() === inputValue);
    const shouldShow = isFocused && uniqueNames.length > 0 && !(exactMatch && uniqueNames.length === 1);
    el.namaGuruCadanganList.hidden = !shouldShow;
  }

  function hideTeacherSuggestions() {
    if (!el.namaGuruCadanganList) {
      return;
    }
    el.namaGuruCadanganList.hidden = true;
  }

  async function chooseTeacherSuggestion(value) {
    const chosen = String(value || "").trim();
    if (!chosen || !el.namaPengisi) {
      hideTeacherSuggestions();
      return;
    }
    el.namaPengisi.value = chosen;
    state.selectedTeacherName = chosen;
    hideTeacherSuggestions();
    fitTeacherNameInputWidth();
    rememberTeacherName(chosen);
    if (state.selectedClass) {
      await renderTableAndPrefill();
    }
  }

  function rememberTeacherName(value) {
    const clean = String(value || "").trim();
    if (!clean || isSystemTeacherName(clean)) {
      return;
    }
    if (!state.teacherNames.includes(clean)) {
      state.teacherNames.push(clean);
      state.teacherNames.sort((a, b) => a.localeCompare(b, "ms"));
      saveTeacherNames();
      renderTeacherSuggestions(state.selectedTeacherName);
    }
    fitTeacherNameInputWidth();
  }

  function isSystemTeacherName(value) {
    const clean = String(value || "").trim().toUpperCase();
    return SYSTEM_TEACHER_NAMES.has(clean);
  }

  function sanitizeTeacherNames(names) {
    const seen = new Set();
    const cleaned = [];
    (Array.isArray(names) ? names : []).forEach((value) => {
      const clean = String(value || "").trim();
      if (!clean || isSystemTeacherName(clean)) {
        return;
      }
      const key = clean.toLowerCase();
      if (seen.has(key)) {
        return;
      }
      seen.add(key);
      cleaned.push(clean);
    });
    return cleaned.sort((a, b) => a.localeCompare(b, "ms"));
  }

  function applyGuruModeToVisibleRows() {
    const rows = [...el.tbody.querySelectorAll("tr[data-row-id]")];
    const isAutoGuru = state.selectedGuruType === "BM" || state.selectedGuruType === "BI";
    rows.forEach((row) => {
      ["bahasa_melayu", "bahasa_inggeris", "lain_lain_bahasa"].forEach((colName) => {
        const input = row.querySelector(`input[data-col="${colName}"]`);
        if (input) {
          input.readOnly = isAutoGuru;
          input.disabled = isAutoGuru;
        }
      });
      updateJumlahAktiviti(row);
    });
  }

  function applyGuruAutoLanguageForRow(row) {
    if (!row) {
      return;
    }
    const nextValues = resolveLanguageValuesFromGuruType(
      state.selectedGuruType,
      {
        bahan_digital: getNumberFromCell(row, "bahan_digital"),
        bahan_bukan_buku: getNumberFromCell(row, "bahan_bukan_buku"),
        fiksyen: getNumberFromCell(row, "fiksyen"),
        bukan_fiksyen: getNumberFromCell(row, "bukan_fiksyen"),
      },
      {
        bahasa_melayu: getNumberFromCell(row, "bahasa_melayu"),
        bahasa_inggeris: getNumberFromCell(row, "bahasa_inggeris"),
        lain_lain_bahasa: getNumberFromCell(row, "lain_lain_bahasa"),
      }
    );
    setNumberToCell(row, "bahasa_melayu", nextValues.bahasa_melayu);
    setNumberToCell(row, "bahasa_inggeris", nextValues.bahasa_inggeris);
    setNumberToCell(row, "lain_lain_bahasa", nextValues.lain_lain_bahasa);
  }

  function resolveLanguageValuesFromGuruType(guruType, materials, currentLanguage) {
    const normalizedType = normalizeGuruType(guruType);
    const totalMaterials =
      clampNilamNumber(materials?.bahan_digital) +
      clampNilamNumber(materials?.bahan_bukan_buku) +
      clampNilamNumber(materials?.fiksyen) +
      clampNilamNumber(materials?.bukan_fiksyen);

    if (normalizedType === "BM") {
      return {
        bahasa_melayu: totalMaterials,
        bahasa_inggeris: 0,
        lain_lain_bahasa: 0,
      };
    }
    if (normalizedType === "BI") {
      return {
        bahasa_melayu: 0,
        bahasa_inggeris: totalMaterials,
        lain_lain_bahasa: 0,
      };
    }
    return {
      bahasa_melayu: clampNilamNumber(currentLanguage?.bahasa_melayu),
      bahasa_inggeris: clampNilamNumber(currentLanguage?.bahasa_inggeris),
      lain_lain_bahasa: clampNilamNumber(currentLanguage?.lain_lain_bahasa),
    };
  }

  function normalizeGuruType(value) {
    const text = String(value || "").trim();
    if (text === "BM" || text === "BI" || text === "Nilam") {
      return text;
    }
    return "Nilam";
  }

  function normalizeDateInput(value) {
    const raw = String(value || "").trim();
    return /^\d{4}-\d{2}-\d{2}$/.test(raw) ? raw : "";
  }

  function normalizeRecordDate(rawDate, tahun, bulan) {
    const normalized = normalizeDateInput(rawDate);
    if (normalized) {
      return normalized;
    }
    const monthIndex = MONTHS.findIndex((m) => m.toLowerCase() === String(bulan || "").trim().toLowerCase());
    if (/^\d{4}$/.test(String(tahun || "")) && monthIndex >= 0) {
      return `${tahun}-${String(monthIndex + 1).padStart(2, "0")}-01`;
    }
    if (/^\d{4}$/.test(String(tahun || ""))) {
      return `${tahun}-01-01`;
    }
    return "1970-01-01";
  }

  function toIsoDate(dateValue) {
    const date = dateValue instanceof Date ? dateValue : new Date(dateValue);
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
  }

  function formatDisplayDate(isoDate) {
    const safe = normalizeDateInput(isoDate);
    if (!safe) {
      return String(isoDate || "-");
    }
    const [year, month, day] = safe.split("-");
    return `${day}/${month}/${year}`;
  }

  function normalizeNameKey(value) {
    return String(value || "").trim().toLowerCase();
  }

  function getSessionRecordKey(record) {
    const tahun = String(record?.tahun || "").trim();
    const bulan = String(record?.bulan || "").trim();
    const noKad = String(record?.no_kad_pengenalan || "").trim();
    const tarikh = normalizeRecordDate(record?.tarikh, tahun, bulan);
    const namaPengisi = normalizeNameKey(record?.nama_pengisi || "");
    const guru = normalizeGuruType(record?.guru);
    if (!tahun || !bulan || !tarikh || !noKad || !namaPengisi || !guru) {
      return "";
    }
    return `${tahun}|${bulan}|${tarikh}|${noKad}|${namaPengisi}|${guru}`;
  }

  function escapeHtml(value) {
    return String(value || "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#39;");
  }

  function sanitizeId(value) {
    return value.toLowerCase().replace(/[^a-z0-9]+/g, "_");
  }

})();
