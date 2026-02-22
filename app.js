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

  const state = {
    rawStudents: [],
    selectedYear: String(new Date().getFullYear()),
    selectedClass: "",
    selectedMonth: "",
    prefillRequestSeq: 0,
  };

  const el = {
    tahun: document.getElementById("tahun"),
    bulan: document.getElementById("bulan"),
    kelas: document.getElementById("kelas"),
    tbody: document.querySelector("#studentTable tbody"),
    status: document.getElementById("status"),
    saveAll: document.getElementById("saveAll"),
    saveAllBottom: document.getElementById("saveAllBottom"),
  };

  init();

  async function init() {
    initYearField();
    initMonthDropdown();
    bindEvents();
    try {
      const students = await loadStudents();
      hydrateStudents(students);
      const syncResult = await autoSyncLocalRecordsToSupabase();
      if (syncResult.mode === "local_only") {
        setStatus(
          `Data murid berjaya dimuatkan: ${students.length} murid. Simpanan merentas peranti belum aktif (isi config.js Supabase).`,
          true
        );
        return;
      }
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
          `Data murid berjaya dimuatkan: ${students.length} murid. ${syncResult.syncedCount} rekod local telah disegerakkan ke cloud.`
        );
        return;
      }
      setStatus(`Data murid berjaya dimuatkan: ${students.length} murid.`);
    } catch (error) {
      setStatus("Gagal memuatkan data murid. Semak `students-data.js` atau fail CSV.", true);
      console.error(error);
    }
  }

  function bindEvents() {
    el.kelas.addEventListener("change", async () => {
      state.selectedClass = el.kelas.value;
      await renderTableAndPrefill();
    });

    el.bulan.addEventListener("change", async () => {
      state.selectedMonth = el.bulan.value;
      if (state.selectedClass) {
        await renderTableAndPrefill();
      }
    });

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

  function initMonthDropdown() {
    MONTHS.forEach((month) => {
      const option = document.createElement("option");
      option.value = month;
      option.textContent = month;
      el.bulan.appendChild(option);
    });

    const currentMonthIndex = new Date().getMonth();
    el.bulan.value = MONTHS[currentMonthIndex];
    state.selectedMonth = MONTHS[currentMonthIndex];
  }

  function initClassDropdown(students) {
    const classes = [...new Set(students.map((s) => s.kelas))].sort();
    classes.forEach((className) => {
      const option = document.createElement("option");
      option.value = className;
      option.textContent = className;
      el.kelas.appendChild(option);
    });
  }

  function resetClassDropdown() {
    el.kelas.innerHTML = '<option value="">Pilih kelas</option>';
    state.selectedClass = "";
  }

  function hydrateStudents(students) {
    state.rawStudents = students;
    resetClassDropdown();
    initClassDropdown(students);
    el.tbody.innerHTML =
      '<tr><td colspan="11" class="empty">Pilih kelas untuk paparkan senarai murid.</td></tr>';
  }

  async function renderTableAndPrefill() {
    renderTable();
    await prefillSavedValuesForSelection();
  }

  function renderTable() {
    if (!state.selectedClass) {
      el.tbody.innerHTML =
        '<tr><td colspan="11" class="empty">Pilih kelas untuk paparkan senarai murid.</td></tr>';
      return;
    }

    const students = state.rawStudents.filter((s) => s.kelas === state.selectedClass);
    if (!students.length) {
      el.tbody.innerHTML =
        '<tr><td colspan="11" class="empty">Tiada murid untuk kelas ini.</td></tr>';
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
        )}" data-no-kad="${escapeHtml(student.no_kad_pengenalan || "")}">
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
          </tr>
        `;
      })
      .join("");

    el.tbody.innerHTML = rowsHtml;
    attachRowInputHandlers();
  }

  function numericInput(colName) {
    return `<input type="number" min="0" max="999" step="1" value="0" data-col="${colName}">`;
  }

  function attachRowInputHandlers() {
    const inputs = el.tbody.querySelectorAll('input[type="number"]:not([readonly])');
    inputs.forEach((input) => {
      input.addEventListener("input", (event) => {
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
    const bm = getNumberFromCell(row, "bahasa_melayu");
    const bi = getNumberFromCell(row, "bahasa_inggeris");
    const lain = getNumberFromCell(row, "lain_lain_bahasa");
    const total = bm + bi + lain;
    row.querySelector('[data-col="jumlah_aktiviti"]').textContent = String(total);
  }

  function getNumberFromCell(row, colName) {
    const input = row.querySelector(`input[data-col="${colName}"]`);
    return Number(input.value) || 0;
  }

  function setNumberToCell(row, colName, value) {
    const input = row.querySelector(`input[data-col="${colName}"]`);
    if (!input) {
      return;
    }
    const number = Number(value);
    if (!Number.isFinite(number)) {
      input.value = "0";
      return;
    }
    input.value = String(Math.max(0, Math.min(999, Math.trunc(number))));
  }

  function collectCurrentRows() {
    const rows = [...el.tbody.querySelectorAll("tr[data-row-id]")];
    return rows.map((row, index) => {
      const nama = (row.dataset.nama || "").trim();
      const kelas = (row.dataset.kelas || "").trim();
      const noKad = (row.dataset.noKad || "").trim();
      if (!noKad) {
        throw new Error(`No. Kad Pengenalan tiada untuk murid ${nama}. Sila kemas kini senarai murid di Admin.`);
      }
      const jumlahAktiviti = Number(
        row.querySelector('[data-col="jumlah_aktiviti"]').textContent || "0"
      );
      const record = {
        no_kad_pengenalan: noKad,
        tahun: state.selectedYear,
        bulan: state.selectedMonth,
        bil: index + 1,
        nama,
        kelas,
        bahan_digital: getNumberFromCell(row, "bahan_digital"),
        bahan_bukan_buku: getNumberFromCell(row, "bahan_bukan_buku"),
        fiksyen: getNumberFromCell(row, "fiksyen"),
        bukan_fiksyen: getNumberFromCell(row, "bukan_fiksyen"),
        bahasa_melayu: getNumberFromCell(row, "bahasa_melayu"),
        bahasa_inggeris: getNumberFromCell(row, "bahasa_inggeris"),
        lain_lain_bahasa: getNumberFromCell(row, "lain_lain_bahasa"),
        jumlah_aktiviti: Number.isFinite(jumlahAktiviti) ? jumlahAktiviti : 0,
        updated_at_client: new Date().toISOString(),
      };

      for (const col of inputColumns) {
        if (!Number.isInteger(record[col]) || record[col] < 0 || record[col] > 999) {
          throw new Error(`Nilai tidak sah untuk ${col} pada murid ${record.nama}`);
        }
      }

      return record;
    });
  }

  async function saveAllRecords() {
    if (!state.selectedClass) {
      setStatus("Sila pilih kelas dahulu.", true);
      return;
    }

    let records = [];
    try {
      records = collectCurrentRows();
    } catch (error) {
      setStatus(error.message, true);
      return;
    }

    if (!records.length) {
      setStatus("Tiada rekod untuk disimpan.", true);
      return;
    }

    const config = window.NILAM_CONFIG || {};
    if (!config.supabaseUrl || !config.supabaseAnonKey) {
      saveLocal(records);
      setStatus(
        "Rekod disimpan dalam localStorage (fallback). Isi `config.js` untuk simpan ke Supabase."
      );
      return;
    }

    const supabaseUrl = config.supabaseUrl.replace(/\/$/, "");
    const endpoint = `${supabaseUrl}/rest/v1/nilam_records?on_conflict=tahun,bulan,no_kad_pengenalan`;

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

      saveLocal(records);
      setStatus(`Berjaya simpan ${records.length} rekod ke Supabase.`);
    } catch (error) {
      console.error(error);
      saveLocal(records);
      setStatus(
        `Simpanan ke Supabase gagal. Rekod disimpan dalam localStorage sebagai sandaran. (${String(
          error.message || "ralat tidak diketahui"
        ).slice(0, 180)})`,
        true
      );
    }
  }

  function saveLocal(records) {
    const key = `nilam_records_${state.selectedYear}_${state.selectedMonth}_${state.selectedClass}`;
    localStorage.setItem(key, JSON.stringify(records));
  }

  async function prefillSavedValuesForSelection() {
    if (!state.selectedClass || !state.selectedMonth) {
      return;
    }

    const requestSeq = state.prefillRequestSeq + 1;
    state.prefillRequestSeq = requestSeq;

    try {
      const savedRecords = await loadSavedRecords(
        state.selectedYear,
        state.selectedMonth,
        state.selectedClass
      );
      if (requestSeq !== state.prefillRequestSeq) {
        return;
      }
      if (!savedRecords.length) {
        return;
      }

      applySavedRecordsToTable(savedRecords);
      setStatus(
        `Data terdahulu dimuatkan untuk ${state.selectedYear} ${state.selectedMonth} ${state.selectedClass} (${savedRecords.length} rekod).`
      );
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
      updateJumlahAktiviti(row);
    });
  }

  async function loadSavedRecords(year, month, kelas) {
    const config = window.NILAM_CONFIG || {};
    if (config.supabaseUrl && config.supabaseAnonKey) {
      try {
        const records = await loadSavedRecordsFromSupabase(year, month, config);
        if (records.length) {
          return records;
        }
      } catch (error) {
        console.error(error);
      }
    }

    return loadSavedRecordsFromLocal(year, month);
  }

  function loadSavedRecordsFromLocal(year, month) {
    const all = [];
    const yearPrefix = `nilam_records_${year}_${month}_`;
    const legacyPrefix = `nilam_records_${month}_`;

    for (let i = 0; i < localStorage.length; i += 1) {
      const key = localStorage.key(i);
      if (!key) {
        continue;
      }
      if (key.startsWith(yearPrefix) || key.startsWith(legacyPrefix)) {
        const parsed = parseStoredRecords(localStorage.getItem(key));
        if (parsed.length) {
          all.push(...parsed);
        }
      }
    }

    return all;
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

  async function loadSavedRecordsFromSupabase(year, month, config) {
    const supabaseUrl = config.supabaseUrl.replace(/\/$/, "");
    const params = new URLSearchParams({
      select:
        "tahun,no_kad_pengenalan,nama,kelas,bulan,bahan_digital,bahan_bukan_buku,fiksyen,bukan_fiksyen,bahasa_melayu,bahasa_inggeris,lain_lain_bahasa,jumlah_aktiviti,updated_at_client",
      tahun: `eq.${year}`,
      bulan: `eq.${month}`,
      order: "updated_at_client.desc",
    });
    const endpoint = `${supabaseUrl}/rest/v1/nilam_records?${params.toString()}`;

    const response = await fetch(endpoint, {
      method: "GET",
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
    return Array.isArray(records) ? records : [];
  }

  async function loadStudents() {
    const config = window.NILAM_CONFIG || {};
    const selectedYear = String(new Date().getFullYear());
    let localStudents = [];
    const overrideRaw = localStorage.getItem(NAMELIST_OVERRIDE_KEY);
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

    try {
      const supabaseStudents = await loadStudentsFromSupabase(config, selectedYear);
      const normalizedSupabase = normalizeStudents(supabaseStudents);
      if (normalizedSupabase.length) {
        return mergeStudentsByNoKad(normalizedSupabase, normalizedLocal);
      }
    } catch (error) {
      console.error("Gagal muat senarai murid dari Supabase", error);
    }

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
      order: "nama_murid.asc",
      limit: "50000",
    });
    const endpoint = `${supabaseUrl}/rest/v1/nilam_students?${params.toString()}`;

    const response = await fetch(endpoint, {
      method: "GET",
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

    primary.forEach((row) => {
      const noKad = String(row.no_kad_pengenalan || "").trim();
      const key = `${row.nama.toLowerCase()}|${row.kelas.toLowerCase()}`;
      if (noKad) {
        byNoKad.set(noKad, row);
      } else if (!byNameClass.has(key)) {
        byNameClass.add(key);
        byNoKad.set(`__${key}`, row);
      }
    });

    fallback.forEach((row) => {
      const noKad = String(row.no_kad_pengenalan || "").trim();
      const key = `${row.nama.toLowerCase()}|${row.kelas.toLowerCase()}`;
      if (noKad) {
        if (!byNoKad.has(noKad)) {
          byNoKad.set(noKad, row);
        }
        return;
      }
      if (!byNameClass.has(key)) {
        byNameClass.add(key);
        byNoKad.set(`__${key}`, row);
      }
    });

    return [...byNoKad.values()];
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
      return { mode: "local_only", syncedCount: 0 };
    }

    const payload = collectLocalRecordsForSync();
    if (!payload.length) {
      return { mode: "cloud_ready", syncedCount: 0 };
    }

    const supabaseUrl = config.supabaseUrl.replace(/\/$/, "");
    const endpoint = `${supabaseUrl}/rest/v1/nilam_records?on_conflict=tahun,bulan,no_kad_pengenalan`;

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
    for (let i = 0; i < localStorage.length; i += 1) {
      const key = localStorage.key(i);
      if (!key || !key.startsWith("nilam_records_")) {
        continue;
      }
      const meta = parseLocalStorageRecordKey(key);
      const parsed = parseStoredRecords(localStorage.getItem(key));
      parsed.forEach((row, index) => {
        const normalized = normalizeRecordForCloud(row, meta, index + 1);
        if (!normalized) {
          return;
        }
        dedup.set(
          `${normalized.tahun}|${normalized.bulan}|${normalized.no_kad_pengenalan}`,
          normalized
        );
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
    if (!noKad || !tahun || !bulan || !kelas || !nama) {
      return null;
    }

    const bahanDigital = clampNilamNumber(row?.bahan_digital);
    const bahanBukanBuku = clampNilamNumber(row?.bahan_bukan_buku);
    const fiksyen = clampNilamNumber(row?.fiksyen);
    const bukanFiksyen = clampNilamNumber(row?.bukan_fiksyen);
    const bahasaMelayu = clampNilamNumber(row?.bahasa_melayu);
    const bahasaInggeris = clampNilamNumber(row?.bahasa_inggeris);
    const lainLainBahasa = clampNilamNumber(row?.lain_lain_bahasa);
    const jumlahAsal = Number(row?.jumlah_aktiviti);
    const jumlahAktiviti = Number.isFinite(jumlahAsal)
      ? clampNilamNumber(jumlahAsal, 2997)
      : bahasaMelayu + bahasaInggeris + lainLainBahasa;
    const bilFromRow = Number(row?.bil);
    const bil = Number.isFinite(bilFromRow) && bilFromRow > 0 ? Math.trunc(bilFromRow) : bilFallback;

    return {
      no_kad_pengenalan: noKad,
      tahun,
      bulan,
      bil,
      nama,
      kelas,
      bahan_digital: bahanDigital,
      bahan_bukan_buku: bahanBukanBuku,
      fiksyen,
      bukan_fiksyen: bukanFiksyen,
      bahasa_melayu: bahasaMelayu,
      bahasa_inggeris: bahasaInggeris,
      lain_lain_bahasa: lainLainBahasa,
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

  function escapeHtml(value) {
    return value
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
