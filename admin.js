(function () {
  "use strict";

  const AUTH_USER = "admin";
  const AUTH_PASS = "nilam_admin";
  const AUTH_SESSION_KEY = "nilam_admin_auth_v1";
  const NAMELIST_OVERRIDE_KEY = "nilam_students_override_v1";
  const MONTHS = window.NILAM_DATA
    ? window.NILAM_DATA.MONTHS
    : [
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

  const state = {
    students: [],
    isManageOpen: false,
    pendingImportYear: "",
    pendingImportMonth: "",
  };

  const el = {
    loginInfo: document.getElementById("adminLoginInfo"),
    panel: document.getElementById("adminPanel"),
    loginBtn: document.getElementById("loginBtn"),
    importBtn: document.getElementById("importNamelistBtn"),
    fileInput: document.getElementById("namelistFileInput"),
    importDataBtn: document.getElementById("importDataBtn"),
    dataFileInput: document.getElementById("dataFileInput"),
    importDataModal: document.getElementById("importDataModal"),
    importYearInput: document.getElementById("importYearInput"),
    importMonthSelect: document.getElementById("importMonthSelect"),
    confirmImportDataBtn: document.getElementById("confirmImportDataBtn"),
    cancelImportDataBtn: document.getElementById("cancelImportDataBtn"),
    resetBtn: document.getElementById("resetDataBtn"),
    manageBtn: document.getElementById("manageStudentsBtn"),
    manageSection: document.getElementById("studentManageSection"),
    manageTbody: document.getElementById("studentManageTbody"),
    addRowBtn: document.getElementById("addStudentRowBtn"),
    saveStudentsBtn: document.getElementById("saveStudentsBtn"),
    logoutBtn: document.getElementById("logoutBtn"),
    status: document.getElementById("adminStatus"),
  };

  init();

  function init() {
    initImportMonthOptions();
    bindEvents();
    if (sessionStorage.getItem(AUTH_SESSION_KEY) === "1") {
      showPanel();
    } else {
      requestLogin();
    }
  }

  function bindEvents() {
    el.loginBtn.addEventListener("click", requestLogin);
    el.importBtn.addEventListener("click", startImport);
    el.fileInput.addEventListener("change", handleImportFile);
    el.importDataBtn.addEventListener("click", startImportData);
    el.dataFileInput.addEventListener("change", handleImportDataFile);
    el.confirmImportDataBtn.addEventListener("click", confirmImportDataSelection);
    el.cancelImportDataBtn.addEventListener("click", closeImportDataDialog);
    el.resetBtn.addEventListener("click", resetAllData);
    el.manageBtn.addEventListener("click", toggleManageStudents);
    el.addRowBtn.addEventListener("click", addStudentRow);
    el.saveStudentsBtn.addEventListener("click", saveManagedStudents);
    el.manageTbody.addEventListener("click", handleManageTableClick);
    el.logoutBtn.addEventListener("click", logout);
  }

  function requestLogin() {
    const username = window.prompt("Admin Username:", "");
    if (username === null) {
      setStatus("Login dibatalkan.", true);
      showLoginOnly();
      return;
    }

    const password = window.prompt("Admin Password:", "");
    if (password === null) {
      setStatus("Login dibatalkan.", true);
      showLoginOnly();
      return;
    }

    if (username === AUTH_USER && password === AUTH_PASS) {
      sessionStorage.setItem(AUTH_SESSION_KEY, "1");
      showPanel();
      setStatus("Login berjaya.");
      return;
    }

    sessionStorage.removeItem(AUTH_SESSION_KEY);
    showLoginOnly();
    setStatus("Username atau password tidak sah.", true);
  }

  function showPanel() {
    el.loginInfo.hidden = true;
    el.panel.hidden = false;
    loadStudentsForManage();
  }

  function showLoginOnly() {
    el.loginInfo.hidden = false;
    el.panel.hidden = true;
    if (el.manageSection) {
      el.manageSection.hidden = true;
    }
    state.isManageOpen = false;
  }

  function logout() {
    sessionStorage.removeItem(AUTH_SESSION_KEY);
    showLoginOnly();
    setStatus("Anda telah logout.");
  }

  function toggleManageStudents() {
    state.isManageOpen = !state.isManageOpen;
    el.manageSection.hidden = !state.isManageOpen;
    if (state.isManageOpen) {
      loadStudentsForManage();
      setStatus("Pengurusan senarai murid dibuka.");
    }
  }

  function loadStudentsForManage() {
    const raw = getCurrentNamelist();
    state.students = raw.map(normalizeStudentRow).filter((row) => row.nama || row.kelas);
    if (!state.students.length) {
      state.students = [emptyStudentRow()];
    }
    renderManageTable();
  }

  function renderManageTable() {
    const rowsHtml = state.students
      .map(
        (row, index) => `
        <tr>
          <td><input type="text" data-field="nama" data-index="${index}" value="${escapeAttr(row.nama)}"></td>
          <td><input type="text" data-field="jantina" data-index="${index}" value="${escapeAttr(
          row.jantina
        )}"></td>
          <td><input type="text" data-field="kelas" data-index="${index}" value="${escapeAttr(row.kelas)}"></td>
          <td><input type="text" data-field="no_kad_pengenalan" data-index="${index}" value="${escapeAttr(
          row.no_kad_pengenalan
        )}"></td>
          <td><input type="text" data-field="email_google_classroom" data-index="${index}" value="${escapeAttr(
          row.email_google_classroom
        )}"></td>
          <td><button type="button" class="mini-btn danger-btn" data-remove-index="${index}">Buang</button></td>
        </tr>
      `
      )
      .join("");

    el.manageTbody.innerHTML = rowsHtml;
  }

  function handleManageTableClick(event) {
    const button = event.target.closest("button[data-remove-index]");
    if (!button) {
      return;
    }

    const index = Number(button.dataset.removeIndex);
    if (!Number.isInteger(index) || index < 0 || index >= state.students.length) {
      return;
    }

    state.students.splice(index, 1);
    if (!state.students.length) {
      state.students.push(emptyStudentRow());
    }
    renderManageTable();
  }

  function addStudentRow() {
    syncTableToState();
    state.students.push(emptyStudentRow());
    renderManageTable();
  }

  async function saveManagedStudents() {
    try {
      syncTableToState();

      const cleaned = state.students
        .map(normalizeStudentRow)
        .filter((row) => row.nama || row.kelas || row.jantina || row.no_kad_pengenalan || row.email_google_classroom);

      const noKadSet = new Set();
      cleaned.forEach((row, idx) => {
        if (!row.nama || !row.kelas || !row.no_kad_pengenalan) {
          throw new Error(`Baris ${idx + 1}: Nama, Kelas dan No. Kad Pengenalan wajib diisi.`);
        }
        if (noKadSet.has(row.no_kad_pengenalan)) {
          throw new Error(`No. Kad Pengenalan berulang dikesan: ${row.no_kad_pengenalan}`);
        }
        noKadSet.add(row.no_kad_pengenalan);
      });

      if (!cleaned.length) {
        throw new Error("Senarai murid kosong. Tambah sekurang-kurangnya 1 murid.");
      }

      localStorage.setItem(NAMELIST_OVERRIDE_KEY, JSON.stringify(cleaned));
      state.students = cleaned;
      renderManageTable();
      await upsertStudentsToSupabase(cleaned);
      setStatus(`Senarai murid berjaya disimpan: ${cleaned.length} murid.`);
    } catch (error) {
      console.error(error);
      setStatus(error.message || "Gagal simpan senarai murid.", true);
    }
  }

  function syncTableToState() {
    const rows = [...el.manageTbody.querySelectorAll("tr")];
    state.students = rows.map((tr) => {
      const get = (field) => {
        const input = tr.querySelector(`input[data-field="${field}"]`);
        return input ? input.value : "";
      };
      return normalizeStudentRow({
        nama: get("nama"),
        jantina: get("jantina"),
        kelas: get("kelas"),
        no_kad_pengenalan: get("no_kad_pengenalan"),
        email_google_classroom: get("email_google_classroom"),
      });
    });
  }

  function startImport() {
    const confirmed = window.confirm(
      "Warning: This will overwrites the current namelist. Continue?"
    );
    if (!confirmed) {
      return;
    }
    el.fileInput.value = "";
    el.fileInput.click();
  }

  async function handleImportFile(event) {
    const [file] = event.target.files || [];
    if (!file) {
      return;
    }

    try {
      const text = await file.text();
      const rows = parseCsv(text);
      const namaColumn = resolveColumnName(rows, ["nama murid", "nama"]);
      const jantinaColumn = resolveColumnName(rows, ["jantina"]);
      const kelasColumn = resolveColumnName(rows, ["kelas 2026", "kelas"]);
      const noKadColumn = resolveColumnName(rows, ["no. kad pengenalan", "no kad pengenalan"]);
      const emailColumn = resolveColumnName(rows, ["email id google classroom", "email google classroom", "email"]);

      if (!namaColumn || !kelasColumn || !noKadColumn) {
        throw new Error(
          "Kolum CSV tidak sah. Wajib ada NAMA MURID, Kelas 2026/Kelas, dan No. Kad Pengenalan."
        );
      }

      const existingByNoKad = new Map(
        getCurrentNamelist()
          .map(normalizeStudentRow)
          .filter((row) => row.no_kad_pengenalan)
          .map((row) => [row.no_kad_pengenalan, row])
      );

      const imported = rows
        .map((row) =>
          normalizeStudentRow({
            nama: row[namaColumn],
            jantina: jantinaColumn ? row[jantinaColumn] : "",
            kelas: row[kelasColumn],
            no_kad_pengenalan: noKadColumn ? row[noKadColumn] : "",
            email_google_classroom: emailColumn ? row[emailColumn] : "",
          })
        )
        .filter((row) => row.nama && row.kelas && row.no_kad_pengenalan);

      const byNoKad = new Map();
      imported.forEach((row) => {
        const existing = existingByNoKad.get(row.no_kad_pengenalan) || {};
        byNoKad.set(row.no_kad_pengenalan, {
          nama: row.nama || existing.nama || "",
          jantina: row.jantina || existing.jantina || "",
          kelas: row.kelas || existing.kelas || "",
          no_kad_pengenalan: row.no_kad_pengenalan,
          email_google_classroom: row.email_google_classroom || existing.email_google_classroom || "",
        });
      });

      const namelist = [...byNoKad.values()].sort((a, b) => {
        const byClass = a.kelas.localeCompare(b.kelas, "ms");
        if (byClass !== 0) {
          return byClass;
        }
        return a.nama.localeCompare(b.nama, "ms");
      });

      if (!namelist.length) {
        throw new Error("Tiada rekod sah dijumpai. Pastikan setiap murid ada No. Kad Pengenalan.");
      }

      localStorage.setItem(NAMELIST_OVERRIDE_KEY, JSON.stringify(namelist));
      state.students = namelist;
      if (state.isManageOpen) {
        renderManageTable();
      }
      await upsertStudentsToSupabase(namelist);
      setStatus(`Namelist berjaya diimport: ${namelist.length} murid.`);
    } catch (error) {
      console.error(error);
      setStatus(error.message || "Import namelist gagal.", true);
    }
  }

  function startImportData() {
    openImportDataDialog();
  }

  function initImportMonthOptions() {
    if (!el.importMonthSelect) {
      return;
    }
    el.importMonthSelect.innerHTML = "";
    MONTHS.forEach((month) => {
      const option = document.createElement("option");
      option.value = month;
      option.textContent = month;
      el.importMonthSelect.appendChild(option);
    });
  }

  function openImportDataDialog() {
    if (!el.importDataModal) {
      return;
    }
    const currentYear = String(new Date().getFullYear());
    if (el.importYearInput) {
      el.importYearInput.value = currentYear;
    }
    if (el.importMonthSelect) {
      el.importMonthSelect.value = MONTHS[new Date().getMonth()];
    }
    el.importDataModal.hidden = false;
  }

  function closeImportDataDialog() {
    if (el.importDataModal) {
      el.importDataModal.hidden = true;
    }
  }

  function confirmImportDataSelection() {
    const year = String(el.importYearInput ? el.importYearInput.value : "").trim();
    if (!/^\d{4}$/.test(year)) {
      setStatus("Tahun tidak sah. Gunakan format 4 digit, contoh 2026.", true);
      return;
    }

    const month = normalizeMonth(el.importMonthSelect ? el.importMonthSelect.value : "");
    if (!month) {
      setStatus("Bulan tidak sah. Sila pilih bulan daripada senarai.", true);
      return;
    }

    state.pendingImportYear = year;
    state.pendingImportMonth = month;
    closeImportDataDialog();
    el.dataFileInput.value = "";
    el.dataFileInput.click();
  }

  async function handleImportDataFile(event) {
    const [file] = event.target.files || [];
    if (!file) {
      return;
    }

    try {
      const year = state.pendingImportYear;
      const month = state.pendingImportMonth;
      if (!year || !month) {
        throw new Error("Tahun/Bulan import tiada. Sila klik butang Import Data semula.");
      }

      const text = await file.text();
      const rows = parseCsv(text);
      const parsed = parseImportedDataRows(rows, year, month);
      if (!parsed.records.length) {
        throw new Error("Tiada rekod data sah dijumpai dalam CSV.");
      }

      const mergedRecords = mergeRecordsByNoKad(year, month, parsed.records);
      persistLocalRecords(mergedRecords);
      mergeStudentsIntoNamelist(parsed.students);
      await upsertImportedRecordsToSupabase(parsed.records);
      await upsertStudentsToSupabase(parsed.students);

      setStatus(
        `Import data berjaya: ${parsed.records.length} rekod diproses untuk ${year} ${month}. Data sedia ada telah di-override ikut No. Kad Pengenalan.`
      );
    } catch (error) {
      console.error(error);
      setStatus(error.message || "Import data gagal.", true);
    } finally {
      state.pendingImportYear = "";
      state.pendingImportMonth = "";
      event.target.value = "";
    }
  }

  async function resetAllData() {
    const confirmed = window.confirm(
      "Warning: This will deletes all the data, not namelist. Continue?"
    );
    if (!confirmed) {
      return;
    }

    let localDeletedCount = 0;
    const localKeys = [];
    for (let i = 0; i < localStorage.length; i += 1) {
      const key = localStorage.key(i);
      if (key && key.startsWith("nilam_records_")) {
        localKeys.push(key);
      }
    }
    localKeys.forEach((key) => {
      localStorage.removeItem(key);
      localDeletedCount += 1;
    });

    let supabaseMessage = "Supabase tidak dikonfigurasi.";
    const config = window.NILAM_CONFIG || {};
    if (config.supabaseUrl && config.supabaseAnonKey) {
      try {
        await deleteAllSupabaseData(config);
        supabaseMessage = "Data Supabase berjaya dipadam.";
      } catch (error) {
        console.error(error);
        supabaseMessage = "Padam data Supabase gagal.";
      }
    }

    setStatus(`Reset selesai. LocalStorage dipadam: ${localDeletedCount} set data. ${supabaseMessage}`);
  }

  async function deleteAllSupabaseData(config) {
    const supabaseUrl = config.supabaseUrl.replace(/\/$/, "");
    const monthFilter = MONTHS.map((month) => `bulan.eq.${encodeURIComponent(month)}`).join(",");
    const endpoint = `${supabaseUrl}/rest/v1/nilam_records?or=(${monthFilter})`;

    const response = await fetch(endpoint, {
      method: "DELETE",
      headers: {
        apikey: config.supabaseAnonKey,
        Authorization: `Bearer ${config.supabaseAnonKey}`,
        Prefer: "return=minimal",
      },
    });

    if (!response.ok) {
      const detail = await response.text();
      throw new Error(`Ralat padam Supabase (${response.status}): ${detail}`);
    }
  }

  function normalizeStudentRow(row) {
    return {
      nama: String(row.nama || "").trim(),
      jantina: String(row.jantina || "").trim(),
      kelas: String(row.kelas || "").trim(),
      no_kad_pengenalan: normalizeNoKad(row.no_kad_pengenalan),
      email_google_classroom: String(row.email_google_classroom || "").trim(),
    };
  }

  function emptyStudentRow() {
    return {
      nama: "",
      jantina: "",
      kelas: "",
      no_kad_pengenalan: "",
      email_google_classroom: "",
    };
  }

  function normalizeMonth(value) {
    const normalized = String(value || "").trim().toLowerCase();
    const found = MONTHS.find((m) => m.toLowerCase() === normalized);
    return found || "";
  }

  function parseImportedDataRows(rows, year, month) {
    const namaColumn = resolveColumnName(rows, ["nama murid", "nama"]);
    const jantinaColumn = resolveColumnName(rows, ["jantina"]);
    const noKadColumn = resolveColumnName(rows, ["no. kad pengenalan", "no kad pengenalan"]);
    const kelasColumn = resolveColumnName(rows, ["kelas 2026", "kelas"]);
    const emailColumn = resolveColumnName(rows, ["email id google classroom", "email google classroom", "email"]);
    const bahanDigitalColumn = resolveColumnName(rows, ["bahan digital"]);
    const bahanBukanBukuColumn = resolveColumnName(rows, ["bahan bukan buku"]);
    const fiksyenColumn = resolveColumnName(rows, ["fiksyen"]);
    const bukanFiksyenColumn = resolveColumnName(rows, ["bukan fiksyen"]);
    const bmColumn = resolveColumnName(rows, ["bahasa melayu"]);
    const biColumn = resolveColumnName(rows, ["bahasa inggeris"]);
    const lainColumn = resolveColumnName(rows, ["lain-lain bahasa", "lain lain bahasa"]);

    if (
      !namaColumn ||
      !noKadColumn ||
      !kelasColumn ||
      !bahanDigitalColumn ||
      !bahanBukanBukuColumn ||
      !fiksyenColumn ||
      !bukanFiksyenColumn ||
      !bmColumn ||
      !biColumn ||
      !lainColumn
    ) {
      throw new Error(
        "Kolum data CSV tidak lengkap. Perlu ada: NAMA MURID, No. Kad Pengenalan, Kelas, Bahan Digital, Bahan Bukan Buku, Fiksyen, Bukan Fiksyen, Bahasa Melayu, Bahasa Inggeris, Lain-lain Bahasa."
      );
    }

    const records = [];
    const students = [];
    rows.forEach((row) => {
      const nama = String(row[namaColumn] || "").trim();
      const noKad = normalizeNoKad(row[noKadColumn]);
      const kelas = String(row[kelasColumn] || "").trim();
      if (!nama || !noKad || !kelas) {
        return;
      }

      const bahanDigital = toInt999(row[bahanDigitalColumn]);
      const bahanBukanBuku = toInt999(row[bahanBukanBukuColumn]);
      const fiksyen = toInt999(row[fiksyenColumn]);
      const bukanFiksyen = toInt999(row[bukanFiksyenColumn]);
      const bahasaMelayu = toInt999(row[bmColumn]);
      const bahasaInggeris = toInt999(row[biColumn]);
      const lainLainBahasa = toInt999(row[lainColumn]);
      const jumlahBacaan = bahasaMelayu + bahasaInggeris + lainLainBahasa;

      records.push({
        tahun: year,
        bulan: month,
        bil: 0,
        no_kad_pengenalan: noKad,
        nama,
        kelas,
        bahan_digital: bahanDigital,
        bahan_bukan_buku: bahanBukanBuku,
        fiksyen,
        bukan_fiksyen: bukanFiksyen,
        bahasa_melayu: bahasaMelayu,
        bahasa_inggeris: bahasaInggeris,
        lain_lain_bahasa: lainLainBahasa,
        jumlah_aktiviti: jumlahBacaan,
        updated_at_client: new Date().toISOString(),
      });

      students.push(
        normalizeStudentRow({
          nama,
          jantina: jantinaColumn ? row[jantinaColumn] : "",
          kelas,
          no_kad_pengenalan: noKad,
          email_google_classroom: emailColumn ? row[emailColumn] : "",
        })
      );
    });

    return {
      records,
      students: dedupeStudentsByNoKad(students),
    };
  }

  function dedupeStudentsByNoKad(students) {
    const map = new Map();
    students.forEach((row) => {
      if (!row.no_kad_pengenalan) {
        return;
      }
      const existing = map.get(row.no_kad_pengenalan);
      if (!existing) {
        map.set(row.no_kad_pengenalan, row);
        return;
      }
      map.set(row.no_kad_pengenalan, {
        ...existing,
        nama: row.nama || existing.nama,
        jantina: row.jantina || existing.jantina,
        kelas: row.kelas || existing.kelas,
        email_google_classroom: row.email_google_classroom || existing.email_google_classroom,
      });
    });
    return [...map.values()];
  }

  function mergeRecordsByNoKad(year, month, importedRecords) {
    const existing = readAllLocalRecords();
    const map = new Map();

    existing.forEach((row) => {
      const key = `${row.tahun}|${row.bulan}|${normalizeNoKad(row.no_kad_pengenalan)}`;
      map.set(key, row);
    });
    importedRecords.forEach((row) => {
      const key = `${row.tahun}|${row.bulan}|${normalizeNoKad(row.no_kad_pengenalan)}`;
      map.set(key, row);
    });

    const merged = [...map.values()];
    const target = merged.filter((r) => r.tahun === year && r.bulan === month);
    const grouped = new Map();
    target.forEach((row) => {
      if (!grouped.has(row.kelas)) {
        grouped.set(row.kelas, []);
      }
      grouped.get(row.kelas).push(row);
    });
    for (const rows of grouped.values()) {
      rows.sort((a, b) => a.nama.localeCompare(b.nama, "ms"));
      rows.forEach((row, idx) => {
        row.bil = idx + 1;
      });
    }
    return merged;
  }

  function readAllLocalRecords() {
    const all = [];
    for (let i = 0; i < localStorage.length; i += 1) {
      const key = localStorage.key(i);
      if (!key || !key.startsWith("nilam_records_")) {
        continue;
      }
      const raw = localStorage.getItem(key);
      if (!raw) {
        continue;
      }
      try {
        const parsed = JSON.parse(raw);
        if (Array.isArray(parsed)) {
          parsed.forEach((row) => {
            const normalized = normalizeImportedRecord(row, key);
            if (normalized) {
              all.push(normalized);
            }
          });
        }
      } catch (error) {
        console.error(error);
      }
    }
    return all;
  }

  function normalizeImportedRecord(row, keyHint) {
    const tahun = String(row.tahun || extractYearFromKey(keyHint) || "").trim();
    const bulan = String(row.bulan || extractMonthFromKey(keyHint) || "").trim();
    const noKad = normalizeNoKad(row.no_kad_pengenalan);
    const nama = String(row.nama || "").trim();
    const kelas = String(row.kelas || "").trim();
    if (!tahun || !bulan || !noKad || !nama || !kelas) {
      return null;
    }

    const bm = toInt999(row.bahasa_melayu);
    const bi = toInt999(row.bahasa_inggeris);
    const lain = toInt999(row.lain_lain_bahasa);
    return {
      tahun,
      bulan,
      bil: Number(row.bil) > 0 ? Number(row.bil) : 0,
      no_kad_pengenalan: noKad,
      nama,
      kelas,
      bahan_digital: toInt999(row.bahan_digital),
      bahan_bukan_buku: toInt999(row.bahan_bukan_buku),
      fiksyen: toInt999(row.fiksyen),
      bukan_fiksyen: toInt999(row.bukan_fiksyen),
      bahasa_melayu: bm,
      bahasa_inggeris: bi,
      lain_lain_bahasa: lain,
      jumlah_aktiviti:
        Number.isFinite(Number(row.jumlah_aktiviti)) && Number(row.jumlah_aktiviti) >= 0
          ? Math.trunc(Number(row.jumlah_aktiviti))
          : bm + bi + lain,
      updated_at_client: row.updated_at_client || new Date().toISOString(),
    };
  }

  function extractYearFromKey(key) {
    const match = String(key || "").match(/^nilam_records_(\d{4})_/);
    return match ? match[1] : "";
  }

  function extractMonthFromKey(key) {
    const value = String(key || "");
    const yearPrefix = value.match(/^nilam_records_\d{4}_(.+?)_/);
    if (yearPrefix) {
      return yearPrefix[1];
    }
    const legacy = value.match(/^nilam_records_(.+?)_/);
    return legacy ? legacy[1] : "";
  }

  function persistLocalRecords(records) {
    const keysToDelete = [];
    for (let i = 0; i < localStorage.length; i += 1) {
      const key = localStorage.key(i);
      if (key && key.startsWith("nilam_records_")) {
        keysToDelete.push(key);
      }
    }
    keysToDelete.forEach((key) => localStorage.removeItem(key));

    const grouped = new Map();
    records.forEach((row) => {
      const key = `nilam_records_${row.tahun}_${row.bulan}_${row.kelas}`;
      if (!grouped.has(key)) {
        grouped.set(key, []);
      }
      grouped.get(key).push({ ...row });
    });

    for (const [key, rows] of grouped.entries()) {
      rows.sort((a, b) => {
        const bilDiff = Number(a.bil || 0) - Number(b.bil || 0);
        if (bilDiff !== 0) {
          return bilDiff;
        }
        return String(a.nama || "").localeCompare(String(b.nama || ""), "ms");
      });
      rows.forEach((row, idx) => {
        if (!Number(row.bil) || Number(row.bil) <= 0) {
          row.bil = idx + 1;
        }
      });
      localStorage.setItem(key, JSON.stringify(rows));
    }
  }

  function mergeStudentsIntoNamelist(importedStudents) {
    const current = getCurrentNamelist().map(normalizeStudentRow);
    const byNoKad = new Map();
    current.forEach((row) => {
      if (row.no_kad_pengenalan) {
        byNoKad.set(row.no_kad_pengenalan, row);
      }
    });

    importedStudents.forEach((row) => {
      if (!row.no_kad_pengenalan) {
        return;
      }
      const existing = byNoKad.get(row.no_kad_pengenalan) || {};
      byNoKad.set(row.no_kad_pengenalan, {
        ...existing,
        nama: row.nama || existing.nama || "",
        jantina: row.jantina || existing.jantina || "",
        kelas: row.kelas || existing.kelas || "",
        no_kad_pengenalan: row.no_kad_pengenalan,
        email_google_classroom: row.email_google_classroom || existing.email_google_classroom || "",
      });
    });

    const merged = [...byNoKad.values()].sort((a, b) => {
      const byClass = a.kelas.localeCompare(b.kelas, "ms");
      if (byClass !== 0) {
        return byClass;
      }
      return a.nama.localeCompare(b.nama, "ms");
    });

    if (merged.length) {
      localStorage.setItem(NAMELIST_OVERRIDE_KEY, JSON.stringify(merged));
      state.students = merged;
      if (state.isManageOpen) {
        renderManageTable();
      }
    }
  }

  async function upsertImportedRecordsToSupabase(records) {
    const config = window.NILAM_CONFIG || {};
    if (!config.supabaseUrl || !config.supabaseAnonKey || !records.length) {
      return;
    }

    const supabaseUrl = config.supabaseUrl.replace(/\/$/, "");
    const endpoint = `${supabaseUrl}/rest/v1/nilam_records?on_conflict=tahun,bulan,no_kad_pengenalan`;
    const response = await fetch(endpoint, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        apikey: config.supabaseAnonKey,
        Authorization: `Bearer ${config.supabaseAnonKey}`,
        Prefer: "resolution=merge-duplicates,return=minimal",
      },
      body: JSON.stringify(records),
    });

    if (!response.ok) {
      const detail = await response.text();
      throw new Error(`Sync import data ke Supabase gagal (${response.status}): ${detail}`);
    }
  }

  function toInt999(value) {
    const number = Number(value);
    if (!Number.isFinite(number)) {
      return 0;
    }
    return Math.max(0, Math.min(999, Math.trunc(number)));
  }

  async function upsertStudentsToSupabase(students) {
    const config = window.NILAM_CONFIG || {};
    if (!config.supabaseUrl || !config.supabaseAnonKey) {
      return;
    }

    const year = String(new Date().getFullYear());
    const kelasField = `kelas_${year}`;
    const payload = students
      .filter((row) => row.no_kad_pengenalan)
      .map((row) => ({
        no_kad_pengenalan: row.no_kad_pengenalan,
        nama_murid: row.nama,
        jantina: row.jantina,
        email_google_classroom: row.email_google_classroom,
        [kelasField]: row.kelas,
      }));

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
      throw new Error(`Sync senarai murid ke Supabase gagal (${response.status}): ${detail}`);
    }
  }

  function getCurrentNamelist() {
    const overrideRaw = localStorage.getItem(NAMELIST_OVERRIDE_KEY);
    if (overrideRaw) {
      try {
        const parsed = JSON.parse(overrideRaw);
        if (Array.isArray(parsed)) {
          return parsed;
        }
      } catch (error) {
        console.error(error);
      }
    }
    return Array.isArray(window.NILAM_STUDENTS) ? window.NILAM_STUDENTS : [];
  }

  function normalizeNoKad(value) {
    return String(value || "")
      .replace(/\s+/g, "")
      .toUpperCase()
      .trim();
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

  function normalizeHeader(value) {
    return String(value || "")
      .replace(/^\uFEFF/, "")
      .toLowerCase()
      .replace(/\s+/g, " ")
      .trim();
  }

  function escapeAttr(value) {
    return String(value || "")
      .replaceAll("&", "&amp;")
      .replaceAll('"', "&quot;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;");
  }

  function setStatus(message, isError) {
    el.status.textContent = message;
    el.status.style.color = isError ? "#b00020" : "";
  }
})();
