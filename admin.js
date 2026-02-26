(function () {
  "use strict";

  const AUTH_USER = "admin";
  const AUTH_PASS = "nilam_admin";
  const AUTH_SESSION_KEY = "nilam_admin_auth_v1";
  const NAMELIST_OVERRIDE_KEY = "nilam_students_override_v1";
  const TEACHER_NAMES_KEY = "nilam_teacher_names_v1";
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
    isPanelHiddenInManage: false,
    currentPage: 1,
    rowsPerPage: 20,
    pendingImportYear: "",
    pendingImportMonth: "",
    pendingImportType: "",
    undoStack: [],
    searchQuery: "",
    pendingCompareResult: null,
    autoSyncTimer: null,
    autoSyncInFlight: false,
    autoSyncQueued: false,
    autoSyncRequireDeactivate: false,
  };

  const el = {
    loginInfo: document.getElementById("adminLoginInfo"),
    panel: document.getElementById("adminPanel"),
    loginBtn: document.getElementById("loginBtn"),
    loginModal: document.getElementById("loginModal"),
    usernameInput: document.getElementById("adminUsernameInput"),
    passwordInput: document.getElementById("adminPasswordInput"),
    submitLoginBtn: document.getElementById("submitLoginBtn"),
    cancelLoginBtn: document.getElementById("cancelLoginBtn"),
    importBtn: document.getElementById("importNamelistBtn"),
    fileInput: document.getElementById("namelistFileInput"),
    importTeachersBtn: document.getElementById("importTeachersBtn"),
    teachersFileInput: document.getElementById("teachersFileInput"),
    importDataBtn: document.getElementById("importDataBtn"),
    dataFileInput: document.getElementById("dataFileInput"),
    importAinsBtn: document.getElementById("importAinsBtn"),
    ainsFileInput: document.getElementById("ainsFileInput"),
    importDataModal: document.getElementById("importDataModal"),
    importYearInput: document.getElementById("importYearInput"),
    importMonthSelect: document.getElementById("importMonthSelect"),
    confirmImportDataBtn: document.getElementById("confirmImportDataBtn"),
    cancelImportDataBtn: document.getElementById("cancelImportDataBtn"),
    resetBtn: document.getElementById("resetDataBtn"),
    manageBtn: document.getElementById("manageStudentsBtn"),
    manageSection: document.getElementById("studentManageSection"),
    toggleAdminPanelBtn: document.getElementById("toggleAdminPanelBtn"),
    rowsPerPageSelect: document.getElementById("manageRowsPerPage"),
    prevPageBtn: document.getElementById("managePrevPageBtn"),
    nextPageBtn: document.getElementById("manageNextPageBtn"),
    pageInput: document.getElementById("managePageInput"),
    pageTotal: document.getElementById("managePageTotal"),
    manageTbody: document.getElementById("studentManageTbody"),
    addRowBtn: document.getElementById("addStudentRowBtn"),
    saveStudentsBtn: document.getElementById("saveStudentsBtn"),
    undoRemoveBtn: document.getElementById("undoRemoveBtn"),
    searchInput: document.getElementById("manageSearchInput"),
    compareBtn: document.getElementById("compareNamelistBtn"),
    compareFileInput: document.getElementById("compareFileInput"),
    compareModal: document.getElementById("compareModal"),
    compareContent: document.getElementById("compareContent"),
    cancelCompareBtn: document.getElementById("cancelCompareBtn"),
    confirmCompareBtn: document.getElementById("confirmCompareBtn"),
    logoutBtn: document.getElementById("logoutBtn"),
    status: document.getElementById("adminStatus"),
  };

  init();

  function init() {
    initImportMonthOptions();
    if (el.rowsPerPageSelect) {
      el.rowsPerPageSelect.value = String(state.rowsPerPage);
    }
    updateAdminPanelToggleLabel();
    bindEvents();
    initInfoBadges();
    if (sessionStorage.getItem(AUTH_SESSION_KEY) === "1") {
      showPanel();
    } else {
      requestLogin();
    }
  }

  function bindEvents() {
    el.loginBtn.addEventListener("click", requestLogin);
    if (el.submitLoginBtn) {
      el.submitLoginBtn.addEventListener("click", submitLogin);
    }
    if (el.cancelLoginBtn) {
      el.cancelLoginBtn.addEventListener("click", cancelLogin);
    }
    if (el.loginModal) {
      el.loginModal.addEventListener("click", handleLoginModalClick);
    }
    if (el.usernameInput) {
      el.usernameInput.addEventListener("keydown", handleLoginInputKeydown);
    }
    if (el.passwordInput) {
      el.passwordInput.addEventListener("keydown", handleLoginInputKeydown);
    }
    el.importBtn.addEventListener("click", startImport);
    el.fileInput.addEventListener("change", handleImportFile);
    if (el.importTeachersBtn) {
      el.importTeachersBtn.addEventListener("click", startImportTeachers);
    }
    if (el.teachersFileInput) {
      el.teachersFileInput.addEventListener("change", handleImportTeachersFile);
    }
    el.importDataBtn.addEventListener("click", startImportData);
    el.dataFileInput.addEventListener("change", handleImportDataFile);
    if (el.importAinsBtn) {
      el.importAinsBtn.addEventListener("click", startImportAins);
    }
    if (el.ainsFileInput) {
      el.ainsFileInput.addEventListener("change", handleImportAinsFile);
    }
    el.confirmImportDataBtn.addEventListener("click", confirmImportDataSelection);
    el.cancelImportDataBtn.addEventListener("click", closeImportDataDialog);
    el.resetBtn.addEventListener("click", resetAllData);
    el.manageBtn.addEventListener("click", toggleManageStudents);
    if (el.toggleAdminPanelBtn) {
      el.toggleAdminPanelBtn.addEventListener("click", toggleAdminPanelVisibility);
    }
    if (el.rowsPerPageSelect) {
      el.rowsPerPageSelect.addEventListener("change", handleRowsPerPageChange);
    }
    if (el.pageInput) {
      el.pageInput.addEventListener("change", commitPageInput);
      el.pageInput.addEventListener("keydown", handlePageInputKeydown);
    }
    if (el.prevPageBtn) {
      el.prevPageBtn.addEventListener("click", goToPrevPage);
    }
    if (el.nextPageBtn) {
      el.nextPageBtn.addEventListener("click", goToNextPage);
    }
    el.addRowBtn.addEventListener("click", addStudentRow);
    el.saveStudentsBtn.addEventListener("click", saveManagedStudents);
    el.manageTbody.addEventListener("click", handleManageTableClick);
    el.manageTbody.addEventListener("input", handleManageTableInput);
    el.manageTbody.addEventListener("change", handleManageTableChange);
    if (el.undoRemoveBtn) {
      el.undoRemoveBtn.addEventListener("click", undoRemove);
    }
    if (el.searchInput) {
      el.searchInput.addEventListener("input", handleSearchInput);
    }
    if (el.compareBtn) {
      el.compareBtn.addEventListener("click", startCompareNamelist);
    }
    if (el.compareFileInput) {
      el.compareFileInput.addEventListener("change", handleCompareFile);
    }
    if (el.cancelCompareBtn) {
      el.cancelCompareBtn.addEventListener("click", closeCompareModal);
    }
    if (el.confirmCompareBtn) {
      el.confirmCompareBtn.addEventListener("click", confirmCompare);
    }
    el.logoutBtn.addEventListener("click", logout);
  }

  function requestLogin() {
    showLoginOnly();
    openLoginDialog();
  }

  function openLoginDialog() {
    if (!el.loginModal) {
      return;
    }
    el.loginModal.hidden = false;
    if (el.usernameInput) {
      el.usernameInput.value = "";
      el.usernameInput.focus();
    }
    if (el.passwordInput) {
      el.passwordInput.value = "";
    }
  }

  function closeLoginDialog() {
    if (!el.loginModal) {
      return;
    }
    el.loginModal.hidden = true;
  }

  function cancelLogin() {
    closeLoginDialog();
    showLoginOnly();
    setStatus("Login dibatalkan.", true);
  }

  function submitLogin() {
    const username = el.usernameInput ? String(el.usernameInput.value || "").trim() : "";
    const password = el.passwordInput ? String(el.passwordInput.value || "") : "";
    if (username === AUTH_USER && password === AUTH_PASS) {
      closeLoginDialog();
      sessionStorage.setItem(AUTH_SESSION_KEY, "1");
      showPanel();
      setStatus("Login berjaya.");
      return;
    }
    if (el.passwordInput) {
      el.passwordInput.value = "";
      el.passwordInput.focus();
    }
    sessionStorage.removeItem(AUTH_SESSION_KEY);
    showLoginOnly();
    setStatus("Username atau password tidak sah.", true);
  }

  function handleLoginModalClick(event) {
    if (event.target !== el.loginModal) {
      return;
    }
    cancelLogin();
  }

  function handleLoginInputKeydown(event) {
    if (event.key === "Escape") {
      event.preventDefault();
      cancelLogin();
      return;
    }
    if (event.key !== "Enter") {
      return;
    }
    event.preventDefault();
    submitLogin();
  }

  function showPanel() {
    el.loginInfo.hidden = true;
    el.panel.hidden = false;
    if (el.logoutBtn) {
      el.logoutBtn.hidden = false;
    }
    state.isPanelHiddenInManage = false;
    updateAdminPanelToggleLabel();
    loadStudentsForManage();
  }

  function showLoginOnly() {
    el.loginInfo.hidden = false;
    el.panel.hidden = true;
    if (el.logoutBtn) {
      el.logoutBtn.hidden = true;
    }
    if (el.manageSection) {
      el.manageSection.hidden = true;
    }
    state.isPanelHiddenInManage = false;
    updateAdminPanelToggleLabel();
    state.isManageOpen = false;
  }

  function initInfoBadges() {
    const badges = Array.from(document.querySelectorAll(".info-badge"));
    if (!badges.length) {
      return;
    }

    const closeAll = () => {
      badges.forEach((badge) => {
        badge.classList.remove("is-open");
        badge.setAttribute("aria-expanded", "false");
      });
    };

    badges.forEach((badge) => {
      badge.addEventListener("click", (event) => {
        event.stopPropagation();
        const isOpen = badge.classList.contains("is-open");
        closeAll();
        if (!isOpen) {
          badge.classList.add("is-open");
          badge.setAttribute("aria-expanded", "true");
        }
      });
      badge.addEventListener("keydown", (event) => {
        if (event.key !== "Enter" && event.key !== " ") {
          return;
        }
        event.preventDefault();
        badge.click();
      });
    });

    document.addEventListener("click", closeAll);
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
      setAdminPanelHidden(true);
      loadStudentsForManage();
      setStatus("Pengurusan senarai murid dibuka.");
      return;
    }
    setAdminPanelHidden(false);
  }

  function loadStudentsForManage() {
    const raw = getCurrentNamelist();
    state.students = raw.map(normalizeStudentRow).filter((row) => row.nama || row.kelas);
    if (!state.students.length) {
      state.students = [emptyStudentRow()];
    }
    state.currentPage = 1;
    state.undoStack = [];
    state.searchQuery = "";
    if (el.searchInput) {
      el.searchInput.value = "";
    }
    updateUndoBtn();
    renderManageTable();
  }

  function renderManageTable() {
    const query = state.searchQuery.toLowerCase();
    const indexed = state.students.map((row, i) => ({ row, originalIndex: i }));
    const filtered = query
      ? indexed.filter(({ row }) =>
          row.nama.toLowerCase().includes(query) ||
          row.no_kad_pengenalan.toLowerCase().includes(query)
        )
      : indexed;

    const totalRows = filtered.length;
    const totalPages = Math.max(1, Math.ceil(totalRows / state.rowsPerPage));
    state.currentPage = Math.min(Math.max(1, state.currentPage), totalPages);
    const start = (state.currentPage - 1) * state.rowsPerPage;
    const end = start + state.rowsPerPage;
    const visibleRows = filtered.slice(start, end);

    const rowsHtml = visibleRows
      .map(
        ({ row, originalIndex }) => `
        <tr>
          <td><input type="text" data-field="nama" data-index="${originalIndex}" value="${escapeAttr(row.nama)}"></td>
          <td><input type="text" data-field="jantina" data-index="${originalIndex}" value="${escapeAttr(row.jantina)}"></td>
          <td><input type="text" data-field="kelas" data-index="${originalIndex}" value="${escapeAttr(row.kelas)}"></td>
          <td><input type="text" data-field="no_kad_pengenalan" data-index="${originalIndex}" value="${escapeAttr(row.no_kad_pengenalan.startsWith("NOIC_") ? "" : row.no_kad_pengenalan)}"></td>
          <td><input type="text" data-field="email_google_classroom" data-index="${originalIndex}" value="${escapeAttr(row.email_google_classroom)}"></td>
          <td><button type="button" class="mini-btn danger-btn" data-remove-index="${originalIndex}">Buang</button></td>
        </tr>
      `
      )
      .join("");

    el.manageTbody.innerHTML = rowsHtml;
    if (el.pageInput) {
      el.pageInput.value = String(state.currentPage);
      el.pageInput.max = String(totalPages);
    }
    if (el.pageTotal) {
      el.pageTotal.textContent = String(totalPages);
    }
    if (el.prevPageBtn) {
      el.prevPageBtn.disabled = state.currentPage <= 1;
    }
    if (el.nextPageBtn) {
      el.nextPageBtn.disabled = state.currentPage >= totalPages;
    }
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

    const student = state.students[index];
    const displayName = [student.nama, student.kelas ? `(${student.kelas})` : ""].filter(Boolean).join(" ");
    if (!window.confirm(`Buang murid ini?\n\n${displayName || "(tanpa nama)"}`)) {
      return;
    }

    state.undoStack.push({ index, row: { ...state.students[index] } });
    state.students.splice(index, 1);
    if (!state.students.length) {
      state.students.push(emptyStudentRow());
    }
    const totalPages = Math.max(1, Math.ceil(state.students.length / state.rowsPerPage));
    state.currentPage = Math.min(state.currentPage, totalPages);
    renderManageTable();
    updateUndoBtn();
    queueAutoSync({ deactivateMissing: true });
  }

  function handleManageTableInput(event) {
    const input = event.target.closest("input[data-field][data-index]");
    if (!input) {
      return;
    }
    const index = Number(input.dataset.index);
    const field = input.dataset.field;
    if (!Number.isInteger(index) || index < 0 || index >= state.students.length || !field) {
      return;
    }
    state.students[index][field] = input.value;
    queueAutoSync({ deactivateMissing: false });
  }

  function handleManageTableChange(event) {
    const input = event.target.closest("input[data-field][data-index]");
    if (!input) {
      return;
    }
    if (input.dataset.field === "kelas") {
      syncCurrentPageToState();
      sortStudentsByClassThenName();
      state.currentPage = 1;
      renderManageTable();
    }
    queueAutoSync({ deactivateMissing: false });
  }

  function sortStudentsByClassThenName() {
    state.students.sort((a, b) => {
      const kelasA = String(a.kelas || "").trim();
      const kelasB = String(b.kelas || "").trim();
      const byClass = kelasA.localeCompare(kelasB, "ms", { numeric: true, sensitivity: "base" });
      if (byClass !== 0) {
        return byClass;
      }
      const namaA = String(a.nama || "").trim();
      const namaB = String(b.nama || "").trim();
      return namaA.localeCompare(namaB, "ms", { sensitivity: "base" });
    });
  }

  function handleRowsPerPageChange() {
    syncCurrentPageToState();
    const next = Number(el.rowsPerPageSelect ? el.rowsPerPageSelect.value : state.rowsPerPage);
    if (next !== 20 && next !== 50 && next !== 100) {
      return;
    }
    state.rowsPerPage = next;
    state.currentPage = 1;
    renderManageTable();
  }

  function goToPrevPage() {
    if (state.currentPage <= 1) {
      return;
    }
    syncCurrentPageToState();
    state.currentPage -= 1;
    renderManageTable();
  }

  function goToNextPage() {
    const totalPages = Math.max(1, Math.ceil(state.students.length / state.rowsPerPage));
    if (state.currentPage >= totalPages) {
      return;
    }
    syncCurrentPageToState();
    state.currentPage += 1;
    renderManageTable();
  }

  function handlePageInputKeydown(event) {
    if (event.key !== "Enter") {
      return;
    }
    event.preventDefault();
    commitPageInput();
  }

  function commitPageInput() {
    const raw = el.pageInput ? el.pageInput.value : "";
    const targetPage = Number(raw);
    const totalPages = Math.max(1, Math.ceil(state.students.length / state.rowsPerPage));
    if (!Number.isInteger(targetPage)) {
      renderManageTable();
      return;
    }
    syncCurrentPageToState();
    state.currentPage = Math.min(Math.max(1, targetPage), totalPages);
    renderManageTable();
  }

  function toggleAdminPanelVisibility() {
    setAdminPanelHidden(!state.isPanelHiddenInManage);
  }

  function setAdminPanelHidden(isHidden) {
    state.isPanelHiddenInManage = Boolean(isHidden);
    el.panel.hidden = state.isPanelHiddenInManage;
    updateAdminPanelToggleLabel();
  }

  function updateAdminPanelToggleLabel() {
    if (!el.toggleAdminPanelBtn) {
      return;
    }
    el.toggleAdminPanelBtn.textContent = state.isPanelHiddenInManage
      ? "Tunjuk Panel Admin"
      : "Sembunyi Panel Admin";
  }

  function addStudentRow() {
    syncCurrentPageToState();
    state.students.push(emptyStudentRow());
    state.currentPage = Math.max(1, Math.ceil(state.students.length / state.rowsPerPage));
    renderManageTable();
    queueAutoSync({ deactivateMissing: false });
  }

  async function saveManagedStudents() {
    try {
      syncCurrentPageToState();
      sortStudentsByClassThenName();

      const noKadSet = new Set();
      const toSave = state.students
        .map(normalizeStudentRow)
        .filter((row) => {
          if (!row.nama || !row.kelas) {
            return false;
          }
          // Skip real-IC duplicate check; synthetic ICs (NOIC_) are generated later.
          if (row.no_kad_pengenalan && !row.no_kad_pengenalan.startsWith("NOIC_")) {
            if (noKadSet.has(row.no_kad_pengenalan)) {
              throw new Error(`No. Kad Pengenalan berulang dikesan: ${row.no_kad_pengenalan}`);
            }
            noKadSet.add(row.no_kad_pengenalan);
          }
          return true;
        });

      if (!toSave.length) {
        throw new Error("Senarai murid kosong. Tambah sekurang-kurangnya 1 murid.");
      }

      localStorage.setItem(NAMELIST_OVERRIDE_KEY, JSON.stringify(toSave));
      state.students = toSave;
      state.undoStack = [];
      updateUndoBtn();
      renderManageTable();
      await upsertStudentsToSupabase(toSave);
      const purgeResult = await syncRecordsToMasterList(toSave);
      setStatus(
        `Senarai murid berjaya disimpan: ${toSave.length} murid. Data murid yang dibuang telah dipadam (Local: ${purgeResult.localDeleted}, Supabase: ${purgeResult.supabaseDeleted}).`
      );
      showPopupStatus("Berjaya disimpan", false);
    } catch (error) {
      console.error(error);
      const message = error.message || "Gagal simpan senarai murid.";
      setStatus(message, true);
      showPopupStatus(message, true);
    }
  }

  function syncTableToState() {
    syncCurrentPageToState();
  }

  function syncCurrentPageToState() {
    const inputs = [...el.manageTbody.querySelectorAll("input[data-field][data-index]")];
    inputs.forEach((input) => {
      const index = Number(input.dataset.index);
      const field = input.dataset.field;
      if (!Number.isInteger(index) || index < 0 || index >= state.students.length || !field) {
        return;
      }
      state.students[index][field] = input.value;
    });
  }

  function queueAutoSync(options = {}) {
    if (!state.isManageOpen) {
      return;
    }
    if (options.deactivateMissing) {
      state.autoSyncRequireDeactivate = true;
    }
    if (state.autoSyncTimer) {
      clearTimeout(state.autoSyncTimer);
    }
    state.autoSyncTimer = setTimeout(() => {
      state.autoSyncTimer = null;
      runAutoSync();
    }, 700);
  }

  async function runAutoSync() {
    if (state.autoSyncInFlight) {
      state.autoSyncQueued = true;
      return;
    }
    state.autoSyncInFlight = true;
    const deactivateMissing = state.autoSyncRequireDeactivate;
    state.autoSyncRequireDeactivate = false;

    try {
      syncCurrentPageToState();
      const noKadSet = new Set();
      const toSync = state.students
        .map(normalizeStudentRow)
        .filter((row) => {
          if (!row.nama || !row.kelas) {
            return false;
          }
          if (row.no_kad_pengenalan && !row.no_kad_pengenalan.startsWith("NOIC_")) {
            if (noKadSet.has(row.no_kad_pengenalan)) {
              throw new Error(`No. Kad Pengenalan berulang dikesan: ${row.no_kad_pengenalan}`);
            }
            noKadSet.add(row.no_kad_pengenalan);
          }
          return true;
        });

      if (!toSync.length) {
        throw new Error("Auto-sync gagal: Senarai murid kosong.");
      }

      localStorage.setItem(NAMELIST_OVERRIDE_KEY, JSON.stringify(toSync));
      await upsertStudentsToSupabase(toSync, { deactivateMissing });
      setStatus(`Auto-sync selesai (${toSync.length} murid).`);
    } catch (error) {
      console.error(error);
      setStatus(error.message || "Auto-sync gagal.", true);
    } finally {
      state.autoSyncInFlight = false;
      if (state.autoSyncQueued) {
        state.autoSyncQueued = false;
        runAutoSync();
      }
    }
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

  function startImportTeachers() {
    if (!el.teachersFileInput) {
      setStatus("Input fail untuk import Nama Guru tidak ditemui.", true);
      return;
    }
    el.teachersFileInput.value = "";
    el.teachersFileInput.click();
  }

  async function handleImportFile(event) {
    const [file] = event.target.files || [];
    if (!file) {
      return;
    }

    try {
      const rows = await readRowsFromFile(file);
      const namaColumn = resolveColumnName(rows, ["nama murid", "nama"]);
      const jantinaColumn = resolveColumnName(rows, ["jantina"]);
      const kelasColumn = resolveColumnName(rows, ["kelas 2026", "kelas"]);
      const noKadColumn = resolveColumnName(rows, ["no. kad pengenalan", "no kad pengenalan"]);
      const emailColumn = resolveColumnName(rows, ["email id google classroom", "email google classroom", "email"]);

      if (!namaColumn || !kelasColumn || !noKadColumn) {
        throw new Error(
          "Kolum fail tidak sah. Wajib ada NAMA MURID, Kelas 2026/Kelas, dan No. Kad Pengenalan."
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
      const purgeResult = await syncRecordsToMasterList(namelist);
      setStatus(
        `Namelist berjaya diimport: ${namelist.length} murid. Data murid yang dibuang telah dipadam (Local: ${purgeResult.localDeleted}, Supabase: ${purgeResult.supabaseDeleted}).`
      );
      showPopupStatus("Berjaya disimpan", false);
    } catch (error) {
      console.error(error);
      const message = error.message || "Import namelist gagal.";
      setStatus(message, true);
      showPopupStatus(message, true);
    } finally {
      event.target.value = "";
    }
  }

  async function handleImportTeachersFile(event) {
    const [file] = event.target.files || [];
    if (!file) {
      return;
    }

    try {
      const rows = await readRowsFromFile(file);
      const namaGuruColumn = resolveColumnName(rows, ["nama guru"]);
      if (!namaGuruColumn) {
        const availableHeaders = rows.length ? Object.keys(rows[0]).join(", ") : "(tiada header)";
        throw new Error(
          `Kolum fail tidak lengkap. Perlu ada header: Nama Guru. Header dikesan: ${availableHeaders}`
        );
      }

      const importedNames = rows
        .map((row) => String(row[namaGuruColumn] || "").trim())
        .filter(Boolean);
      if (!importedNames.length) {
        throw new Error("Tiada nama guru sah dijumpai.");
      }

      const existing = loadStoredTeacherNames();
      const merged = [...new Set([...existing, ...importedNames])].sort((a, b) =>
        a.localeCompare(b, "ms")
      );
      localStorage.setItem(TEACHER_NAMES_KEY, JSON.stringify(merged));

      setStatus(
        `Import Nama Guru berjaya: ${importedNames.length} nama diproses, jumlah senarai ${merged.length}.`
      );
      showPopupStatus("Berjaya disimpan", false);
    } catch (error) {
      console.error(error);
      const message = error.message || "Import Nama Guru gagal.";
      setStatus(message, true);
      showPopupStatus(message, true);
    } finally {
      event.target.value = "";
    }
  }

  function startImportData() {
    state.pendingImportType = "data";
    openImportDataDialog();
  }

  function startImportAins() {
    const now = new Date();
    state.pendingImportType = "ains";
    state.pendingImportYear = String(now.getFullYear());
    state.pendingImportMonth = MONTHS[now.getMonth()];
    if (!el.ainsFileInput) {
      setStatus("Input fail untuk import AINS tidak ditemui.", true);
      return;
    }
    el.ainsFileInput.value = "";
    el.ainsFileInput.click();
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
    if (state.pendingImportType === "ains") {
      if (!el.ainsFileInput) {
        setStatus("Input fail untuk import AINS tidak ditemui.", true);
        return;
      }
      el.ainsFileInput.value = "";
      el.ainsFileInput.click();
      return;
    }
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

      const rows = await readRowsFromFile(file);
      const parsed = parseImportedDataRows(rows, year, month);
      if (!parsed.records.length) {
        throw new Error("Tiada rekod data sah dijumpai dalam fail.");
      }

      const mergedRecords = mergeRecordsByNoKad(year, month, parsed.records);
      persistLocalRecords(mergedRecords);
      mergeStudentsIntoNamelist(parsed.students);
      let syncNote = "";
      try {
        await upsertStudentsToSupabase(parsed.students);
        await upsertImportedRecordsToSupabase(parsed.records);
        syncNote = " Data juga disimpan ke Supabase.";
      } catch (syncError) {
        console.error(syncError);
        const detail = String(syncError?.message || syncError || "").slice(0, 220);
        syncNote = ` Simpanan Supabase gagal (${detail}), tetapi data telah disimpan ke local.`;
      }

      setStatus(
        `Import data berjaya dan disimpan automatik: ${parsed.records.length} rekod diproses untuk ${year} ${month}. Sesi ADMIN_IMPORT bulan ini telah di-override.${syncNote}`
      );
      showPopupStatus("Berjaya disimpan", false);
    } catch (error) {
      console.error(error);
      const message = error.message || "Import data gagal.";
      setStatus(message, true);
      showPopupStatus(message, true);
    } finally {
      state.pendingImportYear = "";
      state.pendingImportMonth = "";
      state.pendingImportType = "";
      event.target.value = "";
    }
  }

  async function handleImportAinsFile(event) {
    const [file] = event.target.files || [];
    if (!file) {
      return;
    }

    try {
      const year = state.pendingImportYear;
      const month = state.pendingImportMonth;
      if (!year || !month) {
        throw new Error("Tahun/Bulan import AINS tiada. Sila klik butang Import Data AINS semula.");
      }

      const rows = await readRowsFromFile(file);
      const parsed = parseImportedAinsRows(rows, year, month);
      if (!parsed.records.length) {
        throw new Error(
          "Tiada padanan data AINS ditemui. Pastikan Nama dan ID DELIMa sama dengan namelist."
        );
      }

      const existingLocal = readAllLocalRecords();
      const clearedLocal = overwriteAinsForPeriod(existingLocal, year, month, parsed.records);
      const existingByKey = new Map();
      clearedLocal.forEach((row) => {
        const key = recordSessionKey(row);
        if (key) {
          existingByKey.set(key, row);
        }
      });
      const ainsOverrideRecords = parsed.records.map((row) =>
        buildRecordWithAinsOverride(row, existingByKey)
      );
      const mergedRecords = mergeRecordsByNoKad(year, month, ainsOverrideRecords, clearedLocal);
      persistLocalRecords(mergedRecords);

      let syncNote = "";
      try {
        await overwriteAinsInSupabase(year, month, parsed.records, ainsOverrideRecords);
        syncNote = " Data AINS juga disimpan ke Supabase.";
      } catch (syncError) {
        console.error(syncError);
        syncNote = " Simpanan Supabase gagal, tetapi data AINS telah disimpan ke local.";
      }

      const unmatchedNote = parsed.unmatchedCount
        ? ` ${parsed.unmatchedCount} baris tidak dipadankan (Nama + ID DELIMa).`
        : "";
      setStatus(
        `Import Data AINS berjaya: ${parsed.records.length} rekod untuk ${year} ${month} telah di-override.${unmatchedNote}${syncNote}`
      );
      showPopupStatus("Berjaya disimpan", false);
    } catch (error) {
      console.error(error);
      const message = error.message || "Import Data AINS gagal.";
      setStatus(message, true);
      showPopupStatus(message, true);
    } finally {
      state.pendingImportYear = "";
      state.pendingImportMonth = "";
      state.pendingImportType = "";
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
        supabaseMessage = `Padam data Supabase gagal: ${String(error.message || "ralat tidak diketahui")}`;
      }
    }

    setStatus(`Reset selesai. LocalStorage dipadam: ${localDeletedCount} set data. ${supabaseMessage}`);
  }

  async function deleteAllSupabaseData(config) {
    const supabaseUrl = config.supabaseUrl.replace(/\/$/, "");
    const headers = {
      apikey: config.supabaseAnonKey,
      Authorization: `Bearer ${config.supabaseAnonKey}`,
      Prefer: "return=minimal",
    };
    const endpoints = [
      `${supabaseUrl}/rest/v1/nilam_records?id=gt.0`,
      `${supabaseUrl}/rest/v1/nilam_records?bulan=not.is.null`,
    ];
    let lastError = "";

    for (const endpoint of endpoints) {
      const response = await fetch(endpoint, {
        method: "DELETE",
        headers,
      });
      if (response.ok) {
        return;
      }
      const detail = await response.text();
      lastError = `Ralat padam Supabase (${response.status}): ${detail}`;
    }

    throw new Error(lastError || "Ralat padam Supabase tidak diketahui.");
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

  function normalizeGuruType(value) {
    const normalized = String(value || "").trim();
    if (normalized === "BM" || normalized === "BI" || normalized === "Nilam") {
      return normalized;
    }
    return "Nilam";
  }

  function defaultTarikhForPeriod(year, month) {
    const safeYear = String(year || "").trim();
    const monthIndex = MONTHS.findIndex(
      (m) => m.toLowerCase() === String(month || "").trim().toLowerCase()
    );
    if (/^\d{4}$/.test(safeYear) && monthIndex >= 0) {
      return `${safeYear}-${String(monthIndex + 1).padStart(2, "0")}-01`;
    }
    return new Date().toISOString().slice(0, 10);
  }

  function normalizeTarikh(value, year, month) {
    const safe = String(value || "").trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(safe)) {
      return safe;
    }
    return defaultTarikhForPeriod(year, month);
  }

  function recordSessionKey(row) {
    const tahun = String(row?.tahun || "").trim();
    const bulan = String(row?.bulan || "").trim();
    const tarikh = normalizeTarikh(row?.tarikh, tahun, bulan);
    const noKad = normalizeNoKad(row?.no_kad_pengenalan);
    const namaPengisi = String(row?.nama_pengisi || "").trim().toLowerCase();
    const guru = normalizeGuruType(row?.guru);
    if (!tahun || !bulan || !tarikh || !noKad || !namaPengisi || !guru) {
      return "";
    }
    return `${tahun}|${bulan}|${tarikh}|${noKad}|${namaPengisi}|${guru}`;
  }

  function parseImportedDataRows(rows, year, month) {
    const namaColumn = resolveColumnName(rows, ["nama murid", "nama"]);
    const jantinaColumn = resolveColumnName(rows, ["jantina"]);
    const noKadColumn = resolveColumnName(rows, ["no. kad pengenalan", "no kad pengenalan"]);
    const kelasColumn = resolveColumnName(rows, ["kelas 2026", "kelas"]) || resolveDynamicKelasColumn(rows);
    const emailColumn = resolveColumnName(rows, ["email id google classroom", "email google classroom", "email"]);
    const bahanDigitalColumn = resolveColumnName(rows, ["bahan digital"]);
    const bahanBukanBukuColumn = resolveColumnName(rows, ["bahan bukan buku"]);
    const fiksyenColumn = resolveColumnName(rows, ["fiksyen"]);
    const bukanFiksyenColumn = resolveColumnName(rows, ["bukan fiksyen"]);
    // Prefer "Rekod" for AINS import source; keep "AINS" as backward-compatible fallback.
    const ainsColumn = resolveColumnName(rows, ["rekod", "ains"]);
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
      const availableHeaders = rows.length ? Object.keys(rows[0]).join(", ") : "(tiada header)";
      throw new Error(
        `Kolum data fail tidak lengkap. Perlu ada: NAMA MURID, No. Kad Pengenalan, Kelas, Bahan Digital, Bahan Bukan Buku, Fiksyen, Bukan Fiksyen, Bahasa Melayu, Bahasa Inggeris, Lain-lain Bahasa. Header dikesan: ${availableHeaders}`
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
      const ains = toInt999(ainsColumn ? row[ainsColumn] : 0);
      const bahasaMelayu = toInt999(row[bmColumn]);
      const bahasaInggeris = toInt999(row[biColumn]);
      const lainLainBahasa = toInt999(row[lainColumn]);
      const jumlahBacaan = bahanDigital + bahanBukanBuku + fiksyen + bukanFiksyen + ains;
      const tarikh = defaultTarikhForPeriod(year, month);

      records.push({
        tahun: year,
        bulan: month,
        tarikh,
        nama_pengisi: "ADMIN_IMPORT",
        guru: "Nilam",
        bil: 1,
        no_kad_pengenalan: noKad,
        nama,
        kelas,
        bahan_digital: bahanDigital,
        bahan_bukan_buku: bahanBukanBuku,
        fiksyen,
        bukan_fiksyen: bukanFiksyen,
        ains,
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
      records: assignBilByClass(records),
      students: dedupeStudentsByNoKad(students),
    };
  }

  function parseImportedAinsRows(rows, year, month) {
    const namaColumn = resolveColumnName(rows, ["nama murid", "nama"]);
    const emailColumn = resolveColumnName(rows, [
      "id delima",
      "id delima murid",
      "id delima (email google classroom)",
      "email id google classroom",
      "email google classroom",
      "email",
    ]);
    const ainsColumn = resolveColumnName(rows, ["rekod"]);

    if (!namaColumn || !emailColumn || !ainsColumn) {
      const availableHeaders = rows.length ? Object.keys(rows[0]).join(", ") : "(tiada header)";
      throw new Error(
        `Kolum fail AINS tidak lengkap. Perlu ada: NAMA MURID, ID DELIMa/Email Google Classroom, dan kolum Rekod. Header dikesan: ${availableHeaders}`
      );
    }

    const namelist = getCurrentNamelist().map(normalizeStudentRow);
    const studentByMatchKey = new Map();
    namelist.forEach((student) => {
      const matchKey = toStudentMatchKey(student.nama, student.email_google_classroom);
      if (!matchKey) {
        return;
      }
      if (!studentByMatchKey.has(matchKey)) {
        studentByMatchKey.set(matchKey, student);
      }
    });

    const nowIso = new Date().toISOString();
    const records = [];
    let unmatchedCount = 0;
    rows.forEach((row) => {
      const namaCsv = String(row[namaColumn] || "").trim();
      const emailCsv = String(row[emailColumn] || "").trim();
      const matchKey = toStudentMatchKey(namaCsv, emailCsv);
      if (!matchKey) {
        return;
      }

      const matchedStudent = studentByMatchKey.get(matchKey);
      if (!matchedStudent || !matchedStudent.kelas || !matchedStudent.nama) {
        unmatchedCount += 1;
        return;
      }

      const noKad = matchedStudent.no_kad_pengenalan || syntheticNoKad(matchedStudent);
      const ains = toInt999(row[ainsColumn]);
      const tarikh = defaultTarikhForPeriod(year, month);

      records.push({
        tahun: year,
        bulan: month,
        tarikh,
        nama_pengisi: "ADMIN_IMPORT",
        guru: "Nilam",
        bil: 1,
        no_kad_pengenalan: noKad,
        nama: matchedStudent.nama,
        kelas: matchedStudent.kelas,
        bahan_digital: 0,
        bahan_bukan_buku: 0,
        fiksyen: 0,
        bukan_fiksyen: 0,
        ains,
        bahasa_melayu: 0,
        bahasa_inggeris: 0,
        lain_lain_bahasa: 0,
        jumlah_aktiviti: ains,
        updated_at_client: nowIso,
      });
    });

    return {
      records: assignBilByClass(dedupeRecordsByNoKad(records)),
      unmatchedCount,
    };
  }

  function dedupeRecordsByNoKad(records) {
    const map = new Map();
    records.forEach((row) => {
      const key = recordSessionKey(row);
      if (!key) {
        return;
      }
      map.set(key, row);
    });
    return [...map.values()];
  }

  function toStudentMatchKey(nama, emailDelima) {
    const namaNorm = String(nama || "").trim().toLowerCase();
    const emailNorm = normalizeDelimaId(emailDelima);
    if (!namaNorm || !emailNorm) {
      return "";
    }
    return `${namaNorm}|${emailNorm}`;
  }

  function normalizeDelimaId(value) {
    return String(value || "").trim().toLowerCase();
  }

  function buildRecordWithAinsOverride(ainsImportRow, existingByKey) {
    const key = recordSessionKey(ainsImportRow);
    const existing = existingByKey.get(key) || {};
    const bahanDigital = toInt999(existing.bahan_digital);
    const bahanBukanBuku = toInt999(existing.bahan_bukan_buku);
    const fiksyen = toInt999(existing.fiksyen);
    const bukanFiksyen = toInt999(existing.bukan_fiksyen);
    const ains = toInt999(ainsImportRow.ains);
    const fallbackBil = toPositiveInt(ainsImportRow.bil, 1);
    const bahasaMelayu = toInt999(existing.bahasa_melayu);
    const bahasaInggeris = toInt999(existing.bahasa_inggeris);
    const lainLainBahasa = toInt999(existing.lain_lain_bahasa);

    return {
      tahun: ainsImportRow.tahun,
      bulan: ainsImportRow.bulan,
      tarikh: normalizeTarikh(ainsImportRow.tarikh, ainsImportRow.tahun, ainsImportRow.bulan),
      nama_pengisi: String(ainsImportRow.nama_pengisi || "ADMIN_IMPORT").trim(),
      guru: normalizeGuruType(ainsImportRow.guru),
      bil: toPositiveInt(existing.bil, fallbackBil),
      no_kad_pengenalan: normalizeNoKad(ainsImportRow.no_kad_pengenalan),
      nama: ainsImportRow.nama,
      kelas: ainsImportRow.kelas,
      bahan_digital: bahanDigital,
      bahan_bukan_buku: bahanBukanBuku,
      fiksyen,
      bukan_fiksyen: bukanFiksyen,
      ains,
      bahasa_melayu: bahasaMelayu,
      bahasa_inggeris: bahasaInggeris,
      lain_lain_bahasa: lainLainBahasa,
      jumlah_aktiviti: computeJumlahBacaan({
        bahan_digital: bahanDigital,
        bahan_bukan_buku: bahanBukanBuku,
        fiksyen,
        bukan_fiksyen: bukanFiksyen,
        ains,
      }),
      updated_at_client: new Date().toISOString(),
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

  function mergeRecordsByNoKad(year, month, importedRecords, baseRecords) {
    const existing = Array.isArray(baseRecords) ? baseRecords : readAllLocalRecords();
    const map = new Map();

    existing.forEach((row) => {
      const key = recordSessionKey(row);
      if (key) {
        map.set(key, row);
      }
    });
    importedRecords.forEach((row) => {
      const key = recordSessionKey(row);
      if (key) {
        map.set(key, row);
      }
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

  function overwriteAinsForPeriod(records, year, month, importedAinsRecords) {
    const targetYear = String(year || "").trim();
    const targetMonth = String(month || "").trim();
    const targetNoKad = new Set(
      (Array.isArray(importedAinsRecords) ? importedAinsRecords : [])
        .map((row) => normalizeNoKad(row?.no_kad_pengenalan))
        .filter(Boolean)
    );
    if (!targetNoKad.size) {
      return Array.isArray(records) ? [...records] : [];
    }

    const nowIso = new Date().toISOString();
    return (Array.isArray(records) ? records : []).map((row) => {
      const rowYear = String(row?.tahun || "").trim();
      const rowMonth = String(row?.bulan || "").trim();
      const rowNoKad = normalizeNoKad(row?.no_kad_pengenalan);
      if (rowYear !== targetYear || rowMonth !== targetMonth || !targetNoKad.has(rowNoKad)) {
        return row;
      }
      const next = {
        ...row,
        ains: 0,
        updated_at_client: nowIso,
      };
      next.jumlah_aktiviti = computeJumlahBacaan(next);
      return next;
    });
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
    const tarikh = normalizeTarikh(row.tarikh, tahun, bulan);
    const namaPengisi = String(row.nama_pengisi || "ADMIN_IMPORT").trim();
    const guru = normalizeGuruType(row.guru);
    const noKad = normalizeNoKad(row.no_kad_pengenalan);
    const nama = String(row.nama || "").trim();
    const kelas = String(row.kelas || "").trim();
    if (!tahun || !bulan || !tarikh || !namaPengisi || !guru || !noKad || !nama || !kelas) {
      return null;
    }

    const bahanDigital = toInt999(row.bahan_digital);
    const bahanBukanBuku = toInt999(row.bahan_bukan_buku);
    const fiksyen = toInt999(row.fiksyen);
    const bukanFiksyen = toInt999(row.bukan_fiksyen);
    const ains = toInt999(row.ains);
    const bm = toInt999(row.bahasa_melayu);
    const bi = toInt999(row.bahasa_inggeris);
    const lain = toInt999(row.lain_lain_bahasa);
    return {
      id: Number.isFinite(Number(row.id)) && Number(row.id) > 0 ? Number(row.id) : undefined,
      tahun,
      bulan,
      tarikh,
      nama_pengisi: namaPengisi,
      guru,
      bil: toPositiveInt(row.bil, 1),
      no_kad_pengenalan: noKad,
      nama,
      kelas,
      bahan_digital: bahanDigital,
      bahan_bukan_buku: bahanBukanBuku,
      fiksyen,
      bukan_fiksyen: bukanFiksyen,
      ains,
      bahasa_melayu: bm,
      bahasa_inggeris: bi,
      lain_lain_bahasa: lain,
      jumlah_aktiviti:
        Number.isFinite(Number(row.jumlah_aktiviti)) && Number(row.jumlah_aktiviti) >= 0
          ? Math.trunc(Number(row.jumlah_aktiviti))
          : computeJumlahBacaan({
              bahan_digital: bahanDigital,
              bahan_bukan_buku: bahanBukanBuku,
              fiksyen,
              bukan_fiksyen: bukanFiksyen,
              ains,
            }),
      updated_at_client: row.updated_at_client || new Date().toISOString(),
    };
  }

  function computeJumlahBacaan(row) {
    return (
      toInt999(row?.bahan_digital) +
      toInt999(row?.bahan_bukan_buku) +
      toInt999(row?.fiksyen) +
      toInt999(row?.bukan_fiksyen) +
      toInt999(row?.ains)
    );
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
    const endpoint =
      `${supabaseUrl}/rest/v1/nilam_records?on_conflict=tahun,bulan,tarikh,no_kad_pengenalan,nama_pengisi,guru`;
    try {
      const payload = (Array.isArray(records) ? records : []).map((row) => ({
        ...row,
        bil: toPositiveInt(row?.bil, 1),
      }));
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

      if (response.ok) {
        return;
      }

      const detail = await response.text();
      const isNoConflictConstraint = response.status === 400 && String(detail || "").includes("42P10");
      if (!isNoConflictConstraint) {
        throw new Error(`Sync import data ke Supabase gagal (${response.status}): ${detail}`);
      }
    } catch (error) {
      const message = String(error && error.message ? error.message : error || "");
      if (!message.includes("42P10")) {
        throw error;
      }
    }

    await replaceMonthlyAdminImportRecordsInSupabase(records, config);
  }

  async function replaceMonthlyAdminImportRecordsInSupabase(records, config) {
    const safeRecords = (Array.isArray(records) ? records.filter(Boolean) : []).map((row) => ({
      ...row,
      bil: toPositiveInt(row?.bil, 1),
    }));
    if (!safeRecords.length) {
      return;
    }

    const year = String(safeRecords[0].tahun || "").trim();
    const month = String(safeRecords[0].bulan || "").trim();
    if (!year || !month) {
      return;
    }

    const adminNames = [...new Set(
      safeRecords
        .map((row) => String(row.nama_pengisi || "").trim())
        .filter(Boolean)
    )];
    const guruTypes = [...new Set(
      safeRecords
        .map((row) => normalizeGuruType(row.guru))
        .filter(Boolean)
    )];

    const supabaseUrl = config.supabaseUrl.replace(/\/$/, "");
    for (const namaPengisi of adminNames) {
      for (const guruType of guruTypes) {
        const params = new URLSearchParams({
          tahun: `eq.${year}`,
          bulan: `eq.${month}`,
          nama_pengisi: `eq.${namaPengisi}`,
          guru: `eq.${guruType}`,
        });
        const deleteEndpoint = `${supabaseUrl}/rest/v1/nilam_records?${params.toString()}`;
        const deleteRes = await fetch(deleteEndpoint, {
          method: "DELETE",
          headers: {
            apikey: config.supabaseAnonKey,
            Authorization: `Bearer ${config.supabaseAnonKey}`,
          },
        });
        if (!deleteRes.ok) {
          const detail = await deleteRes.text();
          throw new Error(`Padam sesi ADMIN_IMPORT gagal (${deleteRes.status}): ${detail}`);
        }
      }
    }

    const insertEndpoint = `${supabaseUrl}/rest/v1/nilam_records`;
    const insertRes = await fetch(insertEndpoint, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        apikey: config.supabaseAnonKey,
        Authorization: `Bearer ${config.supabaseAnonKey}`,
        Prefer: "return=minimal",
      },
      body: JSON.stringify(safeRecords),
    });

    if (!insertRes.ok) {
      const detail = await insertRes.text();
      throw new Error(`Simpan semula sesi ADMIN_IMPORT gagal (${insertRes.status}): ${detail}`);
    }
  }

  async function upsertImportedRecordsToSupabaseById(records) {
    const config = window.NILAM_CONFIG || {};
    if (!config.supabaseUrl || !config.supabaseAnonKey || !records.length) {
      return;
    }

    const supabaseUrl = config.supabaseUrl.replace(/\/$/, "");
    const endpoint = `${supabaseUrl}/rest/v1/nilam_records?on_conflict=id`;
    const payload = records.map((row) => {
      const out = { ...row };
      out.bil = toPositiveInt(out?.bil, 1);
      if (!Number.isFinite(Number(out.id)) || Number(out.id) <= 0) {
        delete out.id;
      } else {
        out.id = Number(out.id);
      }
      return out;
    });

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
      throw new Error(`Sync overwrite AINS ke Supabase gagal (${response.status}): ${detail}`);
    }
  }

  async function overwriteAinsInSupabase(year, month, importedAinsRows, ainsOverrideRecords) {
    const config = window.NILAM_CONFIG || {};
    if (!config.supabaseUrl || !config.supabaseAnonKey) {
      return;
    }

    const targetNoKad = [...new Set(
      (Array.isArray(importedAinsRows) ? importedAinsRows : [])
        .map((row) => normalizeNoKad(row?.no_kad_pengenalan))
        .filter(Boolean)
    )];
    if (!targetNoKad.length) {
      return;
    }

    const existing = await fetchSupabaseRecordsByNoKadForPeriod(year, month, targetNoKad, config);
    const cleared = overwriteAinsForPeriod(existing, year, month, importedAinsRows);
    const byKey = new Map();
    cleared.forEach((row) => {
      const key = recordSessionKey(row);
      if (key) {
        byKey.set(key, row);
      }
    });
    (Array.isArray(ainsOverrideRecords) ? ainsOverrideRecords : []).forEach((row) => {
      const key = recordSessionKey(row);
      if (key) {
        const existingRow = byKey.get(key);
        if (existingRow && Number.isFinite(Number(existingRow.id)) && !Number.isFinite(Number(row.id))) {
          byKey.set(key, { ...row, id: Number(existingRow.id) });
        } else {
          byKey.set(key, row);
        }
      }
    });

    await upsertImportedRecordsToSupabaseById([...byKey.values()]);
  }

  async function fetchSupabaseRecordsByNoKadForPeriod(year, month, noKadList, config) {
    const supabaseUrl = config.supabaseUrl.replace(/\/$/, "");
    const chunkSize = 80;
    const all = [];
    for (let i = 0; i < noKadList.length; i += chunkSize) {
      const chunk = noKadList.slice(i, i + chunkSize);
      const params = new URLSearchParams({
        select:
          "id,tahun,bulan,tarikh,nama_pengisi,guru,bil,no_kad_pengenalan,nama,kelas,bahan_digital,bahan_bukan_buku,fiksyen,bukan_fiksyen,ains,bahasa_melayu,bahasa_inggeris,lain_lain_bahasa,jumlah_aktiviti,updated_at_client",
        tahun: `eq.${year}`,
        bulan: `eq.${month}`,
        no_kad_pengenalan: `in.(${chunk.join(",")})`,
        limit: "50000",
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
        throw new Error(`Semakan rekod Supabase untuk overwrite AINS gagal (${response.status}): ${detail}`);
      }
      const rows = await response.json();
      if (Array.isArray(rows) && rows.length) {
        rows.forEach((row) => {
          const normalized = normalizeImportedRecord(row, "");
          if (normalized) {
            all.push(normalized);
          }
        });
      }
    }
    return all;
  }

  function toInt999(value) {
    const number = Number(value);
    if (!Number.isFinite(number)) {
      return 0;
    }
    return Math.max(0, Math.min(999, Math.trunc(number)));
  }

  function toPositiveInt(value, fallback = 1) {
    const number = Number(value);
    if (Number.isFinite(number) && number > 0) {
      return Math.trunc(number);
    }
    return Math.max(1, Math.trunc(Number(fallback) || 1));
  }

  function assignBilByClass(records) {
    const list = (Array.isArray(records) ? records : []).map((row) => ({
      ...row,
      bil: toPositiveInt(row?.bil, 1),
    }));
    const grouped = new Map();
    list.forEach((row) => {
      const kelas = String(row?.kelas || "").trim();
      if (!kelas) {
        return;
      }
      if (!grouped.has(kelas)) {
        grouped.set(kelas, []);
      }
      grouped.get(kelas).push(row);
    });
    for (const rows of grouped.values()) {
      rows.sort((a, b) => String(a?.nama || "").localeCompare(String(b?.nama || ""), "ms"));
      rows.forEach((row, idx) => {
        row.bil = idx + 1;
      });
    }
    return list;
  }

  // Generates a deterministic synthetic IC for students without one.
  // Prefixed with NOIC_ so it's distinguishable from real IC numbers.
  function syntheticNoKad(row) {
    const kelas = String(row.kelas || "").trim().toUpperCase().replace(/\s+/g, "");
    const nama = String(row.nama || "").trim().toUpperCase().replace(/[^A-Z0-9]/g, "_").slice(0, 30);
    return `NOIC_${kelas}_${nama}`;
  }

  async function upsertStudentsToSupabase(students, options = {}) {
    const deactivateMissing = options.deactivateMissing !== false;
    const config = window.NILAM_CONFIG || {};
    if (!config.supabaseUrl || !config.supabaseAnonKey) {
      return;
    }

    const year = String(new Date().getFullYear());
    const kelasField = `kelas_${year}`;
    const payload = students
      .filter((row) => row.nama && row.kelas)
      .map((row) => ({
        no_kad_pengenalan: row.no_kad_pengenalan || syntheticNoKad(row),
        nama_murid: row.nama,
        jantina: row.jantina,
        email_google_classroom: row.email_google_classroom,
        [kelasField]: row.kelas,
        active: true,
      }));

    const supabaseUrl = config.supabaseUrl.replace(/\/$/, "");
    const headers = {
      "Content-Type": "application/json",
      apikey: config.supabaseAnonKey,
      Authorization: `Bearer ${config.supabaseAnonKey}`,
    };

    // Upsert all active students.
    if (payload.length) {
      const upsertRes = await fetch(`${supabaseUrl}/rest/v1/nilam_students?on_conflict=no_kad_pengenalan`, {
        method: "POST",
        headers: { ...headers, Prefer: "resolution=merge-duplicates,return=minimal" },
        body: JSON.stringify(payload),
      });
      if (!upsertRes.ok) {
        const detail = await upsertRes.text();
        throw new Error(`Sync senarai murid ke Supabase gagal (${upsertRes.status}): ${detail}`);
      }
    }

    // Soft-delete removed students (active=false) instead of deleting,
    // to avoid FK constraint violations from nilam_records references.
    const keepIcs = payload.map((r) => r.no_kad_pengenalan).join(",");
    if (deactivateMissing && keepIcs) {
      const deactivateRes = await fetch(
        `${supabaseUrl}/rest/v1/nilam_students?no_kad_pengenalan=not.in.(${keepIcs})`,
        {
          method: "PATCH",
          headers: { ...headers, Prefer: "return=minimal" },
          body: JSON.stringify({ active: false }),
        }
      );
      if (!deactivateRes.ok) {
        const detail = await deactivateRes.text();
        throw new Error(`Kemaskini status murid lama gagal (${deactivateRes.status}): ${detail}`);
      }
    }
  }

  function buildAllowedKeysByClass(students) {
    const map = new Map();
    (Array.isArray(students) ? students : [])
      .map(normalizeStudentRow)
      .forEach((row) => {
        if (!row.kelas || !row.nama) {
          return;
        }
        if (!map.has(row.kelas)) {
          map.set(row.kelas, new Set());
        }
        const keys = map.get(row.kelas);
        keys.add(`nm:${row.nama.toLowerCase()}`);
        if (row.no_kad_pengenalan) {
          keys.add(`ic:${normalizeNoKad(row.no_kad_pengenalan)}`);
        }
      });
    return map;
  }

  function isRecordAllowedByMaster(row, allowedByClass) {
    const kelas = String(row?.kelas || "").trim();
    if (!kelas) {
      return false;
    }
    const keys = allowedByClass.get(kelas);
    if (!keys || !keys.size) {
      return false;
    }
    const noKad = normalizeNoKad(row?.no_kad_pengenalan);
    if (noKad && keys.has(`ic:${noKad}`)) {
      return true;
    }
    const nama = String(row?.nama || "").trim().toLowerCase();
    return nama ? keys.has(`nm:${nama}`) : false;
  }

  function purgeLocalRecordsByMaster(students) {
    const allowedByClass = buildAllowedKeysByClass(students);
    let deletedCount = 0;

    for (let i = 0; i < localStorage.length; i += 1) {
      const key = localStorage.key(i);
      if (!key || !key.startsWith("nilam_records_")) {
        continue;
      }
      const parsed = parseStoredRecords(localStorage.getItem(key));
      if (!parsed.length) {
        continue;
      }
      const kept = parsed.filter((row) => isRecordAllowedByMaster(row, allowedByClass));
      if (kept.length === parsed.length) {
        continue;
      }
      deletedCount += parsed.length - kept.length;
      if (kept.length) {
        localStorage.setItem(key, JSON.stringify(kept));
      } else {
        localStorage.removeItem(key);
      }
    }

    return deletedCount;
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

  async function purgeSupabaseRecordsByMaster(students, config) {
    if (!config.supabaseUrl || !config.supabaseAnonKey) {
      return 0;
    }

    const allowedByClass = buildAllowedKeysByClass(students);
    const supabaseUrl = config.supabaseUrl.replace(/\/$/, "");
    const headers = {
      apikey: config.supabaseAnonKey,
      Authorization: `Bearer ${config.supabaseAnonKey}`,
    };

    const params = new URLSearchParams({
      select: "id,kelas,nama,no_kad_pengenalan",
      limit: "50000",
    });
    const listRes = await fetch(`${supabaseUrl}/rest/v1/nilam_records?${params.toString()}`, {
      method: "GET",
      headers,
    });
    if (!listRes.ok) {
      const detail = await listRes.text();
      throw new Error(`Gagal semak rekod Supabase (${listRes.status}): ${detail}`);
    }

    const rows = await listRes.json();
    const removableIds = (Array.isArray(rows) ? rows : [])
      .filter((row) => !isRecordAllowedByMaster(row, allowedByClass))
      .map((row) => Number(row.id))
      .filter((id) => Number.isFinite(id) && id > 0);

    if (!removableIds.length) {
      return 0;
    }

    const batchSize = 300;
    for (let i = 0; i < removableIds.length; i += batchSize) {
      const batch = removableIds.slice(i, i + batchSize);
      const deleteRes = await fetch(`${supabaseUrl}/rest/v1/nilam_records?id=in.(${batch.join(",")})`, {
        method: "DELETE",
        headers: {
          ...headers,
          Prefer: "return=minimal",
        },
      });
      if (!deleteRes.ok) {
        const detail = await deleteRes.text();
        throw new Error(`Gagal padam rekod Supabase (${deleteRes.status}): ${detail}`);
      }
    }

    return removableIds.length;
  }

  async function syncRecordsToMasterList(students) {
    const localDeleted = purgeLocalRecordsByMaster(students);
    let supabaseDeleted = 0;
    const config = window.NILAM_CONFIG || {};
    try {
      supabaseDeleted = await purgeSupabaseRecordsByMaster(students, config);
    } catch (error) {
      console.error(error);
    }
    return { localDeleted, supabaseDeleted };
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

  function loadStoredTeacherNames() {
    const raw = localStorage.getItem(TEACHER_NAMES_KEY);
    if (!raw) {
      return [];
    }
    try {
      const parsed = JSON.parse(raw);
      if (!Array.isArray(parsed)) {
        return [];
      }
      return parsed
        .map((value) => String(value || "").trim())
        .filter(Boolean);
    } catch (error) {
      console.error("Gagal parse senarai nama guru", error);
      return [];
    }
  }

  function normalizeNoKad(value) {
    return String(value || "")
      .replace(/\s+/g, "")
      .toUpperCase()
      .trim();
  }

  function isExcelFile(file) {
    const name = String(file?.name || "").toLowerCase();
    return name.endsWith(".xlsx") || name.endsWith(".xls");
  }

  async function readRowsFromFile(file) {
    if (isExcelFile(file)) {
      return readRowsFromExcel(file);
    }
    const text = await file.text();
    return parseCsv(text);
  }

  async function readRowsFromExcel(file) {
    if (!window.XLSX) {
      throw new Error("Parser Excel tidak tersedia. Sila refresh halaman dan cuba semula.");
    }
    const arrayBuffer = await file.arrayBuffer();
    const workbook = window.XLSX.read(arrayBuffer, { type: "array" });
    const firstSheetName = workbook.SheetNames[0];
    if (!firstSheetName) {
      return [];
    }
    const sheet = workbook.Sheets[firstSheetName];
    const rows = window.XLSX.utils.sheet_to_json(sheet, {
      defval: "",
      raw: false,
      blankrows: false,
    });
    return Array.isArray(rows) ? rows : [];
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

  function resolveDynamicKelasColumn(rows) {
    if (!rows.length) {
      return "";
    }
    const firstRowKeys = Object.keys(rows[0]);
    const normalized = firstRowKeys.map((key) => ({
      original: key,
      normalized: normalizeHeader(key),
    }));

    const exact = normalized.find((item) => item.normalized === "kelas");
    if (exact) {
      return exact.original;
    }

    const kelasByYear = normalized.find((item) => /^kelas \d{4}$/.test(item.normalized));
    if (kelasByYear) {
      return kelasByYear.original;
    }

    const startsWithKelas = normalized.find((item) => item.normalized.startsWith("kelas "));
    return startsWithKelas ? startsWithKelas.original : "";
  }

  function normalizeHeader(value) {
    return String(value || "")
      .replace(/^\uFEFF/, "")
      .toLowerCase()
      .replace(/['’`"]/g, "")
      .replace(/[^\p{L}\p{N}]+/gu, " ")
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

  function showPopupStatus(message, isError) {
    window.alert(isError ? String(message || "Operasi gagal.") : "Berjaya disimpan");
  }

  // ── Undo ──────────────────────────────────────────────────────────────────

  function undoRemove() {
    const last = state.undoStack.pop();
    if (!last) {
      return;
    }
    // If only a single empty placeholder exists, remove it first
    if (
      state.students.length === 1 &&
      !state.students[0].nama &&
      !state.students[0].kelas &&
      !state.students[0].no_kad_pengenalan
    ) {
      state.students = [];
    }
    const insertAt = Math.min(last.index, state.students.length);
    state.students.splice(insertAt, 0, last.row);
    state.currentPage = Math.max(1, Math.ceil((insertAt + 1) / state.rowsPerPage));
    renderManageTable();
    updateUndoBtn();
    queueAutoSync({ deactivateMissing: true });
  }

  function updateUndoBtn() {
    if (el.undoRemoveBtn) {
      el.undoRemoveBtn.disabled = state.undoStack.length === 0;
    }
  }

  // ── Search ────────────────────────────────────────────────────────────────

  function handleSearchInput() {
    state.searchQuery = el.searchInput ? el.searchInput.value : "";
    state.currentPage = 1;
    renderManageTable();
  }

  // ── Compare Namelist ──────────────────────────────────────────────────────

  function startCompareNamelist() {
    if (el.compareFileInput) {
      el.compareFileInput.value = "";
      el.compareFileInput.click();
    }
  }

  async function handleCompareFile(event) {
    const [file] = event.target.files || [];
    if (!file) {
      return;
    }

    try {
      const rows = await readRowsFromFile(file);

      const namaColumn = resolveColumnName(rows, ["nama murid", "nama"]);
      const jantinaColumn = resolveColumnName(rows, ["jantina"]);
      const kelasColumn = resolveColumnName(rows, ["kelas 2026", "kelas"]);
      const noKadColumn = resolveColumnName(rows, ["no. kad pengenalan", "no kad pengenalan"]);
      const emailColumn = resolveColumnName(rows, ["email id google classroom", "email google classroom", "email"]);

      if (!namaColumn || !kelasColumn) {
        throw new Error("Kolum fail tidak sah. Wajib ada NAMA MURID dan Kelas.");
      }

      const newStudents = rows
        .map((row) =>
          normalizeStudentRow({
            nama: row[namaColumn],
            jantina: jantinaColumn ? row[jantinaColumn] : "",
            kelas: row[kelasColumn],
            no_kad_pengenalan: noKadColumn ? row[noKadColumn] : "",
            email_google_classroom: emailColumn ? row[emailColumn] : "",
          })
        )
        .filter((row) => row.nama && row.kelas);

      if (!newStudents.length) {
        throw new Error("Tiada murid sah ditemui dalam fail.");
      }

      // Match by IC when present (and not synthetic), else fall back to normalised name.
      function studentKey(r) {
        const ic = String(r.no_kad_pengenalan || "").trim();
        if (ic && !ic.startsWith("NOIC_")) {
          return ic;
        }
        return `__name__${String(r.nama || "").trim().toLowerCase()}`;
      }

      const currentNamelist = getCurrentNamelist().map(normalizeStudentRow);
      const currentByKey = new Map(currentNamelist.map((r) => [studentKey(r), r]));
      const newByKey = new Map();
      newStudents.forEach((r) => newByKey.set(studentKey(r), r));

      const toAdd = newStudents.filter((r) => !currentByKey.has(studentKey(r)));
      const toRemove = currentNamelist.filter((r) => !newByKey.has(studentKey(r)));
      const toUpdateKelas = newStudents
        .filter((r) => {
          const key = studentKey(r);
          return currentByKey.has(key) && currentByKey.get(key).kelas !== r.kelas;
        })
        .map((r) => ({ ...r, kelasLama: currentByKey.get(studentKey(r)).kelas }));
      const toUpdateEmail = newStudents
        .filter((r) => {
          const key = studentKey(r);
          return currentByKey.has(key) &&
            currentByKey.get(key).email_google_classroom !== r.email_google_classroom;
        })
        .map((r) => ({ ...r, emailLama: currentByKey.get(studentKey(r)).email_google_classroom }));

      const finalList = [...newByKey.values()].sort((a, b) => {
        const byKelas = a.kelas.localeCompare(b.kelas, "ms");
        return byKelas !== 0 ? byKelas : a.nama.localeCompare(b.nama, "ms");
      });

      state.pendingCompareResult = { finalList };
      showCompareModal(toAdd, toRemove, toUpdateKelas, toUpdateEmail, newStudents.length);
    } catch (error) {
      console.error(error);
      setStatus(error.message || "Gagal memuatkan fail perbandingan.", true);
    } finally {
      if (event.target) {
        event.target.value = "";
      }
    }
  }

  function showCompareModal(toAdd, toRemove, toUpdateKelas, toUpdateEmail, totalNew) {
    if (!el.compareModal || !el.compareContent) {
      return;
    }

    let html = `<p style="margin:0 0 0.5rem;font-size:0.85rem;color:var(--muted)">Jumlah murid dalam fail baharu: <strong>${totalNew}</strong></p>`;

    if (toAdd.length) {
      html += `<p class="compare-section-title compare-add">Baharu — ${toAdd.length} murid akan ditambah:</p>`;
      html += `<ul class="compare-list">${toAdd.map((r) => `<li>${escapeAttr(r.nama)} (${escapeAttr(r.kelas)})</li>`).join("")}</ul>`;
    } else {
      html += `<p class="compare-section-title">Tiada murid baharu ditemui.</p>`;
    }

    if (toRemove.length) {
      html += `<p class="compare-section-title compare-remove">Dibuang — ${toRemove.length} murid akan dibuang:</p>`;
      html += `<ul class="compare-list">${toRemove.map((r) => `<li>${escapeAttr(r.nama)} (${escapeAttr(r.kelas)})</li>`).join("")}</ul>`;
    } else {
      html += `<p class="compare-section-title">Tiada murid dibuang.</p>`;
    }

    if (toUpdateKelas.length) {
      html += `<p class="compare-section-title compare-update">Kelas Berubah — ${toUpdateKelas.length} murid:</p>`;
      html += `<ul class="compare-list">${toUpdateKelas.map((r) =>
        `<li>${escapeAttr(r.nama)} — ${escapeAttr(r.kelasLama)} → ${escapeAttr(r.kelas)}</li>`
      ).join("")}</ul>`;
    }

    if (toUpdateEmail.length) {
      html += `<p class="compare-section-title compare-update">Email Berubah — ${toUpdateEmail.length} murid:</p>`;
      html += `<ul class="compare-list">${toUpdateEmail.map((r) =>
        `<li>${escapeAttr(r.nama)} — ${escapeAttr(r.emailLama || "(tiada)")} → ${escapeAttr(r.email_google_classroom || "(tiada)")}</li>`
      ).join("")}</ul>`;
    }

    el.compareContent.innerHTML = html;
    el.compareModal.hidden = false;
  }

  function closeCompareModal() {
    if (el.compareModal) {
      el.compareModal.hidden = true;
    }
    state.pendingCompareResult = null;
  }

  async function confirmCompare() {
    if (!state.pendingCompareResult) {
      return;
    }
    const { finalList } = state.pendingCompareResult;
    closeCompareModal();

    try {
      localStorage.setItem(NAMELIST_OVERRIDE_KEY, JSON.stringify(finalList));
      state.students = finalList;
      state.undoStack = [];
      updateUndoBtn();
      state.currentPage = 1;
      if (state.isManageOpen) {
        renderManageTable();
      }
      await upsertStudentsToSupabase(finalList);
      const purgeResult = await syncRecordsToMasterList(finalList);
      setStatus(
        `Senarai nama dikemas kini: ${finalList.length} murid. Data murid yang dibuang telah dipadam (Local: ${purgeResult.localDeleted}, Supabase: ${purgeResult.supabaseDeleted}).`
      );
      showPopupStatus("Berjaya disimpan", false);
    } catch (error) {
      console.error(error);
      const message = error.message || "Gagal kemaskini senarai nama.";
      setStatus(message, true);
      showPopupStatus(message, true);
    }
  }
})();
