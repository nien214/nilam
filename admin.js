(function () {
  "use strict";

  const AUTH_USER = "admin";
  const AUTH_PASS = "nilam_admin";
  const AUTH_SESSION_KEY = "nilam_admin_auth_v1";
  const NAMELIST_OVERRIDE_KEY = "nilam_students_override_v1";
  const TEACHER_NAMES_KEY = "nilam_teacher_names_v1";


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
    isNilamUpdateOpen: false,
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
    nilamUpdateRows: [],
    nilamUpdateLoadedYear: "",
    nilamUpdateLoadedMonth: "",
    nilamUpdateLoadedClass: "",
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
    updateNilamBtn: document.getElementById("updateNilamBtn"),
    manageSection: document.getElementById("studentManageSection"),
    nilamUpdateSection: document.getElementById("nilamUpdateSection"),
    toggleAdminPanelBtn: document.getElementById("toggleAdminPanelBtn"),
    toggleAdminPanelBtnUpdate: document.getElementById("toggleAdminPanelBtnUpdate"),
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
    nilamUpdateYearInput: document.getElementById("updateNilamYearInput"),
    nilamUpdateMonthSelect: document.getElementById("updateNilamMonthSelect"),
    nilamUpdateClassSelect: document.getElementById("updateNilamClassSelect"),
    nilamUpdateSaveBtn: document.getElementById("saveNilamUpdateBtn"),
    nilamUpdateTbody: document.getElementById("nilamUpdateTbody"),
    status: document.getElementById("adminStatus"),
  };

  init();

  function init() {
    initImportMonthOptions();
    initNilamUpdateMonthOptions();
    initNilamUpdateFilters();
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
    if (el.updateNilamBtn) {
      el.updateNilamBtn.addEventListener("click", toggleNilamUpdate);
    }
    if (el.toggleAdminPanelBtn) {
      el.toggleAdminPanelBtn.addEventListener("click", toggleAdminPanelVisibility);
    }
    if (el.toggleAdminPanelBtnUpdate) {
      el.toggleAdminPanelBtnUpdate.addEventListener("click", toggleAdminPanelVisibility);
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
    if (el.nilamUpdateYearInput) {
      el.nilamUpdateYearInput.addEventListener("change", handleNilamUpdateFilterChange);
    }
    if (el.nilamUpdateMonthSelect) {
      el.nilamUpdateMonthSelect.addEventListener("change", handleNilamUpdateFilterChange);
    }
    if (el.nilamUpdateClassSelect) {
      el.nilamUpdateClassSelect.addEventListener("change", handleNilamUpdateFilterChange);
    }
    if (el.nilamUpdateSaveBtn) {
      el.nilamUpdateSaveBtn.addEventListener("click", saveNilamUpdateData);
    }
    if (el.nilamUpdateTbody) {
      el.nilamUpdateTbody.addEventListener("input", handleNilamUpdateTableInput);
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
    if (el.nilamUpdateSection) {
      el.nilamUpdateSection.hidden = true;
    }
    state.isPanelHiddenInManage = false;
    updateAdminPanelToggleLabel();
    state.isManageOpen = false;
    state.isNilamUpdateOpen = false;
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
      state.isNilamUpdateOpen = false;
      if (el.nilamUpdateSection) {
        el.nilamUpdateSection.hidden = true;
      }
      setAdminPanelHidden(true);
      loadStudentsForManage();
      setStatus("Pengurusan senarai murid dibuka.");
      return;
    }
    setAdminPanelHidden(false);
  }

  function toggleNilamUpdate() {
    state.isNilamUpdateOpen = !state.isNilamUpdateOpen;
    if (el.nilamUpdateSection) {
      el.nilamUpdateSection.hidden = !state.isNilamUpdateOpen;
    }
    if (state.isNilamUpdateOpen) {
      state.isManageOpen = false;
      if (el.manageSection) {
        el.manageSection.hidden = true;
      }
      setAdminPanelHidden(true);
      initNilamUpdateFilters();
      handleNilamUpdateFilterChange();
      setStatus("Kemas Kini Data Nilam dibuka.");
      return;
    }
    setAdminPanelHidden(false);
  }

  function initNilamUpdateMonthOptions() {
    if (!el.nilamUpdateMonthSelect) {
      return;
    }
    el.nilamUpdateMonthSelect.innerHTML = "";
    MONTHS.forEach((month) => {
      const option = document.createElement("option");
      option.value = month;
      option.textContent = month;
      el.nilamUpdateMonthSelect.appendChild(option);
    });
  }

  function initNilamUpdateFilters() {
    const now = new Date();
    if (el.nilamUpdateYearInput) {
      const year = String(el.nilamUpdateYearInput.value || "").trim();
      if (!/^\d{4}$/.test(year)) {
        el.nilamUpdateYearInput.value = String(now.getFullYear());
      }
    }
    if (el.nilamUpdateMonthSelect && !normalizeMonth(el.nilamUpdateMonthSelect.value)) {
      el.nilamUpdateMonthSelect.value = MONTHS[now.getMonth()];
    }
    populateNilamUpdateClassOptions();
  }

  function populateNilamUpdateClassOptions() {
    if (!el.nilamUpdateClassSelect) {
      return;
    }
    const previous = String(el.nilamUpdateClassSelect.value || "").trim();
    const classSet = new Set();
    getCurrentNamelist()
      .map(normalizeStudentRow)
      .forEach((row) => {
        const kelas = String(row.kelas || "").trim();
        if (kelas) {
          classSet.add(kelas);
        }
      });
    readAllLocalRecords().forEach((row) => {
      const kelas = String(row?.kelas || "").trim();
      if (kelas) {
        classSet.add(kelas);
      }
    });
    const classes = [...classSet].sort((a, b) => a.localeCompare(b, "ms"));
    el.nilamUpdateClassSelect.innerHTML = '<option value="">Pilih kelas</option>';
    classes.forEach((kelas) => {
      const option = document.createElement("option");
      option.value = kelas;
      option.textContent = kelas;
      el.nilamUpdateClassSelect.appendChild(option);
    });
    if (previous && classSet.has(previous)) {
      el.nilamUpdateClassSelect.value = previous;
    }
  }

  function renderNilamUpdateEmpty(message) {
    if (!el.nilamUpdateTbody) {
      return;
    }
    el.nilamUpdateTbody.innerHTML = `
      <tr>
        <td colspan="14" class="empty">${escapeAttr(message || "Tiada data.")}</td>
      </tr>
    `;
  }

  async function loadStudentsForManage() {
    const config = window.NILAM_CONFIG || {};
    let raw = [];
    let cloudError = "";

    if (config.supabaseUrl && config.supabaseAnonKey) {
      try {
        raw = await fetchStudentsFromSupabase(config, String(new Date().getFullYear()));
      } catch (error) {
        console.error(error);
        cloudError = String(error?.message || error || "").slice(0, 220);
      }
    } else if (CLOUD_ONLY_MODE) {
      cloudError = "Cloud-only mode memerlukan konfigurasi Supabase dalam config.js.";
    }

    if (!raw.length && !CLOUD_ONLY_MODE) {
      raw = getCurrentNamelist();
    }

    state.students = (Array.isArray(raw) ? raw : []).map(normalizeStudentRow).filter((row) => row.nama || row.kelas);
    sortStudentsByClassThenName();
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
    populateNilamUpdateClassOptions();

    if (cloudError && CLOUD_ONLY_MODE) {
      setStatus(`Gagal memuatkan senarai murid dari cloud: ${cloudError}`, true);
    }
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
    const nextLabel = state.isPanelHiddenInManage
      ? "Tunjuk Panel Admin"
      : "Sembunyi Panel Admin";
    if (el.toggleAdminPanelBtn) {
      el.toggleAdminPanelBtn.textContent = nextLabel;
    }
    if (el.toggleAdminPanelBtnUpdate) {
      el.toggleAdminPanelBtnUpdate.textContent = nextLabel;
    }
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

      localSetItem(NAMELIST_OVERRIDE_KEY, JSON.stringify(toSave));
      state.students = toSave;
      state.undoStack = [];
      updateUndoBtn();
      renderManageTable();
      await upsertStudentsToSupabase(toSave);
      const purgeResult = await syncRecordsToMasterList(toSave);
      setStatus(
        `Senarai murid berjaya disimpan: ${toSave.length} murid. Data murid yang dibuang telah dipadam di cloud: ${purgeResult.supabaseDeleted}.`
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

      localSetItem(NAMELIST_OVERRIDE_KEY, JSON.stringify(toSync));
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

      localSetItem(NAMELIST_OVERRIDE_KEY, JSON.stringify(namelist));
      state.students = namelist;
      if (state.isManageOpen) {
        renderManageTable();
      }
      await upsertStudentsToSupabase(namelist);
      const purgeResult = await syncRecordsToMasterList(namelist);
      setStatus(
        `Namelist berjaya diimport: ${namelist.length} murid. Data murid yang dibuang telah dipadam di cloud: ${purgeResult.supabaseDeleted}.`
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

      const existingLocal = loadStoredTeacherNames();
      let existingCloud = [];
      try {
        existingCloud = await fetchTeacherNamesFromSupabase(window.NILAM_CONFIG || {});
      } catch (error) {
        console.error("Gagal muat senarai nama guru dari Supabase", error);
      }
      const existing = [...new Set([...existingLocal, ...existingCloud])];
      const merged = [...new Set([...existing, ...importedNames])].sort((a, b) =>
        a.localeCompare(b, "ms")
      );
      localSetItem(TEACHER_NAMES_KEY, JSON.stringify(merged));

      let syncNote = "";
      try {
        await upsertTeacherNamesToSupabase(merged);
        syncNote = " Senarai nama guru juga disimpan ke Supabase.";
      } catch (syncError) {
        if (CLOUD_ONLY_MODE) {
          throw syncError;
        }
        console.error(syncError);
        const detail = String(syncError?.message || syncError || "").slice(0, 220);
        syncNote = ` Simpanan Supabase gagal (${detail}).`;
      }

      setStatus(
        `Import Nama Guru berjaya: ${importedNames.length} nama diproses, jumlah senarai ${merged.length}.${syncNote}`
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
        if (CLOUD_ONLY_MODE) {
          throw syncError;
        }
        console.error(syncError);
        const detail = String(syncError?.message || syncError || "").slice(0, 220);
        syncNote = ` Simpanan Supabase gagal (${detail}).`;
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

      if (!getCurrentNamelist().length) {
        const config = window.NILAM_CONFIG || {};
        if (!config.supabaseUrl || !config.supabaseAnonKey) {
          throw new Error("Cloud-only mode memerlukan konfigurasi Supabase dalam config.js.");
        }
        state.students = await fetchStudentsFromSupabase(config, year);
        if (state.isManageOpen) {
          renderManageTable();
        }
        populateNilamUpdateClassOptions();
      }

      const rows = await readRowsFromFile(file);
      const parsed = parseImportedAinsRows(rows, year, month);
      if (!parsed.records.length) {
        throw new Error(
          "Tiada padanan data AINS ditemui. Pastikan Nama Murid sama dengan namelist."
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
        await upsertStudentsToSupabase(parsed.matchedStudents || [], { deactivateMissing: false });
        await overwriteAinsInSupabase(year, month, parsed.records, ainsOverrideRecords);
        syncNote = " Data AINS juga disimpan ke Supabase.";
      } catch (syncError) {
        if (CLOUD_ONLY_MODE) {
          throw syncError;
        }
        console.error(syncError);
        const detail = String(syncError?.message || syncError || "").slice(0, 220);
        syncNote = ` Simpanan Supabase gagal (${detail}).`;
      }

      const baseMessage = `Import Data AINS berjaya: ${parsed.records.length} rekod untuk ${year} ${month} telah di-override.`;
      if (!parsed.unmatchedCount) {
        setStatus(`${baseMessage}${syncNote}`);
      } else {
        const unmatchedList = (Array.isArray(parsed.unmatchedRows) ? parsed.unmatchedRows : [])
          .map(
            (row, idx) =>
              `<li>${idx + 1}. Baris ${Number(row.rowNumber) || "?"}: ${escapeAttr(row.nama || "(nama kosong)")} | ${escapeAttr(row.email || "(ID DELIMa kosong)")}</li>`
          )
          .join("");
        const html =
          `${escapeAttr(baseMessage)} ${parsed.unmatchedCount} baris tidak dipadankan (Nama Murid).${escapeAttr(syncNote)}` +
          `<details class="ains-unmatched-block" open>` +
          `<summary>Lihat senarai baris tidak dipadankan</summary>` +
          `<ul class="compare-list ains-unmatched-list">${unmatchedList}</ul>` +
          `</details>`;
        setStatusHtml(html);
      }
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

  async function handleNilamUpdateFilterChange() {
    if (!state.isNilamUpdateOpen) {
      return;
    }
    const year = String(el.nilamUpdateYearInput ? el.nilamUpdateYearInput.value : "").trim();
    const month = normalizeMonth(el.nilamUpdateMonthSelect ? el.nilamUpdateMonthSelect.value : "");
    if (!/^\d{4}$/.test(year) || !month) {
      renderNilamUpdateEmpty("Pilih Tahun, Bulan, dan Kelas untuk memuatkan data.");
      return;
    }
    populateNilamUpdateClassOptions();
    const kelas = String(el.nilamUpdateClassSelect ? el.nilamUpdateClassSelect.value : "").trim();
    if (!kelas) {
      renderNilamUpdateEmpty("Pilih Tahun, Bulan, dan Kelas untuk memuatkan data.");
      return;
    }
    await loadNilamUpdateDataSelection();
  }

  async function loadNilamUpdateDataSelection(options = {}) {
    try {
      const year = String(el.nilamUpdateYearInput ? el.nilamUpdateYearInput.value : "").trim();
      if (!/^\d{4}$/.test(year)) {
        throw new Error("Tahun tidak sah. Gunakan format 4 digit, contoh 2026.");
      }
      const month = normalizeMonth(el.nilamUpdateMonthSelect ? el.nilamUpdateMonthSelect.value : "");
      if (!month) {
        throw new Error("Bulan tidak sah. Sila pilih bulan daripada senarai.");
      }
      populateNilamUpdateClassOptions();
      const kelas = String(el.nilamUpdateClassSelect ? el.nilamUpdateClassSelect.value : "").trim();
      if (!kelas) {
        throw new Error("Sila pilih kelas dahulu.");
      }

      const config = window.NILAM_CONFIG || {};
      if (CLOUD_ONLY_MODE && (!config.supabaseUrl || !config.supabaseAnonKey)) {
        throw new Error("Cloud-only mode memerlukan konfigurasi Supabase dalam config.js.");
      }
      const localRecords = readAllLocalRecords();
      let supabaseRecords = [];
      let cloudWarning = "";
      if (config.supabaseUrl && config.supabaseAnonKey) {
        try {
          supabaseRecords = await fetchAllSupabaseRecordsForAdmin(config);
        } catch (error) {
          if (CLOUD_ONLY_MODE) {
            throw error;
          }
          console.error(error);
          cloudWarning = ` Muat cloud gagal (${String(error?.message || error || "").slice(0, 160)}).`;
        }
      }
      const allRecords = mergeRecordsByLatestSession(supabaseRecords, localRecords);
      const rows = buildNilamUpdateRows(allRecords, year, month, kelas);

      state.nilamUpdateRows = rows;
      state.nilamUpdateLoadedYear = year;
      state.nilamUpdateLoadedMonth = month;
      state.nilamUpdateLoadedClass = kelas;
      renderNilamUpdateTable();

      if (!options.silentStatus) {
        setStatus(`Data dimuatkan: ${rows.length} murid untuk ${year} ${month} ${kelas}.${cloudWarning}`);
      }
    } catch (error) {
      const message = error?.message || "Gagal memuatkan data Kemas Kini NILAM.";
      setStatus(message, true);
      renderNilamUpdateEmpty(message);
    }
  }

  function buildNilamUpdateRows(allRecords, year, month, kelas) {
    const records = Array.isArray(allRecords) ? allRecords : [];
    const classStudents = getCurrentNamelist()
      .map(normalizeStudentRow)
      .filter((row) => String(row.kelas || "").trim() === kelas);

    const byKey = new Map();
    classStudents.forEach((student) => {
      const key = getStudentAggregateKey(student, kelas);
      if (!key || byKey.has(key)) {
        return;
      }
      byKey.set(key, emptyNilamUpdateRow(student, key, kelas));
    });

    records.forEach((row) => {
      if (String(row?.tahun || "").trim() !== year) {
        return;
      }
      if (String(row?.bulan || "").trim() !== month) {
        return;
      }
      if (String(row?.kelas || "").trim() !== kelas) {
        return;
      }
      if (isPureAinsRecord(row)) {
        return;
      }
      const key = getStudentAggregateKey(row, kelas);
      if (!key) {
        return;
      }
      if (!byKey.has(key)) {
        byKey.set(key, emptyNilamUpdateRow(row, key, kelas));
      }
      const slot = byKey.get(key);
      slot.no_kad_pengenalan = normalizeNoKad(slot.no_kad_pengenalan) || normalizeNoKad(row.no_kad_pengenalan);
      slot.nama = slot.nama || String(row.nama || "").trim();
      slot.kelas = slot.kelas || String(row.kelas || "").trim() || kelas;
      slot.bahan_digital += toInt999(row.bahan_digital);
      slot.bahan_bukan_buku += toInt999(row.bahan_bukan_buku);
      slot.fiksyen += toInt999(row.fiksyen);
      slot.bukan_fiksyen += toInt999(row.bukan_fiksyen);
      slot.bahasa_melayu += toInt999(row.bahasa_melayu);
      slot.bahasa_inggeris += toInt999(row.bahasa_inggeris);
      slot.lain_lain_bahasa += toInt999(row.lain_lain_bahasa);
    });

    const yearRecords = records.filter((row) => String(row?.tahun || "").trim() === year);
    const yearMaterialsByKey = computeMaterialsTotalsByKey(yearRecords);
    const allMaterialsByKey = computeMaterialsTotalsByKey(records);
    const yearAinsByKey = computeYearAinsMaxByKey(yearRecords);
    const allTimeAinsByKey = computeAllTimeAinsSumByKey(records);

    const rows = [...byKey.values()].sort((a, b) =>
      String(a.nama || "").localeCompare(String(b.nama || ""), "ms", { sensitivity: "base" })
    );

    rows.forEach((row, index) => {
      row.bil = index + 1;
      recomputeNilamUpdateRow(row);
      row.ains_sepanjang_tahun = yearAinsByKey.get(row.student_key) || 0;
      row.ains_all_time = allTimeAinsByKey.get(row.student_key) || 0;
      row.year_materials_base = Math.max(0, (yearMaterialsByKey.get(row.student_key) || 0) - row.jumlah_aktiviti);
      row.all_materials_base = Math.max(0, (allMaterialsByKey.get(row.student_key) || 0) - row.jumlah_aktiviti);
      row.jumlah_tahun = row.year_materials_base + row.jumlah_aktiviti + row.ains_sepanjang_tahun;
      row.jumlah_all_time = row.all_materials_base + row.jumlah_aktiviti + row.ains_all_time;
    });

    return rows;
  }

  function emptyNilamUpdateRow(source, key, kelasFallback) {
    return {
      bil: 0,
      student_key: key,
      no_kad_pengenalan: normalizeNoKad(source?.no_kad_pengenalan),
      nama: String(source?.nama || "").trim(),
      kelas: String(source?.kelas || "").trim() || kelasFallback,
      bahan_digital: 0,
      bahan_bukan_buku: 0,
      fiksyen: 0,
      bukan_fiksyen: 0,
      bahasa_melayu: 0,
      bahasa_inggeris: 0,
      lain_lain_bahasa: 0,
      jumlah_aktiviti: 0,
      ains_sepanjang_tahun: 0,
      ains_all_time: 0,
      year_materials_base: 0,
      all_materials_base: 0,
      jumlah_tahun: 0,
      jumlah_all_time: 0,
    };
  }

  function getStudentAggregateKey(row, kelasFallback) {
    const noKad = normalizeNoKad(row?.no_kad_pengenalan);
    if (noKad) {
      return `ic:${noKad}`;
    }
    const nama = String(row?.nama || "").trim().toLowerCase();
    const kelas = String(row?.kelas || "").trim().toLowerCase() || String(kelasFallback || "").trim().toLowerCase();
    if (!nama || !kelas) {
      return "";
    }
    return `nm:${nama}|k:${kelas}`;
  }

  function computeMaterialsTotalsByKey(records) {
    const map = new Map();
    (Array.isArray(records) ? records : []).forEach((row) => {
      const key = getStudentAggregateKey(row);
      if (!key) {
        return;
      }
      map.set(key, (map.get(key) || 0) + getMaterialsTotalWithoutAins(row));
    });
    return map;
  }

  function computeYearAinsMaxByKey(records) {
    const map = new Map();
    (Array.isArray(records) ? records : []).forEach((row) => {
      const key = getStudentAggregateKey(row);
      if (!key) {
        return;
      }
      const next = toInt999(row.ains);
      if (!map.has(key) || next > map.get(key)) {
        map.set(key, next);
      }
    });
    return map;
  }

  function computeAllTimeAinsSumByKey(records) {
    const byStudentYear = new Map();
    (Array.isArray(records) ? records : []).forEach((row) => {
      const key = getStudentAggregateKey(row);
      const year = String(row?.tahun || "").trim();
      if (!key || !year) {
        return;
      }
      const yearKey = `${key}|${year}`;
      const next = toInt999(row.ains);
      if (!byStudentYear.has(yearKey) || next > byStudentYear.get(yearKey)) {
        byStudentYear.set(yearKey, next);
      }
    });
    const totals = new Map();
    byStudentYear.forEach((ains, yearKey) => {
      const sep = yearKey.lastIndexOf("|");
      const studentKey = sep >= 0 ? yearKey.slice(0, sep) : yearKey;
      totals.set(studentKey, (totals.get(studentKey) || 0) + ains);
    });
    return totals;
  }

  function getMaterialsTotalWithoutAins(row) {
    return (
      toInt999(row?.bahan_digital) +
      toInt999(row?.bahan_bukan_buku) +
      toInt999(row?.fiksyen) +
      toInt999(row?.bukan_fiksyen)
    );
  }

  function recomputeNilamUpdateRow(row) {
    row.bahan_digital = toInt999(row.bahan_digital);
    row.bahan_bukan_buku = toInt999(row.bahan_bukan_buku);
    row.fiksyen = toInt999(row.fiksyen);
    row.bukan_fiksyen = toInt999(row.bukan_fiksyen);
    row.bahasa_melayu = toInt999(row.bahasa_melayu);
    row.bahasa_inggeris = toInt999(row.bahasa_inggeris);
    row.lain_lain_bahasa = toInt999(row.lain_lain_bahasa);
    row.jumlah_aktiviti = getMaterialsTotalWithoutAins(row);
    row.jumlah_tahun = row.year_materials_base + row.jumlah_aktiviti + row.ains_sepanjang_tahun;
    row.jumlah_all_time = row.all_materials_base + row.jumlah_aktiviti + row.ains_all_time;
  }

  function renderNilamUpdateTable() {
    if (!el.nilamUpdateTbody) {
      return;
    }
    const rows = Array.isArray(state.nilamUpdateRows) ? state.nilamUpdateRows : [];
    if (!rows.length) {
      renderNilamUpdateEmpty("Tiada murid dijumpai untuk kelas dipilih.");
      return;
    }

    el.nilamUpdateTbody.innerHTML = rows
      .map((row, index) => `
        <tr data-row-index="${index}">
          <td>${row.bil}</td>
          <td><span class="cell-static cell-name">${escapeAttr(row.nama)}</span></td>
          <td><span class="cell-static">${escapeAttr(row.kelas)}</span></td>
          <td><input type="number" min="0" max="999" step="1" data-row-index="${index}" data-field="bahan_digital" value="${row.bahan_digital}"></td>
          <td><input type="number" min="0" max="999" step="1" data-row-index="${index}" data-field="bahan_bukan_buku" value="${row.bahan_bukan_buku}"></td>
          <td><input type="number" min="0" max="999" step="1" data-row-index="${index}" data-field="fiksyen" value="${row.fiksyen}"></td>
          <td><input type="number" min="0" max="999" step="1" data-row-index="${index}" data-field="bukan_fiksyen" value="${row.bukan_fiksyen}"></td>
          <td><input type="number" min="0" max="999" step="1" data-row-index="${index}" data-field="bahasa_melayu" value="${row.bahasa_melayu}"></td>
          <td><input type="number" min="0" max="999" step="1" data-row-index="${index}" data-field="bahasa_inggeris" value="${row.bahasa_inggeris}"></td>
          <td><input type="number" min="0" max="999" step="1" data-row-index="${index}" data-field="lain_lain_bahasa" value="${row.lain_lain_bahasa}"></td>
          <td><span class="cell-total" data-col="jumlah_aktiviti">${row.jumlah_aktiviti}</span></td>
          <td><span class="cell-total" data-col="ains_sepanjang_tahun">${row.ains_sepanjang_tahun}</span></td>
          <td><span class="cell-total" data-col="jumlah_tahun">${row.jumlah_tahun}</span></td>
          <td><span class="cell-total" data-col="jumlah_all_time">${row.jumlah_all_time}</span></td>
        </tr>
      `)
      .join("");
  }

  function handleNilamUpdateTableInput(event) {
    const input = event.target.closest("input[data-row-index][data-field]");
    if (!input) {
      return;
    }
    const rowIndex = Number(input.dataset.rowIndex);
    const field = String(input.dataset.field || "").trim();
    if (
      !Number.isInteger(rowIndex) ||
      rowIndex < 0 ||
      rowIndex >= state.nilamUpdateRows.length ||
      !field
    ) {
      return;
    }
    const row = state.nilamUpdateRows[rowIndex];
    row[field] = toInt999(input.value);
    input.value = String(row[field]);
    recomputeNilamUpdateRow(row);
    const tr = input.closest("tr");
    if (!tr) {
      return;
    }
    const monthCell = tr.querySelector('[data-col="jumlah_aktiviti"]');
    const yearCell = tr.querySelector('[data-col="jumlah_tahun"]');
    const allTimeCell = tr.querySelector('[data-col="jumlah_all_time"]');
    if (monthCell) {
      monthCell.textContent = String(row.jumlah_aktiviti);
    }
    if (yearCell) {
      yearCell.textContent = String(row.jumlah_tahun);
    }
    if (allTimeCell) {
      allTimeCell.textContent = String(row.jumlah_all_time);
    }
  }

  function isPureAinsRecord(row) {
    const ains = toInt999(row?.ains);
    const nonAinsTotal =
      getMaterialsTotalWithoutAins(row) +
      toInt999(row?.bahasa_melayu) +
      toInt999(row?.bahasa_inggeris) +
      toInt999(row?.lain_lain_bahasa);
    return ains > 0 && nonAinsTotal === 0;
  }

  function mergeRecordsByLatestSession(supabaseRecords, localRecords) {
    const map = new Map();
    let sequence = 0;
    const put = (row) => {
      const key = recordSessionKey(row);
      if (!key) {
        return;
      }
      sequence += 1;
      const nextTs = parseTimestamp(row?.updated_at_client);
      const current = map.get(key);
      if (!current || nextTs > current.ts || (nextTs === current.ts && sequence >= current.seq)) {
        map.set(key, { row, ts: nextTs, seq: sequence });
      }
    };
    (Array.isArray(supabaseRecords) ? supabaseRecords : []).forEach(put);
    (Array.isArray(localRecords) ? localRecords : []).forEach(put);
    return [...map.values()].map((entry) => entry.row);
  }

  async function fetchAllSupabaseRecordsForAdmin(config) {
    if (!config.supabaseUrl || !config.supabaseAnonKey) {
      return [];
    }
    const supabaseUrl = config.supabaseUrl.replace(/\/$/, "");
    const params = new URLSearchParams({
      select:
        "id,tahun,bulan,tarikh,nama_pengisi,guru,bil,no_kad_pengenalan,nama,kelas,bahan_digital,bahan_bukan_buku,fiksyen,bukan_fiksyen,ains,bahasa_melayu,bahasa_inggeris,lain_lain_bahasa,jumlah_aktiviti,updated_at_client",
      limit: "50000",
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
    const rows = await response.json();
    if (!Array.isArray(rows)) {
      return [];
    }
    const normalized = [];
    rows.forEach((row) => {
      const record = normalizeImportedRecord(row, "");
      if (record) {
        normalized.push(record);
      }
    });
    return normalized;
  }

  function hasAnyNilamInput(row) {
    return (
      toInt999(row?.bahan_digital) > 0 ||
      toInt999(row?.bahan_bukan_buku) > 0 ||
      toInt999(row?.fiksyen) > 0 ||
      toInt999(row?.bukan_fiksyen) > 0 ||
      toInt999(row?.bahasa_melayu) > 0 ||
      toInt999(row?.bahasa_inggeris) > 0 ||
      toInt999(row?.lain_lain_bahasa) > 0
    );
  }

  function isTargetPeriodClass(row, year, month, kelas) {
    return (
      String(row?.tahun || "").trim() === String(year || "").trim() &&
      String(row?.bulan || "").trim() === String(month || "").trim() &&
      String(row?.kelas || "").trim() === String(kelas || "").trim()
    );
  }

  async function saveNilamUpdateData() {
    try {
      const year = String(el.nilamUpdateYearInput ? el.nilamUpdateYearInput.value : "").trim();
      const month = normalizeMonth(el.nilamUpdateMonthSelect ? el.nilamUpdateMonthSelect.value : "");
      const kelas = String(el.nilamUpdateClassSelect ? el.nilamUpdateClassSelect.value : "").trim();
      if (!/^\d{4}$/.test(year) || !month || !kelas) {
        throw new Error("Sila pilih Tahun, Bulan, dan Kelas yang sah.");
      }
      if (
        year !== state.nilamUpdateLoadedYear ||
        month !== state.nilamUpdateLoadedMonth ||
        kelas !== state.nilamUpdateLoadedClass
      ) {
        throw new Error("Penapis telah berubah. Sila pilih semula Kelas untuk muat data terkini.");
      }

      const nowIso = new Date().toISOString();
      const tarikh = defaultTarikhForPeriod(year, month);
      const rowsForSave = state.nilamUpdateRows
        .map((row) => ({ ...row }))
        .filter((row) => hasAnyNilamInput(row));
      rowsForSave.sort((a, b) =>
        String(a.nama || "").localeCompare(String(b.nama || ""), "ms", { sensitivity: "base" })
      );

      const overrideRecords = rowsForSave.map((row, index) => {
        const safeNoKad = normalizeNoKad(row.no_kad_pengenalan) || syntheticNoKad(row);
        const bahanDigital = toInt999(row.bahan_digital);
        const bahanBukanBuku = toInt999(row.bahan_bukan_buku);
        const fiksyen = toInt999(row.fiksyen);
        const bukanFiksyen = toInt999(row.bukan_fiksyen);
        const bahasaMelayu = toInt999(row.bahasa_melayu);
        const bahasaInggeris = toInt999(row.bahasa_inggeris);
        const lainLainBahasa = toInt999(row.lain_lain_bahasa);
        return {
          tahun: year,
          bulan: month,
          tarikh,
          nama_pengisi: "ADMIN_OVERRIDE",
          guru: "Nilam",
          bil: index + 1,
          no_kad_pengenalan: safeNoKad,
          nama: String(row.nama || "").trim(),
          kelas: String(row.kelas || "").trim() || kelas,
          bahan_digital: bahanDigital,
          bahan_bukan_buku: bahanBukanBuku,
          fiksyen,
          bukan_fiksyen: bukanFiksyen,
          ains: 0,
          bahasa_melayu: bahasaMelayu,
          bahasa_inggeris: bahasaInggeris,
          lain_lain_bahasa: lainLainBahasa,
          jumlah_aktiviti: bahanDigital + bahanBukanBuku + fiksyen + bukanFiksyen,
          updated_at_client: nowIso,
        };
      });

      const localRecords = readAllLocalRecords();
      const keptLocal = localRecords.filter((row) => {
        if (!isTargetPeriodClass(row, year, month, kelas)) {
          return true;
        }
        return isPureAinsRecord(row);
      });
      persistLocalRecords([...keptLocal, ...overrideRecords]);

      let syncNote = "";
      const config = window.NILAM_CONFIG || {};
      if (config.supabaseUrl && config.supabaseAnonKey) {
        try {
          const studentsForSync = dedupeStudentsForSync(
            state.nilamUpdateRows.map((row) =>
              normalizeStudentRow({
                nama: row.nama,
                kelas: row.kelas || kelas,
                no_kad_pengenalan: normalizeNoKad(row.no_kad_pengenalan) || syntheticNoKad(row),
              })
            )
          );
          await upsertStudentsToSupabase(studentsForSync, { deactivateMissing: false });
          await overwriteNilamUpdateRecordsInSupabase(year, month, kelas, overrideRecords, config);
          syncNote = " Data juga disimpan ke Supabase.";
        } catch (syncError) {
          if (CLOUD_ONLY_MODE) {
            throw syncError;
          }
          console.error(syncError);
          const detail = String(syncError?.message || syncError || "").slice(0, 220);
          syncNote = ` Simpanan Supabase gagal (${detail}).`;
        }
      }

      await loadNilamUpdateDataSelection({ silentStatus: true });
      setStatus(
        `Kemas kini Data Nilam berjaya: ${rowsForSave.length} rekod disimpan untuk ${year} ${month} ${kelas}.${syncNote}`
      );
      showPopupStatus("Berjaya disimpan", false);
    } catch (error) {
      console.error(error);
      const message = error?.message || "Gagal simpan Kemas Kini Data Nilam.";
      setStatus(message, true);
      showPopupStatus(message, true);
    }
  }

  async function overwriteNilamUpdateRecordsInSupabase(year, month, kelas, overrideRecords, config) {
    const existingRows = await fetchSupabaseRecordsForPeriodClass(year, month, kelas, config);
    const idsToDelete = existingRows
      .filter((row) => !isPureAinsRecord(row))
      .map((row) => Number(row.id))
      .filter((id) => Number.isFinite(id) && id > 0);
    if (idsToDelete.length) {
      await deleteSupabaseRecordsByIds(idsToDelete, config);
    }

    const payload = dedupeRecordsForSupabase(overrideRecords).map((row) => {
      const clean = { ...row };
      delete clean.id;
      return clean;
    });
    if (!payload.length) {
      return;
    }

    const supabaseUrl = config.supabaseUrl.replace(/\/$/, "");
    const insertEndpoint = `${supabaseUrl}/rest/v1/nilam_records`;
    const insertRes = await fetch(insertEndpoint, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        apikey: config.supabaseAnonKey,
        Authorization: `Bearer ${config.supabaseAnonKey}`,
        Prefer: "return=minimal",
      },
      body: JSON.stringify(payload),
    });
    if (!insertRes.ok) {
      const detail = await insertRes.text();
      throw new Error(`Sync kemas kini Nilam ke Supabase gagal (${insertRes.status}): ${detail}`);
    }
  }

  async function fetchSupabaseRecordsForPeriodClass(year, month, kelas, config) {
    const supabaseUrl = config.supabaseUrl.replace(/\/$/, "");
    const params = new URLSearchParams({
      select:
        "id,tahun,bulan,tarikh,nama_pengisi,guru,bil,no_kad_pengenalan,nama,kelas,bahan_digital,bahan_bukan_buku,fiksyen,bukan_fiksyen,ains,bahasa_melayu,bahasa_inggeris,lain_lain_bahasa,jumlah_aktiviti,updated_at_client",
      tahun: `eq.${year}`,
      bulan: `eq.${month}`,
      kelas: `eq.${kelas}`,
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
      throw new Error(`Semakan rekod Supabase untuk kemas kini Nilam gagal (${response.status}): ${detail}`);
    }
    const rows = await response.json();
    const normalized = [];
    (Array.isArray(rows) ? rows : []).forEach((row) => {
      const clean = normalizeImportedRecord(row, "");
      if (clean) {
        normalized.push(clean);
      }
    });
    return normalized;
  }

  async function deleteSupabaseRecordsByIds(ids, config) {
    const supabaseUrl = config.supabaseUrl.replace(/\/$/, "");
    const batchSize = 300;
    for (let i = 0; i < ids.length; i += batchSize) {
      const batch = ids.slice(i, i + batchSize);
      const endpoint = `${supabaseUrl}/rest/v1/nilam_records?id=in.(${batch.join(",")})`;
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
        throw new Error(`Padam rekod Supabase untuk kemas kini Nilam gagal (${response.status}): ${detail}`);
      }
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
    for (let i = 0; i < localLength(); i += 1) {
      const key = localKey(i);
      if (key && key.startsWith("nilam_records_")) {
        localKeys.push(key);
      }
    }
    localKeys.forEach((key) => {
      localRemoveItem(key);
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

    setStatus(`Reset selesai. Data cloud telah dipadam. ${supabaseMessage}`);
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

    if (!namaColumn || !ainsColumn) {
      const availableHeaders = rows.length ? Object.keys(rows[0]).join(", ") : "(tiada header)";
      throw new Error(
        `Kolum fail AINS tidak lengkap. Perlu ada: NAMA MURID dan kolum Rekod. Header dikesan: ${availableHeaders}`
      );
    }

    const namelist = getCurrentNamelist().map(normalizeStudentRow);
    const studentByMatchKey = new Map();
    namelist.forEach((student) => {
      const matchKey = toStudentMatchKey(student.nama);
      if (!matchKey) {
        return;
      }
      if (!studentByMatchKey.has(matchKey)) {
        studentByMatchKey.set(matchKey, student);
      }
    });

    const nowIso = new Date().toISOString();
    const records = [];
    const matchedStudents = [];
    let unmatchedCount = 0;
    const unmatchedRows = [];
    rows.forEach((row, index) => {
      const namaCsv = String(row[namaColumn] || "").trim();
      const emailCsv = String(emailColumn ? row[emailColumn] : "").trim();
      const matchKey = toStudentMatchKey(namaCsv);
      if (!matchKey) {
        return;
      }

      const matchedStudent = studentByMatchKey.get(matchKey);
      if (!matchedStudent || !matchedStudent.kelas || !matchedStudent.nama) {
        unmatchedCount += 1;
        unmatchedRows.push({
          rowNumber: index + 2,
          nama: namaCsv,
          email: emailCsv,
        });
        return;
      }

      matchedStudents.push(
        normalizeStudentRow({
          nama: matchedStudent.nama,
          kelas: matchedStudent.kelas,
          no_kad_pengenalan: matchedStudent.no_kad_pengenalan || "",
          jantina: matchedStudent.jantina || "",
          email_google_classroom: matchedStudent.email_google_classroom || "",
        })
      );

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
      unmatchedRows,
      matchedStudents: dedupeStudentsForSync(matchedStudents),
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

  function toStudentMatchKey(nama) {
    const namaNorm = String(nama || "")
      .trim()
      .toLowerCase()
      .replace(/\s+/g, " ");
    if (!namaNorm) {
      return "";
    }
    return namaNorm;
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

  function dedupeStudentsForSync(students) {
    const map = new Map();
    (Array.isArray(students) ? students : []).forEach((row) => {
      const normalized = normalizeStudentRow(row);
      if (!normalized.nama || !normalized.kelas) {
        return;
      }
      const noKad = normalizeNoKad(normalized.no_kad_pengenalan);
      const key = noKad || `nm:${normalized.nama.toLowerCase()}|k:${normalized.kelas.toLowerCase()}`;
      if (!map.has(key)) {
        map.set(key, normalized);
      }
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
    for (let i = 0; i < localLength(); i += 1) {
      const key = localKey(i);
      if (!key || !key.startsWith("nilam_records_")) {
        continue;
      }
      const raw = localGetItem(key);
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
    for (let i = 0; i < localLength(); i += 1) {
      const key = localKey(i);
      if (key && key.startsWith("nilam_records_")) {
        keysToDelete.push(key);
      }
    }
    keysToDelete.forEach((key) => localRemoveItem(key));

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
      localSetItem(key, JSON.stringify(rows));
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
      localSetItem(NAMELIST_OVERRIDE_KEY, JSON.stringify(merged));
      state.students = merged;
      if (state.isManageOpen) {
        renderManageTable();
      }
    }
  }

  async function upsertImportedRecordsToSupabase(records) {
    const config = window.NILAM_CONFIG || {};
    if (!records.length) {
      return;
    }
    if (!config.supabaseUrl || !config.supabaseAnonKey) {
      if (CLOUD_ONLY_MODE) {
        throw new Error("Cloud-only mode memerlukan konfigurasi Supabase dalam config.js.");
      }
      return;
    }

    const supabaseUrl = config.supabaseUrl.replace(/\/$/, "");
    const endpoint =
      `${supabaseUrl}/rest/v1/nilam_records?on_conflict=tahun,bulan,tarikh,no_kad_pengenalan,nama_pengisi,guru`;
    const payload = dedupeRecordsForSupabase(records);
    try {
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
      const isUpsertDuplicateBatch =
        response.status === 500 &&
        (String(detail || "").includes("21000") ||
          String(detail || "").toLowerCase().includes("cannot affect row a second time"));
      if (isNoConflictConstraint || isUpsertDuplicateBatch) {
        await replaceMonthlyAdminImportRecordsInSupabase(payload, config);
        return;
      }
      if (!isNoConflictConstraint && !isUpsertDuplicateBatch) {
        throw new Error(`Sync import data ke Supabase gagal (${response.status}): ${detail}`);
      }
    } catch (error) {
      const message = String(error && error.message ? error.message : error || "");
      if (!message.includes("42P10") && !message.includes("21000")) {
        throw error;
      }
      await replaceMonthlyAdminImportRecordsInSupabase(payload, config);
      return;
    }
  }

  async function replaceMonthlyAdminImportRecordsInSupabase(records, config) {
    const safeRecords = dedupeRecordsForSupabase(records);
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
    if (!records.length) {
      return;
    }
    if (!config.supabaseUrl || !config.supabaseAnonKey) {
      if (CLOUD_ONLY_MODE) {
        throw new Error("Cloud-only mode memerlukan konfigurasi Supabase dalam config.js.");
      }
      return;
    }

    const supabaseUrl = config.supabaseUrl.replace(/\/$/, "");
    const deduped = dedupeRecordsForSupabase(records);
    const payloadWithId = [];
    const payloadWithoutId = [];
    deduped.forEach((row) => {
      const clean = { ...row, bil: toPositiveInt(row?.bil, 1) };
      const id = Number(clean.id);
      if (Number.isFinite(id) && id > 0) {
        clean.id = id;
        payloadWithId.push(clean);
      } else {
        delete clean.id;
        payloadWithoutId.push(clean);
      }
    });

    if (payloadWithId.length) {
      const updateEndpoint = `${supabaseUrl}/rest/v1/nilam_records?on_conflict=id`;
      const updateRes = await fetch(updateEndpoint, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          apikey: config.supabaseAnonKey,
          Authorization: `Bearer ${config.supabaseAnonKey}`,
          Prefer: "resolution=merge-duplicates,return=minimal",
        },
        body: JSON.stringify(payloadWithId),
      });
      if (!updateRes.ok) {
        const detail = await updateRes.text();
        throw new Error(`Sync overwrite AINS (update) ke Supabase gagal (${updateRes.status}): ${detail}`);
      }
    }

    if (payloadWithoutId.length) {
      const insertEndpoint = `${supabaseUrl}/rest/v1/nilam_records`;
      const insertRes = await fetch(insertEndpoint, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          apikey: config.supabaseAnonKey,
          Authorization: `Bearer ${config.supabaseAnonKey}`,
          Prefer: "return=minimal",
        },
        body: JSON.stringify(payloadWithoutId),
      });
      if (!insertRes.ok) {
        const detail = await insertRes.text();
        throw new Error(`Sync overwrite AINS (insert) ke Supabase gagal (${insertRes.status}): ${detail}`);
      }
    }
  }

  async function overwriteAinsInSupabase(year, month, importedAinsRows, ainsOverrideRecords) {
    const config = window.NILAM_CONFIG || {};
    if (!config.supabaseUrl || !config.supabaseAnonKey) {
      if (CLOUD_ONLY_MODE) {
        throw new Error("Cloud-only mode memerlukan konfigurasi Supabase dalam config.js.");
      }
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

  function dedupeRecordsForSupabase(records) {
    const byKey = new Map();
    (Array.isArray(records) ? records : []).forEach((row, index) => {
      const key = recordSessionKey(row);
      if (!key) {
        return;
      }
      const nextRow = {
        ...row,
        bil: toPositiveInt(row?.bil, 1),
      };
      const nextTs = parseTimestamp(nextRow.updated_at_client);
      const current = byKey.get(key);
      if (!current || nextTs > current.ts || (nextTs === current.ts && index >= current.index)) {
        byKey.set(key, { row: nextRow, ts: nextTs, index });
      }
    });
    return [...byKey.values()].map((entry) => entry.row);
  }

  function parseTimestamp(value) {
    const ms = Date.parse(String(value || ""));
    return Number.isFinite(ms) ? ms : 0;
  }

  // Generates a deterministic synthetic IC for students without one.
  // Prefixed with NOIC_ so it's distinguishable from real IC numbers.
  function syntheticNoKad(row) {
    const kelas = String(row.kelas || "").trim().toUpperCase().replace(/\s+/g, "");
    const nama = String(row.nama || "").trim().toUpperCase().replace(/[^A-Z0-9]/g, "_").slice(0, 30);
    return `NOIC_${kelas}_${nama}`;
  }

  async function fetchStudentsFromSupabase(config, year) {
    if (!config.supabaseUrl || !config.supabaseAnonKey) {
      if (CLOUD_ONLY_MODE) {
        throw new Error("Cloud-only mode memerlukan konfigurasi Supabase dalam config.js.");
      }
      return [];
    }

    const safeYear = /^\d{4}$/.test(String(year || "")) ? String(year) : String(new Date().getFullYear());
    const kelasField = `kelas_${safeYear}`;
    const supabaseUrl = config.supabaseUrl.replace(/\/$/, "");
    const params = new URLSearchParams({
      select: `no_kad_pengenalan,nama_murid,jantina,email_google_classroom,${kelasField},kelas_2026,kelas_2027,kelas_2028,kelas_2029,kelas_2030,kelas_2031`,
      active: "neq.false",
      order: `${kelasField}.asc,nama_murid.asc`,
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
      throw new Error(`Ralat muat senarai murid Supabase (${response.status}): ${detail}`);
    }

    const rows = await response.json();
    if (!Array.isArray(rows)) {
      return [];
    }

    const fallbackClassFields = [
      kelasField,
      "kelas_2026",
      "kelas_2027",
      "kelas_2028",
      "kelas_2029",
      "kelas_2030",
      "kelas_2031",
    ];

    return rows
      .map((row) =>
        normalizeStudentRow({
          nama: row.nama_murid,
          jantina: row.jantina,
          kelas: fallbackClassFields.map((field) => row[field]).find((value) => String(value || "").trim()) || "",
          no_kad_pengenalan: row.no_kad_pengenalan,
          email_google_classroom: row.email_google_classroom,
        })
      )
      .filter((row) => row.nama && row.kelas);
  }

  async function upsertStudentsToSupabase(students, options = {}) {
    const deactivateMissing = options.deactivateMissing !== false;
    const config = window.NILAM_CONFIG || {};
    if (!config.supabaseUrl || !config.supabaseAnonKey) {
      if (CLOUD_ONLY_MODE) {
        throw new Error("Cloud-only mode memerlukan konfigurasi Supabase dalam config.js.");
      }
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

    for (let i = 0; i < localLength(); i += 1) {
      const key = localKey(i);
      if (!key || !key.startsWith("nilam_records_")) {
        continue;
      }
      const parsed = parseStoredRecords(localGetItem(key));
      if (!parsed.length) {
        continue;
      }
      const kept = parsed.filter((row) => isRecordAllowedByMaster(row, allowedByClass));
      if (kept.length === parsed.length) {
        continue;
      }
      deletedCount += parsed.length - kept.length;
      if (kept.length) {
        localSetItem(key, JSON.stringify(kept));
      } else {
        localRemoveItem(key);
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
    const memoryList = (Array.isArray(state.students) ? state.students : [])
      .map(normalizeStudentRow)
      .filter((row) => row.nama && row.kelas);
    if (memoryList.length) {
      return memoryList;
    }
    if (CLOUD_ONLY_MODE) {
      return [];
    }

    const overrideRaw = localGetItem(NAMELIST_OVERRIDE_KEY);
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
    const raw = localGetItem(TEACHER_NAMES_KEY);
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
      .filter(Boolean)
      .sort((a, b) => a.localeCompare(b, "ms"));
  }

  async function upsertTeacherNamesToSupabase(names) {
    const config = window.NILAM_CONFIG || {};
    if (!config.supabaseUrl || !config.supabaseAnonKey) {
      if (CLOUD_ONLY_MODE) {
        throw new Error("Cloud-only mode memerlukan konfigurasi Supabase dalam config.js.");
      }
      return;
    }
    const list = [...new Set((Array.isArray(names) ? names : []).map((name) => String(name || "").trim()).filter(Boolean))];
    if (!list.length) {
      return;
    }

    const supabaseUrl = config.supabaseUrl.replace(/\/$/, "");
    const endpoint = `${supabaseUrl}/rest/v1/nilam_teachers?on_conflict=nama_guru`;
    const chunkSize = 500;
    for (let i = 0; i < list.length; i += chunkSize) {
      const chunk = list.slice(i, i + chunkSize).map((name) => ({ nama_guru: name }));
      const response = await fetch(endpoint, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          apikey: config.supabaseAnonKey,
          Authorization: `Bearer ${config.supabaseAnonKey}`,
          Prefer: "resolution=merge-duplicates,return=minimal",
        },
        body: JSON.stringify(chunk),
      });
      if (!response.ok) {
        const detail = await response.text();
        throw new Error(`Ralat simpan nama guru ke Supabase (${response.status}): ${detail}`);
      }
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

  function setStatusHtml(messageHtml, isError) {
    el.status.innerHTML = messageHtml;
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
      localSetItem(NAMELIST_OVERRIDE_KEY, JSON.stringify(finalList));
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
        `Senarai nama dikemas kini: ${finalList.length} murid. Data murid yang dibuang telah dipadam di cloud: ${purgeResult.supabaseDeleted}.`
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
