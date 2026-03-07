(function () {
  "use strict";

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

  const INPUT_COLUMNS = [
    "bahan_digital",
    "bahan_bukan_buku",
    "fiksyen",
    "bukan_fiksyen",
    "ains",
    "bahasa_melayu",
    "bahasa_inggeris",
    "lain_lain_bahasa",
  ];
  const NAMELIST_OVERRIDE_KEY = "nilam_students_override_v1";


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

  function normalizeText(value) {
    return String(value || "").trim();
  }

  function normalizeNameKey(value) {
    return normalizeText(value).toLowerCase();
  }

  function normalizeGuruType(value) {
    const clean = normalizeText(value);
    if (clean === "BM" || clean === "BI" || clean === "Nilam") {
      return clean;
    }
    return "Nilam";
  }

  function normalizeDateInput(value) {
    const clean = normalizeText(value);
    return /^\d{4}-\d{2}-\d{2}$/.test(clean) ? clean : "";
  }

  function normalizeRecordDate(rawDate, tahun, bulan) {
    const normalized = normalizeDateInput(rawDate);
    if (normalized) {
      return normalized;
    }
    const monthIndex = MONTHS.findIndex((m) => m.toLowerCase() === normalizeText(bulan).toLowerCase());
    if (/^\d{4}$/.test(String(tahun || "")) && monthIndex >= 0) {
      return `${tahun}-${String(monthIndex + 1).padStart(2, "0")}-01`;
    }
    if (/^\d{4}$/.test(String(tahun || ""))) {
      return `${tahun}-01-01`;
    }
    return "";
  }

  function getRecordSessionKey(row) {
    const tahun = normalizeText(row.tahun) || String(new Date().getFullYear());
    const bulan = normalizeText(row.bulan);
    const noKad = normalizeText(row.no_kad_pengenalan);
    const kelas = normalizeText(row.kelas);
    const nama = normalizeText(row.nama);
    const tarikh = normalizeRecordDate(row.tarikh, tahun, bulan);
    const namaPengisi = normalizeNameKey(row.nama_pengisi) || "tidak diketahui";
    const guru = normalizeGuruType(row.guru);
    if (!tahun || !bulan || !kelas || !nama || !tarikh || !guru) {
      return "";
    }
    if (noKad) {
      return `${tahun}|${bulan}|${tarikh}|${noKad}|${namaPengisi}|${guru}`;
    }
    return `${tahun}|${bulan}|${tarikh}|${kelas}|${nama.toLowerCase()}|${namaPengisi}|${guru}`;
  }

  async function getStudents(config) {
    const selectedYear = String(new Date().getFullYear());
    if (!config.supabaseUrl || !config.supabaseAnonKey) {
      if (CLOUD_ONLY_MODE) {
        throw new Error("Cloud-only mode memerlukan konfigurasi Supabase.");
      }
      return [];
    }

    // Supabase is always authoritative — active=neq.false filtered server-side.
    // No-IC students are stored with synthetic NOIC_ IDs in Supabase.
    try {
      const supabaseStudents = await getSupabaseStudents(config, selectedYear);
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
    let localRaw = [];
    const overrideRaw = localGetItem(NAMELIST_OVERRIDE_KEY);
    if (overrideRaw) {
      try {
        const parsed = JSON.parse(overrideRaw);
        if (Array.isArray(parsed)) {
          localRaw = parsed;
        }
      } catch (error) {
        console.error("Gagal parse namelist override", error);
      }
    }
    if (!localRaw.length) {
      localRaw = Array.isArray(window.NILAM_STUDENTS) ? window.NILAM_STUDENTS : [];
    }
    return normalizeStudents(localRaw);
  }

  function normalizeStudents(raw) {
    return (Array.isArray(raw) ? raw : [])
      .map((row) => ({
        nama: normalizeText(row.nama),
        kelas: normalizeText(row.kelas),
        no_kad_pengenalan: normalizeText(row.no_kad_pengenalan),
        jantina: normalizeText(row.jantina),
        email_google_classroom: normalizeText(row.email_google_classroom),
      }))
      .filter((row) => row.nama && row.kelas);
  }

  function getStudentsByClass(students) {
    const map = new Map();
    students.forEach((row) => {
      if (!map.has(row.kelas)) {
        map.set(row.kelas, []);
      }
      map.get(row.kelas).push(row.nama);
    });

    for (const [kelas, names] of map.entries()) {
      const uniqueNames = [...new Set(names)].sort((a, b) => a.localeCompare(b, "ms"));
      map.set(kelas, uniqueNames);
    }

    return map;
  }

  function buildAllowedKeysByClass(students) {
    const map = new Map();
    (Array.isArray(students) ? students : []).forEach((row) => {
      const kelas = normalizeText(row.kelas);
      const nama = normalizeText(row.nama).toLowerCase();
      const noKad = normalizeText(row.no_kad_pengenalan);
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

  function recordMatchesMaster(row, allowedByClass) {
    const kelas = normalizeText(row.kelas);
    const keys = allowedByClass.get(kelas);
    if (!keys || !keys.size) {
      return false;
    }
    const noKad = normalizeText(row.no_kad_pengenalan);
    if (noKad && keys.has(`ic:${noKad}`)) {
      return true;
    }
    const nama = normalizeText(row.nama).toLowerCase();
    return nama ? keys.has(`nm:${nama}`) : false;
  }

  function alignRecordsToCurrentClass(records, students) {
    const classByNoKad = new Map();
    const nameByNoKad = new Map();
    const nameClasses = new Map();

    (Array.isArray(students) ? students : []).forEach((row) => {
      const kelas = normalizeText(row.kelas);
      const nama = normalizeText(row.nama).toLowerCase();
      const noKad = normalizeText(row.no_kad_pengenalan);
      if (!kelas || !nama) {
        return;
      }
      if (noKad) {
        classByNoKad.set(noKad, kelas);
        nameByNoKad.set(noKad, nama);
      }
      if (!nameClasses.has(nama)) {
        nameClasses.set(nama, new Set());
      }
      nameClasses.get(nama).add(kelas);
    });

    return (Array.isArray(records) ? records : []).map((row) => {
      const noKad = normalizeText(row.no_kad_pengenalan);
      const nama = normalizeText(row.nama).toLowerCase();
      const byNoKadClass = noKad ? classByNoKad.get(noKad) : "";
      const byNoKadName = noKad ? nameByNoKad.get(noKad) : "";
      // Trust IC mapping only when IC-owner name matches this record name.
      if (byNoKadClass && byNoKadName && byNoKadName === nama) {
        return { ...row, kelas: byNoKadClass };
      }
      const classSet = nameClasses.get(nama);
      if (classSet && classSet.size === 1) {
        return { ...row, kelas: [...classSet][0] };
      }
      if (byNoKadClass) {
        return { ...row, kelas: byNoKadClass };
      }
      return row;
    });
  }

  function parseTimestamp(value) {
    const time = Date.parse(value || "");
    return Number.isFinite(time) ? time : 0;
  }

  function dedupeRecords(records) {
    const byKey = new Map();

    (Array.isArray(records) ? records : []).forEach((row, index) => {
      const key = getRecordSessionKey(row);
      if (!key) {
        return;
      }

      const nextTs = parseTimestamp(row.updated_at_client);
      const current = byKey.get(key);
      if (!current || nextTs > current.ts || (nextTs === current.ts && index >= current.index)) {
        byKey.set(key, { row, ts: nextTs, index });
      }
    });

    return [...byKey.values()].map((entry) => entry.row);
  }

  function mergeSupabaseAndLocalRecords(supabaseRecords, localRecords) {
    const byKey = new Map();
    dedupeRecords(localRecords).forEach((row) => {
      const key = getRecordSessionKey(row);
      if (!key) {
        return;
      }
      byKey.set(key, row);
    });
    dedupeRecords(supabaseRecords).forEach((row) => {
      const key = getRecordSessionKey(row);
      if (!key) {
        return;
      }
      // Supabase is authoritative for same session key.
      byKey.set(key, row);
    });
    return [...byKey.values()];
  }

  function normalizeRecord(row) {
    const tahun = normalizeText(row.tahun) || String(new Date().getFullYear());
    const bulan = normalizeText(row.bulan);
    const guru = normalizeGuruType(row.guru);
    const out = {
      tahun,
      bulan,
      tarikh: normalizeRecordDate(row.tarikh, tahun, bulan),
      nama_pengisi: normalizeText(row.nama_pengisi) || "Tidak Diketahui",
      guru,
      kelas: normalizeText(row.kelas),
      nama: normalizeText(row.nama),
      no_kad_pengenalan: normalizeText(row.no_kad_pengenalan),
      updated_at_client: row.updated_at_client || "",
    };

    INPUT_COLUMNS.forEach((col) => {
      const number = Number(row[col]);
      out[col] = Number.isFinite(number) ? Math.max(0, Math.trunc(number)) : 0;
    });

    const jumlah = Number(row.jumlah_aktiviti);
    if (Number.isFinite(jumlah)) {
      out.jumlah_aktiviti = Math.max(0, Math.trunc(jumlah));
    } else {
      out.jumlah_aktiviti =
        out.bahan_digital +
        out.bahan_bukan_buku +
        out.fiksyen +
        out.bukan_fiksyen +
        out.ains;
    }

    return out;
  }

  function rowHasAnyInput(row) {
    return INPUT_COLUMNS.some((col) => Number(row[col]) > 0);
  }

  async function fetchAllSupabaseRows(endpoint, headers, pageSize) {
    const rows = [];
    const step = Math.max(1, Math.trunc(Number(pageSize) || 1000));
    let offset = 0;

    // Supabase/PostgREST may cap each response to 1000 rows even when a larger limit is requested.
    while (true) {
      const separator = endpoint.includes("?") ? "&" : "?";
      const pagedEndpoint = `${endpoint}${separator}offset=${offset}&limit=${step}`;
      const response = await fetch(pagedEndpoint, {
        method: "GET",
        cache: "no-store",
        headers,
      });

      if (!response.ok) {
        const detail = await response.text();
        throw new Error(`Ralat muat data Supabase (${response.status}): ${detail}`);
      }

      const page = await response.json();
      if (!Array.isArray(page) || !page.length) {
        break;
      }

      rows.push(...page);
      if (page.length < step) {
        break;
      }
      offset += step;
    }

    return rows;
  }

  function getLocalRecords() {
    const records = [];
    const prefix = "nilam_records_";

    for (let i = 0; i < localLength(); i += 1) {
      const key = localKey(i);
      if (!key || !key.startsWith(prefix)) {
        continue;
      }

      const raw = localGetItem(key);
      if (!raw) {
        continue;
      }

      try {
        const parsed = JSON.parse(raw);
        if (Array.isArray(parsed)) {
          parsed.forEach((row) => records.push(normalizeRecord(row)));
        }
      } catch (error) {
        console.error("Gagal parse localStorage key", key, error);
      }
    }

    return records;
  }

  async function getSupabaseRecords(config) {
    if (!config.supabaseUrl || !config.supabaseAnonKey) {
      return [];
    }

    const supabaseUrl = config.supabaseUrl.replace(/\/$/, "");
    const params = new URLSearchParams({
      select:
        "tahun,bulan,tarikh,nama_pengisi,guru,kelas,nama,no_kad_pengenalan,bahan_digital,bahan_bukan_buku,fiksyen,bukan_fiksyen,ains,bahasa_melayu,bahasa_inggeris,lain_lain_bahasa,jumlah_aktiviti,updated_at_client",
      order: "updated_at_client.desc",
    });

    const endpoint = `${supabaseUrl}/rest/v1/nilam_records?${params.toString()}`;
    const rows = await fetchAllSupabaseRows(
      endpoint,
      {
        apikey: config.supabaseAnonKey,
        Authorization: `Bearer ${config.supabaseAnonKey}`,
      },
      1000
    );
    return rows.map(normalizeRecord);
  }

  async function getSupabaseStudents(config, year) {
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
    });
    const endpoint = `${supabaseUrl}/rest/v1/nilam_students?${params.toString()}`;
    const rows = await fetchAllSupabaseRows(
      endpoint,
      {
        apikey: config.supabaseAnonKey,
        Authorization: `Bearer ${config.supabaseAnonKey}`,
      },
      1000
    );

    return rows.map((row) => ({
      nama: normalizeText(row.nama_murid),
      kelas: normalizeText(row[kelasField]),
      no_kad_pengenalan: normalizeText(row.no_kad_pengenalan),
      jantina: normalizeText(row.jantina),
      email_google_classroom: normalizeText(row.email_google_classroom),
    }));
  }

  async function loadAllData() {
    const config = window.NILAM_CONFIG || {};
    const students = await getStudents(config);
    const studentsByClass = getStudentsByClass(students);

    const localRecords = getLocalRecords();
    let supabaseRecords = [];
    try {
      supabaseRecords = await getSupabaseRecords(config);
    } catch (error) {
      console.error(error);
      if (CLOUD_ONLY_MODE) {
        throw error;
      }
    }

    const allowedByClass = buildAllowedKeysByClass(students);
    const alignedRecords = alignRecordsToCurrentClass(
      mergeSupabaseAndLocalRecords(supabaseRecords, localRecords),
      students
    );
    const records = alignedRecords.filter((row) =>
      recordMatchesMaster(row, allowedByClass)
    );
    return {
      months: MONTHS,
      inputColumns: INPUT_COLUMNS,
      students,
      studentsByClass,
      records,
    };
  }

  window.NILAM_DATA = {
    MONTHS,
    INPUT_COLUMNS,
    NAMELIST_OVERRIDE_KEY,
    loadAllData,
    rowHasAnyInput,
  };
})();
