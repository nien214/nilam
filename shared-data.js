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
    "bahasa_melayu",
    "bahasa_inggeris",
    "lain_lain_bahasa",
  ];
  const NAMELIST_OVERRIDE_KEY = "nilam_students_override_v1";

  function normalizeText(value) {
    return String(value || "").trim();
  }

  async function getStudents(config) {
    const selectedYear = String(new Date().getFullYear());

    // Supabase is always authoritative — active=neq.false filtered server-side.
    // No-IC students are stored with synthetic NOIC_ IDs in Supabase.
    try {
      const supabaseStudents = await getSupabaseStudents(config, selectedYear);
      const normalized = normalizeStudents(supabaseStudents);
      if (normalized.length) {
        return normalized;
      }
    } catch (error) {
      console.error("Gagal muat senarai murid dari Supabase", error);
    }

    // Offline fallback: localStorage override, then bundled data.
    let localRaw = [];
    const overrideRaw = localStorage.getItem(NAMELIST_OVERRIDE_KEY);
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

    records.forEach((row) => {
      const tahun = normalizeText(row.tahun) || String(new Date().getFullYear());
      const bulan = normalizeText(row.bulan);
      const kelas = normalizeText(row.kelas);
      const nama = normalizeText(row.nama);
      const noKad = normalizeText(row.no_kad_pengenalan);
      if (!tahun || !bulan || !kelas || !nama) {
        return;
      }

      const key = noKad
        ? `${tahun}|${bulan}|${noKad}`
        : `${tahun}|${bulan}|${kelas}|${nama}`;
      const current = byKey.get(key);
      if (!current || parseTimestamp(row.updated_at_client) >= parseTimestamp(current.updated_at_client)) {
        byKey.set(key, row);
      }
    });

    return [...byKey.values()];
  }

  function normalizeRecord(row) {
    const out = {
      tahun: normalizeText(row.tahun) || String(new Date().getFullYear()),
      bulan: normalizeText(row.bulan),
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
      out.jumlah_aktiviti = out.bahasa_melayu + out.bahasa_inggeris + out.lain_lain_bahasa;
    }

    return out;
  }

  function rowHasAnyInput(row) {
    return INPUT_COLUMNS.some((col) => Number(row[col]) > 0);
  }

  function getLocalRecords() {
    const records = [];
    const prefix = "nilam_records_";

    for (let i = 0; i < localStorage.length; i += 1) {
      const key = localStorage.key(i);
      if (!key || !key.startsWith(prefix)) {
        continue;
      }

      const raw = localStorage.getItem(key);
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
        "tahun,bulan,kelas,nama,no_kad_pengenalan,bahan_digital,bahan_bukan_buku,fiksyen,bukan_fiksyen,bahasa_melayu,bahasa_inggeris,lain_lain_bahasa,jumlah_aktiviti,updated_at_client",
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
      throw new Error(`Ralat muat data Supabase (${response.status}): ${detail}`);
    }

    const rows = await response.json();
    if (!Array.isArray(rows)) {
      return [];
    }

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
      throw new Error(`Ralat muat senarai murid Supabase (${response.status}): ${detail}`);
    }

    const rows = await response.json();
    if (!Array.isArray(rows)) {
      return [];
    }

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
    }

    const allowedByClass = buildAllowedKeysByClass(students);
    const alignedRecords = alignRecordsToCurrentClass(
      dedupeRecords([...supabaseRecords, ...localRecords]),
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
