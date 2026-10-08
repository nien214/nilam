(function (root, factory) {
  "use strict";

  const api = factory();
  if (typeof module === "object" && module.exports) {
    module.exports = api;
  }
  if (root) {
    root.NILAM_EXPORT = api;
  }
})(typeof window !== "undefined" ? window : globalThis, function () {
  "use strict";

  function cleanText(value) {
    return String(value || "").trim();
  }

  function normalizedName(value) {
    return cleanText(value).toLocaleLowerCase("ms").replace(/\s+/g, " ");
  }

  function nonnegativeInteger(value) {
    const number = Number(value);
    return Number.isFinite(number) && number > 0 ? Math.trunc(number) : 0;
  }

  function readingTotalWithoutAins(record) {
    return (
      nonnegativeInteger(record && record.bahasa_melayu) +
      nonnegativeInteger(record && record.bahasa_inggeris) +
      nonnegativeInteger(record && record.lain_lain_bahasa)
    );
  }

  function isIsoDate(value) {
    return /^\d{4}-\d{2}-\d{2}$/.test(cleanText(value));
  }

  function aggregateRows(records, students, startDate, endDate, includeAins = false, annualRecords = records) {
    if (!isIsoDate(startDate) || !isIsoDate(endDate) || startDate > endDate) {
      throw new Error("Julat tarikh eksport tidak sah.");
    }

    const studentsById = new Map();
    (Array.isArray(students) ? students : []).forEach((student) => {
      const id = cleanText(student && student.no_kad_pengenalan);
      if (id) {
        studentsById.set(id, student);
      }
    });

    const totals = new Map();
    function getRecordKey(record) {
      const id = cleanText(record && record.no_kad_pengenalan);
      const name = cleanText(record && record.nama);
      const kelas = cleanText(record && record.kelas);
      if (!id && !name) {
        return "";
      }
      return id ? `id:${id}` : `nama:${normalizedName(name)}|kelas:${kelas.toLocaleLowerCase("ms")}`;
    }

    (Array.isArray(records) ? records : []).forEach((record, index) => {
      const date = cleanText(record && record.tarikh);
      if (!isIsoDate(date) || date < startDate || date > endDate) {
        return;
      }

      const readingTotal = readingTotalWithoutAins(record);
      if (readingTotal <= 0) {
        return;
      }

      const id = cleanText(record && record.no_kad_pengenalan);
      const name = cleanText(record && record.nama);
      const kelas = cleanText(record && record.kelas);
      const key = getRecordKey(record);
      if (!key) {
        return;
      }

      const timestamp = Date.parse(cleanText(record && record.updated_at_client)) || 0;
      const rank = `${date}|${String(timestamp).padStart(16, "0")}|${String(index).padStart(8, "0")}`;
      const current = totals.get(key) || {
        no_kad_pengenalan: id,
        nama: name,
        kelas,
        bil_bahan_bacaan: 0,
        latestRank: "",
      };
      current.bil_bahan_bacaan += readingTotal;
      if (!current.latestRank || rank >= current.latestRank) {
        current.nama = name || current.nama;
        current.kelas = kelas || current.kelas;
        current.latestRank = rank;
      }
      totals.set(key, current);
    });

    if (includeAins) {
      const highestAinsByStudentYear = new Map();
      (Array.isArray(annualRecords) ? annualRecords : []).forEach((record, index) => {
        const date = cleanText(record && record.tarikh);
        const year = date.slice(0, 4);
        if (!isIsoDate(date) || year < startDate.slice(0, 4) || year > endDate.slice(0, 4)) {
          return;
        }
        const key = getRecordKey(record);
        if (!key) {
          return;
        }
        const ains = nonnegativeInteger(record && record.ains);
        const studentYearKey = `${key}|year:${year}`;
        const timestamp = Date.parse(cleanText(record && record.updated_at_client)) || 0;
        const rank = `${date}|${String(timestamp).padStart(16, "0")}|${String(index).padStart(8, "0")}`;
        const current = highestAinsByStudentYear.get(studentYearKey);
        if (!current || ains > current.ains || (ains === current.ains && rank > current.rank)) {
          highestAinsByStudentYear.set(studentYearKey, {
            key,
            ains,
            rank,
            no_kad_pengenalan: cleanText(record && record.no_kad_pengenalan),
            nama: cleanText(record && record.nama),
            kelas: cleanText(record && record.kelas),
          });
        }
      });

      highestAinsByStudentYear.forEach((entry) => {
        if (entry.ains <= 0) {
          return;
        }
        const current = totals.get(entry.key) || {
          no_kad_pengenalan: entry.no_kad_pengenalan,
          nama: entry.nama,
          kelas: entry.kelas,
          bil_bahan_bacaan: 0,
          latestRank: "",
        };
        current.bil_bahan_bacaan += entry.ains;
        if (!current.latestRank || entry.rank >= current.latestRank) {
          current.nama = entry.nama || current.nama;
          current.kelas = entry.kelas || current.kelas;
          current.latestRank = entry.rank;
        }
        totals.set(entry.key, current);
      });
    }

    const collator = new Intl.Collator("ms", { numeric: true, sensitivity: "base" });
    return [...totals.values()]
      .map((row) => {
        const student = studentsById.get(row.no_kad_pengenalan) || {};
        return {
          nama: cleanText(student.nama || student.nama_murid || row.nama),
          id_delima: cleanText(student.email_google_classroom),
          kelas: row.kelas,
          bil_bahan_bacaan: row.bil_bahan_bacaan,
        };
      })
      .filter((row) => row.nama && row.bil_bahan_bacaan > 0)
      .sort((a, b) => collator.compare(a.kelas, b.kelas) || collator.compare(a.nama, b.nama))
      .map((row, index) => ({ bil: index + 1, ...row }));
  }

  return {
    aggregateRows,
    isIsoDate,
    readingTotalWithoutAins,
  };
});
