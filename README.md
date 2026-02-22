# Perekodan NILAM Bulanan (GitHub Pages)

Borang ini:
- baca senarai murid daripada `senarai_nama_murid_10FEB2026 - ALL.csv`
- pilih `Bulan` (auto pilih bulan semasa) dan `Kelas`
- papar senarai murid ikut kelas
- kira automatik `Jumlah Aktiviti = Bahasa Melayu + Bahasa Inggeris + Lain-lain Bahasa`
- simpan data ke Supabase (sesuai untuk GitHub Pages)

## Fail utama
- `index.html`
- `styles.css`
- `app.js`
- `config.js`
- `supabase.sql`

## Kenapa bukan SQLite
GitHub Pages ialah static hosting (tiada server runtime), jadi SQLite tidak praktikal untuk simpan data berpusat. Pilihan sesuai: Supabase / Firebase / Google Sheets API. Projek ini guna Supabase.

## Setup ringkas
1. Cipta projek Supabase.
2. Jalankan SQL dalam `supabase.sql` (SQL Editor Supabase).
3. Ambil `Project URL` dan `anon public key`.
4. Isi `config.js`:

```js
window.NILAM_CONFIG = {
  supabaseUrl: "https://YOUR_PROJECT.supabase.co",
  supabaseAnonKey: "YOUR_ANON_KEY",
};
```

5. Push semua fail ke repo dan hidupkan GitHub Pages.

## Fallback
Jika `config.js` kosong atau request Supabase gagal, rekod disimpan ke `localStorage` browser dan boleh dimuat turun melalui butang `Muat Turun CSV`.
