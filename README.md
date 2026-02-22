# Perekodan NILAM Bulanan (GitHub Pages)

Borang ini:
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

5. Push semua fail ke repo dan hidupkan GitHub Pages.

## Fallback
Jika `config.js` kosong atau request Supabase gagal, rekod disimpan ke `localStorage` browser dan boleh dimuat turun melalui butang `Muat Turun CSV`.
