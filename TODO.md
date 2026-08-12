# TODO - Fitur My Folder (Penyimpanan File Import)

## Langkah
- [x] Analisis proyek & konfirmasi rencana
- [x] Buat model `ImportedFile` di `siswa/models.py`
- [x] Buat migration `0006_importedfile.py`
- [x] Modifikasi view `import_excel` untuk menyimpan file & record
- [x] Tambah view `penyimpanan_file` di `siswa/views.py`
- [x] Buat template `templates/penyimpanan_file.html`
- [x] Tambah route URL di `siswa/urls.py`
- [x] Tambah menu sidebar "My Folder" (setelah "Import Excel") di `dashboard.html` & `import_excel.html`
- [x] Register model di `siswa/admin.py`
- [x] Jalankan migration & uji
