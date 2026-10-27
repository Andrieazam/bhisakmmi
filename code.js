// V86: FAST START + FIX Nama Arab Santri dari kolom E Sheet Santri.
// V85: FAST START /exec tanpa setupDatabase saat membuka halaman.
// V84: Login cepat + Tahun Pelajaran unik.
// V83 FINAL: portal + tahun pelajaran; pasangan dengan Index V83.
// V82: Tahun Pelajaran final paired deployment.
// V80: direct Tahun Pelajaran reader.
// V79: Menu Tahun Pelajaran + CRUD.
// V77: Nama Arab Santri disimpan di kolom E Sheet Santri.
// V72 FIX SEBENARNYA: satu sumber Nama Arab = Sheet Mapel kolom C.
/**
 * SISTEM DAFTAR NILAI INTENSIF (KMMI)
 * Backend Google Apps Script Real-time Database
 */
/**
 * Reader Sheet aman: nama sheet dibuat case-insensitive.
 * Contoh "Guru", "GURU", atau "guru" tetap dibaca sebagai sheet yang sama.
 * Tidak membuat sheet baru dan tidak mengubah data yang sudah ada.
 */
function getSheetByNameSafe_(ss, wantedName) {
  if (!ss || !wantedName) return null;
  const exact = ss.getSheetByName(wantedName);
  if (exact) return exact;
  const wanted = String(wantedName).trim().toLowerCase();
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (String(sheets[i].getName()).trim().toLowerCase() === wanted) {
      return sheets[i];
    }
  }
  return null;
}
const SHEET_NAMES = {
  PENGATURAN: 'Pengaturan',
  USERS: 'Users',
  GURU: 'Guru',
  SANTRI: 'Santri',
  MAPEL: 'Mapel',
  NILAI: 'Nilai',
  TAHUN_PELAJARAN: 'Tahun Pelajaran',
  TUGAS_GURU: 'Tugas Guru'
};
function getNamaLembagaUntukJudul_() {
  const fallback = 'Sistem Daftar Nilai Intensif (KMMI)';
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get('KMMI_NAMA_LEMBAGA_JUDUL');
    if (cached && String(cached).trim()) return String(cached).trim();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getSheetByNameSafe_(ss, SHEET_NAMES.PENGATURAN);
    if (!sheet) return fallback;
    const values = sheet.getDataRange().getDisplayValues();
    for (let i = 1; i < values.length; i++) {
      const key = String(values[i][0] || '').trim().toLowerCase();
      const value = String(values[i][1] || '').trim();
      if (key === 'nama_sekolah' && value) {
        cache.put('KMMI_NAMA_LEMBAGA_JUDUL', value, 21600);
        return value;
      }
    }
  } catch (err) {
    console.error('getNamaLembagaUntukJudul_', err);
  }
  return fallback;
}
function doGet(e) {
  // Penting: gunakan HtmlTemplate agar \<title> sudah berisi
  // nama lembaga SEBELUM browser membuka halaman.
  // Ini berbeda dengan mengganti document.title setelah JavaScript berjalan.
  const namaLembaga = getNamaLembagaUntukJudul_();
  const template = HtmlService.createTemplateFromFile('Index');
  template.namaLembagaJudul = namaLembaga;
  return template.evaluate()
    .setTitle(namaLembaga)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
/**
 * Setup Otomatis Sheet Database & Default Data jika belum ada
 */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // 1. Sheet Pengaturan
  let shPengaturan = getSheetByNameSafe_(ss, SHEET_NAMES.PENGATURAN);
  if (!shPengaturan) {
    shPengaturan = ss.insertSheet(SHEET_NAMES.PENGATURAN);
    shPengaturan.appendRow(['Kunci', 'Nilai']);
    const defaultSettings = [
      ['nama_sekolah', 'KULLIYATU-L-MU\'ALLIMIN AL-ISLAMIYAH (KMMI)'],
      ['deskripsi', 'Sistem Pengelolaan Daftar Nilai Ujian & Evaluasi Intensif Pondok Pesantren'],
      ['logo_url', 'https\://cdn-icons-png.flaticon.com/512/2991/2991148.png'],
      ['pimpinan', 'K.H. Ahmad Dahlan, M.Pd.'],
      ['informasi_admin', 'Selamat Datang di Sistem Daftar Nilai Intensif KMMI. Harap pastikan seluruh nilai santri telah diinput sebelum tenggat waktu evaluasi semester.']
    ];
    defaultSettings.forEach(r => shPengaturan.appendRow(r));
  }
  // 2. Sheet Users (Default Admin & Guru)
  let shUsers = getSheetByNameSafe_(ss, SHEET_NAMES.USERS);
  if (!shUsers) {
    shUsers = ss.insertSheet(SHEET_NAMES.USERS);
    shUsers.appendRow(['Username', 'Password', 'Nama', 'Role']);
    shUsers.appendRow(['admin', 'admin123', 'Administrator Utama', 'Admin']);
    shUsers.appendRow(['guru', 'guru123', 'Ust. Abdullah, S.Pd.', 'Guru']);
  }
  // 3. Sheet Guru
  let shGuru = getSheetByNameSafe_(ss, SHEET_NAMES.GURU);
  if (!shGuru) {
    shGuru = ss.insertSheet(SHEET_NAMES.GURU);
    shGuru.appendRow(['ID', 'Nama Guru', 'Tingkatan', 'Status']);
    shGuru.appendRow(['GUR-101', 'Ust. Ahmad Fauzi, S.Ag.', 'KMMI 1 - 3', 'Aktif']);
    shGuru.appendRow(['GUR-102', 'Ustdh. Siti Aminah, M.Pd.', 'KMMI 4 - 6', 'Aktif']);
    shGuru.appendRow(['GUR-103', 'Ust. Muhammad Zulkarnaen', 'KMMI 1 - 6', 'Aktif']);
  }
  // 4. Sheet Santri
  let shSantri = getSheetByNameSafe_(ss, SHEET_NAMES.SANTRI);
  if (!shSantri) {
    shSantri = ss.insertSheet(SHEET_NAMES.SANTRI);
    shSantri.appendRow(['ID', 'NIS', 'Nama Santri', 'Kelas', 'Nama Arab']);
    shSantri.appendRow(['SAN-101', '2026001', 'Muhammad Rizky', '1 KMMI A', '']);
    shSantri.appendRow(['SAN-102', '2026002', 'Ahmad Zaki', '1 KMMI A', '']);
    shSantri.appendRow(['SAN-103', '2026003', 'Fatimah Zahra', '2 KMMI B', '']);
    shSantri.appendRow(['SAN-104', '2026004', 'Nurul Hidayah', '2 KMMI B', '']);
  } else {
    const santriLastCol = Math.max(shSantri.getLastColumn(), 4);
    const santriHeader = shSantri.getRange(1, 1, 1, santriLastCol).getValues()[0];
    if (santriHeader.indexOf('Nama Arab') === -1) {
      // Tambahkan kolom E tanpa menggeser ID/NIS/Nama/Kelas yang sudah dipakai sistem.
      shSantri.getRange(1, 5).setValue('Nama Arab');
    }
  }
  // 5. Sheet Mapel
  let shMapel = getSheetByNameSafe_(ss, SHEET_NAMES.MAPEL);
  if (!shMapel) {
    shMapel = ss.insertSheet(SHEET_NAMES.MAPEL);
  }
  normalizeMapelSheetFinal_();
  // 6. Sheet Nilai
  let shNilai = getSheetByNameSafe_(ss, SHEET_NAMES.NILAI);
  if (!shNilai) {
    shNilai = ss.insertSheet(SHEET_NAMES.NILAI);
    shNilai.appendRow(['ID', 'ID Santri', 'NIS', 'Nama Santri', 'Kelas', 'ID Mapel', 'Mata Pelajaran', 'Nilai Akhir', 'Tanggal Update', 'ID Guru', 'Nama Guru']);
  } else {
    const headerNilai = shNilai.getRange(1, 1, 1, Math.max(shNilai.getLastColumn(), 1)).getValues()[0];
    if (headerNilai.indexOf('ID Guru') === -1) shNilai.getRange(1, shNilai.getLastColumn() + 1).setValue('ID Guru');
    const headerNilai2 = shNilai.getRange(1, 1, 1, Math.max(shNilai.getLastColumn(), 1)).getValues()[0];
    if (headerNilai2.indexOf('Nama Guru') === -1) shNilai.getRange(1, shNilai.getLastColumn() + 1).setValue('Nama Guru');
  }
  // Sheet Tahun Pelajaran
  let shTahunPelajaran = getSheetByNameSafe_(ss, SHEET_NAMES.TAHUN_PELAJARAN);
  if (!shTahunPelajaran) {
    shTahunPelajaran = ss.insertSheet(SHEET_NAMES.TAHUN_PELAJARAN);
    shTahunPelajaran.appendRow(['ID', 'Tahun', 'Nama Semester', 'Tahun Mulai', 'Tahun Akhir', 'Tgl Raport', 'Status']);
  }
  // Pastikan REKAP ABSENSI LEBAR siap dan hapus sheet ABSENSI lama.
  _kmmiEnsureAttendanceSheet_();

  // Sheet Tugas Guru: ditambahkan tanpa mengubah sheet atau data lama.
  let shTugasGuru = getSheetByNameSafe_(ss, SHEET_NAMES.TUGAS_GURU);
  if (!shTugasGuru) {
    shTugasGuru = ss.insertSheet(SHEET_NAMES.TUGAS_GURU);
    shTugasGuru.appendRow(['ID', 'ID Guru', 'Nama Guru', 'ID Mapel', 'Mata Pelajaran', 'Kelas', 'Tingkatan', 'Hari', 'Jam Mulai', 'Jam Selesai', 'Tahun Pelajaran', 'Status']);
    shTugasGuru.setFrozenRows(1);
    shTugasGuru.getRange(1, 1, 1, 12).setBackground('#0f766e').setFontColor('#ffffff').setFontWeight('bold');
  }
  // Pastikan Mapel, Nama Arab, dan urutan kolom selalu sinkron.
  normalizeMapelSheetFinal_();
  syncNamaArabMapelV72_();
  // Pastikan satu-satunya penyimpanan absensi tetap REKAP ABSENSI LEBAR.
  _kmmiEnsureAttendanceSheet_();
  formatAllSheets();
}
/**
 * FORMAT SEMUA SHEET
 * - Header: center, tebal, warna hijau/teal
 * - Semua kolom data: center
 * - Kolom nama/deskripsi: rata kiri
 * - Border dan lebar kolom dirapikan
 *
 * ID tetap dipertahankan sebagai ID INTERNAL karena fungsi edit/hapus
 * menggunakan nilai pada kolom pertama.
 */
function normalizeMapelSheetFinal_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheetByNameSafe_(ss, SHEET_NAMES.MAPEL);
  if (!sheet) return;
  const target = ['No', 'Mata Pelajaran', 'Nama Arab', 'Tingkatan', 'ID'];
  const lastRow = sheet.getLastRow();
  const maxCol = Math.max(sheet.getLastColumn(), 1);
  const norm = v => String(v == null ? '' : v).trim().toLowerCase();
  const header = sheet.getRange(1, 1, 1, maxCol).getValues()[0].map(norm);
  const findCol = names => {
    for (const name of names) {
      const idx = header.indexOf(norm(name));
      if (idx >= 0) return idx;
    }
    return -1;
  };
  const colNo = findCol(['no']);
  const colMapel = findCol(['mata pelajaran', 'mapel', 'nama mata pelajaran']);
  const colArab = findCol(['nama arab', 'arab']);
  const colTingkat = findCol(['tingkatan', 'tingkat']);
  const colId = findCol(['id']);
  const sourceRows = lastRow >= 2
    ? sheet.getRange(2, 1, lastRow - 1, maxCol).getValues()
    : [];
  // Ambil ID Mapel yang sudah tersimpan pada Sheet Nilai.
  // Ini penting agar Mufrodat dan mapel lama tidak kehilangan hubungan
  // dengan nilai yang sudah ada.
  const nilaiIdByName = {};
  const nilaiSheet = getSheetByNameSafe_(ss, SHEET_NAMES.NILAI);
  if (nilaiSheet && nilaiSheet.getLastRow() >= 2) {
    const nv = nilaiSheet.getDataRange().getDisplayValues();
    for (let i = 1; i < nv.length; i++) {
      const nama = norm(nv[i][6]);
      const id = String(nv[i][5] || '').trim();
      if (nama && id && !nilaiIdByName[nama]) nilaiIdByName[nama] = id;
    }
  }
  const defaults = {
    'imla': 'الإملاء',
    'al mutholaah': 'المطالعة',
    'shorof': 'الصرف',
    'bahasa arab': 'اللغة العربية',
    'muhadatsah': 'المحادثة',
    'mufrodat': 'المفردات',
    'reading, vocabulary & conversation': 'القراءة والمفردات والمحادثة',
    'al quran & tajwid': 'القرآن والتجويد',
    'shorof-mukhtasor': 'الصرف المختصر',
    'shorof-amtsilah': 'الأمثلة الصرفية',
    'conversation': 'المحادثة',
    'reading & vocabulary': 'القراءة والمفردات',
    'nahwu': 'النحو'
  };
  // ID baru harus deterministik, bukan UUID baru setiap kali fungsi dipanggil.
  const makeStableId = (name, level) => {
    const raw = norm(name) + '|' + norm(level || '-');
    const digest = Utilities.computeDigest(
      Utilities.DigestAlgorithm.MD5,
      raw,
      Utilities.Charset.UTF_8
    );
    const hex = digest.map(b => {
      const n = b < 0 ? b + 256 : b;
      return ('0' + n.toString(16)).slice(-2);
    }).join('').substring(0, 10).toUpperCase();
    return 'MAP-' + hex;
  };
  const valueAt = (row, idx, fallback) => {
    if (idx >= 0 && idx < row.length) return row[idx];
    if (fallback !== undefined && fallback >= 0 && fallback < row.length) return row[fallback];
    return '';
  };
  const rows = [];
  let nomor = 1;
  sourceRows.forEach(row => {
    // Mendukung dua struktur lama:
    // A=ID,B=Mapel,C=Tingkatan,D=Nama Arab
    // atau A=No,B=Mapel,C=Nama Arab,D=Tingkatan.
    const nama = String(valueAt(row, colMapel, 1) || '').trim();
    if (!nama) return;
    let arab = String(valueAt(row, colArab, -1) || '').trim();
    let tingkatan = String(valueAt(row, colTingkat, -1) || '').trim();
    if (!tingkatan) {
      // Struktur lama A=No,B=Mapel,C=Arab,D=Tingkatan.
      tingkatan = String(row[3] || '').trim() || '-';
    }
    if (!arab && colArab < 0 && row.length >= 4) {
      const candidate = String(row[3] || '').trim();
      if (candidate && candidate !== tingkatan) arab = candidate;
    }
    const key = norm(nama);
    if (!arab && defaults[key]) arab = defaults[key];
    // PERTAMA: pertahankan ID yang sudah ada.
    let id = String(valueAt(row, colId, '') || '').trim();
    // KEDUA: jika ID belum ada, pakai ID yang sudah dipakai oleh Sheet Nilai.
    if (!id && nilaiIdByName[key]) id = nilaiIdByName[key];
    // KETIGA: baru buat ID deterministik.
    if (!id) id = makeStableId(nama, tingkatan);
    rows.push([nomor++, nama, arab, tingkatan || '-', id]);
  });
  // Jangan clearContents() lagi. Tulis struktur final tanpa menghapus
  // sheet lain atau data nilai.
  sheet.getRange(1, 1, 1, 5).setValues([target]);
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, 5).setValues(rows);
  }
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 70);
  sheet.setColumnWidth(2, 300);
  sheet.setColumnWidth(3, 260);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 120);
  if (sheet.getMaxColumns() >= 5) sheet.hideColumns(5);
  const endRow = sheet.getLastRow();
  sheet.getRange(1, 1, 1, 5)
    .setBackground('#0f766e')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  if (endRow >= 2) {
    sheet.getRange(2, 1, endRow - 1, 5)
      .setFontFamily('Arial Narrow')
      .setFontSize(12)
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('center');
    sheet.getRange(2, 2, endRow - 1, 1).setHorizontalAlignment('left');
    sheet.getRange(2, 3, endRow - 1, 1)
      .setFontFamily('Arial')
      .setHorizontalAlignment('center');
  }
  sheet.setRowHeight(1, 34);
}
function syncNamaArabMapelV72_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheetByNameSafe_(ss, SHEET_NAMES.MAPEL);
  if (!sheet) return 0;
  // Pastikan struktur final tanpa menghapus data.
  normalizeMapelSheetFinal_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  const values = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  const defaults = {
    'imla': 'الإملاء',
    'al mutholaah': 'المطالعة',
    'shorof': 'الصرف',
    'bahasa arab': 'اللغة العربية',
    'muhadatsah': 'المحادثة',
    'mufrodat': 'المفردات',
    'reading, vocabulary & conversation': 'القراءة والمفردات والمحادثة',
    'al quran & tajwid': 'القرآن والتجويد',
    'shorof-mukhtasor': 'الصرف المختصر',
    'shorof-amtsilah': 'الأمثلة الصرفية',
    'conversation': 'المحادثة',
    'reading & vocabulary': 'القراءة والمفردات',
    'nahwu': 'النحو'
  };
  let updated = 0;
  for (let i = 0; i < values.length; i++) {
    const nama = String(values[i][1] || '').trim();
    const arab = String(values[i][2] || '').trim();
    const key = nama.toLowerCase().replace(/\s+/g, ' ');
    if (nama && !arab && defaults[key]) {
      sheet.getRange(i + 2, 3).setValue(defaults[key]);
      updated++;
    }
  }
  if (updated > 0) touchDataVersion_();
  return updated;
}
function syncNamaArabMapelNow() {
  const count = syncNamaArabMapelV72_();
  return 'Selesai. ' + count + ' Nama Arab diisi/diperbarui langsung pada Sheet Mapel.';
}
function repairMapelSheetNow() {
  normalizeMapelSheetFinal_();
  formatAllSheets();
  return 'Selesai. Sheet Mapel sekarang: No | Mata Pelajaran | Nama Arab | Tingkatan | ID internal (tersembunyi).';
}
function formatAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configs = [
    {
      name: SHEET_NAMES.PENGATURAN,
      widths: [180, 520],
      leftCols: [2]
    },
    {
      name: SHEET_NAMES.USERS,
      widths: [150, 150, 260, 120],
      leftCols: [3]
    },
    {
      name: SHEET_NAMES.GURU,
      widths: [110, 280, 150, 110],
      leftCols: [2]
    },
    {
      name: SHEET_NAMES.SANTRI,
      widths: [110, 140, 300, 120, 260],
      leftCols: [3, 5]
    },
    {
      name: SHEET_NAMES.MAPEL,
      widths: [70, 300, 260, 150, 120],
      leftCols: [2],
      hideCols: [5]
    },
    {
      name: SHEET_NAMES.NILAI,
      widths: [110, 110, 140, 280, 120, 110, 260, 120, 190, 110, 220],
      leftCols: [4, 7, 11]
    },
    {
      name: SHEET_NAMES.TAHUN_PELAJARAN,
      widths: [120, 130, 180, 130, 130, 140, 120],
      leftCols: [3]
    }
  ];
  configs.forEach(cfg => {
    const sheet = ss.getSheetByName(cfg.name);
    if (!sheet) return;
    const lastCol = cfg.widths.length;
    const maxRows = Math.max(sheet.getMaxRows(), 2);
    // Header
    const header = sheet.getRange(1, 1, 1, lastCol);
    header
      .setFontWeight('bold')
      .setFontColor('#ffffff')
      .setBackground('#0f766e')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setWrap(true);
    // Semua data center terlebih dahulu.
    if (maxRows > 1) {
      sheet.getRange(2, 1, maxRows - 1, lastCol)
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');
      // Nama/deskripsi dibuat rata kiri.
      cfg.leftCols.forEach(col => {
        sheet.getRange(2, col, maxRows - 1, 1)
          .setHorizontalAlignment('left');
      });
    }
    // Border rapi untuk area yang terpakai.
    const lastRow = Math.max(sheet.getLastRow(), 1);
    if (lastRow >= 1) {
      sheet.getRange(1, 1, lastRow, lastCol)
        .setBorder(
          true, true, true, true, true, true,
          '#d9e2e7',
          SpreadsheetApp.BorderStyle.SOLID
        );
    }
    // Tinggi header.
    sheet.setRowHeight(1, 34);
    // Lebar kolom.
    cfg.widths.forEach((width, index) => {
      sheet.setColumnWidth(index + 1, width);
    });
    // Kolom internal (misalnya ID Mapel) dapat disembunyikan dari tampilan,
    // tetapi tetap terbaca oleh aplikasi.
    if (cfg.hideCols && Array.isArray(cfg.hideCols)) {
      cfg.hideCols.forEach(col => {
        if (col <= sheet.getMaxColumns()) sheet.hideColumns(col);
      });
    }
    // Header tetap terlihat ketika scroll.
    sheet.setFrozenRows(1);
    // Filter hanya jika belum ada.
    if (lastRow > 1) {
      const existingFilter = sheet.getFilter();
      if (existingFilter) {
        const range = existingFilter.getRange();
        if (range.getNumColumns() !== lastCol || range.getNumRows() !== lastRow) {
          existingFilter.remove();
          sheet.getRange(1, 1, lastRow, lastCol).createFilter();
        }
      } else {
        sheet.getRange(1, 1, lastRow, lastCol).createFilter();
      }
    }
  });
}
function getAllData() {
  // Baca data saja. Setup/format massal tidak boleh dijalankan tiap refresh.
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const getSheetData = (sheetName) => {
    const sheet = getSheetByNameSafe_(ss, sheetName);
    if (!sheet) return [];
    const values = sheet.getDataRange().getDisplayValues();
    if (values.length <= 1) return [];
    const headers = values[0];
    // Sheet Santri: baca berdasarkan nama header yang dinormalisasi.
    // Kolom Nama Arab WAJIB diambil dari header 'Nama Arab' atau fallback
    // kolom E. Ini membuat data Arab tetap terbaca walaupun ada perbedaan
    // huruf besar/kecil atau spasi pada header Spreadsheet.
    if (sheetName === SHEET_NAMES.SANTRI) {
      const h = headers.map(v => String(v == null ? '' : v).trim().toLowerCase());
      const find = names => {
        for (const n of names) {
          const i = h.indexOf(String(n).trim().toLowerCase());
          if (i >= 0) return i;
        }
        return -1;
      };
      const iID = find(['id']);
      const iNIS = find(['nis']);
      const iNama = find(['nama santri', 'nama']);
      const iKelas = find(['kelas', 'class']);
      const iNamaArab = find(['nama arab', 'nama_arab', 'nama arab santri', 'arab']);
      return values.slice(1).map((row, idx) => {
        const obj = {};
        obj.ID = iID >= 0 ? row[iID] : (row[0] || '');
        obj.NIS = iNIS >= 0 ? row[iNIS] : (row[1] || '');
        obj['Nama Santri'] = iNama >= 0 ? row[iNama] : (row[2] || '');
        obj.Kelas = iKelas >= 0 ? row[iKelas] : (row[3] || '');
        // Sheet saat ini menggunakan kolom E sebagai Nama Arab.
        obj['Nama Arab'] = iNamaArab >= 0 ? row[iNamaArab] : (row[4] || '');
        return obj;
      }).filter(o => String(o.NIS || '').trim() || String(o['Nama Santri'] || '').trim());
    }
    // Sheet Mapel: normalisasi hasil baca agar aplikasi SELALU menerima
    // field yang sama, walaupun struktur Spreadsheet lama masih berbeda.
    if (sheetName === SHEET_NAMES.MAPEL) {
      const h = headers.map(v => String(v || '').trim().toLowerCase());
      const find = names => {
        for (const n of names) {
          const i = h.indexOf(String(n).toLowerCase());
          if (i >= 0) return i;
        }
        return -1;
      };
      const iID = find(['id']);
      const iNo = find(['no']);
      const iMapel = find(['mata pelajaran', 'mapel', 'nama mata pelajaran']);
      const iArab = find(['nama arab', 'arab']);
      const iTingkat = find(['tingkatan', 'tingkat']);
      // Bila kolom ID belum ada di Sheet Mapel, cari ID lama
      // dari Sheet Nilai berdasarkan nama mapel. Ini menjaga nilai lama
      // tetap terhubung tanpa harus menunggu migrasi sheet.
      const nilaiIdByName = {};
      const nilaiSheet = getSheetByNameSafe_(ss, SHEET_NAMES.NILAI);
      if (nilaiSheet && nilaiSheet.getLastRow() >= 2) {
        const nv = nilaiSheet.getDataRange().getDisplayValues();
        for (let i = 1; i < nv.length; i++) {
          const nm = String(nv[i][6] || '').trim().toLowerCase().replace(/\s+/g, ' ');
          const id = String(nv[i][5] || '').trim();
          if (nm && id && !nilaiIdByName[nm]) nilaiIdByName[nm] = id;
        }
      }
      const stableId = (nama, tingkat) => {
        const raw = String(nama || '').trim().toLowerCase().replace(/\s+/g, ' ') +
          '|' + String(tingkat || '-').trim().toLowerCase();
        const digest = Utilities.computeDigest(
          Utilities.DigestAlgorithm.MD5, raw, Utilities.Charset.UTF_8
        );
        const hex = digest.map(b => {
          const n = b < 0 ? b + 256 : b;
          return ('0' + n.toString(16)).slice(-2);
        }).join('').substring(0, 10).toUpperCase();
        return 'MAP-' + hex;
      };
      return values.slice(1).map((row, idx) => {
        const obj = {};
        const nama = iMapel >= 0 ? row[iMapel] : (row[1] || '');
        const arab = iArab >= 0 ? row[iArab] : (row[3] || row[2] || '');
        const tingkat = iTingkat >= 0 ? row[iTingkat] : (row[3] || row[2] || '-');
        const nameKey = String(nama || '').trim().toLowerCase().replace(/\s+/g, ' ');
        let id = iID >= 0 ? String(row[iID] || '').trim() : '';
        if (!id) id = nilaiIdByName[nameKey] || stableId(nama, tingkat);
        obj.ID = id;
        obj.No = iNo >= 0 ? row[iNo] : (idx + 1);
        obj['Mata Pelajaran'] = nama;
        obj['Nama Arab'] = arab;
        obj.Tingkatan = tingkat || '-';
        return obj;
      }).filter(o => String(o['Mata Pelajaran'] || '').trim());
    }
    // Sheet Nilai lama memiliki HEADER yang tidak sesuai dengan posisi data.
    // Jangan memindahkan/menghapus data. Untuk Sheet Nilai, baca berdasarkan
    // posisi kolom yang dipakai oleh sistem agar data tampil benar di aplikasi.
    if (sheetName === SHEET_NAMES.NILAI) {
      const nilaiHeaders = [
        'ID',
        'ID Santri',
        'NIS',
        'Nama Santri',
        'Kelas',
        'ID Mapel',
        'Mata Pelajaran',
        'Nilai Akhir',
        'Tanggal Update',
        'ID Guru',
        'Nama Guru'
      ];
      return values.slice(1).map(row => {
        const obj = {};
        nilaiHeaders.forEach((h, i) => {
          obj[h] = row[i] !== undefined ? row[i] : '';
        });
        return obj;
      });
    }
    return values.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => {
        obj[h] = row[i] !== undefined ? row[i] : '';
      });
      return obj;
    });
  };
  const getSettingsData = () => {
    const sheet = getSheetByNameSafe_(ss, SHEET_NAMES.PENGATURAN);
    if (!sheet) return {};
    const values = sheet.getDataRange().getValues();
    let settings = {};
    for (let i = 1; i < values.length; i++) {
      if (values[i][0]) {
        settings[values[i][0]] = values[i][1];
      }
    }
    return settings;
  };
  return {
    pengaturan: getSettingsData(),
    guru: getSheetData(SHEET_NAMES.GURU),
    santri: getSheetData(SHEET_NAMES.SANTRI),
    mapel: getSheetData(SHEET_NAMES.MAPEL),
    nilai: getSheetData(SHEET_NAMES.NILAI),
    tahunPelajaran: getSheetData(SHEET_NAMES.TAHUN_PELAJARAN),
    tugasGuru: getSheetData(SHEET_NAMES.TUGAS_GURU)
  };
}
/**
 * V88: Reader langsung Guru.
 * Nama sheet dan nama header dibaca secara case-insensitive agar
 * "Guru"/"GURU" dan variasi header tetap tampil di aplikasi.
 */
/**
 * Reader langsung Sheet Nilai untuk Leger.
 * Membaca berdasarkan posisi kolom standar agar tidak bergantung pada cache
 * atau payload getAllData yang besar.
 */
function getNilaiDirectV1() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheetByNameSafe_(ss, SHEET_NAMES.NILAI);
  if (!sheet) return [];
  const values = sheet.getDataRange().getDisplayValues();
  if (values.length <= 1) return [];
  const headers = [
    'ID',
    'ID Santri',
    'NIS',
    'Nama Santri',
    'Kelas',
    'ID Mapel',
    'Mata Pelajaran',
    'Nilai Akhir',
    'Tanggal Update',
    'ID Guru',
    'Nama Guru'
  ];
  return values.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      obj[h] = row[i] !== undefined ? row[i] : '';
    });
    return obj;
  }).filter(r =>
    String(r['ID Santri'] || '').trim() ||
    String(r.NIS || '').trim() ||
    String(r['Nama Santri'] || '').trim()
  );
}
function getGuruDirectV88() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheetByNameSafe_(ss, SHEET_NAMES.GURU);
  if (!sheet) return [];
  const values = sheet.getDataRange().getDisplayValues();
  if (!values.length) return [];
  const headers = values[0].map(v => String(v == null ? '' : v).trim().toLowerCase());
  const find = names => {
    for (const name of names) {
      const i = headers.indexOf(String(name).trim().toLowerCase());
      if (i >= 0) return i;
    }
    return -1;
  };
  const iID = find(['id', 'id guru', 'kode', 'kode guru']);
  const iNama = find(['nama guru', 'nama', 'guru']);
  const iTingkat = find(['tingkatan', 'tingkat', 'kelas']);
  const iStatus = find(['status', 'keaktifan']);
  return values.slice(1).map((row, idx) => ({
    ID: iID >= 0 ? row[iID] : (row[0] || ''),
    'Nama Guru': iNama >= 0 ? row[iNama] : (row[1] || ''),
    Tingkatan: iTingkat >= 0 ? row[iTingkat] : (row[2] || '-'),
    Status: iStatus >= 0 ? row[iStatus] : (row[3] || 'Aktif')
  })).filter(r =>
    String(r.ID || '').trim() ||
    String(r['Nama Guru'] || '').trim()
  );
}
/**
 * V88: Reader langsung Mapel.
 * Dipakai frontend saat refresh agar data Sheet menjadi sumber final.
 */
function getMapelDirectV86() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheetByNameSafe_(ss, SHEET_NAMES.MAPEL);
  if (!sheet) return [];
  const values = sheet.getDataRange().getDisplayValues();
  if (values.length <= 1) return [];
  const headers = values[0].map(v => String(v == null ? '' : v).trim().toLowerCase());
  const find = names => {
    for (const name of names) {
      const i = headers.indexOf(String(name).trim().toLowerCase());
      if (i >= 0) return i;
    }
    return -1;
  };
  const iID = find(['id']);
  const iNo = find(['no']);
  const iMapel = find(['mata pelajaran', 'mapel', 'nama mata pelajaran']);
  const iArab = find(['nama arab', 'arab']);
  const iTingkat = find(['tingkatan', 'tingkat']);
  const nilaiIdByName = {};
  const nilaiSheet = getSheetByNameSafe_(ss, SHEET_NAMES.NILAI);
  if (nilaiSheet && nilaiSheet.getLastRow() >= 2) {
    const nv = nilaiSheet.getDataRange().getDisplayValues();
    for (let i = 1; i < nv.length; i++) {
      const nm = String(nv[i][6] || '').trim().toLowerCase().replace(/\s+/g, ' ');
      const id = String(nv[i][5] || '').trim();
      if (nm && id && !nilaiIdByName[nm]) nilaiIdByName[nm] = id;
    }
  }
  const stableId = (nama, tingkat) => {
    const raw = String(nama || '').trim().toLowerCase().replace(/\s+/g, ' ') +
      '|' + String(tingkat || '-').trim().toLowerCase();
    const digest = Utilities.computeDigest(
      Utilities.DigestAlgorithm.MD5, raw, Utilities.Charset.UTF_8
    );
    const hex = digest.map(b => {
      const n = b < 0 ? b + 256 : b;
      return ('0' + n.toString(16)).slice(-2);
    }).join('').substring(0, 10).toUpperCase();
    return 'MAP-' + hex;
  };
  return values.slice(1).map((row, idx) => {
    const nama = iMapel >= 0 ? row[iMapel] : (row[1] || '');
    const tingkat = iTingkat >= 0 ? row[iTingkat] : (row[3] || row[2] || '-');
    const nameKey = String(nama || '').trim().toLowerCase().replace(/\s+/g, ' ');
    let id = iID >= 0 ? String(row[iID] || '').trim() : '';
    if (!id) id = nilaiIdByName[nameKey] || stableId(nama, tingkat);
    return {
      ID: id,
      No: iNo >= 0 ? row[iNo] : String(idx + 1),
      'Mata Pelajaran': nama,
      'Nama Arab': iArab >= 0 ? row[iArab] : (row[2] || row[3] || ''),
      Tingkatan: tingkat || '-'
    };
  }).filter(r => String(r['Mata Pelajaran'] || '').trim());
}
/**
 * V88: Reader Pengaturan langsung untuk pasangan getMapelDirectV86.
 */
function getSettingsDirectV86() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheetByNameSafe_(ss, SHEET_NAMES.PENGATURAN);
  if (!sheet) return {};
  const values = sheet.getDataRange().getDisplayValues();
  const settings = {};
  for (let i = 1; i < values.length; i++) {
    const key = String(values[i][0] || '').trim();
    if (key) settings[key] = values[i][1] || '';
  }
  return settings;
}
/** V87: Reader langsung Sheet Santri.
 * Selalu membaca A:E dan memetakan kolom E sebagai Nama Arab.
 * Jalur ini sengaja dipisahkan dari getAllData agar data Nama Arab
 * tidak bergantung pada cache/normalisasi sheet lain.
 */
function getSantriDirect() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheetByNameSafe_(ss, SHEET_NAMES.SANTRI);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const lastCol = Math.max(5, sheet.getLastColumn());
  const values = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const rows = values.slice(1);
  return rows.map(row => ({
    ID: row[0] || '',
    NIS: row[1] || '',
    'Nama Santri': row[2] || '',
    Kelas: row[3] || '',
    'Nama Arab': row[4] || ''
  })).filter(r => String(r.NIS).trim() || String(r['Nama Santri']).trim());
}

/**
 * Resolusi ID Guru untuk akun login.
 * Prioritas: ID Guru yang tersimpan pada Users, lalu pencocokan nama persis
 * dengan master Guru, kemudian dengan Nama Guru pada Tugas Guru.
 * Fungsi ini hanya membaca data; tidak mengubah sheet atau data pengguna.
 */
function resolveIdGuruUntukAkun_(username, nama, idGuruTersimpan) {
  var saved = String(idGuruTersimpan || '').trim();
  if (saved) return saved;

  var norm = function(v) {
    return String(v == null ? '' : v)
      .trim().toLowerCase()
      .replace(/\s+/g, ' ');
  };
  var namaKey = norm(nama);
  var userKey = norm(username);
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Cocokkan dengan master Guru menggunakan nama yang sama persis.
  var shGuru = getSheetByNameSafe_(ss, SHEET_NAMES.GURU);
  if (shGuru && shGuru.getLastRow() >= 2) {
    var dataGuru = shGuru.getDataRange().getDisplayValues();
    var headersGuru = dataGuru[0].map(norm);
    var idIx = headersGuru.findIndex(function(h) {
      return ['id', 'id guru', 'id_guru', 'idguru', 'kode guru'].indexOf(h) >= 0;
    });
    var namaIx = headersGuru.findIndex(function(h) {
      return ['nama guru', 'nama', 'name'].indexOf(h) >= 0;
    });
    var statusIx = headersGuru.findIndex(function(h) {
      return ['status', 'keaktifan'].indexOf(h) >= 0;
    });
    if (idIx >= 0 && namaIx >= 0 && namaKey) {
      for (var i = 1; i < dataGuru.length; i++) {
        var rowNama = norm(dataGuru[i][namaIx]);
        var aktif = statusIx < 0 || !norm(dataGuru[i][statusIx]) ||
          norm(dataGuru[i][statusIx]) === 'aktif';
        if (aktif && rowNama === namaKey && String(dataGuru[i][idIx] || '').trim()) {
          return String(dataGuru[i][idIx]).trim();
        }
      }
    }
  }

  // Fallback: tugas guru, jika nama akun tidak ditemukan di master Guru.
  var shTugas = getSheetByNameSafe_(ss, SHEET_NAMES.TUGAS_GURU);
  if (shTugas && shTugas.getLastRow() >= 2) {
    var dataTugas = shTugas.getDataRange().getDisplayValues();
    var headersTugas = dataTugas[0].map(norm);
    var idTugasIx = headersTugas.findIndex(function(h) {
      return ['id guru', 'id_guru', 'idguru'].indexOf(h) >= 0;
    });
    var namaTugasIx = headersTugas.findIndex(function(h) {
      return ['nama guru', 'guru'].indexOf(h) >= 0;
    });
    var statusTugasIx = headersTugas.findIndex(function(h) {
      return h === 'status';
    });
    if (idTugasIx >= 0 && namaTugasIx >= 0 && namaKey) {
      for (var j = 1; j < dataTugas.length; j++) {
        var taskName = norm(dataTugas[j][namaTugasIx]);
        var taskActive = statusTugasIx < 0 || !norm(dataTugas[j][statusTugasIx]) ||
          norm(dataTugas[j][statusTugasIx]) === 'aktif';
        if (taskActive && taskName === namaKey && String(dataTugas[j][idTugasIx] || '').trim()) {
          return String(dataTugas[j][idTugasIx]).trim();
        }
      }
    }
  }

  // Username sengaja tidak dianggap sebagai ID Guru, kecuali memang tersimpan
  // eksplisit di kolom ID Guru. Ini mencegah akses tugas milik guru lain.
  return '';
}

function loginUser(username, password) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheetByNameSafe_(ss, SHEET_NAMES.USERS);
  if (!sheet) return { success: false, message: 'Sheet Users tidak ditemukan.' };
  const values = sheet.getDataRange().getDisplayValues();
  if (!values.length) return { success: false, message: 'Sheet Users belum memiliki data.' };
  const headers = values[0].map(v => String(v || '').trim().toLowerCase());
  const col = names => { for (let i = 0; i < names.length; i++) { const n = headers.indexOf(names[i]); if (n >= 0) return n; } return -1; };
  const cUser=col(['username']), cPass=col(['password']), cNama=col(['nama','name']), cRole=col(['role']), cGuru=col(['id guru','id_guru','idguru']), cStatus=col(['status']), cKelasWali=col(['kelas wali','kelas_wali','kelaswali']);
  for (let i=1;i<values.length;i++) {
    const row=values[i];
    if (String(row[cUser]||'').trim().toLowerCase()!==String(username||'').trim().toLowerCase()) continue;
    if (String(row[cPass]||'')!==String(password||'').trim()) continue;
    const status=String(cStatus>=0?row[cStatus]:'Aktif').trim();
    if(status && status.toLowerCase()!=='aktif') return {success:false,message:'Akun ini tidak aktif. Hubungi Administrator.'};
    var loginUsername = cUser >= 0 ? String(row[cUser] || '').trim() : String(username || '').trim();
    var loginNama = cNama >= 0 ? String(row[cNama] || '').trim() : '';
    var savedIdGuru = cGuru >= 0 ? String(row[cGuru] || '').trim() : '';
    var resolvedIdGuru = resolveIdGuruUntukAkun_(loginUsername, loginNama, savedIdGuru);
    return {
      success: true,
      user: {
        username: loginUsername,
        nama: loginNama,
        role: cRole >= 0 ? row[cRole] : 'Guru',
        idGuru: resolvedIdGuru,
        kelasWali: cKelasWali >= 0 ? String(row[cKelasWali] || '').trim() : ''
      }
    };
  }
  return { success: false, message: 'Username atau Password salah!' };
}
function getTahunPelajaranData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheetByNameSafe_(ss, SHEET_NAMES.TAHUN_PELAJARAN);
  if (!sheet) return [];
  const values = sheet.getDataRange().getDisplayValues();
  if (!values.length) return [];
  const headers = values[0].map(v => String(v || '').trim());
  const norm = s => String(s || '').trim().toLowerCase();
  const find = names => {
    for (const name of names) {
      const i = headers.map(norm).indexOf(norm(name));
      if (i >= 0) return i;
    }
    return -1;
  };
  const iID = find(['ID']);
  const iTahun = find(['Tahun']);
  const iSemester = find(['Nama Semester', 'Semester']);
  const iMulai = find(['Tahun Mulai']);
  const iAkhir = find(['Tahun Akhir']);
  const iTgl = find(['Tgl Raport', 'Tanggal Raport']);
  const iStatus = find(['Status']);
  const rows = values.slice(1).map((row, idx) => ({
    ID: iID >= 0 ? row[iID] : '',
    Tahun: iTahun >= 0 ? row[iTahun] : '',
    'Nama Semester': iSemester >= 0 ? row[iSemester] : '',
    'Tahun Mulai': iMulai >= 0 ? row[iMulai] : '',
    'Tahun Akhir': iAkhir >= 0 ? row[iAkhir] : '',
    'Tgl Raport': iTgl >= 0 ? row[iTgl] : '',
    Status: iStatus >= 0 ? row[iStatus] : ''
  })).filter(r => String(r.Tahun || '').trim() || String(r['Nama Semester'] || '').trim());
  // Bila Spreadsheet memiliki baris duplikat, aplikasi hanya mengirim satu
  // untuk kombinasi Tahun + Nama Semester. Data Sheet tidak dihapus.
  const unique = [];
  const seen = {};
  rows.forEach(row => {
    const key = String(row.Tahun || '').trim().toLowerCase() + '|' + String(row['Nama Semester'] || '').trim().toLowerCase();
    if (!seen[key]) {
      seen[key] = true;
      unique.push(row);
    }
  });
  return unique;
}
function saveTahunPelajaran(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheetByNameSafe_(ss, SHEET_NAMES.TAHUN_PELAJARAN);
  if (!sheet) return { success: false, message: 'Sheet Tahun Pelajaran tidak ditemukan.' };
  const values = sheet.getDataRange().getValues();
  const tahun = String(data.Tahun || '').trim();
  const semester = String(data['Nama Semester'] || '').trim();
  const tahunMulai = String(data['Tahun Mulai'] || '').trim();
  const tahunAkhir = String(data['Tahun Akhir'] || '').trim();
  const tglRaport = String(data['Tgl Raport'] || '').trim();
  const status = String(data.Status || 'Tidak Aktif').trim() || 'Tidak Aktif';
  if (!tahun) return { success: false, message: 'Tahun wajib diisi.' };
  if (!semester) return { success: false, message: 'Nama Semester wajib diisi.' };
  if (data.ID) {
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]) === String(data.ID)) {
        sheet.getRange(i + 1, 1, 1, 7).setValues([[
          data.ID, tahun, semester, tahunMulai, tahunAkhir, tglRaport, status
        ]]);
        touchDataVersion_();
        return { success: true, message: 'Tahun Pelajaran berhasil diperbarui!' };
      }
    }
  }
  const newID = 'TP-' + Utilities.getUuid().slice(0, 8).toUpperCase();
  sheet.appendRow([newID, tahun, semester, tahunMulai, tahunAkhir, tglRaport, status]);
  touchDataVersion_();
  return { success: true, message: 'Tahun Pelajaran berhasil ditambahkan!' };
}
function deleteTahunPelajaran(id) {
  return deleteRowByID(SHEET_NAMES.TAHUN_PELAJARAN, id);
}
function deleteAllTahunPelajaran() {
  return clearSheetData(
    SHEET_NAMES.TAHUN_PELAJARAN,
    ['ID', 'Tahun', 'Nama Semester', 'Tahun Mulai', 'Tahun Akhir', 'Tgl Raport', 'Status']
  );
}
/**
 * Versi data ringan untuk sinkronisasi antar perangkat.
 * Client cukup mengecek angka ini. getAllData hanya dipanggil
 * jika versinya berubah.
 */
function touchDataVersion_() {
  const props = PropertiesService.getScriptProperties();
  const now = Date.now();
  const prev = Number(props.getProperty('KMMI_DATA_VERSION') || 0);
  const next = String(Math.max(now, prev + 1));
  props.setProperty('KMMI_DATA_VERSION', next);
  return next;
}
function getDataVersion() {
  // Pemeriksaan versi dibuat O(1): tidak membaca seluruh isi sheet.
  // Perubahan dari fungsi aplikasi memanggil touchDataVersion_().
  // Edit langsung di Spreadsheet ditangani oleh trigger onEdit terpasang.
  const props = PropertiesService.getScriptProperties();
  let version = props.getProperty('KMMI_DATA_VERSION');
  if (!version) {
    version = String(Date.now());
    props.setProperty('KMMI_DATA_VERSION', version);
  }
  return { success: true, version: version };
}

/**
 * Pasang sekali dari editor Apps Script pada salinan spreadsheet.
 * Memastikan edit manual di sheet juga memperbarui versi sinkronisasi.
 * Tidak membaca/menulis isi data; hanya menaikkan penanda versi.
 */
function setupDataVersionTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Spreadsheet aktif tidak ditemukan.');
  const handler = 'onSpreadsheetEditTouchDataVersion_';
  const triggers = ScriptApp.getProjectTriggers();
  const exists = triggers.some(t =>
    t.getHandlerFunction() === handler && t.getEventType() === ScriptApp.EventType.ON_EDIT
  );
  if (!exists) {
    ScriptApp.newTrigger(handler).forSpreadsheet(ss).onEdit().create();
  }
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('KMMI_DATA_VERSION')) {
    props.setProperty('KMMI_DATA_VERSION', String(Date.now()));
  }
  // Cache fingerprint lama tidak lagi dipakai.
  props.deleteProperty('KMMI_DATA_FINGERPRINT');
  props.deleteProperty('KMMI_VERSION_CHECK_AT');
  return { success: true, message: exists
    ? 'Trigger perubahan data sudah terpasang.'
    : 'Trigger perubahan data berhasil dipasang.' };
}

/** Trigger ini hanya untuk edit langsung oleh pengguna di Google Sheets. */
function onSpreadsheetEditTouchDataVersion_(e) {
  try {
    if (!e || !e.range) return;
    touchDataVersion_();
  } catch (err) {
    console.error('Gagal memperbarui versi data setelah edit Spreadsheet:', err);
  }
}

function getPublicBrand() {
  // Sumber utama identitas aplikasi: Google Sheets -> Pengaturan.
  // Dipanggil saat halaman dibuka, termasuk sebelum login.
  const fallback = { nama_sekolah: 'KMMI', deskripsi_sekolah: '', logo_url: '' };
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAMES.PENGATURAN);
    if (!sh) return fallback;
    const values = sh.getDataRange().getDisplayValues();
    const out = Object.assign({}, fallback);
    for (let i = 1; i < values.length; i++) {
      const key = String(values[i][0] || '').trim().toLowerCase();
      const val = String(values[i][1] || '').trim();
      if (key === 'nama_sekolah') out.nama_sekolah = val;
      else if (key === 'deskripsi_sekolah') out.deskripsi_sekolah = val;
      else if (key === 'deskripsi' && !out.deskripsi_sekolah) out.deskripsi_sekolah = val;
      else if (key === 'deskripsi_yayasan' && !out.deskripsi_sekolah) out.deskripsi_sekolah = val;
      else if (key === 'logo_url') out.logo_url = val;
    }
    return out;
  } catch (err) {
    console.warn('getPublicBrand:', err);
    return fallback;
  }
}
function savePengaturan(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheetByNameSafe_(ss, SHEET_NAMES.PENGATURAN);
  sheet.clearContents();
  sheet.appendRow(['Kunci', 'Nilai']);
  for (let key in data) {
    sheet.appendRow([key, data[key]]);
  }
  // Nama lembaga dipakai sebagai judul / nama halaman Web App.
  // Hapus cache supaya perubahan nama langsung dipakai pada pembukaan berikutnya.
  try {
    CacheService.getScriptCache().remove('KMMI_NAMA_LEMBAGA_JUDUL');
  } catch (err) {
    console.warn('Gagal menghapus cache judul lembaga:', err);
  }
  touchDataVersion_();
  return { success: true, message: 'Data Pengaturan Sekolah Berhasil Disimpan & Disinkronkan!' };
}
/**
 * KOP RAPORT TERPUSAT
 * ------------------------------------------------------------
 * KOP tidak lagi bergantung pada localStorage HP/PC.
 * Satu KOP disimpan di Spreadsheet pada sheet khusus KOP_RAPORT.
 *
 * Data gambar dipecah menjadi beberapa baris agar tidak melewati
 * batas ukuran satu sel Google Sheets.
 */
function getKopRaportSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = getSheetByNameSafe_(ss, 'KOP_RAPORT');
  if (!sh) {
    sh = ss.insertSheet('KOP_RAPORT');
    sh.getRange(1, 1, 1, 2).setValues([['No', 'Data']]);
  }
  return sh;
}
function saveKopRaportCentral(dataUrl) {
  const url = String(dataUrl || '').trim();
  if (!url) return { success: false, message: 'Data KOP kosong.' };
  const sh = getKopRaportSheet_();
  // Hapus KOP lama saja; sheet lain tidak disentuh.
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 2).clearContent();
  }
  // Batas aman per sel, di bawah batas maksimum Google Sheets.
  const chunkSize = 45000;
  const rows = [];
  for (let i = 0, no = 1; i < url.length; i += chunkSize, no++) {
    rows.push([no, url.substring(i, i + chunkSize)]);
  }
  if (rows.length) {
    sh.getRange(2, 1, rows.length, 2).setValues(rows);
  }
  touchDataVersion_();
  return {
    success: true,
    message: 'KOP Raport tersimpan terpusat.',
    chunks: rows.length
  };
}
function getKopRaportCentral() {
  const sh = getSheetByNameSafe_(SpreadsheetApp.getActiveSpreadsheet(), 'KOP_RAPORT');
  if (!sh || sh.getLastRow() < 2) {
    return { success: true, kop_raport_url: '' };
  }
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getDisplayValues();
  values.sort((a, b) => Number(a[0] || 0) - Number(b[0] || 0));
  let url = '';
  for (let i = 0; i < values.length; i++) {
    url += String(values[i][1] || '');
  }
  return { success: true, kop_raport_url: url };
}
function saveGuru(guruObj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheetByNameSafe_(ss, SHEET_NAMES.GURU);
  const values = sheet.getDataRange().getValues();
  if (guruObj.ID) {
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]) === String(guruObj.ID)) {
        sheet.getRange(i + 1, 1, 1, 4).setValues([[guruObj.ID, guruObj['Nama Guru'], guruObj['Tingkatan'], guruObj['Status']]]);
        touchDataVersion_();
        return { success: true, message: 'Data Guru Berhasil Diperbarui!' };
      }
    }
  }
  const newID = 'GUR-' + Math.floor(100 + Math.random() * 900);
  sheet.appendRow([newID, guruObj['Nama Guru'], guruObj['Tingkatan'], guruObj['Status'] || 'Aktif']);
  touchDataVersion_();
  return { success: true, message: 'Data Guru Berhasil Ditambahkan!' };
}
function deleteGuru(id) {
  return deleteRowByID(SHEET_NAMES.GURU, id);
}
function deleteGuruBulk(ids) {
  return deleteRowsByIDs(SHEET_NAMES.GURU, ids);
}
function deleteAllGuru() {
  return clearSheetData(SHEET_NAMES.GURU, ['ID', 'Nama Guru', 'Tingkatan', 'Status']);
}
function importGuru(guruList) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheetByNameSafe_(ss, SHEET_NAMES.GURU);
  guruList.forEach(item => {
    const id = item.ID || ('GUR-' + Math.floor(100 + Math.random() * 900));
    sheet.appendRow([id, item['Nama Guru'] || item.Nama || item.nama_guru || 'Guru', item['Tingkatan'] || item.Tingkatan || '-', item['Status'] || item.Status || 'Aktif']);
  });
  touchDataVersion_();
  return { success: true, message: `Berhasil mengimpor ${guruList.length} data Guru!` };
}
function ensureSantriNamaArabColumn_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheetByNameSafe_(ss, SHEET_NAMES.SANTRI);
  if (!sheet) return null;
  if (sheet.getRange(1, 5).getValue() !== 'Nama Arab') {
    sheet.getRange(1, 5).setValue('Nama Arab');
  }
  return sheet;
}
function saveSantri(santriObj) {
  const sheet = ensureSantriNamaArabColumn_();
  if (!sheet) return { success: false, message: 'Sheet Santri tidak ditemukan.' };
  const values = sheet.getDataRange().getValues();
  const namaArab = String(santriObj['Nama Arab'] || '').trim();
  if (santriObj.ID) {
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]) === String(santriObj.ID)) {
        sheet.getRange(i + 1, 1, 1, 5).setValues([[
          santriObj.ID,
          santriObj.NIS,
          santriObj['Nama Santri'],
          santriObj.Kelas,
          namaArab
        ]]);
        touchDataVersion_();
        return { success: true, message: 'Data Santri + Nama Arab Berhasil Diperbarui!' };
      }
    }
  }
  const newID = 'SAN-' + Math.floor(100 + Math.random() * 900);
  sheet.appendRow([
    newID,
    santriObj.NIS,
    santriObj['Nama Santri'],
    santriObj.Kelas,
    namaArab
  ]);
  touchDataVersion_();
  return { success: true, message: 'Data Santri + Nama Arab Berhasil Ditambahkan!' };
}
function deleteSantri(id) {
  return deleteRowByID(SHEET_NAMES.SANTRI, id);
}
function deleteSantriBulk(ids) {
  return deleteRowsByIDs(SHEET_NAMES.SANTRI, ids);
}
function deleteAllSantri() {
  return clearSheetData(SHEET_NAMES.SANTRI, ['ID', 'NIS', 'Nama Santri', 'Kelas', 'Nama Arab']);
}
function importSantri(santriList) {
  const sheet = ensureSantriNamaArabColumn_();
  if (!sheet) return { success: false, message: 'Sheet Santri tidak ditemukan.' };
  if (!Array.isArray(santriList)) return { success: false, message: 'Data Excel tidak valid.' };
  let count = 0;
  santriList.forEach(item => {
    const id = item.ID || ('SAN-' + Math.floor(100 + Math.random() * 900));
    const nis = item.NIS || item.nis || '-';
    const nama = item['Nama Santri'] || item.Nama || item.nama_santri || 'Santri';
    const kelas = item.Kelas || item.kelas || '-';
    const namaArab = item['Nama Arab'] || item.Nama_Arab || item['Nama_Arab'] || item.Arab || item.arab || '';
    sheet.appendRow([id, nis, nama, kelas, namaArab]);
    count++;
  });
  touchDataVersion_();
  return {
    success: true,
    message: `Berhasil mengimpor ${count} data Santri + Nama Arab!`
  };
}
function saveMapel(mapelObj) {
  normalizeMapelSheetFinal_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheetByNameSafe_(ss, SHEET_NAMES.MAPEL);
  if (!sheet) return { success: false, message: 'Sheet Mapel tidak ditemukan.' };
  const values = sheet.getDataRange().getValues();
  const nama = String(mapelObj['Mata Pelajaran'] || '').trim();
  const arab = String(mapelObj['Nama Arab'] || '').trim();
  const tingkatan = String(mapelObj.Tingkatan || '-').trim() || '-';
  if (!nama) return { success: false, message: 'Nama Mata Pelajaran wajib diisi.' };
  // Struktur: A No, B Mata Pelajaran, C Nama Arab, D Tingkatan, E ID.
  if (mapelObj.ID) {
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][4]) === String(mapelObj.ID)) {
        sheet.getRange(i + 1, 2, 1, 4).setValues([[
          nama, arab, tingkatan, mapelObj.ID
        ]]);
        touchDataVersion_();
        return { success: true, message: 'Mata Pelajaran + Nama Arab Berhasil Diperbarui!' };
      }
    }
  }
  // Cegah data ganda berdasarkan Mata Pelajaran + Tingkatan.
  for (let i = 1; i < values.length; i++) {
    if (
      String(values[i][1]).trim().toLowerCase() === nama.toLowerCase() &&
      String(values[i][3] || '-').trim().toLowerCase() === tingkatan.toLowerCase()
    ) {
      // Bila hanya Nama Arab yang baru diisi, perbarui baris yang sudah ada.
      if (arab && !String(values[i][2] || '').trim()) {
        sheet.getRange(i + 1, 3).setValue(arab);
        touchDataVersion_();
        return { success: true, message: 'Nama Arab berhasil diperbarui pada Mata Pelajaran yang sudah ada!' };
      }
      return { success: false, message: 'Mata Pelajaran tersebut sudah ada, sehingga tidak dibuat data ganda.' };
    }
  }
  const newID = 'MAP-' + Utilities.getUuid().slice(0, 8).toUpperCase();
  const newNo = Math.max(0, ...values.slice(1).map(r => Number(r[0]) || 0)) + 1;
  sheet.appendRow([newNo, nama, arab, tingkatan, newID]);
  touchDataVersion_();
  return { success: true, message: 'Mata Pelajaran + Nama Arab Berhasil Ditambahkan!' };
}
function deleteMapel(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheetByNameSafe_(ss, SHEET_NAMES.MAPEL);
  if (!sheet) return { success: false, message: 'Sheet Mapel tidak ditemukan.' };
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][4]) === String(id)) {
      sheet.deleteRow(i + 1);
      renumberMapel_();
      touchDataVersion_();
      return { success: true, message: 'Mata Pelajaran Berhasil Dihapus!' };
    }
  }
  return { success: false, message: 'Mata Pelajaran tidak ditemukan!' };
}
function deleteMapelBulk(ids) {
  const wanted = (ids || []).map(String);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheetByNameSafe_(ss, SHEET_NAMES.MAPEL);
  if (!sheet) return { success: false, message: 'Sheet Mapel tidak ditemukan.' };
  const values = sheet.getDataRange().getValues();
  let count = 0;
  for (let i = values.length - 1; i >= 1; i--) {
    if (wanted.includes(String(values[i][4]))) {
      sheet.deleteRow(i + 1);
      count++;
    }
  }
  renumberMapel_();
  touchDataVersion_();
  return { success: true, message: `Berhasil menghapus ${count} Mata Pelajaran.` };
}
function deleteAllMapel() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = getSheetByNameSafe_(ss, SHEET_NAMES.MAPEL);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.MAPEL);
  }
  sheet.clearContents();
  sheet.clearFormats();
  sheet.getRange(1, 1, 1, 5).setValues([['No', 'Mata Pelajaran', 'Nama Arab', 'Tingkatan', 'ID']]);
  formatAllSheets();
  touchDataVersion_();
  return { success: true, message: 'Semua data Mata Pelajaran berhasil dihapus.' };
}
function importMapel(mapelList) {
  normalizeMapelSheetFinal_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheetByNameSafe_(ss, SHEET_NAMES.MAPEL);
  if (!sheet) return { success: false, message: 'Sheet Mapel tidak ditemukan.' };
  if (!Array.isArray(mapelList)) return { success: false, message: 'Data Excel tidak valid.' };
  const values = sheet.getDataRange().getValues();
  const existing = {};
  for (let i = 1; i < values.length; i++) {
    const nama = String(values[i][1] || '').trim();
    const tingkatan = String(values[i][3] || '-').trim() || '-';
    if (nama) {
      const key = (nama + '|' + tingkatan).toLowerCase().replace(/\s+/g, ' ');
      existing[key] = {
        row: i + 1,
        id: String(values[i][4] || ''),
        arab: String(values[i][2] || '').trim()
      };
    }
  }
  const seenImport = {};
  const rowsToAdd = [];
  let skipped = 0;
  let updated = 0;
  let nextNo = Math.max(0, ...values.slice(1).map(r => Number(r[0]) || 0)) + 1;
  mapelList.forEach(item => {
    const nama = String(
      item['Mata Pelajaran'] ||
      item.Mapel ||
      item.mata_pelajaran ||
      item['Nama Mata Pelajaran'] ||
      ''
    ).trim();
    const tingkatan = String(item.Tingkatan || item.tingkatan || '-').trim() || '-';
    const arab = String(
      item['Nama Arab'] ||
      item.Nama_Arab ||
      item['Nama_Arab'] ||
      item.Arab ||
      item.arab ||
      ''
    ).trim();
    if (!nama) return;
    const key = (nama + '|' + tingkatan).toLowerCase().replace(/\s+/g, ' ');
    if (seenImport[key]) {
      skipped++;
      return;
    }
    seenImport[key] = true;
    if (existing[key]) {
      if (arab && arab !== existing[key].arab) {
        sheet.getRange(existing[key].row, 3).setValue(arab);
        updated++;
      } else {
        skipped++;
      }
      return;
    }
    const id = String(item.ID || '').trim() || ('MAP-' + Utilities.getUuid().slice(0, 8).toUpperCase());
    rowsToAdd.push([nextNo++, nama, arab, tingkatan, id]);
  });   
  if (rowsToAdd.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAdd.length, 5).setValues(rowsToAdd);
  }
  const autoFilled = syncNamaArabMapelV72_();
  touchDataVersion_();
  return {
    success: true,
    message: `Sinkronisasi selesai. ${rowsToAdd.length} Mata Pelajaran baru, ${updated} Nama Arab diperbarui, ${autoFilled} Nama Arab otomatis diisi, ${skipped} data dilewati.`
  };
}
function renumberMapel_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheetByNameSafe_(ss, SHEET_NAMES.MAPEL);
  if (!sheet || sheet.getLastRow() < 2) return;
  const n = sheet.getLastRow() - 1;
  const nums = [];
  for (let i = 1; i <= n; i++) nums.push([i]);
  sheet.getRange(2, 1, n, 1).setValues(nums);
}
function saveNilaiBatch(nilaiArray) {
  if (!Array.isArray(nilaiArray)) {
    return { success: false, message: 'Data nilai tidak valid.' };
  }
  for (let i = 0; i < nilaiArray.length; i++) {
    const nilai = Number(nilaiArray[i].nilai_akhir);
    if (!isFinite(nilai) || nilai < 1 || nilai > 10) {
      return { success: false, message: 'Nilai harus berada pada rentang 1 sampai 10.' };
    }
  }

  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getSheetByNameSafe_(ss, SHEET_NAMES.NILAI);
    if (!sheet) return { success: false, message: 'Sheet Nilai tidak ditemukan.' };

    const lastRow = sheet.getLastRow();
    const values = lastRow >= 1
      ? sheet.getRange(1, 1, lastRow, Math.max(11, sheet.getLastColumn())).getValues()
      : [];
    const now = new Date().toLocaleString('id-ID');
    let updated = 0;
    let added = 0;

    // Satu nilai unik per Santri + Mata Pelajaran.
    // ID Guru bukan bagian dari kunci agar nilai lama tetap diperbarui
    // walaupun kolom ID Guru pada data lama kosong/berbeda.
    const sameStudent = (existing, item) => {
      const existingId = String(existing[1] || '').trim();
      const incomingId = String(item.id_santri || '').trim();
      if (existingId && incomingId && existingId === incomingId) return true;
      const existingNis = String(existing[2] || '').trim();
      const incomingNis = String(item.nis || '').trim();
      return !!existingNis && !!incomingNis && existingNis === incomingNis;
    };
    const sameSubject = (existing, item) => {
      const existingId = String(existing[5] || '').trim();
      const incomingId = String(item.id_mapel || '').trim();
      if (existingId && incomingId && existingId === incomingId) return true;
      const existingName = String(existing[6] || '').trim().toLowerCase().replace(/\s+/g, ' ');
      const incomingName = String(item.mapel || '').trim().toLowerCase().replace(/\s+/g, ' ');
      return !!existingName && !!incomingName && existingName === incomingName;
    };

    nilaiArray.forEach(item => {
      let rowIndex = -1;
      // Cari rekaman lama, termasuk data lama yang ID Gurunya kosong.
      for (let i = 1; i < values.length; i++) {
        if (sameStudent(values[i], item) && sameSubject(values[i], item)) {
          rowIndex = i;
          break;
        }
      }

      const record = [
        rowIndex >= 1 && values[rowIndex][0]
          ? values[rowIndex][0]
          : ('NIL-' + Utilities.getUuid().slice(0, 8).toUpperCase()),
        item.id_santri || '',
        item.nis || '',
        item.nama_santri || '',
        item.kelas || '',
        item.id_mapel || '',
        item.mapel || '',
        Number(item.nilai_akhir),
        now,
        item.id_guru || '',
        item.nama_guru || ''
      ];

      if (rowIndex >= 1) {
        // Perbarui baris yang ada, tidak menambah baris baru.
        sheet.getRange(rowIndex + 1, 1, 1, 11).setValues([record]);
        values[rowIndex] = record;
        updated++;
      } else {
        // Baris baru hanya untuk kombinasi Santri + Mapel yang belum ada.
        sheet.appendRow(record);
        values.push(record);
        added++;
      }
    });

    touchDataVersion_();
    return {
      success: true,
      updated: updated,
      added: added,
      message: 'Nilai tersimpan. ' + updated + ' nilai diperbarui' +
        (added ? ', ' + added + ' nilai baru ditambahkan.' : '; tidak ada baris baru ditambahkan.')
    };
  } finally {
    lock.releaseLock();
  }
}

function deleteRowByID(sheetName, id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheetByNameSafe_(ss, sheetName);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      touchDataVersion_();
      return { success: true, message: 'Data Berhasil Dihapus!' };
    }
  }
  return { success: false, message: 'Data tidak ditemukan!' };
}
function deleteRowsByIDs(sheetName, ids) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheetByNameSafe_(ss, sheetName);
  const values = sheet.getDataRange().getValues();
  let count = 0;
  for (let i = values.length - 1; i >= 1; i--) {
    if (ids.includes(String(values[i][0]))) {
      sheet.deleteRow(i + 1);
      count++;
    }
  }
  touchDataVersion_();
  return { success: true, message: `${count} Data Terpilih Berhasil Dihapus!` };
}
function clearSheetData(sheetName, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheetByNameSafe_(ss, sheetName);
  sheet.clearContents();
  sheet.appendRow(headers);
  touchDataVersion_();
  return { success: true, message: 'Semua Data Berhasil Dihapus!' };
}

/* ================================================================
   TAMBAHAN BACKEND MENU AKUN LOGIN GURU
   Tempelkan di BAGIAN PALING BAWAH Code.gs yang sedang dipakai.
   Struktur sheet Users: Username | Password | Nama | Role | ID Guru | Status
   Fungsi ini tidak menghapus atau mengubah data nilai/master Guru.
   ================================================================ */
function getGuruAccounts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = getSheetByNameSafe_(ss, 'Users');
  if (!sh) throw new Error('Sheet Users tidak ditemukan.');
  var values = sh.getDataRange().getDisplayValues();
  if (!values.length) return [];
  var headers = values[0].map(function(v) { return String(v || '').trim().toLowerCase(); });
  function col_(names) {
    for (var i = 0; i < names.length; i++) {
      var at = headers.indexOf(String(names[i]).trim().toLowerCase());
      if (at >= 0) return at;
    }
    return -1;
  }
  var cUser = col_(['username']), cName = col_(['nama', 'name']),
      cRole = col_(['role']), cGuru = col_(['id guru', 'id_guru', 'idguru']),
      cStatus = col_(['status']), cKelasWali = col_(['kelas wali', 'kelas_wali', 'kelaswali']);
  if (cUser < 0 || cName < 0 || cRole < 0) {
    throw new Error('Header Users harus memiliki Username, Nama, dan Role.');
  }
  var out = [];
  for (var r = 1; r < values.length; r++) {
    var row = values[r], role = String(row[cRole] || '').trim();
    var roleLower = role.toLowerCase();
    if (roleLower !== 'guru' && roleLower !== 'wali kelas' && roleLower !== 'guru + wali kelas' && roleLower !== 'guru+wali kelas' && roleLower !== 'wali kelas + guru') continue;
    out.push({
      username: String(row[cUser] || '').trim(),
      nama: String(row[cName] || '').trim(),
      role: role,
      idGuru: cGuru >= 0 ? String(row[cGuru] || '').trim() : '',
      kelasWali: cKelasWali >= 0 ? String(row[cKelasWali] || '').trim() : '',
      status: cStatus >= 0 ? String(row[cStatus] || 'Aktif').trim() || 'Aktif' : 'Aktif'
    });
  }
  return out;
}

function saveGuruAccount(payload) {
  payload = payload || {};
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = getSheetByNameSafe_(ss, 'Users');
  if (!sh) return { success: false, message: 'Sheet Users tidak ditemukan.' };

  var lastCol = Math.max(sh.getLastColumn(), 1);
  var headerValues = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  var headers = headerValues.map(function(v) { return String(v || '').trim().toLowerCase(); });

  function col_(names) {
    for (var i = 0; i < names.length; i++) {
      var at = headers.indexOf(String(names[i]).trim().toLowerCase());
      if (at >= 0) return at;
    }
    return -1;
  }
  function ensureColumn_(names, label) {
    var at = col_(names);
    if (at >= 0) return at;
    lastCol++;
    sh.getRange(1, lastCol).setValue(label);
    headers.push(label.toLowerCase());
    return headers.length - 1;
  }

  var cUser = col_(['username']), cPass = col_(['password']),
      cName = col_(['nama', 'name']), cRole = col_(['role']);
  if (cUser < 0 || cPass < 0 || cName < 0 || cRole < 0) {
    return { success: false, message: 'Header sheet Users minimal harus memiliki Username, Password, Nama, dan Role.' };
  }

  // Tambahkan kolom tambahan di paling kanan saja; data akun lama tidak digeser.
  var cGuru = ensureColumn_(['id guru', 'id_guru', 'idguru'], 'ID Guru');
  var cStatus = ensureColumn_(['status'], 'Status');
  var cKelasWali = ensureColumn_(['kelas wali', 'kelas_wali', 'kelaswali'], 'Kelas Wali');

  var values = sh.getDataRange().getDisplayValues();
  var username = String(payload.username || '').trim();
  var nama = String(payload.nama || payload.name || '').trim();
  var password = String(payload.password || '');
  var original = String(payload.originalUsername || '').trim();
  var roleInput = String(payload.role || 'Guru').trim().toLowerCase();
  var role = (roleInput === 'wali kelas') ? 'Wali Kelas' : ((roleInput === 'guru + wali kelas' || roleInput === 'guru+wali kelas' || roleInput === 'wali kelas + guru') ? 'Guru + Wali Kelas' : 'Guru');
  var idGuru = String(payload.idGuru || payload['ID Guru'] || '').trim();
  var kelasWali = String(payload.kelasWali || payload['Kelas Wali'] || payload.kelas || '').trim();
  var status = String(payload.status || 'Aktif').trim() || 'Aktif';

  if (['guru','wali kelas','guru + wali kelas','guru+wali kelas','wali kelas + guru'].indexOf(roleInput) < 0) {
    return { success: false, message: 'Role akun tidak valid. Pilih Guru, Wali Kelas, atau Guru + Wali Kelas.' };
  }
  if (!username || !nama || (!original && !password)) {
    return { success: false, message: 'Nama, username, dan password wajib diisi untuk akun baru.' };
  }
  if (!idGuru) {
    return { success: false, message: 'Silakan pilih nama guru dari daftar agar ID Guru terisi.' };
  }
  var needsWaliAssignment = role === 'Wali Kelas' || role === 'Guru + Wali Kelas';
  if (needsWaliAssignment && !kelasWali) {
    return { success: false, message: 'Kelas Wali wajib diisi untuk akun Wali Kelas atau Guru + Wali Kelas.' };
  }
  if (needsWaliAssignment) {
    var santriKelas = getSantriDirect();
    var kelasValid = santriKelas.some(function(s) {
      return String(s.Kelas || '').trim().toLowerCase() === kelasWali.toLowerCase();
    });
    if (!kelasValid) {
      return { success: false, message: 'Kelas Wali belum ditemukan pada data Santri. Periksa penulisan kelas.' };
    }
  }

  var targetRow = -1, oldPassword = '';
  for (var r = 1; r < values.length; r++) {
    var currentUser = String(values[r][cUser] || '').trim();
    var currentRole = String(values[r][cRole] || '').trim().toLowerCase();

    if (currentUser.toLowerCase() === username.toLowerCase() &&
        currentUser.toLowerCase() !== original.toLowerCase()) {
      return { success: false, message: 'Username sudah digunakan akun lain.' };
    }

    if (original && currentUser.toLowerCase() === original.toLowerCase() &&
        (currentRole === 'guru' || currentRole === 'wali kelas' || currentRole === 'guru + wali kelas' || currentRole === 'guru+wali kelas' || currentRole === 'wali kelas + guru')) {
      targetRow = r + 1;
      oldPassword = String(values[r][cPass] || '');
    }
  }
  if (original && targetRow < 0) {
    return { success: false, message: 'Akun Guru/Wali Kelas yang akan diedit tidak ditemukan.' };
  }

  var width = Math.max(sh.getLastColumn(), headers.length);
  var rowData;
  if (targetRow > 0) {
    rowData = sh.getRange(targetRow, 1, 1, width).getValues()[0];
    while (rowData.length < width) rowData.push('');
  } else {
    rowData = new Array(width).fill('');
  }

  rowData[cUser] = username;
  rowData[cPass] = password || oldPassword;
  rowData[cName] = nama;
  // Jangan paksa role menjadi Guru: simpan role yang benar-benar dipilih di formulir.
  rowData[cRole] = role;
  rowData[cGuru] = idGuru;
  rowData[cStatus] = status;
  rowData[cKelasWali] = needsWaliAssignment ? kelasWali : '';

  if (targetRow > 0) {
    sh.getRange(targetRow, 1, 1, width).setValues([rowData]);
  } else {
    sh.appendRow(rowData);
  }

  if (typeof touchDataVersion_ === 'function') touchDataVersion_();

  // Sistem akses Wali Kelas membaca sheet "Penugasan Wali Kelas".
  // Sinkronkan pilihan kelas dari formulir ke sheet penugasan yang sudah ada.
  if (needsWaliAssignment && status.toLowerCase() === 'aktif' &&
      typeof saveWaliKelasAssignment === 'function') {
    var assignmentResult = saveWaliKelasAssignment({ username: username, kelas: kelasWali });
    if (!assignmentResult || !assignmentResult.success) {
      return {
        success: false,
        message: 'Akun tersimpan, tetapi penetapan kelas belum berhasil: ' +
          (assignmentResult && assignmentResult.message ? assignmentResult.message : 'silakan periksa sheet Penugasan Wali Kelas.')
      };
    }
  } else if (!needsWaliAssignment && typeof deleteWaliKelasAssignment === 'function') {
    // Jika role diubah menjadi Guru biasa, lepaskan penetapan wali lama agar akses wali tidak tertinggal.
    try { deleteWaliKelasAssignment(username); } catch (ignore) {}
  }

  return {
    success: true,
    message: 'Akun ' + role + (targetRow > 0 ? ' berhasil diperbarui.' : ' berhasil ditambahkan.') +
      (needsWaliAssignment ? ' Kelas ' + kelasWali + ' juga sudah ditetapkan.' : '')
  };
}

function deleteGuruAccount(username) {
  username = String(username || '').trim();
  if (!username) return { success: false, message: 'Username akun tidak valid.' };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = getSheetByNameSafe_(ss, 'Users');
  if (!sh) throw new Error('Sheet Users tidak ditemukan.');
  var values = sh.getDataRange().getDisplayValues();
  if (!values.length) return { success: false, message: 'Sheet Users kosong.' };
  var headers = values[0].map(function(v) { return String(v || '').trim().toLowerCase(); });
  var cUser = headers.indexOf('username'), cRole = headers.indexOf('role');
  if (cUser < 0 || cRole < 0) return { success: false, message: 'Header Username atau Role tidak ditemukan.' };
  for (var r = values.length - 1; r >= 1; r--) {
    if (String(values[r][cUser] || '').trim().toLowerCase() === username.toLowerCase()) {
      var role = String(values[r][cRole] || '').trim().toLowerCase();
      if (role !== 'guru' && role !== 'wali kelas' && role !== 'guru + wali kelas' && role !== 'guru+wali kelas' && role !== 'wali kelas + guru') {
        return { success: false, message: 'Menu ini hanya dapat menghapus akun Guru atau Wali Kelas.' };
      }
      sh.deleteRow(r + 1);
      if (typeof touchDataVersion_ === 'function') touchDataVersion_();
      return { success: true, message: 'Akun login berhasil dihapus. Data master Guru dan nilai tetap aman.' };
    }
  }
  return { success: false, message: 'Akun Guru/Wali Kelas tidak ditemukan.' };
}


/* TUGAS GURU: CRUD aman; tidak menghapus master guru, mapel, atau nilai. */
function ensureTugasGuruSheet_() {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let sh=getSheetByNameSafe_(ss,SHEET_NAMES.TUGAS_GURU);
  const headers=['ID','ID Guru','Nama Guru','ID Mapel','Mata Pelajaran','Kelas','Tingkatan','Hari','Jam Mulai','Jam Selesai','Tahun Pelajaran','Status'];
  if(!sh){sh=ss.insertSheet(SHEET_NAMES.TUGAS_GURU);sh.getRange(1,1,1,headers.length).setValues([headers]);sh.setFrozenRows(1);sh.getRange(1,1,1,headers.length).setBackground('#0f766e').setFontColor('#ffffff').setFontWeight('bold');}
  else if(sh.getLastRow()===0)sh.getRange(1,1,1,headers.length).setValues([headers]);
  return sh;
}
function getTugasGuru() {
  const sh=ensureTugasGuruSheet_(),v=sh.getDataRange().getDisplayValues();
  if(v.length<=1)return [];
  const h=v[0].map(x=>String(x||'').trim().toLowerCase());
  const col=names=>{for(let j=0;j<names.length;j++){const k=h.indexOf(names[j]);if(k>=0)return k;}return -1;};
  const f=[['ID',['id']],['ID Guru',['id guru','id_guru','idguru']],['Nama Guru',['nama guru']],['ID Mapel',['id mapel','id_mapel','idmapel']],['Mata Pelajaran',['mata pelajaran','mapel']],['Kelas',['kelas']],['Tingkatan',['tingkatan','tingkat']],['Hari',['hari']],['Jam Mulai',['jam mulai']],['Jam Selesai',['jam selesai']],['Tahun Pelajaran',['tahun pelajaran','tahun']],['Status',['status']]];
  return v.slice(1).map(r=>{const o={};f.forEach(x=>{const c=col(x[1]);o[x[0]]=c>=0?r[c]:'';});return o;}).filter(o=>o.ID&&o['ID Guru']&&o['ID Mapel']);
}
// Normalisasi identitas untuk membandingkan data lama dan input baru.
// Menghapus spasi/karakter tak terlihat dan menyamakan pemisah tahun 2026-2027 / 2026/2027.
function kmmiNormTaskText_(value) {
  return String(value == null ? '' : value)
    .normalize('NFKC')
    .replace(/[\u200B-\u200D\uFEFF\u00AD]/g, '')
    .toLowerCase()
    .trim()
    .replace(/[‐‑‒–—−]/g, '-')
    .replace(/[\s\u00A0]+/g, '');
}
function kmmiNormTaskYear_(value) {
  return kmmiNormTaskText_(value)
    .replace(/[\/\\]/g, '-')
    .replace(/-+/g, '-');
}
// Format tampilan/penyimpanan tahun pelajaran wajib memakai strip (-).
// Contoh: 2026/2027, 2026\2027, atau 2026–2027 menjadi 2026-2027.
function kmmiFormatTaskYear_(value) {
  return String(value == null ? '' : value)
    .normalize('NFKC')
    .replace(/[\u200B-\u200D\uFEFF\u00AD]/g, '')
    .trim()
    .replace(/[‐‑‒–—−]/g, '-')
    .replace(/[\/\\]/g, '-')
    .replace(/\s*-\s*/g, '-')
    .replace(/-+/g, '-');
}

// Jalankan manual satu kali jika ingin menyeragamkan format tahun di Sheet Tugas Guru.
// Fungsi ini hanya mengganti tanda pemisah tahun; tidak menghapus atau menggabungkan baris.
function seragamkanStripTahunPelajaranTugasGuru() {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = ensureTugasGuruSheet_();
    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow < 2 || lastCol < 1) {
      return {success:true, updated:0, message:'Tidak ada data Tugas Guru untuk diseragamkan.'};
    }
    const values = sh.getRange(1, 1, lastRow, lastCol).getDisplayValues();
    const headers = values[0].map(v => String(v == null ? '' : v).trim().toLowerCase().replace(/\s+/g, ' '));
    const yearCol = headers.findIndex(h => h === 'tahun pelajaran' || h === 'tahun');
    if (yearCol < 0) {
      return {success:false, updated:0, message:'Kolom Tahun Pelajaran tidak ditemukan. Tidak ada data yang diubah.'};
    }
    let updated = 0;
    for (let r = 1; r < values.length; r++) {
      const before = String(values[r][yearCol] == null ? '' : values[r][yearCol]).trim();
      const after = kmmiFormatTaskYear_(before);
      if (before && after !== before) {
        sh.getRange(r + 1, yearCol + 1).setValue(after);
        updated++;
      }
    }
    if (updated && typeof touchDataVersion_ === 'function') touchDataVersion_();
    SpreadsheetApp.flush();
    const message = updated + ' baris Tahun Pelajaran diseragamkan memakai tanda strip (-). Data tidak dihapus.';
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) ss.toast(message, 'Format Tahun Pelajaran', 8);
    return {success:true, updated:updated, message:message};
  } finally {
    lock.releaseLock();
  }
}

function saveTugasGuru(p) {
  p = p || {};
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = ensureTugasGuruSheet_();
    const v = sh.getDataRange().getDisplayValues();
    const h = (v[0] || []).map(x => String(x || '').trim().toLowerCase());
    const fields = ['ID','ID Guru','Nama Guru','ID Mapel','Mata Pelajaran','Kelas','Tingkatan','Hari','Jam Mulai','Jam Selesai','Tahun Pelajaran','Status'];
    const d = {
      'ID': String(p.id || '').trim() || 'TG-' + Utilities.getUuid().slice(0,8).toUpperCase(),
      'ID Guru': String(p.idGuru || '').trim(),
      'Nama Guru': String(p.namaGuru || '').trim(),
      'ID Mapel': String(p.idMapel || '').trim(),
      'Mata Pelajaran': String(p.mapel || '').trim(),
      'Kelas': String(p.kelas || '').trim(),
      'Tingkatan': String(p.tingkatan || '').trim(),
      'Hari': String(p.hari || '').trim(),
      'Jam Mulai': String(p.jamMulai || '').trim(),
      'Jam Selesai': String(p.jamSelesai || '').trim(),
      'Tahun Pelajaran': kmmiFormatTaskYear_(p.tahunPelajaran || ''),
      'Status': String(p.status || 'Aktif').trim()
    };
    if (!d['ID Guru'] || !d['Nama Guru'] || !d['ID Mapel'] || !d['Mata Pelajaran'] || !d.Kelas || !d['Tahun Pelajaran']) {
      return {success:false, message:'Guru, mata pelajaran, kelas, dan tahun pelajaran wajib dipilih.'};
    }
    if (d['Jam Mulai'] && d['Jam Selesai'] && d['Jam Selesai'] <= d['Jam Mulai']) {
      return {success:false, message:'Jam selesai harus setelah jam mulai.'};
    }
    const col = name => h.indexOf(name.toLowerCase());
    const idc=col('ID'), gc=col('ID Guru'), gnc=col('Nama Guru'), mc=col('ID Mapel'), mnc=col('Mata Pelajaran'), kc=col('Kelas'), yc=col('Tahun Pelajaran');
    const editId = String(p.id || '').trim();
    const norm = kmmiNormTaskText_;
    const sameSubject = (row) => {
      const rowMapelId = mc >= 0 ? String(row[mc] || '').trim() : '';
      const rowMapelName = mnc >= 0 ? norm(row[mnc]) : '';
      const incomingMapelId = d['ID Mapel'];
      const incomingMapelName = norm(d['Mata Pelajaran']);
      // Cocokkan ID DAN nama: nama mapel sama tetap dianggap sama meski ID master berbeda.
      return (!!rowMapelId && !!incomingMapelId && rowMapelId === incomingMapelId) ||
             (!!rowMapelName && !!incomingMapelName && rowMapelName === incomingMapelName);
    };
    let targetRow = -1;
    for (let r=1; r<v.length; r++) {
      const row = v[r];
      const rowId = idc >= 0 ? String(row[idc] || '').trim() : '';
      if (editId && rowId === editId) { targetRow = r + 1; continue; }
      // Nama guru menjadi fallback identitas. Ini penting karena data lama
      // kadang memakai ID Guru berbeda meskipun nama guru yang tampil sama.
      const rowGuruId = gc >= 0 ? norm(row[gc]) : '';
      const rowGuruName = gnc >= 0 ? norm(row[gnc]) : '';
      const incomingGuruId = norm(d['ID Guru']);
      const incomingGuruName = norm(d['Nama Guru']);
      const sameTeacher = (!!rowGuruId && !!incomingGuruId && rowGuruId === incomingGuruId) ||
                          (!!rowGuruName && !!incomingGuruName && rowGuruName === incomingGuruName);
      const sameClass = kc >= 0 && norm(row[kc]) === norm(d.Kelas);
      const sameYear = yc >= 0 && kmmiNormTaskYear_(row[yc]) === kmmiNormTaskYear_(d['Tahun Pelajaran']);
      if (sameTeacher && sameClass && sameYear && sameSubject(row)) {
        return {success:false, message:'TUGAS SUDAH TERDAFTAR. Guru, mata pelajaran, kelas, dan tahun pelajaran ini sudah memiliki penugasan. Satu penugasan berlaku selama satu tahun pelajaran dengan jadwal pertemuan mingguan. Silakan gunakan tombol Edit untuk mengubah hari atau jam mengajar. Sistem tidak akan membuat penugasan baru yang sama.'};
      }
    }
    if (editId && targetRow < 0) {
      return {success:false, message:'Tugas yang akan diedit tidak ditemukan. Muat ulang daftar tugas lalu coba kembali.'};
    }
    const rowValues = fields.map(x => d[x]);
    if (targetRow > 0) sh.getRange(targetRow,1,1,rowValues.length).setValues([rowValues]);
    else sh.appendRow(rowValues);
    if (typeof touchDataVersion_ === 'function') touchDataVersion_();
    return {success:true, message:targetRow > 0 ? 'Jadwal tugas guru berhasil diperbarui.' : 'Tugas guru berhasil ditambahkan. Jadwal berlaku mingguan selama tahun pelajaran.'};
  } finally {
    lock.releaseLock();
  }
}
/**
 * Pemeriksaan duplikat lama Tugas Guru.
 * Aturan identitas: nama guru + nama mata pelajaran + kelas + tahun pelajaran.
 * Baris paling atas dipertahankan; duplikat yang lebih bawah dihapus.
 * Jalankan manual satu kali setelah membuat salinan spreadsheet.
 */
function bersihkanDuplikatTugasGuruSekali() {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    const sh = ensureTugasGuruSheet_();
    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow < 2 || lastCol < 1) {
      const emptyMsg = 'Tidak ada data tugas untuk diperiksa.';
      if (ss) ss.toast(emptyMsg, 'Pemeriksaan Duplikat Tugas Guru', 8);
      Logger.log(emptyMsg);
      return {success:true, removed:0, kept:0, message:emptyMsg};
    }
    const values = sh.getRange(1, 1, lastRow, lastCol).getDisplayValues();
    const headers = values[0].map(v => String(v == null ? '' : v).normalize('NFKC').replace(/[\u200B-\u200D\uFEFF\u00AD]/g, '').toLowerCase().trim().replace(/[\s\u00A0]+/g, ' '));
    const findCol = names => headers.findIndex(h => names.some(n => h === n));
    const cId = findCol(['id']);
    const cGuruId = findCol(['id guru', 'idguru']);
    const cGuru = findCol(['nama guru', 'guru']);
    const cMapelId = findCol(['id mapel', 'idmapel']);
    const cMapel = findCol(['mata pelajaran', 'mapel']);
    const cKelas = findCol(['kelas']);
    const cTahun = findCol(['tahun pelajaran', 'tahun']);
    if ([cGuru,cMapel,cKelas,cTahun].some(c => c < 0)) {
      const msg = 'Header tugas belum cocok. Kolom yang terbaca: ' + headers.join(' | ') + '. Tidak ada data yang dihapus.';
      if (ss) ss.toast(msg, 'Pemeriksaan Duplikat Tugas Guru', 10);
      Logger.log(msg);
      return {success:false, removed:0, message:msg};
    }
    const seen = new Map();
    const rowsToDelete = [];
    const duplicateDetails = [];
    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      const guruId = cGuruId >= 0 ? kmmiNormTaskText_(row[cGuruId]) : '';
      const guruName = kmmiNormTaskText_(row[cGuru]);
      const mapelId = cMapelId >= 0 ? kmmiNormTaskText_(row[cMapelId]) : '';
      const mapelName = kmmiNormTaskText_(row[cMapel]);
      const kelas = kmmiNormTaskText_(row[cKelas]);
      const tahun = kmmiNormTaskYear_(row[cTahun]);
      if ((!guruId && !guruName) || (!mapelId && !mapelName) || !kelas || !tahun) continue;
      // Gunakan ID bila ada; fallback ke nama. Tahun dinormalisasi agar tanda / dan - setara.
      const teacherKey = guruName || guruId;
      const subjectKey = mapelName || mapelId;
      const key = [teacherKey, subjectKey, kelas, tahun].join('|');
      if (seen.has(key)) {
        rowsToDelete.push(r + 1);
        duplicateDetails.push('baris ' + (r + 1) + ' duplikat dari baris ' + seen.get(key) + ': ' + row[cGuru] + ' / ' + row[cMapel] + ' / kelas ' + row[cKelas] + ' / ' + row[cTahun]);
      } else {
        seen.set(key, r + 1);
      }
    }
    // Hapus dari bawah supaya nomor baris yang belum diproses tidak bergeser.
    for (let i = rowsToDelete.length - 1; i >= 0; i--) sh.deleteRow(rowsToDelete[i]);
    if (rowsToDelete.length && typeof touchDataVersion_ === 'function') touchDataVersion_();
    SpreadsheetApp.flush();
    const msg = rowsToDelete.length
      ? rowsToDelete.length + ' baris duplikat berhasil dihapus. Baris yang dipertahankan: ' + seen.size + ' tugas unik.'
      : 'Tidak ditemukan duplikat menurut ID/nama guru, ID/nama mapel, kelas, dan tahun yang sudah dinormalisasi.';
    if (ss) ss.toast(msg, 'Pemeriksaan Duplikat Tugas Guru', 10);
    Logger.log(msg);
    duplicateDetails.forEach(d => Logger.log(d));
    return {success:true, removed:rowsToDelete.length, kept:seen.size, message:msg, details:duplicateDetails};
  } finally {
    lock.releaseLock();
  }
}

function deleteTugasGuru(id) {
  id = String(id == null ? '' : id).trim();
  if (!id) return { success: false, message: 'ID tugas tidak valid.' };

  // Gunakan header kolom ID, jangan menganggap ID selalu berada di kolom A.
  // Ini menjaga kompatibilitas dengan sheet Tugas Guru yang urutan kolomnya berbeda.
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = ensureTugasGuruSheet_();
    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow < 2 || lastCol < 1) {
      return { success: false, message: 'Tugas guru tidak ditemukan.' };
    }

    const values = sh.getRange(1, 1, lastRow, lastCol).getDisplayValues();
    const headers = values[0].map(function (x) {
      return String(x == null ? '' : x).trim().toLowerCase().replace(/\s+/g, ' ');
    });
    const idCol = headers.indexOf('id');
    if (idCol < 0) {
      return { success: false, message: 'Kolom ID tidak ditemukan pada sheet Tugas Guru. Data tidak dihapus.' };
    }

    for (let r = values.length - 1; r >= 1; r--) {
      if (String(values[r][idCol] == null ? '' : values[r][idCol]).trim() === id) {
        sh.deleteRow(r + 1);
        if (typeof touchDataVersion_ === 'function') touchDataVersion_();
        return { success: true, message: 'Tugas guru berhasil dihapus. Data guru dan nilai tetap aman.' };
      }
    }
    return { success: false, message: 'ID tugas tidak ditemukan pada kolom ID. Tidak ada data yang dihapus.' };
  } finally {
    lock.releaseLock();
  }
}


/* PENUGASAN INPUT NILAI - terpisah dari Tugas Guru/Mengajar.
   Sheet baru hanya dibuat saat menu ini dibuka atau assignment disimpan. */
function ensurePenugasanNilaiSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = getSheetByNameSafe_(ss, 'Penugasan Input Nilai');
  var headers = ['ID', 'ID Guru', 'Nama Guru', 'ID Mapel', 'Mata Pelajaran', 'Kelas', 'Tahun Pelajaran', 'Status'];
  if (!sh) {
    sh = ss.insertSheet('Penugasan Input Nilai');
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setBackground('#0f766e').setFontColor('#ffffff').setFontWeight('bold');
  } else if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sh;
}
function getPenugasanInputNilai() {
  var sh = ensurePenugasanNilaiSheet_();
  var values = sh.getDataRange().getDisplayValues();
  if (values.length < 2) return [];
  var headers = values[0].map(function(x){return String(x || '').trim().toLowerCase();});
  function col(names){for(var i=0;i<names.length;i++){var n=headers.indexOf(names[i]);if(n>=0)return n;}return -1;}
  var ix = {
    id:col(['id']), idGuru:col(['id guru','id_guru']), namaGuru:col(['nama guru','nama_guru']),
    idMapel:col(['id mapel','id_mapel']), mapel:col(['mata pelajaran','mapel']),
    kelas:col(['kelas']), tahun:col(['tahun pelajaran','tahun']), status:col(['status'])
  };
  return values.slice(1).map(function(r){return {
    ID:r[ix.id]||'', 'ID Guru':r[ix.idGuru]||'', 'Nama Guru':r[ix.namaGuru]||'',
    'ID Mapel':r[ix.idMapel]||'', 'Mata Pelajaran':r[ix.mapel]||'',
    Kelas:r[ix.kelas]||'', 'Tahun Pelajaran':r[ix.tahun]||'', Status:r[ix.status]||'Aktif'
  };}).filter(function(o){return o.ID && o['ID Guru'] && o['ID Mapel'] && o.Kelas;});
}
function savePenugasanInputNilai(p) {
  p = p || {};
  var sh = ensurePenugasanNilaiSheet_(), values = sh.getDataRange().getDisplayValues();
  var id = String(p.id || '').trim();
  var d = [
    id || ('PN-' + Utilities.getUuid().slice(0,8).toUpperCase()),
    String(p.idGuru || '').trim(), String(p.namaGuru || '').trim(),
    String(p.idMapel || '').trim(), String(p.mapel || '').trim(),
    String(p.kelas || '').trim(), String(p.tahunPelajaran || '').trim(),
    String(p.status || 'Aktif').trim() || 'Aktif'
  ];
  if (!d[1] || !d[2] || !d[3] || !d[4] || !d[5]) return {success:false,message:'Guru, mata pelajaran, dan kelas wajib dipilih.'};
  for (var i=1;i<values.length;i++) {
    if (String(values[i][0]||'')===d[0]) continue;
    if (String(values[i][1]||'')===d[1] && String(values[i][3]||'')===d[3] && String(values[i][5]||'')===d[5] && String(values[i][6]||'')===d[6] && String(values[i][7]||'Aktif').toLowerCase()==='aktif' && d[7].toLowerCase()==='aktif') return {success:false,message:'Penugasan input nilai yang sama sudah ada.'};
  }
  var target=-1;
  for (var j=1;j<values.length;j++) if(String(values[j][0]||'')===d[0]) {target=j+1;break;}
  if(target>0) sh.getRange(target,1,1,d.length).setValues([d]); else sh.appendRow(d);
  touchDataVersion_();
  return {success:true,message:target>0?'Penugasan input nilai diperbarui.':'Penugasan input nilai ditambahkan.'};
}
function deletePenugasanInputNilai(id) {
  id=String(id||'').trim(); if(!id)return {success:false,message:'ID penugasan tidak valid.'};
  var sh=ensurePenugasanNilaiSheet_(), v=sh.getDataRange().getDisplayValues();
  for(var i=v.length-1;i>=1;i--) if(String(v[i][0]||'')===id){sh.deleteRow(i+1);touchDataVersion_();return {success:true,message:'Penugasan input nilai dihapus.'};}
  return {success:false,message:'Penugasan tidak ditemukan.'};
}

/* ===== KMMI PATCH: BACKEND ABSENSI GURU + WALI KELAS =====
   Tambahan fungsi kompatibilitas untuk Index terbaru.
   Tidak menghapus atau mengganti sheet/data lama; header yang belum ada
   hanya ditambahkan di kolom paling kanan.
*/
function _kmmiNorm_(v) {
  return String(v == null ? '' : v).trim().toLowerCase().replace(/[\s_-]+/g, ' ');
}
function _kmmiFindCol_(headers, aliases) {
  var h = headers.map(_kmmiNorm_);
  for (var i = 0; i < aliases.length; i++) {
    var ix = h.indexOf(_kmmiNorm_(aliases[i]));
    if (ix >= 0) return ix;
  }
  return -1;
}
function _kmmiWideSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = 'REKAP ABSENSI LEBAR';
  var sh = getSheetByNameSafe_(ss, name);
  if (!sh) sh = ss.insertSheet(name);
  var base = ['Tanggal', 'Kelas', 'NIS', 'Nama Santri'];
  if (sh.getLastRow() === 0) sh.getRange(1,1,1,base.length).setValues([base]);
  var h = sh.getRange(1,1,1,Math.max(1,sh.getLastColumn())).getDisplayValues()[0];
  if (!h[0] || _kmmiNorm_(h[0]) !== 'tanggal') {
    // Preserve any existing visible recap by moving it aside before initializing storage.
    var old = getSheetByNameSafe_(ss, 'REKAP ABSENSI LEBAR LAMA');
    if (!old) { old = ss.insertSheet('REKAP ABSENSI LEBAR LAMA'); }
    if (sh.getLastRow() > 0 && old.getLastRow() === 0) {
      var vals = sh.getDataRange().getValues();
      old.getRange(1,1,vals.length,vals[0].length).setValues(vals);
    }
    sh.clear();
    sh.getRange(1,1,1,base.length).setValues([base]);
    h = base.slice();
  }
  sh.setFrozenRows(1);
  return sh;
}
function _kmmiEnsureAttendanceSheet_() {
  // Satu-satunya sheet penyimpanan absensi adalah REKAP ABSENSI LEBAR.
  // Pastikan sheet tujuan sudah tersedia sebelum menghapus sheet ABSENSI lama.
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tujuan = _kmmiWideSheet_();
  var lama = getSheetByNameSafe_(ss, 'ABSENSI');
  if (lama && lama.getSheetId() !== tujuan.getSheetId()) {
    ss.deleteSheet(lama);
  }
  return tujuan;
}

// Jalankan fungsi ini satu kali jika sheet ABSENSI lama masih terlihat.
// Fungsi menolak menghapus sheet lama jika REKAP ABSENSI LEBAR belum siap.
function hapusSheetAbsensiLama() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tujuan = getSheetByNameSafe_(ss, 'REKAP ABSENSI LEBAR');
  if (!tujuan || tujuan.getLastRow() < 1) {
    throw new Error('REKAP ABSENSI LEBAR belum tersedia. Sheet ABSENSI tidak dihapus.');
  }
  var lama = getSheetByNameSafe_(ss, 'ABSENSI');
  if (!lama) return 'Sheet ABSENSI sudah tidak ada. REKAP ABSENSI LEBAR tetap digunakan.';
  if (lama.getSheetId() === tujuan.getSheetId()) {
    throw new Error('Sheet tujuan tidak boleh dihapus.');
  }
  ss.deleteSheet(lama);
  return 'Sheet ABSENSI lama berhasil dihapus. Penyimpanan tetap di REKAP ABSENSI LEBAR.';
}
function _kmmiGetRows_(sheet) {
  if (!sheet || sheet.getLastRow() < 1) return {headers: [], rows: []};
  var vals = sheet.getDataRange().getDisplayValues();
  var headers = vals[0] || [];
  // Ubah bentuk lebar menjadi record virtual agar fungsi baca lama tetap kompatibel.
  if (_kmmiNorm_(headers[0]) === 'tanggal' && _kmmiNorm_(headers[1]) === 'kelas' && _kmmiNorm_(headers[2]) === 'nis') {
    var virtualHeaders = ['ID','NIS','Nama Santri','Kelas','Tanggal','Status','Keterangan','ID Guru','Username','Nama Penginput','Mata Pelajaran','Waktu Update'];
    var rows = [];
    for (var r=1; r<vals.length; r++) {
      var line=vals[r];
      if (!line[0] || !line[1] || !line[2]) continue;
      for (var c=4; c<headers.length; c++) {
        var hn=String(headers[c]||'');
        if (!hn || hn.indexOf('__META__')===0) continue;
        var status=String(line[c]||'').trim();
        if (!status) continue;
        var metaCol=headers.indexOf('__META__'+hn);
        var meta={};
        if (metaCol>=0 && line[metaCol]) { try { meta=JSON.parse(line[metaCol]); } catch(e) {} }
        rows.push([
          meta.id || '', line[2], line[3], line[1], line[0], status,
          meta.keterangan || '', meta.idGuru || '', meta.username || '',
          meta.namaPenginput || '', hn, meta.waktuUpdate || ''
        ]);
      }
    }
    return {headers:virtualHeaders, rows:rows};
  }
  return {headers: headers, rows: vals.slice(1)};
}
function _kmmiAttendanceObject_(row, headers) {
  function get(aliases) { var i = _kmmiFindCol_(headers, aliases); return i >= 0 ? row[i] : ''; }
  return {
    ID: get(['ID']),
    nis: get(['NIS', 'ID Santri']),
    NIS: get(['NIS', 'ID Santri']),
    nama: get(['Nama Santri', 'Nama']),
    'Nama Santri': get(['Nama Santri', 'Nama']),
    kelas: get(['Kelas']),
    Kelas: get(['Kelas']),
    mapel: get(['Mata Pelajaran', 'Mapel']),
    'Mata Pelajaran': get(['Mata Pelajaran', 'Mapel']),
    tanggal: get(['Tanggal', 'Tanggal Absensi', 'Date']),
    Tanggal: get(['Tanggal', 'Tanggal Absensi', 'Date']),
    status: get(['Status', 'Kehadiran']),
    Status: get(['Status', 'Kehadiran']),
    keterangan: get(['Keterangan', 'Catatan']),
    Keterangan: get(['Keterangan', 'Catatan']),
    idGuru: get(['ID Guru', 'ID_Guru']),
    username: get(['Username', 'Akun']),
    namaPenginput: get(['Nama Penginput', 'Nama Guru', 'Nama Wali'])
  };
}
function _kmmiWriteAttendance_(records, idGuru, username, namaPenginput, mapel) {
  if (!Array.isArray(records) || !records.length) {
    return {success:false,message:'Tidak ada data absensi untuk disimpan.'};
  }
  var lock=LockService.getDocumentLock();
  if (!lock.tryLock(10000)) {
    return {success:false,message:'Absensi sedang disimpan oleh proses lain. Coba lagi sebentar.'};
  }
  try {
    var sh=_kmmiWideSheet_();
    var user=String(username||'').trim(), guru=String(idGuru||'').trim();
    var subject=String(mapel||'').trim();
    var added=0, updated=0, skipped=0;
    var keyNorm=function(v){return String(v==null?'':v).trim().toLowerCase().replace(/\s+/g,' ');};

    function dateKey(v) {
      if (v instanceof Date && !isNaN(v.getTime())) {
        return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
      var s=String(v==null?'':v).trim();
      if (!s) return '';
      var m=s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
      if (m) return m[1]+'-'+('0'+m[2]).slice(-2)+'-'+('0'+m[3]).slice(-2);
      m=s.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})$/);
      if (m) return m[3]+'-'+('0'+m[2]).slice(-2)+'-'+('0'+m[1]).slice(-2);
      var d=new Date(s);
      return isNaN(d.getTime()) ? s.slice(0,10) :
        Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
    function rowKey(date, kelas, nis) {
      return dateKey(date)+'|'+keyNorm(kelas)+'|'+String(nis==null?'':nis).trim();
    }

    // Siapkan seluruh kolom mapel dan metadata sebelum membaca data,
    // sehingga tidak perlu menambah kolom di tengah proses tiap santri.
    var lastCol=Math.max(4,sh.getLastColumn());
    var headers=sh.getRange(1,1,1,lastCol).getDisplayValues()[0];
    while(headers.length<4) headers.push('');
    var subjects={};
    records.forEach(function(r) {
      r=r||{};
      var s=String(subject||r.mapel||r['Mata Pelajaran']||'Pelajaran Wali Kelas').trim();
      if(s) subjects[s]=true;
    });
    Object.keys(subjects).forEach(function(s) {
      if(headers.indexOf(s)<0) headers.push(s);
      var meta='__META__'+s;
      if(headers.indexOf(meta)<0) headers.push(meta);
    });

    // Tambah kolom baru sekaligus, bukan satu per santri.
    var actualLastCol=sh.getLastColumn();
    if(headers.length>actualLastCol) {
      if(actualLastCol<1) actualLastCol=1;
      sh.insertColumnsAfter(actualLastCol,headers.length-actualLastCol);
    }
    sh.getRange(1,1,1,headers.length).setValues([headers]);

    // Ambil isi sheet sekali saja.
    var lastRow=sh.getLastRow();
    var rawRows=lastRow>=2 ? sh.getRange(2,1,lastRow-1,headers.length).getValues() : [];
    var displayRows=lastRow>=2 ? sh.getRange(2,1,lastRow-1,headers.length).getDisplayValues() : [];
    var indexByKey={};
    for(var i=0;i<displayRows.length;i++) {
      var k=rowKey(displayRows[i][0],displayRows[i][1],displayRows[i][2]);
      if(k && k.indexOf('|')>=0 && indexByKey[k]===undefined) indexByKey[k]=i;
    }

    // Perbarui buffer di memori. Jika satu santri muncul dua kali pada
    // batch yang sama, kedua record diarahkan ke baris yang sama.
    records.forEach(function(r) {
      r=r||{};
      var nis=String(r.nis||r.NIS||'').trim();
      var kelas=String(r.kelas||r.Kelas||'').trim();
      var tanggal=dateKey(r.tanggal||r.Tanggal||'');
      var status=String(r.status||r.Status||'').trim();
      var thisSubject=String(subject||r.mapel||r['Mata Pelajaran']||'Pelajaran Wali Kelas').trim();
      if(!nis||!kelas||!tanggal||!status||headers.indexOf(thisSubject)<0) {
        skipped++;
        return;
      }

      var k=rowKey(tanggal,kelas,nis);
      var idx=indexByKey[k];
      if(idx===undefined) {
        idx=rawRows.length;
        var newRow=new Array(headers.length).fill('');
        newRow[0]=tanggal;
        newRow[1]=kelas;
        newRow[2]=nis;
        newRow[3]=String(r.nama||r['Nama Santri']||'').trim();
        rawRows.push(newRow);
        displayRows.push(newRow.map(function(v){return String(v==null?'':v);}));
        indexByKey[k]=idx;
        added++;
      } else {
        updated++;
        if(!rawRows[idx][3]) rawRows[idx][3]=String(r.nama||r['Nama Santri']||'').trim();
      }

      var statusIx=headers.indexOf(thisSubject);
      rawRows[idx][statusIx]=status;
      var metaIx=headers.indexOf('__META__'+thisSubject);
      var oldMeta={};
      var oldRaw=rawRows[idx][metaIx];
      if(oldRaw) { try { oldMeta=JSON.parse(String(oldRaw))||{}; } catch(e) {} }
      var meta={
        id:oldMeta.id||('ABS-'+Utilities.getUuid().slice(0,8).toUpperCase()),
        keterangan:String(r.keterangan||r.Keterangan||oldMeta.keterangan||'').trim(),
        idGuru:guru||String(r.idGuru||oldMeta.idGuru||'').trim(),
        username:user||String(r.username||oldMeta.username||'').trim(),
        namaPenginput:String(namaPenginput||r.namaPenginput||oldMeta.namaPenginput||'').trim(),
        waktuUpdate:new Date().toISOString()
      };
      rawRows[idx][metaIx]=JSON.stringify(meta);
    });

    // Satu penulisan berkelompok untuk seluruh baris yang tersimpan.
    if(rawRows.length) {
      sh.getRange(2,1,rawRows.length,headers.length).setValues(rawRows);
    }
    headers.forEach(function(h,i) {
      if(String(h).indexOf('__META__')===0) {
        try { sh.hideColumns(i+1); } catch(e) {}
      }
    });
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,headers.length)
      .setBackground('#0f766e').setFontColor('#ffffff').setFontWeight('bold').setWrap(true);
    try { if(typeof touchDataVersion_==='function') touchDataVersion_(); } catch(e) {}
    SpreadsheetApp.flush();
    return {
      success:true,
      message:'Absensi tersimpan di REKAP ABSENSI LEBAR.',
      added:added, updated:updated, skipped:skipped
    };
  } catch(err) {
    return {success:false,message:'Gagal menyimpan absensi: '+(err&&err.message?err.message:err)};
  } finally {
    lock.releaseLock();
  }
}
function getKelasAbsensiGuru(idGuru, username) {
  var id = String(idGuru || '').trim();
  var user = String(username || '').trim();
  var seen = {}, result = [];
  // Kelas dari tugas mengajar Guru.
  if (id) {
    getTugasGuru().filter(function(t) {
      var active = _kmmiNorm_(t.Status || 'Aktif');
      return String(t['ID Guru'] || '').trim() === id && (!active || active === 'aktif');
    }).forEach(function(t) {
      var kelas=String(t.Kelas || '').trim();
      if(kelas && !seen[kelas]) { seen[kelas]=true; result.push({kelas:kelas,mapel:String(t['Mata Pelajaran'] || '').trim(),idMapel:String(t['ID Mapel'] || '').trim()}); }
    });
  }
  // Kelas Wali ditambahkan untuk akun Wali murni maupun Guru yang juga Wali.
  if (user) {
    getWaliKelasAssignments().forEach(function(a) {
      var active = _kmmiNorm_(a.Status || 'Aktif');
      var kelas=String(a.Kelas || '').trim();
      if(String(a.Username || '').trim().toLowerCase()===user.toLowerCase() && (!active || active==='aktif') && kelas && !seen[kelas]) {
        seen[kelas]=true; result.push({kelas:kelas,mapel:'Absensi Santri',idMapel:'WALI'});
      }
    });
  }
  return result.sort(function(a,b){return a.kelas.localeCompare(b.kelas,undefined,{numeric:true});});
}
function saveAbsensiSantriBatchTugasGuru(records, idGuru, username, namaGuru) {
  var id = String(idGuru || '').trim(), user = String(username || '').trim();
  if (!id && !user) return {success:false,message:'Akun belum terhubung ke ID Guru/username.'};
  var allowed = {};
  // Kelas mengajar aktif.
  if (id) getTugasGuru().filter(function(t) {
    return String(t['ID Guru'] || '').trim() === id && _kmmiNorm_(t.Status || 'Aktif') === 'aktif';
  }).forEach(function(t){var k=String(t.Kelas || '').trim();if(k)allowed[k]='guru';});
  // Kelas Wali yang aktif dan ditetapkan Admin.
  if (user) getWaliKelasAssignments().forEach(function(a) {
    if(String(a.Username || '').trim().toLowerCase()===user.toLowerCase() && _kmmiNorm_(a.Status || 'Aktif')==='aktif') {
      var k=String(a.Kelas || '').trim();if(k && !allowed[k])allowed[k]='wali';
    }
  });
  var filtered = (records || []).filter(function(r){return allowed[String(r.kelas || r.Kelas || '').trim()];});
  if (!filtered.length) return {success:false,message:'Kelas absensi tidak cocok dengan Tugas Guru atau penugasan Wali Kelas aktif.'};
  // Simpan identitas akun agar status tersimpan dapat dibaca kembali pada akun yang sama.
  return _kmmiWriteAttendance_(filtered,id,user,namaGuru,'');
}
function _kmmiEnsureWaliAssignmentSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = getSheetByNameSafe_(ss, 'Penugasan Wali Kelas');
  if (!sh) {
    sh = ss.insertSheet('Penugasan Wali Kelas');
    sh.getRange(1,1,1,5).setValues([['ID','Username','Nama Wali Kelas','Kelas','Status']]);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,5).setBackground('#6d28d9').setFontColor('#ffffff').setFontWeight('bold');
  } else if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,5).setValues([['ID','Username','Nama Wali Kelas','Kelas','Status']]);
  }
  return sh;
}
function getWaliKelasUsersForAdmin() {
  // Akun Guru yang sudah ada juga boleh ditetapkan sebagai Wali Kelas.
  // Role dasar tidak diubah; penugasan tetap disimpan terpisah.
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = getSheetByNameSafe_(ss,'Users');
  if (!sh || sh.getLastRow() < 2) return [];
  var vals = sh.getDataRange().getDisplayValues(), h = vals[0], ci = _kmmiFindCol_(h,['Username']), cn = _kmmiFindCol_(h,['Nama','Name']), cr = _kmmiFindCol_(h,['Role']), cs = _kmmiFindCol_(h,['Status']);
  if (ci < 0 || cr < 0) return [];
  return vals.slice(1).filter(function(r){
    var role = _kmmiNorm_(r[cr]);
    var active = !r[cs] || _kmmiNorm_(r[cs]) === 'aktif';
    return active && (role === 'guru' || role === 'wali kelas' || role === 'guru + wali kelas' || role === 'guru+wali kelas' || role === 'wali kelas + guru');
  }).map(function(r){return {username:String(r[ci]||'').trim(),nama:cn>=0?String(r[cn]||'').trim():'',role:cr>=0?String(r[cr]||'').trim():'',status:cs>=0?String(r[cs]||'Aktif').trim():'Aktif'};}).filter(function(u){return u.username;});
}
function getWaliKelasAssignments() {
  var sh = _kmmiEnsureWaliAssignmentSheet_(), d = _kmmiGetRows_(sh);
  var h=d.headers, rows=d.rows;
  var iu=_kmmiFindCol_(h,['Username']), inama=_kmmiFindCol_(h,['Nama Wali Kelas','Nama Wali','Nama']), ik=_kmmiFindCol_(h,['Kelas']), ist=_kmmiFindCol_(h,['Status']), iid=_kmmiFindCol_(h,['ID']);
  return rows.map(function(r){return {ID:iid>=0?r[iid]:'',Username:iu>=0?r[iu]:'','Nama Wali Kelas':inama>=0?r[inama]:'',Kelas:ik>=0?r[ik]:'',Status:ist>=0?r[ist]:'Aktif'};}).filter(function(x){return x.Username && x.Kelas;});
}
function saveWaliKelasAssignment(payload) {
  payload=payload||{};
  var username=String(payload.username||payload.Username||'').trim(), kelas=String(payload.kelas||payload.Kelas||'').trim();
  if(!username||!kelas)return {success:false,message:'Pilih akun Wali Kelas dan kelas.'};
  var users=getWaliKelasUsersForAdmin(), user=null;
  users.forEach(function(u){if(u.username.toLowerCase()===username.toLowerCase())user=u;});
  if(!user)return {success:false,message:'Akun tidak ditemukan atau role bukan Guru/Wali Kelas.'};
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=_kmmiEnsureWaliAssignmentSheet_(), d=_kmmiGetRows_(sh), h=d.headers;
  var iu=_kmmiFindCol_(h,['Username']), ik=_kmmiFindCol_(h,['Kelas']), idc=_kmmiFindCol_(h,['ID']), inama=_kmmiFindCol_(h,['Nama Wali Kelas']), ist=_kmmiFindCol_(h,['Status']);
  var santri=getSantriDirect(), classExists=santri.some(function(s){return String(s.Kelas||'').trim()===kelas;});
  if(!classExists)return {success:false,message:'Kelas tersebut belum ada pada data Santri.'};
  var target=-1;
  for(var i=0;i<d.rows.length;i++)if(String(d.rows[i][iu]||'').trim().toLowerCase()===username.toLowerCase()){target=i+2;break;}
  var out=target>=2?sh.getRange(target,1,1,sh.getLastColumn()).getValues()[0]:new Array(sh.getLastColumn()).fill('');
  out[idc]='WALI-'+Utilities.getUuid().slice(0,8).toUpperCase();out[iu]=username;out[inama]=user.nama;out[ik]=kelas;out[ist]='Aktif';
  if(target>=2)sh.getRange(target,1,1,out.length).setValues([out]);else sh.appendRow(out);
  try{touchDataVersion_();}catch(e){}
  return {success:true,message:'Penetapan kelas Wali Kelas berhasil disimpan.'};
}
function deleteWaliKelasAssignment(username) {
  username = String(username || '').trim();
  if (!username) return {success:false, message:'Username wajib diisi.'};
  var sh = _kmmiEnsureWaliAssignmentSheet_();
  var d = _kmmiGetRows_(sh), h = d.headers;
  var iu = _kmmiFindCol_(h, ['Username']);
  if (iu < 0) return {success:false, message:'Kolom Username tidak ditemukan.'};
  var target = -1;
  for (var i = 0; i < d.rows.length; i++) {
    if (String(d.rows[i][iu] || '').trim().toLowerCase() === username.toLowerCase()) { target = i + 2; break; }
  }
  if (target < 2) return {success:false, message:'Penetapan akun tersebut tidak ditemukan.'};
  sh.deleteRow(target);
  try { touchDataVersion_(); } catch (e) {}
  return {success:true, message:'Penetapan Wali Kelas berhasil dihapus. Akun login dan data santri tetap aman.'};
}
function getWaliKelasView(username) {
  var user=String(username||'').trim();
  if(!user)return {success:false,message:'Username akun tidak ditemukan.'};
  // Akses Wali Kelas ditentukan oleh penugasan aktif, bukan role dasar akun.
  // Ini memungkinkan satu akun Guru yang ditetapkan Admin tetap menggunakan login yang sama.
  var assigns=getWaliKelasAssignments(), assignment=null;
  assigns.forEach(function(a){if(String(a.Username||'').trim().toLowerCase()===user.toLowerCase()&&_kmmiNorm_(a.Status||'Aktif')==='aktif')assignment=a;});
  if(!assignment)return {success:false,message:'Kelas Wali belum ditetapkan oleh Admin.'};
  var kelas=String(assignment.Kelas||'').trim();
  var santri=getSantriDirect().filter(function(s){return String(s.Kelas||'').trim()===kelas;});
  var attSheet=_kmmiEnsureAttendanceSheet_(), ad=_kmmiGetRows_(attSheet);
  var absensi=ad.rows.map(function(r){return _kmmiAttendanceObject_(r,ad.headers);}).filter(function(a){return String(a.kelas||'').trim()===kelas;});
  var nilai=getNilaiDirectV1().filter(function(n){return String(n.Kelas||'').trim()===kelas;}).map(function(n){return {
    nis:n.NIS||'',NIS:n.NIS||'', 'nama santri':n['Nama Santri']||'', 'Nama Santri':n['Nama Santri']||'',
    'mata pelajaran':n['Mata Pelajaran']||'', 'Mata Pelajaran':n['Mata Pelajaran']||'',
    'nilai akhir':n['Nilai Akhir']||'', 'Nilai Akhir':n['Nilai Akhir']||'',
    'tanggal update':n['Tanggal Update']||'', 'Tanggal Update':n['Tanggal Update']||''
  };});
  return {success:true,kelas:kelas,santri:santri,absensi:absensi,nilai:nilai,nama:user};
}
/**
 * Wali Kelas hanya boleh memperbaiki nama santri yang masih terdaftar
 * pada kelas penugasannya. NIS dan kelas tidak dapat diubah dari fitur ini.
 */
function updateWaliSantriName(username, nis, namaBaru) {
  var user = String(username || '').trim();
  var keyNis = String(nis || '').trim();
  var name = String(namaBaru || '').trim().replace(/\s+/g, ' ');
  if (!user || !keyNis) return { success: false, message: 'Akun dan NIS santri wajib diisi.' };
  if (!name) return { success: false, message: 'Nama santri tidak boleh kosong.' };
  if (name.length > 150) return { success: false, message: 'Nama terlalu panjang (maksimal 150 karakter).' };
  var view = getWaliKelasView(user);
  if (!view || !view.success) return { success: false, message: (view && view.message) || 'Akses Wali Kelas tidak valid.' };
  var kelas = String(view.kelas || '').trim();
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = getSheetByNameSafe_(SpreadsheetApp.getActiveSpreadsheet(), SHEET_NAMES.SANTRI);
    if (!sh || sh.getLastRow() < 2) return { success: false, message: 'Data Santri tidak ditemukan.' };
    var vals = sh.getRange(2, 1, sh.getLastRow() - 1, Math.max(5, sh.getLastColumn())).getDisplayValues();
    for (var i = 0; i < vals.length; i++) {
      var rowNis = String(vals[i][1] || '').trim();
      var rowClass = String(vals[i][3] || '').trim();
      if (rowNis === keyNis && rowClass.toLowerCase() === kelas.toLowerCase()) {
        sh.getRange(i + 2, 3).setValue(name);
        SpreadsheetApp.flush();
        touchDataVersion_();
        return { success: true, message: 'Nama santri berhasil diperbarui.' };
      }
    }
    return { success: false, message: 'Santri tidak ditemukan di kelas wali yang ditugaskan.' };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Menghapus hanya baris master Santri pada kelas yang ditugaskan kepada Wali.
 * Riwayat nilai dan absensi tidak disentuh.
 */
function deleteWaliSantri(username, nis) {
  var user = String(username || '').trim();
  var keyNis = String(nis || '').trim();
  if (!user || !keyNis) return { success: false, message: 'Akun dan NIS santri wajib diisi.' };
  var view = getWaliKelasView(user);
  if (!view || !view.success) return { success: false, message: (view && view.message) || 'Akses Wali Kelas tidak valid.' };
  var kelas = String(view.kelas || '').trim();
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = getSheetByNameSafe_(SpreadsheetApp.getActiveSpreadsheet(), SHEET_NAMES.SANTRI);
    if (!sh || sh.getLastRow() < 2) return { success: false, message: 'Data Santri tidak ditemukan.' };
    var vals = sh.getRange(2, 1, sh.getLastRow() - 1, Math.max(5, sh.getLastColumn())).getDisplayValues();
    for (var i = vals.length - 1; i >= 0; i--) {
      var rowNis = String(vals[i][1] || '').trim();
      var rowClass = String(vals[i][3] || '').trim();
      if (rowNis === keyNis && rowClass.toLowerCase() === kelas.toLowerCase()) {
        sh.deleteRow(i + 2);
        SpreadsheetApp.flush();
        touchDataVersion_();
        return { success: true, message: 'Santri dihapus dari daftar aktif kelas ' + kelas + '. Riwayat nilai dan absensi lama tetap tersimpan.' };
      }
    }
    return { success: false, message: 'Santri tidak ditemukan di kelas wali yang ditugaskan.' };
  } finally {
    lock.releaseLock();
  }
}

function saveWaliKelasAbsensi(username, records) {
  var user=String(username||'').trim();
  var view=getWaliKelasView(user);
  if(!view || !view.success)return {success:false,message:(view&&view.message)||'Akses Wali Kelas tidak valid.'};
  var allowed=String(view.kelas||'').trim();
  var filtered=(records||[]).filter(function(r){return String(r.kelas||r.Kelas||'').trim()===allowed;});
  if(!filtered.length)return {success:false,message:'Tidak ada data absensi untuk kelas wali ini.'};
  return _kmmiWriteAttendance_(filtered,'',user,user,'');
}



/* ================================================================
   ABSENSI: pembacaan status dan data tersimpan
   Dibuat selaras dengan kunci penyimpanan: akun guru + kelas + tanggal.
   Tidak membuat baris baru saat membaca status.
   ================================================================ */
function _kmmiAttendanceOwnerMatches_(row, headers, idGuru, username) {
  var ru = _kmmiFindCol_(headers, ['Username', 'Akun']);
  var ri = _kmmiFindCol_(headers, ['ID Guru', 'ID_Guru']);
  var user = String(username || '').trim();
  var id = String(idGuru || '').trim();
  if (user && ru >= 0) return String(row[ru] || '').trim().toLowerCase() === user.toLowerCase();
  if (id && ri >= 0) return String(row[ri] || '').trim() === id;
  return false;
}
function getAbsensiSantriKelasTugas(payload) {
  payload = payload || {};
  var tanggal = String(payload.tanggal || '').trim().slice(0, 10);
  var kelas = String(payload.kelas || '').trim();
  var idGuru = String(payload.idGuru || payload['ID Guru'] || '').trim();
  var username = String(payload.username || '').trim();
  if (!tanggal || !kelas) return {success:false, rows:[], message:'Tanggal dan kelas wajib diisi.'};
  var sh = _kmmiEnsureAttendanceSheet_();
  var data = _kmmiGetRows_(sh);
  // Filter berdasarkan baris mentah supaya pemeriksaan pemilik memakai kolom asli.
  var rows = data.rows.filter(function(row) {
    var dateIx = _kmmiFindCol_(data.headers, ['Tanggal', 'Tanggal Absensi', 'Date']);
    var classIx = _kmmiFindCol_(data.headers, ['Kelas']);
    return dateIx >= 0 && classIx >= 0 &&
      String(row[dateIx] || '').trim().slice(0, 10) === tanggal &&
      String(row[classIx] || '').trim() === kelas &&
      _kmmiAttendanceOwnerMatches_(row, data.headers, idGuru, username);
  }).map(function(row) { return _kmmiAttendanceObject_(row, data.headers); });
  return {success:true, rows:rows};
}
function getAbsensiStatusTugasGuru(payload) {
  payload = payload || {};
  var tanggal = String(payload.tanggal || '').trim().slice(0, 10);
  var kelasList = Array.isArray(payload.kelasList) ? payload.kelasList.map(function(k) { return String(k || '').trim(); }) : [];
  var idGuru = String(payload.idGuru || payload['ID Guru'] || '').trim();
  var username = String(payload.username || '').trim();
  var statusByClass = {};
  kelasList.forEach(function(k) { if (k) statusByClass[k] = {sudah:false}; });
  if (!tanggal || !kelasList.length) return {success:true, statusByClass:statusByClass};
  var sh = _kmmiEnsureAttendanceSheet_();
  var data = _kmmiGetRows_(sh);
  var dateIx = _kmmiFindCol_(data.headers, ['Tanggal', 'Tanggal Absensi', 'Date']);
  var classIx = _kmmiFindCol_(data.headers, ['Kelas']);
  if (dateIx < 0 || classIx < 0) return {success:true, statusByClass:statusByClass};
  data.rows.forEach(function(row) {
    var kelas = String(row[classIx] || '').trim();
    if (statusByClass[kelas] && String(row[dateIx] || '').trim().slice(0, 10) === tanggal &&
        _kmmiAttendanceOwnerMatches_(row, data.headers, idGuru, username)) {
      statusByClass[kelas] = {sudah:true};
    }
  });
  return {success:true, statusByClass:statusByClass};
}


/** Rekap absensi santri untuk tampilan harian dan bulanan.
 * Tidak mengubah data. Guru hanya menerima absensi miliknya; Admin/Staff dapat melihat semua.
 */
function getAbsensiRekapSantri(payload) {
  payload = payload || {};
  var username = String(payload.username || '').trim();
  var idGuru = String(payload.idGuru || payload['ID Guru'] || '').trim();
  var from = String(payload.tanggalAwal || payload.from || '').trim().slice(0, 10);
  var to = String(payload.tanggalAkhir || payload.to || '').trim().slice(0, 10);
  var wantedClass = String(payload.kelas || '').trim();
  var wantedMapel = String(payload.mapel || '').trim();
  var norm = function(v) {
    return String(v == null ? '' : v).trim().toLowerCase()
      .replace(/[_-]+/g, ' ').replace(/\s+/g, ' ');
  };

  // Hak akses dan kelas Wali dibaca ulang dari Sheet Users, bukan dipercaya
  // hanya dari payload browser. Ini menjaga Wali tetap terbatas pada kelasnya.
  var role = String(payload.role || payload.Role || '').trim();
  var kelasWali = String(payload.kelasWali || payload['Kelas Wali'] || '').trim();
  var namaAkun = String(payload.nama || payload.namaPenginput || '').trim();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var usersSheet = getSheetByNameSafe_(ss, SHEET_NAMES.USERS || 'Users');
  if (usersSheet && usersSheet.getLastRow() >= 2 && username) {
    var userValues = usersSheet.getDataRange().getDisplayValues();
    var userHeaders = (userValues[0] || []).map(norm);
    var findUserCol = function(names) {
      for (var i = 0; i < names.length; i++) {
        var ix = userHeaders.indexOf(norm(names[i]));
        if (ix >= 0) return ix;
      }
      return -1;
    };
    var userIx = findUserCol(['Username', 'User Name', 'Akun']);
    var roleIx = findUserCol(['Role', 'Peran']);
    var classWaliIx = findUserCol(['Kelas Wali', 'Kelas_Wali', 'KelasWali', 'Kelas Tanggung Jawab']);
    var nameIx = findUserCol(['Nama', 'Nama Lengkap', 'Name']);
    var idGuruIx = findUserCol(['ID Guru', 'ID_Guru', 'IDGuru']);
    for (var ui = 1; ui < userValues.length; ui++) {
      if (userIx >= 0 && norm(userValues[ui][userIx]) === norm(username)) {
        if (roleIx >= 0 && String(userValues[ui][roleIx] || '').trim()) role = String(userValues[ui][roleIx]).trim();
        if (classWaliIx >= 0 && String(userValues[ui][classWaliIx] || '').trim()) kelasWali = String(userValues[ui][classWaliIx]).trim();
        if (nameIx >= 0 && String(userValues[ui][nameIx] || '').trim()) namaAkun = String(userValues[ui][nameIx]).trim();
        if (!idGuru && idGuruIx >= 0) idGuru = String(userValues[ui][idGuruIx] || '').trim();
        break;
      }
    }
  }

  // Sumber kebenaran penugasan Wali: sheet Penugasan Wali Kelas.
  // Akun gabungan kadang menyimpan role dasar 'Guru' di Users; jangan sampai
  // kelas lain terbuka hanya karena role dasar itu tidak memuat kata Wali.
  var waliAssignment = null;
  if (username) {
    try {
      getWaliKelasAssignments().forEach(function(a) {
        if (norm(a.Username) === norm(username) && norm(a.Status || 'Aktif') === 'aktif') {
          waliAssignment = a;
        }
      });
    } catch (assignmentError) {}
  }
  if (waliAssignment && String(waliAssignment.Kelas || '').trim()) {
    kelasWali = String(waliAssignment.Kelas).trim();
    if (norm(role).indexOf('wali') < 0) role = role ? role + ' + Wali Kelas' : 'Wali Kelas';
  }

  var roleKey = norm(role);
  var isWali = roleKey.indexOf('wali') >= 0;
  var isPrivileged = ['admin', 'administrator', 'administrator utama', 'staff', 'staf'].indexOf(roleKey) >= 0;
  // Mendukung akun gabungan Guru+Wali tanpa memperlebar akses rekap kelas lain.
  if (isWali && !kelasWali) {
    var roleClass = String(role || '').match(/wali\s*kelas\s*[-:|]\s*(.+)$/i);
    if (roleClass && roleClass[1]) kelasWali = String(roleClass[1]).trim();
    if (!kelasWali) kelasWali = String(payload.kelasWali || payload['Kelas Wali'] || '').trim();
  }
  if (isWali) {
    if (!kelasWali) {
      return {success:false, rows:[], tugas:[], message:'Kelas Wali belum terdeteksi pada akun. Periksa kolom Kelas Wali di Sheet Users.'};
    }
    // Wali tidak dapat mengganti kelas melalui parameter browser.
    if (wantedClass && norm(wantedClass) !== norm(kelasWali)) {
      return {success:false, rows:[], tugas:[], message:'Akses rekap dibatasi pada kelas tanggung jawab Wali: ' + kelasWali + '.'};
    }
    wantedClass = kelasWali;
  }

  var sh = _kmmiEnsureAttendanceSheet_();
  var data = _kmmiGetRows_(sh);
  var dateIx = _kmmiFindCol_(data.headers, ['Tanggal', 'Tanggal Absensi', 'Date']);
  var classIx = _kmmiFindCol_(data.headers, ['Kelas']);
  var mapelIx = _kmmiFindCol_(data.headers, ['Mata Pelajaran', 'Mapel']);
  var usernameIx = _kmmiFindCol_(data.headers, ['Username', 'Akun']);
  var idGuruIx = _kmmiFindCol_(data.headers, ['ID Guru', 'ID_Guru']);

  var allTasks = getTugasGuru().filter(function(t) {
    var status = norm(t.Status || 'Aktif');
    return !status || status === 'aktif';
  });
  var tasks;
  if (isPrivileged) {
    tasks = allTasks.slice();
  } else if (isWali) {
    tasks = allTasks.filter(function(t) { return norm(t.Kelas) === norm(kelasWali); });
  } else {
    tasks = allTasks.filter(function(t) {
      return idGuru && String(t['ID Guru'] || '').trim() === idGuru;
    });
  }

  var subjectByOwnerClass = {};
  tasks.forEach(function(t) {
    var c = String(t.Kelas || '').trim();
    var m = String(t['Mata Pelajaran'] || '').trim();
    var owner = String(t['ID Guru'] || '').trim();
    if (!c || !m || !owner) return;
    var key = owner + '|' + norm(c);
    if (!subjectByOwnerClass[key]) subjectByOwnerClass[key] = [];
    if (subjectByOwnerClass[key].indexOf(m) < 0) subjectByOwnerClass[key].push(m);
  });

  var rows = data.rows.filter(function(row) {
    var d = dateIx >= 0 ? String(row[dateIx] || '').trim().slice(0, 10) : '';
    var c = classIx >= 0 ? String(row[classIx] || '').trim() : '';
    if (!d || (from && d < from) || (to && d > to)) return false;
    if (wantedClass && norm(c) !== norm(wantedClass)) return false;
    if (isWali) return true; // seluruh guru, tetapi hanya kelas Wali yang sah
    if (isPrivileged) return true;
    return _kmmiAttendanceOwnerMatches_(row, data.headers, idGuru, username);
  }).map(function(row) {
    var o = _kmmiAttendanceObject_(row, data.headers);
    var m = mapelIx >= 0 ? String(row[mapelIx] || '').trim() : '';
    var rowUsername = usernameIx >= 0 ? String(row[usernameIx] || '').trim() : '';
    var rowIdGuru = idGuruIx >= 0 ? String(row[idGuruIx] || '').trim() : '';
    var ownWaliEntry = isWali && username && norm(rowUsername) === norm(username);

    // Absensi yang dimasukkan Wali merupakan kategori tersendiri, bukan
    // disalin ke setiap mata pelajaran dan dihitung berkali-kali.
    if (ownWaliEntry && (!m || norm(m) === 'absensi santri' || norm(m) === 'absensi wali')) {
      m = 'Absensi Wali';
    }
    if (!m) {
      var ownerSubjects = subjectByOwnerClass[rowIdGuru + '|' + norm(o.kelas || o.Kelas)] || [];
      if (ownerSubjects.length === 1) m = ownerSubjects[0];
      else m = 'Mapel Belum Ditentukan';
    }
    o.mapel = m || 'Mapel Belum Ditentukan';
    o['Mata Pelajaran'] = o.mapel;
    return o;
  }).filter(function(o) {
    return !wantedMapel || norm(o.mapel) === norm(wantedMapel);
  });

  // Tambahkan kategori yang memang muncul pada data absensi, agar dapat
  // dipilih di Rekap Bulanan walaupun bukan tugas mengajar reguler.
  var taskKeys = {};
  tasks.forEach(function(t) {
    var key = norm(t.Kelas) + '|' + norm(t['Mata Pelajaran']);
    taskKeys[key] = true;
  });
  rows.forEach(function(r) {
    var kelas = String(r.kelas || r.Kelas || '').trim();
    var mapel = String(r.mapel || '').trim();
    var key = norm(kelas) + '|' + norm(mapel);
    if (kelas && mapel && !taskKeys[key]) {
      tasks.push({Kelas:kelas, 'Mata Pelajaran':mapel, Status:'Aktif'});
      taskKeys[key] = true;
    }
  });
  if (isWali && !taskKeys[norm(kelasWali) + '|absensi wali']) {
    tasks.push({Kelas:kelasWali, 'Mata Pelajaran':'Absensi Wali', Status:'Aktif'});
  }

  return {
    success: true,
    rows: rows,
    tugas: tasks.map(function(t) {
      return {
        mapel: String(t['Mata Pelajaran'] || '').trim(),
        kelas: String(t.Kelas || '').trim(),
        status: String(t.Status || 'Aktif').trim(),
        namaGuru: String(t['Nama Guru'] || '').trim()
      };
    }),
    kelasWali: isWali ? kelasWali : ''
  };
}

/**
 * Tambahan untuk menu Lihat Juara Mapel pada akun Wali Kelas.
 * Tempel di bagian paling bawah Code.gs yang sudah berisi helper V32:
 * juaraMapelSheet_() dan juaraMapelRowObject_().
 * Mengembalikan hanya pengajuan berstatus Disetujui dari kelas yang diminta.
 */
function getJuaraMapelForWali(payload) {
  try {
    payload = payload || {};
    var role = String(payload.role || '').trim().toUpperCase().replace(/[_-]+/g, ' ').replace(/\s+/g, ' ');
    if (role !== 'WALI KELAS') {
      return { success: false, rows: [], message: 'Akses hanya untuk tampilan Wali Kelas.' };
    }
    var kelas = String(payload.classId || payload.kelas || '').trim();
    if (!kelas) {
      return { success: false, rows: [], message: 'Kelas Wali belum terdeteksi. Periksa penetapan kelas Wali terlebih dahulu.' };
    }
    var sh = juaraMapelSheet_();
    var last = sh.getLastRow();
    if (last < 2) return { success: true, rows: [] };
    var values = sh.getRange(2, 1, last - 1, 14).getDisplayValues();
    var norm = function(v) { return String(v == null ? '' : v).trim().toLowerCase().replace(/\s+/g, ' '); };
    var rows = values
      .filter(function(r) { return String(r[0] || '').trim(); })
      .map(juaraMapelRowObject_)
      .filter(function(r) {
        return norm(r.classId) === norm(kelas) && norm(r.status) === 'disetujui';
      })
      .sort(function(a, b) {
        return String(a.semester).localeCompare(String(b.semester)) ||
          String(a.subjectName).localeCompare(String(b.subjectName));
      });
    return { success: true, rows: rows };
  } catch (err) {
    return { success: false, rows: [], message: err && err.message ? err.message : String(err) };
  }
}


// ============================================================
// TAMBAHAN JUARA MAPEL V32 — Guru mengajukan, Admin menyetujui,
// Wali Kelas hanya melihat hasil yang sudah disetujui.
// Ditambahkan tanpa mengganti fungsi backend yang sudah ada.
// ============================================================

/*
 * V32 — JUARA MATA PELAJARAN PER SEMESTER
 * Tempelkan seluruh isi file ini di BAGIAN PALING BAWAH Code.gs yang aktif.
 * Tidak mengganti fungsi absensi, rekap, NIS, logo, pengaturan lembaga, atau Informasi.
 * Data satu usulan disimpan per kombinasi Semester + Kelas + Mata Pelajaran,
 * sehingga pengiriman ulang memperbarui baris lama, bukan menambah duplikat.
 */

function juaraMapelSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Spreadsheet utama tidak ditemukan.');
  var sh = ss.getSheetByName('Juara_Mapel');
  var headers = ['record_id','semester','class_id','subject_name','schedule_id','teacher_id','teacher_name','student_nis','student_name','status','created_at','updated_at','reviewed_by','reviewed_at'];
  if (!sh) sh = ss.insertSheet('Juara_Mapel');
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#D1FAE5');
  } else {
    var current = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), headers.length)).getDisplayValues()[0];
    if (String(current[0] || '').trim().toLowerCase() !== 'record_id') {
      throw new Error('Header sheet Juara_Mapel tidak sesuai. Jangan mengubah header sheet ini secara manual.');
    }
  }
  return sh;
}

function juaraMapelNorm_(v) { return String(v == null ? '' : v).trim().toLowerCase().replace(/\s+/g, ' '); }
function juaraMapelRole_(p) { return String(p && p.role || '').trim().toUpperCase().replace(/[_-]+/g, ' ').replace(/\s*\+\s*/g, '+').replace(/\s+/g, ' '); }
function juaraMapelIsGuruRole_(role) { return role === 'GURU' || role === 'GURU+WALI' || role === 'GURU WALI' || role === 'GURU DAN WALI' || role === 'GURU & WALI'; }
function juaraMapelIdentity_(p) { return String(p && (p.teacherId || p.userId || p.username || p.teacherName || p.name) || '').trim(); }
function juaraMapelRowObject_(r) {
  return {
    recordId: String(r[0] || ''), semester: String(r[1] || ''), classId: String(r[2] || ''),
    subjectName: String(r[3] || ''), scheduleId: String(r[4] || ''), teacherId: String(r[5] || ''),
    teacherName: String(r[6] || ''), studentNis: String(r[7] || ''), studentName: String(r[8] || ''),
    status: String(r[9] || 'Menunggu Persetujuan'), createdAt: String(r[10] || ''), updatedAt: String(r[11] || ''),
    reviewedBy: String(r[12] || ''), reviewedAt: String(r[13] || '')
  };
}

function getJuaraMapelData(payload) {
  try {
    payload = payload || {};
    var sh = juaraMapelSheet_();
    var last = sh.getLastRow();
    if (last < 2) return { success: true, rows: [] };
    var values = sh.getRange(2, 1, last - 1, 14).getDisplayValues();
    var admin = juaraMapelRole_(payload) === 'ADMIN';
    var identity = juaraMapelNorm_(juaraMapelIdentity_(payload));
    var rows = values.filter(function(r) { return String(r[0] || '').trim(); }).map(juaraMapelRowObject_);
    if (!admin) {
      rows = rows.filter(function(r) {
        return juaraMapelNorm_(r.teacherId) === identity ||
          (!r.teacherId && juaraMapelNorm_(r.teacherName) === juaraMapelNorm_(payload.name || payload.username));
      });
    } else {
      rows = rows.filter(function(r) { return juaraMapelNorm_(r.status) !== 'dibatalkan'; });
    }
    rows.sort(function(a,b) { return String(b.updatedAt || '').localeCompare(String(a.updatedAt || '')); });
    return { success: true, rows: rows };
  } catch (err) {
    return { success: false, rows: [], message: err && err.message ? err.message : String(err) };
  }
}

function saveJuaraMapel(payload) {
  var lock = LockService.getScriptLock();
  try {
    payload = payload || {};
    var role = juaraMapelRole_(payload);
    if (!juaraMapelIsGuruRole_(role) && role !== 'ADMIN') return { success: false, message: 'Akun ini tidak memiliki akses untuk mengirim juara.' };
    var semester = String(payload.semester || '').trim();
    var kelas = String(payload.classId || '').trim();
    var mapel = String(payload.subjectName || '').trim();
    var teacherId = String(payload.teacherId || payload.userId || payload.username || '').trim();
    var teacherName = String(payload.teacherName || payload.name || payload.username || '').trim();
    var nis = String(payload.studentNis || '').trim();
    var nama = String(payload.studentName || '').trim();
    if (!semester || !kelas || !mapel || !teacherName || !nis || !nama) {
      return { success: false, message: 'Semester, kelas, mata pelajaran, guru, NIS, dan nama santri wajib lengkap.' };
    }
    if (!lock.tryLock(10000)) return { success: false, message: 'Server sedang memproses pengajuan lain. Silakan coba lagi.' };
    var sh = juaraMapelSheet_();
    var last = sh.getLastRow();
    var matchRow = 0;
    var values = last >= 2 ? sh.getRange(2, 1, last - 1, 14).getDisplayValues() : [];
    for (var i = 0; i < values.length; i++) {
      var r = values[i];
      if (juaraMapelNorm_(r[1]) === juaraMapelNorm_(semester) &&
          juaraMapelNorm_(r[2]) === juaraMapelNorm_(kelas) &&
          juaraMapelNorm_(r[3]) === juaraMapelNorm_(mapel)) {
        matchRow = i + 2;
        var existingTeacherId = String(r[5] || '').trim();
        var sameTeacher = juaraMapelNorm_(existingTeacherId) === juaraMapelNorm_(teacherId) ||
          (!existingTeacherId && juaraMapelNorm_(r[6]) === juaraMapelNorm_(teacherName));
        if (!sameTeacher && role !== 'ADMIN') {
          return { success: false, message: 'Juara untuk kelas dan mata pelajaran ini sudah diajukan oleh guru lain. Hubungi Admin jika perlu diperbaiki.' };
        }
        break;
      }
    }
    var now = new Date().toISOString();
    if (matchRow) {
      var old = sh.getRange(matchRow, 1, 1, 14).getValues()[0];
      sh.getRange(matchRow, 1, 1, 14).setValues([[
        old[0], semester, kelas, mapel, String(payload.scheduleId || ''), teacherId, teacherName,
        nis, nama, 'Menunggu Persetujuan', old[10] || now, now, '', ''
      ]]);
      if (typeof touchDataVersion_ === 'function') touchDataVersion_();
      return { success: true, message: 'Pilihan juara diperbarui. Menunggu persetujuan Admin.', updated: true };
    }
    var id = 'JMP_' + Utilities.getUuid();
    sh.appendRow([id, semester, kelas, mapel, String(payload.scheduleId || ''), teacherId, teacherName, nis, nama, 'Menunggu Persetujuan', now, now, '', '']);
    if (typeof touchDataVersion_ === 'function') touchDataVersion_();
    return { success: true, message: 'Juara berhasil dikirim. Menunggu persetujuan Admin.', updated: false };
  } catch (err) {
    return { success: false, message: err && err.message ? err.message : String(err) };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function reviewJuaraMapel(payload) {
  try {
    payload = payload || {};
    if (juaraMapelRole_(payload) !== 'ADMIN') return { success: false, message: 'Hanya Admin yang dapat menyetujui atau menolak juara.' };
    var recordId = String(payload.recordId || '').trim();
    var status = String(payload.status || '').trim();
    var reviewer = String(payload.reviewedBy || 'Administrator').trim();
    if (!recordId || (status !== 'Disetujui' && status !== 'Ditolak')) return { success: false, message: 'Data persetujuan tidak valid.' };
    var sh = juaraMapelSheet_();
    var last = sh.getLastRow();
    if (last < 2) return { success: false, message: 'Belum ada pengajuan juara.' };
    var ids = sh.getRange(2, 1, last - 1, 1).getDisplayValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === recordId) {
        sh.getRange(i + 2, 10, 1, 1).setValue(status);
        sh.getRange(i + 2, 13, 1, 2).setValues([[reviewer, new Date().toISOString()]]);
        if (typeof touchDataVersion_ === 'function') touchDataVersion_();
        return { success: true, message: 'Status juara diperbarui menjadi ' + status + '.' };
      }
    }
    return { success: false, message: 'Pengajuan juara tidak ditemukan.' };
  } catch (err) {
    return { success: false, message: err && err.message ? err.message : String(err) };
  }
}


/**
 * Membatalkan pengajuan Juara Mapel milik Guru yang sedang login.
 * Baris dipertahankan sebagai jejak audit; data tidak dihapus permanen.
 * Pengajuan yang sudah disetujui Admin tidak dapat dibatalkan oleh Guru.
 */
function deleteJuaraMapel(payload) {
  var lock = LockService.getScriptLock();
  try {
    payload = payload || {};
    if (!juaraMapelIsGuruRole_(juaraMapelRole_(payload))) {
      return { success: false, message: 'Hanya akun Guru atau Guru+Wali yang dapat membatalkan pengajuannya sendiri.' };
    }
    var recordId = String(payload.recordId || '').trim();
    var identity = juaraMapelNorm_(juaraMapelIdentity_(payload));
    if (!recordId || !identity) {
      return { success: false, message: 'Identitas Guru atau ID pengajuan tidak lengkap.' };
    }
    if (!lock.tryLock(10000)) {
      return { success: false, message: 'Server sedang memproses pengajuan lain. Silakan coba lagi.' };
    }
    var sh = juaraMapelSheet_();
    var last = sh.getLastRow();
    if (last < 2) return { success: false, message: 'Pengajuan tidak ditemukan.' };
    var values = sh.getRange(2, 1, last - 1, 14).getDisplayValues();
    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      if (String(row[0] || '').trim() !== recordId) continue;
      var ownerId = juaraMapelNorm_(row[5]);
      var ownerName = juaraMapelNorm_(row[6]);
      var sameOwner = ownerId ? ownerId === identity :
        ownerName === juaraMapelNorm_(payload.teacherName || payload.name || payload.username);
      if (!sameOwner) return { success: false, message: 'Akses ditolak. Guru hanya dapat membatalkan pengajuannya sendiri.' };
      var status = juaraMapelNorm_(row[9]);
      if (status === 'disetujui') return { success: false, message: 'Pengajuan sudah disetujui Admin dan tidak dapat dibatalkan oleh Guru.' };
      if (status === 'dibatalkan') return { success: false, message: 'Pengajuan ini sudah dibatalkan sebelumnya.' };
      sh.getRange(i + 2, 10).setValue('Dibatalkan');
      sh.getRange(i + 2, 12).setValue(new Date().toISOString());
      sh.getRange(i + 2, 13, 1, 2).clearContent();
      if (typeof touchDataVersion_ === 'function') touchDataVersion_();
      return { success: true, message: 'Pengajuan dibatalkan. Riwayat tetap disimpan.' };
    }
    return { success: false, message: 'Pengajuan Juara Mapel tidak ditemukan.' };
  } catch (err) {
    return { success: false, message: err && err.message ? err.message : String(err) };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function getJuaraMapelForWali(payload) {
  try {
    payload = payload || {};
    var role = String(payload.role || '').trim().toUpperCase().replace(/[_-]+/g, ' ').replace(/\s+/g, ' ');
    if (role !== 'WALI KELAS') {
      return { success: false, rows: [], message: 'Akses hanya untuk tampilan Wali Kelas.' };
    }
    var kelas = String(payload.classId || payload.kelas || '').trim();
    if (!kelas) {
      return { success: false, rows: [], message: 'Kelas Wali belum terdeteksi. Periksa penetapan kelas Wali terlebih dahulu.' };
    }
    var sh = juaraMapelSheet_();
    var last = sh.getLastRow();
    if (last < 2) return { success: true, rows: [] };
    var values = sh.getRange(2, 1, last - 1, 14).getDisplayValues();
    var norm = function(v) { return String(v == null ? '' : v).trim().toLowerCase().replace(/\s+/g, ' '); };
    var rows = values
      .filter(function(r) { return String(r[0] || '').trim(); })
      .map(juaraMapelRowObject_)
      .filter(function(r) {
        return norm(r.classId) === norm(kelas) && norm(r.status) === 'disetujui';
      })
      .sort(function(a, b) {
        return String(a.semester).localeCompare(String(b.semester)) ||
          String(a.subjectName).localeCompare(String(b.subjectName));
      });
    return { success: true, rows: rows };
  } catch (err) {
    return { success: false, rows: [], message: err && err.message ? err.message : String(err) };
  }
}


/**
 * Jalankan sekali secara manual dari editor Apps Script.
 * Membuat sheet Juara_Mapel dan headernya tanpa menghapus data lama.
 */
function setupJuaraMapelSheetNow() {
  try {
    var sh = juaraMapelSheet_();
    SpreadsheetApp.flush();
    return {
      success: true,
      sheetName: sh.getName(),
      lastRow: sh.getLastRow(),
      message: 'Sheet Juara_Mapel siap digunakan.'
    };
  } catch (err) {
    throw new Error('Gagal menyiapkan sheet Juara_Mapel: ' +
      (err && err.message ? err.message : String(err)));
  }
}

/**
 * ADD-ON UJI COBA REKAP ABSENSI MELEBAR
 * Tempelkan di BAGIAN PALING BAWAH Code.gs aktif.
 * Tidak mengganti fungsi simpan absensi dan tidak mengubah sheet ABSENSI.
 * Jalankan manual: buatRekapAbsensiLebar()
 */
function buatRekapAbsensiLebar() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sh=_kmmiWideSheet_();
  var headers=sh.getRange(1,1,1,Math.max(1,sh.getLastColumn())).getDisplayValues()[0];
  if (sh.getLastRow()<1) throw new Error('Sheet REKAP ABSENSI LEBAR belum siap.');
  sh.setFrozenRows(1);
  sh.getRange(1,1,1,headers.length).setBackground('#0f766e').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  if(sh.getLastRow()>1) sh.getRange(2,1,sh.getLastRow()-1,headers.length).setVerticalAlignment('middle').setWrap(true);
  sh.setColumnWidth(1,105); sh.setColumnWidth(2,95); sh.setColumnWidth(3,115); sh.setColumnWidth(4,230);
  for(var c=4;c<headers.length;c++) {
    if(String(headers[c]||'').indexOf('__META__')===0){try{sh.hideColumns(c+1);}catch(e){}}
    else sh.setColumnWidth(c+1,145);
  }
  return 'REKAP ABSENSI LEBAR sudah menjadi penyimpanan utama. Data tidak dibangun ulang atau dihapus.';
}

/**
 * MIGRASI SEKALI SAJA: salin data ABSENSI lama ke REKAP ABSENSI LEBAR.
 * Tidak menghapus sheet sumber. Jalankan manual, lalu bandingkan jumlah data sebelum menghapus sumber.
 */
function migrasiAbsensiLamaKeLebar() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var old=null;
  ['ABSENSI','Absensi Santri','Absensi','Kehadiran Santri'].some(function(n){old=getSheetByNameSafe_(ss,n);return !!old;});
  if(!old || old.getLastRow()<2) return {success:true,message:'Tidak ada data lama untuk dimigrasikan.',migrated:0};
  var d=_kmmiGetRows_(old), h=d.headers, rows=d.rows;
  var ix={nis:_kmmiFindCol_(h,['NIS','ID Santri']),nama:_kmmiFindCol_(h,['Nama Santri','Nama']),kelas:_kmmiFindCol_(h,['Kelas']),tanggal:_kmmiFindCol_(h,['Tanggal','Tanggal Absensi','Date']),status:_kmmiFindCol_(h,['Status','Kehadiran']),ket:_kmmiFindCol_(h,['Keterangan','Catatan']),guru:_kmmiFindCol_(h,['ID Guru']),user:_kmmiFindCol_(h,['Username','Akun']),penginput:_kmmiFindCol_(h,['Nama Penginput','Nama Guru','Nama Wali']),mapel:_kmmiFindCol_(h,['Mata Pelajaran','Mapel'])};
  ['nis','nama','kelas','tanggal','status'].forEach(function(k){if(ix[k]<0)throw new Error('Kolom wajib sumber lama tidak ditemukan: '+k);});
  var count=0;
  rows.forEach(function(row){
    var rec={nis:row[ix.nis],nama:row[ix.nama],kelas:row[ix.kelas],tanggal:String(row[ix.tanggal]||'').slice(0,10),status:row[ix.status],keterangan:ix.ket>=0?row[ix.ket]:'',idGuru:ix.guru>=0?row[ix.guru]:'',username:ix.user>=0?row[ix.user]:'',namaPenginput:ix.penginput>=0?row[ix.penginput]:'',mapel:ix.mapel>=0?row[ix.mapel]:'Pelajaran Wali Kelas'};
    if(!rec.mapel)rec.mapel='Pelajaran Wali Kelas';
    var result=_kmmiWriteAttendance_([rec],rec.idGuru,rec.username,rec.namaPenginput,rec.mapel);
    if(!result.success)throw new Error(result.message||'Gagal memigrasikan NIS '+rec.nis);
    count++;
  });
  SpreadsheetApp.flush();
  return {success:true,migrated:count,message:'Migrasi selesai. Sheet ABSENSI lama tetap dipertahankan untuk pemeriksaan.'};
}
                   
