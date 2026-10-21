// CONFIGURATION
const SPREADSHEET_ID = '1PkmW4hah19NQjfd__Sz2c3ZGONO4GWDuukVs4ViLGYE';

const URL_DATAFORM = `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/gviz/tq?tqx=out:json&sheet=DataForm`;
const URL_USERS = `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/gviz/tq?tqx=out:json&sheet=Users`;
const URL_ATTENDANCE = `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/gviz/tq?tqx=out:json&sheet=Attendance`;
const URL_SALARYSETTINGS = `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/gviz/tq?tqx=out:json&sheet=SalarySettings`;

// Global State
let rawPayrollData = [];
let processedPayrollSummary = [];

let usersMap = {};
let rawAttendanceData = [];
let processedAttendanceSummary = [];

let tutorSalarySettings = []; // Data tarif honor per tutor (khusus role Tutor)

// Master List Honor Kegiatan (Sesuai Gambar Acuan Fisik)
const daftarHonorKegiatanMaster = [
  { no: 1, kegiatan: "HONOR TRANSPORT", honor: 20000 },
  { no: 2, kegiatan: "HONOR KEGIATAN NON AKADEMIS", honor: 30000 },
  { no: 3, kegiatan: "HONOR KEGIATAN NON AKADEMIS (LAMA)", honor: 50000 },
  { no: 4, kegiatan: "HONOR BIMTEK", honor: 50000 },
  { no: 5, kegiatan: "HONOR PENGAWAS TKA SD", honor: 55000 },
  { no: 6, kegiatan: "HONOR PENGAWAS TKA SMP", honor: 65000 },
  { no: 7, kegiatan: "HONOR PENGAWAS TKA SMA", honor: 75000 },
  { no: 8, kegiatan: "HONOR LEMBUR WEEKEND", honor: 100000 },
  { no: 9, kegiatan: "HONOR PEMBUATAN SOAL UJIAN SD", honor: null },
  { no: 10, kegiatan: "HONOR PEMBUATAN SOAL UJIAN SMP", honor: null },
  { no: 11, kegiatan: "HONOR PEMBUATAN SOAL UJIAN SMA", honor: null },
  { no: 12, kegiatan: "HONOR PENGAWAS UJIAN", honor: null },
  { no: 13, kegiatan: "HONOR PEMBAGIAN RAPORT WALIKELAS", honor: 300000 },
  { no: 14, kegiatan: "HONOR WEEKLY REPORT (PERLEMBAR)", honor: 3750 }
];

// Format Angka ke Rupiah
function formatRupiah(number) {
  if (number === null || number === undefined || number === '' || isNaN(number) || number === 0) {
    return '-';
  }
  return new Intl.NumberFormat('id-ID', { style: 'currency', currency: 'IDR', maximumFractionDigits: 0 }).format(number);
}

// Mapping bulan Bahasa Indonesia
const indonesianMonths = {
  'januari': '01', 'jan': '01',
  'februari': '02', 'feb': '02',
  'maret': '03', 'mar': '03',
  'april': '04', 'apr': '04',
  'mei': '05',
  'juni': '06', 'jun': '06',
  'juli': '07', 'jul': '07',
  'agustus': '08', 'ags': '08', 'agu': '08', 'agt': '08',
  'september': '09', 'sep': '09', 'sept': '09',
  'oktober': '10', 'okt': '10',
  'november': '11', 'nov': '11', 'nopember': '11', 'nop': '11',
  'desember': '12', 'des': '12'
};

// Helper Null-Safe Cell Reader
function getCellValue(row, colIndex) {
  if (!row || !row.c || !row.c[colIndex]) return '';
  const cell = row.c[colIndex];
  return cell.v !== null && cell.v !== undefined ? cell.v : (cell.f || '');
}

function getCellFormatted(row, colIndex) {
  if (!row || !row.c || !row.c[colIndex]) return '';
  const cell = row.c[colIndex];
  return cell.f || (cell.v !== null && cell.v !== undefined ? String(cell.v) : '');
}

// Konversi Multi-Format Tanggal ke YYYY-MM-DD
function parseToIsoDate(rawVal, formattedVal) {
  const rawStr = String(rawVal || '').trim();
  const fmtStr = String(formattedVal || '').trim();

  if (rawStr.includes('Date(')) {
    const matches = rawStr.match(/\d+/g);
    if (matches && matches.length >= 3) {
      const y = matches[0];
      const m = String(Number(matches[1]) + 1).padStart(2, '0');
      const d = String(matches[2]).padStart(2, '0');
      return `${y}-${m}-${d}`;
    }
  }

  const target = fmtStr || rawStr;
  if (!target) return '';

  const matchIso = target.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if (matchIso) {
    return `${matchIso[1]}-${matchIso[2].padStart(2, '0')}-${matchIso[3].padStart(2, '0')}`;
  }

  const matchNumeric = target.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (matchNumeric) {
    return `${matchNumeric[3]}-${matchNumeric[2].padStart(2, '0')}-${matchNumeric[1].padStart(2, '0')}`;
  }

  const clean = target.replace(/^(senin|selasa|rabu|kamis|jumat|jum\'at|sabtu|minggu)[,\s]*/i, '').trim();
  const matchText = clean.match(/(\d{1,2})[\s\/\-]+([a-zA-Z]+)[\s\/\-]+(\d{4})/);
  if (matchText) {
    const day = matchText[1].padStart(2, '0');
    const monthStr = matchText[2].toLowerCase();
    const year = matchText[3];
    const month = indonesianMonths[monthStr] || '01';
    return `${year}-${month}-${day}`;
  }

  return target;
}

// Helper Fetch & Parse JSON GViz API
async function fetchGVizJson(url) {
  const response = await fetch(url);
  if (!response.ok) throw new Error(`HTTP ${response.status}`);
  const text = await response.text();
  const firstOpen = text.indexOf('{');
  const lastClose = text.lastIndexOf('}');
  if (firstOpen === -1 || lastClose === -1) throw new Error("Akses ditolak ke Spreadsheet");
  const jsonString = text.substring(firstOpen, lastClose + 1);
  return JSON.parse(jsonString);
}

// FETCH ALL DATA (Urutan: Users -> SalarySettings -> DataForm -> Attendance)
async function fetchAllData() {
  await fetchAttendanceAndUsers();
  await fetchSalarySettings();
  await fetchPayrollData();
}

// ========================================================
// Helper untuk Menghitung Tarif per Sesi Mengajar Tutor
// ========================================================
function calculateSessionHonor(tutorName, category, kelas) {
  const nameClean = tutorName.trim().toLowerCase();
  
  // Cari data setting tarif tutor dari tutorSalarySettings
  const tutorSetting = tutorSalarySettings.find(t => t.nama.trim().toLowerCase() === nameClean);

  const cat = (category || '').toUpperCase();
  const kls = (kelas || '').toUpperCase();

  // Jika setting spesifik tutor ditemukan:
  if (tutorSetting) {
    if (cat.includes('ARIANNA') || kls.includes('ARIANNA')) {
      return tutorSetting.honorHybrid || tutorSetting.honorSAC || 75000;
    }
    if (cat.includes('PESANTREN') || kls.includes('PESANTREN')) {
      return tutorSetting.honorPesantren || tutorSetting.honorLOSSD || 75000;
    }
    if (cat.includes('LOS SD') || kls.includes('LOS SD') || (kls.includes('SD') && cat.includes('LOS'))) {
      return tutorSetting.honorLOSSD || 50000;
    }
    if (cat.includes('LOS SMP') || kls.includes('LOS SMP') || (kls.includes('SMP') && cat.includes('LOS'))) {
      return tutorSetting.honorLOSSMP || tutorSetting.honorLOSSD || 107000;
    }
    if (cat.includes('LOS SMA') || kls.includes('LOS SMA') || (kls.includes('SMA') && cat.includes('LOS'))) {
      if (cat.includes('LUAR') || kls.includes('LUAR')) {
        return tutorSetting.honorLOSSMALK || 137500;
      }
      return tutorSetting.honorLOSSMADK || 125000;
    }
    if (cat.includes('ZAHRA') || kls.includes('ZAHRA') || cat.includes('LOS')) {
      // Hybrid Zahra / General LOS -> ambil tarif LOS SD / LOS SMP / default
      return tutorSetting.honorLOSSD || tutorSetting.honorLOSSMP || 50000;
    }
    
    // Default jika kategori SAC
    return tutorSetting.honorSAC || 50000;
  }

  // Fallback default jika nama tutor belum terdaftar di setting tarif
  if (cat.includes('LOS') || cat.includes('ZAHRA')) return 45000;
  return 35000;
}

// ========================================================
// 1. PAGE BERANDA: PAYROLL TUTOR (DATAFORM)
// ========================================================

async function fetchPayrollData() {
  const tbody = document.getElementById('payroll-table-body');
  if (tbody) tbody.innerHTML = `<tr><td colspan="6" class="p-8 text-center text-slate-400">Memuat data payroll...</td></tr>`;

  try {
    const data = await fetchGVizJson(URL_DATAFORM);
    const rows = data.table.rows || [];

    rawPayrollData = rows.map(row => {
      const isoDate = parseToIsoDate(getCellValue(row, 0), getCellFormatted(row, 0));
      return {
        tanggal: isoDate,
        tutor: String(getCellValue(row, 1)).trim(),
        kategori: String(getCellValue(row, 2)).trim().toUpperCase(),
        kelas: String(getCellValue(row, 5)).trim().toUpperCase()
      };
    }).filter(item => item.tutor !== '' && item.tutor !== 'Tanpa Nama' && item.tutor !== 'Nama Tutor');

    processPayrollData();
  } catch (err) {
    console.error("Error Payroll:", err);
    if (tbody) tbody.innerHTML = `<tr><td colspan="6" class="p-8 text-center text-red-600 font-medium">⚠️ Gagal memuat data DataForm: ${err.message}</td></tr>`;
  }
}

function processPayrollData() {
  const startDate = document.getElementById('start-date').value;
  const endDate = document.getElementById('end-date').value;

  const filteredRows = rawPayrollData.filter(item => {
    if (!item.tanggal) return true;
    if (startDate && item.tanggal < startDate) return false;
    if (endDate && item.tanggal > endDate) return false;
    return true;
  });

  const tutorMap = {};
  filteredRows.forEach(item => {
    const nama = item.tutor;
    if (!tutorMap[nama]) {
      tutorMap[nama] = { nama, countSAC: 0, countLOS: 0, subtotalSAC: 0, subtotalLOS: 0 };
    }

    const honorRate = calculateSessionHonor(nama, item.kategori, item.kelas);

    const isLOS = item.kategori.includes('LOS') || 
                  item.kategori.includes('ZAHRA') || 
                  item.kelas.includes('ZAHRA') ||
                  item.kategori.includes('HYBRID ZAHRA');

    if (isLOS) {
      tutorMap[nama].countLOS += 1;
      tutorMap[nama].subtotalLOS += honorRate;
    } else {
      tutorMap[nama].countSAC += 1;
      tutorMap[nama].subtotalSAC += honorRate;
    }
  });

  processedPayrollSummary = Object.values(tutorMap).map(tutor => {
    return {
      ...tutor,
      totalHonor: tutor.subtotalSAC + tutor.subtotalLOS
    };
  });

  processedPayrollSummary.sort((a, b) => a.nama.localeCompare(b.nama));
  renderPayrollTable(processedPayrollSummary);
  updatePayrollStats(processedPayrollSummary);
}

function renderPayrollTable(data) {
  const tbody = document.getElementById('payroll-table-body');
  if (!tbody) return;
  tbody.innerHTML = '';

  if (data.length === 0) {
    tbody.innerHTML = `<tr><td colspan="6" class="p-8 text-center text-slate-400">Tidak ada data payroll pada rentang tanggal ini.</td></tr>`;
    return;
  }

  data.forEach(tutor => {
    const tr = document.createElement('tr');
    tr.className = 'hover:bg-slate-50 transition border-b border-slate-200';
    tr.innerHTML = `
      <td class="p-3.5 font-bold text-slate-900">${tutor.nama}</td>
      <td class="p-3.5 text-center font-medium">${tutor.countSAC} Sesi</td>
      <td class="p-3.5 text-center font-medium">${tutor.countLOS} Sesi</td>
      <td class="p-3.5 text-right text-slate-600">${formatRupiah(tutor.subtotalSAC)}</td>
      <td class="p-3.5 text-right text-slate-600">${formatRupiah(tutor.subtotalLOS)}</td>
      <td class="p-3.5 text-right font-bold text-[#2271b1]">${formatRupiah(tutor.totalHonor)}</td>
    `;
    tbody.appendChild(tr);
  });
}

function updatePayrollStats(data) {
  const totalSAC = data.reduce((a, b) => a + b.countSAC, 0);
  const totalLOS = data.reduce((a, b) => a + b.countLOS, 0);
  const totalHonor = data.reduce((a, b) => a + b.totalHonor, 0);

  document.getElementById('stat-sac-count').innerText = `${totalSAC} Sesi`;
  document.getElementById('stat-los-count').innerText = `${totalLOS} Sesi`;
  document.getElementById('stat-tutor-count').innerText = `${data.length} Orang`;
  document.getElementById('stat-total-honor').innerText = formatRupiah(totalHonor);
}

function filterPayrollTable() {
  const query = document.getElementById('payroll-search').value.toLowerCase();
  const filtered = processedPayrollSummary.filter(t => t.nama.toLowerCase().includes(query));
  renderPayrollTable(filtered);
}

// Ekspor Payroll
function exportPayrollExcel() {
  if (!processedPayrollSummary.length) return alert("Tidak ada data payroll untuk diekspor!");
  const excelData = processedPayrollSummary.map(item => ({
    "Nama Tutor": item.nama,
    "Sesi SAC": item.countSAC,
    "Sesi LOS": item.countLOS,
    "Subtotal SAC (Rp)": item.subtotalSAC,
    "Subtotal LOS (Rp)": item.subtotalLOS,
    "Total Honor (Rp)": item.totalHonor
  }));
  const ws = XLSX.utils.json_to_sheet(excelData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Payroll Tutor");
  XLSX.writeFile(wb, "Payroll_Tutor_Carnation.xlsx");
}

function exportPayrollCSV() {
  if (!processedPayrollSummary.length) return alert("Tidak ada data payroll untuk diekspor!");
  let csv = "data:text/csv;charset=utf-8,\uFEFFNama Tutor,Sesi SAC,Sesi LOS,Subtotal SAC,Subtotal LOS,Total Honor\n";
  processedPayrollSummary.forEach(item => {
    csv += `"${item.nama}",${item.countSAC},${item.countLOS},${item.subtotalSAC},${item.subtotalLOS},${item.totalHonor}\n`;
  });
  const link = document.createElement("a");
  link.setAttribute("href", encodeURI(csv));
  link.setAttribute("download", "Payroll_Tutor_Carnation.csv");
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}

// ========================================================
// 2. PAGE REKAP ABSENSI (ATTENDANCE + USERS)
// ========================================================

async function fetchAttendanceAndUsers() {
  const tbody = document.getElementById('absensi-table-body');
  if (tbody) tbody.innerHTML = `<tr><td colspan="7" class="p-8 text-center text-slate-400">Memuat data absensi tutor...</td></tr>`;

  try {
    const usersData = await fetchGVizJson(URL_USERS);
    const userRows = usersData.table.rows || [];
    usersMap = {};

    userRows.forEach(row => {
      const idUser = String(getCellValue(row, 0)).trim();
      const namaUser = String(getCellValue(row, 1)).trim();
      const jabatanUser = String(getCellValue(row, 2)).trim();
      if (idUser) {
        usersMap[idUser] = { nama: namaUser, jabatan: jabatanUser };
      }
    });

    const attendanceData = await fetchGVizJson(URL_ATTENDANCE);
    const attRows = attendanceData.table.rows || [];

    rawAttendanceData = attRows.map(row => {
      const userId = String(getCellValue(row, 1)).trim();
      const userObj = usersMap[userId] || { nama: '', jabatan: '' };

      const isoDate = parseToIsoDate(getCellValue(row, 2), getCellFormatted(row, 2));

      return {
        id: getCellValue(row, 0),
        userId: userId,
        nama: userObj.nama,
        jabatan: userObj.jabatan,
        tanggal: isoDate,
        tanggalFmt: getCellFormatted(row, 2) || getCellValue(row, 2),
        jamMasuk: getCellValue(row, 3) || '-',
        jamKeluar: getCellValue(row, 4) || '-',
        status: String(getCellValue(row, 9) || 'Hadir').trim(),
        fotoIn: getCellValue(row, 10)
      };
    }).filter(item => item.nama !== '' && item.tanggal !== '' && item.jabatan.toLowerCase().includes('tutor'));

    processAttendanceData();

  } catch (err) {
    console.error("Error Attendance:", err);
    if (tbody) tbody.innerHTML = `<tr><td colspan="7" class="p-8 text-center text-red-600 font-medium">⚠️ Gagal memuat data Absensi: ${err.message}</td></tr>`;
  }
}

function processAttendanceData() {
  const startDate = document.getElementById('absensi-start-date').value;
  const endDate = document.getElementById('absensi-end-date').value;
  const statusFilter = document.getElementById('absensi-status-filter').value;

  processedAttendanceSummary = rawAttendanceData.filter(item => {
    if (startDate && item.tanggal < startDate) return false;
    if (endDate && item.tanggal > endDate) return false;
    if (statusFilter !== 'ALL' && item.status.toUpperCase() !== statusFilter) return false;
    return true;
  });

  processedAttendanceSummary.sort((a, b) => b.tanggal.localeCompare(a.tanggal));

  renderAttendanceTable(processedAttendanceSummary);
  updateAttendanceStats(processedAttendanceSummary);
}

function renderAttendanceTable(data) {
  const tbody = document.getElementById('absensi-table-body');
  if (!tbody) return;
  tbody.innerHTML = '';

  if (data.length === 0) {
    tbody.innerHTML = `<tr><td colspan="7" class="p-8 text-center text-slate-400">Tidak ada record absensi tutor sesuai filter.</td></tr>`;
    return;
  }

  data.forEach(item => {
    const isLate = item.status.toLowerCase() === 'terlambat';
    const statusBadge = isLate 
      ? `<span class="px-2 py-0.5 rounded text-[10px] font-bold bg-amber-100 text-amber-800 border border-amber-300">Terlambat</span>`
      : `<span class="px-2 py-0.5 rounded text-[10px] font-bold bg-emerald-100 text-emerald-800 border border-emerald-300">Hadir</span>`;

    const fotoBtn = item.fotoIn 
      ? `<a href="${item.fotoIn}" target="_blank" class="text-[#2271b1] hover:underline flex items-center justify-center gap-1"><i data-lucide="image" class="w-3.5 h-3.5"></i> Lihat Foto</a>`
      : `<span class="text-slate-400">-</span>`;

    const tr = document.createElement('tr');
    tr.className = 'hover:bg-slate-50 transition border-b border-slate-200';
    tr.innerHTML = `
      <td class="p-3.5 font-bold text-slate-900">${item.nama}</td>
      <td class="p-3.5 text-slate-600">${item.jabatan}</td>
      <td class="p-3.5 font-medium">${item.tanggalFmt || item.tanggal}</td>
      <td class="p-3.5 text-center font-mono font-semibold">${item.jamMasuk}</td>
      <td class="p-3.5 text-center font-mono font-semibold">${item.jamKeluar}</td>
      <td class="p-3.5 text-center">${statusBadge}</td>
      <td class="p-3.5 text-center">${fotoBtn}</td>
    `;
    tbody.appendChild(tr);
  });

  if (typeof lucide !== 'undefined') lucide.createIcons();
}

function updateAttendanceStats(data) {
  const totalRec = data.length;
  const hadirRec = data.filter(d => d.status.toLowerCase() === 'hadir').length;
  const lateRec = data.filter(d => d.status.toLowerCase() === 'terlambat').length;

  document.getElementById('stat-absensi-total').innerText = `${totalRec} Record`;
  document.getElementById('stat-absensi-hadir').innerText = `${hadirRec} Record`;
  document.getElementById('stat-absensi-terlambat').innerText = `${lateRec} Record`;
}

function filterAttendanceTable() {
  const query = document.getElementById('absensi-search').value.toLowerCase();
  const filtered = processedAttendanceSummary.filter(item => 
    item.nama.toLowerCase().includes(query) || 
    item.jabatan.toLowerCase().includes(query)
  );
  renderAttendanceTable(filtered);
}

function exportAttendanceExcel() {
  if (!processedAttendanceSummary.length) return alert("Tidak ada data absensi untuk diekspor!");
  const excelData = processedAttendanceSummary.map(item => ({
    "Nama Tutor": item.nama,
    "Jabatan": item.jabatan,
    "Tanggal": item.tanggalFmt || item.tanggal,
    "Jam Masuk": item.jamMasuk,
    "Jam Keluar": item.jamKeluar,
    "Status": item.status
  }));
  const ws = XLSX.utils.json_to_sheet(excelData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Rekap Absensi Tutor");
  XLSX.writeFile(wb, "Rekap_Absensi_Tutor_Carnation.xlsx");
}

function exportAttendanceCSV() {
  if (!processedAttendanceSummary.length) return alert("Tidak ada data absensi untuk diekspor!");
  let csv = "data:text/csv;charset=utf-8,\uFEFFNama Tutor,Jabatan,Tanggal,Jam Masuk,Jam Keluar,Status\n";
  processedAttendanceSummary.forEach(item => {
    csv += `"${item.nama}","${item.jabatan}","${item.tanggalFmt || item.tanggal}","${item.jamMasuk}","${item.jamKeluar}","${item.status}"\n`;
  });
  const link = document.createElement("a");
  link.setAttribute("href", encodeURI(csv));
  link.setAttribute("download", "Rekap_Absensi_Tutor_Carnation.csv");
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}

// ========================================================
// 3. PAGE TARIF HONOR TUTOR (INPUT NOMINAL & REKAP KHUSUS ROLE TUTOR)
// ========================================================

async function fetchSalarySettings() {
  const tbody = document.getElementById('tarif-tutor-table-body');
  if (tbody) tbody.innerHTML = `<tr><td colspan="9" class="p-8 text-center text-slate-400">Memuat data tarif honor tutor...</td></tr>`;

  try {
    const data = await fetchGVizJson(URL_SALARYSETTINGS);
    const rows = data.table.rows || [];

    const tutorNamesFromUsers = new Set(
      Object.values(usersMap)
        .filter(u => u.jabatan && u.jabatan.toLowerCase().includes('tutor'))
        .map(u => u.nama.trim().toLowerCase())
    );

    tutorSalarySettings = [];

    rows.forEach(row => {
      const nama = String(getCellValue(row, 0)).trim();
      const namaLower = nama.toLowerCase();

      const isTutorRole = tutorNamesFromUsers.size === 0 
        ? (!namaLower.includes('admin') && !namaLower.includes('office') && !namaLower.includes('direktur'))
        : tutorNamesFromUsers.has(namaLower);

      if (nama && nama !== 'Nama Karyawan' && isTutorRole) {
        const honorSAC = parseFloat(getCellValue(row, 1)) || 50000;
        const honorLOSSD = parseFloat(getCellValue(row, 10)) || 0;
        const honorLOSSMP = parseFloat(getCellValue(row, 11)) || 0;
        const honorPesantren = parseFloat(getCellValue(row, 13)) || 0;
        const honorHybrid = parseFloat(getCellValue(row, 14)) || 75000;
        const honorLOSSMADK = parseFloat(getCellValue(row, 15)) || 0;
        const honorLOSSMALK = parseFloat(getCellValue(row, 16)) || 0;

        tutorSalarySettings.push({
          nama,
          honorSAC,
          honorHybrid,
          honorLOSSD,
          honorLOSSMP,
          honorPesantren,
          honorLOSSMADK,
          honorLOSSMALK
        });
      }
    });

    tutorSalarySettings.sort((a, b) => a.nama.localeCompare(b.nama));

    populateTutorDropdown();
    renderTarifTutorTable();
    renderKegiatanTable();

  } catch (err) {
    console.error("Error SalarySettings:", err);
    useDefaultSalarySettings();
  }
}

function useDefaultSalarySettings() {
  const defaultTutors = [
    "ANNE NURMELIYANTI, S.Pd", "SUHAYATI, S.Sos", "AGUSTINA PRAJAWATI", "FADIL MUHAMMAD, S.Pd",
    "INDRA LESMANA, S.Pd", "LIA LIANTINA, S.Farm", "SHIFA AWIYAH", "YENI MAULIDAH",
    "ALIYA MASRURAH", "MAJIIDAH TUNNISWAH", "ANGGIE BAROKAH", "MOHAMMAD GALANG ARIFKI",
    "TANTI APRILIANTI, S.Pd", "GALIH PANGESTU WIJAYA", "RIBKA CHRISTIANI", "SYIFA'A NUR AFIYAH, S.Pd", "RIA RIYANTI, S.Pd"
  ];

  tutorSalarySettings = defaultTutors.map(nama => ({
    nama,
    honorSAC: 50000,
    honorHybrid: 75000,
    honorLOSSD: 0,
    honorLOSSMP: 0,
    honorPesantren: 0,
    honorLOSSMADK: 0,
    honorLOSSMALK: 0
  }));

  populateTutorDropdown();
  renderTarifTutorTable();
  renderKegiatanTable();
}

function populateTutorDropdown() {
  const select = document.getElementById('select-tutor-name');
  if (!select) return;

  select.innerHTML = '<option value="">-- Pilih Nama Tutor --</option>';
  tutorSalarySettings.forEach((tutor, idx) => {
    const opt = document.createElement('option');
    opt.value = idx;
    opt.textContent = tutor.nama;
    select.appendChild(opt);
  });
}

function onTutorSelectChange() {
  const select = document.getElementById('select-tutor-name');
  const idx = select.value;

  if (idx === "") {
    resetTutorTarifForm();
    return;
  }

  const t = tutorSalarySettings[idx];
  document.getElementById('input-honor-sac').value = t.honorSAC || '';
  document.getElementById('input-honor-hybrid').value = t.honorHybrid || '';
  document.getElementById('input-honor-los-sd').value = t.honorLOSSD || '';
  document.getElementById('input-honor-los-smp').value = t.honorLOSSMP || '';
  document.getElementById('input-honor-pesantren').value = t.honorPesantren || '';
  document.getElementById('input-honor-los-sma-dk').value = t.honorLOSSMADK || '';
  document.getElementById('input-honor-los-sma-lk').value = t.honorLOSSMALK || '';
}

function saveTutorTarif() {
  const select = document.getElementById('select-tutor-name');
  const idx = select.value;

  if (idx === "") {
    alert("Silakan pilih Nama Tutor terlebih dahulu!");
    return;
  }

  tutorSalarySettings[idx].honorSAC = parseFloat(document.getElementById('input-honor-sac').value) || 0;
  tutorSalarySettings[idx].honorHybrid = parseFloat(document.getElementById('input-honor-hybrid').value) || 0;
  tutorSalarySettings[idx].honorLOSSD = parseFloat(document.getElementById('input-honor-los-sd').value) || 0;
  tutorSalarySettings[idx].honorLOSSMP = parseFloat(document.getElementById('input-honor-los-smp').value) || 0;
  tutorSalarySettings[idx].honorPesantren = parseFloat(document.getElementById('input-honor-pesantren').value) || 0;
  tutorSalarySettings[idx].honorLOSSMADK = parseFloat(document.getElementById('input-honor-los-sma-dk').value) || 0;
  tutorSalarySettings[idx].honorLOSSMALK = parseFloat(document.getElementById('input-honor-los-sma-lk').value) || 0;

  renderTarifTutorTable();
  processPayrollData(); // Update juga perhitungan Beranda secara langsung
  alert(`Nominal honor untuk ${tutorSalarySettings[idx].nama} berhasil diperbarui!`);
}

function resetTutorTarifForm() {
  document.getElementById('select-tutor-name').value = "";
  document.getElementById('input-honor-sac').value = "";
  document.getElementById('input-honor-hybrid').value = "";
  document.getElementById('input-honor-los-sd').value = "";
  document.getElementById('input-honor-los-smp').value = "";
  document.getElementById('input-honor-pesantren').value = "";
  document.getElementById('input-honor-los-sma-dk').value = "";
  document.getElementById('input-honor-los-sma-lk').value = "";
}

function renderTarifTutorTable() {
  const tbody = document.getElementById('tarif-tutor-table-body');
  if (!tbody) return;
  tbody.innerHTML = '';

  if (tutorSalarySettings.length === 0) {
    tbody.innerHTML = `<tr><td colspan="9" class="p-8 text-center text-slate-400">Belum ada data tarif tutor.</td></tr>`;
    return;
  }

  tutorSalarySettings.forEach((item, idx) => {
    const tr = document.createElement('tr');
    tr.className = 'hover:bg-slate-50 transition border-b border-slate-200';
    tr.innerHTML = `
      <td class="p-2.5 text-center font-semibold text-slate-500">${idx + 1}</td>
      <td class="p-2.5 font-bold text-slate-900">${item.nama}</td>
      <td class="p-2.5 text-right font-medium text-slate-700">${formatRupiah(item.honorSAC)}</td>
      <td class="p-2.5 text-right font-medium text-slate-700">${formatRupiah(item.honorHybrid)}</td>
      <td class="p-2.5 text-right font-medium text-slate-700">${formatRupiah(item.honorLOSSD)}</td>
      <td class="p-2.5 text-right font-medium text-slate-700">${formatRupiah(item.honorLOSSMP)}</td>
      <td class="p-2.5 text-right font-medium text-slate-700">${formatRupiah(item.honorPesantren)}</td>
      <td class="p-2.5 text-right font-medium text-slate-700">${formatRupiah(item.honorLOSSMADK)}</td>
      <td class="p-2.5 text-right font-medium text-slate-700">${formatRupiah(item.honorLOSSMALK)}</td>
    `;
    tbody.appendChild(tr);
  });
}

function renderKegiatanTable() {
  const tbody = document.getElementById('kegiatan-table-body');
  if (!tbody) return;
  tbody.innerHTML = '';

  daftarHonorKegiatanMaster.forEach(item => {
    const tr = document.createElement('tr');
    tr.className = 'hover:bg-slate-50 transition border-b border-slate-200';
    tr.innerHTML = `
      <td class="p-2.5 text-center font-semibold text-slate-500">${item.no}</td>
      <td class="p-2.5 font-semibold text-slate-800">${item.kegiatan}</td>
      <td class="p-2.5 text-right font-bold text-[#2271b1]">${formatRupiah(item.honor)}</td>
    `;
    tbody.appendChild(tr);
  });
}

// ========================================================
// 4. NAVIGATION TAB SWITCHER
// ========================================================

function switchTab(tab) {
  const pageBeranda = document.getElementById('page-beranda');
  const pageAbsensi = document.getElementById('page-absensi');
  const pageTarif = document.getElementById('page-tarif');

  const navBeranda = document.getElementById('nav-beranda');
  const navAbsensi = document.getElementById('nav-absensi');
  const navTarif = document.getElementById('nav-tarif');

  const pageTitleTop = document.getElementById('page-title-topbar');

  pageBeranda.classList.add('hidden');
  pageAbsensi.classList.add('hidden');
  pageTarif.classList.add('hidden');

  navBeranda.className = 'w-full flex items-center gap-3 px-3 py-2.5 rounded text-xs font-semibold hover:bg-[#2c3338] text-slate-300 hover:text-white transition';
  navAbsensi.className = 'w-full flex items-center gap-3 px-3 py-2.5 rounded text-xs font-semibold hover:bg-[#2c3338] text-slate-300 hover:text-white transition';
  navTarif.className = 'w-full flex items-center gap-3 px-3 py-2.5 rounded text-xs font-semibold hover:bg-[#2c3338] text-slate-300 hover:text-white transition';

  if (tab === 'beranda') {
    pageBeranda.classList.remove('hidden');
    navBeranda.className = 'w-full flex items-center gap-3 px-3 py-2.5 rounded text-xs font-semibold bg-[#2271b1] text-white transition';
    pageTitleTop.innerText = "Beranda (Payroll Tutor)";
  } else if (tab === 'absensi') {
    pageAbsensi.classList.remove('hidden');
    navAbsensi.className = 'w-full flex items-center gap-3 px-3 py-2.5 rounded text-xs font-semibold bg-[#2271b1] text-white transition';
    pageTitleTop.innerText = "Rekap Absensi Tutor";
  } else if (tab === 'tarif') {
    pageTarif.classList.remove('hidden');
    navTarif.className = 'w-full flex items-center gap-3 px-3 py-2.5 rounded text-xs font-semibold bg-[#2271b1] text-white transition';
    pageTitleTop.innerText = "Setting Tarif Honor Tutor";
  }

  if (typeof lucide !== 'undefined') lucide.createIcons();
}

// INITIALIZATION
window.addEventListener('DOMContentLoaded', () => {
  if (typeof lucide !== 'undefined') lucide.createIcons();
  fetchAllData();
});