// CONFIGURATION
const SPREADSHEET_ID = '1PkmW4hah19NQjfd__Sz2c3ZGONO4GWDuukVs4ViLGYE';

const URL_DATAFORM = `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/gviz/tq?tqx=out:json&sheet=DataForm`;
const URL_USERS = `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/gviz/tq?tqx=out:json&sheet=Users`;
const URL_ATTENDANCE = `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/gviz/tq?tqx=out:json&sheet=Attendance`;

// Global State
let rawPayrollData = [];
let processedPayrollSummary = [];

let usersMap = {}; // { id_user: { nama, jabatan } }
let rawAttendanceData = [];
let processedAttendanceSummary = [];

// Format Angka ke Rupiah
function formatRupiah(number) {
  return new Intl.NumberFormat('id-ID', { style: 'currency', currency: 'IDR', maximumFractionDigits: 0 }).format(number || 0);
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

// Helper untuk Fetch & Parse JSON GViz API
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

// FETCH ALL DATA (DataForm + Users + Attendance)
async function fetchAllData() {
  await fetchPayrollData();
  await fetchAttendanceAndUsers();
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
        kategori: String(getCellValue(row, 2)).trim().toUpperCase()
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
  const rateSAC = parseFloat(document.getElementById('rate-sac').value) || 0;
  const rateLOS = parseFloat(document.getElementById('rate-los').value) || 0;

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
      tutorMap[nama] = { nama, countSAC: 0, countLOS: 0 };
    }
    if (item.kategori.includes('LOS')) {
      tutorMap[nama].countLOS += 1;
    } else {
      tutorMap[nama].countSAC += 1;
    }
  });

  processedPayrollSummary = Object.values(tutorMap).map(tutor => {
    const subtotalSAC = tutor.countSAC * rateSAC;
    const subtotalLOS = tutor.countLOS * rateLOS;
    return { ...tutor, subtotalSAC, subtotalLOS, totalHonor: subtotalSAC + subtotalLOS };
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
  if (tbody) tbody.innerHTML = `<tr><td colspan="7" class="p-8 text-center text-slate-400">Memuat data absensi & pegawai...</td></tr>`;

  try {
    // A. Fetch Sheet Users
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

    // B. Fetch Sheet Attendance
    const attendanceData = await fetchGVizJson(URL_ATTENDANCE);
    const attRows = attendanceData.table.rows || [];

    rawAttendanceData = attRows.map(row => {
      const userId = String(getCellValue(row, 1)).trim();
      const userObj = usersMap[userId] || { nama: 'Pegawai Tanpa Nama', jabatan: 'Pegawai' };

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
    }).filter(item => item.nama !== 'Pegawai Tanpa Nama' && item.tanggal !== '');

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

  // Urutkan dari tanggal terbaru
  processedAttendanceSummary.sort((a, b) => b.tanggal.localeCompare(a.tanggal));

  renderAttendanceTable(processedAttendanceSummary);
  updateAttendanceStats(processedAttendanceSummary);
}

function renderAttendanceTable(data) {
  const tbody = document.getElementById('absensi-table-body');
  if (!tbody) return;
  tbody.innerHTML = '';

  if (data.length === 0) {
    tbody.innerHTML = `<tr><td colspan="7" class="p-8 text-center text-slate-400">Tidak ada record absensi sesuai filter.</td></tr>`;
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

// Ekspor Absensi
function exportAttendanceExcel() {
  if (!processedAttendanceSummary.length) return alert("Tidak ada data absensi untuk diekspor!");
  const excelData = processedAttendanceSummary.map(item => ({
    "Nama Pegawai": item.nama,
    "Jabatan": item.jabatan,
    "Tanggal": item.tanggalFmt || item.tanggal,
    "Jam Masuk": item.jamMasuk,
    "Jam Keluar": item.jamKeluar,
    "Status": item.status
  }));
  const ws = XLSX.utils.json_to_sheet(excelData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Rekap Absensi");
  XLSX.writeFile(wb, "Rekap_Absensi_Carnation.xlsx");
}

function exportAttendanceCSV() {
  if (!processedAttendanceSummary.length) return alert("Tidak ada data absensi untuk diekspor!");
  let csv = "data:text/csv;charset=utf-8,\uFEFFNama Pegawai,Jabatan,Tanggal,Jam Masuk,Jam Keluar,Status\n";
  processedAttendanceSummary.forEach(item => {
    csv += `"${item.nama}","${item.jabatan}","${item.tanggalFmt || item.tanggal}","${item.jamMasuk}","${item.jamKeluar}","${item.status}"\n`;
  });
  const link = document.createElement("a");
  link.setAttribute("href", encodeURI(csv));
  link.setAttribute("download", "Rekap_Absensi_Carnation.csv");
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}

// ========================================================
// 3. NAVIGATION TAB SWITCHER
// ========================================================

function switchTab(tab) {
  const pageBeranda = document.getElementById('page-beranda');
  const pageAbsensi = document.getElementById('page-absensi');
  const navBeranda = document.getElementById('nav-beranda');
  const navAbsensi = document.getElementById('nav-absensi');
  const pageTitleTop = document.getElementById('page-title-topbar');

  if (tab === 'beranda') {
    pageBeranda.classList.remove('hidden');
    pageAbsensi.classList.add('hidden');

    navBeranda.className = 'w-full flex items-center gap-3 px-3 py-2.5 rounded text-xs font-semibold bg-[#2271b1] text-white transition';
    navAbsensi.className = 'w-full flex items-center gap-3 px-3 py-2.5 rounded text-xs font-semibold hover:bg-[#2c3338] text-slate-300 hover:text-white transition';

    pageTitleTop.innerText = "Beranda (Payroll Tutor)";
  } else if (tab === 'absensi') {
    pageBeranda.classList.add('hidden');
    pageAbsensi.classList.remove('hidden');

    navAbsensi.className = 'w-full flex items-center gap-3 px-3 py-2.5 rounded text-xs font-semibold bg-[#2271b1] text-white transition';
    navBeranda.className = 'w-full flex items-center gap-3 px-3 py-2.5 rounded text-xs font-semibold hover:bg-[#2c3338] text-slate-300 hover:text-white transition';

    pageTitleTop.innerText = "Rekap Absensi Pegawai";
  }

  if (typeof lucide !== 'undefined') lucide.createIcons();
}

// INITIALIZATION
window.addEventListener('DOMContentLoaded', () => {
  if (typeof lucide !== 'undefined') lucide.createIcons();
  fetchAllData();
});