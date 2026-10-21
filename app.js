// CONFIGURATION
const SPREADSHEET_ID = '1PkmW4hah19NQjfd__Sz2c3ZGONO4GWDuukVs4ViLGYE';
const SHEET_NAME = 'DataForm'; 

const GViz_URL = `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/gviz/tq?tqx=out:json&sheet=${SHEET_NAME}`;

let rawAttendanceData = [];
let processedTutorSummary = [];

// Format Angka ke Rupiah
function formatRupiah(number) {
  return new Intl.NumberFormat('id-ID', { style: 'currency', currency: 'IDR', maximumFractionDigits: 0 }).format(number || 0);
}

// Mapping nama bulan Bahasa Indonesia
const indonesianMonths = {
  'januari': '01', 'jan': '01',
  'februari': '02', 'feb': '02',
  'maret': '03', 'mar': '03',
  'april': '04', 'apr': '04',
  'mei': '05',
  'juni': '06', 'jun': '06',
  'juli': '07', 'jul': '07',
  'agustus': '08', 'ags': '08', 'agu': '08',
  'september': '09', 'sep': '09',
  'oktober': '10', 'okt': '10',
  'november': '11', 'nov': '11',
  'desember': '12', 'des': '12'
};

// Helper membaca cell secara aman (Null-Safe)
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

// Konversi DD/MM/YYYY, Date(Y,M,D), atau "Senin, 5 Oktober 2026" ke YYYY-MM-DD
function parseToIsoDate(rawVal, formattedVal) {
  let strVal = String(formattedVal || rawVal || '').trim();
  if (!strVal) return '';

  // 1. Format GViz Date(YYYY,M,D)
  if (strVal.includes('Date(')) {
    const matches = strVal.match(/\d+/g);
    if (matches && matches.length >= 3) {
      const y = matches[0];
      const m = String(Number(matches[1]) + 1).padStart(2, '0');
      const d = String(matches[2]).padStart(2, '0');
      return `${y}-${m}-${d}`;
    }
  }

  // 2. Format Angka DD/MM/YYYY
  if (strVal.includes('/')) {
    const parts = strVal.split('/');
    if (parts.length === 3) {
      const d = parts[0].padStart(2, '0');
      const m = parts[1].padStart(2, '0');
      const y = parts[2].trim();
      return `${y}-${m}-${d}`;
    }
  }

  // 3. Format ISO YYYY-MM-DD
  if (/^\d{4}-\d{2}-\d{2}$/.test(strVal)) {
    return strVal;
  }

  // 4. Format Teks Indonesia (contoh: "Senin, 5 Oktober 2026" atau "5 Oktober 2026")
  let cleanStr = strVal.replace(/^(senin|selasa|rabu|kamis|jumat|sabtu|minggu)[,\s]*/i, '').trim();
  const parts = cleanStr.split(/\s+/);
  
  if (parts.length >= 3) {
    const day = parts[0].padStart(2, '0');
    const monthStr = parts[1].toLowerCase();
    const year = parts[2];
    const month = indonesianMonths[monthStr] || '01';
    
    if (!isNaN(day) && !isNaN(year)) {
      return `${year}-${month}-${day}`;
    }
  }

  return strVal;
}

// Fetch Data dari Google Sheets via GViz API
async function fetchPayrollData() {
  const tbody = document.getElementById('payroll-table-body');
  if (tbody) {
    tbody.innerHTML = `<tr><td colspan="6" class="p-8 text-center text-slate-400">Memperbarui data absensi dari Google Sheets...</td></tr>`;
  }

  try {
    const response = await fetch(GViz_URL);
    
    if (!response.ok) {
      throw new Error(`HTTP Error ${response.status}: Gagal terhubung ke Google Sheets.`);
    }

    const text = await response.text();

    const firstOpen = text.indexOf('{');
    const lastClose = text.lastIndexOf('}');
    
    if (firstOpen === -1 || lastClose === -1) {
      throw new Error("Izin akses ditolak. Pastikan Google Sheet dibagikan ke 'Siapa saja yang memiliki link'.");
    }

    const jsonString = text.substring(firstOpen, lastClose + 1);
    const data = JSON.parse(jsonString);

    if (!data.table || !data.table.rows) {
      throw new Error("Struktur data spreadsheet kosong.");
    }

    const rows = data.table.rows;

    // Mapping Kolom DataForm:
    // [0]: Tanggal | [1]: Nama Tutor | [2]: Mengajar (SAC/LOS) | [5]: Kelas | [6]: Mapel
    rawAttendanceData = rows.map(row => {
      const valTanggal = getCellValue(row, 0);
      const fmtTanggal = getCellFormatted(row, 0);
      const isoDate = parseToIsoDate(valTanggal, fmtTanggal);

      const tutorName = String(getCellValue(row, 1)).trim();
      const kategoriStr = String(getCellValue(row, 2)).trim().toUpperCase();

      return {
        tanggal: isoDate,
        tutor: tutorName,
        kategori: kategoriStr,
        kelas: getCellValue(row, 5),
        mapel: getCellValue(row, 6)
      };
    }).filter(item => item.tutor !== '' && item.tutor !== 'Tanpa Nama' && item.tutor !== 'Nama Tutor');

    processData();

  } catch (error) {
    console.error('Error fetching data:', error);
    if (tbody) {
      tbody.innerHTML = `
        <tr>
          <td colspan="6" class="p-8 text-center text-red-600 font-medium">
            ⚠️ ${error.message}<br>
            <span class="text-xs text-slate-500 font-normal mt-1 block">
              Pastikan Google Sheet di-Share dengan akses "Siapa saja yang memiliki link" sebagai Pelihat (Viewer).
            </span>
          </td>
        </tr>`;
    }
  }
}

// Filter Berdasarkan Rentang Tanggal & Hitung Sesi Per Tutor
function processData() {
  const startDate = document.getElementById('start-date').value;
  const endDate = document.getElementById('end-date').value;
  const rateSAC = parseFloat(document.getElementById('rate-sac').value) || 0;
  const rateLOS = parseFloat(document.getElementById('rate-los').value) || 0;

  const filteredRows = rawAttendanceData.filter(item => {
    if (!item.tanggal) return true;
    if (startDate && item.tanggal < startDate) return false;
    if (endDate && item.tanggal > endDate) return false;
    return true;
  });

  const tutorMap = {};

  filteredRows.forEach(item => {
    const nama = item.tutor;
    if (!tutorMap[nama]) {
      tutorMap[nama] = {
        nama: nama,
        countSAC: 0,
        countLOS: 0
      };
    }

    if (item.kategori.includes('LOS')) {
      tutorMap[nama].countLOS += 1;
    } else {
      tutorMap[nama].countSAC += 1;
    }
  });

  processedTutorSummary = Object.values(tutorMap).map(tutor => {
    const subtotalSAC = tutor.countSAC * rateSAC;
    const subtotalLOS = tutor.countLOS * rateLOS;
    return {
      ...tutor,
      subtotalSAC,
      subtotalLOS,
      totalHonor: subtotalSAC + subtotalLOS
    };
  });

  processedTutorSummary.sort((a, b) => a.nama.localeCompare(b.nama));

  renderTable(processedTutorSummary);
  updateStats(processedTutorSummary);
}

// Render Tabel
function renderTable(data) {
  const tbody = document.getElementById('payroll-table-body');
  if (!tbody) return;
  
  tbody.innerHTML = '';

  if (data.length === 0) {
    tbody.innerHTML = `<tr><td colspan="6" class="p-8 text-center text-slate-400">Tidak ada data absensi pada rentang tanggal ini.</td></tr>`;
    return;
  }

  data.forEach(tutor => {
    const tr = document.createElement('tr');
    tr.className = 'hover:bg-slate-50 transition border-b border-slate-100';
    tr.innerHTML = `
      <td class="p-4 font-semibold text-slate-900">${tutor.nama}</td>
      <td class="p-4 text-center font-medium text-slate-700">${tutor.countSAC} Sesi</td>
      <td class="p-4 text-center font-medium text-slate-700">${tutor.countLOS} Sesi</td>
      <td class="p-4 text-right text-slate-600">${formatRupiah(tutor.subtotalSAC)}</td>
      <td class="p-4 text-right text-slate-600">${formatRupiah(tutor.subtotalLOS)}</td>
      <td class="p-4 text-right font-bold text-indigo-600">${formatRupiah(tutor.totalHonor)}</td>
    `;
    tbody.appendChild(tr);
  });
}

// Update Statistik Ringkasan
function updateStats(data) {
  const totalSAC = data.reduce((acc, curr) => acc + curr.countSAC, 0);
  const totalLOS = data.reduce((acc, curr) => acc + curr.countLOS, 0);
  const totalHonor = data.reduce((acc, curr) => acc + curr.totalHonor, 0);

  document.getElementById('stat-sac-count').innerText = `${totalSAC} Sesi`;
  document.getElementById('stat-los-count').innerText = `${totalLOS} Sesi`;
  document.getElementById('stat-tutor-count').innerText = `${data.length} Orang`;
  document.getElementById('stat-total-honor').innerText = formatRupiah(totalHonor);
}

// Fitur Filter / Pencarian
function filterData() {
  const query = document.getElementById('search-input').value.toLowerCase();
  const filtered = processedTutorSummary.filter(item => 
    item.nama.toLowerCase().includes(query)
  );
  renderTable(filtered);
}

// FUNGSI EKSPOR EXCEL & CSV
function exportToExcel() {
  if (!processedTutorSummary || processedTutorSummary.length === 0) {
    alert("Tidak ada data untuk diekspor!");
    return;
  }

  const startDate = document.getElementById('start-date').value || 'Awal';
  const endDate = document.getElementById('end-date').value || 'Akhir';

  const excelData = processedTutorSummary.map(item => ({
    "Nama Tutor": item.nama,
    "Sesi SAC": item.countSAC,
    "Sesi LOS": item.countLOS,
    "Subtotal SAC (Rp)": item.subtotalSAC,
    "Subtotal LOS (Rp)": item.subtotalLOS,
    "Total Honor (Rp)": item.totalHonor
  }));

  const worksheet = XLSX.utils.json_to_sheet(excelData);
  worksheet['!cols'] = [
    { wch: 30 }, { wch: 12 }, { wch: 12 }, { wch: 20 }, { wch: 20 }, { wch: 22 }
  ];

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Rekap Honor");
  XLSX.writeFile(workbook, `Rekap_Honor_Tutor_${startDate}_sd_${endDate}.xlsx`);
}

function exportToCSV() {
  if (!processedTutorSummary || processedTutorSummary.length === 0) {
    alert("Tidak ada data untuk diekspor!");
    return;
  }

  const startDate = document.getElementById('start-date').value || 'Awal';
  const endDate = document.getElementById('end-date').value || 'Akhir';

  let csvContent = "data:text/csv;charset=utf-8,\uFEFF";
  csvContent += "Nama Tutor,Sesi SAC,Sesi LOS,Subtotal SAC (Rp),Subtotal LOS (Rp),Total Honor (Rp)\n";

  processedTutorSummary.forEach(item => {
    const row = [
      `"${item.nama}"`,
      item.countSAC,
      item.countLOS,
      item.subtotalSAC,
      item.subtotalLOS,
      item.totalHonor
    ].join(",");
    csvContent += row + "\n";
  });

  const encodedUri = encodeURI(csvContent);
  const link = document.createElement("a");
  link.setAttribute("href", encodedUri);
  link.setAttribute("download", `Rekap_Honor_Tutor_${startDate}_sd_${endDate}.csv`);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}

// Inisialisasi Otomatis saat Halaman Dimuat
window.addEventListener('DOMContentLoaded', () => {
  if (typeof lucide !== 'undefined') {
    lucide.createIcons();
  }
  fetchPayrollData();
});