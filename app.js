// CONFIGURATION: Ganti dengan ID Spreadsheet Anda
const SPREADSHEET_ID = '1PkmW4hah19NQjfd__Sz2c3ZGONO4GWDuukVs4ViLGYE';
const SHEET_NAME = 'DataForm'; // Nama tab di spreadsheet Anda

const GViz_URL = `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/gviz/tq?tqx=out:json&sheet=${SHEET_NAME}`;

let rawPayrollData = [];

// Format Angka ke Rupiah
function formatRupiah(number) {
  return new Intl.NumberFormat('id-ID', { style: 'currency', currency: 'IDR', maximumFractionDigits: 0 }).format(number || 0);
}

// Fetch Data dari Google Sheets via GViz API
async function fetchPayrollData() {
  const tbody = document.getElementById('payroll-table-body');
  tbody.innerHTML = `<tr><td colspan="7" class="p-8 text-center text-slate-400">Memperbarui data...</td></tr>`;

  try {
    const response = await fetch(GViz_URL);
    const text = await response.text();
    
    // Parsing response GViz JSON (Menghapus pembungkus fungsi google.visualization.Query.setResponse)
    const jsonString = text.substring(47, text.length - 2);
    const data = JSON.parse(jsonString);

    const rows = data.table.rows;
    
    // Mapping baris spreadsheet ke objek JSON
    // Kolom 0: Nama, Kolom 1: Jabatan, Kolom 2: Gaji Pokok, Kolom 3: Tunjangan, Kolom 4: Potongan, Kolom 5: Status
    rawPayrollData = rows.map(row => {
      const gP = row.c[2] ? (row.c[2].v || 0) : 0;
      const tj = row.c[3] ? (row.c[3].v || 0) : 0;
      const pt = row.c[4] ? (row.c[4].v || 0) : 0;
      return {
        nama: row.c[0] ? row.c[0].v : '-',
        jabatan: row.c[1] ? row.c[1].v : '-',
        gajiPokok: gP,
        tunjangan: tj,
        potongan: pt,
        total: (gP + tj) - pt,
        status: row.c[5] ? row.c[5].v : 'Proses'
      };
    });

    renderTable(rawPayrollData);
    updateStats(rawPayrollData);

  } catch (error) {
    console.error('Error fetching data:', error);
    tbody.innerHTML = `<tr><td colspan="7" class="p-8 text-center text-red-500">Gagal mengambil data. Pastikan ID Spreadsheet benar dan link publik telah diaktifkan.</td></tr>`;
  }
}

// Render Data ke Tabel
function renderTable(data) {
  const tbody = document.getElementById('payroll-table-body');
  tbody.innerHTML = '';

  if (data.length === 0) {
    tbody.innerHTML = `<tr><td colspan="7" class="p-8 text-center text-slate-400">Tidak ada data ditemukan.</td></tr>`;
    return;
  }

  data.forEach(item => {
    const isPaid = item.status.toLowerCase() === 'lunas' || item.status.toLowerCase() === 'paid';
    const statusClass = isPaid ? 'bg-emerald-50 text-emerald-700 border-emerald-200' : 'bg-amber-50 text-amber-700 border-amber-200';

    const tr = document.createElement('tr');
    tr.className = 'hover:bg-slate-50 transition border-b border-slate-100';
    tr.innerHTML = `
      <td class="p-4 font-medium text-slate-900">${item.nama}</td>
      <td class="p-4 text-slate-600">${item.jabatan}</td>
      <td class="p-4 text-slate-600">${formatRupiah(item.gajiPokok)}</td>
      <td class="p-4 text-slate-600">${formatRupiah(item.tunjangan)}</td>
      <td class="p-4 text-slate-600">${formatRupiah(item.potongan)}</td>
      <td class="p-4 font-bold text-slate-900">${formatRupiah(item.total)}</td>
      <td class="p-4">
        <span class="px-2.5 py-1 text-xs font-semibold rounded-full border ${statusClass}">
          ${item.status}
        </span>
      </td>
    `;
    tbody.appendChild(tr);
  });
}

// Update Statistik Dashboard
function updateStats(data) {
  const totalPengeluaran = data.reduce((acc, curr) => acc + curr.total, 0);
  document.getElementById('stat-total').innerText = formatRupiah(totalPengeluaran);
  document.getElementById('stat-count').innerText = data.length;
}

// Fitur Filter / Pencarian
function filterData() {
  const query = document.getElementById('search-input').value.toLowerCase();
  const filtered = rawPayrollData.filter(item => 
    item.nama.toLowerCase().includes(query) || 
    item.jabatan.toLowerCase().includes(query)
  );
  renderTable(filtered);
}

// Jalankan saat pertama kali dibuka
document.addEventListener('DOMContentLoaded', fetchPayrollData);// CONFIGURATION: ID Spreadsheet & Sheet
const SPREADSHEET_ID = '1PkmW4hah19NQjfd__Sz2c3ZGONO4GWDuukVs4ViLGYE';
const SHEET_NAME = 'DataForm'; 

const GViz_URL = `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/gviz/tq?tqx=out:json&sheet=${SHEET_NAME}`;

let rawAttendanceData = [];
let processedTutorSummary = [];

// Format Angka ke Rupiah
function formatRupiah(number) {
  return new Intl.NumberFormat('id-ID', { style: 'currency', currency: 'IDR', maximumFractionDigits: 0 }).format(number || 0);
}

// Helper untuk mengubah tanggal dari berbagai format (DD/MM/YYYY atau Date(Y,M,D)) ke YYYY-MM-DD
function parseToIsoDate(rawVal, formattedVal) {
  let strVal = formattedVal || rawVal || '';
  if (!strVal) return '';

  // Kasus format GViz Date(YYYY,M,D)
  if (typeof strVal === 'string' && strVal.includes('Date(')) {
    const matches = strVal.match(/\d+/g);
    if (matches && matches.length >= 3) {
      const y = matches[0];
      const m = String(Number(matches[1]) + 1).padStart(2, '0');
      const d = String(matches[2]).padStart(2, '0');
      return `${y}-${m}-${d}`;
    }
  }

  // Kasus format teks DD/MM/YYYY
  if (typeof strVal === 'string' && strVal.includes('/')) {
    const parts = strVal.split('/');
    if (parts.length === 3) {
      const d = parts[0].padStart(2, '0');
      const m = parts[1].padStart(2, '0');
      const y = parts[2].trim();
      return `${y}-${m}-${d}`;
    }
  }

  return strVal;
}

// Fetch Data dari Google Sheets via GViz API
async function fetchPayrollData() {
  const tbody = document.getElementById('payroll-table-body');
  tbody.innerHTML = `<tr><td colspan="6" class="p-8 text-center text-slate-400">Memperbarui data absensi...</td></tr>`;

  try {
    const response = await fetch(GViz_URL);
    const text = await response.text();
    
    // Parsing JSON GViz
    const jsonString = text.substring(47, text.length - 2);
    const data = JSON.parse(jsonString);

    const rows = data.table.rows;
    
    // Mapping Kolom DataForm:
    // [0]: Tanggal | [1]: Nama Tutor | [2]: Mengajar (SAC/LOS) | [3]: Jam Mulai | [4]: Jam Selesai | [5]: Kelas | [6]: Mapel
    rawAttendanceData = rows.map(row => {
      const cellTanggalVal = row.c[0] ? row.c[0].v : '';
      const cellTanggalFormatted = row.c[0] ? row.c[0].f : '';
      const isoDate = parseToIsoDate(cellTanggalVal, cellTanggalFormatted);

      return {
        tanggal: isoDate,
        tutor: row.c[1] && row.c[1].v ? String(row.c[1].v).trim() : 'Tanpa Nama',
        kategori: row.c[2] && row.c[2].v ? String(row.c[2].v).trim().toUpperCase() : 'SAC',
        kelas: row.c[5] ? row.c[5].v : '',
        mapel: row.c[6] ? row.c[6].v : ''
      };
    }).filter(item => item.tutor !== 'Tanpa Nama' && item.tutor !== 'Nama Tutor');

    processData();

  } catch (error) {
    console.error('Error fetching data:', error);
    tbody.innerHTML = `<tr><td colspan="6" class="p-8 text-center text-red-500">Gagal mengambil data. Pastikan Spreadsheet ID benar dan link publik telah diaktifkan.</td></tr>`;
  }
}

// Filter Berdasarkan Rentang Tanggal & Hitung Sesi Per Tutor
function processData() {
  const startDate = document.getElementById('start-date').value;
  const endDate = document.getElementById('end-date').value;
  const rateSAC = parseFloat(document.getElementById('rate-sac').value) || 0;
  const rateLOS = parseFloat(document.getElementById('rate-los').value) || 0;

  // Filter baris sesuai tanggal jika diisi
  const filteredRows = rawAttendanceData.filter(item => {
    if (!item.tanggal) return true;
    if (startDate && item.tanggal < startDate) return false;
    if (endDate && item.tanggal > endDate) return false;
    return true;
  });

  // Agregasi / Grouping berdasarkan Nama Tutor
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

    // Setiap baris di DataForm dihitung 1 sesi (SAC atau LOS)
    if (item.kategori.includes('LOS')) {
      tutorMap[nama].countLOS += 1;
    } else {
      tutorMap[nama].countSAC += 1;
    }
  });

  // Kalkulasi Honor Subtotal & Total
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

  // Urutkan berdasarkan Nama Tutor (Alphabetical)
  processedTutorSummary.sort((a, b) => a.nama.localeCompare(b.nama));

  renderTable(processedTutorSummary);
  updateStats(processedTutorSummary);
}

// Render Tabel
function renderTable(data) {
  const tbody = document.getElementById('payroll-table-body');
  tbody.innerHTML = '';

  if (data.length === 0) {
    tbody.innerHTML = `<tr><td colspan="6" class="p-8 text-center text-slate-400">Tidak ada data absensi pada rentang tanggal yang dipilih.</td></tr>`;
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

// Jalankan saat halaman siap
document.addEventListener('DOMContentLoaded', fetchPayrollData);