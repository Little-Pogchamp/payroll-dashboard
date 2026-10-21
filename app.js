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
document.addEventListener('DOMContentLoaded', fetchPayrollData);