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

// Master List Honor Kegiatan
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

// Format Angka ke Rupiah Formatted
function formatRupiah(number) {
  if (number === null || number === undefined || number === '' || isNaN(number) || number === 0) {
    return '-';
  }
  return new Intl.NumberFormat('id-ID', { style: 'currency', currency: 'IDR', maximumFractionDigits: 0 }).format(number);
}

function formatNumberExcel(number) {
  if (!number) return '0';
  return new Intl.NumberFormat('id-ID').format(number);
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

// FETCH ALL DATA (Urutan: Attendance & Users -> SalarySettings -> Payroll)
async function fetchAllData() {
  await fetchAttendanceAndUsers();
  await fetchSalarySettings();
  await fetchPayrollData();
}

// Helper Deteksi Tarif & Label Per Sesi
function calculateSessionInfo(tutorName, category, kelas) {
  const nameClean = tutorName.trim().toLowerCase();
  const tutorSetting = tutorSalarySettings.find(t => t.nama.trim().toLowerCase() === nameClean);

  const cat = (category || '').toUpperCase();
  const kls = (kelas || '').toUpperCase();

  let rate = 50000;
  let labelTag = 'SAC';
  let detailClass = kls || 'Umum';

  if (tutorSetting) {
    if (cat.includes('ARIANNA') || kls.includes('ARIANNA')) {
      rate = tutorSetting.honorHybrid || tutorSetting.honorSAC || 75000;
      labelTag = 'Hybrid Arianna';
    } else if (cat.includes('PESANTREN') || kls.includes('PESANTREN')) {
      rate = tutorSetting.honorPesantren || tutorSetting.honorLOSSD || 75000;
      labelTag = 'LOS Pesantren';
      detailClass = 'Pesantren (Cibogo)';
    } else if (cat.includes('LOS SD') || kls.includes('LOS SD') || (kls.includes('SD') && cat.includes('LOS'))) {
      rate = tutorSetting.honorLOSSD || 97500;
      labelTag = 'LOS SD';
      detailClass = kls ? `SD (${kls})` : 'SD (Gendis/Aydin)';
    } else if (cat.includes('LOS SMP') || kls.includes('LOS SMP') || (kls.includes('SMP') && cat.includes('LOS'))) {
      rate = tutorSetting.honorLOSSMP || 107000;
      labelTag = 'LOS SMP';
      detailClass = kls ? `SMP (${kls})` : 'SMP (Kenzie)';
    } else if (cat.includes('LOS SMA') || kls.includes('LOS SMA') || (kls.includes('SMA') && cat.includes('LOS'))) {
      if (cat.includes('LUAR') || kls.includes('LUAR')) {
        rate = tutorSetting.honorLOSSMALK || 137500;
      } else {
        rate = tutorSetting.honorLOSSMADK || 125000;
      }
      labelTag = 'LOS SMA';
      detailClass = kls ? `SMA (${kls})` : 'SMA (Fadil)';
    } else if (cat.includes('ZAHRA') || kls.includes('ZAHRA') || cat.includes('LOS')) {
      rate = tutorSetting.honorLOSSD || tutorSetting.honorLOSSMP || 97500;
      labelTag = 'LOS SD/General';
      detailClass = kls ? `LOS (${kls})` : 'LOS';
    } else {
      rate = tutorSetting.honorSAC || 50000;
      labelTag = 'SAC';
    }
  }

  return { rate, labelTag, detailClass };
}

// ========================================================
// 1. PAGE BERANDA: PAYROLL TUTOR (DATAFORM + ATTENDANCE LOS)
// ========================================================

async function fetchPayrollData() {
  const tbody = document.getElementById('payroll-table-body');
  if (tbody) tbody.innerHTML = `<tr><td colspan="8" class="p-8 text-center text-slate-400">Memuat data payroll...</td></tr>`;

  try {
    const data = await fetchGVizJson(URL_DATAFORM);
    const rows = data.table.rows || [];

    // 1. Ambil data dari sheet DataForm
    const dataFormItems = rows.map(row => {
      const isoDate = parseToIsoDate(getCellValue(row, 0), getCellFormatted(row, 0));
      return {
        tanggal: isoDate,
        tutor: String(getCellValue(row, 1)).trim(),
        kategori: String(getCellValue(row, 2)).trim().toUpperCase(),
        kelas: String(getCellValue(row, 5)).trim(),
        source: 'DataForm'
      };
    }).filter(item => item.tutor !== '' && item.tutor !== 'Tanpa Nama' && item.tutor !== 'Nama Tutor');

    // 2. Ambil presensi LOS dari sheet Attendance
    const attendanceLosItems = [];
    rawAttendanceData.forEach(att => {
      const attText = `${att.status || ''} ${att.keterangan || ''}`.toUpperCase();
      const isLOS = attText.includes('LOS') || attText.includes('ZAHRA');

      if (isLOS && att.nama) {
        let categoryStr = 'LOS';
        if (attText.includes('HYBRID ZAHRA') || attText.includes('ZAHRA')) {
          categoryStr = 'HYBRID ZAHRA';
        } else if (attText.includes('LOS SD')) {
          categoryStr = 'LOS SD';
        } else if (attText.includes('LOS SMP')) {
          categoryStr = 'LOS SMP';
        } else if (attText.includes('LOS SMA')) {
          categoryStr = 'LOS SMA';
        } else if (attText.includes('PESANTREN')) {
          categoryStr = 'LOS PESANTREN';
        }

        attendanceLosItems.push({
          tanggal: att.tanggal,
          tutor: att.nama,
          kategori: categoryStr,
          kelas: att.keterangan || 'LOS',
          source: 'Attendance'
        });
      }
    });

    // 3. Gabungkan kedua sumber data (Deduplikasi agar jika tutor mencatat LOS pada tanggal sama di DataForm tidak terhitung ganda)
    const combinedData = [...dataFormItems];

    attendanceLosItems.forEach(attItem => {
      const isDuplicate = dataFormItems.some(dfItem => 
        dfItem.tutor.toLowerCase() === attItem.tutor.toLowerCase() &&
        dfItem.tanggal === attItem.tanggal &&
        (dfItem.kategori.includes('LOS') || dfItem.kategori.includes('ZAHRA'))
      );

      if (!isDuplicate) {
        combinedData.push(attItem);
      }
    });

    rawPayrollData = combinedData;

    processPayrollData();
  } catch (err) {
    console.error("Error Payroll:", err);
    if (tbody) tbody.innerHTML = `<tr><td colspan="8" class="p-8 text-center text-red-600 font-medium">⚠️ Gagal memuat data DataForm: ${err.message}</td></tr>`;
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
      tutorMap[nama] = {
        nama,
        datesSet: new Set(),
        countSAC: 0,
        countLOS: 0,
        totalHonor: 0,
        losClassCounts: {},
        tarifBreakdownMap: {},
        kegiatanNotes: []
      };
    }

    if (item.tanggal) {
      tutorMap[nama].datesSet.add(item.tanggal);
    }

    const info = calculateSessionInfo(nama, item.kategori, item.kelas);
    tutorMap[nama].totalHonor += info.rate;

    const isLOS = item.kategori.includes('LOS') || 
                  item.kategori.includes('ZAHRA') || 
                  item.kelas.toUpperCase().includes('ZAHRA') ||
                  item.kategori.includes('HYBRID ZAHRA');

    if (isLOS) {
      tutorMap[nama].countLOS += 1;
      const cLabel = info.detailClass;
      tutorMap[nama].losClassCounts[cLabel] = (tutorMap[nama].losClassCounts[cLabel] || 0) + 1;
    } else {
      tutorMap[nama].countSAC += 1;
    }

    tutorMap[nama].tarifBreakdownMap[info.labelTag] = info.rate;
  });

  processedPayrollSummary = Object.values(tutorMap).map(tutor => {
    const rincianParts = [];
    for (const [klsLabel, count] of Object.entries(tutor.losClassCounts)) {
      rincianParts.push(`${count}x ${klsLabel}`);
    }
    const rincianString = rincianParts.length > 0 ? rincianParts.join(', ') : '-';

    const breakdownTags = [];
    for (const [tag, rateVal] of Object.entries(tutor.tarifBreakdownMap)) {
      breakdownTags.push(`${tag} = ${formatNumberExcel(rateVal)}`);
    }

    return {
      nama: tutor.nama,
      hadir: tutor.datesSet.size || (tutor.countSAC + tutor.countLOS),
      countSAC: tutor.countSAC,
      countLOS: tutor.countLOS,
      rincian: rincianString,
      kegiatan: tutor.kegiatanNotes.length > 0 ? tutor.kegiatanNotes.join(', ') : '-',
      totalHonor: tutor.totalHonor,
      breakdownTags: breakdownTags
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
    tbody.innerHTML = `<tr><td colspan="8" class="p-8 text-center text-slate-400">Tidak ada data payroll pada rentang tanggal ini.</td></tr>`;
    return;
  }

  data.forEach((tutor, idx) => {
    const tr = document.createElement('tr');
    tr.className = 'hover:bg-slate-50 transition border-b border-slate-200';
    tr.innerHTML = `
      <td class="p-3 text-center font-semibold text-slate-500">${idx + 1}</td>
      <td class="p-3 font-bold text-slate-900">${tutor.nama}</td>
      <td class="p-3 text-center font-semibold text-slate-700">${tutor.hadir}</td>
      <td class="p-3 text-center font-semibold text-slate-700">${tutor.countSAC}</td>
      <td class="p-3 text-center font-semibold text-slate-700">${tutor.countLOS}</td>
      <td class="p-3 text-slate-600 text-[11px]">${tutor.rincian}</td>
      <td class="p-3 text-center text-slate-500 text-[11px]">${tutor.kegiatan}</td>
      <td class="p-3 text-right font-bold text-[#059669]">${formatRupiah(tutor.totalHonor)}</td>
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

// EKSPOR EXCEL FORMAT RESMI
function exportPayrollExcel() {
  if (!processedPayrollSummary || processedPayrollSummary.length === 0) {
    alert("Tidak ada data payroll untuk diekspor!");
    return;
  }

  const startDate = document.getElementById('start-date').value || 'Awal';
  const endDate = document.getElementById('end-date').value || 'Akhir';

  let totalSACAll = 0;
  let totalLOSAll = 0;
  let totalHonorAll = 0;

  let rowsHtml = '';
  processedPayrollSummary.forEach((tutor, idx) => {
    totalSACAll += tutor.countSAC;
    totalLOSAll += tutor.countLOS;
    totalHonorAll += tutor.totalHonor;

    let tagsCells = '';
    for (let i = 0; i < 5; i++) {
      const tagText = tutor.breakdownTags[i] || '';
      tagsCells += `<td style="font-size:8.5pt;">${tagText}</td>`;
    }

    rowsHtml += `
      <tr>
        <td style="text-align:center;">${idx + 1}</td>
        <td>${tutor.nama}</td>
        <td style="text-align:center;">${tutor.hadir}</td>
        <td style="text-align:center;">${tutor.countSAC}</td>
        <td style="text-align:center;">${tutor.countLOS}</td>
        <td style="font-size:8.5pt;">${tutor.rincian}</td>
        <td style="text-align:center; font-size:8.5pt;">${tutor.kegiatan}</td>
        <td style="text-align:right; font-weight:bold;">Rp&nbsp;${formatNumberExcel(tutor.totalHonor)}</td>
        <td style="border:none;"></td>
        ${tagsCells}
      </tr>
    `;
  });

  const fullExcelHtml = `
    <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
    <head>
      <meta charset="UTF-8">
      <!--[if gte mso 9]>
      <xml>
        <x:ExcelWorkbook>
          <x:ExcelWorksheets>
            <x:ExcelWorksheet>
              <x:Name>Laporan Honor</x:Name>
              <x:WorksheetOptions>
                <x:Print><x:ValidPrinterInfo/><x:PaperSizeIndex>9</x:PaperSizeIndex><x:Scale>70</x:Scale><x:FitToPage/></x:Print>
                <x:FitToPage/>
                <x:PageSetup>
                  <x:Layout x:Orientation="Landscape"/>
                  <x:PageMargins x:Header="0.3" x:Footer="0.3" x:Top="0.5" x:Bottom="0.5" x:Left="0.5" x:Right="0.5"/>
                </x:PageSetup>
              </x:WorksheetOptions>
            </x:ExcelWorksheet>
          </x:ExcelWorksheets>
        </x:ExcelWorkbook>
      </xml>
      <![endif]-->
      <style>
        @page { size: A4 landscape; margin: 0.5in; }
        body { font-family: Calibri, sans-serif; }
        table { border-collapse: collapse; margin: 0 auto; }
        th, td { border: 0.5pt solid #000; padding: 6px; font-size: 9.5pt; vertical-align: middle; }
        .th-header { background-color: #059669; color: #ffffff; font-weight: bold; text-align: center; }
      </style>
    </head>
    <body>
      <div style="width:1450px; margin:0 auto;">
        <h3 style="margin-bottom:4px; margin-left: 4px;">HONOR TUTOR 2026/2027</h3>
        <p style="margin-bottom:12px; margin-left: 4px;"><strong>Periode:</strong> ${startDate} s/d ${endDate}</p>
        <table style="width:1450px;">
          <colgroup>
            <col width="30"><col width="160"><col width="45"><col width="45"><col width="65"><col width="220"><col width="120"><col width="90"><col width="20"><col width="120"><col width="130"><col width="140"><col width="120"><col width="120">
          </colgroup>
          <thead>
            <tr>
              <th style="background:#059669; color:#fff;">No</th>
              <th style="background:#059669; color:#fff;">Nama Tutor</th>
              <th style="background:#059669; color:#fff;">Hadir</th>
              <th style="background:#059669; color:#fff;">SAC</th>
              <th style="background:#059669; color:#fff;">LOS Total</th>
              <th style="background:#059669; color:#fff;">Rincian Sesi Mengajar</th>
              <th style="background:#059669; color:#fff;">Kegiatan</th>
              <th style="background:#059669; color:#fff;">Total Honor</th>
              <th style="border:none;"></th>
              <th colspan="5" style="border:none;"></th>
            </tr>
          </thead>
          <tbody>
            ${rowsHtml}
            <tr style="background:#e2e8f0; font-weight:bold;">
              <td colspan="3" style="text-align:center;">TOTAL</td>
              <td style="text-align:center;">${totalSACAll}</td>
              <td style="text-align:center;">${totalLOSAll}</td>
              <td></td><td></td>
              <td style="text-align:right;">Rp&nbsp;${formatNumberExcel(totalHonorAll)}</td>
              <td style="border:none;"></td>
              <td colspan="5" style="border:none;"></td>
            </tr>
          </tbody>
        </table>
      </div>
    </body>
    </html>
  `;

  const blob = new Blob(['\uFEFF' + fullExcelHtml], { type: 'application/vnd.ms-excel;charset=utf-8' });
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = `Rekap_payroll_tutor_${startDate}_s_d_${endDate}.xls`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}

function exportPayrollCSV() {
  if (!processedPayrollSummary.length) return alert("Tidak ada data payroll untuk diekspor!");
  const startDate = document.getElementById('start-date').value || 'Awal';
  const endDate = document.getElementById('end-date').value || 'Akhir';

  let csv = "data:text/csv;charset=utf-8,\uFEFFNo,Nama Tutor,Hadir,SAC,LOS Total,Rincian Sesi Mengajar,Kegiatan,Total Honor,Breakdown Tarif\n";
  processedPayrollSummary.forEach((item, idx) => {
    csv += `${idx + 1},"${item.nama}",${item.hadir},${item.countSAC},${item.countLOS},"${item.rincian}","${item.kegiatan}",${item.totalHonor},"${item.breakdownTags.join('; ')}"\n`;
  });

  const link = document.createElement("a");
  link.setAttribute("href", encodeURI(csv));
  link.setAttribute("download", `Rekap_payroll_tutor_${startDate}_s_d_${endDate}.csv`);
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

      const col5 = String(getCellValue(row, 5)).trim();
      const col6 = String(getCellValue(row, 6)).trim();
      const col7 = String(getCellValue(row, 7)).trim();
      const col8 = String(getCellValue(row, 8)).trim();
      const col9 = String(getCellValue(row, 9)).trim();

      const combinedExtra = [col5, col6, col7, col8, col9].filter(Boolean).join(' ');

      return {
        id: getCellValue(row, 0),
        userId: userId,
        nama: userObj.nama,
        jabatan: userObj.jabatan,
        tanggal: isoDate,
        tanggalFmt: getCellFormatted(row, 2) || getCellValue(row, 2),
        jamMasuk: getCellValue(row, 3) || '-',
        jamKeluar: getCellValue(row, 4) || '-',
        keterangan: combinedExtra,
        status: col9 || 'Hadir',
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
        const honorLOSSD = parseFloat(getCellValue(row, 10)) || 97500;
        const honorLOSSMP = parseFloat(getCellValue(row, 11)) || 107000;
        const honorPesantren = parseFloat(getCellValue(row, 13)) || 75000;
        const honorHybrid = parseFloat(getCellValue(row, 14)) || 75000;
        const honorLOSSMADK = parseFloat(getCellValue(row, 15)) || 125000;
        const honorLOSSMALK = parseFloat(getCellValue(row, 16)) || 137500;

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
    honorLOSSD: 97500,
    honorLOSSMP: 107000,
    honorPesantren: 75000,
    honorLOSSMADK: 125000,
    honorLOSSMALK: 137500
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
  processPayrollData();
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