document.getElementById('tahun-sekarang').textContent = new Date().getFullYear();
let cleanedPO = [];
let generatedASN = [];
let groupedByRak = [];

function getSmartSplit(partNo, orderQty, pallet) {
  if (pallet === 'B1') return [orderQty];
  const p = String(partNo).trim();
  if (p === '48820BZ060' && orderQty === 15) return [7, 8];
  if (p === '48820BZ080' && orderQty === 21) return [10, 11];
  if (p === '48820BZ130' && orderQty === 13) return [6, 7];
  if (p === '488200D040' && orderQty === 13) return [8, 5];

  let maxPerRak = 10;

  if (p === '488200D040') {
    maxPerRak = 8;
    } else if (p === '450460K060' || p === '450460K070') {
    maxPerRak = 20; 
  }
   if (p === 'PARTLAIN01') { maxPerRak = 5; }
  const splits = [];
  let rem = orderQty;
  while (rem > 0) {
    const take = Math.min(rem, maxPerRak);
    splits.push(take);
    rem -= take;
  }
  return splits;
}

document.getElementById('poFile').addEventListener('change', function(e) {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function(evt) {
    const data = new Uint8Array(evt.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(firstSheet, { defval: "" });

    const validRows = rows.filter(r => {
      const no = String(r['NO.'] || '').trim();
      return /^\d+$/.test(no) && r['P/O NUMBER'] && r['PARTS NO.'];
    });

    validRows.sort((a, b) => {
      const palA = String(a['PALLETIZE MARK']).trim();
      const palB = String(b['PALLETIZE MARK']).trim();
      if (palA !== palB) return palA.localeCompare(palB);
      return 0;
    });

    const caseB1 = document.getElementById('caseB1').value.trim();
    const caseD1 = document.getElementById('caseD1').value.trim();
    const caseD1Big = document.getElementById('caseD1Big').value.trim();

    cleanedPO = validRows.map(r => {
      const pallet = String(r['PALLETIZE MARK']).trim();
      const orderQty = parseInt(r['ORDER QTY']) || 0;
      const poNo = String(r['P/O NUMBER']).trim();

      let caseMark = '2';
      let caseNo = caseD1;
      if (pallet === 'B1') {
        caseMark = '1';
        caseNo = caseB1;
      } else if (orderQty >= 100 || !poNo.startsWith('SM')) {
        caseMark = '3';
        caseNo = caseD1Big;
      }

      return {
        vendor: String(r['VENDOR CODE']).trim(),
        poNo: poNo,
        itemNo: String(r['ITEM NO']).trim().padStart(4, '0'),
        partNo: String(r['PARTS NO.']).trim(),
        orderQty: orderQty,
        pallet: pallet,
        shipDate: String(r['DELIVERY SCHEDULE DATE']).trim(),
        caseMark: caseMark,
        caseNo: caseNo,
        splits: getSmartSplit(r['PARTS NO.'], orderQty, pallet)
      };
    });

    renderPOTable();
    buildAllData();
    document.getElementById('configSection').style.display = 'block';
  };
  reader.readAsArrayBuffer(file);
});

function runValidation() {
  if (cleanedPO.length === 0) return;

  const caseB1 = (document.getElementById('caseB1')?.value || '').trim().toUpperCase();
  const caseD1 = (document.getElementById('caseD1')?.value || '').trim().toUpperCase();
  const caseD1Big = (document.getElementById('caseD1Big')?.value || '').trim().toUpperCase();

  const errB1 = document.getElementById('errB1');
  const errD1 = document.getElementById('errD1');
  const errD1Big = document.getElementById('errD1Big');

  let isAllValid = true;
  if (caseB1 && !caseB1.startsWith('B')) {
    if (errB1) errB1.style.display = 'block';
    isAllValid = false;
  } else {
    if (errB1) errB1.style.display = 'none';
  }
  if (caseD1 && !caseD1.startsWith('D')) {
    if (errD1) errD1.style.display = 'block';
    isAllValid = false;
  } else {
    if (errD1) errD1.style.display = 'none';
  }
  if (caseD1Big && !caseD1Big.startsWith('D')) {
    if (errD1Big) errD1Big.style.display = 'block';
    isAllValid = false;
  } else {
    if (errD1Big) errD1Big.style.display = 'none';
  }
  cleanedPO.forEach(item => {
    if (item.pallet === 'B1') item.caseNo = caseB1;
    else if (item.caseMark === '3') item.caseNo = caseD1Big;
    else item.caseNo = caseD1;

    // Cek selisih kuantitas per baris
    const totalSplit = item.splits.reduce((a, b) => a + b, 0);
    if (totalSplit !== item.orderQty) {
      item.qtyError = true;
      isAllValid = false;
    } else {
      item.qtyError = false;
    }
  });
  const btnPrint = document.getElementById('btnPrint');
  const btnDownload = document.getElementById('btnDownload');
  if (btnPrint) btnPrint.disabled = !isAllValid;
  if (btnDownload) btnDownload.disabled = !isAllValid;
  renderPOTable();
  if (isAllValid) {
    document.getElementById('alertTopArea').innerHTML = ''; 
    buildAllData();
  } else {
    const alertTop = document.getElementById('alertTopArea');
    if (alertTop) {
      alertTop.innerHTML = '<div style="color:#dc2626; text-align:center; font-weight:bold; padding:12px; background:#fee2e2; border: 2px solid #f87171; border-radius:8px; margin-bottom: 14px;">⚠️ Case No Masih Salah! Silakan betulkan dulu cantik. 😍</div>';
    }
    const partArea = document.getElementById('partContentArea');
    if (partArea) {
      partArea.innerHTML = '';
    }
  }
}
function renderPOTable() {
  const tbody = document.getElementById('poTableBody');
  tbody.innerHTML = '';
  cleanedPO.forEach((item, idx) => {
    const isB1 = item.pallet === 'B1';
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${idx + 1}</td>
      <td><b>${item.pallet}</b></td>
      <td>${item.caseMark}</td>
      <td><b>${item.caseNo}</b></td>
      <td>${item.vendor}</td>
      <td>${item.poNo}</td>
      <td>${item.itemNo}</td>
      <td><b>${item.partNo}</b></td>
      <td><b>${item.orderQty}</b></td>
      <td>${isB1 ? '🔒 Gabung 1 tempat' : 'Pecah tempat'}</td>
      <td>
        <input type="text" class="split-input" value="${item.splits.join(', ')}" 
          ${isB1 ? 'readonly style="background:#fef9c3;font-weight:bold;"' : ''} 
          onchange="updateSplit(${idx}, this.value)">
      </td>
    `;
    tbody.appendChild(tr);
  });
}

function updateSplit(idx, val) {
  const nums = val.split(',').map(s => parseInt(s.trim())).filter(n => !isNaN(n) && n > 0);
  if (nums.length > 0) {
    cleanedPO[idx].splits = nums;
    buildAllData();
  }
}

function buildAllData() {
  generatedASN = [];
  const rakMap = new Map();
  const hasB1 = cleanedPO.some(i => i.pallet === 'B1');
  let currentD1Rak = hasB1 ? 1 : 0;

  cleanedPO.forEach(item => {
    if (item.pallet === 'B1') {
      item.splits.forEach(qty => {
        const rowObj = {
          NO_RAK: "1",
          SHIPMENT_DATE: item.shipDate,
          CASE_NO: item.caseNo,
          CASE_PALLETIZE_MARK: "1",
          PO_NO: item.poNo,
          PO_DATE: "",
          ITEM_NO: item.itemNo,
          FRANCHISE: "",
          PART_NO: item.partNo,
          DELIVERY_QUANTITY: qty,
          VENDOR_CODE: item.vendor,
          PALLET: item.pallet
        };
        generatedASN.push(rowObj);
        if (!rakMap.has("1")) rakMap.set("1", []);
        rakMap.get("1").push(rowObj);
      });
    } else {
      item.splits.forEach(qty => {
        currentD1Rak++;
        const rakStr = String(currentD1Rak);
        const rowObj = {
          NO_RAK: rakStr,
          SHIPMENT_DATE: item.shipDate,
          CASE_NO: item.caseNo,
          CASE_PALLETIZE_MARK: String(item.caseMark),
          PO_NO: item.poNo,
          PO_DATE: "",
          ITEM_NO: item.itemNo,
          FRANCHISE: "",
          PART_NO: item.partNo,
          DELIVERY_QUANTITY: qty,
          VENDOR_CODE: item.vendor,
          PALLET: item.pallet
        };
        generatedASN.push(rowObj);
        rakMap.set(rakStr, [rowObj]);
      });
    }
  });

  groupedByRak = Array.from(rakMap.entries());
  renderPartContentSlips();

  const totalQty = generatedASN.reduce((acc, r) => acc + r.DELIVERY_QUANTITY, 0);
  const sumBar = document.getElementById('summaryBar');
  sumBar.style.display = 'block';
  sumBar.innerText = `✅ Total Baris ASN: ${generatedASN.length} Baris | Total Label PART CONTENT Siap Print: ${groupedByRak.length} | Total Qty: ${totalQty} Pcs`;
}
function renderPartContentSlips() {
  const container = document.getElementById('partContentArea');
  container.innerHTML = '';
  groupedByRak.forEach(([rakNo, rows]) => {
    const isMultiRow = rows.length > 1 || rows[0].PALLET === 'B1';
    let rowsHtml = '';

    // 1 garis penanda part gabung
    // if (isMultiRow) {
    //   rowsHtml += `
    //     <tr class="pc-row">
    //       <td class="rak-box"></td>
    //       <td class="bordered"></td>
    //       <td class="bordered"></td>
    //       <td class="bordered"></td>
    //       <td class="bordered"></td>
    //       <td class="bordered"></td>
    //       <td class="bordered"></td>
    //       <td class="bordered"></td>
    //       <td class="bordered"></td>
    //       <td class="bordered"></td>
    //       <td class="bordered"></td>
    //       <td class="bordered"></td>
    //     </tr>
    //   `;
    // }

    rows.forEach(r => {
      rowsHtml += `
        <tr class="pc-row">
          <td class="${isMultiRow ? 'rak-box' : 'col-rak'}">${r.NO_RAK}</td>
          <td class="bordered">${r.SHIPMENT_DATE}</td>
          <td class="bordered" style="text-align:left;">${r.CASE_NO}</td>
          <td class="bordered" style="text-align:left;">${r.CASE_PALLETIZE_MARK}</td>
          <td class="bordered" style="text-align:left;">${r.PO_NO}</td>
          <td class="bordered">${r.PO_DATE}</td>
          <td class="bordered" style="text-align:left;">${r.ITEM_NO}</td>
          <td class="bordered">${r.FRANCHISE}</td>
          <td class="bordered" style="text-align:left;">${r.PART_NO}</td>
          <td class="bordered" style="text-align:right;">${r.DELIVERY_QUANTITY}</td>
          <td class="bordered" style="text-align:left;">${r.VENDOR_CODE}</td>
          <td class="${isMultiRow ? 'bordered' : 'col-pal'}" style="text-align:left;">${r.PALLET}</td>
        </tr>
      `;
    });

    const card = document.createElement('div');
    card.className = 'pc-card';
    card.innerHTML = `
      <table class="pc-table">
        <tr>
          <td class="col-rak"></td>
          <td colspan="10" class="pc-title">PART CONTENT</td>
          <td class="col-pal"></td>
        </tr>
        <tr><td colspan="12" class="pc-spacer"></td></tr>
        <tr>
          <td class="col-rak"></td>
          <td class="pc-hdr-ylw">SHIPMENT<br>DATE</td>
          <td class="pc-hdr-ylw">CASE NO</td>
          <td class="pc-hdr-wht">CASE<br>PALLETIZE<br>MARK</td>
          <td class="pc-hdr-ylw">PO NO</td>
          <td class="pc-hdr-wht">PO DATE</td>
          <td class="pc-hdr-ylw">ITEM NO</td>
          <td class="pc-hdr-wht">FRANCHISE</td>
          <td class="pc-hdr-ylw">PART NO</td>
          <td class="pc-hdr-ylw">DELIVERY<br>QUANTITY</td>
          <td class="pc-hdr-ylw">VENDOR<br>CODE</td>
          <td class="col-pal"></td>
        </tr>
        ${rowsHtml}
      </table>
    `;
    container.appendChild(card);
  });
}

function downloadASNExcel() {
  if (generatedASN.length === 0) return alert("Belum ada data! Upload file dulu.");
  const wsData = [];
  const merges = [];
  wsData.push([
    "DELIVERY DATE", "NO RAK", "SHIPMENT DATE", "CASE NO", "CASE PALLETIZE MARK",
    "PO NO", "PO DATE", "ITEM NO", "FRANCHISE", "PART NO", "DELIVERY QUANTITY",
    "VENDOR CODE", "PALLET", "PREPARE", "ADMIN"
  ]);
  generatedASN.forEach(r => {
    wsData.push([
      "", r.NO_RAK, r.SHIPMENT_DATE, r.CASE_NO, r.CASE_PALLETIZE_MARK,
      r.PO_NO, "", r.ITEM_NO, "", r.PART_NO, r.DELIVERY_QUANTITY,
      r.VENDOR_CODE, r.PALLET, "", ""
    ]);
  });
  groupedByRak.forEach(([rakNo, rows]) => {
    wsData.push([]); // Baris kosong pemisah
    const titleRowIdx = wsData.length;
    wsData.push(["", "", "PART CONTENT", "", "", "", "", "", "", "", "", "", ""]);
    merges.push({ s: { r: titleRowIdx, c: 2 }, e: { r: titleRowIdx, c: 11 } });

    wsData.push([]); // Baris spasi bawah judul
    wsData.push([
      "", "", "SHIPMENT DATE", "CASE NO", "CASE PALLETIZE MARK",
      "PO NO", "PO DATE", "ITEM NO", "FRANCHISE", "PART NO",
      "DELIVERY QUANTITY", "VENDOR CODE", ""
    ]);
    // 1 garis penanda part gabung
    // if (rows.length > 1 || rows[0].PALLET === 'B1') {
    //   wsData.push(["", "", "", "", "", "", "", "", "", "", "", "", ""]);
    // }

    rows.forEach(r => {
      wsData.push([
        "", r.NO_RAK, r.SHIPMENT_DATE, r.CASE_NO, r.CASE_PALLETIZE_MARK,
        r.PO_NO, "", r.ITEM_NO, "", r.PART_NO, r.DELIVERY_QUANTITY,
        r.VENDOR_CODE, r.PALLET
      ]);
    });
  });

  const ws = XLSX.utils.aoa_to_sheet(wsData);
  ws['!merges'] = merges;
  ws['!cols'] = [
    { wch: 14 }, { wch: 8 }, { wch: 15 }, { wch: 16 }, { wch: 21 },
    { wch: 15 }, { wch: 14 }, { wch: 12 }, { wch: 14 }, { wch: 19 },
    { wch: 19 }, { wch: 15 }, { wch: 8 }, { wch: 10 }, { wch: 10 }
  ];
  const borderAll = {
    top: { style: "thin", color: { rgb: "000000" } },
    bottom: { style: "thin", color: { rgb: "000000" } },
    left: { style: "thin", color: { rgb: "000000" } },
    right: { style: "thin", color: { rgb: "000000" } }
  };
  const yellowCols = [2, 3, 5, 7, 9, 10, 11]; 

  for (let R = 0; R < wsData.length; R++) {
    const row = wsData[R];
    if (!row || row.length === 0) continue; 

    // A. BIKIN GARIS UNTUK TABEL UTAMA ASN (ATAS)
    if (R <= generatedASN.length) {
      for (let C = 0; C < 15; C++) {
        const ref = XLSX.utils.encode_cell({ r: R, c: C });
        if (!ws[ref]) ws[ref] = { t: 's', v: '' };
        
        ws[ref].s = {
          font: { name: "Calibri", sz: R === 0 ? 12 : 14, bold: R === 0 },
          alignment: { horizontal: "center", vertical: "center", wrapText: true },
          border: borderAll 
        };
        if (R === 0 && yellowCols.includes(C)) {
          ws[ref].s.fill = { fgColor: { rgb: "FFFF00" } };
        }
      }
    }
    else {
      if (row[2] === "PART CONTENT") {
        for (let C = 2; C <= 11; C++) {
          const ref = XLSX.utils.encode_cell({ r: R, c: C });
          if (!ws[ref]) ws[ref] = { t: 's', v: '' };
          ws[ref].s = {
            font: { name: "Times New Roman", sz: 24 },
            alignment: { horizontal: "center", vertical: "center" },
            border: borderAll
          };
        }
      }
      else if (row[2] === "SHIPMENT DATE") {
        for (let C = 2; C <= 11; C++) {
          const ref = XLSX.utils.encode_cell({ r: R, c: C });
          if (!ws[ref]) continue;
          const isYlw = yellowCols.includes(C);
          ws[ref].s = {
            font: { name: "Calibri", sz: 13, bold: isYlw },
            fill: isYlw ? { fgColor: { rgb: "FFFF00" } } : undefined,
            alignment: { horizontal: "center", vertical: "center", wrapText: true },
            border: borderAll
          };
        }
      }
      else if (R > generatedASN.length + 1 && (row[2] !== undefined || row[1] === "")) {
        for (let C = 1; C <= 12; C++) {
          const ref = XLSX.utils.encode_cell({ r: R, c: C });
          if (!ws[ref]) ws[ref] = { t: 's', v: '' };
          ws[ref].s = {
            font: { name: "Calibri", sz: 18 },
            alignment: { horizontal: "center", vertical: "center" },
            border: (C >= 2 && C <= 11) || row[12] === 'B1' || (C === 1 && row[1]) ? borderAll : undefined
          };
        }
      }
    }
  }

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "ASN TEMPLATE");
  const shipDate = generatedASN[0]?.SHIPMENT_DATE || "EXPORT";
  XLSX.writeFile(wb, `ASN_Template_${shipDate}_PPICSMC.xlsx`);
}
