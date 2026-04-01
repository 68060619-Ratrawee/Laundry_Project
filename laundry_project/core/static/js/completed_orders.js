function showItemsModal(billId, items, qty, unit, price) {
    document.getElementById('modalBillId').innerText = billId;
    document.getElementById('modalTotalQty').innerText = qty;
    document.getElementById('modalUnit').innerText = unit;
    document.getElementById('modalTotalPrice').innerText = price;
    
    const itemListContainer = document.getElementById('modalItemList');
    
    const totalQty = parseInt(qty) || 1;
    const totalPrice = parseFloat(price) || 0;
    const unitPrice = totalPrice / totalQty;

    const itemArray = items.split(',').map(i => i.trim()).filter(i => i !== "");
    const itemCounts = {};
    itemArray.forEach(i => { itemCounts[i] = (itemCounts[i] || 0) + 1; });
    
    let htmlContent = '<div class="list-group list-group-flush border rounded-3 overflow-hidden">';
    
    Object.keys(itemCounts).forEach(itemName => {
        const itemQty = itemCounts[itemName];
        const itemLinePrice = (itemQty * unitPrice).toFixed(2);
        
        htmlContent += `
            <div class="list-group-item d-flex justify-content-between align-items-center py-2 bg-white">
                <span class="text-secondary small"><i class="fa-solid fa-check text-success me-2"></i>${itemName}</span>
                <div>
                    <span class="badge bg-light text-primary border border-primary-subtle fw-bold me-2">x ${itemQty} ${unit}</span>
                    <span class="text-dark fw-bold small">฿${itemLinePrice}</span>
                </div>
            </div>`;
    });
    
    htmlContent += '</div>';
    itemListContainer.innerHTML = htmlContent;
    var myModal = new bootstrap.Modal(document.getElementById('itemsModal'));
    myModal.show();
}

document.addEventListener("DOMContentLoaded", function() {
    const fBill = document.getElementById('filterBill');
    const fCus = document.getElementById('filterCustomer');
    const fStart = document.getElementById('filterStartDate');
    const fEnd = document.getElementById('filterEndDate');
    const pSize = document.getElementById('pageSize');
    const searchBtn = document.getElementById('searchBtn'); 
    const pgControls = document.getElementById('paginationControls');
    const noMatch = document.getElementById('noMatchRow');
    
    const allRows = Array.from(document.querySelectorAll('#historyTableBody tr.order-row'));
    let currentPage = 1;

    function renderTable() {
        const billV = fBill.value.toLowerCase();
        const cusV = fCus.value.toLowerCase();
        const startV = fStart.value;
        const endV = fEnd.value;
        const size = parseInt(pSize.value);

        const filteredRows = allRows.filter(row => {
            const billText = row.querySelector('.col-bill').textContent.toLowerCase();
            const cusName = row.querySelector('.cus-name').textContent.toLowerCase();
            const rowStart = row.getAttribute('data-start-date');
            const rowEnd = row.getAttribute('data-end-date');

            return billText.includes(billV) && cusName.includes(cusV) && 
                   (startV === "" || rowStart === startV) && (endV === "" || rowEnd === endV);
        });

        const total = filteredRows.length;
        const totalP = Math.ceil(total / size) || 1;
        if (currentPage > totalP) currentPage = totalP;

        const startIdx = (currentPage - 1) * size;
        const endIdx = startIdx + size;

        allRows.forEach(row => row.style.display = 'none');
        filteredRows.slice(startIdx, endIdx).forEach(row => row.style.display = '');

        if (total === 0 && allRows.length > 0) {
            noMatch.style.display = '';
        } else {
            noMatch.style.display = 'none';
        }

        const startDisp = total > 0 ? startIdx + 1 : 0;
        const endDisp = Math.min(endIdx, total);
        document.getElementById('pageInfo').textContent = `แสดงที่ ${startDisp} ถึง ${endDisp} จากทั้งหมด ${total} รายการ`;
        
        renderPagination(totalP);
    }

    function renderPagination(totalP) {
        pgControls.innerHTML = '';
        const createBtn = (label, target, active = false, disabled = false) => {
            const li = document.createElement('li');
            li.className = `page-item ${active ? 'active' : ''} ${disabled ? 'disabled' : ''}`;
            const a = document.createElement('a');
            a.className = 'page-link shadow-none';
            a.href = 'javascript:void(0);';
            a.textContent = label;
            if (!disabled) {
                a.onclick = () => { currentPage = target; renderTable(); window.scrollTo({top: 0, behavior: 'smooth'}); };
            }
            li.appendChild(a);
            return li;
        };

        pgControls.appendChild(createBtn('ก่อนหน้า', currentPage - 1, false, currentPage === 1));
        for (let i = 1; i <= totalP; i++) {
            pgControls.appendChild(createBtn(i, i, currentPage === i));
        }
        pgControls.appendChild(createBtn('ถัดไป', currentPage + 1, false, currentPage === totalP || totalP === 0));
    }

    searchBtn.addEventListener('click', () => { 
        currentPage = 1; 
        renderTable(); 
    });

    pSize.addEventListener('change', () => { 
        currentPage = 1; 
        renderTable(); 
    });

    renderTable();
});

async function exportToExcel() {
    const fBill = document.getElementById('filterBill').value.toLowerCase();
    const fCus = document.getElementById('filterCustomer').value.toLowerCase();
    const fStart = document.getElementById('filterStartDate').value;
    const fEnd = document.getElementById('filterEndDate').value;
    
    const allRows = Array.from(document.querySelectorAll('#historyTableBody tr.order-row'));
    
    const exportRows = allRows.filter(row => {
        const billText = row.querySelector('.col-bill').textContent.toLowerCase();
        const cusName = row.querySelector('.cus-name').textContent.toLowerCase();
        const rowStart = row.getAttribute('data-start-date');
        const rowEnd = row.getAttribute('data-end-date');

        return billText.includes(fBill) && cusName.includes(fCus) && 
               (fStart === "" || rowStart === fStart) && (fEnd === "" || rowEnd === fEnd);
    });

    if (exportRows.length === 0) {
        Swal.fire({ icon: 'warning', title: 'ไม่มีข้อมูล', text: 'ไม่พบข้อมูลสำหรับการ Export', confirmButtonColor: '#0d6efd' });
        return;
    }

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('รายงานประวัติงานที่เสร็จสิ้น');

    worksheet.columns = [
        { key: 'billId', width: 20 },
        { key: 'date', width: 25 },
        { key: 'completedDate', width: 25 }, 
        { key: 'customer', width: 35 },
        { key: 'item', width: 45 },
        { key: 'qty', width: 20 },
        { key: 'price', width: 20 }
    ];

    worksheet.mergeCells('A1:G1');
    const titleRow = worksheet.getRow(1);
    titleRow.getCell(1).value = 'ประวัติงานที่เสร็จสิ้น (พร้อมรายได้สุทธิ)';
    titleRow.getCell(1).font = { size: 16, bold: true };
    titleRow.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };

    worksheet.mergeCells('A2:G2');
    const subtitleRow = worksheet.getRow(2);
    const now = new Date();
    const formattedDate = now.toLocaleDateString('th-TH') + ' เวลา ' + now.toLocaleTimeString('th-TH');
    subtitleRow.getCell(1).value = `Export ณ วันที่: ${formattedDate}`;
    subtitleRow.getCell(1).font = { size: 16 };
    subtitleRow.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };

    const headerRow = worksheet.getRow(3);
    headerRow.values = ['รหัสบิล', 'วันที่รับออเดอร์', 'วันที่สิ้นสุด', 'ชื่อลูกค้า', 'รายการผ้า', 'จำนวน/ตัว', 'รวมสุทธิ'];
    for (let i = 1; i <= 7; i++) {
        let cell = headerRow.getCell(i);
        cell.font = { size: 16, bold: true, color: { argb: 'FFFFFFFF' } }; 
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0D6EFD' } }; 
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
    }

    let grandTotal = 0;

    exportRows.forEach((row, index) => {
        const billId = row.querySelector('.col-bill').textContent.trim();
        const orderDate = row.querySelector('td:nth-child(2)').textContent.trim();
        const completedDate = row.querySelector('td:nth-child(6)').textContent.trim(); 
        const cusName = row.querySelector('.cus-name').textContent.trim();
        
        const itemsStr = row.getAttribute('data-items') || '';
        const totalQty = parseInt(row.getAttribute('data-qty') || '0');
        const price = parseFloat(row.getAttribute('data-price') || '0.00');

        const unitPrice = totalQty > 0 ? (price / totalQty) : 0;

        grandTotal += price;

        const itemArray = itemsStr.split(',').map(i => i.trim()).filter(i => i !== "");
        const itemCounts = {};
        itemArray.forEach(i => { itemCounts[i] = (itemCounts[i] || 0) + 1; });
        const uniqueItems = Object.keys(itemCounts);

        if (uniqueItems.length === 1) itemCounts[uniqueItems[0]] = totalQty;

        const bgColor = (index % 2 === 0) ? 'FFF8F9FA' : 'FFFFFFFF'; 

        if (uniqueItems.length === 0) {
            let newRow = worksheet.addRow({ billId: billId, date: orderDate, completedDate: completedDate, customer: cusName, item: '-', qty: totalQty, price: price });
            styleDataRow(newRow, bgColor);
        } else {
            uniqueItems.forEach((itemName, i) => {
                const itemQty = itemCounts[itemName];
                const itemLinePrice = itemQty * unitPrice; 

                let newRow;
                if (i === 0) {
                    newRow = worksheet.addRow({ billId: billId, date: orderDate, completedDate: completedDate, customer: cusName, item: itemName, qty: itemQty, price: itemLinePrice });
                } else {
                    newRow = worksheet.addRow({ billId: '', date: '', completedDate: '', customer: '', item: itemName, qty: itemQty, price: itemLinePrice });
                }
                styleDataRow(newRow, bgColor);
            });
        }
    });

    let totalRow = worksheet.addRow({ billId: '', date: '', completedDate: '', customer: '', item: '', qty: 'ยอดรวมสุทธิทั้งหมด', price: grandTotal });
    for (let i = 1; i <= 7; i++) {
        let cell = totalRow.getCell(i);
        cell.font = { size: 16, bold: true, color: { argb: 'FF000000' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC107' } }; 
        cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
    }
    totalRow.getCell(6).alignment = { horizontal: 'right', vertical: 'middle' };
    totalRow.getCell(7).numFmt = '#,##0.00';
    totalRow.getCell(7).alignment = { horizontal: 'right', vertical: 'middle' };

    const buffer = await workbook.xlsx.writeBuffer();
    const dateStr = new Date().toISOString().slice(0,10);
    saveAs(new Blob([buffer]), `ประวัติงานที่เสร็จสิ้น (พร้อมรายได้สุทธิ)_${dateStr}.xlsx`);
}

function styleDataRow(row, bgColor) {
    for (let i = 1; i <= 7; i++) {
        let cell = row.getCell(i);
        cell.font = { size: 16 }; 
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgColor } };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
    }
    row.getCell(4).alignment = { horizontal: 'left', vertical: 'middle' }; 
    row.getCell(5).alignment = { horizontal: 'left', vertical: 'middle' }; 
    
    const priceCell = row.getCell(7);
    if(priceCell.value !== '') {
        priceCell.numFmt = '#,##0.00';
        priceCell.alignment = { horizontal: 'right', vertical: 'middle' }; 
    }
}