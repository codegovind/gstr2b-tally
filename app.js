document.getElementById('convertBtn').addEventListener('click', () => {
    const fileInput = document.getElementById('jsonInput');

    if (!fileInput.files.length) {
        showStatus("Please select a JSON file first.", "error");
        return;
    }

    const file = fileInput.files[0];
    const reader = new FileReader();

    reader.onload = function(e) {
        try {
            const data = JSON.parse(e.target.result);
            processData(data);
            showStatus("Excel generated successfully!", "success");
        } catch (error) {
            console.error(error);
            showStatus("Invalid JSON format or missing data.", "error");
        }
    };

    reader.readAsText(file);
});

function showStatus(message, type) {
    const statusText = document.getElementById('status');
    statusText.textContent = message;
    statusText.className = `status ${type}`;
}

function processData(jsonData) {
    const headers = [
        "Voucher Date", "Voucher Type Name", "Voucher Number",
        "Buyer/Supplier - Address", "Buyer/Supplier - Pincode",
        "Ledger Name", "Ledger Amount", "Ledger Amount Dr/Cr",
        "Item Name", "Billed Quantity", "Item Rate", "Item Rate per",
        "Item Amount", "Change Mode "
    ];

    let excelData = [headers];

    function extractAndAddTaxes(taxObject, addRow, drCr = "Dr") {
        const txval = taxObject.txval || 0;
        const cgst = taxObject.cgst || 0;
        const sgst = taxObject.sgst || 0;
        const igst = taxObject.igst || 0;
        const cess = taxObject.cess || 0;

        let sideTotal = 0;
        if (txval > 0) { addRow("Purchase A/c", txval, drCr); sideTotal += txval; }
        if (cgst > 0)  { addRow("CGST", cgst, drCr);          sideTotal += cgst; }
        if (sgst > 0)  { addRow("SGST", sgst, drCr);          sideTotal += sgst; }
        if (igst > 0)  { addRow("IGST", igst, drCr);          sideTotal += igst; }
        if (cess > 0)  { addRow("Cess", cess, drCr);          sideTotal += cess; }
        return sideTotal;
    }

    function addRoundOff(addRow, val, sideTotal, drCrOnTaxSide) {
        // Safety check: if val is 0, don't attempt to round off the entire invoice
        if (val === 0) return; 
        
        const diff = Math.round((Math.abs(val) - sideTotal) * 100) / 100;
        if (Math.abs(diff) >= 0.01) {
            const roundOffDrCr = diff > 0 ? drCrOnTaxSide : (drCrOnTaxSide === "Dr" ? "Cr" : "Dr");
            addRow("Round Off", Math.abs(diff), roundOffDrCr);
        }
    }

    // 2. Extract and Process B2B data
    const b2bData = jsonData?.data?.docdata?.b2b;
    if (b2bData && Array.isArray(b2bData)) {
        b2bData.forEach(supplier => {
            const ctin = supplier.ctin || "";
            const trdnm = supplier.trdnm || ctin;

            if (supplier.inv && Array.isArray(supplier.inv)) {
                supplier.inv.forEach(inv => {
                    const inum = inv.inum || "";
                    const idt = inv.dt || inv.idt || "";
                    const val = inv.val || 0;

                    const addRow = (ledgerName, amount, drCr) => {
                        excelData.push([
                            idt, "Purchase", inum, "", "",
                            ledgerName, amount, drCr,
                            "", "", "", "", "", "Accounting Invoice"
                        ]);
                    };

                    addRow(trdnm, val, "Cr");

                    let sideTotal = 0;
                    if (inv.itms && Array.isArray(inv.itms)) {
                        inv.itms.forEach(itm => {
                            sideTotal += extractAndAddTaxes(itm.itm_det || {}, addRow, "Dr");
                        });
                    } else {
                        sideTotal = extractAndAddTaxes(inv, addRow, "Dr");
                    }

                    // Fix JavaScript float precision before rounding
                    sideTotal = Math.round(sideTotal * 100) / 100;
                    addRoundOff(addRow, val, sideTotal, "Dr");
                });
            }
        });
    }

    // 3. Extract and Process CDNR data
    const cdnrData = jsonData?.data?.docdata?.cdnr;
    if (cdnrData && Array.isArray(cdnrData)) {
        cdnrData.forEach(supplier => {
            const ctin = supplier.ctin || "";
            const trdnm = supplier.trdnm || ctin;

            if (supplier.nt && Array.isArray(supplier.nt)) {
                supplier.nt.forEach(note => {
                    const ntnum = note.ntnum || "";
                    const idt = note.dt || note.idt || "";
                    const val = note.val || 0;

                    const typ = note.typ || "";
                    const isCreditNote = typ === "C";
                    const voucherType = isCreditNote ? "Debit Note" : "Credit Note";
                    const supplierDrCr = isCreditNote ? "Dr" : "Cr";
                    const taxDrCr = isCreditNote ? "Cr" : "Dr";

                    const addRow = (ledgerName, amount, drCr) => {
                        excelData.push([
                            idt, voucherType, ntnum, "", "",
                            ledgerName, amount, drCr,
                            "", "", "", "", "", "Accounting Invoice"
                        ]);
                    };

                    addRow(trdnm, val, supplierDrCr);

                    let sideTotal = 0;
                    if (note.itms && Array.isArray(note.itms)) {
                        note.itms.forEach(itm => {
                            sideTotal += extractAndAddTaxes(itm.itm_det || {}, addRow, taxDrCr);
                        });
                    } else {
                        sideTotal = extractAndAddTaxes(note, addRow, taxDrCr);
                    }

                    // Fix JavaScript float precision before rounding
                    sideTotal = Math.round(sideTotal * 100) / 100;
                    addRoundOff(addRow, val, sideTotal, taxDrCr);
                });
            }
        });
    }

    if (excelData.length === 1) {
        throw new Error("No valid B2B or CDNR data found to export.");
    }

    const worksheet = XLSX.utils.aoa_to_sheet(excelData);
    const workbook = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(workbook, worksheet, "Accounting Voucher");
    XLSX.writeFile(workbook, "Purchase_With_Notes.xlsx");
}
