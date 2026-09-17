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
            
            // Clear the file input field after successful download
            fileInput.value = ''; 
            
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

function formatTallyDate(dateStr) {
    if (!dateStr) return "";
    const parts = dateStr.split('-');
    if (parts.length === 3) {
        const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
        return `${parts[0]}-${months[parseInt(parts[1], 10) - 1]}-${parts[2]}`;
    }
    return dateStr;
}

function processData(jsonData) {
    // NEW: Dictionary to convert POS numbers into Tally State Names
    const gstStates = {
        "01": "Jammu & Kashmir", "02": "Himachal Pradesh", "03": "Punjab", "04": "Chandigarh",
        "05": "Uttarakhand", "06": "Haryana", "07": "Delhi", "08": "Rajasthan",
        "09": "Uttar Pradesh", "10": "Bihar", "11": "Sikkim", "12": "Arunachal Pradesh",
        "13": "Nagaland", "14": "Manipur", "15": "Mizoram", "16": "Tripura",
        "17": "Meghalaya", "18": "Assam", "19": "West Bengal", "20": "Jharkhand",
        "21": "Odisha", "22": "Chhattisgarh", "23": "Madhya Pradesh", "24": "Gujarat",
        "25": "Daman & Diu", "26": "Dadra & Nagar Haveli and Daman & Diu", "27": "Maharashtra",
        "28": "Andhra Pradesh", "29": "Karnataka", "30": "Goa", "31": "Lakshadweep",
        "32": "Kerala", "33": "Tamil Nadu", "34": "Puducherry", "35": "Andaman & Nicobar Islands",
        "36": "Telangana", "37": "Andhra Pradesh", "38": "Ladakh"
    };

    const headers = [
        "Voucher Date", "Voucher Type Name", "Voucher Number", "GSTIN/UIN", "Place of Supply",
        "Supplier Invoice Date", "Narration", 
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

        let rate = taxObject.rt || 0;
        if (!rate && txval > 0) {
            rate = Math.round(((cgst + sgst + igst) / txval) * 100);
        }
        
        const standardRates = [0, 5, 12, 18, 28];
        rate = standardRates.reduce((prev, curr) => Math.abs(curr - rate) < Math.abs(prev - rate) ? curr : prev);

        let purLedger = `PURCHASE ${rate}%`;
        let cgstLedger = `INPUT CGST ${rate/2}%`; 
        let sgstLedger = `INPUT SGST ${rate/2}%`;
        let igstLedger = `INPUT IGST ${rate}%`; 

        if (rate === 18) {
            purLedger = "PURCHASE 18%";
            cgstLedger = "INPUT CSGT 9 %"; 
            sgstLedger = "INPUT SGST 9 %";
        } else if (rate === 28) {
            purLedger = "PURCHASE 28 %";
            cgstLedger = "INPUT CGST 14 %";
            sgstLedger = "INPUT SGST 14 %";
        } else if (rate === 5) {
            purLedger = "PURCHASE 5%";
            cgstLedger = "INPUT CGST 2.5 %";
            sgstLedger = "INPUT SGST 2.5 %";
        } else if (rate === 0) {
            purLedger = "PURCHASE 0%";
        }

        let sideTotal = 0;
        if (txval > 0) { addRow(purLedger, txval, drCr); sideTotal += txval; }
        if (cgst > 0)  { addRow(cgstLedger, cgst, drCr);  sideTotal += cgst; }
        if (sgst > 0)  { addRow(sgstLedger, sgst, drCr);  sideTotal += sgst; }
        if (igst > 0)  { addRow(igstLedger, igst, drCr);  sideTotal += igst; }
        if (cess > 0)  { addRow("Cess", cess, drCr);      sideTotal += cess; }
        
        return sideTotal;
    }

    function addRoundOff(addRow, val, sideTotal, drCrOnTaxSide) {
        if (val === 0) return; 
        
        const diff = Math.round((Math.abs(val) - sideTotal) * 100) / 100;
        if (Math.abs(diff) >= 0.01) {
            const roundOffDrCr = diff > 0 ? drCrOnTaxSide : (drCrOnTaxSide === "Dr" ? "Cr" : "Dr");
            addRow("Round Off", Math.abs(diff), roundOffDrCr);
        }
    }

    const b2bData = jsonData?.data?.docdata?.b2b;
    if (b2bData && Array.isArray(b2bData)) {
        b2bData.forEach(supplier => {
            const ctin = supplier.ctin || "";
            const trdnm = supplier.trdnm || ctin;
            const filingDate = supplier.supfildt || "";
            const period = supplier.supprd || "";

            if (supplier.inv && Array.isArray(supplier.inv)) {
                supplier.inv.forEach(inv => {
                    const inum = inv.inum || "";
                    const idt = formatTallyDate(inv.dt || inv.idt || "");
                    const val = inv.val || 0;
                    
                    // UPDATED: Swap "27" for "Maharashtra"
                    const rawPos = inv.pos || "";
                    const posName = gstStates[rawPos] || rawPos;
                    
                    const rcm = inv.rev === "Y" ? "Yes" : "No";
                    const itc = inv.itcavl === "Y" ? "Eligible" : "Ineligible";
                    const autoNarration = `GSTR-2B Import | Filed: ${filingDate} | Period: ${period} | RCM: ${rcm} | ITC: ${itc}`;

                    const addRow = (ledgerName, amount, drCr, gstin = "", posVal = "", supInvDate = "", narr = "") => {
                        excelData.push([
                            idt, "Purchase", inum, gstin, posVal, 
                            supInvDate, narr, 
                            "", "",
                            ledgerName, amount, drCr,
                            "", "", "", "", "", "Accounting Invoice"
                        ]);
                    };

                    addRow(trdnm, val, "Cr", ctin, posName, idt, autoNarration);

                    let sideTotal = 0;
                    if (inv.itms && Array.isArray(inv.itms)) {
                        inv.itms.forEach(itm => {
                            sideTotal += extractAndAddTaxes(itm.itm_det || {}, addRow, "Dr");
                        });
                    } else {
                        sideTotal = extractAndAddTaxes(inv, addRow, "Dr");
                    }

                    sideTotal = Math.round(sideTotal * 100) / 100;
                    addRoundOff(addRow, val, sideTotal, "Dr");
                });
            }
        });
    }

    const cdnrData = jsonData?.data?.docdata?.cdnr;
    if (cdnrData && Array.isArray(cdnrData)) {
        cdnrData.forEach(supplier => {
            const ctin = supplier.ctin || "";
            const trdnm = supplier.trdnm || ctin;
            const filingDate = supplier.supfildt || "";
            const period = supplier.supprd || "";

            if (supplier.nt && Array.isArray(supplier.nt)) {
                supplier.nt.forEach(note => {
                    const ntnum = note.ntnum || "";
                    const idt = formatTallyDate(note.dt || note.idt || "");
                    const val = note.val || 0;
                    
                    // UPDATED: Swap "27" for "Maharashtra"
                    const rawPos = note.pos || "";
                    const posName = gstStates[rawPos] || rawPos;

                    const typ = note.typ || "";
                    const isCreditNote = typ === "C";
                    const voucherType = isCreditNote ? "Debit Note" : "Credit Note";
                    const supplierDrCr = isCreditNote ? "Dr" : "Cr";
                    const taxDrCr = isCreditNote ? "Cr" : "Dr";

                    const rcm = note.rev === "Y" ? "Yes" : "No";
                    const itc = note.itcavl === "Y" ? "Eligible" : "Ineligible";
                    const autoNarration = `GSTR-2B Note | Filed: ${filingDate} | Period: ${period} | RCM: ${rcm} | ITC: ${itc}`;

                    const addRow = (ledgerName, amount, drCr, gstin = "", posVal = "", supInvDate = "", narr = "") => {
                        excelData.push([
                            idt, voucherType, ntnum, gstin, posVal, 
                            supInvDate, narr,
                            "", "",
                            ledgerName, amount, drCr,
                            "", "", "", "", "", "Accounting Invoice"
                        ]);
                    };

                    addRow(trdnm, val, supplierDrCr, ctin, posName, idt, autoNarration);

                    let sideTotal = 0;
                    if (note.itms && Array.isArray(note.itms)) {
                        note.itms.forEach(itm => {
                            sideTotal += extractAndAddTaxes(itm.itm_det || {}, addRow, taxDrCr);
                        });
                    } else {
                        sideTotal = extractAndAddTaxes(note, addRow, taxDrCr);
                    }

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
