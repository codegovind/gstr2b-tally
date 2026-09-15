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
            showStatus("Invalid JSON format or missing B2B data.", "error");
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
    // 1. Setup the Headers matching the Tally Template
    const headers = [
        "Voucher Date", "Voucher Type Name", "Voucher Number", 
        "Buyer/Supplier - Address", "Buyer/Supplier - Pincode", 
        "Ledger Name", "Ledger Amount", "Ledger Amount Dr/Cr", 
        "Item Name", "Billed Quantity", "Item Rate", "Item Rate per", 
        "Item Amount", "Change Mode"
    ];

    let excelData = [headers];

    // 2. Extract B2B data safely
    const b2bData = jsonData?.data?.docdata?.b2b;
    
    if (!b2bData || !Array.isArray(b2bData)) {
        throw new Error("No B2B data found in JSON.");
    }

    // 3. Process each supplier and invoice
    b2bData.forEach(supplier => {
        const ctin = supplier.ctin || "";
        const trdnm = supplier.trdnm || ctin; 

        if (supplier.inv && Array.isArray(supplier.inv)) {
            supplier.inv.forEach(inv => {
                const inum = inv.inum || "";
                const idt = inv.idt || "";
                const val = inv.val || 0;

                const addRow = (ledgerName, amount, drCr) => {
                    excelData.push([
                        idt, "Purchase", inum, "", "", 
                        ledgerName, amount, drCr, 
                        "", "", "", "", "", "Accounting Invoice"
                    ]);
                };

                // ROW 1: Credit Supplier
                addRow(trdnm, val, "Cr");

                if (inv.itms && Array.isArray(inv.itms)) {
                    inv.itms.forEach(itm => {
                        const itmDet = itm.itm_det || {};
                        const txval = itmDet.txval || 0;
                        const cgst = itmDet.cgst || 0;
                        const sgst = itmDet.sgst || 0;
                        const igst = itmDet.igst || 0;

                        // ROW 2: Debit Standard Purchase Account
                        if (txval > 0) addRow("Purchase A/c", txval, "Dr");

                        // ROWS 3, 4, 5: Debit Taxes
                        if (cgst > 0) addRow("CGST", cgst, "Dr");
                        if (sgst > 0) addRow("SGST", sgst, "Dr");
                        if (igst > 0) addRow("IGST", igst, "Dr");
                    });
                }
            });
        }
    });

    // 4. Generate and download the Excel file using SheetJS
    const worksheet = XLSX.utils.aoa_to_sheet(excelData);
    const workbook = XLSX.utils.book_new();
    
    XLSX.utils.book_append_sheet(workbook, worksheet, "Accounting Voucher");
    XLSX.writeFile(workbook, "Purchase_ReadyToImport.xlsx");
}