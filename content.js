// Inject Interceptor to capture last fetch
const s = document.createElement('script');
s.src = chrome.runtime.getURL('injected.js');
s.onload = function () {
    this.remove();
    console.log("Interceptor injected successfully.");
};
(document.head || document.documentElement).appendChild(s);

// PDF.js Configuration
if (typeof pdfjsLib !== 'undefined') {
    pdfjsLib.GlobalWorkerOptions.workerSrc = chrome.runtime.getURL('pdf.worker.min.js');
}

let capturedConfig = null;

// Listen for intercepted data from injected.js
window.addEventListener("message", (event) => {
    if (event.source !== window) return;
    if (event.data.type && event.data.type === "TAX_INTERCEPT") {
        console.log("Interceptor captured data:", event.data);
        capturedConfig = event.data;
        notifyPopup("SUDAH SIAP! Klik tombol 'Ambil Seluruh Data' untuk memulai.");
    }
});

chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    // 0. CHECK STATUS (New popup opened)
    if (request.action === "check_status") {
        if (capturedConfig) {
            sendResponse({ status: "ready", message: "SUDAH SIAP! Klik tombol 'Ambil Seluruh Data' untuk memulai." });
        } else {
            sendResponse({ status: "waiting", message: "Tunggu Sebentar lagi disiapin..." });
        }
        return; // Synchronous response
    }

    if (request.action === "start_download_pdf") {
        if (!capturedConfig) {
            notifyPopup("ERROR: No data captured. Refresh & Load Table first.");
            sendResponse({ status: "error", message: "No data captured" });
            return;
        }

        // Use captured config exactly as is (no modification to rows)
        const urlToUse = capturedConfig.url;
        const tokenToUse = capturedConfig.auth;
        let payloadToUse = null;
        try {
            payloadToUse = typeof capturedConfig.payload === 'string' ? JSON.parse(capturedConfig.payload) : capturedConfig.payload;
        } catch (e) { }

        if (!payloadToUse) {
            notifyPopup("Payload Error");
            sendResponse({ status: "error", message: "Payload Error" });
            return;
        }

        // Pass captured response body as 4th argument
        startPdfDownloadProcess(urlToUse, tokenToUse, payloadToUse, capturedConfig.response);
        sendResponse({ status: "started" });
    }

    if (request.action === "start_detailed_excel") {
        if (!capturedConfig) {
            notifyPopup("ERROR: No data captured. Refresh & Load Table first.");
            sendResponse({ status: "error", message: "No data captured" });
            return;
        }

        const urlToUse = capturedConfig.url;
        const tokenToUse = capturedConfig.auth;
        let payloadToUse = null;
        try {
            payloadToUse = typeof capturedConfig.payload === 'string' ? JSON.parse(capturedConfig.payload) : capturedConfig.payload;
        } catch (e) { }

        if (!payloadToUse) {
            notifyPopup("Payload Error");
            sendResponse({ status: "error", message: "Payload Error" });
            return;
        }

        startDetailedExcelProcess(urlToUse, tokenToUse, payloadToUse, capturedConfig.response);
        sendResponse({ status: "started" });
    }

    if (request.action === "start_fetch") {

        // STRICT MODE: MUST HAVE CAPTURED CONFIG
        if (!capturedConfig) {
            notifyPopup("ERROR: No data captured yet. Please refresh the page and wait for the table to load.");
            sendResponse({ status: "error" });
            return;
        }

        const urlToUse = capturedConfig.url;
        const tokenToUse = capturedConfig.auth;
        let payloadToUse = null;

        try {
            payloadToUse = typeof capturedConfig.payload === 'string' ? JSON.parse(capturedConfig.payload) : capturedConfig.payload;
        } catch (e) {
            console.error("Failed to parse captured payload", e);
        }

        if (!payloadToUse) {
            notifyPopup("ERROR: Captured payload is invalid.");
            sendResponse({ status: "error" });
            return;
        }

        // Force Rows to 500
        if (typeof payloadToUse.Rows !== 'undefined') {
            payloadToUse.Rows = 500;
        }

        startFetching(urlToUse, tokenToUse, payloadToUse, capturedConfig.response);
        sendResponse({ status: "started" });
    }

    // Asynchronous response handling
    return true;
});

// (Stray Logic Removed)



async function startFetching(url, token, userPayload, initialData) {
    let allData = [];

    // Ensure userPayload is an object, if null init as empty
    if (!userPayload) {
        userPayload = { First: 0, Rows: 500 };
    }

    // Ensure First exists, default to 0 if not present
    if (typeof userPayload.First === 'undefined') {
        userPayload.First = 0;
    }

    let currentOffset = userPayload.First;
    let hasMore = true;

    try {
        notifyPopup(`Starting Loop... (Start Offset: ${currentOffset})`);

        while (hasMore) {
            // Construct payload for this iteration
            // We only modify 'First'. 'Rows' takes whatever the user put in the JSON.
            const payload = {
                ...userPayload,
                First: currentOffset
            };

            const headers = {
                "Content-Type": "application/json",
                "Accept": "application/json"
            };
            if (token) {
                headers["Authorization"] = token;
            }

            console.log(`Fetching Offset: ${currentOffset}`);

            let data;

            // Optimization: Use captured initial data for the first URL/Offset if matches
            // We assume the captured response corresponds to the 'userPayload' state (start offset)
            if (initialData && currentOffset === userPayload.First) {
                console.log("Using Initial Captured Data for first batch.");
                data = initialData;
                initialData = null; // Consume it so we don't reuse it
            } else {
                const response = await fetch(url, {
                    method: "POST",
                    headers: headers,
                    body: JSON.stringify(payload),
                    credentials: 'include'
                });

                if (!response.ok) {
                    const errorText = await response.text();
                    console.error("Server Error Response:", errorText);
                    throw new Error(`HTTP ${response.status}: ${errorText.substring(0, 100)}...`);
                }
                data = await response.json();
            }

            // Smart search for the array of data
            let items = findItemsArray(data);

            if (!items) {
                const debugStr = JSON.stringify(data).substring(0, 150);
                console.error("Structure unknown or no data found.", data);
                notifyPopup(`Warning: No data array found. Response: ${debugStr}`);
                // Do not add raw data to allData, avoids empty CSV rows
                hasMore = false;
            } else {
                if (items.length > 0) {
                    allData = allData.concat(items);
                    notifyPopup(`Collected ${allData.length} items...`);
                    currentOffset += items.length;
                } else {
                    hasMore = false;
                }
            }

            // Rate limiting precaution
            if (hasMore) {
                await new Promise(r => setTimeout(r, 300));
            }
        }


        notifyPopup(`Done! Total items: ${Array.isArray(allData) ? allData.length : 'Unknown'}. Downloading Excel...`);
        exportToExcel(allData, url);

    } catch (error) {
        console.error(error);
        notifyPopup(`Error: ${error.message}`);
    }
}

function notifyPopup(message) {
    // Safely try to notify popup. If closed, ignore error.
    try {
        console.log("Notifying Popup:", message);
        chrome.runtime.sendMessage({ action: "update_status", message: message }, (response) => {
            if (chrome.runtime.lastError) {
                // Ignore "Receiving end does not exist" which happens if popup is closed
            }
        });
    } catch (e) {
        // ignore
    }
}

function exportToExcel(data, sourceUrl = "") {
    if (!Array.isArray(data) || data.length === 0) {
        notifyPopup("ERROR: No data to export (Empty Array).");
        return;
    }

    // Determine type based on URL (Output vs Input)
    const isInputInvoice = sourceUrl && sourceUrl.toLowerCase().includes("inputinvoice/list");

    // Helper for Transaction Codes
    const getTransDesc = (code) => {
        const map = {
            '01': 'Kepada Pihak yang Bukan Pemungut PPN',
            '02': 'Kepada Pemungut Bendaharawan',
            '03': 'Kepada Pemungut Selain Bendaharawan',
            '04': 'DPP Nilai Lain',
            '05': 'Besaran Tertentu',
            '06': 'Penyerahan Lainnya',
            '07': 'Penyerahan yang PPN-nya Tidak Dipungut',
            '08': 'Penyerahan yang PPN-nya Dibebaskan',
            '09': 'Penyerahan Aktiva'
        };
        return map[code] ? `${code} - ${map[code]}` : code;
    };

    // Helper for Month Name (Indonesian)
    const getMonthName = (dateStr) => {
        if (!dateStr) return "";
        const months = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
        try {
            const d = new Date(dateStr);
            return months[d.getMonth()] || "";
        } catch (e) { return ""; }
    };

    // Helper for Date Format (YYYY-MM-DD)
    const fmtDate = (dateStr) => {
        if (!dateStr) return "";
        return dateStr.substring(0, 10); // 2025-11-07
    };

    let headers = [];
    let rows = [];

    if (isInputInvoice) {
        // --- CONFIG FOR INPUT INVOICE (Pajak Masukan) ---
        headers = [
            "NPWP Penjual",
            "Nama Penjual",
            "Nomor Faktur Pajak",
            "Tanggal Faktur Pajak",
            "Masa Pajak",
            "Tahun",
            "Masa Pajak Pengkreditkan",
            "Tahun Pajak Pengkreditan",
            "Status Faktur",
            "Harga Jual/Penggantian/DPP",
            "DPP Nilai Lain/DPP",
            "PPN",
            "PPnBM",
            "Perekam",
            "Nomor SP2D",
            "Valid",
            "Dilaporkan",
            "Dilaporkan oleh Penjual"
        ];

        // Mapping for Input Invoice
        rows = data.map(item => {
            return [
                item.SellerTIN || "",
                item.SellerTaxpayerName || "",
                item.TaxInvoiceNumber || "",
                fmtDate(item.TaxInvoiceDate),
                getMonthName(item.TaxInvoiceDate), // Masa Pajak
                item.TaxInvoiceYear || "",
                item.PeriodCredit || "", // Masa Pajak Pengkreditkan
                item.YearCredit || "",   // Tahun Pajak Pengkreditan
                item.TaxInvoiceStatus || "",
                item.SellingPrice || 0,
                item.OtherTaxBase || 0,
                item.VAT || 0,
                item.STLG || 0,
                item.Signer || "", // Perekam
                item.SP2DNumber || "",
                (item.Valid ? "TRUE" : "FALSE"),
                (item.ReportedByBuyer ? "TRUE" : "FALSE"), // Dilaporkan (oleh kita/pembeli)
                (item.ReportedBySeller ? "TRUE" : "FALSE")
            ];
        });

    } else {
        // --- CONFIG FOR OUTPUT INVOICE (Pajak Keluaran) - Default ---
        headers = [
            "NPWP Pembeli / Identitas",
            "Nama Pembeli",
            "Kode Transaksi",
            "Nomor Faktur Pajak",
            "Tanggal Faktur Pajak",
            "Masa Pajak",
            "Tahun",
            "Status Faktur",
            "ESignStatus",
            "Harga Jual/Penggantian/DPP",
            "DPP Nilai Lain/DPP",
            "PPN",
            "PPnBM",
            "Penandatangan",
            "Referensi",
            "Dilaporkan oleh Penjual",
            "Dilaporkan oleh Pemungut PPN"
        ];

        rows = data.map(item => {
            const taxNumber = item.TaxInvoiceNumber || "";
            let transCode = "";
            
            if (taxNumber.length >= 2) {
                transCode = taxNumber.substring(0, 2);
            } else if (item.TaxInvoiceCode && item.TaxInvoiceCode.length >= 2) {
                // Fallback: Ambil 2 digit terakhir dari TaxInvoiceCode (Misal "TD.00304" -> "04")
                transCode = item.TaxInvoiceCode.slice(-2);
            }
            const reportedByVat = (item.ReportedByVATCollector === null || item.ReportedByVATCollector === undefined) ? "" : (item.ReportedByVATCollector ? "TRUE" : "FALSE");

            return [
                item.BuyerTIN || "",
                item.BuyerName || item.BuyerTaxpayerNameClear || "",
                getTransDesc(transCode),
                taxNumber,
                fmtDate(item.TaxInvoiceDate),
                getMonthName(item.TaxInvoiceDate), // Masa Pajak
                item.TaxInvoiceYear || "",
                item.TaxInvoiceStatus || "",
                item.ESignStatus || "",
                item.SellingPrice || 0,
                item.OtherTaxBase || 0,
                item.VAT || 0,
                item.STLG || 0,
                item.Signer || "",
                item.Reference || "",
                item.ReportedBySeller ? "TRUE" : "FALSE",
                reportedByVat
            ];
        });
    }

    // Rows mapping already done above based on type logic

    // 3. Generate XML Content (SpreadsheetXML 2003)
    // This allows us to set proper types (String vs Number) to preserve leading zeros in Excel
    let xml = '<?xml version="1.0"?>\n';
    xml += '<?mso-application progid="Excel.Sheet"?>\n';
    xml += '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"\n';
    xml += ' xmlns:o="urn:schemas-microsoft-com:office:office"\n';
    xml += ' xmlns:x="urn:schemas-microsoft-com:office:excel"\n';
    xml += ' xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"\n';
    xml += ' xmlns:html="http://www.w3.org/TR/REC-html40">\n';
    xml += ' <Worksheet ss:Name="Sheet1">\n';
    xml += '  <Table>\n';

    // Header Row
    xml += '   <Row>\n';
    headers.forEach(h => {
        xml += `    <Cell><Data ss:Type="String">${h}</Data></Cell>\n`;
    });
    xml += '   </Row>\n';

    // Data Rows
    rows.forEach(row => {
        xml += '   <Row>\n';
        row.forEach(field => {
            // Escape special XML characters
            let cleanVal = String(field || "")
                .replace(/&/g, "&amp;")
                .replace(/</g, "&lt;")
                .replace(/>/g, "&gt;")
                .replace(/"/g, "&quot;")
                .replace(/'/g, "&apos;");

            // Force String type for all fields to be safe and preserve "00..."
            xml += `    <Cell><Data ss:Type="String">${cleanVal}</Data></Cell>\n`;
        });
        xml += '   </Row>\n';
    });

    xml += '  </Table>\n';
    xml += ' </Worksheet>\n';
    xml += '</Workbook>';

    // 4. Download as XLS (XML Spreadsheet)
    const blob = new Blob([xml], { type: "application/vnd.ms-excel" });
    const url = URL.createObjectURL(blob);

    // Create filename
    const filename = `tax_data_${new Date().toISOString().slice(0, 10)}.xls`;

    const downloadLink = document.createElement('a');
    downloadLink.href = url;
    downloadLink.download = filename;
    document.body.appendChild(downloadLink);
    downloadLink.click();
    document.body.removeChild(downloadLink);
    URL.revokeObjectURL(url);
}

// PDF DOWNLOAD LOGIC
async function startPdfDownloadProcess(listUrl, token, listPayload, capturedResponse) {
    try {
        let items = [];

        // 1. Prefer Captured Response (Instant)
        if (capturedResponse) {
            notifyPopup("Using Captured Response List (Instant)...");
            // reused finding logic
            const data = capturedResponse;
            items = findItemsArray(data);
        }

        // 2. Fetch list manually if no captured response (Fallback)
        if (!items || items.length === 0) {
            notifyPopup("Fetching list for PDF (Fallback)...");
            const response = await fetch(listUrl, {
                method: "POST",
                headers: {
                    "Content-Type": "application/json",
                    "Authorization": token
                },
                body: JSON.stringify(listPayload),
                credentials: 'include'
            });

            if (!response.ok) throw new Error("Failed to fetch list");

            const data = await response.json();
            // Smart find
            items = findItemsArray(data);
        }

        if (!items || items.length === 0) {
            notifyPopup("No items found in list to download.");
            return;
        }

        notifyPopup(`Found ${items.length} items. Starting PDF download...`);

        // 2. Loop and Download
        const pdfApiUrl = "https://coretaxdjp.pajak.go.id/einvoiceportal/api/DownloadInvoice/download-invoice-document";

        let successCount = 0;
        for (let i = 0; i < items.length; i++) {
            const item = items[i];
            notifyPopup(`Downloading PDF ${i + 1}/${items.length}...`);

            // Construct Payload
            // DocumentDate is current time as per user request
            const now = new Date();
            // Format: YYYY-MM-DDTHH:mm:ss (No 'Z' as per user example)
            const currentFormatted = now.toISOString().split('.')[0];

            // AUTO-DETECT Menu Type
            // Logic: 
            // 1. Check URL for 'input' (Incoming) vs 'output' (Outgoing)
            // 2. Check URL for 'return' OR Tax Number for 'RET' prefix
            const lowerUrl = (listUrl || "").toLowerCase();
            const isInput = lowerUrl.includes("input"); // e.g. inputinvoice/list
            
            // Check for Retur characteristic (URL has 'return' OR Number starts with RET)
            const isReturn = lowerUrl.includes("return") || 
                             (item.TaxInvoiceNumber && typeof item.TaxInvoiceNumber === 'string' && item.TaxInvoiceNumber.startsWith("RET"));

            let menuType = "Outgoing"; // Default
            if (isInput) {
                menuType = isReturn ? "IncomingReturn" : "Incoming";
            } else {
                menuType = isReturn ? "OutgoingReturn" : "Outgoing";
            }

            const pdfPayload = {
                "EInvoiceRecordIdentifier": item.RecordId || item.recordId,
                "EInvoiceAggregateIdentifier": item.AggregateIdentifier || item.aggregateIdentifier,
                "DocumentAggregateIdentifier": item.DocumentFormAggregateIdentifier || item.documentFormAggregateIdentifier,
                "TaxpayerAggregateIdentifier": item.SellerTaxpayerAggregateIdentifier || item.sellerTaxpayerAggregateIdentifier,
                "LetterNumber": item.TaxInvoiceNumber,
                "DocumentDate": currentFormatted,
                "EInvoiceMenuType": menuType,
                "TaxInvoiceStatus": item.TaxInvoiceStatus
            };
            
            console.log("Preparing PDF Payload:", JSON.stringify(pdfPayload));

            try {
                // Construct specific filename: OutputTaxInvoice-RecordId-SellerTIN-TaxNumber-BuyerTIN
                // Using fallback "0" or "X" if field missing to keep structure
                await downloadSinglePdf(pdfApiUrl, token, pdfPayload, customFilename);
                successCount++;
            } catch (err) {
                console.error("Failed PDF", item.TaxInvoiceNumber, err);
            }

            // Delay to be safe
            await new Promise(r => setTimeout(r, 800));
        }

        notifyPopup(`Finished! Downloaded ${successCount}/${items.length} PDFs.`);

    } catch (e) {
        console.error(e);
        notifyPopup("Error downloading PDFs: " + e.message);
    }
}

// DETAILED EXCEL PROCESS (The User's specific request)
async function startDetailedExcelProcess(listUrl, token, listPayload, capturedResponse) {
    try {
        let items = [];
        if (capturedResponse) {
            items = findItemsArray(capturedResponse);
        }

        if (!items || items.length === 0) {
            notifyPopup("Fetching list for Detailed Export...");
            const response = await fetch(listUrl, {
                method: "POST",
                headers: { "Content-Type": "application/json", "Authorization": token },
                body: JSON.stringify(listPayload),
                credentials: 'include'
            });
            if (!response.ok) throw new Error("Failed to fetch list");
            const data = await response.json();
            items = findItemsArray(data);
        }

        if (!items || items.length === 0) {
            notifyPopup("No items found to process.");
            return;
        }

        notifyPopup(`Processing ${items.length} invoices. This involves downloading PDFs...`);

        const pdfApiUrl = "https://coretaxdjp.pajak.go.id/einvoiceportal/api/DownloadInvoice/download-invoice-document";
        const allRows = [];

        // Optimasi: Gunakan Batch Processing agar lebih cepat (Turbo Mode)
        const BATCH_SIZE = 5; // Proses 5 PDF sekaligus
        for (let i = 0; i < items.length; i += BATCH_SIZE) {
            const batch = items.slice(i, i + BATCH_SIZE);
            
            await Promise.all(batch.map(async (invoice, bIdx) => {
                const currentIdx = i + bIdx + 1;
                notifyPopup(`Processing ${currentIdx}/${items.length}: ${invoice.TaxInvoiceNumber || invoice.LetterNumber}...`);

                // 1. Fetch PDF binary/json
                const now = new Date();
                const currentFormatted = now.toISOString().split('.')[0];
                const lowerUrl = (listUrl || "").toLowerCase();
                const isInput = lowerUrl.includes("input");
                
                const apiNo = invoice.LetterNumber || invoice.TaxInvoiceNumber || invoice.taxInvoiceNumber || "";
                const isStartingWithRet = String(apiNo).toUpperCase().startsWith("RET");
                const isReturn = lowerUrl.includes("return") || isStartingWithRet;
                
                let menuType = isInput ? (isReturn ? "IncomingReturn" : "Incoming") : (isReturn ? "OutgoingReturn" : "Outgoing");

                const pdfPayload = {
                    "EInvoiceRecordIdentifier": invoice.RecordId || invoice.recordId,
                    "EInvoiceAggregateIdentifier": invoice.AggregateIdentifier || invoice.aggregateIdentifier,
                    "DocumentAggregateIdentifier": invoice.DocumentFormAggregateIdentifier || invoice.documentFormAggregateIdentifier,
                    "TaxpayerAggregateIdentifier": invoice.SellerTaxpayerAggregateIdentifier || invoice.sellerTaxpayerAggregateIdentifier,
                    "LetterNumber": apiNo,
                    "DocumentDate": currentFormatted,
                    "EInvoiceMenuType": menuType,
                    "TaxInvoiceStatus": invoice.TaxInvoiceStatus
                };

                try {
                    const pdfResponse = await fetch(pdfApiUrl, {
                        method: "POST",
                        headers: { "Content-Type": "application/json", "Authorization": token },
                        body: JSON.stringify(pdfPayload),
                        credentials: 'include'
                    });

                    if (pdfResponse.ok) {
                        const json = await pdfResponse.json();
                        let base64 = null;
                        if (json.Payload && json.Payload.Message) base64 = json.Payload.Message.Data;
                        else if (json.Data) base64 = json.Data;

                        if (base64) {
                            // 2. Extract Data from PDF
                            const extracted = await extractInfoFromPdf(base64); 
                            const pdfItems = (extracted && extracted.items && extracted.items.length > 0) 
                                ? extracted.items 
                                : [{ productName: "Extraction Failed", unit: "", unitPrice: 0, total: 0 }];

                            // 3. Create Excel Rows
                            let fNoFaktur = "";
                            let fNoRetur = "";
                            if (isStartingWithRet) fNoRetur = apiNo; else fNoFaktur = apiNo;
                            
                            if (extracted.returnNumber) fNoRetur = extracted.returnNumber;
                            if (extracted.refInvoice) fNoFaktur = extracted.refInvoice;

                            pdfItems.forEach((item, index) => {
                                const isLast = (index === pdfItems.length - 1);
                                allRows.push([
                                    (invoice.TaxInvoiceDate || "").substring(0, 10),
                                    fNoFaktur, fNoRetur,
                                    item.productName || "",                            // 3: NAMA BARANG
                                    item.qty || 0,                                     // 4: JUMLAH SATUAN
                                    item.unit || "",                                   // 5: SATUAN
                                    item.unitPrice || 0,                               // 6: HARGA SATUAN PER ITEM
                                    item.total || 0,                                   // 7: TOTAL HARGA PER ITEM
                                    isLast ? (invoice.SellingPrice || 0) : "",         // 8: TOTAL PERFAKTUR (DPP)
                                    isLast ? (invoice.VAT || 0) : "",                  // 9: PPN
                                    isLast ? ((invoice.SellingPrice || 0) + (invoice.VAT || 0)) : "" // 10: TOTAL DPP + PPN
                                ]);
                            });
                            // Spacer
                            allRows.push(["", "", "", "", "", "", "", "", "", "", ""]);
                        }
                    }
                } catch (err) {
                    console.error("Detailed analyze failed", apiNo, err);
                    allRows.push([(invoice.TaxInvoiceDate || "").substring(0, 10), apiNo, "", "ERROR: PDF FAIL", "", 0, 0, 0, 0, 0, 0]);
                    allRows.push(["", "", "", "", "", "", "", "", "", "", ""]);
                }
            }));

            // Jeda antar batch (Turbo: 200ms)
            await new Promise(r => setTimeout(r, 200));
        }

        // 4. Generate the Final Excel
        const headers = ["TANGGAL", "NO FAKTUR", "NO RETUR", "NAMA BARANG", "JUMLAH SATUAN", "SATUAN", "HARGA SATUAN PER ITEM", "TOTAL HARGA PER ITEM", "TOTAL PERFAKTUR (DPP)", "PPN", "TOTAL DPP PERFAKTUR + PPN"];
        exportDetailedToExcel(headers, allRows);
        notifyPopup("Detailed Excel Exported!");

    } catch (e) {
        console.error(e);
        notifyPopup("Detailed Export Error: " + e.message);
    }
}

function exportDetailedToExcel(headers, rows) {
    let xml = '<?xml version="1.0"?>\n<?mso-application progid="Excel.Sheet"?>\n';
    xml += '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" xmlns:html="http://www.w3.org/TR/REC-html40">\n';
    xml += ' <Worksheet ss:Name="Sheet1">\n  <Table>\n';
    
    // Headers
    xml += '   <Row>\n';
    headers.forEach(h => { xml += `    <Cell><Data ss:Type="String">${h}</Data></Cell>\n`; });
    xml += '   </Row>\n';

    // Rows
    rows.forEach(row => {
        xml += '   <Row>\n';
        row.forEach((field, idx) => {
            let type = "String";
            // Numeric for price/qty columns ([4] is Qty, [6:] are prices) but ONLY if there is data
            const numericIndices = [4, 6, 7, 8, 9, 10];
            if (numericIndices.includes(idx) && field !== "" && field !== null && field !== undefined) {
                type = "Number";
            }
            
            let cleanVal = String(field || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
            xml += `    <Cell><Data ss:Type="${type}">${cleanVal}</Data></Cell>\n`;
        });
        xml += '   </Row>\n';
    });

    xml += '  </Table>\n </Worksheet>\n</Workbook>';
    const blob = new Blob([xml], { type: "application/vnd.ms-excel" });
    const url = URL.createObjectURL(blob);
    const filename = `tax_item_details_${new Date().toISOString().slice(0, 10)}.xls`;
    const a = document.createElement('a');
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
}

async function downloadSinglePdf(url, token, payload, filenameLabel) {
    const response = await fetch(url, {
        method: "POST",
        headers: {
            "Content-Type": "application/json",
            "Authorization": token
        },
        body: JSON.stringify(payload),
        credentials: 'include'
    });

    if (!response.ok) {
        const errText = await response.text();
        throw new Error("Server error " + response.status + " | " + errText.substring(0, 150));
    }

    const contentType = response.headers.get("content-type");

    // STRICTLY use the provided filename (plus .pdf)
    // We ignore server content-disposition because user requires specific format
    let finalFilename = filenameLabel;
    if (!finalFilename.toLowerCase().endsWith(".pdf")) {
        finalFilename += ".pdf";
    }

    // Check if response is JSON (likely Base64 wrapped)
    if (contentType && contentType.includes("application/json")) {
        const json = await response.json();

        let base64Data = null;

        // Check deep structure consistent with user provided example
        if (json.Payload && json.Payload.Message) {
            if (json.Payload.Message.Data) base64Data = json.Payload.Message.Data;
            // IGNORE Server Filename, force our own format
        }
        // Fallback checks
        else if (json.Data) base64Data = json.Data;
        else if (json.data) base64Data = json.data;

        if (base64Data) {
            try {
                // FEATURE: Extract Product Name for better filename
                const extractedInfo = await extractInfoFromPdf(base64Data);
                if (extractedInfo && extractedInfo.productName) {
                    console.log("Extracted Product Name:", extractedInfo.productName);
                    // Add product name to filename prefix (limit length)
                    const cleanName = extractedInfo.productName.substring(0, 30).replace(/[^a-zA-Z0-9]/g, '_');
                    finalFilename = `${cleanName}-${finalFilename}`;
                }

                // Convert Base64 to Blob manually to avoid URL length limits
                const binaryString = atob(base64Data);
                const len = binaryString.length;
                const bytes = new Uint8Array(len);
                for (let i = 0; i < len; i++) {
                    bytes[i] = binaryString.charCodeAt(i);
                }
                const blob = new Blob([bytes], { type: "application/pdf" });

                const blobUrl = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = blobUrl;
                a.download = finalFilename;
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                URL.revokeObjectURL(blobUrl);
                return;
            } catch (e) {
                throw new Error("Failed to decode PDF Base64: " + e.message);
            }
        }

        // If JSON but no data found, probably an error message
        throw new Error("API returned JSON but no PDF data found: " + (json.Message || JSON.stringify(json).substring(0, 100)));
    }

    // Assume standard binary response (Blob) - unlikely now based on user report
    const blob = await response.blob();
    if (blob.size < 100) {
        throw new Error("PDF file too small, likely corrupted.");
    }

    const blobUrl = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = blobUrl;
    a.download = finalFilename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(blobUrl);
}

function findItemsArray(data) {
    if (!data) return null;
    if (Array.isArray(data)) return data;

    // 1. Common paths
    if (Array.isArray(data.data)) return data.data;
    if (Array.isArray(data.Items)) return data.Items;
    if (Array.isArray(data.rows)) return data.rows;
    if (Array.isArray(data.value)) return data.value;

    // 2. Deep path (Coretax style)
    if (data.Payload && data.Payload.Message && Array.isArray(data.Payload.Message.Data)) {
        return data.Payload.Message.Data;
    }

    // 3. Fallback: Recursive search (limited depth)
    const MAX_DEPTH = 3;
    function search(obj, depth) {
        if (depth > MAX_DEPTH) return null;
        if (!obj || typeof obj !== 'object') return null;

        // Check immediate children first
        for (const key in obj) {
            if (Array.isArray(obj[key]) && obj[key].length > 0) {
                // Assume this is it
                return obj[key];
            }
        }

        // Go deeper
        for (const key in obj) {
            if (obj[key] && typeof obj[key] === 'object') {
                const res = search(obj[key], depth + 1);
                if (res) return res;
            }
        }
        return null;
    }

    return search(data, 0);
}

// PDF EXTRACTION HELPER
async function extractInfoFromPdf(base64) {
    if (typeof pdfjsLib === 'undefined') return null;

    try {
        const binaryString = atob(base64);
        const len = binaryString.length;
        const bytes = new Uint8Array(len);
        for (let i = 0; i < len; i++) {
            bytes[i] = binaryString.charCodeAt(i);
        }

        const loadingTask = pdfjsLib.getDocument({ data: bytes });
        const pdf = await loadingTask.promise;
        const page = await pdf.getPage(1);
        const textContent = await page.getTextContent();
        
        const strings = textContent.items.map(item => item.str.trim()).filter(s => s.length > 0);
        console.log("PDF TEXT STRINGS:", strings);

        const foundItems = [];
        
        // Find indices of headers to bound the search
        const headerIdx = strings.findIndex(s => s.includes("Nama Barang Kena Pajak"));
        // Boundary: end of items table usually starts with "Harga Jual" summary or "Potongan Harga"
        let footerIdx = strings.findIndex(s => s.includes("Potongan Harga"));
        if (footerIdx === -1) footerIdx = strings.findIndex(s => s.includes("Harga Jual / Penggantian / Uang Muka"));
        if (footerIdx === -1) footerIdx = strings.length;

        if (headerIdx !== -1) {
            // Loop through potential rows
            for (let i = headerIdx; i < footerIdx - 2; i++) {
                // Heuristic: No (digit) -> Code (digit strings) -> Name (text) -> Details (Rp...x...) -> Total (price string)
                if (/^\d+$/.test(strings[i]) && strings[i].length <= 3) {
                    const no = strings[i];
                    const nextIsCode = /^\d+$/.test(strings[i+1]) || (strings[i+1].length >= 5);
                    
                    if (nextIsCode) {
                        const code = strings[i+1];
                        const name = strings[i+2];
                        let unit = "";
                        let unitPrice = 0;
                        let qty = 0;
                        let lineTotal = 0;

                        // Details line usually "Rp 3.040,54 x 4,00 Piece"
                        // Or might be split into multiple strings
                        // Look ahead for the Rp line
                        for (let j = i + 3; j < i + 10 && j < footerIdx; j++) {
                            if (strings[j].includes("Rp") && strings[j].includes("x")) {
                                const detailText = strings[j];
                                // Regex: Rp (price) x (qty) (unit)
                                // Example: Rp 3.040,54 x 4,00 Piece
                                const match = detailText.match(/Rp\s*([\d.,]+)\s*x\s*([\d.,]+)\s*(\w+)/);
                                if (match) {
                                    unitPrice = parseFloat(match[1].replace(/\./g, '').replace(',', '.'));
                                    qty = parseFloat(match[2].replace(/\./g, '').replace(',', '.'));
                                    unit = match[3];
                                }
                                
                                // The line total is usually one of the next strings that looks like a price but NOT the detail line
                                // Example: "12.162,16"
                                for (let k = j + 1; k < j + 4 && k < footerIdx; k++) {
                                     // Check if string contains . and , and looks like a price
                                     if (/^[\d.]+,\d{2}$/.test(strings[k])) {
                                         lineTotal = parseFloat(strings[k].replace(/\./g, '').replace(',', '.'));
                                         break;
                                     }
                                }
                                break;
                            }
                        }

                        if (name) {
                            foundItems.push({
                                productName: name,
                                unit: unit,
                                unitPrice: unitPrice,
                                qty: qty,
                                total: lineTotal
                            });
                        }
                    }
                }
            }
        }

        // NOMOR NOTA RETUR & REFERENSI EXTRACTION
        let returnNumber = "";
        let refInvoice = "";

        // 1. Look for document number in headers (RET...)
        const retIdx = strings.findIndex(s => s.toUpperCase().includes("NOMOR") && s.toUpperCase().includes("RET"));
        if (retIdx !== -1) {
            const raw = strings[retIdx];
            const matchRet = raw.match(/(RET\d+)/i);
            if (matchRet) returnNumber = matchRet[1];
        }

        // 2. Look for Original Invoice Reference
        // Matches common variations: "atas nomor Faktur Pajak", "Nomor Faktur Pajak yang diretur", etc.
        const refIdx = strings.findIndex(s => {
            const lower = s.toLowerCase();
            return (lower.includes("faktur pajak") && (lower.includes("atas nomor") || lower.includes("diretur") || lower.includes("nomor:")));
        });

        if (refIdx !== -1) {
            const raw = strings[refIdx];
            // Match exactly 16 digits or 13 digits (standard FP length)
            const match = raw.match(/(\d{13,16})/); 
            if (match) {
                refInvoice = match[1];
            } else if (strings[refIdx + 1]) {
                const matchNext = strings[refIdx + 1].match(/(\d{13,16})/);
                if (matchNext) refInvoice = matchNext[1];
            }
        }

        return {
            items: foundItems,
            returnNumber: returnNumber,
            refInvoice: refInvoice,
            productName: foundItems.length > 0 ? foundItems[0].productName : ""
        };

    } catch (err) {
        console.warn("PDF extraction failed", err);
        return null;
    }
}
