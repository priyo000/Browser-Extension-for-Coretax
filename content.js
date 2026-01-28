// Inject Interceptor to capture last fetch
const s = document.createElement('script');
s.src = chrome.runtime.getURL('injected.js');
s.onload = function () {
    this.remove();
    console.log("Interceptor injected successfully.");
};
(document.head || document.documentElement).appendChild(s);

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
            const transCode = taxNumber.length >= 2 ? taxNumber.substring(0, 2) : "";
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
                const fRecordId = item.RecordId || "UnknownID";
                const fSeller = item.SellerTIN || "UnknownSeller";
                const fTaxNo = item.TaxInvoiceNumber || "UnknownTaxNo";
                const fBuyer = item.BuyerTIN || "UnknownBuyer";

                const customFilename = `OutputTaxInvoice-${fRecordId}-${fSeller}-${fTaxNo}-${fBuyer}`;

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
