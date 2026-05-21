console.log("script.js loaded!");
// import { createClient } from "https://cdn.jsdelivr.net/npm/@supabase/supabase-js/+esm";

const { createClient } = window.supabase;

const supabaseUrl = 'https://jkykahgmkdxyfcwexoeh.supabase.co'
const supabaseKey = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImpreWthaGdta2R4eWZjd2V4b2VoIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzkxMDEwNjIsImV4cCI6MjA5NDY3NzA2Mn0.YrMvQjiG8p_jACKvzqmS9cFdhbQJi0jC343cTR4BA3E';
const supabase = createClient(supabaseUrl, supabaseKey)

document.getElementById('uploadFirstFile').addEventListener('change', handleFileUpload);

let firstFileData = [];
let secondFileData = {};
let mergedData = [];
let handleToBankMap = {};
let ifscToBankMap = {};
let originCategoryMap = {};
let categoryWebsiteMap = {};
let isMerged = false;
let websiteExcelData = [];

async function loadStaticJson() {
    // Load the static JSON file only once
    const mergeType = document.getElementById('mergeTypeDropdown').value;

    let jsonFilePath = '';
    if (mergeType === 'upi' || mergeType === 'credit_netbanking' || mergeType === 'not_found' || mergeType === 'crypto' || mergeType === 'investment_web') {
        jsonFilePath = 'json/secondFile.json'; // JSON for UPI
    } else if (mergeType === 'telegram') {
        jsonFilePath = 'json/telegram_wtsp.json'; // JSON for Telegram
    }
    else if (mergeType === 'investment_scam') {
        jsonFilePath = 'json/investment_scam.json'; // JSON for Telegram
    }

    // Load the selected JSON file
    const selectedJsonFile = await fetch(jsonFilePath);
    secondFileData = await selectedJsonFile.json();

    const handleBankFile = await fetch('json/handleBankName.json');
    const handleFileData = await handleBankFile.json();

    const ifscBankFile = await fetch('json/ifscBankName.json');
    const ifscFileData = await ifscBankFile.json();

    // const originWebsite = await fetch('json/originWebsite.json');
    // const origin = await originWebsite.json();

    // const categoryWebsite = await fetch('json/categoryWebsite.json');
    // const category = await categoryWebsite.json();

    // Assuming the handles and bank names are in `Sheet1`
    handleFileData.Sheet2.forEach(item => {
        if (item.Handle && item.Bank_name) {
            handleToBankMap[item.Handle.toLowerCase()] = item.Bank_name;
        }
    });

    ifscFileData.Sheet3.forEach(item => {
        if (item.ifsc_code && item.bank_name) {
            ifscToBankMap[item.ifsc_code] = item.bank_name;
        }
    })

    // origin.Sheet1.forEach(item => {
    //     if (item.url && item.origin) {
    //         originWebsiteMap[item.url] = item.origin;
    //     }
    // })

    // category.Sheet1.forEach(item => {
    //     if (item.url && item.category) {
    //         categoryWebsiteMap[item.url] = item.category;
    //     }
    // })

    originCategoryMap = await fetchOriginFromSupabase();
    categoryWebsiteMap = await fetchCategoryFromSupabase();


    const sheet1Data = secondFileData.Sheet1; // All objects from Sheet1
    const sheet2Data = handleFileData.sheet2; // All objects from sheet2
    const sheet3Data = ifscFileData.sheet3;
    // const sheet4Data = origin.sheet1;
    // const sheet5Data = category.sheet1;
    const sheet4Data = originCategoryMap.sheet1;
    const sheet5Data = categoryWebsiteMap.sheet1;

    // secondFileData = { sheet1Data, sheet2Data, sheet3Data, sheet4Data, sheet5Data };
    secondFileData = { sheet1Data, sheet2Data, sheet3Data };

    return secondFileData;
}

async function fetchAllRows(tableName) {
    let allRows = [];
    let batchSize = 1000;
    let start = 0;
    let hasMore = true;
    while (hasMore) {
        const { data, error } = await supabase
            .from(tableName)
            .select('*')
            .range(start, start + batchSize - 1);
        if (error) { console.error(error); break; }
        allRows = allRows.concat(data);
        hasMore = data.length === batchSize;
        start += batchSize;
    }
    return allRows;
}


async function fetchOriginFromSupabase() {
    const rows = await fetchAllRows('Website_origin_category');
    const originWebsiteMap = {};
    rows.forEach(row => {
        originWebsiteMap[row.url] = {
            origin: row.origin || null,
            Category: row.Category || null
        }
    });
    return originWebsiteMap;
}

async function fetchCategoryFromSupabase() {
    const rows = await fetchAllRows('categoryWebsite');
    const categoryWebsiteMap = {};
    rows.forEach(row => {
        categoryWebsiteMap[row.url] = row.category;
    });
    return categoryWebsiteMap;
}


async function handleFileUpload(event) {
    const file = event.target.files[0];
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    firstFileData = XLSX.utils.sheet_to_json(sheet);
}

function extractDomain(url) {
    try {
        const parsedUrl = new URL(url);
        let domain = parsedUrl.hostname;
        domain = domain.replace(/^www\./, '');
        return domain;  // Extracts the domain (without 'https://' or path)
    } catch (e) {
        return 'NA';  // In case of empty column
    }
}

function determineType(upiVpa) {
    const upiVpaStr = String(upiVpa).trim();
    if (!upiVpaStr) return 'Bank Account';

    // Check if it's a UPI ID (contains @)
    if (upiVpaStr.includes('@')) {
        return 'UPI';
    }

    // Check if it's a phone number (basic check for 10 digits)
    const phonePattern = /^\d{10}$/;
    if (phonePattern.test(upiVpaStr)) {
        return 'Wallet';
    }

    return 'Bank Account'; // If neither, return NA
}

function extractTimestampFromUrl(url) {
    // Extract the number from the URL (after 'npci-')
    const match = url.match(/(?:npci|mfilterit|without_header)-(\d+)_/);
    // const match = url.match(/--(\d+)--/);
    if (match && match[1]) {
        return parseInt(match[1], 10);  // Convert the matched number to an integer
    }
    return null; // Return null if no number found
}

function convertTimestampToDate(timestamp) {
    if (timestamp) {
        const date = new Date(0); // Start with Unix epoch (1970-01-01)
        date.setSeconds(timestamp); // Add seconds
        // Adjust for your timezone if needed (e.g., GMT+5:30)
        date.setHours(date.getHours() + 5); // Adjust for hours
        date.setMinutes(date.getMinutes() + 30); // Adjust for minutes
        return date.toISOString().slice(0, 10);
    }
    return 'Invalid Timestamp'; // Return this if the timestamp is not valid
}

function determinePlatform(url) {
    if (url.includes('wa')) {
        return 'WhatsApp';
    } else if (url.includes('telegram')) {
        return 'Telegram';
    } else if (url.includes('t.me')) {
        return 'Telegram';
    } else if (url.includes('instagram')) {
        return 'Instagram';
    } else if (url.includes('facebook')) {
        return 'Facebook';
    }
    return 'NA';
}

function convertToDateTime(npciNumber) {
    if (npciNumber) {
        const date = new Date(0); // Start with Unix epoch (1970-01-01)
        date.setSeconds(npciNumber); // Add seconds
        // Adjust for your timezone if needed (e.g., GMT+5:30)
        date.setHours(date.getHours() + 5); // Adjust for hours
        date.setMinutes(date.getMinutes() + 30); // Adjust for minutes
        return date.toISOString().slice(0, 19).replace('T', ' ');
    } // Convert string to number
}

async function previewData() {
    if (firstFileData.length === 0 || secondFileData.length === 0) {
        alert('Please upload the first file and ensure the JSON file is loaded.');
        return;
    }

    await loadStaticJson();

    // Merge each Excel row with the full JSON row structure
    mergedData = firstFileData.map(excelRow => {

        const mergeType = document.getElementById('mergeTypeDropdown').value;

        let ss_url = excelRow?.npci_mfilterit_without_header ? excelRow.npci_mfilterit_without_header : '';
        let npciUrl = ''
        let mfilterit = '';
        let without_header = '';

        if (ss_url.includes('npci')) {
            mfilterit = ss_url.replace('npci', 'mfilterit');
            without_header = ss_url.replace('npci', 'without_header');
            npciUrl = ss_url;
        } else if (ss_url.includes('mfilterit')) {
            npciUrl = ss_url.replace('mfilterit', 'npci');
            without_header = ss_url.replace('mfilterit', 'without_header');
            mfilterit = ss_url
        } else if (ss_url.includes('without_header')) {
            npciUrl = ss_url.replace('without_header', 'npci');
            mfilterit = ss_url.replace('without_header', 'mfilterit');
            without_header = ss_url;
        }

        const npci_mfilterit = [mfilterit, npciUrl, without_header].filter(Boolean).join(',');

        let bankName = "NA";

        let upiHandle = 'NA';
        let ifscCode = 'NA';

        // if (mergeType === 'upi' || mergeType === 'telegram' || mergeType === 'investment_scam' || mergeType === 'investment_web') {
        //     upiHandle = excelRow?.upi_vpa && String(excelRow.upi_vpa).includes('@')
        //         ? String(excelRow.upi_vpa).split('@')[1].toLowerCase()
        //         : 'NA';
        //         console.log("UPI Handle:", upiHandle);

        //     // Extract IFSC code
        //     ifscCode = excelRow?.ifsc_code && excelRow.ifsc_code !== 'NA'
        //         ? excelRow.ifsc_code.trim().substring(0, 4).toUpperCase()
        //         : null;

        //     // Prioritize IFSC-based bank lookup if IFSC code exists
        //     if (ifscCode && ifscToBankMap[ifscCode]) {
        //         bankName = ifscToBankMap[ifscCode];
        //     }
        //     // Fallback to UPI handle-based lookup if no valid IFSC code
        //     else if (upiHandle && handleToBankMap[upiHandle]) {
        //         bankName = handleToBankMap[upiHandle];
        //         console.log("Bank Name from UPI Handle:", bankName);
        //     }
        // } else if (mergeType === 'credit_netbanking') {
        //     bankName = excelRow?.bank_name || '';
        // }

        if (['upi', 'telegram', 'investment_scam', 'investment_web'].includes(mergeType)) {
            upiHandle = excelRow?.upi_vpa && String(excelRow.upi_vpa).includes('@')
                ? String(excelRow.upi_vpa).split('@')[1].toLowerCase()
                : 'NA';

            const ifscRaw = String(excelRow?.ifsc_code ?? '').trim();
            const ifscCode = (ifscRaw && ifscRaw !== 'NA')
                ? ifscRaw.substring(0, 4).toUpperCase()
                : null;

            if (ifscCode && ifscToBankMap[ifscCode]) {
                bankName = ifscToBankMap[ifscCode];
            } else if (upiHandle && handleToBankMap[upiHandle]) {
                bankName = handleToBankMap[upiHandle];
            } else {
                console.log('No matching bank found for IFSC or UPI handle.');
            }
        } else if (mergeType === 'credit_netbanking') {
            bankName = excelRow?.bank_name || 'NA';
        }

        const upiType = mergeType === 'upi' || mergeType === 'telegram' || mergeType === 'investment_scam' || mergeType === 'investment_web'
            ? determineType(excelRow?.upi_vpa || '')
            : mergeType === 'credit_netbanking'
                ? excelRow?.platform?.replace('banking', 'Banking')
                : mergeType === 'crypto'
                    ? 'Crypto'
                    : 'NA';

        // Extract the timestamp from the URL and convert it to a date
        const timestamp = extractTimestampFromUrl(excelRow?.npci_mfilterit_without_header); // Adjust the column name as needed
        const date = convertTimestampToDate(timestamp)

        const dateTime = convertToDateTime(timestamp);

        function normalize(url) {
            return (url || '')
                .toLowerCase()
                .trim()
                .replace(/^https?:\/\//, '')
                .replace(/^www\./, '')
                .replace(/\/$/, '');
        }

        const origin =
            (mergeType === 'upi' ||
                mergeType === 'credit_netbanking' ||
                mergeType === 'not_found' ||
                mergeType === 'crypto' ||
                mergeType === 'investment_web')
                ? (
                    (() => {
                        const cleanUrl = normalize(excelRow.website_url);
                        const foundKey = Object.keys(originCategoryMap).find(key =>
                            normalize(key) === cleanUrl
                        );
                        return foundKey ? originCategoryMap[foundKey].origin : 'NA';
                    })()
                )
                : (mergeType==="telegram")
                    ? "India"
                    : "NA";

        const categoryMap_Inevst_scam = {
            "t.me": "Telegram",
            "telegram.org": "Telegram",
            "wa.me": "Whatsapp",
            "facebook.com": "Facebook",
            "instagram.com": "Instagram",
            "threads.com": "Thread",
            "youtube.com": "YouTube",
            "x.com": "X"
        }

        const category =
            (mergeType === 'upi' ||
                mergeType === 'credit_netbanking' ||
                mergeType === 'not_found' ||
                mergeType === 'crypto')
                ? (
                    (() => {
                        const cleanUrl = normalize(excelRow.website_url);
                        const foundKey = Object.keys(originCategoryMap).find(key =>
                            normalize(key) === cleanUrl
                        );
                        return foundKey ? originCategoryMap[foundKey].Category : 'NA';
                    })()
                )
                : (mergeType === "investment_scam" || mergeType === "investment_web")
                    ? (excelRow?.category || '')
                    : mergeType === 'telegram'
                        ? excelRow?.category
                        : "NA"

        const search_for = mergeType === 'investment_scam'
            ? (() => {
                const url = excelRow?.website_url || "";
                const match = Object.keys(categoryMap_Inevst_scam).find((domain) =>
                    url.includes(domain)
                );
                return match ? categoryMap_Inevst_scam[match] : "NA";
            })()
            : mergeType === 'upi' || mergeType === 'credit_netbanking' || mergeType === 'not_found' || mergeType === 'crypto' || mergeType === 'investment_web'
                ? 'Web'
                : mergeType === 'telegram'
                    ? 'Messaging Channel Platforms'
                    : 'NA';

        const paymentUrl = mergeType === 'upi' || mergeType === 'crypto' || mergeType === 'investment_web'
            ? (excelRow?.payment_gateway_url || 'NA')
            : mergeType === 'credit_netbanking'
                ? (excelRow?.destination_url || '')
                : "NA";

        const upiUrl = mergeType === 'upi' || mergeType === 'crypto' || mergeType === 'investment_web'
            ? (excelRow?.payment_gateway_url || 'NA')
            : "NA";

        const intermediateUrl1 = excelRow?.intermediate_url_1 ? excelRow?.intermediate_url_1 : '';
        const intermediateUrl2 = excelRow?.intermediate_url_2 ? excelRow?.intermediate_url_2 : '';
        const intermediateUrl3 = excelRow?.intermediate_url_3 ? excelRow?.intermediate_url_3 : '';
        const intermediateUrl4 = excelRow?.intermediate_url_4 ? excelRow?.intermediate_url_4 : '';

        const intermediateUrls = mergeType === 'upi' || mergeType === 'crypto' || mergeType === 'investment_web'
            ? (excelRow?.payment_gateway_url || 'NA')
            : mergeType === 'credit_netbanking'
                ? [intermediateUrl1, intermediateUrl2, intermediateUrl3, intermediateUrl4]
                    .filter(Boolean)
                    .join(',') // Join domains with commas
                : 'NA';

        const intermediateDomainName =
            mergeType === 'credit_netbanking'
                ? [intermediateUrl1, intermediateUrl2, intermediateUrl3, intermediateUrl4]
                    .filter(Boolean) // Remove empty or null values
                    .map(extractDomain) // Extract domain from each URL
                    .join(',') // Join domains with commas
                : '';

        const paymentIntermediateUrls = mergeType === 'upi' || mergeType === 'crypto' || mergeType === 'investment_web'
            ? (extractDomain(excelRow?.payment_gateway_url || '') === extractDomain(excelRow?.website_url || '')
                ? 'NA'
                : extractDomain(excelRow?.payment_gateway_url || ''))
            : mergeType === 'credit_netbanking'
                ? intermediateDomainName
                : 'NA';

        const bankAccountNumber = mergeType === 'upi' || mergeType === 'telegram' || mergeType === 'investment_scam' || mergeType === 'investment_web'
            ? excelRow?.bank_account_number || ''
            : mergeType === 'credit_netbanking'
                ? 'NA'
                : 'NA';

        const ifsc = mergeType === 'upi' || mergeType === 'telegram' || mergeType === 'investment_scam' || mergeType === 'investment_web'
            ? excelRow?.ifsc_code || ''
            : mergeType === 'credit_netbanking'
                ? 'NA'
                : 'NA';

        const upiId = mergeType === 'upi' || mergeType === 'telegram' || mergeType === 'investment_scam' || mergeType === 'investment_web'
            ? excelRow?.upi_vpa || ''
            : mergeType === 'credit_netbanking'
                ? 'NA'
                : 'NA';

        const accHolderName = mergeType === 'upi' || mergeType === 'telegram' || mergeType === 'investment_scam' || mergeType === 'investment_web'
            ? excelRow?.account_holder_name
            : mergeType === "credit_netbanking"
                ? excelRow?.account_holder_name
                    ? excelRow?.account_holder_name
                    : "NA"
                : 'NA';

        const branchName = mergeType === 'not_found'
            ? ''
            : "NA"

        const crypto_wallet_id = mergeType === 'crypto'
            ? excelRow?.Crypto_wallet_id
            : "NA"

        const crypto_platform = mergeType === 'crypto'
            ? excelRow?.Crypto_platform
            : "NA"

        const crypto_wallet = excelRow?.Balance_in_crypto_wallet
        const balance_in_crypto_wallet = mergeType === 'crypto'
            ? (crypto_wallet !== undefined && crypto_wallet !== null ? String(crypto_wallet) : "NA")
            : "NA"

        const transaction_count_value = excelRow?.Crypto_wallet_transaction_count;
        const crypto_wallet_transaction_count = mergeType === 'crypto'
            ? (transaction_count_value !== undefined && transaction_count_value !== null ? String(transaction_count_value) : "NA")
            : "NA"

        const contact_no = mergeType === 'investment_scam'
            ? excelRow?.contact_no
            : 'NA'

        const feature_type = mergeType === 'investment_web' || mergeType === 'investment_scam'
            ? "BS Investment Scam"
            : "BS Money Laundering"

        const scam_type = mergeType === 'investment_scam' || mergeType === 'investment_web'
            ? (excelRow?.category || '')
            : ""

        return {
            ...secondFileData.sheet1Data[0], // Start with the full JSON structure as the base,
            bank_account_number: bankAccountNumber, // Account Number
            ifsc_code: ifsc, //IFSC Code
            upi_vpa: upiId, // upi id
            ac_holder_name: accHolderName, //account holder name
            website_url: excelRow?.website_url || secondFileData.sheet1Data[0].website_url, //website url
            payment_gateway_intermediate_url: intermediateUrls, //payment gateway url
            payment_gateway_url: paymentUrl, //payment gateway url
            upi_url: upiUrl,
            transaction_method: excelRow?.transaction_method || secondFileData.sheet1Data[0].transaction_method, // Transaction Method
            screenshot: npci_mfilterit,
            screenshot_case_report_link: npci_mfilterit,
            handle: upiHandle,
            payment_gateway_name: paymentIntermediateUrls,
            upi_bank_account_wallet: upiType,
            inserted_date: date,
            case_generated_time: dateTime,
            bank_name: bankName,
            origin: origin,
            category_of_website: category,
            // platform: platform,
            bank_branch_details: branchName,
            crypto_wallet_id: crypto_wallet_id,
            crypto_platform: crypto_platform,
            balance_in_crypto_wallet: balance_in_crypto_wallet,
            crypto_wallet_transaction_count: crypto_wallet_transaction_count,
            web_contact_no: contact_no,
            search_for: search_for,
            feature_type: feature_type,
            scam_type: scam_type
        };
    });

    displayPreview(mergedData);

    return true;
}

function displayPreview(data) {
    const container = document.getElementById("previewContainer");
    container.innerHTML = "";

    const table = document.createElement("table");
    table.style.borderCollapse = "collapse";
    table.style.width = "100%";

    // Generate table headers
    const headerRow = document.createElement("tr");
    Object.keys(data[0]).forEach(column => {
        const th = document.createElement("th");
        th.textContent = column;
        th.style.border = "1px solid rgb(41 39 68)";
        th.style.padding = "8px 2px";
        th.style.backgroundColor = "#5a5693";
        th.style.color = "#fff";
        th.style.fontSize = "15px"
        th.style.fontWeight = "500"
        th.style.textAlign = "center";
        th.style.fontFamily = "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif";
        headerRow.appendChild(th);
    });
    table.appendChild(headerRow);

    // Populate table rows with data
    data.forEach(row => {
        const rowElement = document.createElement("tr");
        Object.values(row).forEach(cell => {
            const cellElement = document.createElement("td");
            cellElement.textContent = cell || "";  // Show empty if cell is undefined
            cellElement.style.border = "1px solid rgb(41 39 68)";
            cellElement.style.padding = "8px";
            rowElement.appendChild(cellElement);
        });
        table.appendChild(rowElement);
    });

    container.appendChild(table);
}

function downloadUpdatedFile() {
    const ws = XLSX.utils.json_to_sheet(mergedData);
    const csvData = XLSX.utils.sheet_to_csv(ws); // Convert worksheet to CSV format

    // Create a Blob from the CSV data
    const blob = new Blob([csvData], { type: 'text/csv;charset=utf-8;' });

    // Create a download link
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'MergedFile.csv'; // Set the file name
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);

    setTimeout(() => {
        location.reload(); // Reload the page after a slight delay
    }, 500);
}

function normalize(url) {
    if (!url) return "";

    // If url is not a string, convert it
    if (typeof url !== "string") {
        url = String(url);
    }

    return (url || '')
        .trim()
        .replace(/^https?:\/\//, '')
        .replace(/^www\./, '')
        .replace(/\/$/, '');
}

function setStatus(message) {
    const statusBox = document.getElementById("importStatus");
    if (statusBox) statusBox.innerText = message;
}

// XLSX and Supabase CDN must be loaded
document.getElementById('updateButton').addEventListener('click', handleUpdateDB);

async function handleUpdateDB() {
    const fileInput = document.getElementById('websiteExcelFile');
    const collection = "Website_origin_category";   // your combined table

    if (!fileInput.files[0]) {
        alert('Please select a file first!');
        return;
    }

    setStatus("Reading Excel file...");

    // Read file
    const data = await fileInput.files[0].arrayBuffer();
    const workbook = XLSX.read(data, { type: 'array' });
    const firstSheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[firstSheetName];
    const rows = XLSX.utils.sheet_to_json(sheet);

    console.log("Total Excel rows loaded:", rows.length);

    // Fetch all existing URLs
    let allExisting = [];
    let start = 0, batchSize = 1000, hasMore = true;

    while (hasMore) {
        const { data, error } = await supabase
            .from(collection)
            .select('url')
            .range(start, start + batchSize - 1);

        if (error) {
            alert('Error fetching existing data');
            return;
        }

        if (data) allExisting = allExisting.concat(data);
        hasMore = (data && data.length === batchSize);
        start += batchSize;
    }

    const existingUrlsSet = new Set(allExisting.map(item => normalize(item.Website)));

    // Prepare rows (common table)
    const formattedRows = rows
        .filter(row => row.Website && !existingUrlsSet.has(normalize(row.Website)))
        .map(row => ({
            url: String(row.Website).trim(),
            origin: row.Origin ? String(row.Origin).trim() : null,
            Category: row.Category ? String(row.Category).trim() : null
        }));

    console.log("Rows to insert in DB:", formattedRows.length);

    if (formattedRows.length === 0) {
        alert('All URLs already exist! No new unique rows to import.');
        return;
    }

    setStatus("Importing data into database...");

    // Insert to Supabase
    const { error } = await supabase.from(collection).insert(formattedRows);

    if (error) {
        alert('Error updating Supabase: ' + error.message);
    } else {
        alert(`Imported ${formattedRows.length} NEW unique rows into ${collection}!`);
        fileInput.value = '';
    }
}

window.handleMergeDownload = handleMergeDownload;

async function handleMergeDownload() {
    console.log("handleMergeDownload triggered");
    showProgressBar(); // SHOW BAR as soon as merge starts!
    try {
        const button = document.getElementById('mergeDownloadBtn');
        if (!isMerged) {
            const success = await previewData();
            if (!success) {
                hideProgressBar();
                return;
            }
            isMerged = true;
            button.textContent = 'Download';
            button.classList.add('ready-to-download');
        } else {
            downloadUpdatedFile();
            setTimeout(() => {
                button.textContent = 'Merge & Preview Data';
                isMerged = false;
            }, 1500);
        }
    } finally {
        hideProgressBar(); // HIDE BAR as soon as merge/preview completes!
    }
}

// document.getElementById('mergeDownloadBtn').addEventListener('click', handleMergeDownload);

// document.getElementById("mergeDownloadBtn").addEventListener("click", () => console.log("Button clicked"));

function showProgressBar() {
    document.getElementById("top-progress-bar").style.display = "block";
}
function hideProgressBar() {
    document.getElementById("top-progress-bar").style.display = "none";
}

// Load the static JSON once when the page loads
loadStaticJson();