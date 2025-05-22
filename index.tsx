
// Ensure Mammoth and XLSX types are available if developers use TypeScript,
// by declaring them if they are loaded from CDN.
declare var mammoth: any;
declare var XLSX: any;
declare var html2pdf: any;
declare var pdfjsLib: any;
// Fix: Provide a more specific type declaration for the docx global variable
// to help TypeScript resolve types like docx.Paragraph when used in type annotations.
declare var docx: {
    Paragraph: new (options?: any) => any; // Used as new docx.Paragraph({})
    PageBreak: new (options?: any) => any; // Used as new docx.PageBreak()
    Document: new (options?: any) => any;  // Used as new docx.Document({})
    Packer: {
        toBlob: (document: any) => Promise<Blob>; // Used as docx.Packer.toBlob(doc)
    };
};

const fileInput = document.getElementById('fileInput') as HTMLInputElement;
const viewer = document.getElementById('viewer') as HTMLDivElement;
const fileNameDisplay = document.getElementById('fileName') as HTMLSpanElement;
const fileTypeDisplay = document.getElementById('fileType') as HTMLSpanElement;
const fileInfoSection = document.getElementById('file-info-section') as HTMLElement;
const loadingIndicator = document.getElementById('loading-indicator') as HTMLDivElement;
const fileInputLabel = document.querySelector('.file-input-label') as HTMLLabelElement;

const recentFilesListElement = document.getElementById('recent-files-list') as HTMLUListElement;
const noRecentFilesElement = document.getElementById('no-recent-files') as HTMLParagraphElement;
const clearHistoryButton = document.getElementById('clear-history-button') as HTMLButtonElement;
const recentDocumentsSection = document.getElementById('recent-documents-section') as HTMLElement;

// Conversion Tools Elements
const convertDocxToPdfButton = document.getElementById('convertDocxToPdfButton') as HTMLButtonElement;
const convertPdfToDocxButton = document.getElementById('convertPdfToDocxButton') as HTMLButtonElement;
const convertPdfToExcelButton = document.getElementById('convertPdfToExcelButton') as HTMLButtonElement;
const convertExcelToPdfButton = document.getElementById('convertExcelToPdfButton') as HTMLButtonElement;
const conversionStatusElement = document.getElementById('conversion-status') as HTMLParagraphElement;


const MAX_RECENT_FILES = 5;
const RECENT_FILES_KEY = 'recentFileViewerFiles';

interface RecentFile {
    name: string;
    typeIdentifier: 'pdf' | 'docx' | 'excel';
    data: string; // Base64 encoded ArrayBuffer for docx/excel, or DataURL for PDF
    originalMimeType: string;
}

interface CurrentViewedFile {
    name: string;
    typeIdentifier: 'pdf' | 'docx' | 'excel';
    rawData: ArrayBuffer | string; // ArrayBuffer for docx/excel, DataURL string for PDF
    originalMimeType: string;
}
let currentViewedFile: CurrentViewedFile | null = null;
let currentPdfObjectUrl: string | null = null;

// Initialize PDF.js worker
if (typeof pdfjsLib !== 'undefined') {
    pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
}


// --- Helper Functions ---
function arrayBufferToBase64(buffer: ArrayBuffer): string {
    let binary = '';
    const bytes = new Uint8Array(buffer);
    const len = bytes.byteLength;
    for (let i = 0; i < len; i++) {
        binary += String.fromCharCode(bytes[i]);
    }
    return window.btoa(binary);
}

function base64ToArrayBuffer(base64: string): ArrayBuffer {
    const binary_string = window.atob(base64);
    const len = binary_string.length;
    const bytes = new Uint8Array(len);
    for (let i = 0; i < len; i++) {
        bytes[i] = binary_string.charCodeAt(i);
    }
    return bytes.buffer;
}

function dataURLtoBlob(dataUrl: string): Blob | null {
    try {
        const parts = dataUrl.split(',');
        if (parts.length !== 2) {
            console.error("Invalid Data URL: does not have two parts");
            return null;
        }
        const metaPart = parts[0];
        const base64Data = parts[1];

        const mimeMatch = metaPart.match(/:(.*?);/);
        if (!mimeMatch || mimeMatch.length < 2) {
            console.error("Invalid Data URL: MIME type not found in meta part");
            return null;
        }
        const mimeType = mimeMatch[1];

        const byteString = atob(base64Data);
        const ab = new ArrayBuffer(byteString.length);
        const ia = new Uint8Array(ab);
        for (let i = 0; i < byteString.length; i++) {
            ia[i] = byteString.charCodeAt(i);
        }
        return new Blob([ab], { type: mimeType });
    } catch (e) {
        console.error("Error converting Data URL to Blob:", e);
        return null;
    }
}

function revokeCurrentPdfUrl(): void {
    if (currentPdfObjectUrl) {
        URL.revokeObjectURL(currentPdfObjectUrl);
        currentPdfObjectUrl = null;
    }
}

function downloadFile(blob: Blob, filename: string): void {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}


// --- UI Update Functions ---
function showLoading(isLoading: boolean): void {
    if (isLoading) {
        revokeCurrentPdfUrl();
        loadingIndicator.classList.remove('hidden');
        viewer.innerHTML = '';
        viewer.classList.add('hidden');
    } else {
        loadingIndicator.classList.add('hidden');
        viewer.classList.remove('hidden');
    }
}

function displayError(message: string): void {
    revokeCurrentPdfUrl();
    viewer.innerHTML = `<p style="color: red; text-align: center;">Error: ${message}</p>`;
    viewer.classList.remove('hidden');
    loadingIndicator.classList.add('hidden');
    fileInfoSection.classList.add('hidden');
    currentViewedFile = null; // Reset current file on error
    updateConversionButtonsState(); // Update buttons accordingly
}

function displayFileInfo(name: string, type: string): void {
    fileNameDisplay.textContent = name;
    fileTypeDisplay.textContent = type || 'N/A';
    fileInfoSection.classList.remove('hidden');
}

function updateConversionButtonsState(): void {
    convertDocxToPdfButton.disabled = !(currentViewedFile?.typeIdentifier === 'docx');
    convertPdfToDocxButton.disabled = !(currentViewedFile?.typeIdentifier === 'pdf');
    convertPdfToExcelButton.disabled = !(currentViewedFile?.typeIdentifier === 'pdf');
    convertExcelToPdfButton.disabled = !(currentViewedFile?.typeIdentifier === 'excel');
}

function setConversionStatus(message: string, isError: boolean = false): void {
    conversionStatusElement.textContent = message;
    conversionStatusElement.style.color = isError ? 'red' : 'inherit';
    if (!message) {
        conversionStatusElement.classList.add('hidden');
    } else {
        conversionStatusElement.classList.remove('hidden');
    }
}


// --- Rendering Logic ---
function renderPdfData(pdfDataUrl: string, fileName: string): void {
    revokeCurrentPdfUrl();
    const blob = dataURLtoBlob(pdfDataUrl);
    if (!blob) {
        displayError(`Could not process PDF data for ${fileName}. Invalid or corrupt Data URL.`);
        return;
    }
    currentPdfObjectUrl = URL.createObjectURL(blob);
    viewer.innerHTML = `<iframe src="${currentPdfObjectUrl}" type="application/pdf" width="100%" height="580px" style="border: none;" title="${fileName} preview"></iframe>`;
    displayFileInfo(fileName, 'application/pdf');
    showLoading(false);
}

function renderDocxData(arrayBuffer: ArrayBuffer, fileName: string): void {
    if (typeof mammoth === 'undefined') {
        displayError('Mammoth.js library is not loaded. Cannot preview DOCX files.');
        return;
    }
    revokeCurrentPdfUrl();
    mammoth.convertToHtml({ arrayBuffer })
        .then((result: { value: string; messages: any[] }) => {
            viewer.innerHTML = result.value;
            displayFileInfo(fileName, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
            showLoading(false);
        })
        .catch((err: any) => {
            console.error('Error converting DOCX from data:', err);
            displayError(`Could not process DOCX data for ${fileName}. ${err.message || ''}`);
        });
}

function renderExcelData(arrayBuffer: ArrayBuffer, fileName: string, originalMimeType: string): void {
    if (typeof XLSX === 'undefined') {
        displayError('SheetJS (XLSX) library is not loaded. Cannot preview Excel files.');
        return;
    }
    revokeCurrentPdfUrl();
    try {
        const data = new Uint8Array(arrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        if (!firstSheetName) {
            displayError(`Excel file '${fileName}' appears to be empty or has no sheets.`);
            return;
        }
        const worksheet = workbook.Sheets[firstSheetName];
        const html = XLSX.utils.sheet_to_html(worksheet);
        viewer.innerHTML = html;
        displayFileInfo(fileName, originalMimeType || (fileName.endsWith('.xlsx') ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' : 'application/vnd.ms-excel'));
        showLoading(false);
    } catch (err: any) {
        console.error(`Error processing Excel file '${fileName}':`, err);
        displayError(`Could not process Excel file '${fileName}'. ${err.message || ''}`);
    }
}

// --- Recent Files Logic ---
function getRecentFiles(): RecentFile[] {
    const storedFiles = localStorage.getItem(RECENT_FILES_KEY);
    return storedFiles ? JSON.parse(storedFiles) : [];
}

function saveRecentFiles(files: RecentFile[]): void {
    localStorage.setItem(RECENT_FILES_KEY, JSON.stringify(files));
}

function addFileToRecents(fileName: string, typeIdentifier: 'pdf' | 'docx' | 'excel', data: string, originalMimeType: string): void {
    let recents = getRecentFiles();
    recents = recents.filter(f => f.name !== fileName);
    const newRecentFile: RecentFile = { name: fileName, typeIdentifier, data, originalMimeType };
    recents.unshift(newRecentFile);
    if (recents.length > MAX_RECENT_FILES) {
        recents.pop();
    }
    saveRecentFiles(recents);
    populateRecentFilesList();
}

function populateRecentFilesList(): void {
    const recents = getRecentFiles();
    recentFilesListElement.innerHTML = '';
    if (recents.length === 0) {
        noRecentFilesElement.classList.remove('hidden');
        clearHistoryButton.classList.add('hidden');
        if (recentDocumentsSection) recentDocumentsSection.classList.add('hidden');
    } else {
        noRecentFilesElement.classList.add('hidden');
        clearHistoryButton.classList.remove('hidden');
        if (recentDocumentsSection) recentDocumentsSection.classList.remove('hidden');
        recents.forEach((file, index) => {
            const listItem = document.createElement('li');
            listItem.textContent = file.name;
            listItem.setAttribute('role', 'button');
            listItem.setAttribute('tabindex', '0');
            listItem.setAttribute('aria-label', `Open ${file.name}`);
            listItem.dataset.index = index.toString();
            listItem.addEventListener('click', () => loadRecentFile(file));
            listItem.addEventListener('keydown', (event) => {
                if (event.key === 'Enter' || event.key === ' ') {
                    event.preventDefault();
                    loadRecentFile(file);
                }
            });
            recentFilesListElement.appendChild(listItem);
        });
    }
}

function loadRecentFile(file: RecentFile): void {
    showLoading(true);
    setConversionStatus('');
    try {
        if (file.typeIdentifier === 'pdf') {
            currentViewedFile = { name: file.name, typeIdentifier: 'pdf', rawData: file.data, originalMimeType: file.originalMimeType };
            renderPdfData(file.data, file.name);
        } else if (file.typeIdentifier === 'docx') {
            const arrayBuffer = base64ToArrayBuffer(file.data);
            currentViewedFile = { name: file.name, typeIdentifier: 'docx', rawData: arrayBuffer, originalMimeType: file.originalMimeType };
            renderDocxData(arrayBuffer, file.name);
        } else if (file.typeIdentifier === 'excel') {
            const arrayBuffer = base64ToArrayBuffer(file.data);
            currentViewedFile = { name: file.name, typeIdentifier: 'excel', rawData: arrayBuffer, originalMimeType: file.originalMimeType };
            renderExcelData(arrayBuffer, file.name, file.originalMimeType);
        }
        updateConversionButtonsState();
    } catch (error: any) {
        console.error("Error loading recent file:", error);
        displayError(`Could not load recent file ${file.name}. ${error.message || ''}`);
        currentViewedFile = null;
        updateConversionButtonsState();
    }
}

function clearRecentFilesHistory(): void {
    localStorage.removeItem(RECENT_FILES_KEY);
    populateRecentFilesList();
    revokeCurrentPdfUrl();
    viewer.innerHTML = '<p>Select a PDF, DOCX, XLSX, or XLS file to view its content here.</p>';
    fileInfoSection.classList.add('hidden');
    loadingIndicator.classList.add('hidden');
    currentViewedFile = null;
    updateConversionButtonsState();
    setConversionStatus('');
}

// --- Processing Logic for New Files ---
function processAndRenderFile(file: File): void {
    const fileName = file.name;
    const fileExtension = fileName.split('.').pop()?.toLowerCase();
    const reader = new FileReader();

    showLoading(true);
    setConversionStatus('');

    reader.onload = (e) => {
        const result = e.target?.result;
        if (!result) {
            displayError('File content could not be read.');
            return;
        }

        let typeId: 'pdf' | 'docx' | 'excel' | undefined;
        let rawDataForCurrent: ArrayBuffer | string | undefined;

        if (file.type === 'application/pdf' || fileExtension === 'pdf') {
            typeId = 'pdf';
            rawDataForCurrent = result as string;
            renderPdfData(rawDataForCurrent, fileName);
            addFileToRecents(fileName, 'pdf', rawDataForCurrent, file.type);
        } else if (fileExtension === 'docx') {
            typeId = 'docx';
            rawDataForCurrent = result as ArrayBuffer;
            renderDocxData(rawDataForCurrent, fileName);
            addFileToRecents(fileName, 'docx', arrayBufferToBase64(rawDataForCurrent), file.type);
        } else if (fileExtension === 'xlsx' || fileExtension === 'xls') {
            typeId = 'excel';
            rawDataForCurrent = result as ArrayBuffer;
            renderExcelData(rawDataForCurrent, fileName, file.type);
            addFileToRecents(fileName, 'excel', arrayBufferToBase64(rawDataForCurrent), file.type);
        } else {
            displayError(`Unsupported file type: .${fileExtension}. Please select a PDF, DOCX, XLSX, or XLS file.`);
        }

        if (typeId && rawDataForCurrent !== undefined) {
            currentViewedFile = { name: fileName, typeIdentifier: typeId, rawData: rawDataForCurrent, originalMimeType: file.type };
        } else {
            currentViewedFile = null;
        }
        updateConversionButtonsState();
    };

    reader.onerror = () => {
        displayError(`Could not read file: ${fileName}.`);
        currentViewedFile = null;
        updateConversionButtonsState();
    };

    if (file.type === 'application/pdf' || fileExtension === 'pdf') {
        reader.readAsDataURL(file);
    } else if (fileExtension === 'docx' || fileExtension === 'xlsx' || fileExtension === 'xls') {
        reader.readAsArrayBuffer(file);
    } else {
        showLoading(false);
        currentViewedFile = null;
        updateConversionButtonsState();
        if (fileExtension === 'doc') {
            displayError('.doc files are not directly supported for preview. Please convert to .docx.');
        } else {
            displayError(`Unsupported file type: .${fileExtension}. Please select a PDF, DOCX, XLSX, or XLS file.`);
        }
    }
}

// --- Conversion Functions ---
async function handleDocxToPdf(): Promise<void> {
    if (!currentViewedFile || currentViewedFile.typeIdentifier !== 'docx' || !(currentViewedFile.rawData instanceof ArrayBuffer)) {
        setConversionStatus('No DOCX file loaded or data is invalid.', true);
        return;
    }
    if (typeof mammoth === 'undefined' || typeof html2pdf === 'undefined') {
        setConversionStatus('Conversion library (Mammoth or html2pdf) not loaded.', true);
        return;
    }
    setConversionStatus('Converting DOCX to PDF...');
    try {
        const { value: html } = await mammoth.convertToHtml({ arrayBuffer: currentViewedFile.rawData });
        const options = {
            margin: 1,
            filename: `${currentViewedFile.name.split('.').slice(0, -1).join('.')}.pdf`,
            image: { type: 'jpeg', quality: 0.98 },
            html2canvas: { scale: 2, useCORS: true }, // useCORS for external images if any
            jsPDF: { unit: 'in', format: 'letter', orientation: 'portrait' }
        };
        await html2pdf().from(html).set(options).save();
        setConversionStatus('DOCX to PDF conversion complete. Downloading...');
    } catch (error: any) {
        console.error('Error converting DOCX to PDF:', error);
        setConversionStatus(`DOCX to PDF conversion failed: ${error.message}`, true);
    }
}

async function handlePdfToDocx(): Promise<void> {
    if (!currentViewedFile || currentViewedFile.typeIdentifier !== 'pdf' || typeof currentViewedFile.rawData !== 'string') {
        setConversionStatus('No PDF file loaded or data is invalid.', true);
        return;
    }
    if (typeof pdfjsLib === 'undefined' || typeof docx === 'undefined') {
        setConversionStatus('Conversion library (PDF.js or docx) not loaded.', true);
        return;
    }
    setConversionStatus('Converting PDF to DOCX...');
    try {
        const pdfData = atob(currentViewedFile.rawData.substring(currentViewedFile.rawData.indexOf(',') + 1));
        const pdfArrayBuffer = new Uint8Array(pdfData.length);
        for (let i = 0; i < pdfData.length; i++) {
            pdfArrayBuffer[i] = pdfData.charCodeAt(i);
        }

        const pdf = await pdfjsLib.getDocument({ data: pdfArrayBuffer }).promise;
        const numPages = pdf.numPages;
        const paragraphs: docx.Paragraph[] = [];

        for (let i = 1; i <= numPages; i++) {
            const page = await pdf.getPage(i);
            const textContent = await page.getTextContent();
            const pageText = textContent.items.map((item: any) => item.str).join(' ');
            if (pageText.trim()) {
                 paragraphs.push(new docx.Paragraph({ text: pageText }));
            }
            if (i < numPages) { // Add a page break except for the last page's content
                 paragraphs.push(new docx.Paragraph({ children: [new docx.PageBreak()] }));
            }
        }
        
        if (paragraphs.length === 0) {
             paragraphs.push(new docx.Paragraph({ text: "No text content found in PDF." }));
        }

        const doc = new docx.Document({
            sections: [{
                properties: {},
                children: paragraphs,
            }],
        });

        const blob = await docx.Packer.toBlob(doc);
        downloadFile(blob, `${currentViewedFile.name.split('.').slice(0, -1).join('.')}.docx`);
        setConversionStatus('PDF to DOCX conversion complete. Downloading...');
    } catch (error: any) {
        console.error('Error converting PDF to DOCX:', error);
        setConversionStatus(`PDF to DOCX conversion failed: ${error.message}`, true);
    }
}

async function handlePdfToExcel(): Promise<void> {
    if (!currentViewedFile || currentViewedFile.typeIdentifier !== 'pdf' || typeof currentViewedFile.rawData !== 'string') {
        setConversionStatus('No PDF file loaded or data is invalid.', true);
        return;
    }
    if (typeof pdfjsLib === 'undefined' || typeof XLSX === 'undefined') {
        setConversionStatus('Conversion library (PDF.js or XLSX) not loaded.', true);
        return;
    }
    setConversionStatus('Converting PDF to Excel...');
    try {
        const pdfData = atob(currentViewedFile.rawData.substring(currentViewedFile.rawData.indexOf(',') + 1));
        const pdfArrayBuffer = new Uint8Array(pdfData.length);
        for (let i = 0; i < pdfData.length; i++) {
            pdfArrayBuffer[i] = pdfData.charCodeAt(i);
        }

        const pdf = await pdfjsLib.getDocument({ data: pdfArrayBuffer }).promise;
        const numPages = pdf.numPages;
        const excelData: { Page: number; Line: number; Text: string }[] = [];
        let lineNumGlobal = 1;

        for (let i = 1; i <= numPages; i++) {
            const page = await pdf.getPage(i);
            const textContent = await page.getTextContent();
            
            // Basic line reconstruction: group items by y-coordinate (transform[5])
            const lines: { y: number, texts: string[] }[] = [];
            textContent.items.forEach((item: any) => {
                const y = parseFloat(item.transform[5].toFixed(2)); // Approximate y
                let line = lines.find(l => Math.abs(l.y - y) < 5); // Tolerance for y-coordinate
                if (!line) {
                    line = { y, texts: [] };
                    lines.push(line);
                }
                line.texts.push(item.str);
            });
            // Sort lines by y-coordinate (top to bottom)
            lines.sort((a,b) => b.y - a.y);


            lines.forEach(line => {
                 const lineText = line.texts.join(' '); // Join text items in the same "line"
                 if(lineText.trim()){
                    excelData.push({ Page: i, Line: lineNumGlobal++, Text: lineText });
                 }
            });
             if (lines.length === 0 && textContent.items.length > 0) { // Fallback if no lines formed but text exists
                const pageText = textContent.items.map((item: any) => item.str).join(' ');
                 if(pageText.trim()){
                    excelData.push({ Page: i, Line: lineNumGlobal++, Text: pageText });
                 }
            }
        }
        
        if (excelData.length === 0) {
             excelData.push({ Page: 1, Line: 1, Text: "No text content found in PDF." });
        }

        const worksheet = XLSX.utils.json_to_sheet(excelData);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Extracted Text');
        
        // XLSX.writeFile is for Node.js, for browser use write to get ArrayBuffer then Blob
        const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        
        downloadFile(blob, `${currentViewedFile.name.split('.').slice(0, -1).join('.')}.xlsx`);
        setConversionStatus('PDF to Excel conversion complete. Downloading...');
    } catch (error: any) {
        console.error('Error converting PDF to Excel:', error);
        setConversionStatus(`PDF to Excel conversion failed: ${error.message}`, true);
    }
}

async function handleExcelToPdf(): Promise<void> {
    if (!currentViewedFile || currentViewedFile.typeIdentifier !== 'excel' || !(currentViewedFile.rawData instanceof ArrayBuffer)) {
        setConversionStatus('No Excel file loaded or data is invalid.', true);
        return;
    }
    if (typeof XLSX === 'undefined' || typeof html2pdf === 'undefined') {
        setConversionStatus('Conversion library (XLSX or html2pdf) not loaded.', true);
        return;
    }
    setConversionStatus('Converting Excel to PDF...');
    try {
        const data = new Uint8Array(currentViewedFile.rawData);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        if (!firstSheetName) {
            setConversionStatus('Excel file is empty or has no sheets.', true);
            return;
        }
        const worksheet = workbook.Sheets[firstSheetName];
        const htmlTable = XLSX.utils.sheet_to_html(worksheet);

        // Wrap HTML table in a full HTML structure for better PDF rendering if needed, or use directly.
        // html2pdf can often handle table fragments well.
        const element = document.createElement('div');
        element.innerHTML = htmlTable;
        // Apply some basic styling to the table for PDF output
        const tableElement = element.querySelector('table');
        if (tableElement) {
            tableElement.style.borderCollapse = 'collapse';
            tableElement.style.width = '100%';
            const thtd = element.querySelectorAll('th, td');
            thtd.forEach(cell => {
                (cell as HTMLElement).style.border = '1px solid black';
                (cell as HTMLElement).style.padding = '5px';
            });
        }


        const options = {
            margin: 0.5,
            filename: `${currentViewedFile.name.split('.').slice(0, -1).join('.')}.pdf`,
            image: { type: 'jpeg', quality: 0.98 },
            html2canvas: { scale: 2, useCORS: true },
            jsPDF: { unit: 'in', format: 'letter', orientation: 'landscape' } // Landscape often better for Excel
        };
        await html2pdf().from(element).set(options).save();
        setConversionStatus('Excel to PDF conversion complete. Downloading...');
    } catch (error: any) {
        console.error('Error converting Excel to PDF:', error);
        setConversionStatus(`Excel to PDF conversion failed: ${error.message}`, true);
    }
}


// --- Initialization and Event Listeners ---
if (fileInput && viewer && fileNameDisplay && fileTypeDisplay && fileInfoSection && loadingIndicator && fileInputLabel && recentFilesListElement && noRecentFilesElement && clearHistoryButton && recentDocumentsSection && convertDocxToPdfButton && convertPdfToDocxButton && convertPdfToExcelButton && convertExcelToPdfButton && conversionStatusElement) {
    fileInput.addEventListener('change', (event) => {
        const files = (event.target as HTMLInputElement).files;
        if (!files || files.length === 0) {
            return;
        }
        processAndRenderFile(files[0]);
    });

    fileInputLabel.addEventListener('keydown', (event) => {
        if (event.key === 'Enter' || event.key === ' ') {
            event.preventDefault();
            fileInput.click();
        }
    });

    clearHistoryButton.addEventListener('click', clearRecentFilesHistory);

    // Conversion button listeners
    convertDocxToPdfButton.addEventListener('click', handleDocxToPdf);
    convertPdfToDocxButton.addEventListener('click', handlePdfToDocx);
    convertPdfToExcelButton.addEventListener('click', handlePdfToExcel);
    convertExcelToPdfButton.addEventListener('click', handleExcelToPdf);


    populateRecentFilesList();
    updateConversionButtonsState(); // Initial state
    setConversionStatus(''); // Clear any initial status

    if (viewer.innerHTML.trim() === '' || viewer.textContent?.trim() === 'Select a PDF, DOCX, XLSX, or XLS file to view its content here.') {
        if (getRecentFiles().length === 0) {
             viewer.innerHTML = '<p>Select a PDF, DOCX, XLSX, or XLS file to view its content here.</p>';
        }
    }

} else {
    console.error('One or more essential UI elements are missing from the DOM.');
    const appContainer = document.getElementById('app-container');
    if (appContainer) {
        appContainer.innerHTML = '<p style="color: red; text-align: center; padding: 20px;">Error: Application could not initialize. Essential elements missing.</p>';
    }
}
