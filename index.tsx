/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
*/
import { GoogleGenAI, Type } from '@google/genai';

// This is a workaround for the fact that the google global object is not
// available in the module scope.
declare global {
    const gapi: any;
    namespace google {
        namespace accounts {
            namespace oauth2 {
                interface TokenResponse {
                    access_token: string;
                    expires_in: number;
                    scope: string;
                    token_type: string;
                    error?: string;
                    error_description?: string;
                    error_uri?: string;
                }

                interface TokenClientConfig {
                    client_id: string;
                    scope: string;
                    callback: (tokenResponse: TokenResponse) => void;
                }

                interface OverridableTokenClientConfig {
                    prompt: string;
                }

                interface TokenClient {
                    requestAccessToken: (overrideConfig?: OverridableTokenClientConfig) => void;
                }

                function initTokenClient(config: TokenClientConfig): TokenClient;
            }
        }
    }
}


// --- DOM ELEMENT REFERENCES ---
const fileInput = document.getElementById('file-upload') as HTMLInputElement;
const fileLabel = document.getElementById('file-label');
const imagePreview = document.getElementById('image-preview') as HTMLImageElement;
const extractButton = document.getElementById('extract-button') as HTMLButtonElement;
const loader = document.getElementById('loader');
const loaderMessage = document.getElementById('loader-message') as HTMLParagraphElement;
const resultsBody = document.getElementById('results-body') as HTMLTableSectionElement;
const errorMessage = document.getElementById('error-message');
const fileUploader = document.querySelector('.file-uploader');
const connectGSheetButton = document.getElementById('connect-g-sheet-button') as HTMLButtonElement;
const exportButton = document.getElementById('export-button') as HTMLButtonElement;
const gSheetStatus = document.getElementById('g-sheet-status');
const exportResult = document.getElementById('export-result');

// --- STATE ---
let imageBase64: string | null = null;
let isLoading = false;
let extractedData: any[] = [];
let tokenClient: google.accounts.oauth2.TokenClient;
let loadingInterval: number | null = null;

// --- CONSTANTS ---
const GOOGLE_CLIENT_ID = process.env.GOOGLE_CLIENT_ID;
const GEMINI_API_KEY = process.env.API_KEY; 
const SCOPES = 'https://www.googleapis.com/auth/spreadsheets';
const LOADING_MESSAGES = [
    "Uploading and securing your document...",
    "Warming up the AI model...",
    "Scanning the document for text...",
    "Identifying headers and columns...",
    "Extracting individual records...",
    "Structuring the data into JSON format...",
    "Almost done, finalizing the results..."
];


// --- GEMINI SETUP ---
const ai = new GoogleGenAI({apiKey: GEMINI_API_KEY});

const textPart = {
    text: `From the provided document image, extract the data into a structured JSON format.
The document contains header information (like Sender GSTIN, Customer ID) and a table of records.
For each row in the table, create a JSON object. This object should include all the header information combined with the specific data from that row (Serial No, Article Number, etc.).
The final output should be an array of these JSON objects, where each object represents one complete record.`
};

const responseSchema = {
    type: Type.ARRAY,
    items: {
        type: Type.OBJECT,
        properties: {
            senderGstin: { type: Type.STRING },
            bookingOfficeGstin: { type: Type.STRING },
            customerId: { type: Type.STRING },
            contractId: { type: Type.STRING },
            customerName: { type: Type.STRING },
            bookingRefId: { type: Type.STRING },
            serialNo: { type: Type.STRING },
            articleNumber: { type: Type.STRING },
            productType: { type: Type.STRING },
            weight: { type: Type.NUMBER },
            senderName: { type: Type.STRING },
            receiverName: { type: Type.STRING },
            baseTariff: { type: Type.NUMBER },
            remarks: { type: Type.STRING },
            createdBy: { type: Type.STRING },
            createdOn: { type: Type.STRING },
            bulkReference: { type: Type.STRING }
        },
    }
};

const columnOrder = [
    'senderGstin', 'bookingOfficeGstin', 'customerId', 'contractId', 'customerName',
    'bookingRefId', 'serialNo', 'articleNumber', 'productType', 'weight',
    'senderName', 'receiverName', 'baseTariff', 'remarks', 'createdBy',
    'createdOn', 'bulkReference'
];

// --- GOOGLE API INITIALIZATION ---
function gapiLoaded() {
    gapi.load('client', initializeGapiClient);
}
async function initializeGapiClient() {
    await gapi.client.init({
      apiKey: GEMINI_API_KEY,
      discoveryDocs: ['https://sheets.googleapis.com/$discovery/rest?version=v4'],
    });
}
function gisLoaded() {
    tokenClient = google.accounts.oauth2.initTokenClient({
        client_id: GOOGLE_CLIENT_ID,
        scope: SCOPES,
        callback: (tokenResponse) => {
             if (tokenResponse.error) {
                displayError(`Google Auth Error: ${tokenResponse.error}`);
                return;
            }
            gSheetStatus.textContent = 'Status: Connected';
            updateButtonStates();
        },
    });
}
// Manually attach to window because of module scope
(window as any).gapiLoaded = gapiLoaded;
(window as any).gisLoaded = gisLoaded;

// --- FUNCTIONS ---

/**
 * Updates button states based on the current application state.
 */
function updateButtonStates() {
    const dataAvailable = extractedData.length > 0;
    const gSheetConnected = gapi?.client?.getToken() !== null;

    extractButton.disabled = isLoading || !imageBase64;
    connectGSheetButton.disabled = isLoading || gSheetConnected;
    exportButton.disabled = isLoading || !dataAvailable || !gSheetConnected;
}

/**
 * Handles the file selection, reads the file as base64, and updates the UI.
 */
function handleFileChange(event: Event) {
    const files = (event.target as HTMLInputElement).files;
    if (!files || files.length === 0) return;

    const file = files[0];
    const reader = new FileReader();

    reader.onloadend = () => {
        imageBase64 = reader.result as string;
        imagePreview.src = imageBase64;
        imagePreview.classList.remove('hidden');
        fileLabel.textContent = file.name;
        updateButtonStates();
    };

    reader.readAsDataURL(file);
}

/**
 * Renders the extracted data into the results table.
 */
function renderTable(data: any[]) {
    resultsBody.innerHTML = ''; // Clear previous results

    if (!data || data.length === 0) {
        const row = resultsBody.insertRow();
        const cell = row.insertCell();
        cell.colSpan = columnOrder.length;
        cell.textContent = 'No data extracted from the document.';
        cell.style.textAlign = 'center';
        return;
    }

    data.forEach(item => {
        const row = resultsBody.insertRow();
        columnOrder.forEach(key => {
            const cell = row.insertCell();
            cell.textContent = item[key] !== undefined && item[key] !== null ? String(item[key]) : 'N/A';
        });
    });
}

/**
 * Sets the loading state of the application.
 */
function setLoading(loading: boolean) {
    isLoading = loading;
    loader.classList.toggle('hidden', !loading);
    updateButtonStates();

    if (loading) {
        let messageIndex = 0;
        loaderMessage.textContent = LOADING_MESSAGES[0];
        loadingInterval = window.setInterval(() => {
            messageIndex = (messageIndex + 1) % LOADING_MESSAGES.length;
            loaderMessage.textContent = LOADING_MESSAGES[messageIndex];
        }, 2500);
    } else {
        if (loadingInterval) {
            clearInterval(loadingInterval);
            loadingInterval = null;
        }
    }
}

/**
 * Displays an error message to the user.
 */
function displayError(message: string) {
    errorMessage.textContent = message;
    errorMessage.classList.remove('hidden');
}

/**
 * Clears any previous results or errors from the UI.
 */
function clearOutput() {
    resultsBody.innerHTML = '';
    errorMessage.classList.add('hidden');
    exportResult.classList.add('hidden');
    exportResult.textContent = '';
    extractedData = [];
    updateButtonStates();
}

/**
 * Main function to call the Gemini API and extract data.
 */
async function handleExtractData() {
    if (!imageBase64 || isLoading) return;

    setLoading(true);
    clearOutput();

    try {
        const imagePart = {
            inlineData: {
                mimeType: imageBase64.match(/:(.*?);/)[1],
                data: imageBase64.split(',')[1],
            },
        };

        const response = await ai.models.generateContent({
            model: 'gemini-2.5-flash',
            contents: { parts: [textPart, imagePart] },
            config: {
                responseMimeType: "application/json",
                responseSchema: responseSchema,
            },
        });
        
        const data = JSON.parse(response.text);
        extractedData = data;
        renderTable(data);

    } catch (error) {
        console.error("Error extracting data:", error);
        displayError("An error occurred while extracting data. Please check the console for details.");
        extractedData = [];
    } finally {
        setLoading(false);
    }
}

/**
 * Initiates the Google Sheets authentication flow.
 */
function handleAuthClick() {
    if (gapi.client.getToken() === null) {
        tokenClient.requestAccessToken({prompt: 'consent'});
    } else {
        tokenClient.requestAccessToken({prompt: ''});
    }
}

/**
 * Handles exporting the extracted data to a new Google Sheet.
 */
async function handleExportClick() {
    if (extractedData.length === 0 || isLoading) return;

    setLoading(true);
    exportResult.classList.add('hidden');
    exportResult.className = 'hidden';

    try {
        // 1. Create a new spreadsheet
        const createResponse = await gapi.client.sheets.spreadsheets.create({
            properties: {
                title: `Document Extraction - ${new Date().toLocaleString()}`,
            },
        });

        const spreadsheetId = createResponse.result.spreadsheetId;
        const spreadsheetUrl = createResponse.result.spreadsheetUrl;
        
        // 2. Format data for the Sheets API
        const headerRow = columnOrder.map(key => {
            const spaced = key.replace(/([A-Z])/g, ' $1');
            return spaced.charAt(0).toUpperCase() + spaced.slice(1);
        });
        const dataRows = extractedData.map(item => columnOrder.map(key => item[key] ?? ''));
        const values = [headerRow, ...dataRows];

        // 3. Write data to the new sheet
        await gapi.client.sheets.spreadsheets.values.update({
            spreadsheetId: spreadsheetId,
            range: 'Sheet1!A1',
            valueInputOption: 'USER_ENTERED',
            resource: {
                values: values,
            },
        });

        // 4. Show success message with a link
        exportResult.innerHTML = `Successfully exported! <a href="${spreadsheetUrl}" target="_blank" rel="noopener noreferrer">Open Google Sheet</a>`;
        exportResult.classList.add('success');
        exportResult.classList.remove('hidden');

    } catch (error) {
        console.error("Error exporting to Google Sheets:", error);
        const gapiError = error as any;
        exportResult.textContent = `Export failed. ${gapiError.result?.error?.message || 'Check console for details.'}`;
        exportResult.classList.add('error');
        exportResult.classList.remove('hidden');
    } finally {
        setLoading(false);
    }
}


// --- EVENT LISTENERS ---
fileInput.addEventListener('change', handleFileChange);
extractButton.addEventListener('click', handleExtractData);
connectGSheetButton.addEventListener('click', handleAuthClick);
exportButton.addEventListener('click', handleExportClick);

// Drag and drop functionality
fileUploader.addEventListener('dragover', (event) => {
    event.preventDefault();
    fileUploader.classList.add('dragover');
});

fileUploader.addEventListener('dragleave', () => {
    fileUploader.classList.remove('dragover');
});

fileUploader.addEventListener('drop', (event) => {
    event.preventDefault();
    fileUploader.classList.remove('dragover');
    fileInput.files = (event as DragEvent).dataTransfer.files;
    const changeEvent = new Event('change');
    fileInput.dispatchEvent(changeEvent);
});

// Initial button state update on load
document.addEventListener('DOMContentLoaded', updateButtonStates);