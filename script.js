// --- CONFIGURAZIONE ---

// IL TUO LINK ORIGINALE (Quello corto)
const FILE_LINK = "https://1drv.ms/x/c/ac1a912c65f087d9/IQSCrL3EW_MNQJi_FLvY8KNJAXXS-7KsHuornWAqYgAoNnE";

// --- ELEMENTI DOM ---
const inputCodice = document.getElementById('input-codice');
const btnPlay = document.getElementById('btn-play');
const btnReset = document.getElementById('btn-reset');
const loadingOverlay = document.getElementById('loading-overlay');

const phRisultati = document.getElementById('placeholder-risultati');
const listaDati = document.getElementById('lista-dati');
const errCodice = document.getElementById('error-codice');

const phImmagine = document.getElementById('placeholder-immagine');
const imgResult = document.getElementById('result-image');
const errImmagine = document.getElementById('error-immagine');

// --- FUNZIONI ---

function resetApp() {
    inputCodice.value = "";
    phRisultati.classList.remove('hidden');
    listaDati.classList.add('hidden');
    errCodice.classList.add('hidden');
    listaDati.innerHTML = "";
    phImmagine.classList.remove('hidden');
    imgResult.classList.add('hidden');
    errImmagine.classList.add('hidden');
    imgResult.src = "";
}

// Funzione di supporto per il download sicuro
async function downloadExcel(url, methodDescription) {
    try {
        console.log(`Tentativo ${methodDescription}: ${url}`);
        
        // Usa CodeTabs per bypassare i blocchi CORS e seguire i redirect
        const proxyUrl = `https://api.codetabs.com/v1/proxy?quest=${encodeURIComponent(url)}`;
        
        const response = await fetch(proxyUrl, { method: 'GET', cache: 'no-store' });
        
        if (!response.ok) return null; // Errore HTTP

        const arrayBuffer = await response.arrayBuffer();
        
        // CONTROLLO FONDAMENTALE:
        // Se il primo byte è 60 (<), abbiamo scaricato una pagina HTML (Errore).
        // Se il primo byte è 80 (P), è un file ZIP/Excel (Successo).
        const firstByte = new Uint8Array(arrayBuffer)[0];
        
        if (firstByte === 60) {
            console.warn(`Fallito ${methodDescription}: Ricevuto HTML invece di Excel.`);
            return null;
        }

        return arrayBuffer; // Successo! Restituisce il file binario
    } catch (e) {
        console.warn(`Errore ${methodDescription}:`, e.message);
        return null;
    }
}

async function eseguiRicerca() {
    const rawCode = inputCodice.value.trim();
    if (!rawCode) return; 

    const searchCode = rawCode.replace(/\s+/g, '').toUpperCase(); 
    loadingOverlay.classList.remove('hidden');

    try {
        let excelData = null;

        // --- STRATEGIA 1: API UFFICIALE (Base64) ---
        // Converte il link in una richiesta API formale
        let base64 = btoa(FILE_LINK).replace(/\//g, '_').replace(/\+/g, '-').replace(/=+$/, '');
        const url1 = `https://api.onedrive.com/v1.0/shares/u!${base64}/root/content`;
        
        if (!excelData) excelData = await downloadExcel(url1, "API OneDrive");

        // --- STRATEGIA 2: GRAPH API (Moderna) ---
        // A volte api.onedrive.com fallisce, proviamo graph.microsoft.com
        const url2 = `https://graph.microsoft.com/v1.0/shares/u!${base64}/driveItem/content`;
        
        if (!excelData) excelData = await downloadExcel(url2, "Graph API");

        // --- STRATEGIA 3: DOWNLOAD DIRETTO FORZATO ---
        // Modifica manuale del link corto
        const url3 = FILE_LINK.replace('/x/c/', '/download/c/');
        
        if (!excelData) excelData = await downloadExcel(url3, "Download Diretto");


        // SE TUTTO FALLISCE
        if (!excelData) {
            throw new Error("Impossibile scaricare il file. Microsoft sta bloccando le connessioni esterne. Riprova più tardi.");
        }

        // LETTURA EXCEL
        const workbook = XLSX.read(excelData, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        
        // Converte in matrice (array di array)
        const matrix = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

        if (!matrix || matrix.length === 0) throw new Error("File Excel vuoto.");

        // Ritardo estetico
        setTimeout(() => {
            elaboraDati(matrix, searchCode);
            loadingOverlay.classList.add('hidden');
        }, 3000); 

    } catch (error) {
        loadingOverlay.classList.add('hidden');
        alert("Errore: " + error.message);
    }
}

function elaboraDati(matrix, searchCode) {
    errCodice.classList.add('hidden');
    errImmagine.classList.add('hidden');
    phRisultati.classList.add('hidden');
    phImmagine.classList.add('hidden');
    listaDati.innerHTML = "";

    let recordTrovato = null;
    let headers = [];

    // Scorre le righe
    for (let i = 0; i < matrix.length; i++) {
        const row = matrix[i];
        for (let j = 0; j < row.length; j++) {
            let cell = String(row[j]).toUpperCase().replace(/\s+/g, '');
            if (cell === searchCode) {
                recordTrovato = row;
                // Gestione Intestazioni
                if (i > 0) headers = matrix[0]; 
                else headers = row.map((_, idx) => `Dato ${idx + 1}`);
                break;
            }
        }
        if (recordTrovato) break;
    }

    if (!recordTrovato) {
        listaDati.classList.add('hidden');
        errCodice.classList.remove('hidden');
        imgResult.classList.add('hidden');
        errImmagine.classList.remove('hidden');
        return;
    }

    // Mostra Dati
    listaDati.classList.remove('hidden');
    let imageLinkFound = "";
    const maxLen = Math.max(headers.length, recordTrovato.length);

    for (let k = 0; k < maxLen; k++) {
        let key = headers[k] || `Colonna ${k+1}`;
        let val = recordTrovato[k];

        if (!val || String(val).trim() === "") continue;

        let keyLower = String(key).toLowerCase();
        
        // Cerca Immagine
        if (keyLower.includes('immagine') || keyLower.includes('link') || keyLower.includes('foto') || keyLower.includes('url')) {
            imageLinkFound = val;
            continue;
        }

        // Non mostrare il codice stesso
        let valClean = String(val).toUpperCase().replace(/\s+/g, '');
        if (valClean === searchCode) continue;

        const div = document.createElement('div');
        div.className = 'data-item';
        div.innerText = `- ${key}: ${val}`;
        listaDati.appendChild(div);
    }

    // Gestione Immagine
    if (imageLinkFound) {
        let imgUrl = String(imageLinkFound);
        if (imgUrl.includes("drive.google.com")) {
             imgUrl = imgUrl.replace("/view", "/preview").replace("open?id=", "uc?id=");
        }
        imgResult.src = imgUrl;
        imgResult.classList.remove('hidden');
        imgResult.onerror = () => { imgResult.classList.add('hidden'); errImmagine.classList.remove('hidden'); };
    } else {
        errImmagine.classList.remove('hidden');
    }
}

// Event Listeners
btnPlay.addEventListener('click', eseguiRicerca);
btnReset.addEventListener('click', resetApp);
inputCodice.addEventListener('keypress', (e) => {
    if (e.key === 'Enter') eseguiRicerca();
});