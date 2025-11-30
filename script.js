// --- CONFIGURAZIONE ---

// ID del file Google Drive
const FILE_ID = "1MUxjFGP4l3tHTckFkW1DA5QaUJwex4xx";

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

async function scaricaConFallback(googleUrl) {
    const timestamp = new Date().getTime();
    const proxies = [
        `https://corsproxy.io/?${encodeURIComponent(googleUrl)}&t=${timestamp}`,
        `https://api.codetabs.com/v1/proxy?quest=${encodeURIComponent(googleUrl)}&t=${timestamp}`,
        `https://api.allorigins.win/raw?url=${encodeURIComponent(googleUrl)}&t=${timestamp}`
    ];

    let lastError = null;
    for (let proxyUrl of proxies) {
        try {
            const response = await fetch(proxyUrl, { method: 'GET', cache: 'no-store' });
            if (!response.ok) throw new Error(`HTTP ${response.status}`);
            const arrayBuffer = await response.arrayBuffer();
            const firstByte = new Uint8Array(arrayBuffer)[0];
            if (firstByte === 60) throw new Error("Ricevuto HTML invece di Excel");
            return arrayBuffer;
        } catch (e) {
            console.warn("Proxy fallito, provo il prossimo:", e);
            lastError = e;
        }
    }
    throw lastError || new Error("Impossibile scaricare il file Excel.");
}

async function eseguiRicerca() {
    const rawCode = inputCodice.value.trim();
    if (!rawCode) return; 

    const searchCode = rawCode.replace(/\s+/g, '').toUpperCase(); 
    loadingOverlay.classList.remove('hidden');

    try {
        const exportUrl = `https://docs.google.com/spreadsheets/d/${FILE_ID}/export?format=xlsx`;
        const arrayBuffer = await scaricaConFallback(exportUrl);

        if (!arrayBuffer) throw new Error("Download fallito.");

        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        
        const matrix = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

        if (!matrix || matrix.length === 0) throw new Error("File Excel vuoto.");

        setTimeout(() => {
            elaboraDati(matrix, searchCode);
            loadingOverlay.classList.add('hidden');
        }, 1500); 

    } catch (error) {
        console.error(error);
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

    // 1. CERCA IL CODICE
    for (let i = 0; i < matrix.length; i++) {
        const row = matrix[i];
        for (let j = 0; j < row.length; j++) {
            let cell = String(row[j]).toUpperCase().replace(/\s+/g, '');
            if (cell === searchCode) {
                recordTrovato = row;
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

    // 2. MOSTRA DATI
    listaDati.classList.remove('hidden');
    let imageLinkFound = "";
    
    const maxLen = Math.max(headers.length, recordTrovato.length);

    for (let k = 0; k < maxLen; k++) {
        let key = String(headers[k] || "").trim();
        let val = recordTrovato[k];

        if (!val || String(val).trim() === "") continue;

        let keyLower = key.toLowerCase();
        let valString = String(val);

        // RILEVA FOTO (Cerca "FOTO" o Link Drive)
        let isImageHeader = keyLower.includes('foto') || keyLower.includes('immagine');
        let isDriveLink = valString.includes('drive.google.com') || valString.includes('docs.google.com');

        if (isImageHeader || isDriveLink) {
            imageLinkFound = valString;
            continue; 
        }

        // FILTRO ASTERISCO (*)
        if (!key.startsWith('*')) {
            continue; 
        }

        let displayKey = key.replace('*', '').trim();
        let valClean = valString.toUpperCase().replace(/\s+/g, '');
        if (valClean === searchCode) continue;

        const div = document.createElement('div');
        div.className = 'data-item';
        div.innerText = `- ${displayKey}: ${val}`;
        listaDati.appendChild(div);
    }

    // 3. CARICA IMMAGINE (CORRETTO)
    if (imageLinkFound) {
        let imgUrl = imageLinkFound;
        
        // Estrazione ID precisa per link Google Drive
        let driveId = null;

        if (imgUrl.includes("/d/")) {
            // Formato: .../d/1c5DF5q-KvhG4JevQhMSHqJeYzz2_wwDx/view...
            let parts = imgUrl.split('/d/');
            if (parts.length > 1) {
                driveId = parts[1].split('/')[0];
            }
        } else if (imgUrl.includes("id=")) {
            // Formato: ...?id=1c5DF5q...
            driveId = imgUrl.split('id=')[1].split('&')[0];
        }

        if (driveId) {
            // Costruisce il link diretto ufficiale
            imgUrl = `https://drive.google.com/uc?export=view&id=${driveId}`;
            
            // TRUCCO: Imposta 'no-referrer' per evitare che Google blocchi l'immagine
            imgResult.referrerPolicy = "no-referrer";
        }

        console.log("URL Immagine generato:", imgUrl);

        imgResult.src = imgUrl;
        imgResult.classList.remove('hidden');
        
        imgResult.onerror = function() {
            console.error("Errore caricamento immagine (Forse permessi privati?):", imgUrl);
            imgResult.classList.add('hidden');
            errImmagine.classList.remove('hidden');
        };
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