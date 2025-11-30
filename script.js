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
    imgResult.onerror = null; // Reset gestore errori
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
            console.log("Tentativo download:", proxyUrl);
            const response = await fetch(proxyUrl, { method: 'GET', cache: 'no-store' });
            if (!response.ok) throw new Error(`HTTP ${response.status}`);
            const arrayBuffer = await response.arrayBuffer();
            const firstByte = new Uint8Array(arrayBuffer)[0];
            if (firstByte === 60) throw new Error("Ricevuto HTML invece di Excel");
            return arrayBuffer;
        } catch (e) {
            console.warn("Proxy fallito:", e);
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
    // Reset interfaccia
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

    // 2. ELABORA I DATI
    listaDati.classList.remove('hidden');
    let imageLinkFound = "";
    
    const maxLen = Math.max(headers.length, recordTrovato.length);

    for (let k = 0; k < maxLen; k++) {
        let key = String(headers[k] || "").trim();
        let val = recordTrovato[k];

        if (!val || String(val).trim() === "") continue;

        let keyLower = key.toLowerCase();
        let valString = String(val);

        // --- STEP A: CATTURA IMMAGINE ---
        let isImageHeader = keyLower.includes('foto') || keyLower.includes('immagine');
        let isDriveLink = valString.includes('drive.google.com') || valString.includes('docs.google.com');

        if (isImageHeader || isDriveLink) {
            imageLinkFound = valString;
            continue; 
        }

        // --- STEP B: FILTRO ASTERISCO RIGIDO ---
        // Se non inizia con *, SALTA e vai al prossimo giro.
        if (!key.startsWith('*')) {
            continue; 
        }

        // --- STEP C: PULIZIA TESTO ---
        let displayKey = key.replace('*', '').trim();
        let valClean = valString.toUpperCase().replace(/\s+/g, '');
        if (valClean === searchCode) continue;

        const div = document.createElement('div');
        div.className = 'data-item';
        // Il trattino estetico viene aggiunto qui
        div.innerText = `- ${displayKey}: ${val}`;
        listaDati.appendChild(div);
    }

    // 3. CARICA IMMAGINE (DOPPIO METODO)
    if (imageLinkFound) {
        let imgUrl = imageLinkFound;
        let driveId = null;

        if (imgUrl.includes("/d/")) {
            let idMatch = imgUrl.match(/\/d\/(.*?)\//);
            if (idMatch) driveId = idMatch[1];
        } else if (imgUrl.includes("id=")) {
            driveId = imgUrl.split('id=')[1].split('&')[0];
        }

        if (driveId) {
            // Metodo 1: Google User Content (con no-referrer)
            imgUrl = `https://drive.google.com/uc?export=view&id=${driveId}`;
            
            // Metodo 2 (Fallback): LH3 (Server immagini Google)
            // Se il metodo 1 fallisce, questo scatta automaticamente
            imgResult.onerror = () => {
                console.warn("Metodo 1 fallito. Provo Metodo 2...");
                imgResult.onerror = () => { // Se fallisce anche il 2, mostra errore
                     imgResult.classList.add('hidden');
                     errImmagine.classList.remove('hidden');
                };
                imgResult.src = `https://lh3.googleusercontent.com/d/${driveId}`;
            };
        } else {
             // Fallback per link non-Drive (se mai ne avessi)
             imgResult.onerror = () => {
                imgResult.classList.add('hidden');
                errImmagine.classList.remove('hidden');
             };
        }

        console.log("Carico immagine:", imgUrl);

        // Nasconde il referrer per bypassare i blocchi base
        imgResult.referrerPolicy = "no-referrer";
        imgResult.src = imgUrl;
        imgResult.classList.remove('hidden');
        
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