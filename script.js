// --- CONFIGURAZIONE ---

// ID del file Google Drive (estratto dal tuo link)
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

// Funzione che tenta di scaricare usando diversi proxy in sequenza
async function scaricaConFallback(googleUrl) {
    const timestamp = new Date().getTime();
    
    // Lista di "Ponti" (Proxy) da provare in ordine
    const proxies = [
        // 1. CorsProxy.io (Veloce e stabile)
        `https://corsproxy.io/?${encodeURIComponent(googleUrl)}&t=${timestamp}`,
        // 2. CodeTabs (Ottimo per i redirect)
        `https://api.codetabs.com/v1/proxy?quest=${encodeURIComponent(googleUrl)}&t=${timestamp}`,
        // 3. AllOrigins (Backup finale)
        `https://api.allorigins.win/raw?url=${encodeURIComponent(googleUrl)}&t=${timestamp}`
    ];

    let lastError = null;

    // Prova i proxy uno alla volta
    for (let proxyUrl of proxies) {
        try {
            console.log("Tentativo download con:", proxyUrl);
            const response = await fetch(proxyUrl, { method: 'GET', cache: 'no-store' });
            
            if (!response.ok) throw new Error(`HTTP ${response.status}`);
            
            const arrayBuffer = await response.arrayBuffer();
            
            // Verifica che non sia una pagina di errore HTML
            const firstByte = new Uint8Array(arrayBuffer)[0];
            if (firstByte === 60) { // 60 = '<'
                throw new Error("Ricevuto HTML invece di Excel");
            }
            
            return arrayBuffer; // Successo! Usciamo dal ciclo e restituiamo il file
            
        } catch (e) {
            console.warn("Proxy fallito, provo il prossimo...", e);
            lastError = e;
            // Continua col prossimo proxy nel ciclo
        }
    }
    
    // Se siamo qui, tutti i proxy hanno fallito
    throw lastError || new Error("Tutti i tentativi di connessione sono falliti.");
}

async function eseguiRicerca() {
    const rawCode = inputCodice.value.trim();
    if (!rawCode) return; 

    const searchCode = rawCode.replace(/\s+/g, '').toUpperCase(); 
    loadingOverlay.classList.remove('hidden');

    try {
        // Costruiamo il link diretto per l'export in Excel da Google Drive
        const exportUrl = `https://docs.google.com/spreadsheets/d/${FILE_ID}/export?format=xlsx`;

        // Avviamo il download con la strategia "Carro Armato" (prova più strade)
        const arrayBuffer = await scaricaConFallback(exportUrl);

        if (!arrayBuffer) throw new Error("Impossibile scaricare il file.");

        // LETTURA EXCEL
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        
        // Converte in matrice dati
        const matrix = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

        if (!matrix || matrix.length === 0) throw new Error("Il file Excel è vuoto.");

        // Ritardo estetico
        setTimeout(() => {
            elaboraDati(matrix, searchCode);
            loadingOverlay.classList.add('hidden');
        }, 2000); 

    } catch (error) {
        console.error(error);
        loadingOverlay.classList.add('hidden');
        alert("Errore di Connessione: " + error.message + "\n\nProva a ricaricare la pagina o cambia rete (Wi-Fi/4G).");
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

    // Ricerca nelle righe
    for (let i = 0; i < matrix.length; i++) {
        const row = matrix[i];
        for (let j = 0; j < row.length; j++) {
            let cell = String(row[j]).toUpperCase().replace(/\s+/g, '');
            if (cell === searchCode) {
                recordTrovato = row;
                // Intestazioni
                if (i > 0) headers = matrix[0]; 
                else headers = row.map((_, idx) => `Dato ${idx + 1}`);
                break;
            }
        }
        if (recordTrovato) break;
    }

    // Se non trovato
    if (!recordTrovato) {
        listaDati.classList.add('hidden');
        errCodice.classList.remove('hidden');
        imgResult.classList.add('hidden');
        errImmagine.classList.remove('hidden');
        return;
    }

    // Se trovato
    listaDati.classList.remove('hidden');
    let imageLinkFound = "";
    const maxLen = Math.max(headers.length, recordTrovato.length);

    for (let k = 0; k < maxLen; k++) {
        let key = headers[k] || `Colonna ${k+1}`;
        let val = recordTrovato[k];

        if (!val || String(val).trim() === "") continue;

        let keyLower = String(key).toLowerCase();
        let valString = String(val);

        // Rileva Immagini/Link
        let isImageColumn = keyLower.includes('immagine') || keyLower.includes('foto') || keyLower.includes('url') || keyLower.includes('link');
        let isDriveLink = valString.includes('drive.google.com') || valString.includes('docs.google.com');

        if (isImageColumn || isDriveLink) {
            imageLinkFound = valString;
            continue; 
        }

        // Non ristampare il codice
        let valClean = valString.toUpperCase().replace(/\s+/g, '');
        if (valClean === searchCode) continue;

        const div = document.createElement('div');
        div.className = 'data-item';
        div.innerText = `- ${key}: ${val}`;
        listaDati.appendChild(div);
    }

    // Mostra Immagine
    if (imageLinkFound) {
        let imgUrl = imageLinkFound;
        // Fix link Google Drive
        if (imgUrl.includes("/d/")) {
            let idMatch = imgUrl.match(/\/d\/(.*?)\//);
            if (idMatch) imgUrl = `https://drive.google.com/uc?export=view&id=${idMatch[1]}`;
        }
        imgResult.src = imgUrl;
        imgResult.classList.remove('hidden');
        imgResult.onerror = () => {
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