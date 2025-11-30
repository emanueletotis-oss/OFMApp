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

// Funzione Download Robusto
async function scaricaConFallback(googleUrl) {
    const timestamp = new Date().getTime();
    
    // Proxy multipli
    const proxies = [
        `https://corsproxy.io/?${encodeURIComponent(googleUrl)}&t=${timestamp}`,
        `https://api.codetabs.com/v1/proxy?quest=${encodeURIComponent(googleUrl)}&t=${timestamp}`,
        `https://api.allorigins.win/raw?url=${encodeURIComponent(googleUrl)}&t=${timestamp}`
    ];

    let lastError = null;

    for (let proxyUrl of proxies) {
        try {
            console.log("Download...", proxyUrl);
            const response = await fetch(proxyUrl, { method: 'GET', cache: 'no-store' });
            
            if (!response.ok) throw new Error(`HTTP ${response.status}`);
            
            const arrayBuffer = await response.arrayBuffer();
            
            const firstByte = new Uint8Array(arrayBuffer)[0];
            if (firstByte === 60) throw new Error("Ricevuto HTML invece di Excel");
            
            return arrayBuffer;
        } catch (e) {
            console.warn("Proxy errore:", e);
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
        
        // Converte in matrice (array di array)
        // defval: "" ci assicura che le celle vuote siano stringhe vuote
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

    // 1. CERCA LA RIGA GIUSTA
    for (let i = 0; i < matrix.length; i++) {
        const row = matrix[i];
        
        // Cerca in tutte le celle della riga
        for (let j = 0; j < row.length; j++) {
            let cell = String(row[j]).toUpperCase().replace(/\s+/g, '');
            if (cell === searchCode) {
                recordTrovato = row;
                // Le intestazioni sono sempre nella prima riga (indice 0)
                headers = matrix[0]; 
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

    // 2. VISUALIZZA I DATI
    listaDati.classList.remove('hidden');
    let imageLinkFound = "";
    
    // Scansioniamo tutte le colonne (fino alla lunghezza dell'intestazione o della riga)
    const maxLen = Math.max(headers.length, recordTrovato.length);

    for (let k = 0; k < maxLen; k++) {
        // Recupera Titolo Colonna e Valore Cella
        let key = String(headers[k] || `Colonna ${k+1}`).trim();
        let val = recordTrovato[k];

        // Se il valore è vuoto o null, salta la colonna (così nasconde le celle vuote)
        if (val === undefined || val === null || String(val).trim() === "") continue;

        let keyLower = key.toLowerCase();
        let valString = String(val);
        let valClean = valString.toUpperCase().replace(/\s+/g, '');

        // A. È IL CODICE CERCATO? -> Nascondilo
        if (valClean === searchCode) continue;

        // B. È LA FOTO? -> Salvala per dopo, non scriverla
        // Cerca se il titolo è esattamente "FOTO" o contiene "IMMAGINE"
        if (keyLower === 'foto' || keyLower.includes('immagine') || keyLower.includes('link')) {
            imageLinkFound = valString;
            continue; 
        }

        // C. È UN DATO NORMALE? -> Scrivilo
        const div = document.createElement('div');
        div.className = 'data-item';
        // Mostra Titolo: Valore
        div.innerText = `- ${key}: ${val}`;
        listaDati.appendChild(div);
    }

    // 3. VISUALIZZA L'IMMAGINE
    if (imageLinkFound) {
        let imgUrl = imageLinkFound;
        
        // Correzione Link Google Drive
        // Trasforma: https://drive.google.com/file/d/XXX/view?usp=drive_link
        // In: https://drive.google.com/uc?export=view&id=XXX
        if (imgUrl.includes("/d/")) {
            let idMatch = imgUrl.match(/\/d\/(.*?)\//);
            if (idMatch) {
                let imgId = idMatch[1];
                imgUrl = `https://drive.google.com/uc?export=view&id=${imgId}`;
            }
        }

        imgResult.src = imgUrl;
        imgResult.classList.remove('hidden');
        
        imgResult.onerror = () => {
            imgResult.classList.add('hidden');
            errImmagine.classList.remove('hidden');
        };
    } else {
        // Nessun link foto trovato -> Messaggio errore dorato
        errImmagine.classList.remove('hidden');
    }
}

// Event Listeners
btnPlay.addEventListener('click', eseguiRicerca);
btnReset.addEventListener('click', resetApp);
inputCodice.addEventListener('keypress', (e) => {
    if (e.key === 'Enter') eseguiRicerca();
});