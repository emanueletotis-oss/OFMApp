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

// Funzione "Carro Armato" per scaricare il file (3 livelli di sicurezza)
async function scaricaConFallback(googleUrl) {
    const timestamp = new Date().getTime();
    
    // Proxy multipli per garantire il download anche con 4G
    const proxies = [
        `https://corsproxy.io/?${encodeURIComponent(googleUrl)}&t=${timestamp}`,
        `https://api.codetabs.com/v1/proxy?quest=${encodeURIComponent(googleUrl)}&t=${timestamp}`,
        `https://api.allorigins.win/raw?url=${encodeURIComponent(googleUrl)}&t=${timestamp}`
    ];

    let lastError = null;

    for (let proxyUrl of proxies) {
        try {
            console.log("Download tentativo:", proxyUrl);
            const response = await fetch(proxyUrl, { method: 'GET', cache: 'no-store' });
            
            if (!response.ok) throw new Error(`HTTP ${response.status}`);
            
            const arrayBuffer = await response.arrayBuffer();
            
            // Controllo anti-pagina-web
            const firstByte = new Uint8Array(arrayBuffer)[0];
            if (firstByte === 60) throw new Error("Ricevuto HTML invece di Excel");
            
            return arrayBuffer;
        } catch (e) {
            console.warn("Proxy fallito:", e);
            lastError = e;
        }
    }
    throw lastError || new Error("Impossibile scaricare il file Excel. Controlla la connessione.");
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

    // 1. CERCA IL CODICE NELLE RIGHE
    for (let i = 0; i < matrix.length; i++) {
        const row = matrix[i];
        for (let j = 0; j < row.length; j++) {
            let cell = String(row[j]).toUpperCase().replace(/\s+/g, '');
            if (cell === searchCode) {
                recordTrovato = row;
                // Intestazioni: usa la prima riga del file (indice 0)
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

    // 2. MOSTRA I DATI
    listaDati.classList.remove('hidden');
    let imageLinkFound = "";
    
    // Controlliamo se l'utente sta usando il filtro asterisco (*)
    // Se almeno una colonna ha *, attiveremo la modalità "Solo Asterischi"
    const usaFiltroAsterisco = headers.some(h => String(h).trim().startsWith('*'));

    const maxLen = Math.max(headers.length, recordTrovato.length);

    for (let k = 0; k < maxLen; k++) {
        let key = String(headers[k] || `Colonna ${k+1}`).trim();
        let val = recordTrovato[k];

        if (!val || String(val).trim() === "") continue;

        let keyLower = key.toLowerCase();
        let valString = String(val);

        // --- RILEVAMENTO IMMAGINE ---
        // Se la colonna si chiama "FOTO" o contiene un link, è un'immagine
        let isImageHeader = keyLower.includes('foto') || keyLower.includes('immagine') || keyLower.includes('link');
        let isUrlContent = valString.includes('drive.google.com') || valString.includes('docs.google.com');

        if (isImageHeader || isUrlContent) {
            imageLinkFound = valString;
            continue; // Non scriverla come testo
        }

        // --- FILTRO ASTERISCO ---
        // Se nel file Excel ci sono colonne con *, mostriamo SOLO quelle.
        if (usaFiltroAsterisco && !key.startsWith('*')) {
            continue; // Nascondi colonne senza * (es. "fuffa")
        }

        // Pulisce il nome per la visualizzazione (toglie l'asterisco)
        let displayKey = key.replace('*', '').trim();

        // Non mostrare il codice cercato nell'elenco
        let valClean = valString.toUpperCase().replace(/\s+/g, '');
        if (valClean === searchCode) continue;

        const div = document.createElement('div');
        div.className = 'data-item';
        div.innerText = `- ${displayKey}: ${val}`;
        listaDati.appendChild(div);
    }

    // 3. CARICA IMMAGINE NEL RIQUADRO
    if (imageLinkFound) {
        let imgUrl = imageLinkFound;
        
        // Fix automatico link Google Drive (trasforma 'view' in link diretto)
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
        errImmagine.classList.remove('hidden');
    }
}

// Event Listeners
btnPlay.addEventListener('click', eseguiRicerca);
btnReset.addEventListener('click', resetApp);
inputCodice.addEventListener('keypress', (e) => {
    if (e.key === 'Enter') eseguiRicerca();
});