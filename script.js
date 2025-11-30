// --- CONFIGURAZIONE ---

// 1. INCOLLA QUI IL LINK LUNGHISSIMO (Quello preso dalla barra degli indirizzi dopo aver aperto il file)
// Deve iniziare con https://onedrive.live.com/....
const LONG_LINK = "https://onedrive.live.com/edit?cid=ac1a912c65f087d9&id=AC1A912C65F087D9!sc4bdac82f35b400d98bf14bbd8f0a349&resid=AC1A912C65F087D9!sc4bdac82f35b400d98bf14bbd8f0a349&ithint=file%2Cxlsx&embed=1&migratedtospo=true&redeem=aHR0cHM6Ly8xZHJ2Lm1zL3gvYy9hYzFhOTEyYzY1ZjA4N2Q5L0lRU0NyTDNFV19NTlFKaV9GTHZZOEtOSkFYWFMtN0tzSHVvcm5XQXFZZ0FvTm5F&wdo=2";

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

// Funzione che scarica e controlla se è un vero Excel
async function tryFetch(url, proxyName) {
    try {
        const response = await fetch(url, { method: 'GET', cache: 'no-store' });
        if (!response.ok) throw new Error(`HTTP ${response.status}`);
        
        const arrayBuffer = await response.arrayBuffer();
        
        // Controllo primo byte: 60 = '<' (HTML, errore), 80 = 'P' (Excel/Zip, corretto)
        const firstByte = new Uint8Array(arrayBuffer)[0];
        if (firstByte === 60) return null; // È HTML, quindi fallito
        
        return arrayBuffer;
    } catch (e) {
        console.warn(`${proxyName} errore:`, e.message);
        return null;
    }
}

async function eseguiRicerca() {
    const rawCode = inputCodice.value.trim();
    if (!rawCode) return; 

    // Controllo se l'utente ha incollato il link
    if (LONG_LINK.includes("INCOLLA_QUI")) {
        alert("Errore: Devi aprire il file script.js e incollare il link lungo di OneDrive!");
        return;
    }

    const searchCode = rawCode.replace(/\s+/g, '').toUpperCase(); 
    loadingOverlay.classList.remove('hidden');

    try {
        // --- TRASFORMAZIONE LINK ---
        // Prende il link lungo (edit o view) e lo forza a 'download'
        // Esempio: onedrive.live.com/edit.aspx?... -> onedrive.live.com/download?...
        let downloadUrl = LONG_LINK.replace('/edit.aspx', '/download').replace('/view.aspx', '/download');
        
        // Se non ha sostituito nulla (magari il link è diverso), proviamo ad aggiungere ?download=1
        if (!downloadUrl.includes('/download')) {
             if (downloadUrl.includes('?')) downloadUrl += "&download=1";
             else downloadUrl += "?download=1";
        }

        const timestamp = new Date().getTime();
        let excelData = null;

        // --- STRATEGIA PROXY (AllOrigins è il più robusto per OneDrive) ---
        // Usiamo solo questo perché è quello che gestisce meglio i redirect di onedrive.live.com
        const proxyUrl = `https://api.allorigins.win/raw?url=${encodeURIComponent(downloadUrl)}&t=${timestamp}`;
        
        excelData = await tryFetch(proxyUrl, "AllOrigins");

        // Se fallisce, proviamo CorsProxy come backup
        if (!excelData) {
            const proxy2 = `https://corsproxy.io/?${encodeURIComponent(downloadUrl)}&t=${timestamp}`;
            excelData = await tryFetch(proxy2, "CorsProxy");
        }

        if (!excelData) throw new Error("Link non valido o blocco connessione. Assicurati di usare il Link Lungo dalla barra degli indirizzi.");

        // LETTURA EXCEL
        const workbook = XLSX.read(excelData, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        
        // Matrice dati
        const matrix = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

        if (!matrix || matrix.length === 0) throw new Error("Il file Excel sembra vuoto.");

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

    // Cerca il codice in tutta la matrice
    for (let i = 0; i < matrix.length; i++) {
        const row = matrix[i];
        for (let j = 0; j < row.length; j++) {
            let cell = String(row[j]).toUpperCase().replace(/\s+/g, '');
            if (cell === searchCode) {
                recordTrovato = row;
                // Intestazioni: prova la riga 0, o header generici
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

    // Mostra dati
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

    // Visualizza Immagine
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