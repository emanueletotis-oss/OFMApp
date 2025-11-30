// --- CONFIGURAZIONE ---

// Ho estratto e ricostruito il tuo link originale pulito. NON TOCCARLO.
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

// Funzione che converte il link in un comando API ufficiale Microsoft
function getApiUrl(shareUrl) {
    let base64 = btoa(shareUrl);
    base64 = base64.replace(/\//g, '_').replace(/\+/g, '-').replace(/=+$/, '');
    return "https://api.onedrive.com/v1.0/shares/u!" + base64 + "/root/content";
}

async function eseguiRicerca() {
    const rawCode = inputCodice.value.trim();
    if (!rawCode) return; 

    const searchCode = rawCode.replace(/\s+/g, '').toUpperCase(); 
    loadingOverlay.classList.remove('hidden');

    try {
        // 1. Genera il link API (che punta al file binario, non alla pagina web)
        const apiUrl = getApiUrl(FILE_LINK);
        const timestamp = new Date().getTime();

        // 2. Usa CodeTabs (Il proxy più affidabile per i redirect)
        // Aggiungiamo il timestamp per evitare che l'iPhone usi la cache vecchia
        const proxyUrl = `https://api.codetabs.com/v1/proxy?quest=${encodeURIComponent(apiUrl)}&t=${timestamp}`;
        
        console.log("Scaricando da:", proxyUrl);

        const response = await fetch(proxyUrl, { method: 'GET', cache: 'no-store' });
        
        if (!response.ok) throw new Error("Errore connessione: " + response.status);
        
        const arrayBuffer = await response.arrayBuffer();

        // 3. CONTROLLO FONDAMENTALE
        // Se il file inizia con '<' (byte 60), è HTML (pagina di errore Microsoft), non Excel.
        const firstByte = new Uint8Array(arrayBuffer)[0];
        if (firstByte === 60) {
            // Convertiamo il buffer in testo per vedere l'errore (solo per debug)
            const textDecoder = new TextDecoder();
            const text = textDecoder.decode(arrayBuffer).substring(0, 100);
            console.error("Ricevuto HTML:", text);
            throw new Error("Microsoft ha bloccato il download diretto. Riprova tra 1 minuto.");
        }

        // 4. Lettura Excel
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        
        // Estrai i dati come matrice (Raw)
        const matrix = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

        if (!matrix || matrix.length === 0) throw new Error("File Excel vuoto.");

        // Ritardo estetico
        setTimeout(() => {
            elaboraDati(matrix, searchCode);
            loadingOverlay.classList.add('hidden');
        }, 3000); 

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

    // Cerca il codice
    for (let i = 0; i < matrix.length; i++) {
        const row = matrix[i];
        for (let j = 0; j < row.length; j++) {
            let cell = String(row[j]).toUpperCase().replace(/\s+/g, '');
            if (cell === searchCode) {
                recordTrovato = row;
                // Intestazioni: se siamo alla riga i, l'header è alla riga 0
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

    listaDati.classList.remove('hidden');
    let imageLinkFound = "";
    const maxLen = Math.max(headers.length, recordTrovato.length);

    for (let k = 0; k < maxLen; k++) {
        let key = headers[k] || `Colonna ${k+1}`;
        let val = recordTrovato[k];

        if (!val || String(val).trim() === "") continue;

        let keyLower = String(key).toLowerCase();
        
        // Rileva link immagine
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

    // Mostra immagine
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