// --- CONFIGURAZIONE ---

// IL TUO LINK ORIGINALE (Lo lasciamo così com'è, lo script lo trasformerà)
const USER_LINK = "https://1drv.ms/x/c/ac1a912c65f087d9/IQSCrL3EW_MNQJi_FLvY8KNJAXXS-7KsHuornWAqYgAoNnE";

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

// QUESTA È LA FUNZIONE MAGICA
// Trasforma il link "1drv.ms" in un link API ufficiale che forza il download
function creaLinkUfficialeOneDrive(url) {
    // 1. Codifica il link in Base64
    let encoded = btoa(url);
    // 2. Rendi la stringa sicura per URL (sostituisci caratteri speciali)
    encoded = encoded.replace(/\//g, '_').replace(/\+/g, '-');
    // 3. Rimuovi i simboli '=' finali
    encoded = encoded.replace(/=+$/, '');
    // 4. Crea l'URL che punta direttamente al file
    return "https://api.onedrive.com/v1.0/shares/u!" + encoded + "/root/content";
}

async function tryFetch(url, proxyName) {
    try {
        const response = await fetch(url, { method: 'GET', cache: 'no-store' });
        if (!response.ok) throw new Error(`HTTP ${response.status}`);
        
        const arrayBuffer = await response.arrayBuffer();
        
        // Verifica se è un file vero (se inizia con '<' è HTML, quindi errore)
        const firstByte = new Uint8Array(arrayBuffer)[0];
        if (firstByte === 60) return null; 
        
        return arrayBuffer;
    } catch (e) {
        console.warn(`${proxyName} fallito:`, e.message);
        return null;
    }
}

async function eseguiRicerca() {
    const rawCode = inputCodice.value.trim();
    if (!rawCode) return; 

    // Pulisce il codice (toglie spazi e mette maiuscolo)
    const searchCode = rawCode.replace(/\s+/g, '').toUpperCase(); 
    loadingOverlay.classList.remove('hidden');

    try {
        // 1. Generiamo il link API Microsoft
        const apiLink = creaLinkUfficialeOneDrive(USER_LINK);
        const timestamp = new Date().getTime();
        let excelData = null;

        // 2. Tentativi di download tramite Proxy (necessario per aggirare i blocchi)
        
        // Tentativo A: CorsProxy (Veloce e gestisce bene i redirect API)
        if (!excelData) {
            const urlA = `https://corsproxy.io/?${encodeURIComponent(apiLink)}&t=${timestamp}`;
            excelData = await tryFetch(urlA, "CorsProxy");
        }

        // Tentativo B: CodeTabs (Molto affidabile)
        if (!excelData) {
            const urlB = `https://api.codetabs.com/v1/proxy?quest=${encodeURIComponent(apiLink)}&t=${timestamp}`;
            excelData = await tryFetch(urlB, "CodeTabs");
        }

        if (!excelData) throw new Error("Impossibile scaricare il file Excel. Riprova tra poco.");

        // 3. Lettura del file
        const workbook = XLSX.read(excelData, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        
        // Converte in matrice di dati
        const matrix = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

        if (!matrix || matrix.length === 0) throw new Error("Il file Excel scaricato è vuoto.");

        // Ritardo estetico
        setTimeout(() => {
            elaboraDati(matrix, searchCode);
            loadingOverlay.classList.add('hidden');
        }, 3000); 

    } catch (error) {
        loadingOverlay.classList.add('hidden');
        alert("Errore tecnico: " + error.message);
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

    // Cerchiamo il codice in ogni riga
    for (let i = 0; i < matrix.length; i++) {
        const row = matrix[i];
        
        for (let j = 0; j < row.length; j++) {
            // Pulisce la cella da spazi e formattazione
            let cell = String(row[j]).toUpperCase().replace(/\s+/g, '');
            
            if (cell === searchCode) {
                recordTrovato = row;
                // Cerca di capire dove sono le intestazioni (di solito riga 0)
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

    // Trovato! Mostriamo i dati
    listaDati.classList.remove('hidden');
    let imageLinkFound = "";

    const maxLen = Math.max(headers.length, recordTrovato.length);
    for (let k = 0; k < maxLen; k++) {
        let key = headers[k] || `Colonna ${k+1}`;
        let val = recordTrovato[k];

        if (!val || String(val).trim() === "") continue;

        let keyLower = String(key).toLowerCase();
        
        // Cerca colonna immagine
        if (keyLower.includes('immagine') || keyLower.includes('link') || keyLower.includes('foto') || keyLower.includes('url')) {
            imageLinkFound = val;
            continue;
        }

        // Non mostrare il codice stesso nell'elenco
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
        // Correzione automatica link drive
        if (imgUrl.includes("drive.google.com")) {
             imgUrl = imgUrl.replace("/view", "/preview").replace("open?id=", "uc?id=");
        }
        imgResult.src = imgUrl;
        imgResult.classList.remove('hidden');
        
        imgResult.onerror = () => {
            imgResult.classList.add('hidden');
            errImmagine.classList.remove('hidden');
        };
    } else {
        imgResult.classList.add('hidden');
        errImmagine.classList.remove('hidden');
    }
}

// Event Listeners
btnPlay.addEventListener('click', eseguiRicerca);
btnReset.addEventListener('click', resetApp);
inputCodice.addEventListener('keypress', (e) => {
    if (e.key === 'Enter') eseguiRicerca();
});