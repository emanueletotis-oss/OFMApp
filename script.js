// --- CONFIGURAZIONE ---

// Il tuo link di condivisione (Lo lasciamo così com'è)
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

// Funzione magica: Converte il link di condivisione in un link API diretto
function generateOneDriveApiUrl(shareUrl) {
    // 1. Codifica il link in Base64
    let base64Value = btoa(shareUrl);
    // 2. Rendi la stringa sicura per l'URL (sostituisci caratteri speciali)
    base64Value = base64Value.replace(/\//g, '_').replace(/\+/g, '-');
    // 3. Rimuovi i simboli '=' alla fine
    base64Value = base64Value.replace(/=+$/, '');
    // 4. Costruisci l'URL dell'API di OneDrive
    return "https://api.onedrive.com/v1.0/shares/u!" + base64Value + "/root/content";
}

async function tryFetch(proxyUrl) {
    try {
        console.log("Tentativo con:", proxyUrl);
        const response = await fetch(proxyUrl, { method: 'GET', cache: 'no-store' });
        if (!response.ok) throw new Error("HTTP " + response.status);
        
        const arrayBuffer = await response.arrayBuffer();
        
        // Controllo se è ancora HTML (vuol dire che il proxy ha fallito il redirect)
        const firstByte = new Uint8Array(arrayBuffer)[0];
        if (firstByte === 60) throw new Error("Ricevuto HTML invece di Excel"); 
        
        return arrayBuffer;
    } catch (e) {
        console.warn("Proxy fallito:", e);
        return null;
    }
}

async function eseguiRicerca() {
    const rawCode = inputCodice.value.trim();
    if (!rawCode) return; 

    const searchCode = rawCode.toUpperCase(); 
    loadingOverlay.classList.remove('hidden');

    try {
        // 1. Genera il link API ufficiale
        const apiUrl = generateOneDriveApiUrl(USER_LINK);
        
        // 2. Usa un proxy per chiamare l'API (necessario per scaricare file su web)
        // Usiamo allorigins che segue bene i reindirizzamenti dell'API
        let excelData = null;
        
        const proxyUrl = "https://api.allorigins.win/raw?url=" + encodeURIComponent(apiUrl) + "&rand=" + new Date().getTime();
        
        excelData = await tryFetch(proxyUrl);

        if (!excelData) {
            // Backup: prova con corsproxy se il primo fallisce
            const proxy2 = "https://corsproxy.io/?" + encodeURIComponent(apiUrl) + "&rand=" + new Date().getTime();
            excelData = await tryFetch(proxy2);
        }

        if (!excelData) throw new Error("Impossibile contattare OneDrive. Verifica la connessione.");

        // 3. Elaborazione File
        const workbook = XLSX.read(excelData, { type: 'array' });
        
        if (!workbook.SheetNames.length) throw new Error("File Excel non valido.");
        
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

        // Ritardo estetico
        setTimeout(() => {
            elaboraDati(jsonData, searchCode);
            loadingOverlay.classList.add('hidden');
        }, 3500); 

    } catch (error) {
        console.error("ERRORE:", error);
        loadingOverlay.classList.add('hidden');
        alert("Errore di caricamento: " + error.message + "\n\nAssicurati che il file Excel non sia vuoto o protetto da password.");
    }
}

function elaboraDati(data, code) {
    errCodice.classList.add('hidden');
    errImmagine.classList.add('hidden');
    phRisultati.classList.add('hidden');
    phImmagine.classList.add('hidden');
    listaDati.innerHTML = "";

    let recordTrovato = null;

    for (let row of data) {
        let rowCode = "";
        if (row['Codice']) rowCode = row['Codice'];
        else if (row['codice']) rowCode = row['codice'];
        else rowCode = Object.values(row)[0]; 
        
        if (String(rowCode).toUpperCase() === code) {
            recordTrovato = row;
            break;
        }
    }

    if (!recordTrovato) {
        listaDati.classList.add('hidden');
        errCodice.classList.remove('hidden');
        imgResult.classList.add('hidden');
        errImmagine.classList.remove('hidden');
        return;
    }

    listaDati.classList.remove('hidden');

    const keys = Object.keys(recordTrovato);
    let imageLinkFound = "";

    keys.forEach(key => {
        const val = recordTrovato[key];
        const keyLower = key.toLowerCase();

        if (keyLower.includes('immagine') || keyLower.includes('link') || keyLower.includes('foto') || keyLower.includes('url')) {
            imageLinkFound = val; 
            return;
        }

        if (keyLower !== 'codice' && val !== "" && val !== undefined) {
            const div = document.createElement('div');
            div.className = 'data-item';
            div.innerText = `- ${key}: ${val}`;
            listaDati.appendChild(div);
        }
    });

    if (imageLinkFound && String(imageLinkFound).trim() !== "") {
        let imgUrl = String(imageLinkFound);
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

btnPlay.addEventListener('click', eseguiRicerca);
btnReset.addEventListener('click', resetApp);
inputCodice.addEventListener('keypress', (e) => {
    if (e.key === 'Enter') eseguiRicerca();
});