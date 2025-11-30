// --- CONFIGURAZIONE ---

// LINK ONEDRIVE (Non toccare)
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

// Funzione helper per il download
async function tryFetch(url, proxyName) {
    try {
        console.log(`Tentativo con ${proxyName}:`, url);
        const response = await fetch(url, { method: 'GET', cache: 'no-store' });
        
        if (!response.ok) throw new Error(`HTTP ${response.status}`);
        
        const arrayBuffer = await response.arrayBuffer();
        
        const firstByte = new Uint8Array(arrayBuffer)[0];
        if (firstByte === 60) throw new Error("Ricevuto HTML invece di Excel"); 
        
        return arrayBuffer;
    } catch (e) {
        console.warn(`${proxyName} fallito:`, e.message);
        return null;
    }
}

async function eseguiRicerca() {
    // PULIZIA INPUT UTENTE: Rimuove spazi prima e dopo
    const rawCode = inputCodice.value.trim();
    if (!rawCode) return; 

    // Normalizza tutto in maiuscolo per il confronto
    const searchCode = rawCode.toUpperCase(); 
    loadingOverlay.classList.remove('hidden');

    try {
        let directLink = USER_LINK.replace('/x/', '/download/').split('?')[0];
        let excelData = null;
        const timestamp = new Date().getTime();

        // 1. CodeTabs
        if (!excelData) {
            const url1 = `https://api.codetabs.com/v1/proxy?quest=${encodeURIComponent(directLink)}&t=${timestamp}`;
            excelData = await tryFetch(url1, "CodeTabs");
        }

        // 2. CorsProxy
        if (!excelData) {
            const url2 = `https://corsproxy.io/?${encodeURIComponent(directLink)}&t=${timestamp}`;
            excelData = await tryFetch(url2, "CorsProxy");
        }

        // 3. AllOrigins
        if (!excelData) {
            const url3 = `https://api.allorigins.win/raw?url=${encodeURIComponent(directLink)}&t=${timestamp}`;
            excelData = await tryFetch(url3, "AllOrigins");
        }

        if (!excelData) throw new Error("Connessione instabile. Riprova.");

        // LETTURA EXCEL
        const workbook = XLSX.read(excelData, { type: 'array' });
        
        // Cerca il primo foglio che ha dati
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        
        // Converte in JSON grezzo
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

        // Ritardo animazione
        setTimeout(() => {
            elaboraDati(jsonData, searchCode);
            loadingOverlay.classList.add('hidden');
        }, 3500); 

    } catch (error) {
        console.error("ERRORE:", error);
        loadingOverlay.classList.add('hidden');
        alert("Errore: " + error.message);
    }
}

function elaboraDati(data, searchCode) {
    errCodice.classList.add('hidden');
    errImmagine.classList.add('hidden');
    phRisultati.classList.add('hidden');
    phImmagine.classList.add('hidden');
    listaDati.innerHTML = "";

    let recordTrovato = null;

    // --- RICERCA INTELLIGENTE ---
    // Scorre tutte le righe del file Excel
    for (let row of data) {
        // Controlla ogni valore dentro la riga (non solo la colonna 'Codice')
        // Questo risolve il problema se la colonna ha un nome diverso
        const values = Object.values(row);
        
        for (let val of values) {
            // Pulisce il valore Excel da spazi e lo mette maiuscolo
            let valClean = String(val).trim().toUpperCase();
            
            // Confronto esatto
            if (valClean === searchCode) {
                recordTrovato = row;
                break; // Trovato! Esce dal ciclo interno
            }
        }
        if (recordTrovato) break; // Trovato! Esce dal ciclo delle righe
    }

    // --- SE NON TROVATO ---
    if (!recordTrovato) {
        listaDati.classList.add('hidden');
        errCodice.classList.remove('hidden');
        imgResult.classList.add('hidden');
        errImmagine.classList.remove('hidden');
        return;
    }

    // --- SE TROVATO ---
    listaDati.classList.remove('hidden');

    const keys = Object.keys(recordTrovato);
    let imageLinkFound = "";

    keys.forEach(key => {
        let val = recordTrovato[key];
        // Pulisce anche il valore da visualizzare se è una stringa
        if (typeof val === 'string') val = val.trim();

        const keyLower = key.toLowerCase();

        // Cerca link immagine
        if (keyLower.includes('immagine') || keyLower.includes('link') || keyLower.includes('foto') || keyLower.includes('url')) {
            imageLinkFound = val; 
            return;
        }

        // Escludiamo la colonna che contiene il codice stesso per non ripeterlo
        // e i valori vuoti
        let isCodeColumn = String(val).toUpperCase() === searchCode;
        
        if (!isCodeColumn && val !== "" && val !== undefined) {
            const div = document.createElement('div');
            div.className = 'data-item';
            div.innerText = `- ${key}: ${val}`;
            listaDati.appendChild(div);
        }
    });

    // Gestione Immagine
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

// Event Listeners
btnPlay.addEventListener('click', eseguiRicerca);
btnReset.addEventListener('click', resetApp);
inputCodice.addEventListener('keypress', (e) => {
    if (e.key === 'Enter') eseguiRicerca();
});