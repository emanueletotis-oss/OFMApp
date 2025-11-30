// --- CONFIGURAZIONE ---

// 1. IL TUO LINK ORIGINALE (Preso dal tuo messaggio)
// Non lo modifichiamo, lo usiamo così com'è.
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

// Funzione di supporto per tentare il download
async function tryFetch(proxyUrl) {
    try {
        console.log("Tentativo download da:", proxyUrl);
        const response = await fetch(proxyUrl, {
            method: 'GET',
            cache: 'no-store' // IMPORTANTE: Forza l'iPhone a scaricare sempre il file nuovo
        });
        
        if (!response.ok) throw new Error("Status: " + response.status);
        
        const arrayBuffer = await response.arrayBuffer();
        
        // Controllo se è HTML (Errore) invece di Excel
        const firstByte = new Uint8Array(arrayBuffer)[0];
        if (firstByte === 60) throw new Error("Ricevuto HTML invece di Excel"); // 60 = '<'
        
        return arrayBuffer;
    } catch (e) {
        console.warn("Tentativo fallito:", e);
        return null; // Ritorna null se fallisce
    }
}

async function eseguiRicerca() {
    const rawCode = inputCodice.value.trim();
    if (!rawCode) return; 

    const searchCode = rawCode.toUpperCase(); 
    loadingOverlay.classList.remove('hidden');

    try {
        // PREPARAZIONE LINK CORRETTA
        // Invece di sostituire /x/, aggiungiamo semplicemente ?download=1 alla fine.
        // Questo dice a OneDrive di scaricare il file invece di aprirlo.
        let downloadLink = USER_LINK;
        
        // Se c'è già un '?', usiamo '&', altrimenti '?'
        if (downloadLink.includes('?')) {
            downloadLink += "&download=1";
        } else {
            downloadLink += "?download=1";
        }

        // STRATEGIA PROXY (AllOrigins è il più stabile per il mobile)
        let excelData = null;

        // Costruiamo l'URL del proxy
        // Aggiungiamo un timestamp casuale per evitare che l'iPhone usi la cache vecchia
        const proxyUrl = "https://api.allorigins.win/raw?url=" + encodeURIComponent(downloadLink) + "&rand=" + new Date().getTime();
        
        excelData = await tryFetch(proxyUrl);

        // Se fallisce, proviamo il secondo proxy (backup)
        if (!excelData) {
             const proxy2 = "https://corsproxy.io/?" + encodeURIComponent(downloadLink) + "&rand=" + new Date().getTime();
             excelData = await tryFetch(proxy2);
        }

        if (!excelData) {
            throw new Error("Impossibile scaricare il file. Controlla la connessione internet.");
        }

        // ELABORAZIONE FILE
        const workbook = XLSX.read(excelData, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

        if (jsonData.length === 0) throw new Error("Il file Excel è vuoto.");

        // Ritardo estetico richiesto
        setTimeout(() => {
            elaboraDati(jsonData, searchCode);
            loadingOverlay.classList.add('hidden');
        }, 3500); 

    } catch (error) {
        console.error("ERRORE:", error);
        loadingOverlay.classList.add('hidden');
        alert("Errore tecnico: " + error.message + "\n\nSe il problema persiste, prova a chiudere e riaprire la pagina.");
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