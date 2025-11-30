// --- CONFIGURAZIONE ---

// Usiamo un proxy CORS (corsproxy.io) per aggirare il blocco di sicurezza del browser.
// Il link originale è quello di OneDrive con ?download=1 alla fine.
const ONEDRIVE_LINK = "https://1drv.ms/x/c/ac1a912c65f087d9/IQSCrL3EW_MNQJi_FLvY8KNJAXXS-7KsHuornWAqYgAoNnE?download=1";
const EXCEL_URL = "https://corsproxy.io/?" + encodeURIComponent(ONEDRIVE_LINK);

// --- ELEMENTI DOM ---
const inputCodice = document.getElementById('input-codice');
const btnPlay = document.getElementById('btn-play');
const btnReset = document.getElementById('btn-reset');
const loadingOverlay = document.getElementById('loading-overlay');

// Area Risultati
const phRisultati = document.getElementById('placeholder-risultati');
const listaDati = document.getElementById('lista-dati');
const errCodice = document.getElementById('error-codice');

// Area Immagine
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

async function eseguiRicerca() {
    const rawCode = inputCodice.value.trim();
    if (!rawCode) return; 

    const searchCode = rawCode.toUpperCase(); 

    loadingOverlay.classList.remove('hidden');

    try {
        // Scarica il file Excel tramite il Proxy
        // Aggiungiamo un timestamp per evitare la cache
        const response = await fetch(EXCEL_URL + "&t=" + new Date().getTime());
        
        if (!response.ok) throw new Error("Errore scaricamento Excel");

        const arrayBuffer = await response.arrayBuffer();

        // Leggi i dati
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

        // Ritardo per animazione
        setTimeout(() => {
            elaboraDati(jsonData, searchCode);
            loadingOverlay.classList.add('hidden');
        }, 3500); 

    } catch (error) {
        console.error("Errore:", error);
        loadingOverlay.classList.add('hidden');
        // Messaggio più specifico per debug (visibile solo in console sviluppatore se serve)
        // L'utente vedrà ancora il popup
        alert("Errore di connessione: Impossibile leggere il file Excel.\nAssicurati che il link OneDrive sia pubblico.");
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

    if (imageLinkFound && imageLinkFound.trim() !== "") {
        if (imageLinkFound.includes("drive.google.com")) {
             imageLinkFound = imageLinkFound.replace("/view", "/preview").replace("open?id=", "uc?id=");
        }
        imgResult.src = imageLinkFound;
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