// --- CONFIGURAZIONE ---

// Link originale di OneDrive
const ONEDRIVE_LINK = "https://1drv.ms/x/c/ac1a912c65f087d9/IQSCrL3EW_MNQJi_FLvY8KNJAXXS-7KsHuornWAqYgAoNnE?download=1";

// NUOVO PROXY: Usiamo 'allorigins' che gestisce meglio i redirect di OneDrive
// Aggiungiamo un timestamp per evitare che il browser usi la cache vecchia
const EXCEL_URL = "https://api.allorigins.win/raw?url=" + encodeURIComponent(ONEDRIVE_LINK);

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

async function eseguiRicerca() {
    const rawCode = inputCodice.value.trim();
    if (!rawCode) return; 

    const searchCode = rawCode.toUpperCase(); 

    loadingOverlay.classList.remove('hidden');

    try {
        // Aggiungiamo &timestamp alla fine del proxy url per forzare il refresh
        const finalUrl = EXCEL_URL + "&timestamp=" + new Date().getTime();
        
        console.log("Tentativo download da:", finalUrl); // Per debug

        const response = await fetch(finalUrl);
        
        if (!response.ok) {
            throw new Error(`Errore HTTP: ${response.status}`);
        }

        const arrayBuffer = await response.arrayBuffer();

        // CONTROLLO DI SICUREZZA
        // A volte OneDrive restituisce una pagina HTML di errore invece del file Excel.
        // Se il file inizia con "<", è probabile che sia HTML e non un file XLSX.
        const firstByte = new Uint8Array(arrayBuffer)[0];
        if (firstByte === 60) { // 60 è il codice ASCII per '<'
            throw new Error("OneDrive ha restituito una pagina Web invece del file Excel. Link scaduto o bloccato.");
        }

        // Lettura Excel
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const sheetName = workbook.SheetNames[0]; // Prende il primo foglio
        const worksheet = workbook.Sheets[sheetName];
        
        // Converte in JSON
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

        if (jsonData.length === 0) {
            throw new Error("Il file Excel sembra vuoto o non leggibile.");
        }

        // Ritardo per animazione
        setTimeout(() => {
            elaboraDati(jsonData, searchCode);
            loadingOverlay.classList.add('hidden');
        }, 3500); 

    } catch (error) {
        console.error("Dettaglio Errore:", error);
        loadingOverlay.classList.add('hidden');
        // Mostriamo l'errore specifico nel popup
        alert("Errore tecnico: " + error.message + "\n\nControlla che il link OneDrive sia ancora valido e pubblico.");
    }
}

function elaboraDati(data, code) {
    errCodice.classList.add('hidden');
    errImmagine.classList.add('hidden');
    phRisultati.classList.add('hidden');
    phImmagine.classList.add('hidden');
    listaDati.innerHTML = "";

    let recordTrovato = null;

    // Cerca la riga
    for (let row of data) {
        // Cerca colonna 'Codice' o usa la prima colonna disponibile
        let rowCode = "";
        if (row['Codice']) rowCode = row['Codice'];
        else if (row['codice']) rowCode = row['codice'];
        else rowCode = Object.values(row)[0]; 
        
        if (String(rowCode).toUpperCase() === code) {
            recordTrovato = row;
            break;
        }
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

    const keys = Object.keys(recordTrovato);
    let imageLinkFound = "";

    keys.forEach(key => {
        const val = recordTrovato[key];
        const keyLower = key.toLowerCase();

        // Cerca link immagine nascosto
        if (keyLower.includes('immagine') || keyLower.includes('link') || keyLower.includes('foto') || keyLower.includes('url')) {
            imageLinkFound = val; 
            return;
        }

        // Mostra dati (escludendo codice e vuoti)
        if (keyLower !== 'codice' && val !== "" && val !== undefined) {
            const div = document.createElement('div');
            div.className = 'data-item';
            div.innerText = `- ${key}: ${val}`;
            listaDati.appendChild(div);
        }
    });

    // Gestione Immagine
    if (imageLinkFound && String(imageLinkFound).trim() !== "") {
        // Fix link drive
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