// --- CONFIGURAZIONE ---

// 1. IL TUO LINK ESATTO (Preso dal tuo messaggio)
// Il codice sotto lo correggerà automaticamente per il download.
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

async function eseguiRicerca() {
    const rawCode = inputCodice.value.trim();
    if (!rawCode) return; 

    const searchCode = rawCode.toUpperCase(); 
    loadingOverlay.classList.remove('hidden');

    try {
        // --- TRUCCO PER IL NUOVO ONEDRIVE ---
        // I nuovi link hanno /x/ che apre il visualizzatore web (HTML).
        // Noi cambiamo /x/ con /download/ per forzare il file grezzo.
        let downloadUrl = USER_LINK.replace('/x/', '/download/');
        
        // Puliamo eventuali parametri extra che potrebbero dare fastidio
        downloadUrl = downloadUrl.split('?')[0];

        // Usiamo corsproxy.io che gestisce bene i file binari
        const proxyUrl = "https://corsproxy.io/?" + encodeURIComponent(downloadUrl);
        
        // Aggiungiamo timestamp per evitare la cache del browser
        const finalUrl = proxyUrl + "&t=" + new Date().getTime();

        console.log("Scaricando da:", finalUrl); // Per debug in console

        const response = await fetch(finalUrl);
        
        if (!response.ok) throw new Error("Errore scaricamento: " + response.status);

        const arrayBuffer = await response.arrayBuffer();

        // CONTROLLO DI SICUREZZA:
        // Se il primo byte è '<', significa che abbiamo scaricato una pagina web HTML (errore) 
        // invece del file Excel (binario).
        const firstByte = new Uint8Array(arrayBuffer)[0];
        if (firstByte === 60) { 
            throw new Error("Il link non permette il download diretto. Verifica che il file non sia protetto da password.");
        }

        // Lettura del file Excel
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

        // Ritardo estetico (3.5 secondi)
        setTimeout(() => {
            elaboraDati(jsonData, searchCode);
            loadingOverlay.classList.add('hidden');
        }, 3500); 

    } catch (error) {
        console.error("Errore:", error);
        loadingOverlay.classList.add('hidden');
        alert("Errore tecnico: " + error.message + "\n\nAssicurati che il file Excel su OneDrive non sia stato spostato o rinominato.");
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
        // Cerca colonna codice
        if (row['Codice']) rowCode = row['Codice'];
        else if (row['codice']) rowCode = row['codice'];
        else rowCode = Object.values(row)[0]; // Fallback prima colonna
        
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

        // Cerca URL immagine nascosto
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
        let imgUrl = String(imageLinkFound);
        // Fix per link drive view -> preview
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