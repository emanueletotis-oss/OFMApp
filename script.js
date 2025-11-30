// --- CONFIGURAZIONE ---

// ID del tuo file Google Drive (estratto dal link che mi hai dato)
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

async function eseguiRicerca() {
    const rawCode = inputCodice.value.trim();
    if (!rawCode) return; 

    // Pulisce il codice (toglie spazi e mette maiuscolo)
    const searchCode = rawCode.replace(/\s+/g, '').toUpperCase(); 
    loadingOverlay.classList.remove('hidden');

    try {
        // 1. COSTRUZIONE LINK EXPORT
        // Questo link dice a Google: "Dammi il file ID in formato Excel (.xlsx)"
        let exportUrl = `https://docs.google.com/spreadsheets/d/${FILE_ID}/export?format=xlsx`;
        
        // Aggiungiamo un numero casuale per evitare che il telefono usi la memoria vecchia (cache)
        exportUrl += "&t=" + new Date().getTime();

        // 2. DOWNLOAD TRAMITE PROXY (AllOrigins)
        // Usiamo questo ponte per evitare che il browser blocchi il download per sicurezza
        const proxyUrl = `https://api.allorigins.win/raw?url=${encodeURIComponent(exportUrl)}`;
        
        console.log("Scaricando da:", proxyUrl);

        const response = await fetch(proxyUrl, { method: 'GET', cache: 'no-store' });
        
        if (!response.ok) throw new Error("Errore connessione: " + response.status);

        const arrayBuffer = await response.arrayBuffer();

        // Controllo se il file è valido (se inizia con '<' è una pagina web di errore/login)
        const firstByte = new Uint8Array(arrayBuffer)[0];
        if (firstByte === 60) throw new Error("Accesso negato. Assicurati che il file su Drive sia 'Chiunque abbia il link'.");

        // 3. LETTURA EXCEL
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const sheetName = workbook.SheetNames[0]; // Prende il primo foglio
        const worksheet = workbook.Sheets[sheetName];
        
        // Converte in matrice (Tutti i dati grezzi)
        const matrix = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

        if (!matrix || matrix.length === 0) throw new Error("Il file Excel è vuoto.");

        // Ritardo estetico (3 secondi) per mostrare l'animazione
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

    // --- RICERCA INTELLIGENTE ---
    // Scorre tutte le righe del file
    for (let i = 0; i < matrix.length; i++) {
        const row = matrix[i];
        
        // Cerca il codice in ogni cella della riga
        for (let j = 0; j < row.length; j++) {
            let cell = String(row[j]).toUpperCase().replace(/\s+/g, '');
            
            if (cell === searchCode) {
                recordTrovato = row;
                
                // Cerca di individuare le intestazioni (solitamente riga 0)
                if (i > 0) headers = matrix[0]; 
                else headers = row.map((_, idx) => `Dato ${idx + 1}`);
                break;
            }
        }
        if (recordTrovato) break;
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
    let imageLinkFound = "";
    const maxLen = Math.max(headers.length, recordTrovato.length);

    for (let k = 0; k < maxLen; k++) {
        let key = headers[k] || `Colonna ${k+1}`;
        let val = recordTrovato[k];

        if (!val || String(val).trim() === "") continue;

        let keyLower = String(key).toLowerCase();
        let valString = String(val);

        // --- RILEVAMENTO IMMAGINE ---
        // Se la colonna si chiama "Immagine", "Foto", "Link" 
        // OPPURE se il valore sembra un link di Google Drive
        let isImageColumn = keyLower.includes('immagine') || keyLower.includes('foto') || keyLower.includes('url') || keyLower.includes('link');
        let isDriveLink = valString.includes('drive.google.com') || valString.includes('docs.google.com');

        if (isImageColumn || isDriveLink) {
            imageLinkFound = valString;
            continue; // Non stampare il link nell'elenco testuale
        }

        // Non ristampare il codice cercato nell'elenco
        let valClean = valString.toUpperCase().replace(/\s+/g, '');
        if (valClean === searchCode) continue;

        // Stampa il dato
        const div = document.createElement('div');
        div.className = 'data-item';
        div.innerText = `- ${key}: ${val}`;
        listaDati.appendChild(div);
    }

    // --- VISUALIZZAZIONE IMMAGINE ---
    if (imageLinkFound) {
        let imgUrl = imageLinkFound;

        // FIX LINK DRIVE: Se hai incollato il link normale di condivisione foto,
        // questo codice lo trasforma automaticamente in un link visibile nell'app.
        if (imgUrl.includes("/d/")) {
            let idMatch = imgUrl.match(/\/d\/(.*?)\//);
            if (idMatch) {
                let imgId = idMatch[1];
                // Trasforma view -> preview (più compatibile)
                imgUrl = `https://drive.google.com/uc?export=view&id=${imgId}`;
            }
        }

        imgResult.src = imgUrl;
        imgResult.classList.remove('hidden');
        
        imgResult.onerror = () => {
            console.warn("Impossibile caricare immagine:", imgUrl);
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