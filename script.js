// --- CONFIGURAZIONE ---

// Link diretto al file Excel su OneDrive (aggiornato con il tuo codice specifico)
// ?download=1 forza lo scaricamento del file invece dell'apertura della pagina web
const EXCEL_URL = "https://1drv.ms/x/c/ac1a912c65f087d9/IQSCrL3EW_MNQJi_FLvY8KNJAXXS-7KsHuornWAqYgAoNnE?download=1";

// --- ELEMENTI DOM (Interfaccia Utente) ---
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

/**
 * Resetta l'interfaccia allo stato iniziale
 */
function resetApp() {
    inputCodice.value = "";
    
    // Ripristina area risultati
    phRisultati.classList.remove('hidden');
    listaDati.classList.add('hidden');
    errCodice.classList.add('hidden');
    listaDati.innerHTML = "";
    
    // Ripristina area immagine
    phImmagine.classList.remove('hidden');
    imgResult.classList.add('hidden');
    errImmagine.classList.add('hidden');
    imgResult.src = "";
}

/**
 * Funzione principale: Scarica Excel, cerca il codice, mostra i risultati
 */
async function eseguiRicerca() {
    const rawCode = inputCodice.value.trim();
    if (!rawCode) return; // Se vuoto non fa nulla

    const searchCode = rawCode.toUpperCase(); // Normalizza in maiuscolo

    // 1. Avvia l'animazione di caricamento (Overlay scuro con cerchio pulsante)
    loadingOverlay.classList.remove('hidden');

    try {
        // 2. Scarica il file Excel più recente
        // Aggiungiamo un timestamp (&t=...) per evitare che il browser usi una versione vecchia salvata nella cache
        const response = await fetch(EXCEL_URL + "&t=" + new Date().getTime());
        
        if (!response.ok) throw new Error("Errore durante lo scaricamento del file Excel");

        const arrayBuffer = await response.arrayBuffer();

        // 3. Leggi i dati usando la libreria SheetJS
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        
        // Prende il primo foglio di lavoro disponibile
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        
        // Converte il foglio in un oggetto JSON (Lista di dati)
        // defval: "" assicura che le celle vuote non rompano il codice
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

        // 4. Ritardo artificiale per mostrare l'animazione (Richiesta utente: > 3 sec)
        setTimeout(() => {
            elaboraDati(jsonData, searchCode);
            loadingOverlay.classList.add('hidden'); // Nasconde animazione alla fine
        }, 3500); // 3500 millisecondi = 3.5 secondi

    } catch (error) {
        console.error("Errore:", error);
        loadingOverlay.classList.add('hidden');
        alert("Errore di connessione: Impossibile leggere il file Excel aggiornato. Controlla il link o la connessione internet.");
    }
}

/**
 * Cerca il codice nei dati scaricati e aggiorna l'interfaccia
 */
function elaboraDati(data, code) {
    // Nascondi tutti i messaggi precedenti
    errCodice.classList.add('hidden');
    errImmagine.classList.add('hidden');
    phRisultati.classList.add('hidden');
    phImmagine.classList.add('hidden');
    listaDati.innerHTML = "";

    // 1. CERCA LA RIGA CORRISPONDENTE AL CODICE
    let recordTrovato = null;

    for (let row of data) {
        // Cerca una colonna che si chiami "Codice" (case insensitive) oppure usa la prima colonna
        let rowCode = "";
        
        if (row['Codice']) rowCode = row['Codice'];
        else if (row['codice']) rowCode = row['codice'];
        else rowCode = Object.values(row)[0]; // Fallback: prende il valore della prima colonna
        
        // Confronto insensibile a maiuscole/minuscole
        if (String(rowCode).toUpperCase() === code) {
            recordTrovato = row;
            break;
        }
    }

    // --- CASO 1: CODICE NON TROVATO ---
    if (!recordTrovato) {
        listaDati.classList.add('hidden');
        errCodice.classList.remove('hidden'); // Mostra triangolo rosso
        
        imgResult.classList.add('hidden');
        errImmagine.classList.remove('hidden'); // Mostra triangolo dorato (anche l'immagine manca se manca il codice)
        return;
    }

    // --- CASO 2: CODICE TROVATO ---
    listaDati.classList.remove('hidden');

    // Estrai le chiavi (intestazioni colonne)
    const keys = Object.keys(recordTrovato);
    let imageLinkFound = "";

    keys.forEach(key => {
        const val = recordTrovato[key];
        const keyLower = key.toLowerCase();

        // Logica per identificare se questa colonna è un link immagine
        if (keyLower.includes('immagine') || keyLower.includes('link') || keyLower.includes('foto') || keyLower.includes('url')) {
            imageLinkFound = val; // Salviamo il link ma non lo stampiamo nel testo
            return;
        }

        // Stampiamo il dato solo se:
        // 1. Non è la colonna "Codice"
        // 2. Il valore non è vuoto
        if (keyLower !== 'codice' && val !== "" && val !== undefined) {
            const div = document.createElement('div');
            div.className = 'data-item';
            div.innerText = `- ${key}: ${val}`;
            listaDati.appendChild(div);
        }
    });

    // --- GESTIONE IMMAGINE ---
    // Se nel file excel c'era una colonna con un link, lo usiamo.
    if (imageLinkFound && imageLinkFound.trim() !== "") {
        
        // Fix per link Google Drive se necessario (converte 'view' o 'open' in link diretto se pubblico)
        if (imageLinkFound.includes("drive.google.com")) {
             imageLinkFound = imageLinkFound.replace("/view", "/preview").replace("open?id=", "uc?id=");
        }

        imgResult.src = imageLinkFound;
        imgResult.classList.remove('hidden');
        
        // Se l'immagine non si carica (link rotto o privato), mostra errore
        imgResult.onerror = () => {
            imgResult.classList.add('hidden');
            errImmagine.classList.remove('hidden');
        };

    } else {
        // Nessun link immagine trovato nella riga Excel
        imgResult.classList.add('hidden');
        errImmagine.classList.remove('hidden');
    }
}

// --- GESTIONE EVENTI (Click e Tasti) ---

// Click sul tasto Play
btnPlay.addEventListener('click', eseguiRicerca);

// Click sul tasto Reset
btnReset.addEventListener('click', resetApp);

// Pressione tasto "Invio" nella casella di testo
inputCodice.addEventListener('keypress', (e) => {
    if (e.key === 'Enter') eseguiRicerca();
});