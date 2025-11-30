// --- CONFIGURAZIONE ---
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

async function tryFetch(url, proxyName) {
    try {
        console.log(`Tentativo ${proxyName}:`, url);
        const response = await fetch(url, { method: 'GET', cache: 'no-store' });
        if (!response.ok) throw new Error(`HTTP ${response.status}`);
        const arrayBuffer = await response.arrayBuffer();
        const firstByte = new Uint8Array(arrayBuffer)[0];
        if (firstByte === 60) throw new Error("Ricevuto HTML invece di Excel"); 
        return arrayBuffer;
    } catch (e) {
        console.warn(`${proxyName} errore:`, e.message);
        return null;
    }
}

async function eseguiRicerca() {
    const rawCode = inputCodice.value.trim();
    if (!rawCode) return; 

    // Normalizziamo il codice cercato (rimuoviamo spazi e tutto maiuscolo)
    // Sostituiamo anche eventuali spazi interni per sicurezza
    const searchCode = rawCode.replace(/\s+/g, '').toUpperCase(); 
    
    loadingOverlay.classList.remove('hidden');

    try {
        let directLink = USER_LINK.replace('/x/', '/download/').split('?')[0];
        let excelData = null;
        const timestamp = new Date().getTime();

        // TENTATIVI DOWNLOAD
        if (!excelData) excelData = await tryFetch(`https://api.codetabs.com/v1/proxy?quest=${encodeURIComponent(directLink)}&t=${timestamp}`, "CodeTabs");
        if (!excelData) excelData = await tryFetch(`https://corsproxy.io/?${encodeURIComponent(directLink)}&t=${timestamp}`, "CorsProxy");
        if (!excelData) excelData = await tryFetch(`https://api.allorigins.win/raw?url=${encodeURIComponent(directLink)}&t=${timestamp}`, "AllOrigins");

        if (!excelData) throw new Error("Connessione fallita. Riprova.");

        // LETTURA EXCEL AVANZATA
        const workbook = XLSX.read(excelData, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        
        // Convertiamo usando 'header: 1' per ottenere una matrice grezza (Array di Array)
        // Questo evita problemi se le intestazioni non sono nella riga 1
        const jsonMatrix = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

        // Ritardo estetico
        setTimeout(() => {
            elaboraDatiMatrice(jsonMatrix, searchCode);
            loadingOverlay.classList.add('hidden');
        }, 3000); 

    } catch (error) {
        loadingOverlay.classList.add('hidden');
        alert("Errore tecnico: " + error.message);
    }
}

function elaboraDatiMatrice(matrix, searchCode) {
    errCodice.classList.add('hidden');
    errImmagine.classList.add('hidden');
    phRisultati.classList.add('hidden');
    phImmagine.classList.add('hidden');
    listaDati.innerHTML = "";

    let headers = [];
    let recordTrovato = null;
    let headerRowIndex = -1;

    // 1. CERCA LA RIGA DELLE INTESTAZIONI E IL DATO
    // Scansioniamo tutta la matrice per trovare dove sono i dati
    for (let i = 0; i < matrix.length; i++) {
        const row = matrix[i];
        
        // Controlliamo se in questa riga c'è il codice cercato
        for (let j = 0; j < row.length; j++) {
            let cellVal = String(row[j]).replace(/\s+/g, '').toUpperCase(); // Pulisce cella
            
            if (cellVal === searchCode) {
                recordTrovato = row;
                
                // Se abbiamo trovato il dato, cerchiamo di capire dove sono le intestazioni.
                // Assumiamo che le intestazioni siano nella riga PRECEDENTE (i-1) o la PRIMA riga (0)
                // Se siamo alla riga 0, non ci sono intestazioni sopra, quindi usiamo indici generici.
                if (i > 0) {
                    headers = matrix[0]; // Proviamo a prendere la riga 0 come intestazioni standard
                } else {
                    headers = row.map((_, idx) => `Colonna ${idx + 1}`); // Intestazioni fittizie
                }
                break;
            }
        }
        if (recordTrovato) break;
    }

    // --- DEBUGGING FONDAMENTALE ---
    // Se non troviamo nulla, mostriamo all'utente cosa ha letto l'app
    if (!recordTrovato) {
        listaDati.classList.add('hidden');
        errCodice.classList.remove('hidden');
        imgResult.classList.add('hidden');
        errImmagine.classList.remove('hidden');

        // Creiamo un messaggio di debug per capire cosa sta succedendo
        let debugMsg = "DIAGNOSTICA (Fai uno screenshot):\n";
        debugMsg += `Ho scaricato ${matrix.length} righe.\n`;
        debugMsg += `Codice cercato: "${searchCode}"\n`;
        
        if (matrix.length > 0) {
            // Mostra il contenuto delle prime 3 righe per capire la struttura
            let firstRows = matrix.slice(0, 3).map(r => JSON.stringify(r)).join("\n");
            debugMsg += `\nPrime righe lette dal file:\n${firstRows}`;
        } else {
            debugMsg += "\nIl file sembra vuoto!";
        }
        
        alert(debugMsg);
        return;
    }

    // --- MOSTRA I DATI TROVATI ---
    listaDati.classList.remove('hidden');

    let imageLinkFound = "";

    // Mappa headers con valori
    // Se headers è vuoto o più corto della riga, gestiamo l'errore
    const maxLen = Math.max(headers.length, recordTrovato.length);

    for (let k = 0; k < maxLen; k++) {
        let key = headers[k] || `Colonna ${k+1}`; // Nome colonna
        let val = recordTrovato[k];              // Valore cella

        if (val === undefined || val === null || String(val).trim() === "") continue;

        let keyClean = String(key).toLowerCase();
        let valClean = String(val).replace(/\s+/g, '').toUpperCase();

        // Salta la cella che contiene il codice stesso
        if (valClean === searchCode) continue;

        // Cerca Immagine
        if (keyClean.includes('immagine') || keyClean.includes('link') || keyClean.includes('foto') || keyClean.includes('url')) {
            imageLinkFound = val;
            continue;
        }

        const div = document.createElement('div');
        div.className = 'data-item';
        div.innerText = `- ${key}: ${val}`;
        listaDati.appendChild(div);
    }

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