// --- CONFIGURAZIONE ---
const USER_LINK = "https://1drv.ms/x/c/ac1a912c65f087d9/IQSCrL3EW_MNQJi_FLvY8KNJAXXS-7KsHuornWAqYgAoNnE";

// --- ELEMENTI DOM ---
const inputCodice = document.getElementById('input-codice');
const btnPlay = document.getElementById('btn-play');
const btnReset = document.getElementById('btn-reset');
const loadingOverlay = document.getElementById('loading-overlay');
const listaDati = document.getElementById('lista-dati');
const errCodice = document.getElementById('error-codice');
const imgResult = document.getElementById('result-image');
const errImmagine = document.getElementById('error-immagine');
const phRisultati = document.getElementById('placeholder-risultati');
const phImmagine = document.getElementById('placeholder-immagine');

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
    const rawCode = inputCodice.value.trim();
    if (!rawCode) return; 

    // Pulisce il codice cercato (toglie spazi e mette maiuscolo)
    const searchCode = rawCode.replace(/\s+/g, '').toUpperCase(); 
    loadingOverlay.classList.remove('hidden');

    try {
        let directLink = USER_LINK.replace('/x/', '/download/').split('?')[0];
        let excelData = null;
        const timestamp = new Date().getTime();

        // Tentativi Download a Cascata
        if (!excelData) excelData = await tryFetch(`https://api.codetabs.com/v1/proxy?quest=${encodeURIComponent(directLink)}&t=${timestamp}`, "CodeTabs");
        if (!excelData) excelData = await tryFetch(`https://corsproxy.io/?${encodeURIComponent(directLink)}&t=${timestamp}`, "CorsProxy");
        if (!excelData) excelData = await tryFetch(`https://api.allorigins.win/raw?url=${encodeURIComponent(directLink)}&t=${timestamp}`, "AllOrigins");

        if (!excelData) throw new Error("Impossibile scaricare il file. Riprova.");

        // LETTURA RAW (Grezza)
        const workbook = XLSX.read(excelData, { type: 'array' });
        
        // Prende il PRIMO foglio
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        
        // Converte in MATRICE (Array di Array) per vedere tutto, anche intestazioni spostate
        const matrix = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

        loadingOverlay.classList.add('hidden');

        // --- DIAGNOSTICA VISIVA (Quello che mi serve vedere) ---
        let msg = "DIAGNOSTICA (Fai screen e mandamelo):\n\n";
        msg += "Righe totali lette: " + matrix.length + "\n";
        msg += "Codice cercato: [" + searchCode + "]\n\n";
        
        if (matrix.length > 0) msg += "RIGA 1: " + JSON.stringify(matrix[0]) + "\n";
        if (matrix.length > 1) msg += "RIGA 2: " + JSON.stringify(matrix[1]) + "\n";
        if (matrix.length > 2) msg += "RIGA 3: " + JSON.stringify(matrix[2]) + "\n";
        if (matrix.length > 3) msg += "RIGA 4: " + JSON.stringify(matrix[3]) + "\n";

        // Mostra il popup con i dati grezzi
        alert(msg);

        // Ora proviamo comunque a cercare
        elaboraDati(matrix, searchCode);

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

    // Cerchiamo in TUTTE le righe
    for (let i = 0; i < matrix.length; i++) {
        const row = matrix[i];
        
        // Controlliamo ogni cella della riga
        for (let j = 0; j < row.length; j++) {
            let cell = String(row[j]).toUpperCase().replace(/\s+/g, ''); // Pulizia estrema
            
            if (cell === searchCode) {
                recordTrovato = row;
                
                // Tentativo di trovare le intestazioni
                // Se il dato è alla riga 5, usiamo la riga 4 come header. 
                // Se è alla riga 0, usiamo header generici.
                if (i > 0) headers = matrix[0]; // Assumiamo header sempre in riga 1 per ora
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

    // Trovato!
    listaDati.classList.remove('hidden');
    let imageLinkFound = "";

    // Stampa dati
    const maxLen = Math.max(headers.length, recordTrovato.length);
    for (let k = 0; k < maxLen; k++) {
        let key = headers[k] || `Colonna ${k+1}`;
        let val = recordTrovato[k];

        if (!val || String(val).trim() === "") continue;

        let keyLower = String(key).toLowerCase();
        
        // Rileva Immagine
        if (keyLower.includes('immagine') || keyLower.includes('link') || keyLower.includes('url')) {
            imageLinkFound = val;
            continue;
        }

        // Evita di ristampare il codice stesso
        let valClean = String(val).toUpperCase().replace(/\s+/g, '');
        if (valClean === searchCode) continue;

        const div = document.createElement('div');
        div.className = 'data-item';
        div.innerText = `- ${key}: ${val}`;
        listaDati.appendChild(div);
    }

    if (imageLinkFound) {
        let imgUrl = String(imageLinkFound);
        if (imgUrl.includes("drive.google.com")) {
             imgUrl = imgUrl.replace("/view", "/preview").replace("open?id=", "uc?id=");
        }
        imgResult.src = imgUrl;
        imgResult.classList.remove('hidden');
    } else {
        errImmagine.classList.remove('hidden');
    }
}

btnPlay.addEventListener('click', eseguiRicerca);
btnReset.addEventListener('click', resetApp);
inputCodice.addEventListener('keypress', (e) => {
    if (e.key === 'Enter') eseguiRicerca();
});