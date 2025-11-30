// --- CONFIGURAZIONE ---

// Il tuo link 1drv.ms (che trasformeremo automaticamente in link di download)
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

// Funzione di supporto per tentare il download da diversi proxy
async function tryFetch(proxyUrl) {
    try {
        console.log("Tentativo download da:", proxyUrl);
        const response = await fetch(proxyUrl, {
            method: 'GET',
            cache: 'no-store' // IMPORTANTE: Forza l'iPhone a non usare la cache
        });
        
        if (!response.ok) throw new Error("Status: " + response.status);
        
        const arrayBuffer = await response.arrayBuffer();
        
        // Controllo se è HTML (Errore) invece di Excel
        const firstByte = new Uint8Array(arrayBuffer)[0];
        if (firstByte === 60) throw new Error("Ricevuto HTML invece di Excel"); // 60 = '<'
        
        return arrayBuffer;
    } catch (e) {
        console.warn("Proxy fallito:", e);
        return null; // Ritorna null se fallisce
    }
}

async function eseguiRicerca() {
    const rawCode = inputCodice.value.trim();
    if (!rawCode) return; 

    const searchCode = rawCode.toUpperCase(); 
    loadingOverlay.classList.remove('hidden');

    try {
        // PREPARAZIONE LINK
        // 1. Modifica il link per puntare al download
        let downloadLink = USER_LINK.replace('/x/', '/download/');
        // Rimuove parametri extra
        downloadLink = downloadLink.split('?')[0]; 

        // STRATEGIA DOPPIO PROXY
        let excelData = null;

        // TENTATIVO 1: Usa 'AllOrigins' (spesso funziona meglio su mobile)
        const proxy1 = "https://api.allorigins.win/raw?url=" + encodeURIComponent(downloadLink) + "&timestamp=" + new Date().getTime();
        excelData = await tryFetch(proxy1);

        // TENTATIVO 2: Se il primo fallisce, usa 'CorsProxy'
        if (!excelData) {
            console.log("Primo proxy fallito, provo il secondo...");
            const proxy2 = "https://corsproxy.io/?" + encodeURIComponent(downloadLink) + "&timestamp=" + new Date().getTime();
            excelData = await tryFetch(proxy2);
        }

        // Se entrambi falliscono
        if (!excelData) {
            throw new Error("Connessione instabile. Impossibile scaricare il file da OneDrive con la rete attuale.");
        }

        // ELABORAZIONE FILE
        const workbook = XLSX.read(excelData, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

        if (jsonData.length === 0) throw new Error("Il file Excel è vuoto.");

        // Ritardo estetico
        setTimeout(() => {
            elaboraDati(jsonData, searchCode);
            loadingOverlay.classList.add('hidden');
        }, 3500); 

    } catch (error) {
        console.error("ERRORE FATALE:", error);
        loadingOverlay.classList.add('hidden');
        alert(error.message + "\n\nSuggerimento: Prova a disattivare e riattivare il Wi-Fi/Dati.");
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