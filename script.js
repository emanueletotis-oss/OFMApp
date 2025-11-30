// --- CONFIGURAZIONE ---

// IL TUO LINK (Non toccarlo, lo script lo gestisce in automatico)
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

// Funzione helper per tentare il download
async function tryFetch(url, proxyName) {
    try {
        console.log(`Tentativo con ${proxyName}:`, url);
        const response = await fetch(url, { method: 'GET', cache: 'no-store' });
        
        if (!response.ok) throw new Error(`HTTP ${response.status}`);
        
        const arrayBuffer = await response.arrayBuffer();
        
        // Verifica se abbiamo scaricato HTML (errore) invece di un file binario
        const firstByte = new Uint8Array(arrayBuffer)[0];
        if (firstByte === 60) throw new Error("Ricevuto HTML (pagina web) invece del file Excel"); 
        
        return arrayBuffer;
    } catch (e) {
        console.warn(`${proxyName} fallito:`, e.message);
        return null;
    }
}

async function eseguiRicerca() {
    const rawCode = inputCodice.value.trim();
    if (!rawCode) return; 

    const searchCode = rawCode.toUpperCase(); 
    loadingOverlay.classList.remove('hidden');

    try {
        // 1. TRASFORMAZIONE CHIRURGICA DEL LINK
        // Trasformiamo: https://1drv.ms/x/c/... -> https://1drv.ms/download/c/...
        // Questo è il metodo più sicuro per i nuovi link "Streamline" di OneDrive.
        let directLink = USER_LINK.replace('/x/', '/download/');
        directLink = directLink.split('?')[0]; // Rimuove parametri sporchi

        let excelData = null;
        const timestamp = new Date().getTime();

        // --- STRATEGIA A CASCATA (3 Tentativi) ---

        // TENTATIVO 1: CodeTabs (Molto affidabile per file Excel)
        if (!excelData) {
            const url1 = `https://api.codetabs.com/v1/proxy?quest=${encodeURIComponent(directLink)}&t=${timestamp}`;
            excelData = await tryFetch(url1, "CodeTabs");
        }

        // TENTATIVO 2: CorsProxy.io (Veloce, ma a volte bloccato su mobile)
        if (!excelData) {
            const url2 = `https://corsproxy.io/?${encodeURIComponent(directLink)}&t=${timestamp}`;
            excelData = await tryFetch(url2, "CorsProxy");
        }

        // TENTATIVO 3: AllOrigins (Lento ma spesso passa ovunque)
        if (!excelData) {
            const url3 = `https://api.allorigins.win/raw?url=${encodeURIComponent(directLink)}&t=${timestamp}`;
            excelData = await tryFetch(url3, "AllOrigins");
        }

        // Se tutti falliscono
        if (!excelData) {
            throw new Error("Tutti i tentativi di connessione sono falliti. Il firewall della rete mobile potrebbe bloccare i download.");
        }

        // 2. ELABORAZIONE FILE
        const workbook = XLSX.read(excelData, { type: 'array' });
        
        if (!workbook.SheetNames.length) throw new Error("File Excel vuoto o illeggibile.");
        
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

        // Ritardo estetico per l'animazione
        setTimeout(() => {
            elaboraDati(jsonData, searchCode);
            loadingOverlay.classList.add('hidden');
        }, 3500); 

    } catch (error) {
        console.error("ERRORE CRITICO:", error);
        loadingOverlay.classList.add('hidden');
        alert("Errore: " + error.message + "\n\nProva a disattivare il Wi-Fi e usare solo dati, o viceversa.");
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

// Event Listeners
btnPlay.addEventListener('click', eseguiRicerca);
btnReset.addEventListener('click', resetApp);
inputCodice.addEventListener('keypress', (e) => {
    if (e.key === 'Enter') eseguiRicerca();
});