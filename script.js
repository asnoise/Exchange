/**
 * Excel Currency Converter
 * Logic: Reads Column Q as Currency, Column R as INR Output
 */

const Converter = (() => {
    const state = {
        wb: null,
        fileName: '',
        processedWb: null,
        rates: null,
        fallbackRates: {
            EUR: 102.0, USD: 87.0, GBP: 118.0, JPY: 0.60,
            CAD: 64.0, AUD: 57.0, SGD: 67.0, AED: 24.0, INR: 1.0
        }
    };

    const dom = {
        fileInput: document.getElementById('fileInput'),
        dropZone: document.getElementById('dropZone'),
        convertBtn: document.getElementById('convertBtn'),
        downloadBtn: document.getElementById('downloadBtn'),
        fileNameDisplay: document.getElementById('fileNameDisplay'),
        progressContainer: document.getElementById('progressContainer'),
        progressFill: document.getElementById('progressFill'),
        progressStatus: document.getElementById('progressStatus'),
        progressPercent: document.getElementById('progressPercent'),
        summaryGrid: document.getElementById('summaryGrid'),
        statRows: document.getElementById('statRows'),
        statTime: document.getElementById('statTime'),
        statApi: document.getElementById('statApi'),
        themeToggle: document.getElementById('themeToggle'),
        toastContainer: document.getElementById('toastContainer')
    };

    const init = () => {
        bindEvents();
        fetchRates();
        checkTheme();
    };

    const bindEvents = () => {
        dom.dropZone.addEventListener('click', () => dom.fileInput.click());
        dom.fileInput.addEventListener('change', e => handleFile(e.target.files[0]));
        
        dom.dropZone.addEventListener('dragover', e => { e.preventDefault(); dom.dropZone.style.borderColor = 'var(--primary)'; });
        dom.dropZone.addEventListener('dragleave', () => dom.dropZone.style.borderColor = '');
        dom.dropZone.addEventListener('drop', e => {
            e.preventDefault();
            dom.dropZone.style.borderColor = '';
            handleFile(e.dataTransfer.files[0]);
        });

        dom.convertBtn.addEventListener('click', processExcel);
        dom.downloadBtn.addEventListener('click', downloadOutput);
        dom.themeToggle.addEventListener('click', toggleTheme);
    };

    async function fetchRates() {
        try {
            const res = await fetch('https://open.er-api.com/v6/latest/USD');
            const data = await res.json();
            const inrBase = data.rates.INR;
            
            // Normalize all rates to INR (1 Unit = X INR)
            state.rates = {};
            for (const [curr, rate] of Object.entries(data.rates)) {
                state.rates[curr] = (1 / rate) * inrBase;
            }
            dom.statApi.textContent = "Live API";
            notify("Exchange rates updated", "success");
        } catch (e) {
            state.rates = state.fallbackRates;
            dom.statApi.textContent = "Fallback";
            notify("Using offline exchange rates", "error");
        }
    }

    function handleFile(file) {
        if (!file) return;
        if (!file.name.match(/\.(xlsx|xls)$/)) {
            return notify("Please upload an Excel file (.xlsx or .xls)", "error");
        }

        state.fileName = file.name;
        dom.fileNameDisplay.textContent = file.name;
        
        const reader = new FileReader();
        reader.onload = e => {
            const data = new Uint8Array(e.target.result);
            state.wb = XLSX.read(data, { type: 'array' });
            dom.convertBtn.disabled = false;
            dom.downloadBtn.disabled = true;
            dom.summaryGrid.style.display = 'none';
            notify("File loaded successfully");
        };
        reader.readAsArrayBuffer(file);
    }

    async function processExcel() {
        if (!state.wb) return;
        const startTime = performance.now();
        
        dom.convertBtn.disabled = true;
        dom.progressContainer.style.display = 'block';
        updateProgress(10, "Reading Sheet...");

        try {
            const sheetName = state.wb.SheetNames[0];
            const ws = state.wb.Sheets[sheetName];
            
            // Convert to Array of Arrays to strictly control Column Q (16) and R (17)
            const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
            
            if (data.length < 2) throw new Error("Excel sheet is empty or has no data rows.");

            // Smart Detection for Earnings Column ($)
            // We search headers for "Earnings", "Amount", or index P (15) as fallback
            const headers = data[0];
            let earningsIdx = headers.findIndex(h => {
                const s = String(h).toLowerCase();
                return s.includes('earnings($)') || s.includes('earning ($)') || s.includes('earnings') || s.includes('amount');
            });

            // Fallback to Column P (15) if earnings column not found by name
            if (earningsIdx === -1) earningsIdx = 15;

            // Column Q = Index 16 (Currency Source)
            // Column R = Index 17 (INR Target)
            const currencyIdx = 16; 
            const inrIdx = 17;

            // Update Header for Column R
            data[0][inrIdx] = "Earning (₹)";

            const totalRows = data.length - 1;
            
            // Process rows (skip header)
            for (let i = 1; i < data.length; i++) {
                const row = data[i];
                
                // Ensure the row has enough columns
                while (row.length <= inrIdx) row.push("");

                const currCode = String(row[currencyIdx] || '').trim().toUpperCase();
                const amount = parseFloat(row[earningsIdx]) || 0;
                
                const rate = state.rates[currCode] || state.rates['USD'];
                const converted = amount * rate;
                
                row[inrIdx] = Number(converted.toFixed(2));

                if (i % 100 === 0) {
                    const pct = 10 + Math.floor((i / totalRows) * 80);
                    updateProgress(pct, `Converting row ${i}...`);
                    // Allow UI to breathe for large files
                    if (i % 1000 === 0) await new Promise(r => setTimeout(r, 0));
                }
            }

            updateProgress(95, "Finalizing file...");

            // Rebuild Workbook
            const newWs = XLSX.utils.aoa_to_sheet(data);
            const newWb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(newWb, newWs, sheetName);

            state.processedWb = newWb;
            
            const endTime = performance.now();
            dom.statRows.textContent = totalRows.toLocaleString();
            dom.statTime.textContent = ((endTime - startTime) / 1000).toFixed(2) + 's';
            
            updateProgress(100, "Conversion Complete");
            dom.downloadBtn.disabled = false;
            dom.summaryGrid.style.display = 'grid';
            notify("Processing finished", "success");

        } catch (err) {
            notify(err.message, "error");
            dom.convertBtn.disabled = false;
        }
    }

    function downloadOutput() {
        if (!state.processedWb) return;
        const outName = `INR_Converted_${state.fileName}`;
        XLSX.writeFile(state.processedWb, outName);
        notify("Downloading updated file...");
    }

    // UI Helpers
    function updateProgress(pct, status) {
        dom.progressFill.style.width = pct + '%';
        dom.progressPercent.textContent = pct + '%';
        dom.progressStatus.textContent = status;
    }

    function notify(msg, type = 'success') {
        const t = document.createElement('div');
        t.className = `toast ${type}`;
        t.innerHTML = `<i class="fas ${type === 'success' ? 'fa-check-circle' : 'fa-exclamation-circle'}"></i><span>${msg}</span>`;
        dom.toastContainer.appendChild(t);
        setTimeout(() => {
            t.style.opacity = '0';
            setTimeout(() => t.remove(), 500);
        }, 3000);
    }

    function toggleTheme() {
        const isDark = document.documentElement.getAttribute('data-theme') === 'dark';
        const newTheme = isDark ? 'light' : 'dark';
        document.documentElement.setAttribute('data-theme', newTheme);
        localStorage.setItem('theme', newTheme);
        dom.themeToggle.innerHTML = isDark ? '<i class="fas fa-moon"></i>' : '<i class="fas fa-sun"></i>';
    }

    function checkTheme() {
        const saved = localStorage.getItem('theme') || 'light';
        document.documentElement.setAttribute('data-theme', saved);
        dom.themeToggle.innerHTML = saved === 'dark' ? '<i class="fas fa-sun"></i>' : '<i class="fas fa-moon"></i>';
    }

    return { init };
})();

document.addEventListener('DOMContentLoaded', Converter.init);