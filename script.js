/**
 * Excel Currency Converter - Vanilla JS Engine
 * Handles XLSX parsing, API fetching, and Conversion logic
 */

const App = (() => {
    // Application State
    const state = {
        workbook: null,
        fileName: "",
        convertedWorkbook: null,
        exchangeRates: null,
        isDarkMode: false,
        fallbackRates: {
            EUR: 102,
            USD: 87,
            GBP: 118,
            AED: 24,
            CAD: 64,
            AUD: 57,
            JPY: 0.60,
            SGD: 67,
            INR: 1
        }
    };

    // Cache DOM Elements
    const elements = {
        dropZone: document.getElementById('dropZone'),
        fileInput: document.getElementById('fileInput'),
        fileNameDisplay: document.getElementById('fileNameDisplay'),
        convertBtn: document.getElementById('convertBtn'),
        downloadBtn: document.getElementById('downloadBtn'),
        progressContainer: document.getElementById('progressContainer'),
        progressFill: document.getElementById('progressFill'),
        progressStatus: document.getElementById('progressStatus'),
        progressPercent: document.getElementById('progressPercent'),
        summaryGrid: document.getElementById('summaryGrid'),
        statRows: document.getElementById('statRows'),
        statCurrencies: document.getElementById('statCurrencies'),
        statTime: document.getElementById('statTime'),
        statSource: document.getElementById('statSource'),
        themeToggle: document.getElementById('themeToggle'),
        toastContainer: document.getElementById('toastContainer')
    };

    // Initialize the App
    const init = () => {
        setupEventListeners();
        fetchExchangeRates();
        loadTheme();
    };

    const setupEventListeners = () => {
        // Drag and Drop
        elements.dropZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            elements.dropZone.classList.add('drag-over');
        });

        elements.dropZone.addEventListener('dragleave', () => {
            elements.dropZone.classList.remove('drag-over');
        });

        elements.dropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            elements.dropZone.classList.remove('drag-over');
            const file = e.dataTransfer.files[0];
            handleFileSelection(file);
        });

        elements.dropZone.addEventListener('click', () => elements.fileInput.click());
        
        elements.fileInput.addEventListener('change', (e) => {
            handleFileSelection(e.target.files[0]);
        });

        // Conversion Logic
        elements.convertBtn.addEventListener('click', processConversion);

        // Download Logic
        elements.downloadBtn.addEventListener('click', downloadFile);

        // Theme Toggle
        elements.themeToggle.addEventListener('click', toggleTheme);
    };

    /**
     * Fetch latest rates or use fallback
     */
    async function fetchExchangeRates() {
        try {
            const response = await fetch('https://open.er-api.com/v6/latest/USD');
            if (!response.ok) throw new Error('API unreachable');
            
            const data = await response.json();
            const usdToInr = data.rates.INR;
            
            // Map rates relative to INR
            // If base is USD: 1 Unit = (1 / Rate_of_Unit) * Rate_of_INR
            state.exchangeRates = {};
            for (const [curr, rate] of Object.entries(data.rates)) {
                state.exchangeRates[curr] = (1 / rate) * usdToInr;
            }
            state.statSource.textContent = "Live API";
            showToast("Exchange rates updated successfully", "success");
        } catch (error) {
            console.error("Using Fallback Rates:", error);
            state.exchangeRates = state.fallbackRates;
            state.statSource.textContent = "Fallback";
            showToast("API Offline. Using fallback rates.", "error");
        }
    }

    /**
     * Validate and Load Excel File
     */
    function handleFileSelection(file) {
        if (!file) return;

        const validTypes = [
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'application/vnd.ms-excel'
        ];

        if (!validTypes.includes(file.type) && !file.name.endsWith('.xlsx') && !file.name.endsWith('.xls')) {
            showToast("Invalid file type. Please upload an Excel file.", "error");
            return;
        }

        state.fileName = file.name;
        elements.fileNameDisplay.textContent = file.name;
        
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                state.workbook = XLSX.read(data, { type: 'array', cellStyles: true });
                elements.convertBtn.disabled = false;
                elements.downloadBtn.disabled = true;
                elements.summaryGrid.classList.remove('show');
                showToast("File loaded successfully", "success");
            } catch (err) {
                showToast("Error reading Excel file", "error");
            }
        };
        reader.readAsArrayBuffer(file);
    }

    /**
     * Core Conversion Logic
     */
    async function processConversion() {
        if (!state.workbook) return;

        const startTime = performance.now();
        elements.convertBtn.disabled = true;
        elements.progressContainer.style.display = 'block';
        updateProgress(10, "Initializing...");

        try {
            const firstSheetName = state.workbook.SheetNames[0];
            const worksheet = state.workbook.Sheets[firstSheetName];
            
            // Convert sheet to JSON for easier processing
            const rows = XLSX.utils.sheet_to_json(worksheet, { defval: null });
            
            if (rows.length === 0) throw new Error("The worksheet is empty.");

            // Identify Columns
            const headers = Object.keys(rows[0]);
            const colMap = findRequiredColumns(headers);

            if (!colMap.currency || !colMap.earnings) {
                throw new Error("Missing required columns: Platforms Currency or Earnings($)");
            }

            updateProgress(30, `Processing ${rows.length} rows...`);

            const uniqueCurrencies = new Set();
            
            // Iterate through rows and convert
            const processedData = rows.map((row, index) => {
                const currencyCode = String(row[colMap.currency] || '').toUpperCase().trim();
                const earningsAmount = parseFloat(row[colMap.earnings]) || 0;
                
                let rate = state.exchangeRates[currencyCode] || 0;
                
                // Fallback logic for specific variations if not found
                if (!rate && currencyCode.includes('USD')) rate = state.exchangeRates.USD;
                if (!rate && currencyCode.includes('EUR')) rate = state.exchangeRates.EUR;

                const inrValue = earningsAmount * rate;
                row[colMap.targetInr] = parseFloat(inrValue.toFixed(2));
                
                if (currencyCode) uniqueCurrencies.add(currencyCode);

                // Update UI every 500 rows to keep main thread alive
                if (index % 500 === 0) {
                    const pct = 30 + Math.floor((index / rows.length) * 60);
                    updateProgress(pct, `Converting row ${index}...`);
                }

                return row;
            });

            updateProgress(90, "Finalizing workbook...");

            // Create new Worksheet from processed data
            const newWs = XLSX.utils.json_to_sheet(processedData);
            
            // Create a clone of the workbook to keep original intact
            const newWb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(newWb, newWs, firstSheetName);
            
            // Copy over other sheets if they exist
            state.workbook.SheetNames.slice(1).forEach(name => {
                XLSX.utils.book_append_sheet(newWb, state.workbook.Sheets[name], name);
            });

            state.convertedWorkbook = newWb;
            
            const endTime = performance.now();
            const duration = ((endTime - startTime) / 1000).toFixed(2);

            // Update Stats
            elements.statRows.textContent = rows.length.toLocaleString();
            elements.statCurrencies.textContent = uniqueCurrencies.size;
            elements.statTime.textContent = `${duration}s`;
            
            updateProgress(100, "Conversion Complete!");
            elements.downloadBtn.disabled = false;
            elements.summaryGrid.classList.add('show');
            showToast("Conversion successful!", "success");

        } catch (err) {
            console.error(err);
            showToast(err.message, "error");
            elements.convertBtn.disabled = false;
            updateProgress(0, "Error occurred");
        }
    }

    /**
     * Map headers to expected logic
     */
    function findRequiredColumns(headers) {
        const map = {
            currency: null,
            earnings: null,
            targetInr: "Earning (₹)"
        };

        headers.forEach(h => {
            const clean = h.toLowerCase().trim();
            
            // Check Currency
            if (clean === "platforms currency" || clean === "currency" || clean === "customer territory") {
                map.currency = h;
            }
            
            // Check Earnings
            if (clean === "earnings($)" || clean === "earnings ($)" || clean === "earning($)" || clean === "amount") {
                map.earnings = h;
            }

            // Check if INR already exists
            if (clean === "earning(₹)" || clean === "earning (₹)" || clean === "inr") {
                map.targetInr = h;
            }
        });

        return map;
    }

    function downloadFile() {
        if (!state.convertedWorkbook) return;
        
        const timestamp = new Date().getTime();
        const outputName = `converted_${state.fileName.split('.')[0]}_${timestamp}.xlsx`;
        
        XLSX.writeFile(state.convertedWorkbook, outputName);
        showToast("Download started", "success");
    }

    /**
     * UI Helper: Update Progress
     */
    function updateProgress(percent, status) {
        elements.progressFill.style.width = `${percent}%`;
        elements.progressPercent.textContent = `${percent}%`;
        elements.progressStatus.textContent = status;
    }

    /**
     * UI Helper: Toast Notifications
     */
    function showToast(message, type = "success") {
        const toast = document.createElement('div');
        toast.className = `toast ${type}`;
        
        const icon = type === 'success' ? 'fa-check-circle' : 'fa-exclamation-triangle';
        
        toast.innerHTML = `
            <i class="fas ${icon}"></i>
            <span>${message}</span>
        `;
        
        elements.toastContainer.appendChild(toast);
        
        setTimeout(() => {
            toast.style.opacity = '0';
            toast.style.transform = 'translateX(100%)';
            setTimeout(() => toast.remove(), 300);
        }, 4000);
    }

    /**
     * UI Helper: Theme Toggle
     */
    function toggleTheme() {
        state.isDarkMode = !state.isDarkMode;
        const theme = state.isDarkMode ? 'dark' : 'light';
        document.documentElement.setAttribute('data-theme', theme);
        elements.themeToggle.innerHTML = state.isDarkMode ? '<i class="fas fa-sun"></i>' : '<i class="fas fa-moon"></i>';
        localStorage.setItem('theme', theme);
    }

    function loadTheme() {
        const saved = localStorage.getItem('theme') || 'light';
        state.isDarkMode = saved === 'dark';
        document.documentElement.setAttribute('data-theme', saved);
        elements.themeToggle.innerHTML = state.isDarkMode ? '<i class="fas fa-sun"></i>' : '<i class="fas fa-moon"></i>';
    }

    return { init };
})();

// Launch App
document.addEventListener('DOMContentLoaded', App.init);