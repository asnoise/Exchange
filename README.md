# Excel Currency Converter (Col Q to R)

A professional, high-speed tool to convert Excel earnings into Indian Rupees (INR). This version is optimized to read **Column Q** for source currency and output the result into **Column R**.

## ✨ Key Features
- **Strict Column Mapping**: Specifically designed for reports where **Column Q** contains the currency code (USD, EUR, GBP) and **Column R** is reserved for the INR calculation.
- **Smart Detection**: Automatically finds the "Earnings" value column by scanning headers for keywords like "Earnings($)" or "Amount".
- **Real-Time Data**: Connects to the ExchangeRate-API for live mid-market rates.
- **Privacy First**: No server-side processing. Your financial data stays in your browser.
- **Dark Mode**: Optimized UI for both light and dark environments.

## 🛠 Tech Stack
- **JavaScript (ES6+)**: Pure vanilla logic for maximum performance.
- **SheetJS**: Advanced Excel manipulation without formatting loss.
- **CSS3**: Modern glassmorphism with responsive layout.

## 🚀 How to Deploy on GitHub Pages
1. Create a new repository on GitHub.
2. Upload `index.html`, `style.css`, and `script.js`.
3. Go to **Settings > Pages**.
4. Select the `main` branch and click **Save**.
5. Your application will be live at `https://yourusername.github.io/your-repo-name/`.

## 📖 Usage Guide
1. **Prepare File**: Ensure your Excel file has currencies in Column Q.
2. **Upload**: Drag your file into the dashboard.
3. **Convert**: Click 'Start Conversion'. The app will fetch rates and calculate values.
4. **Download**: Get your file back with the "Earning (₹)" column fully populated in Column R.

## 📝 Supported Currencies
Supports all major world currencies via API fallback, including:
- USD, EUR, GBP, AED, CAD, AUD, JPY, SGD.

## ⚖️ License
MIT License - Free for personal and commercial use.
