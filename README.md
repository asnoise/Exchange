# Excel Currency Converter (Global to INR)

A high-performance, client-side web application that converts earnings in an Excel file from various global currencies (USD, EUR, GBP, etc.) into Indian Rupees (INR) using live exchange rates.

![License](https://img.shields.io/badge/license-MIT-blue.svg)
![Tech](https://img.shields.io/badge/tech-Vanilla%20JS-yellow.svg)
![Security](https://img.shields.io/badge/privacy-Client--Side-green.svg)

## 🚀 Features

- **Automatic Column Detection**: Smartly identifies columns like "Platforms Currency" and "Earnings ($)".
- **Live Exchange Rates**: Fetches real-time data from `open.er-api.com`.
- **Privacy Focused**: All processing happens in your browser. No data is uploaded to any server.
- **Glassmorphic UI**: Modern, responsive design with Dark Mode support.
- **Large File Support**: Optimized to handle workbooks with 10,000+ rows using efficient memory management.
- **No Backend**: Deployable on GitHub Pages, Vercel, or Netlify with zero configuration.

## 🛠️ Supported Currencies

- **USD** (US Dollar)
- **EUR** (Euro)
- **GBP** (British Pound)
- **JPY** (Japanese Yen)
- **CAD** (Canadian Dollar)
- **AUD** (Australian Dollar)
- **SGD** (Singapore Dollar)
- **AED** (UAE Dirham)
- *And many more via live API*

## 📦 Installation & Deployment

### Local Setup
1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/excel-currency-converter.git