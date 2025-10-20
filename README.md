# Digital Receipt Automation

An automated data extraction and organization tool for government-issued digital receipts.  
This system combines **Selenium**, **HTTPX**, and **async I/O** to log in, capture authorization tokens, retrieve large batches of receipt data, and export them into Excel files for analysis or record-keeping.

---

## 🚀 Overview

This project automates the process of retrieving and structuring e-receipt data from the official Vietnamese government e-invoice portal (`hoadondientu.gdt.gov.vn`).  
It handles browser authentication, API token extraction, and high-speed asynchronous data collection, turning previously manual tasks into fully automated workflows.

The app includes:
- Automated login with **Selenium** (Chrome in debug mode)
- Secure **authorization token capture** via Chrome DevTools WebSocket
- Parallelized **data retrieval** using **asyncio** + **HTTPX**
- Export of organized results into **Excel (via Pandas)**
- Simple **Tkinter GUI** to select company, date range, and receipt type

---

## 🧩 Features

- 🔐 Automated browser login and token management  
- ⚡ Parallel async requests for up to **10,000 receipts/hour**  
- 🧾 Support for both **Sales (Bán ra)** and **Purchases (Mua vào)** receipts  
- 📊 Structured export of financial data to Excel for reporting and audits  
- 🪟 GUI interface for non-technical users (Tkinter executable)  

---

## 🧠 Technical Stack

- **Python 3.11+**
- **Selenium** — automated Chrome login and DOM interaction  
- **HTTPX** — async API communication  
- **Asyncio** — concurrent I/O handling  
- **WebSocket / DevTools Protocol** — authorization token capture  
- **Pandas** — data cleaning and export  
- **Tkinter** — graphical interface for running the tool  

---

## ⚙️ How It Works

1. Launches Chrome in remote-debug mode and connects via Selenium.  
2. Intercepts authorization headers through Chrome’s WebSocket API (`getAuth.py`).  
3. Uses the valid token to call government APIs asynchronously.  
4. Retrieves invoice details, including company info, tax codes, itemized costs, and payment methods.  
5. Saves results into structured Excel sheets for analysis.

---

## 🧰 File Structure

| File | Description |
|------|--------------|
| `main.py` | Main application logic: login, data retrieval, and export |
| `getAuth.py` | Handles WebSocket connection to extract authorization token |
| `path.json` | Local configuration for driver and executable paths |
| `users/` | Stores user credentials (pickled dictionaries) |
| `output/` | Folder where exported Excel files are saved |

---

## 🪄 Usage

1. Ensure you have **Google Chrome** and **ChromeDriver** installed.  
2. Update `path.json` with the correct paths:
   ```json
   {
     "exe_path": "C:/path/to/exe",
     "driver_path": "C:/path/to/chromedriver.exe",
     "batch_script": "C:/path/to/open_chrome_debug.bat"
   }
