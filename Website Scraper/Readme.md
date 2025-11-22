Creator - NAMA Krityam

# 🧠 Adaptive Universal Web Scraper (GUI-Based)

[![Python](https://img.shields.io/badge/Python-3.10%2B-blue.svg)]()
[![Playwright](https://img.shields.io/badge/Automation-Playwright-green.svg)]()
[![GUI](https://img.shields.io/badge/Desktop-Tkinter-orange.svg)]()
[![Status](https://img.shields.io/badge/Build-Stable-success.svg)]()

A GUI-driven, adaptive web scraping tool developed during an **SDE internship project**.

This application automatically detects tables, lists, pagination, and infinite-scroll structures — extracting structured data without modifying code for each site.  
It supports image validation, partial saving, progress tracking, and safe interruption handling.

---

## 📌 Key Features

✔ GUI interface — no coding required to run  
✔ Works on dynamic, paginated, and infinite-scroll pages  
✔ Auto-detects:
- tables
- list-based content
- buttons for pagination
- scrollable containers

✔ Export to Excel (multi-sheet supported)  
✔ Optional image download + status validation  
✔ Safe STOP button with partial save  
✔ No duplicate data (normalization logic)  
✔ Runs headless or visible mode  
✔ Threaded execution — UI stays responsive  

---

## 🛠 Tech Stack

- **Python 3.10+**
- Playwright (browser automation)
- Tkinter (GUI)
- pandas + openpyxl (Excel export)
- httpx (image validation & downloading)
- asyncio + threading
- logging

---

## 📦 Installation

1. Install Python libraries

Run:

- pip install --upgrade pip
- pip install playwright pandas httpx openpyxl

2. Install Playwright browsers
 
- playwright install chromium (Optional if you want full support)

- playwright install

3. (Windows users) If tkinter is missing
Most Python distributions include it.
If not, install Python from:

https://www.python.org/downloads/

and ensure "tcl/tk" is enabled.

▶️ How to Run the Application
In the project folder:

- Run in terminal
python gui_scraper.py
(Use the filename that contains the GUI code — for example gui_scraper.py)

The window will open with:

✅ URL input
✅ Excel filename input
✅ Checkboxes & options
✅ Log viewer
✅ Start / Stop buttons
  


🧭 Usage Instructions:-

1. Enter the Website URL
Example:
https://www.example.com/items 

2. Enter Output Excel Filename
Example:
mydata.xlsx
(Extension auto-adds if missing)

3. Optional Settings
- download first image per row
- headless mode
- max pagination attempts
- image download concurrency

4. Click Start Scraping
The scraper will:

- open browser
- detect content type
- extract data progressively
- update logs in real time
- save results to Excel

5. Click STOP anytime
✔ scraping halts safely
✔ partial data is saved automatically

📂 Output Example

project/
│── scraper_gui.py
│── extracted_data.xlsx
│── scraped_images/
│── scraper_log_2025_11_23.log
│── README.md
Excel may contain:

- one sheet (simple pages)
- multiple sheets (multiple tables/lists)
- Columns vary based on page detection but always include:
- serial_no
- extracted fields
- links and images (if found)
- validation columns (when enabled)

📝 Logging
# The application logs:

- navigation events
- extraction progress
- container/table detection
- scroll cycles
- image validation status
- STOP requests
- save confirmations

# Logs appear:

✅ live in the GUI
✅ saved to file automatically

Example:

scraper_log_20251123_002355.log
🔁 Interruption & Recovery
✔ STOP button triggers safe exit
✔ Partial results are written to Excel
✔ Browser closes cleanly
✔ No data corruption
✔ UI remains responsive

⚠️ Notes & Considerations:-

- dynamic websites may require extra load time
- content structure changes may need selector updates
- image availability depends on the source
- use only on allowed and permitted websites

📄 License :-

- This tool is provided strictly for educational and testing purposes.
- Use responsibly and comply with all legal and ethical guidelines.