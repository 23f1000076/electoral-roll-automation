# Electoral Roll Automation Pipeline

## Overview

The **Electoral Roll Automation Pipeline** is an end-to-end Python automation system that downloads electoral roll PDFs from the Election Commission of India (ECI) website, extracts voter information using OCR, converts the extracted data into structured CSV files, and stores the records in a MySQL database.

The project combines **Selenium**, **OCR (Tesseract)**, **Image Processing**, **PDF Parsing**, **Pandas**, and **MySQL** into a single automated workflow.

---

# Features

* Automated navigation of the ECI Electoral Roll portal
* Automatic selection of State, District, Constituency and Language
* Manual CAPTCHA support (required by the portal)
* Bulk PDF downloading
* Automatic PDF renaming and folder organization
* OCR-based extraction of voter information
* Metadata extraction from cover pages
* CSV generation
* Automatic MySQL database creation
* Duplicate voter detection
* OCR validation and review reports
* Resume-friendly folder structure

---

# Workflow

```text
Election Commission Website
        │
        ▼
Select State
        │
        ▼
Select District
        │
        ▼
Select Constituency
        │
        ▼
Select Language
        │
        ▼
User Solves CAPTCHA
        │
        ▼
Download Electoral Roll PDFs
        │
        ▼
Rename PDFs
        │
        ▼
Move PDFs into Folder Structure
        │
        ▼
Convert PDF → Images
        │
        ▼
OCR using Tesseract
        │
        ▼
Extract Metadata
        │
        ▼
Extract Voter Records
        │
        ▼
Create CSV
        │
        ▼
Create MySQL Database
        │
        ▼
Insert Records
        │
        ▼
Next Constituency
```

---

# Project Structure

```
project/
│
├── electoral_pipeline.py
├── requirements.txt
├── README.md
│
├── downloads/
│     └── State/
│           └── District/
│                 └── Constituency/
│                        ├── 001 - Part Name.pdf
│                        ├── 001 - Part Name.csv


---

# Technologies Used

| Technology    | Purpose                        |
| ------------- | ------------------------------ |
| Python        | Main programming language      |
| Selenium      | Browser automation             |
| Tesseract OCR | Text recognition               |
| Pillow        | Image preprocessing            |
| pdf2image     | Convert PDF pages into images  |
| PyPDF         | Read PDF information           |
| Pandas        | CSV creation and manipulation  |
| MySQL         | Database storage               |
| Regex         | Data extraction and validation |

---

# Prerequisites

Install the following software before running the project.

## Python

Python 3.10 or later

---

## Firefox Browser

Install Mozilla Firefox.

---

## GeckoDriver

Download the GeckoDriver matching your Firefox version.

Place it in your system PATH.

---

## Tesseract OCR

Download and install:

https://github.com/tesseract-ocr/tesseract

Example installation path:

```
C:\Program Files\Tesseract-OCR\
```

---

## Poppler

Download Poppler for Windows.

Example path:

```
C:\Program Files\poppler-26.02.0\
```

---

## MySQL Server

Install MySQL Community Server.

Create a user or use the root account.

Example configuration:

```
Host : localhost

User : root

Password : root

Database : election_db
```

---

# Python Dependencies

Install all required libraries.

```bash
pip install selenium
pip install pandas
pip install pytesseract
pip install pdf2image
pip install pillow
pip install pypdf
pip install mysql-connector-python
```

or

```bash
pip install -r requirements.txt
```

---

# Configuration

Modify the configuration section according to your system.

```python
download_folder = r"C:\Users\Admin\Downloads"

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

POPPLER_PATH = r"C:\Program Files\poppler-26.02.0\Library\bin"

STATIC_STATE_NAME = "Meghalaya"

STATIC_YEAR = "2025"

STATIC_ROLL_TYPE = "Final Roll - 2025"
```

Update MySQL credentials.

```python
DB_CONFIG = {
    "host": "localhost",
    "user": "root",
    "password": "root",
    "database": "election_db"
}
```

---

# Running the Project

Run the script.

```bash
python electoral_pipeline.py
```

The application will

* Open the Election Commission website
* Navigate automatically
* Ask you to solve the CAPTCHA
* Download PDFs
* Process OCR
* Generate CSV files
* Store records in MySQL

---

# Output Structure

```
Downloads/

└── Meghalaya

      └── East Khasi Hills

              └── Shillong

                       001 - XYZ.pdf

                       001 - XYZ.csv


---

# Database Schema

The application automatically creates the table.

Main columns include

* Voter ID
* Name
* Guardian Name
* Guardian Type
* House Number
* Age
* Gender
* Village
* District
* Mandal
* Polling Station
* Polling Address
* Assembly Constituency
* Parliamentary Constituency
* Source File

---

# Error Handling

The pipeline includes handling for

* OCR failures
* Invalid voter cards
* Duplicate voter IDs
* Missing PDFs
* Locked files
* Website alerts
* Selenium timeouts
* Missing metadata
* Invalid ages
* Download failures
* Missing pages

Review reports are automatically generated whenever necessary.

---

# Advantages

* Fully automated workflow
* Eliminates manual data entry
* Organized folder structure
* Automatic database creation
* Duplicate prevention
* Review reports for OCR errors
* Easy to extend
* Modular architecture
* Reusable OCR pipeline
* Suitable for large-scale data extraction

---

# Limitations

* CAPTCHA must be solved manually.
* OCR accuracy depends on scan quality.
* Changes to the ECI website may require Selenium updates.
* Hardcoded paths should be replaced with environment variables.
* Processing large datasets is CPU-intensive.

---

# Future Improvements

* Docker support
* Linux compatibility
* Multi-threaded OCR
* Automatic retry mechanism
* Environment variable configuration
* Logging framework
* Cloud storage support
* Dashboard for monitoring
* REST API integration
* AI-based OCR models (EasyOCR, PaddleOCR, TrOCR)

---

# Troubleshooting

### PDFs are not downloading

* Verify GeckoDriver installation.
* Ensure Firefox allows automatic downloads.
* Solve the CAPTCHA correctly.

### OCR returns incorrect text

* Verify the Tesseract installation path.
* Increase image DPI if necessary.
* Ensure the PDF quality is readable.

### MySQL connection failed

* Verify MySQL Server is running.
* Check host, username, password, and database configuration.

### Poppler not found

Verify the Poppler installation path in the configuration.

---

# Contributing

Contributions are welcome.

If you find bugs or have feature suggestions, feel free to create an issue or submit a pull request.

---

# Disclaimer

This project is intended for educational, research, and data engineering purposes.

Please ensure that your use of the Election Commission of India website complies with all applicable laws, website terms of use, and data privacy regulations. CAPTCHA handling remains manual to respect the portal's anti-automation safeguards.

---

# License

This project is released under the MIT License. Feel free to use, modify, and distribute it according to the license terms.
