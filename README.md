# 💻 Laptop Data Scraper | Scrapy + Playwright + Excel Automation

### 🧠 Overview
This project automates the extraction of **laptop product data** from a structured e-commerce test site using **Scrapy** and **Playwright**. It demonstrates a production-grade, end-to-end **web scraping pipeline** — from dynamic content rendering to clean, client-ready Excel reporting.

---

## 🚀 Features
✅ Handles **dynamic pages** using Playwright (for JavaScript-rendered content)  
✅ Maintains **exact sequential ordering** of products across pagination  
✅ Performs **data cleaning, deduplication, and validation**  
✅ Exports results to a **professionally formatted Excel file** using OpenPyXL  
✅ Adds **metadata** including data source and timestamp  
✅ Fully **automated, reproducible, and efficient** — ideal for portfolio demonstration or client projects

---

## 🛠️ Tech Stack

| Component | Purpose |
|------------|----------|
| **Python** | Core language |
| **Scrapy** | Web scraping framework |
| **Playwright** | Handles dynamic (JavaScript-rendered) pages |
| **OpenPyXL** | Excel file creation and formatting |
| **Datetime** | Timestamp and metadata handling |

---

## 📊 Output Example

**File Name:** `Product_Details.xlsx`  
**Sheet:** `Products`

| Title | Price | Description | Reviews |
|--------|--------|--------------|----------|
| Laptop Model A | $499.00 | Lightweight and portable | 120 |
| Laptop Model B | $899.00 | High performance for gaming | 235 |

### ✨ Excel Styling Includes:
- 🎨 Dark blue header with white bold centered text  
- 🧾 Alternate light-grey row shading  
- 💲 Numeric formatting for prices and reviews  
- 📌 Frozen header row and auto-filter enabled  
- 🕓 Metadata (source URL + timestamp) appended at the end  

---

## 📂 Project Structure
```
laptops_spider.py          # Main Scrapy + Playwright spider
Product_Details.xlsx       # Generated Excel output
README.md                  # Documentation file
```

---

## ⚙️ Installation & Setup

### 1. Clone the Repository
```bash
git clone https://github.com/Onyeksman/laptop-scraper.git
cd laptop-scraper
```

### 2. Install Dependencies
```bash
pip install scrapy scrapy-playwright openpyxl
python -m playwright install
```

### 3. Run the Spider
```bash
python laptops_spider.py
```

> The spider automatically launches Playwright, scrapes the site, and saves the cleaned dataset to `Product_Details.xlsx`.

---

## 🧩 Key Highlights
- Ensures **data accuracy** and **consistency** across all pages  
- Removes **duplicates** and replaces missing values with `"N/A"`  
- Applies **Excel formatting automation** for a polished, professional dataset  
- Ready for **market research**, **competitive analysis**, or **inventory tracking**

---

## 👤 Author
**Onyekachi Ejimofor**  
💼 *Lead Python Developer — Web Scraping & Data Automation*  
📧 onyeife@gmail.com  
🌐 https://github.com/Onyeksman

---

## 📜 License
This project is released under the [MIT License](LICENSE).
