# Web Scraping & Excel Automation

An end-to-end Python automation project that extracts structured product data from a website and transforms it into a professional Excel report with analytics and visual dashboards.

## Project Overview

This project demonstrates a complete workflow for turning web data into a client-ready Excel deliverable:

**Website → Browser Automation → Data Extraction → Excel Automation → Analytics Dashboard**

The scraper uses Playwright to collect product information from [Books to Scrape](http://books.toscrape.com/) and OpenPyXL to generate and format the final Excel report.

> **Note:** Books to Scrape is used as a demo scraping target for portfolio purposes.

## What This Project Does

The scraper:

1. Opens the target website using Playwright.
2. Collects products from the listing page.
3. Visits each product's detail page.
4. Extracts structured product information.
5. Converts the collected information into an organized Excel workbook.
6. Applies professional formatting and conditional formatting.
7. Generates summary statistics and charts.
8. Produces a ready-to-use Excel report.

## Key Features

### Web Scraping
- Browser automation with Playwright
- Product listing extraction
- Deep scraping of individual product detail pages
- Category extraction
- UPC extraction
- Product type extraction
- Price extraction
- Tax extraction
- Rating extraction
- Review count extraction
- Stock status and quantity extraction
- Product description extraction

### Excel Automation
- Automated workbook generation
- Structured product catalog
- Professional headers and formatting
- Zebra row styling
- Auto-filter
- Freeze panes
- Conditional stock formatting
- Automatic column sizing
- Rating visualization

### Analytics Dashboard
The generated workbook includes:
- Total products scraped
- Total available stock
- Average price
- Lowest price
- Highest price
- Product distribution by category
- Rating distribution
- Top 10 products with the lowest stock
- Bar charts
- Pie chart

## Technology Stack

| Technology | Purpose |
|---|---|
| Python | Core programming language |
| Playwright | Browser automation and web scraping |
| OpenPyXL | Excel workbook generation and formatting |
| Regex | Structured text extraction |
| Asyncio | Asynchronous browser operations |

## Project Workflow

```text
                    TARGET WEBSITE
                          │
                          ▼
                 Playwright Browser
                          │
                          ▼
                 Product Listing
                          │
                          ▼
              Product Detail Pages
                          │
                          ▼
                Data Extraction
                          │
                          ▼
                Structured Dataset
                          │
                          ▼
                 Excel Automation
                    ┌─────┴─────┐
                    ▼           ▼
              Data Catalog   Dashboard
                    │           │
                    └─────┬─────┘
                          ▼
                 Client-Ready Report
```

## Project Results

The current demo run successfully collected **20 products** and generated an Excel workbook containing:

- **Detailed Book Catalog**
- **Dashboard & Statistics**
- 13 structured data fields
- Category analysis
- Rating analysis
- Stock analysis
- Automated charts

## Excel Output

The generated workbook is available in:

```text
output/daftar_buku_premium_lengkap.xlsx
```

### Detailed Catalog

![Detailed Book Catalog](screenshots/detailed-catalog.png)

### Analytics Dashboard

![Analytics Dashboard](screenshots/analytics-dashboard.png)

## Example Data Fields

The detailed catalog contains fields such as:

```text
Book Title
Category
UPC
Product Type
Price (Excl. Tax)
Price (Incl. Tax)
Tax
Rating
Review Count
Stock Status
Stock Quantity
Short Description
```

## Installation

### 1. Clone the repository

```bash
git clone https://github.com/YOUR-USERNAME/web-scraping-excel-automation.git
cd web-scraping-excel-automation
```

### 2. Install Python dependencies

```bash
pip install -r requirements.txt
```

### 3. Install Playwright Chromium

```bash
playwright install chromium
```

## Usage

Run the scraper with:

```bash
python scraper.py
```

The script will open a Chromium browser, scrape the target website, process the collected data, and generate the Excel report.

## Project Structure

```text
web-scraping-excel-automation/
│
├── scraper.py
├── requirements.txt
├── README.md
├── .gitignore
│
├── output/
│   └── daftar_buku_premium_lengkap.xlsx
│
├── screenshots/
│   ├── detailed-catalog.png
│   └── analytics-dashboard.png
│
└── docs/
```

## Business Use Cases

The same type of workflow can be adapted for business data collection and reporting, including:

- E-commerce product data collection
- Competitor price research
- Product catalog creation
- Inventory monitoring
- Market research
- Website-to-Excel data automation
- Structured data collection
- Automated reporting

## Why This Project Matters

The goal is not only to scrape a website.

The project demonstrates how raw web information can be transformed into a structured, readable, and useful business report.

**Raw Web Data → Structured Data → Excel Report → Business Insights**

## Limitations

This portfolio project uses a demonstration website and a fixed scraping structure. Different websites require different selectors, pagination handling, authentication flows, anti-bot considerations, and data-cleaning rules.

The project is intended as a demonstration of web scraping and Excel automation capabilities.

## Future Improvements

Potential improvements include:

- Pagination support
- Configurable target URLs
- Retry handling
- Logging
- Data validation
- Duplicate detection
- Configurable output formats
- More advanced analytics

## Author

**Theofilus**

Python • Web Scraping • Excel Automation • Data Processing

---

If you need automated web data collection and conversion into structured Excel reports, this project demonstrates the workflow and technical capabilities behind that service.
