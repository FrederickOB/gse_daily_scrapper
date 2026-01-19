# GSE Trading Data Scraper

## Overview
Automated web scraper that extracts yesterday's trading data from the Ghana Stock Exchange (GSE) website and saves it to Excel for Power BI dashboard integration.

## Features
- Scrapes data directly from HTML table cells (not Excel downloads)
- Filters for previous day's data automatically
- Appends new data to existing Excel file
- Prevents duplicate entries
- Ready for task scheduler automation

## Requirements
```
pandas
selenium
webdriver-manager
```

## Setup
1. Install dependencies: `pip install pandas selenium webdriver-manager`
2. Ensure Chrome browser is installed
3. Create data directory: `/Users/mac/Desktop/data/`

## Usage
Run manually: `python main.py`

## Automation
Set up Windows Task Scheduler to run daily at desired time:
- Program: `python`
- Arguments: `main.py`
- Start in: `/Users/mac/Desktop/GSE script/`

## Output
Data saved to: `/Users/mac/Desktop/data/gse_trading_data_latest.xlsx`

Refresh Power BI dashboard to view latest metrics.