# SGD Exchange Rate Analyser

A Python project that downloads 3 years of SGD exchange rate data, analyses trends, and generates a formatted Excel report with embedded charts.

Built by **Thuta Lin** as a portfolio project to demonstrate data analysis skills using Python and Excel.

---

## What It Does

- **Downloads** live SGD exchange rate data vs 6 major currencies (USD, EUR, GBP, JPY, AUD, CNY) using the Yahoo Finance API
- **Analyses** the data: 3-year statistics, volatility, year-over-year change, and monthly averages
- **Visualises** trends with two charts (30-day rolling average line chart + YoY bar chart)
- **Exports** a fully formatted multi-sheet Excel workbook with embedded charts

---

## Output

Running the script produces a single Excel file in the `output/` folder with 4 sheets:

| Sheet | Contents |
|---|---|
| Summary | Key stats per currency (mean, high, low, volatility, YoY change) |
| Historical Data | Full daily exchange rate table for 3 years |
| Monthly Averages | Rates averaged by month (easier to spot trends) |
| Charts | Two charts embedded as images |

---

## How to Run

**1. Clone the repository**
```bash
git clone https://github.com/YOUR_USERNAME/sgd-fx-tracker.git
cd sgd-fx-tracker
```

**2. Install dependencies**
```bash
pip install -r requirements.txt
```

**3. Run the script**
```bash
python main.py
```

The Excel report will be saved to the `output/` folder automatically.

---

## Project Structure

```
sgd-fx-tracker/
├── main.py            
├── requirements.txt   
├── README.md         
└── output/            
```

---

## Tools & Libraries Used

| Tool | Purpose |
|---|---|
| `yfinance` | Fetches exchange rate data from Yahoo Finance |
| `pandas` | Data cleaning, analysis, and aggregation |
| `matplotlib` | Chart creation |
| `openpyxl` | Excel report generation and formatting |

---

## Currency Pairs Tracked

| Pair | Meaning |
|---|---|
| SGD/USD | Singapore Dollar vs US Dollar |
| SGD/EUR | Singapore Dollar vs Euro |
| SGD/GBP | Singapore Dollar vs British Pound |
| SGD/JPY | Singapore Dollar vs Japanese Yen |
| SGD/AUD | Singapore Dollar vs Australian Dollar |
| SGD/CNY | Singapore Dollar vs Chinese Yuan |

---

## About

This project was built as part of a data portfolio to demonstrate:
- Python scripting and API data retrieval
- Financial data analysis using pandas
- Data visualisation with matplotlib
- Automated Excel reporting with openpyxl

---

## Acknowledgement

This project was built with the assistance of [Claude](https://claude.ai) (Anthropic's AI assistant), which helped with code structure, chart design, and setting up the GitHub Actions automation. The analysis logic, financial context, and project direction were guided by me as part of learning how to apply Python to real-world data problems.
