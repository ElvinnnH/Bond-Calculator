# Bond Calculator

A Streamlit web application for calculating bond yields and prices with proper market conventions, day count selection, benchmark & credit spread decomposition, and Excel bond watchlist integration.

## Key Features

- Dual calculation modes:
  - Calculate Yield to Maturity (YTM) from clean price
  - Calculate clean price from target YTM
- Market-convention schedule-based pricing (fractional first period, clean/dirty price separation)
- Treasury benchmark integration (10Y / 20Y / 30Y) and credit spread decomposition stacked chart
- Day count conventions: Actual/Actual and 30/360 US
- Modified Duration (schedule-based) calculation
- Coupon frequency: Annual, Semi-annual, Quarterly
- Excel bond watchlist integration (sheet `watchlist`): select bond to pre-fill Issuer, Rating, Coupon, Maturity
- Bond List tab with rating and years-to-maturity filters
- Accrued interest and dirty price reporting
- Risk indicators and credit rating color cues

## Excel Watchlist
Place `Corporate Bonds.xlsx` in the project root with a sheet named `watchlist` including columns (or close variants):
`ISIN`, `Issuer`, `S&P Rating`, `Coupon`, `Maturity Date`
The loader attempts alternative header names and normalizes coupon values ("5" or "5%" or `0.05`).

## Installation (Local)

```bash
python -m venv .venv
source .venv/bin/activate  # macOS/Linux
# On Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## Run Locally

```bash
streamlit run bond_calculator.py
```
App opens at: http://localhost:8501

## Deployment: Streamlit Community Cloud
1. Create public GitHub repo (already: `https://github.com/ElvinnnH/Bond-Calculator.git`).
2. Ensure files are committed: `bond_calculator.py`, `requirements.txt`, `Corporate Bonds.xlsx` (if you want shared data), `README.md`.
3. Go to https://share.streamlit.io, click "Deploy app" and select the repo, branch `main` (or default), and entry file `bond_calculator.py`.
4. Deploy to obtain a URL like `https://bond-calculator.streamlit.app`.
5. Share that URL; any user can access without coding knowledge.
6. To update: push new commits; the app redeploys automatically.

### Git Quickstart
```bash
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/ElvinnnH/Bond-Calculator.git
git push -u origin main
```
Later updates:
```bash
git add .
git commit -m "Update bond calculator"
git push
```

## Usage Workflow
1. Use sidebar to select an Excel bond or manual entry.
2. Set purchase & maturity dates, coupon frequency, day count, treasury benchmark.
3. Choose calculation mode (Price → YTM or YTM → Price) in Calculator tab.
4. View stacked bar chart (Benchmark vs Credit Spread) and risk metrics.
5. Explore the full list in "Bond List" tab with filters.

## Notes
- Clean price is the user input/output; dirty price = clean + accrued.
- YTM solving uses numerical root-finding over full cashflow schedule.
- Modified duration uses discounted cashflow time weights (fractional first period).

## Disclaimer
Educational estimates only; no investment advice.
