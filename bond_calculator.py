import streamlit as st
import numpy as np
import pandas as pd
from datetime import datetime, date, timedelta
from scipy.optimize import fsolve
import math
import altair as alt  # Added for charting
from dateutil.relativedelta import relativedelta  # For coupon schedule

# Attempt to load bond list from Excel with robust column mapping (updated to use 'watchlist' sheet)
try:
    # NEW: Load specific sheet 'watchlist' (case-insensitive)
    xls = pd.ExcelFile("Corporate Bonds.xlsx")
    target_sheet = None
    for s in xls.sheet_names:
        if s.lower() == "watchlist":
            target_sheet = s
            break
    if target_sheet is None:
        st.sidebar.warning("Sheet 'watchlist' not found in Corporate Bonds.xlsx.")
        raw_df = None
    else:
        raw_df = xls.parse(target_sheet)

    if raw_df is not None:
        raw_df.columns = [c.strip() for c in raw_df.columns]
        lower_map = {c.lower(): c for c in raw_df.columns}
        required_lower = ["isin", "issuer", "s&p rating", "coupon", "maturity date"]
        alt_names = {
            "s&p rating": ["s&p", "rating", "sp rating", "s&p rating"],
            "maturity date": ["maturity", "maturitydate", "mat date", "maturity date"],
            "coupon": ["coupon", "coupon rate", "cpn"],
        }
        resolved = {}
        for key in required_lower:
            if key in lower_map:
                resolved[key] = lower_map[key]
            else:
                found = None
                for alt in alt_names.get(key, []):
                    if alt in lower_map:
                        found = lower_map[alt]
                        break
                if found:
                    resolved[key] = found
        if len(resolved) == len(required_lower):
            bonds_df = raw_df[[resolved[k] for k in required_lower]].copy()
            bonds_df.columns = ["ISIN", "Issuer", "S&P Rating", "Coupon", "Maturity Date"]
            # Clean coupon column (remove % and convert to float)
            def _parse_coupon(x):
                if pd.isna(x):
                    return None
                if isinstance(x, (int, float)):
                    return float(x)
                s = str(x).strip()
                if s.endswith('%'):
                    s = s[:-1]
                s = s.replace(',', '')
                try:
                    return float(s)
                except Exception:
                    return None
            bonds_df['Coupon'] = bonds_df['Coupon'].apply(_parse_coupon)
            # Normalize coupon to percent for display (if fractional like 0.05 -> 5.0)
            bonds_df['CouponDisplayPct'] = bonds_df['Coupon'].apply(lambda v: v*100 if v is not None and v <= 1 else v)
            bonds_df['Maturity Date'] = pd.to_datetime(bonds_df['Maturity Date'], errors='coerce').dt.date
        else:
            bonds_df = None
            missing = set([c.title() for c in required_lower]) - set([c.lower() for c in resolved.keys()])
            st.sidebar.warning(f"Excel sheet 'watchlist' loaded but missing required columns: {', '.join(missing)}")
    else:
        bonds_df = None
except Exception:
    bonds_df = None
    st.sidebar.info("Could not load Corporate Bonds.xlsx or sheet 'watchlist'.")

# Set page config
st.set_page_config(
    page_title="Bond Calculator",
    page_icon="📊",
    layout="wide"
)

# Title and description
st.title("📊 Bond Calculator")
st.markdown("Calculate Yield to Maturity (YTM) from price or Price from YTM for bonds. Includes Treasury benchmark and credit spread decomposition.")

def days_between_dates(start_date, end_date):
    """Calculate the number of days between two dates"""
    return (end_date - start_date).days

def generate_coupon_schedule(purchase_date: date, maturity_date: date, coupon_frequency: int):
    """Generate ascending list of coupon dates from first after purchase to maturity (inclusive). Assume maturity is coupon date."""
    months_step = int(12 / coupon_frequency)
    dates = []
    current = maturity_date
    while current > purchase_date and len(dates) < 1000:  # safety cap
        dates.append(current)
        current = current - relativedelta(months=months_step)
    dates.sort()
    return dates

def find_adjacent_coupon_dates(purchase_date: date, maturity_date: date, coupon_frequency: int):
    """Forward schedule from purchase to maturity; return prev and next coupon dates."""
    months_step = int(12 / coupon_frequency)
    # Build forward schedule starting from maturity backwards then reverse
    dates = []
    current = maturity_date
    while current > purchase_date and len(dates) < 1000:
        dates.append(current)
        current -= relativedelta(months=months_step)
    dates.sort()  # ascending after purchase
    if not dates:
        return None, None, []
    next_coupon = dates[0]
    prev_coupon = next_coupon - relativedelta(months=months_step)
    return prev_coupon, next_coupon, dates

# Update: include freq param and fix 30/360 denominator

def day_count_fraction(prev_coupon: date, settlement: date, next_coupon: date, convention: str, freq: int):
    if convention == "Actual/Actual":
        num = (settlement - prev_coupon).days
        den = (next_coupon - prev_coupon).days
        return num / den if den > 0 else 0.0
    elif convention == "30/360 US":
        # US 30/360 day count within a coupon period
        def _dmy(d):
            return d.year, d.month, d.day
        y1, m1, d1 = _dmy(prev_coupon)
        y2, m2, d2 = _dmy(settlement)
        if d1 == 31:
            d1 = 30
        if d2 == 31 and d1 == 30:
            d2 = 30
        num = 360 * (y2 - y1) + 30 * (m2 - m1) + (d2 - d1)
        den = 360 / freq
        return (num / den) if den > 0 else 0.0
    else:
        return 0.0

# Helper: 30/360 US day difference between two dates

def days_30_360_us(start_d: date, end_d: date) -> int:
    y1, m1, d1 = start_d.year, start_d.month, start_d.day
    y2, m2, d2 = end_d.year, end_d.month, end_d.day
    if d1 == 31:
        d1 = 30
    if d2 == 31 and d1 == 30:
        d2 = 30
    return 360 * (y2 - y1) + 30 * (m2 - m1) + (d2 - d1)

def compute_coupon_context(purchase_date: date, maturity_date: date, coupon_frequency: int, convention: str):
    prev_coupon, next_coupon, schedule = find_adjacent_coupon_dates(purchase_date, maturity_date, coupon_frequency)
    if prev_coupon is None:
        return 0, 0, 0, 0.0
    if convention == "30/360 US":
        days_in_period = int(360 / coupon_frequency)
        days_since_prev = days_30_360_us(prev_coupon, purchase_date)
        days_to_next = max(days_in_period - days_since_prev, 0)
    else:  # Actual/Actual
        days_in_period = (next_coupon - prev_coupon).days
        days_since_prev = (purchase_date - prev_coupon).days
        days_to_next = (next_coupon - purchase_date).days
    fraction_elapsed = (days_since_prev / days_in_period) if days_in_period > 0 else 0.0
    remaining_coupons = len(schedule)
    return remaining_coupons, int(days_to_next), int(days_in_period), float(fraction_elapsed)

def calculate_bond_price(face_value, coupon_rate, ytm, remaining_coupons, days_to_next_coupon, days_in_period, fraction_elapsed, coupon_frequency=2):
    """Pricer with proper fractional first period. remaining_coupons includes next coupon."""
    coupon_payment = (coupon_rate * face_value) / coupon_frequency
    accrued_interest = coupon_payment * fraction_elapsed
    if ytm == 0:
        clean_price = face_value + coupon_payment * remaining_coupons
        dirty_price = clean_price + accrued_interest
        return clean_price, dirty_price, accrued_interest
    first_fraction = days_to_next_coupon / days_in_period if days_in_period > 0 else 0
    pv_coupons = 0.0
    for i in range(remaining_coupons):  # i=0 is next coupon
        t = first_fraction + i  # time in coupon periods
        discount = (1 + ytm / coupon_frequency) ** t
        pv_coupons += coupon_payment / discount
    t_face = first_fraction + (remaining_coupons - 1)
    pv_face_value = face_value / ((1 + ytm / coupon_frequency) ** t_face)
    clean_price = pv_coupons + pv_face_value
    dirty_price = clean_price + accrued_interest
    return clean_price, dirty_price, accrued_interest

def calculate_ytm_from_price(face_value, coupon_rate, market_dirty_price, remaining_coupons, days_to_next_coupon, days_in_period, fraction_elapsed, coupon_frequency=2):
    coupon_payment = (coupon_rate * face_value) / coupon_frequency
    accrued_interest = coupon_payment * fraction_elapsed
    def price_diff(ytm):
        _, dirty_p, _ = calculate_bond_price(face_value, coupon_rate, ytm, remaining_coupons, days_to_next_coupon, days_in_period, fraction_elapsed, coupon_frequency)
        return dirty_p - market_dirty_price
    try:
        ytm = fsolve(price_diff, coupon_rate)[0]
        return max(ytm, 0.0)
    except Exception:
        return None

# Full schedule-based pricing (no par shortcut)

def build_coupon_schedule(settlement: date, maturity: date, freq: int):
    months_step = int(12 / freq)
    dates = []
    current = maturity
    while current > settlement and len(dates) < 1000:
        dates.append(current)
        current -= relativedelta(months=months_step)
    dates.sort()
    return dates

def accrual_context(settlement: date, freq: int, next_coupon: date, convention: str):
    months_step = int(12 / freq)
    prev_coupon = next_coupon - relativedelta(months=months_step)
    if convention == "30/360 US":
        days_in_period = int(360 / freq)
        accrued_days = days_30_360_us(prev_coupon, settlement)
        days_to_next = max(days_in_period - accrued_days, 0)
    else:  # Actual/Actual
        days_in_period = (next_coupon - prev_coupon).days
        accrued_days = (settlement - prev_coupon).days
        days_to_next = (next_coupon - settlement).days
    fraction_elapsed = (accrued_days / days_in_period) if days_in_period > 0 else 0.0
    return fraction_elapsed, days_in_period, days_to_next

def price_dirty(face_value, coupon_rate, ytm, freq, schedule, settlement, fraction_elapsed, days_in_period, days_to_next):
    coupon = (coupon_rate * face_value) / freq
    # Fraction until first coupon (remaining part of current period)
    first_fraction = days_to_next / days_in_period if days_in_period > 0 else 0.0
    periods = len(schedule)
    dirty = 0.0
    for i in range(periods):
        exponent = first_fraction + i  # fractional first + integer steps
        cf = coupon if i < periods - 1 else coupon + face_value
        dirty += cf / ((1 + ytm / freq) ** exponent)
    return dirty

def solve_ytm_schedule(face_value, coupon_rate, clean_price_obs, freq, settlement, maturity, convention):
    schedule = build_coupon_schedule(settlement, maturity, freq)
    if not schedule:
        return None
    next_coupon = schedule[0]
    fraction_elapsed, days_in_period, days_to_next = accrual_context(settlement, freq, next_coupon, convention)
    coupon = (coupon_rate * face_value) / freq
    accrued = coupon * fraction_elapsed
    dirty_obs = clean_price_obs + accrued
    def diff(y):
        return price_dirty(face_value, coupon_rate, y, freq, schedule, settlement, fraction_elapsed, days_in_period, days_to_next) - dirty_obs
    try:
        ytm = fsolve(diff, coupon_rate)[0]
        return max(ytm, 0.0)
    except Exception:
        return None

def price_from_ytm_schedule(face_value, coupon_rate, ytm, freq, settlement, maturity, convention):
    schedule = build_coupon_schedule(settlement, maturity, freq)
    if not schedule:
        return None, None, None
    next_coupon = schedule[0]
    fraction_elapsed, days_in_period, days_to_next = accrual_context(settlement, freq, next_coupon, convention)
    coupon = (coupon_rate * face_value) / freq
    accrued = coupon * fraction_elapsed
    dirty = price_dirty(face_value, coupon_rate, ytm, freq, schedule, settlement, fraction_elapsed, days_in_period, days_to_next)
    clean = dirty - accrued
    return clean, dirty, accrued

# NEW: Modified Duration calculation (Macaulay / (1 + ytm/freq))

def modified_duration(face_value, coupon_rate, ytm, freq, settlement, maturity, convention):
    """Compute modified duration using full schedule with fractional first period based on selected day count."""
    if ytm is None:
        return None
    schedule = build_coupon_schedule(settlement, maturity, freq)
    if not schedule:
        return None
    next_coupon = schedule[0]
    fraction_elapsed, days_in_period, days_to_next = accrual_context(settlement, freq, next_coupon, convention)
    first_fraction = (days_to_next / days_in_period) if days_in_period > 0 else 0.0
    coupon = (coupon_rate * face_value) / freq
    periods = len(schedule)
    dirty_price = price_dirty(face_value, coupon_rate, ytm, freq, schedule, settlement, fraction_elapsed, days_in_period, days_to_next)
    if dirty_price == 0:
        return None
    macaulay_num = 0.0
    for i in range(periods):
        t_periods = first_fraction + i  # time in coupon periods to cash flow
        cf = coupon if i < periods - 1 else coupon + face_value
        discount = (1 + ytm / freq) ** t_periods
        pv = cf / discount
        t_years = t_periods / freq  # convert coupon periods (including fractional first) to years
        macaulay_num += t_years * pv
    macaulay = macaulay_num / dirty_price
    mod_duration = macaulay / (1 + ytm / freq)
    return mod_duration

# Helper pricing functions (placed early to avoid NameError)

def market_price(face_value, coupon_rate, ytm, remaining_coupons, fraction_elapsed, coupon_frequency):
    coupon = (coupon_rate * face_value) / coupon_frequency
    accrued = coupon * fraction_elapsed
    pv_coupons = sum(coupon / ((1 + ytm / coupon_frequency) ** i) for i in range(1, remaining_coupons + 1))
    pv_face = face_value / ((1 + ytm / coupon_frequency) ** remaining_coupons)
    dirty = pv_coupons + pv_face
    clean = dirty - accrued
    return clean, dirty, accrued

def solve_ytm(face_value, coupon_rate, market_clean_price, fraction_elapsed, remaining_coupons, coupon_frequency):
    coupon = (coupon_rate * face_value) / coupon_frequency
    accrued = coupon * fraction_elapsed
    market_dirty = market_clean_price + accrued
    if abs(market_clean_price - face_value) < 1e-6:
        return coupon_rate
    def diff(y):
        pv_c = sum(coupon / ((1 + y / coupon_frequency) ** i) for i in range(1, remaining_coupons + 1))
        pv_f = face_value / ((1 + y / coupon_frequency) ** remaining_coupons)
        return pv_c + pv_f - market_dirty
    try:
        y = fsolve(diff, coupon_rate)[0]
        return max(y, 0.0)
    except Exception:
        return None

# Ensure pricing helper functions defined before use

def market_convention_price(face_value, coupon_rate, ytm, remaining_coupons, fraction_elapsed, coupon_frequency):
    coupon = (coupon_rate * face_value) / coupon_frequency
    accrued = coupon * fraction_elapsed
    if ytm == 0:
        dirty = face_value + coupon * remaining_coupons
    else:
        pv_coupons = sum(coupon / ((1 + ytm / coupon_frequency) ** i) for i in range(1, remaining_coupons + 1))
        pv_face = face_value / ((1 + ytm / coupon_frequency) ** remaining_coupons)
        dirty = pv_coupons + pv_face
    clean = dirty - accrued
    return clean, dirty, accrued

def exact_fraction_price(face_value, coupon_rate, ytm, remaining_coupons, days_to_next_coupon, days_in_period, fraction_elapsed, coupon_frequency):
    coupon = (coupon_rate * face_value) / coupon_frequency
    accrued = coupon * fraction_elapsed
    if ytm == 0:
        clean = face_value + coupon * remaining_coupons
        return clean, clean + accrued, accrued
    first_frac = days_to_next_coupon / days_in_period if days_in_period > 0 else 0
    pv_coupons = 0.0
    for i in range(remaining_coupons):
        t = first_frac + i
        pv_coupons += coupon / ((1 + ytm / coupon_frequency) ** t)
    t_face = first_frac + (remaining_coupons - 1)
    pv_face = face_value / ((1 + ytm / coupon_frequency) ** t_face)
    clean = pv_coupons + pv_face
    return clean, clean + accrued, accrued

def price_wrapper(face_value, coupon_rate, ytm, remaining_coupons, days_to_next_coupon, days_in_period, fraction_elapsed, coupon_frequency, mode):
    if mode == "Market":
        return market_convention_price(face_value, coupon_rate, ytm, remaining_coupons, fraction_elapsed, coupon_frequency)
    return exact_fraction_price(face_value, coupon_rate, ytm, remaining_coupons, days_to_next_coupon, days_in_period, fraction_elapsed, coupon_frequency)

def solve_ytm_from_dirty(face_value, coupon_rate, market_dirty, remaining_coupons, days_to_next_coupon, days_in_period, fraction_elapsed, coupon_frequency, mode):
    coupon = (coupon_rate * face_value) / coupon_frequency
    accrued = coupon * fraction_elapsed
    def diff(ytm):
        _, dirty, _ = price_wrapper(face_value, coupon_rate, ytm, remaining_coupons, days_to_next_coupon, days_in_period, fraction_elapsed, coupon_frequency, mode)
        return dirty - market_dirty
    if mode == "Market":
        par_dirty = face_value + accrued
        if abs(market_dirty - par_dirty) < 1e-6:
            return coupon_rate
    try:
        return max(fsolve(diff, coupon_rate)[0], 0.0)
    except Exception:
        return None

# Sidebar for bond information
st.sidebar.header("Bond Information")

# Bond selection from Excel (if available)
if bonds_df is not None:
    bond_options = ["Manual Entry"] + [f"{row.ISIN} | {row.Issuer}" for row in bonds_df.itertuples()]
    selected_bond = st.sidebar.selectbox("Select Bond (Excel)", bond_options)
    if selected_bond != "Manual Entry":
        idx = bond_options.index(selected_bond) - 1
        row = bonds_df.iloc[idx]
        st.session_state['issuer_name_default'] = str(row['Issuer'])
        st.session_state['rating_default'] = str(row.get('S&P Rating',''))
        try:
            # Use CouponDisplayPct for UI value
            coupon_pct = row.get('CouponDisplayPct')
            st.session_state['coupon_default'] = float(coupon_pct) if coupon_pct is not None else 5.0
        except Exception:
            st.session_state['coupon_default'] = 5.0
        try:
            st.session_state['maturity_date_default'] = row['Maturity Date'] if isinstance(row['Maturity Date'], date) else date(2030,12,31)
        except Exception:
            st.session_state['maturity_date_default'] = date(2030,12,31)
    else:
        st.session_state.setdefault('issuer_name_default', 'Sample Issuer')
        st.session_state.setdefault('rating_default', 'A')
        st.session_state.setdefault('coupon_default', 5.0)
        st.session_state.setdefault('maturity_date_default', date(2030,12,31))
else:
    selected_bond = None
    st.sidebar.info("Bond list unavailable or columns missing.")
    st.session_state.setdefault('issuer_name_default', 'Sample Issuer')
    st.session_state.setdefault('rating_default', 'A')
    st.session_state.setdefault('coupon_default', 5.0)
    st.session_state.setdefault('maturity_date_default', date(2030,12,31))

# Issuer details input (renamed from Bond Name)
rating_list = ["AAA", "AA+", "AA", "AA-", "A+", "A", "A-", "BBB+", "BBB", "BBB-", "BB+", "BB", "BB-", "B+", "B", "B-", "CCC+", "CCC", "CCC-", "CC", "C", "D"]
issuer_name = st.sidebar.text_input("Issuer Name", value=st.session_state.get('issuer_name_default', 'Sample Issuer'), key='issuer_name')
# Find rating index safely
_default_rating = st.session_state.get('rating_default', 'A')
try:
    rating_index = rating_list.index(_default_rating)
except ValueError:
    rating_index = rating_list.index('A')
rating = st.sidebar.selectbox("Rating", rating_list, index=rating_index, key='rating')

# Date inputs
col1, col2 = st.sidebar.columns(2)
with col1:
    purchase_date = st.date_input("Purchase Date", value=date.today(), key='purchase_date')
with col2:
    maturity_date = st.date_input("Maturity Date", value=st.session_state.get('maturity_date_default', date(2030, 12, 31)), max_value=date(2080, 12, 31), key='maturity_date')

# Bond parameters
face_value = st.sidebar.number_input("Face Value ($)", value=100.0, min_value=0.01, key='face_value')  # Changed default to 100
coupon_rate = st.sidebar.number_input("Annual Coupon Rate (%)", value=st.session_state.get('coupon_default', 5.0), min_value=0.0, max_value=20.0, step=0.1, key='coupon_rate') / 100

# Treasury benchmark selection and yield input
benchmark_tenor = st.sidebar.selectbox("US Treasury Benchmark Tenor", ["10Y", "20Y", "30Y"], index=0)
treasury_yield = st.sidebar.number_input(f"{benchmark_tenor} Treasury Yield (%)", value=4.0, min_value=0.0, max_value=15.0, step=0.01) / 100

# Coupon frequency
coupon_frequency = st.sidebar.selectbox("Coupon Frequency", [1, 2, 4], index=1, format_func=lambda x: f"{x}x per year ({'Annual' if x==1 else 'Semi-annual' if x==2 else 'Quarterly'})")

# Day count convention selection and pricing mode (pricing remains Market)
day_count_convention = st.sidebar.selectbox("Day Count Convention", ["Actual/Actual", "30/360 US"], index=0)
pricing_mode = "Market"
st.sidebar.caption(f"Using {day_count_convention} day count & Market pricing convention")

# Validation
if maturity_date <= purchase_date:
    st.error("Maturity date must be after purchase date!")
    st.stop()

remaining_coupons, days_to_next_coupon, days_in_period, fraction_elapsed = compute_coupon_context(purchase_date, maturity_date, coupon_frequency, day_count_convention)

# Add tabs: Calculator and Bond List
calc_tab, list_tab = st.tabs(["Calculator", "Bond List"])  # NEW

with calc_tab:
    st.header("Calculation Mode")
    calc_mode = st.radio("Choose calculation method:", ["Calculate YTM from Price", "Calculate Price from YTM"], horizontal=True)

    col_left, col_right = st.columns(2)

    if calc_mode == "Calculate YTM from Price":
        with col_left:
            st.subheader("Input Clean Price → Calculate YTM")
            market_clean_price = st.number_input("Market Clean Price ($)", value=100.0, min_value=0.01)
            if st.button("Calculate YTM", type="primary"):
                ytm = solve_ytm_schedule(face_value, coupon_rate, market_clean_price, coupon_frequency, purchase_date, maturity_date, day_count_convention)
                schedule = build_coupon_schedule(purchase_date, maturity_date, coupon_frequency)
                if schedule:
                    next_coupon = schedule[0]
                    fraction_elapsed, days_in_period, days_to_next = accrual_context(purchase_date, coupon_frequency, next_coupon, day_count_convention)
                    coupon_payment = (coupon_rate * face_value) / coupon_frequency
                    accrued_interest = coupon_payment * fraction_elapsed
                    dirty_price = market_clean_price + accrued_interest
                else:
                    accrued_interest = 0.0
                    dirty_price = market_clean_price
                if ytm is not None:
                    st.session_state['current_ytm'] = ytm  # store current YTM
                    st.success(f"Yield to Maturity: {ytm*100:.3f}%")
                    credit_spread = ytm - treasury_yield
                    st.metric("YTM (Annual)", f"{ytm*100:.3f}%")
                    st.metric(f"{benchmark_tenor} Treasury", f"{treasury_yield*100:.3f}%")
                    st.metric("Credit Spread", f"{credit_spread*100:.3f}%")
                    st.metric("Accrued Interest", f"${accrued_interest:.4f}")
                    st.metric("Dirty Price", f"${dirty_price:.4f}")
                    st.metric("Clean Price", f"${market_clean_price:.4f}")
                    # Chart across price range
                    price_range = np.arange(max(0.01, market_clean_price - 5), market_clean_price + 5.001, 0.1)
                    rows = []
                    for p in price_range:
                        y = solve_ytm_schedule(face_value, coupon_rate, p, coupon_frequency, purchase_date, maturity_date, day_count_convention)
                        if y is None:
                            continue
                        rows.append({"Price": p, "Benchmark": treasury_yield*100, "Credit Spread": max(y - treasury_yield, 0)*100})
                    if rows:
                        chart_df = pd.DataFrame(rows)
                        stacked_df = chart_df.melt(id_vars=["Price"], value_vars=["Benchmark", "Credit Spread"], var_name="Component", value_name="YieldPct")
                        base = alt.Chart(stacked_df).mark_bar().encode(
                            x=alt.X("Price:Q", title="Market Clean Price ($)"),
                            y=alt.Y("YieldPct:Q", stack="zero", title="YTM Components (%)"),
                            color=alt.Color(
                                "Component:N",
                                title="Component",
                                scale=alt.Scale(domain=["Benchmark", "Credit Spread"], range=["#1f77b4", "#ff7f0e"])
                            ),
                            order=alt.Order("Component:N", sort="ascending"),
                            tooltip=["Price", "Component", alt.Tooltip("YieldPct:Q", title="%", format=".3f")]
                        ).properties(height=400)
                        current_ytm_pct = ytm * 100 if ytm is not None else None
                        point_rows = []
                        if current_ytm_pct is not None:
                            point_rows.append({"Price": market_clean_price, "YieldPct": treasury_yield*100, "Label": "Benchmark"})
                            point_rows.append({"Price": market_clean_price, "YieldPct": current_ytm_pct, "Label": "YTM"})
                        points_df = pd.DataFrame(point_rows) if point_rows else pd.DataFrame(columns=["Price","YieldPct","Label"]) 
                        vline = alt.Chart(pd.DataFrame({"Price": [market_clean_price]})).mark_rule(color="#666", strokeDash=[5,5]).encode(x="Price:Q")
                        bench_point = alt.Chart(points_df[points_df.get("Label") == "Benchmark"]).mark_point(size=80, color="#1f77b4").encode(
                            x="Price:Q", y=alt.Y("YieldPct:Q", title=None), tooltip=["Price", alt.Tooltip("YieldPct:Q", title="Benchmark %", format=".3f")]
                        )
                        ytm_point = alt.Chart(points_df[points_df.get("Label") == "YTM"]).mark_point(size=80, color="#111").encode(
                            x="Price:Q", y=alt.Y("YieldPct:Q"), tooltip=["Price", alt.Tooltip("YieldPct:Q", title="YTM %", format=".3f")]
                        )
                        chart = base + vline + bench_point + ytm_point
                        st.altair_chart(chart, use_container_width=True)
                    else:
                        st.warning("Chart data unavailable.")
                else:
                    st.error("YTM solve failed.")
    else:
        with col_left:
            st.subheader("Input YTM → Calculate Price")
            target_ytm = st.number_input("Target YTM (%)", value=5.0, min_value=0.0, max_value=20.0, step=0.1) / 100
            if st.button("Calculate Price", type="primary"):
                clean_price, dirty_price, accrued_interest = price_from_ytm_schedule(face_value, coupon_rate, target_ytm, coupon_frequency, purchase_date, maturity_date, day_count_convention)
                if clean_price is None:
                    st.error("Cannot build schedule.")
                else:
                    st.session_state['current_ytm'] = target_ytm  # store current YTM
                    credit_spread = target_ytm - treasury_yield
                    st.success(f"Clean Price: ${clean_price:.4f}")
                    st.metric("Dirty Price", f"${dirty_price:.4f}")
                    st.metric("Accrued Interest", f"${accrued_interest:.4f}")
                    st.metric("YTM", f"{target_ytm*100:.3f}%")
                    st.metric(f"{benchmark_tenor} Treasury", f"{treasury_yield*100:.3f}%")
                    st.metric("Credit Spread", f"{credit_spread*100:.3f}%")
                    st.metric("Current Yield", f"{(coupon_rate * face_value / clean_price)*100:.3f}%")
                    if clean_price > face_value:
                        st.info(f"Premium: ${clean_price - face_value:.2f}")
                    elif clean_price < face_value:
                        st.warning(f"Discount: ${face_value - clean_price:.2f}")
                    else:
                        st.success("Par")

    with col_right:
        st.subheader("Bond Summary")
        schedule_for_summary = build_coupon_schedule(purchase_date, maturity_date, coupon_frequency)
        next_coupon_str = schedule_for_summary[0].strftime("%Y-%m-%d") if schedule_for_summary else "-"
        summary_data = {
            "Attribute": ["Issuer Name","Rating","Face Value","Coupon Rate","Purchase Date","Maturity Date","Days to Maturity","Coupon Frequency","Treasury Benchmark Tenor","Treasury Yield","Next Coupon","Days to Next Coupon","Accrued Fraction"],
            "Value": [issuer_name, rating, f"${face_value:,.2f}", f"{coupon_rate*100:.2f}%", purchase_date.strftime("%Y-%m-%d"), maturity_date.strftime("%Y-%m-%d"), f"{(maturity_date - purchase_date).days:,} days", f"{coupon_frequency}x", benchmark_tenor, f"{treasury_yield*100:.2f}%", next_coupon_str, f"{int(days_to_next_coupon)}", f"{fraction_elapsed:.4f}"]
        }
        st.table(pd.DataFrame(summary_data))

        st.subheader("Risk Information")
        years_to_maturity = (maturity_date - purchase_date).days / 365
        st.metric("Years to Maturity", f"{years_to_maturity:.2f}")
        current_ytm_val = st.session_state.get('current_ytm')
        if current_ytm_val is not None:
            mdur = modified_duration(face_value, coupon_rate, current_ytm_val, coupon_frequency, purchase_date, maturity_date, day_count_convention)
            if mdur is not None:
                st.metric("Modified Duration", f"{mdur:.4f}")
            else:
                st.metric("Modified Duration", "N/A")
        else:
            st.metric("Modified Duration", "Calculate YTM/Price first")
        risk_colors = {"AAA": "🟢", "AA+": "🟢", "AA": "🟢", "AA-": "🟢", "A+": "🟡", "A": "🟡", "A-": "🟡", "BBB+": "🟠", "BBB": "🟠", "BBB-": "🟠", "BB+": "🔴", "BB": "🔴", "BB-": "🔴", "B+": "🔴", "B": "🔴", "B-": "🔴", "CCC+": "⚫", "CCC": "⚫", "CCC-": "⚫", "CC": "⚫", "C": "⚫", "D": "⚫"}
        st.markdown(f"**Credit Risk:** {risk_colors.get(rating, '❓')} {rating}")

    st.markdown("---")
    st.markdown("**Disclaimer:** Estimates only. YTM decomposition assumes linear separation of benchmark yield and credit spread.")

with list_tab:
    st.subheader("Bond List (Excel)")
    if bonds_df is None:
        st.info("Bond list unavailable or columns missing.")
    else:
        # Prepare display DataFrame
        display_df = bonds_df.copy()
        # Compute years to maturity relative to Purchase Date input
        display_df['Years to Maturity'] = display_df['Maturity Date'].apply(lambda d: (d - purchase_date).days/365 if pd.notnull(d) else None)
        # Rating filter
        available_ratings = sorted([r for r in display_df['S&P Rating'].dropna().unique()])
        selected_ratings = st.multiselect("Filter by Rating", available_ratings, default=available_ratings)
        if selected_ratings:
            display_df = display_df[display_df['S&P Rating'].isin(selected_ratings)]
        # Optional min/max years to maturity filter
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            min_years = st.number_input("Min Years to Maturity", value=0.0, min_value=0.0, step=0.5)
        with col_f2:
            max_years = st.number_input("Max Years to Maturity", value=float(display_df['Years to Maturity'].max() if display_df['Years to Maturity'].notnull().any() else 50), min_value=0.0, step=0.5)
        display_df = display_df[(display_df['Years to Maturity'].isna()) | ((display_df['Years to Maturity'] >= min_years) & (display_df['Years to Maturity'] <= max_years))]
        # Final formatting
        out_df = display_df[['ISIN','Issuer','S&P Rating','CouponDisplayPct','Maturity Date','Years to Maturity']].rename(columns={'CouponDisplayPct':'Coupon (%)'})
        st.dataframe(out_df, use_container_width=True)
        st.caption("Coupon (%) column shows percentage (e.g., 5 means 5%). Years to Maturity computed vs selected Purchase Date.")
