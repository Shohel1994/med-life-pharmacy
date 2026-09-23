"""
=================================================================================
 MED LIFE PHARMACY - ERP System
 (v4: Sales/Purchase Return + Expenses + Profit & Loss + Reports Hub (PDF/Excel))
 Built with: Streamlit + Supabase (supabase-py)
 Apps developed by Shohel Rana - ARJ Studio
=================================================================================

REQUIREMENTS (requirements.txt):
   streamlit>=1.40
   supabase
   pandas
   reportlab
   openpyxl

STREAMLIT SECRETS (.streamlit/secrets.toml):
   SUPABASE_URL = "https://xxxxxxxxxxxx.supabase.co"
   SUPABASE_KEY = "your-supabase-anon-or-service-key"

EXISTING TABLES: master_medicines, medicines, purchases, supplier_ledger,
                 sales, customer_ledger, opening_stock, app_settings, ledger_txn,
                 medicine_batches

NEW THINGS - run this ONCE in Supabase -> SQL Editor (safe to re-run). The app
also shows this SQL on-screen (Settings / Reports Hub) if a table is missing.

   -- (1) Batch / expiry table (one row per received batch)
   create table if not exists medicine_batches (
       id            bigint generated always as identity primary key,
       medicine_name text not null,
       medicine_type text,
       batch_no      text,
       expiry_date   date,
       qty_in        integer not null default 0,
       qty_left      integer not null default 0,
       source        text,            -- 'Purchase' / 'Opening' / 'Existing stock'
       created_at    timestamptz default now()
   );
   create index if not exists idx_batches_name   on medicine_batches (medicine_name);
   create index if not exists idx_batches_expiry on medicine_batches (expiry_date);

   -- (2) Expiry + batch columns on purchase & opening stock history
   alter table purchases     add column if not exists expiry_date date;
   alter table purchases     add column if not exists batch_no text;
   alter table purchases     add column if not exists invoice_no text;
   alter table opening_stock add column if not exists expiry_date date;
   alter table opening_stock add column if not exists batch_no text;

   -- (3) Remember last sale price + average cost price per medicine
   alter table medicines add column if not exists sale_price numeric default 0;
   alter table medicines add column if not exists avg_cost   numeric default 0;

   -- (4) Customer (sales) returns
   create table if not exists sales_returns (
       id             bigint generated always as identity primary key,
       return_date    date not null default current_date,
       voucher_no     text,
       customer_name  text,
       customer_phone text,
       items          text,               -- JSON [{name,qty,unit_price,subtotal}]
       total_amount   numeric not null default 0,
       refund_mode    text,               -- 'Cash Refund' / 'Adjust Due' / 'Store Credit'
       note           text,
       created_at     timestamptz default now()
   );
   create index if not exists idx_sales_returns_voucher on sales_returns (voucher_no);

   -- (5) Supplier (purchase) returns
   create table if not exists purchase_returns (
       id            bigint generated always as identity primary key,
       return_date   date not null default current_date,
       supplier_name text not null,
       medicine_name text not null,
       medicine_type text,
       quantity      integer not null default 0,
       unit_cost     numeric not null default 0,
       total_amount  numeric not null default 0,
       settle_mode   text,                -- 'Reduce Payable Due' / 'Refund Received'
       note          text,
       created_at    timestamptz default now()
   );

   -- (6) Business running costs (salary, electricity, rent, etc.)
   create table if not exists expenses (
       id          bigint generated always as identity primary key,
       exp_date    date not null default current_date,
       category    text not null,        -- Staff Salary / Electricity Bill / Water Bill / ...
       description text,
       amount      numeric not null default 0,
       created_at  timestamptz default now()
   );
   create index if not exists idx_expenses_date on expenses (exp_date);

NOTE: If you use Row Level Security, allow your key to read/write all of the
above tables.
=================================================================================
"""

import calendar
import io
import json
import random
import re
import zlib
from datetime import date, datetime, timedelta
from xml.sax.saxutils import escape

import pandas as pd
import streamlit as st
from supabase import create_client, Client

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, A5, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

# =================================================================================
# PAGE CONFIG - MUST BE FIRST STREAMLIT COMMAND
# =================================================================================
st.set_page_config(
    page_title="Med Life Pharmacy | ERP",
    page_icon="💊",
    layout="wide",
    initial_sidebar_state="expanded",
)


def _st_version() -> tuple:
    try:
        return tuple(int(x) for x in st.__version__.split(".")[:2])
    except Exception:
        return (1, 40)


# Works on old and new Streamlit versions (use_container_width was replaced by width="stretch")
STRETCH = {"width": "stretch"} if _st_version() >= (1, 50) else {"use_container_width": True}

CREDIT_TEXT = "Apps developed by Shohel Rana - ARJ Studio"

EXPIRY_SQL = """create table if not exists medicine_batches (
    id bigint generated always as identity primary key,
    medicine_name text not null,
    medicine_type text,
    batch_no text,
    expiry_date date,
    qty_in integer not null default 0,
    qty_left integer not null default 0,
    source text,
    created_at timestamptz default now()
);
create index if not exists idx_batches_name on medicine_batches (medicine_name);
create index if not exists idx_batches_expiry on medicine_batches (expiry_date);
alter table purchases     add column if not exists expiry_date date;
alter table purchases     add column if not exists batch_no text;
alter table purchases     add column if not exists invoice_no text;
alter table opening_stock add column if not exists expiry_date date;
alter table opening_stock add column if not exists batch_no text;
alter table medicines     add column if not exists sale_price numeric default 0;"""

FULL_SQL = EXPIRY_SQL + """
alter table medicines add column if not exists avg_cost numeric default 0;

create table if not exists sales_returns (
    id bigint generated always as identity primary key,
    return_date date not null default current_date,
    voucher_no text,
    customer_name text,
    customer_phone text,
    items text,
    total_amount numeric not null default 0,
    refund_mode text,
    note text,
    created_at timestamptz default now()
);
create index if not exists idx_sales_returns_voucher on sales_returns (voucher_no);

create table if not exists purchase_returns (
    id bigint generated always as identity primary key,
    return_date date not null default current_date,
    supplier_name text not null,
    medicine_name text not null,
    medicine_type text,
    quantity integer not null default 0,
    unit_cost numeric not null default 0,
    total_amount numeric not null default 0,
    settle_mode text,
    note text,
    created_at timestamptz default now()
);

create table if not exists expenses (
    id bigint generated always as identity primary key,
    exp_date date not null default current_date,
    category text not null,
    description text,
    amount numeric not null default 0,
    created_at timestamptz default now()
);
create index if not exists idx_expenses_date on expenses (exp_date);"""

EXPENSE_CATEGORIES = [
    "Staff Salary", "Shop Rent", "Electricity Bill", "Water Bill", "Internet Bill",
    "Gas Bill", "Transport", "Maintenance / Repair", "Marketing", "Other",
]
REFUND_MODES = ["Adjust Against Due", "Cash Refund", "Store Credit"]
SETTLE_MODES = ["Reduce Payable Due", "Refund Received in Cash"]

# =================================================================================
# CUSTOM CSS - ERP STYLE (top menu, green sidebar, compact rows)
# =================================================================================
st.markdown(
    """
    <style>
        .stApp { background-color: #f4f6f8; }
        .block-container { padding-top: 3.4rem; padding-bottom: 0.8rem; }
        div[data-testid="stVerticalBlock"] { gap: 0.55rem; }
        label[data-testid="stWidgetLabel"] p { font-size: 0.78rem; font-weight: 600; margin-bottom: 0; color: #1e3a8a; }

        .erp-header {
            background: linear-gradient(90deg, #0b3d33 0%, #0f766e 60%, #14b8a6 100%);
            padding: 10px 22px; border-radius: 10px; color: white; margin-top: 0.3rem;
        }
        .erp-header h1 { margin: 0; font-size: 1.3rem; font-weight: 700; color: white; }
        .erp-header p { margin: 2px 0 0 0; font-size: 0.78rem; opacity: 0.9; }

        .sec-title {
            background: repeating-linear-gradient(45deg, #dbe7f3, #dbe7f3 4px, #c9d9ea 4px, #c9d9ea 8px);
            border: 1px solid #9db4cc; border-radius: 3px; padding: 3px 10px;
            font-weight: 700; font-size: 0.85rem; color: #0b2e4a; margin: 0 0 4px 0;
        }

        /* ---- Top menu bar ---- */
        .st-key-topnav {
            background-color: #0f766e; padding: 4px 12px 6px 12px;
            border-radius: 0 0 10px 10px; margin-bottom: 0.6rem;
        }
        .st-key-topnav div[role="radiogroup"] { gap: 6px; }
        .st-key-topnav label p { color: #ffffff !important; font-weight: 600; font-size: 0.95rem; }
        .st-key-topnav label > div:first-child { display: none; }
        .st-key-topnav label { padding: 6px 16px; border-radius: 6px; cursor: pointer; }
        .st-key-topnav label:has(input:checked) { background-color: #0b3d33; }
        .st-key-topnav label:hover { background-color: #14b8a6; }

        /* ---- Left green sidebar ---- */
        section[data-testid="stSidebar"] { background-color: #0b3d33; }
        section[data-testid="stSidebar"] * { color: #eafaf5 !important; }
        section[data-testid="stSidebar"] input { color: #111 !important; }
        section[data-testid="stSidebar"] .stRadio label { font-size: 0.95rem; }
        section[data-testid="stSidebar"] label[data-testid="stWidgetLabel"] p { color: #eafaf5 !important; }
        .arj-credit {
            margin-top: 18px; padding: 10px 6px; text-align: center; font-size: 0.78rem;
            border-top: 1px solid rgba(255,255,255,0.25); letter-spacing: 0.3px;
        }
        .side-alert { padding: 6px 10px; border-radius: 6px; font-size: 0.82rem; font-weight: 700; margin: 4px 0; }
        section[data-testid="stSidebar"] .side-alert.red { background: #fee2e2; }
        section[data-testid="stSidebar"] .side-alert.red, section[data-testid="stSidebar"] .side-alert.red * { color: #991b1b !important; }
        section[data-testid="stSidebar"] .side-alert.amber { background: #fef3c7; }
        section[data-testid="stSidebar"] .side-alert.amber, section[data-testid="stSidebar"] .side-alert.amber * { color: #92400e !important; }

        div[data-testid="stForm"] {
            background-color: white; padding: 20px; border-radius: 12px;
            border: 1px solid #e2e8f0; box-shadow: 0 1px 4px rgba(0,0,0,0.05);
        }
        div.stButton > button, div.stFormSubmitButton > button, div.stDownloadButton > button {
            border-radius: 6px; font-weight: 600; height: 2.4em;
        }
        div[data-testid="stMetric"] {
            background-color: white; padding: 8px 12px; border-radius: 8px;
            border: 1px solid #e2e8f0; box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        }
        h2, h3 { color: #0b3d33; }

        .badge-cash { background-color: #dcfce7; color: #166534; padding: 3px 10px;
            border-radius: 20px; font-size: 0.8rem; font-weight: 700; }
        .badge-credit { background-color: #fef3c7; color: #92400e; padding: 3px 10px;
            border-radius: 20px; font-size: 0.8rem; font-weight: 700; }

        .voucher-box {
            background-color: white; border: 1.5px dashed #0f766e; border-radius: 10px;
            padding: 18px 20px; font-family: 'Courier New', monospace; color: #111827; margin-top: 0.4rem;
        }
        .voucher-box h3 { text-align: center; margin: 0 0 2px 0; color: #0b3d33; }
        .voucher-box .v-sub { text-align: center; font-size: 0.8rem; color: #444; margin-bottom: 10px; }
        .voucher-box hr { border: none; border-top: 1px dashed #999; margin: 8px 0; }
        .voucher-box table { width: 100%; border-collapse: collapse; font-size: 0.85rem; }
        .voucher-box th { text-align: left; border-bottom: 1px solid #999; padding: 4px 2px; }
        .voucher-box td { padding: 4px 2px; }
        .voucher-box .v-right { text-align: right; }
        .voucher-box .v-total-row td { font-weight: 700; font-size: 1.0rem; border-top: 1px dashed #999; }
    </style>
    """,
    unsafe_allow_html=True,
)

NEW_CUSTOM_LABEL = "-- New / Custom Medicine --"
NEW_SUPPLIER_LABEL = "-- New / Custom Supplier --"
NEW_CUSTOMER_LABEL = "-- Walk-in / New Customer --"
DEFAULT_DISCOUNT_PCT = 5.0
DEFAULT_EXPIRY_ALERT_DAYS = 90
UNIT_BOX = "Box / Carton"
UNIT_PCS = "Pcs (Loose)"

MEDICINE_TYPES = [
    "Tablet", "Capsule", "Syrup", "Suspension", "Injection", "Syringe",
    "Drops", "Ointment/Cream", "Inhaler", "Powder/Sachet", "IV Fluid/Saline", "Other",
]

SHOP_NAME = "Med Life Pharmacy"
SHOP_ADDRESS = "Chachkoir, Khalifa Para, Gurudaspur, Natore"


def sec(title: str):
    st.markdown(f'<div class="sec-title">{escape(title)}</div>', unsafe_allow_html=True)


def spacer_line():
    st.markdown("<div style='height:1.7rem'></div>", unsafe_allow_html=True)


# =================================================================================
# SUPABASE CONNECTION
# =================================================================================
@st.cache_resource
def init_connection() -> Client:
    return create_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_KEY"])


try:
    supabase: Client = init_connection()
except Exception as e:
    st.error(f"❌ Could not connect to Supabase. Check secrets.toml. Details: {e}")
    st.stop()


# =================================================================================
# SMALL HELPERS
# =================================================================================
def fetch_all(build, page: int = 1000, max_rows: int = 50000) -> list:
    """Supabase returns max 1000 rows per request. This pages through ALL rows.
    `build` is a function returning a fresh query (already ordered)."""
    rows, start = [], 0
    while True:
        data = build().range(start, start + page - 1).execute().data or []
        rows.extend(data)
        if len(data) < page or len(rows) >= max_rows:
            break
        start += page
    return rows


def fmt_money(val) -> str:
    try:
        return f"৳{float(val):,.2f}"
    except Exception:
        return "৳0.00"


def tk(v) -> str:
    """PDF-safe money (default PDF fonts have no ৳ glyph)."""
    try:
        return f"Tk {float(v):,.2f}"
    except Exception:
        return "Tk 0.00"


def month_end(y: int, m: int) -> date:
    return date(int(y), int(m), calendar.monthrange(int(y), int(m))[1])


def parse_expiry(v):
    """Accepts 'MM/YYYY', 'YYYY-MM', 'YYYY-MM-DD', Excel dates... Returns a date (month-end for month-only) or None."""
    try:
        if v is None or (not isinstance(v, str) and pd.isna(v)):
            return None
        s = str(v).strip()
        if not s or s.lower() in ("nan", "nat", "none", "-"):
            return None
        m = re.match(r"^(\d{1,2})[/\-.](\d{4})$", s)
        if m and 1 <= int(m.group(1)) <= 12:
            return month_end(m.group(2), m.group(1))
        m = re.match(r"^(\d{4})[/\-.](\d{1,2})$", s)
        if m and 1 <= int(m.group(2)) <= 12:
            return month_end(m.group(1), m.group(2))
        d = pd.to_datetime(s, errors="coerce", dayfirst=True)
        if pd.isna(d):
            return None
        d = d.date()
        # Excel turns "03/2027" into 1-Mar-2027 -> treat a 1st-of-month as that month's end
        return month_end(d.year, d.month) if d.day == 1 else d
    except Exception:
        return None


def fmt_exp(v) -> str:
    try:
        if v is None or pd.isna(v):
            return "-"
        return pd.Timestamp(v).strftime("%m/%Y")
    except Exception:
        return "-"


# =================================================================================
# DATA ACCESS - lazy per page, only needed columns, cached, fully paged
# =================================================================================
@st.cache_data(ttl=3600, show_spinner="Loading medicine catalogue (first time only)…")
def fetch_master_medicines() -> pd.DataFrame:
    try:
        rows = fetch_all(lambda: supabase.table("master_medicines").select("brand_name,company_name,form")
                         .order("brand_name").order("company_name").order("form"), max_rows=100000)
        return pd.DataFrame(rows).drop_duplicates()
    except Exception as e:
        st.error(f"❌ Failed to fetch master medicine catalogue: {e}")
        return pd.DataFrame()


@st.cache_data(ttl=60, show_spinner=False)
def fetch_medicines() -> pd.DataFrame:
    rows = None
    for cols in ("id,name,medicine_type,stock,created_at,sale_price,avg_cost",
                "id,name,medicine_type,stock,created_at,sale_price", "id,name,medicine_type,stock,created_at"):
        try:
            rows = fetch_all(lambda: supabase.table("medicines").select(cols).order("name").order("id"))
            break
        except Exception:
            rows = None
    if rows is None:
        st.error("❌ Failed to fetch inventory.")
        return pd.DataFrame()
    df = pd.DataFrame(rows)
    if not df.empty:
        df["stock"] = pd.to_numeric(df["stock"], errors="coerce").fillna(0).astype(int)
        if "medicine_type" not in df.columns:
            df["medicine_type"] = ""
        df["medicine_type"] = df["medicine_type"].fillna("")
        for c in ("sale_price", "avg_cost"):
            if c not in df.columns:
                df[c] = 0.0
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
    return df


@st.cache_data(ttl=60, show_spinner=False)
def get_pick_list() -> pd.DataFrame:
    """One searchable list = my own stock names + the master catalogue."""
    rows, seen = [], set()
    master = fetch_master_medicines()
    master_rows = []
    if not master.empty:
        for brand, company, form in zip(master["brand_name"].fillna(""), master["company_name"].fillna(""), master["form"].fillna("")):
            brand, company, form = str(brand).strip(), str(company).strip(), str(form).strip()
            if not brand:
                continue
            label = f"{brand} ({form}) - {company}" if form or company else brand
            master_rows.append((label, brand, form))
            seen.add(brand.lower())
    meds = fetch_medicines()
    if not meds.empty:
        for nm, tp in zip(meds["name"], meds["medicine_type"]):
            if str(nm).strip() and str(nm).lower() not in seen:
                rows.append((f"{nm}  [my stock]", str(nm), str(tp)))
    rows.extend(master_rows)
    df = pd.DataFrame(rows, columns=["label", "name", "form"])
    df["lc"] = df["label"].str.lower()
    return df


@st.cache_data(ttl=60, show_spinner=False)
def fetch_purchases(limit: int = 200, credit_only: bool = False) -> pd.DataFrame:
    try:
        q = supabase.table("purchases").select("*").order("purchase_date", desc=True).limit(limit)
        if credit_only:
            q = q.eq("payment_type", "Credit")
        df = pd.DataFrame(q.execute().data)
        if not df.empty:
            df["purchase_date"] = pd.to_datetime(df["purchase_date"], errors="coerce")
            for col in ["quantity", "total_amount", "paid_amount", "due_amount"]:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        return df
    except Exception as e:
        st.error(f"❌ Failed to fetch purchase history: {e}")
        return pd.DataFrame()


@st.cache_data(ttl=60, show_spinner=False)
def fetch_purchases_range(start: str, end: str) -> pd.DataFrame:
    try:
        rows = fetch_all(lambda: supabase.table("purchases").select("*").gte("purchase_date", start)
                         .lte("purchase_date", end).order("purchase_date", desc=True).order("id"), max_rows=20000)
        df = pd.DataFrame(rows)
        if not df.empty:
            df["purchase_date"] = pd.to_datetime(df["purchase_date"], errors="coerce")
            for col in ["quantity", "total_amount", "paid_amount", "due_amount"]:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        return df
    except Exception as e:
        st.error(f"❌ Failed to fetch purchase history: {e}")
        return pd.DataFrame()


@st.cache_data(ttl=120, show_spinner=False)
def fetch_supplier_names() -> list:
    try:
        res = supabase.table("purchases").select("supplier_name").limit(2000).execute()
        names = {r["supplier_name"] for r in res.data if r.get("supplier_name")}
        res2 = supabase.table("supplier_ledger").select("supplier_name").execute()
        names |= {r["supplier_name"] for r in res2.data if r.get("supplier_name")}
        return sorted(names)
    except Exception:
        return []


@st.cache_data(ttl=60, show_spinner=False)
def fetch_supplier_ledger() -> pd.DataFrame:
    try:
        rows = fetch_all(lambda: supabase.table("supplier_ledger").select("*").order("total_due", desc=True).order("id"))
        df = pd.DataFrame(rows)
        if not df.empty:
            df["total_due"] = pd.to_numeric(df["total_due"], errors="coerce").fillna(0)
        return df
    except Exception as e:
        st.error(f"❌ Failed to fetch supplier ledger: {e}")
        return pd.DataFrame()


@st.cache_data(ttl=30, show_spinner=False)
def fetch_sales(start: str, end: str) -> pd.DataFrame:
    try:
        rows = fetch_all(lambda: supabase.table("sales").select("*").gte("sale_date", start).lte("sale_date", end)
                         .order("created_at", desc=True).order("id"), max_rows=20000)
        df = pd.DataFrame(rows)
        if not df.empty:
            df["sale_date"] = pd.to_datetime(df["sale_date"], errors="coerce")
            for col in ["subtotal", "discount", "total_amount", "paid_amount", "due_amount"]:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        return df
    except Exception as e:
        st.error(f"❌ Failed to fetch sales history: {e}")
        return pd.DataFrame()


@st.cache_data(ttl=60, show_spinner=False)
def fetch_customer_ledger() -> pd.DataFrame:
    try:
        rows = fetch_all(lambda: supabase.table("customer_ledger").select("*").order("total_due", desc=True).order("id"))
        df = pd.DataFrame(rows)
        if not df.empty:
            df["total_due"] = pd.to_numeric(df["total_due"], errors="coerce").fillna(0)
        return df
    except Exception as e:
        st.error(f"❌ Failed to fetch customer ledger: {e}")
        return pd.DataFrame()


@st.cache_data(ttl=30, show_spinner=False)
def fetch_txn(party_type: str, party_key: str) -> pd.DataFrame:
    """Full history (statement) of one customer / supplier."""
    try:
        res = (
            supabase.table("ledger_txn").select("*").eq("party_type", party_type).eq("party_key", party_key)
            .order("created_at").limit(1000).execute()
        )
        return pd.DataFrame(res.data)
    except Exception:
        return pd.DataFrame()


@st.cache_data(ttl=60, show_spinner=False)
def fetch_opening_stock(limit: int = 200) -> pd.DataFrame:
    try:
        res = supabase.table("opening_stock").select("*").order("created_at", desc=True).limit(limit).execute()
        return pd.DataFrame(res.data)
    except Exception:
        return pd.DataFrame()


@st.cache_data(ttl=120, show_spinner=False)
def fetch_settings() -> dict:
    try:
        res = supabase.table("app_settings").select("key,value").execute()
        return {r["key"]: r["value"] for r in res.data}
    except Exception:
        return {}


def set_setting(key: str, value: str) -> bool:
    try:
        supabase.table("app_settings").upsert({"key": key, "value": value}).execute()
        fetch_settings.clear()
        return True
    except Exception as e:
        st.error(f"❌ Could not save setting (is the `app_settings` table created?): {e}")
        return False


def get_alert_days() -> int:
    try:
        return max(1, int(float(fetch_settings().get("expiry_alert_days", DEFAULT_EXPIRY_ALERT_DAYS))))
    except Exception:
        return DEFAULT_EXPIRY_ALERT_DAYS


# ---------------------------------------------------------------------------------
# EXPIRY / BATCH DATA
# ---------------------------------------------------------------------------------
@st.cache_data(ttl=300, show_spinner=False)
def batches_ready() -> bool:
    try:
        supabase.table("medicine_batches").select("id").limit(1).execute()
        return True
    except Exception:
        return False


@st.cache_data(ttl=60, show_spinner=False)
def fetch_batches() -> pd.DataFrame:
    """All batches that still have stock (qty_left > 0)."""
    try:
        rows = fetch_all(lambda: supabase.table("medicine_batches")
                         .select("id,medicine_name,medicine_type,batch_no,expiry_date,qty_left,source")
                         .gt("qty_left", 0).order("expiry_date").order("id"))
        df = pd.DataFrame(rows)
        if not df.empty:
            df["expiry_date"] = pd.to_datetime(df["expiry_date"], errors="coerce")
            df["qty_left"] = pd.to_numeric(df["qty_left"], errors="coerce").fillna(0).astype(int)
            df["days_left"] = (df["expiry_date"] - pd.Timestamp(date.today())).dt.days
            df["batch_no"] = df["batch_no"].fillna("")
            df["medicine_type"] = df["medicine_type"].fillna("")
        return df
    except Exception:
        return pd.DataFrame()


def expiry_maps():
    """Returns ({name: nearest NON-expired expiry}, {name: expired qty})."""
    b = fetch_batches()
    if b.empty:
        return {}, {}
    d = b[b["expiry_date"].notna()]
    valid = d[d["days_left"] >= 0].groupby("medicine_name")["expiry_date"].min().to_dict()
    expired = d[d["days_left"] < 0].groupby("medicine_name")["qty_left"].sum().to_dict()
    return valid, expired


# ---------------------------------------------------------------------------------
# RETURNS + EXPENSES DATA
# ---------------------------------------------------------------------------------
@st.cache_data(ttl=300, show_spinner=False)
def table_ready(name: str) -> bool:
    try:
        supabase.table(name).select("id").limit(1).execute()
        return True
    except Exception:
        return False


@st.cache_data(ttl=60, show_spinner=False)
def fetch_sales_returns(start: str, end: str) -> pd.DataFrame:
    try:
        rows = fetch_all(lambda: supabase.table("sales_returns").select("*").gte("return_date", start)
                         .lte("return_date", end).order("return_date", desc=True).order("id"), max_rows=20000)
        df = pd.DataFrame(rows)
        if not df.empty:
            df["return_date"] = pd.to_datetime(df["return_date"], errors="coerce")
            df["total_amount"] = pd.to_numeric(df["total_amount"], errors="coerce").fillna(0)
        return df
    except Exception:
        return pd.DataFrame()


@st.cache_data(ttl=30, show_spinner=False)
def returned_qty_for_voucher(voucher_no: str) -> dict:
    """{medicine_name: already-returned qty} for one voucher."""
    try:
        res = supabase.table("sales_returns").select("items").eq("voucher_no", voucher_no).execute()
        out = {}
        for r in res.data or []:
            for it in json.loads(r.get("items") or "[]"):
                out[it["name"]] = out.get(it["name"], 0) + int(it.get("qty", 0))
        return out
    except Exception:
        return {}


@st.cache_data(ttl=60, show_spinner=False)
def fetch_purchase_returns(start: str, end: str) -> pd.DataFrame:
    try:
        rows = fetch_all(lambda: supabase.table("purchase_returns").select("*").gte("return_date", start)
                         .lte("return_date", end).order("return_date", desc=True).order("id"), max_rows=20000)
        df = pd.DataFrame(rows)
        if not df.empty:
            df["return_date"] = pd.to_datetime(df["return_date"], errors="coerce")
            for c in ("quantity", "unit_cost", "total_amount"):
                df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
        return df
    except Exception:
        return pd.DataFrame()


@st.cache_data(ttl=60, show_spinner=False)
def fetch_expenses(start: str, end: str) -> pd.DataFrame:
    try:
        rows = fetch_all(lambda: supabase.table("expenses").select("*").gte("exp_date", start)
                         .lte("exp_date", end).order("exp_date", desc=True).order("id"), max_rows=20000)
        df = pd.DataFrame(rows)
        if not df.empty:
            df["exp_date"] = pd.to_datetime(df["exp_date"], errors="coerce")
            df["amount"] = pd.to_numeric(df["amount"], errors="coerce").fillna(0)
        return df
    except Exception:
        return pd.DataFrame()


def insert_expense(payload: dict) -> bool:
    try:
        supabase.table("expenses").insert(payload).execute()
        return True
    except Exception as e:
        st.error(f"❌ Could not save expense (is the `expenses` table created?): {e}")
        return False


def delete_expense(expense_id) -> bool:
    try:
        supabase.table("expenses").delete().eq("id", expense_id).execute()
        return True
    except Exception as e:
        st.error(f"❌ Could not delete expense: {e}")
        return False


def insert_sales_return(payload: dict) -> bool:
    try:
        supabase.table("sales_returns").insert(payload).execute()
        returned_qty_for_voucher.clear()
        return True
    except Exception as e:
        st.error(f"❌ Could not save sales return (is the `sales_returns` table created?): {e}")
        return False


def insert_purchase_return(payload: dict) -> bool:
    try:
        supabase.table("purchase_returns").insert(payload).execute()
        return True
    except Exception as e:
        st.error(f"❌ Could not save purchase return (is the `purchase_returns` table created?): {e}")
        return False


def update_avg_cost(name: str, added_qty: int, unit_cost: float, prev_stock: int):
    """Weighted-average cost, used for Profit & Loss (COGS). Never blocks the caller."""
    if added_qty <= 0 or unit_cost is None or unit_cost <= 0:
        return
    try:
        med = find_medicine_by_name(name)
        if not med:
            return
        cur = supabase.table("medicines").select("avg_cost").eq("id", med["id"]).limit(1).execute().data
        old_avg = float(cur[0].get("avg_cost") or 0) if cur else 0.0
        old_qty = max(int(prev_stock), 0)
        new_avg = ((old_avg * old_qty) + (float(unit_cost) * added_qty)) / max(old_qty + added_qty, 1)
        supabase.table("medicines").update({"avg_cost": round(new_avg, 4)}).eq("id", med["id"]).execute()
    except Exception:
        pass  # cost tracking must never block a purchase/opening-stock save


def bulk_update_avg_cost(rows: list):
    """rows: [{name, added_qty, unit_cost}]. Call AFTER stock has already been increased
    (so current 'stock' already includes added_qty); avg_cost itself is untouched until here."""
    rows = [r for r in rows if r.get("unit_cost") and float(r["unit_cost"]) > 0 and int(r.get("added_qty") or 0) > 0]
    if not rows:
        return
    try:
        names = list({r["name"] for r in rows})
        cur = {}
        for i in range(0, len(names), 100):
            res = supabase.table("medicines").select("id,name,stock,avg_cost").in_("name", names[i:i + 100]).execute()
            for r in res.data or []:
                cur[r["name"]] = r
        updates = []
        for r in rows:
            m = cur.get(r["name"])
            if not m:
                continue
            added = int(r["added_qty"])
            prev_stock = max(int(m["stock"] or 0) - added, 0)
            old_avg = float(m.get("avg_cost") or 0)
            new_avg = ((old_avg * prev_stock) + (float(r["unit_cost"]) * added)) / max(prev_stock + added, 1)
            updates.append({"id": m["id"], "name": m["name"], "avg_cost": round(new_avg, 4)})
        for i in range(0, len(updates), 200):
            supabase.table("medicines").upsert(updates[i:i + 200]).execute()
    except Exception:
        pass


def get_avg_cost_map() -> dict:
    meds = fetch_medicines()
    if meds.empty or "avg_cost" not in meds.columns:
        return {}
    return dict(zip(meds["name"], pd.to_numeric(meds["avg_cost"], errors="coerce").fillna(0)))


def clear_data_caches():
    for fn in (fetch_medicines, fetch_purchases, fetch_purchases_range, fetch_supplier_ledger, fetch_sales,
               fetch_customer_ledger, fetch_opening_stock, fetch_supplier_names, fetch_txn, fetch_batches,
               get_pick_list, fetch_sales_returns, fetch_purchase_returns, fetch_expenses, returned_qty_for_voucher):
        fn.clear()


def clear_all_caches():
    clear_data_caches()
    fetch_master_medicines.clear()
    fetch_settings.clear()
    batches_ready.clear()
    table_ready.clear()


# ---------------------------------------------------------------------------------
# WRITE HELPERS
# ---------------------------------------------------------------------------------
def _insert_fallback(table: str, rows, optional_cols: tuple) -> bool:
    """Insert rows; if optional columns don't exist yet in the table, retry without them."""
    rows = rows if isinstance(rows, list) else [rows]
    try:
        for i in range(0, len(rows), 200):
            supabase.table(table).insert(rows[i:i + 200]).execute()
        return True
    except Exception:
        try:
            slim = [{k: v for k, v in r.items() if k not in optional_cols} for r in rows]
            for i in range(0, len(slim), 200):
                supabase.table(table).insert(slim[i:i + 200]).execute()
            st.warning(f"⚠️ Saved to `{table}` without some optional columns. Run the SQL from the Near Expiry page to enable them.")
            return True
        except Exception as e:
            st.error(f"❌ Could not save to `{table}`: {e}")
            return False


def find_medicine_by_name(name: str):
    try:
        res = supabase.table("medicines").select("id,stock").eq("name", name).limit(1).execute()
        return res.data[0] if res.data else None
    except Exception as e:
        st.error(f"❌ Failed to look up medicine: {e}")
        return None


def insert_medicine(name: str, stock: int, medicine_type: str, sale_price: float = 0.0) -> bool:
    payload = {"name": name, "stock": stock, "medicine_type": medicine_type}
    if sale_price and sale_price > 0:
        payload["sale_price"] = float(sale_price)
    try:
        supabase.table("medicines").insert(payload).execute()
        return True
    except Exception as e:
        if "sale_price" in payload:
            payload.pop("sale_price")
            try:
                supabase.table("medicines").insert(payload).execute()
                return True
            except Exception as e2:
                e = e2
        st.error(f"❌ Failed to save new medicine: {e}")
        return False


def update_medicine_stock(med_id, new_stock: int, medicine_type: str = None, sale_price: float = None) -> bool:
    payload = {"stock": new_stock}
    if medicine_type:
        payload["medicine_type"] = medicine_type
    if sale_price and sale_price > 0:
        payload["sale_price"] = float(sale_price)
    try:
        supabase.table("medicines").update(payload).eq("id", med_id).execute()
        return True
    except Exception as e:
        if "sale_price" in payload:
            payload.pop("sale_price")
            try:
                supabase.table("medicines").update(payload).eq("id", med_id).execute()
                return True
            except Exception as e2:
                e = e2
        st.error(f"❌ Failed to update stock: {e}")
        return False


def save_or_restock(name: str, added_qty: int, medicine_type: str, sale_price: float = 0.0) -> tuple:
    name = name.strip()
    if not name:
        return False, "Medicine name cannot be empty."
    existing = find_medicine_by_name(name)
    if existing:
        new_stock = int(existing["stock"]) + added_qty
        if update_medicine_stock(existing["id"], new_stock, medicine_type, sale_price):
            return True, f"'{name}' already existed — stock increased to {new_stock} pcs."
        return False, "Failed to update existing medicine stock."
    if insert_medicine(name, added_qty, medicine_type, sale_price):
        return True, f"'{name}' added as a new item with {added_qty} pcs in stock."
    return False, "Failed to save the new medicine."


def bulk_restock(items: list):
    """Fast import: ONE read + a few bulk writes instead of 2 requests per row.
    items: dicts with name, qty, mtype, sale_price. Returns (set_of_saved_names, failed_count)."""
    agg = {}
    for it in items:
        nm = str(it["name"]).strip()
        if not nm or int(it["qty"]) <= 0:
            continue
        a = agg.setdefault(nm, {"qty": 0, "mtype": it["mtype"], "sale_price": 0.0})
        a["qty"] += int(it["qty"])
        a["mtype"] = it["mtype"]
        if float(it.get("sale_price") or 0) > 0:
            a["sale_price"] = float(it["sale_price"])
    try:
        existing = {r["name"]: r for r in fetch_all(lambda: supabase.table("medicines").select("id,name,stock").order("id"))}
    except Exception as e:
        st.error(f"❌ Could not read inventory: {e}")
        return set(), len(agg)

    upd_price, upd_plain, ins = [], [], []
    for nm, a in agg.items():
        ex = existing.get(nm)
        if ex:
            row = {"id": ex["id"], "name": nm, "stock": int(ex["stock"] or 0) + a["qty"], "medicine_type": a["mtype"]}
            if a["sale_price"] > 0:
                row["sale_price"] = a["sale_price"]
                upd_price.append(row)
            else:
                upd_plain.append(row)
        else:
            ins.append({"name": nm, "stock": a["qty"], "medicine_type": a["mtype"], "sale_price": a["sale_price"]})

    done, failed = set(), 0

    def run(rows, fn):
        nonlocal failed
        for i in range(0, len(rows), 200):
            chunk = rows[i:i + 200]
            try:
                fn(chunk)
            except Exception:
                try:
                    fn([{k: v for k, v in r.items() if k != "sale_price"} for r in chunk])
                except Exception as e:
                    failed += len(chunk)
                    st.error(f"❌ A batch of {len(chunk)} items failed: {e}")
                    continue
            done.update(r["name"] for r in chunk)

    run(upd_price, lambda c: supabase.table("medicines").upsert(c).execute())
    run(upd_plain, lambda c: supabase.table("medicines").upsert(c).execute())
    run(ins, lambda c: supabase.table("medicines").insert(c).execute())
    return done, failed


def add_batches(rows: list) -> bool:
    if not rows:
        return True
    try:
        for i in range(0, len(rows), 200):
            supabase.table("medicine_batches").insert(rows[i:i + 200]).execute()
        fetch_batches.clear()
        return True
    except Exception as e:
        st.warning(f"⚠️ Stock was saved, but the expiry batch could NOT be saved. "
                   f"Create the `medicine_batches` table (see ⏰ Near Expiry Report page). Details: {e}")
        return False


def batch_row(name, mtype, batch_no, expiry, qty, source) -> dict:
    return {"medicine_name": name, "medicine_type": mtype, "batch_no": batch_no or None,
            "expiry_date": expiry.isoformat() if expiry else None,
            "qty_in": int(qty), "qty_left": int(qty), "source": source}


def consume_batches(name: str, qty: int):
    """Reduce batch quantities after a sale: earliest NON-expired batch first (FEFO); expired batches last."""
    try:
        res = (supabase.table("medicine_batches").select("id,qty_left,expiry_date")
               .eq("medicine_name", name).gt("qty_left", 0).execute())
        today_s = date.today().isoformat()
        far = "9999-12-31"

        def key(r):
            exp = r.get("expiry_date") or far
            return (exp < today_s, exp)  # non-expired first, then earliest expiry

        remaining = int(qty)
        for r in sorted(res.data or [], key=key):
            if remaining <= 0:
                break
            take = min(remaining, int(r["qty_left"]))
            supabase.table("medicine_batches").update({"qty_left": int(r["qty_left"]) - take}).eq("id", r["id"]).execute()
            remaining -= take
        fetch_batches.clear()
    except Exception:
        pass  # batch tracking must never block a sale


def write_off_batch(batch_id: int, med_name: str, qty: int) -> bool:
    """Remove qty from a batch AND from total stock (expired / damaged / returned)."""
    try:
        cur = supabase.table("medicine_batches").select("qty_left").eq("id", batch_id).limit(1).execute().data
        if not cur:
            st.error("❌ Batch not found.")
            return False
        left = int(cur[0]["qty_left"])
        qty = min(qty, left)
        supabase.table("medicine_batches").update({"qty_left": left - qty}).eq("id", batch_id).execute()
        med = find_medicine_by_name(med_name)
        if med:
            update_medicine_stock(med["id"], max(int(med["stock"]) - qty, 0))
        return True
    except Exception as e:
        st.error(f"❌ Write-off failed: {e}")
        return False


def insert_purchase(payload) -> bool:
    """Accepts a single purchase dict OR a list of dicts (one multi-item invoice)."""
    return _insert_fallback("purchases", payload, ("expiry_date", "batch_no", "invoice_no"))


def log_txn(party_type, party_key, party_name, txn_type, ref_no, due_added, paid, balance_after):
    """Save one ledger history line. Silent if the table isn't created yet (app keeps working)."""
    try:
        supabase.table("ledger_txn").insert({
            "txn_date": date.today().isoformat(), "party_type": party_type, "party_key": party_key,
            "party_name": party_name, "txn_type": txn_type, "ref_no": ref_no,
            "due_added": float(due_added), "paid": float(paid), "balance_after": float(balance_after),
        }).execute()
    except Exception:
        pass


def upsert_supplier_due(supplier_name: str, delta_due: float):
    """Add to supplier's total due. Returns NEW balance, or None on failure."""
    try:
        existing = supabase.table("supplier_ledger").select("total_due").eq("supplier_name", supplier_name).execute()
        now = datetime.now().isoformat()
        if existing.data:
            new_due = float(existing.data[0]["total_due"] or 0) + delta_due
            supabase.table("supplier_ledger").update({"total_due": new_due, "updated_at": now}).eq(
                "supplier_name", supplier_name).execute()
        else:
            new_due = delta_due
            supabase.table("supplier_ledger").insert(
                {"supplier_name": supplier_name, "total_due": new_due, "updated_at": now}).execute()
        return new_due
    except Exception as e:
        st.error(f"❌ Failed to update supplier due: {e}")
        return None


def pay_supplier(supplier_id, new_due: float) -> bool:
    try:
        supabase.table("supplier_ledger").update(
            {"total_due": new_due, "updated_at": datetime.now().isoformat()}).eq("id", supplier_id).execute()
        return True
    except Exception as e:
        st.error(f"❌ Failed to record supplier payment: {e}")
        return False


def generate_voucher_no() -> str:
    return f"MLP-{datetime.now().strftime('%Y%m%d-%H%M%S')}-{random.randint(100, 999)}"


def insert_sale(payload: dict):
    try:
        res = supabase.table("sales").insert(payload).execute()
        return res.data[0] if res.data else payload
    except Exception as e:
        st.error(f"❌ Failed to record sale: {e}")
        return None


def upsert_customer_due(name: str, phone: str, delta_due: float):
    """Add to customer's total due (same phone = same customer). Returns NEW balance, or None on failure."""
    try:
        existing = supabase.table("customer_ledger").select("total_due").eq("customer_phone", phone).execute()
        now = datetime.now().isoformat()
        if existing.data:
            new_due = float(existing.data[0]["total_due"] or 0) + delta_due
            supabase.table("customer_ledger").update(
                {"total_due": new_due, "customer_name": name, "updated_at": now}).eq("customer_phone", phone).execute()
        else:
            new_due = delta_due
            supabase.table("customer_ledger").insert(
                {"customer_name": name, "customer_phone": phone, "total_due": new_due, "updated_at": now}).execute()
        return new_due
    except Exception as e:
        st.error(f"❌ Failed to update customer due: {e}")
        return None


def receive_customer_payment(customer_id, new_due: float) -> bool:
    try:
        supabase.table("customer_ledger").update(
            {"total_due": new_due, "updated_at": datetime.now().isoformat()}).eq("id", customer_id).execute()
        return True
    except Exception as e:
        st.error(f"❌ Failed to record customer payment: {e}")
        return False


def insert_opening_rows(rows: list) -> bool:
    return _insert_fallback("opening_stock", rows,
                            ("purchase_unit", "box_quantity", "units_per_box", "expiry_date", "batch_no"))


# =================================================================================
# PDF HELPERS  (generic table PDF + voucher PDF)
# =================================================================================
@st.cache_data(show_spinner=False, max_entries=40)
def build_table_pdf(title: str, subtitle: str, headers: list, rows: list, weights: list,
                    landscape_mode: bool, summary: list, right_cols: list, row_bg: list, generated: str) -> bytes:
    """One reusable A4 table report (header on every page, page numbers, colour-coded rows)."""
    ps = landscape(A4) if landscape_mode else A4
    margin = 12 * mm
    avail = ps[0] - 2 * margin
    styles = getSampleStyleSheet()
    t_style = ParagraphStyle("t", parent=styles["Title"], fontSize=15, textColor=colors.HexColor("#0b3d33"), spaceAfter=1)
    center = ParagraphStyle("c", parent=styles["Normal"], alignment=1, fontSize=8.5, textColor=colors.HexColor("#444444"))
    h2 = ParagraphStyle("h2", parent=styles["Heading2"], alignment=1, fontSize=12, textColor=colors.HexColor("#0f766e"), spaceBefore=4, spaceAfter=2)
    normal = ParagraphStyle("n", parent=styles["Normal"], fontSize=9, leading=12)
    cell = ParagraphStyle("cell", parent=styles["Normal"], fontSize=8, leading=10)
    cell_r = ParagraphStyle("cellr", parent=cell, alignment=2)
    head = ParagraphStyle("head", parent=cell, textColor=colors.white, fontName="Helvetica-Bold")
    head_r = ParagraphStyle("headr", parent=head, alignment=2)

    buf = io.BytesIO()

    def footer(canvas, doc):
        canvas.saveState()
        canvas.setFont("Helvetica", 7.5)
        canvas.setFillColor(colors.HexColor("#666666"))
        canvas.drawString(margin, 7 * mm, f"{SHOP_NAME}  ·  {CREDIT_TEXT}")
        canvas.drawRightString(ps[0] - margin, 7 * mm, f"Page {doc.page}")
        canvas.restoreState()

    doc = SimpleDocTemplate(buf, pagesize=ps, leftMargin=margin, rightMargin=margin, topMargin=11 * mm,
                            bottomMargin=14 * mm, title=title, author="Shohel Rana - ARJ Studio")
    story = [Paragraph(escape(SHOP_NAME), t_style), Paragraph(escape(SHOP_ADDRESS), center),
             Paragraph(escape(title), h2)]
    if subtitle:
        story.append(Paragraph(escape(subtitle), center))
    story.append(Paragraph(f"Generated: {escape(generated)}", center))
    story.append(Spacer(1, 5))
    for line in summary or []:
        story.append(Paragraph(escape(line), normal))
    if summary:
        story.append(Spacer(1, 5))

    if not rows:
        story.append(Paragraph("No records.", normal))
    else:
        data = [[Paragraph(escape(str(h)), head_r if i in right_cols else head) for i, h in enumerate(headers)]]
        for r in rows:
            data.append([Paragraph(escape(str(c)), cell_r if i in right_cols else cell) for i, c in enumerate(r)])
        tot_w = float(sum(weights))
        tbl = Table(data, colWidths=[avail * w / tot_w for w in weights], repeatRows=1)
        style = [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0f766e")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#cbd5e1")),
            ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f1f5f9")]),
        ]
        for i, bg in enumerate(row_bg or [], start=1):
            if bg:
                style.append(("BACKGROUND", (0, i), (-1, i), colors.HexColor(bg)))
        tbl.setStyle(TableStyle(style))
        story.append(tbl)
    doc.build(story, onFirstPage=footer, onLaterPages=footer)
    return buf.getvalue()


def pdf_button(label: str, key: str, file_name: str, title: str, headers: list, rows: list, weights: list,
               subtitle: str = "", landscape_mode: bool = False, summary: list = None,
               right_cols: list = None, row_bg: list = None):
    data = build_table_pdf(title, subtitle, headers, rows, weights, landscape_mode, summary or [],
                           right_cols or [], row_bg or [], datetime.now().strftime("%d %b %Y, %I:%M %p"))
    st.download_button(label, data=data, file_name=file_name, mime="application/pdf", key=key, **STRETCH)


def build_excel(sheets: dict) -> bytes:
    """sheets: {sheet_name: DataFrame}. Auto-widens columns a little."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for name, df in sheets.items():
            safe = str(name)[:31] or "Sheet1"
            df.to_excel(writer, sheet_name=safe, index=False)
            ws = writer.sheets[safe]
            for i, col in enumerate(df.columns, start=1):
                if len(df) > 0:
                    longest = df[col].astype(str).str.len().clip(upper=40).max()
                    longest = 12 if pd.isna(longest) else int(longest)
                else:
                    longest = 12
                width = min(max(12, longest + 2), 42)
                ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = width
    return buf.getvalue()


def excel_button(label: str, key: str, file_name: str, sheets: dict):
    data = build_excel(sheets)
    st.download_button(label, data=data, file_name=file_name,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       key=key, **STRETCH)


def export_buttons(base_key: str, file_stub: str, pdf_kwargs: dict, excel_sheets: dict):
    """Side-by-side PDF + Excel download buttons for a report."""
    c1, c2 = st.columns(2)
    with c1:
        pdf_button("🖨️ Download PDF", f"pdf_{base_key}", f"{file_stub}.pdf", **pdf_kwargs)
    with c2:
        excel_button("📊 Download Excel", f"xlsx_{base_key}", f"{file_stub}.xlsx", excel_sheets)


def discount_pct_of(sale_row: dict) -> float:
    try:
        sub = float(sale_row.get("subtotal", 0) or 0)
        return round(float(sale_row.get("discount", 0) or 0) / sub * 100, 2) if sub > 0 else 0.0
    except Exception:
        return 0.0


def _parse_items(sale_row: dict) -> list:
    try:
        return json.loads(sale_row.get("items") or "[]")
    except Exception:
        return []


def _date_str(sale_row: dict) -> str:
    d = sale_row.get("sale_date", "")
    return d.strftime("%Y-%m-%d") if hasattr(d, "strftime") else str(d)


def render_voucher_html(sale_row: dict) -> str:
    items = _parse_items(sale_row)
    rows_html = ""
    for it in items:
        rows_html += (
            f"<tr><td>{escape(str(it.get('name', '')))}"
            f"{' (' + escape(str(it.get('medicine_type'))) + ')' if it.get('medicine_type') else ''}</td>"
            f"<td class='v-right'>{it.get('qty', 0)}</td>"
            f"<td class='v-right'>{fmt_money(it.get('unit_price', 0))}</td>"
            f"<td class='v-right'>{fmt_money(it.get('subtotal', 0))}</td></tr>"
        )
    payment_mode = sale_row.get("payment_mode", "")
    badge_class = "badge-cash" if payment_mode == "Cash" else "badge-credit"
    pct = discount_pct_of(sale_row)
    cust = escape(str(sale_row.get("customer_name") or "Walk-in Customer"))
    phone = sale_row.get("customer_phone")
    return f"""
    <div class="voucher-box">
        <h3>💊 {SHOP_NAME}</h3>
        <div class="v-sub">{SHOP_ADDRESS}</div>
        <hr>
        <div><b>Voucher No:</b> {escape(str(sale_row.get('voucher_no', '')))}</div>
        <div><b>Date:</b> {_date_str(sale_row)}</div>
        <div><b>Customer:</b> {cust} {('· ' + escape(str(phone))) if phone else ''}</div>
        <hr>
        <table>
            <tr><th>Item</th><th class="v-right">Qty</th><th class="v-right">Price</th><th class="v-right">Amount</th></tr>
            {rows_html}
        </table>
        <hr>
        <table>
            <tr><td>Subtotal</td><td class="v-right">{fmt_money(sale_row.get('subtotal', 0))}</td></tr>
            <tr><td>Discount ({pct:g}%)</td><td class="v-right">- {fmt_money(sale_row.get('discount', 0))}</td></tr>
            <tr class="v-total-row"><td>Total Payable</td><td class="v-right">{fmt_money(sale_row.get('total_amount', 0))}</td></tr>
            <tr><td>Paid Amount</td><td class="v-right">{fmt_money(sale_row.get('paid_amount', 0))}</td></tr>
            <tr><td>Due Amount</td><td class="v-right">{fmt_money(sale_row.get('due_amount', 0))}</td></tr>
        </table>
        <hr>
        <div>Payment Mode: <span class="{badge_class}">{payment_mode}</span></div>
        <div class="v-sub" style="margin-top:10px;">Thank you for shopping with us!</div>
    </div>
    """


def build_voucher_pdf(sale_row: dict) -> bytes:
    """Downloadable PDF voucher (A5)."""
    items = _parse_items(sale_row)
    styles = getSampleStyleSheet()
    title = ParagraphStyle("t", parent=styles["Title"], fontSize=16, textColor=colors.HexColor("#0b3d33"), spaceAfter=2)
    center = ParagraphStyle("c", parent=styles["Normal"], alignment=1, fontSize=8.5, textColor=colors.HexColor("#444444"))
    normal = ParagraphStyle("n", parent=styles["Normal"], fontSize=9, leading=12)

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A5, leftMargin=12 * mm, rightMargin=12 * mm, topMargin=12 * mm, bottomMargin=12 * mm,
        title=f"Voucher {sale_row.get('voucher_no', '')}", author="Shohel Rana - ARJ Studio",
    )
    story = [
        Paragraph(SHOP_NAME, title),
        Paragraph(escape(SHOP_ADDRESS), center),
        Spacer(1, 6),
        Paragraph(f"<b>Voucher No:</b> {escape(str(sale_row.get('voucher_no', '')))}", normal),
        Paragraph(f"<b>Date:</b> {_date_str(sale_row)}", normal),
        Paragraph(
            f"<b>Customer:</b> {escape(str(sale_row.get('customer_name') or 'Walk-in Customer'))}"
            + (f" &nbsp;|&nbsp; {escape(str(sale_row.get('customer_phone')))}" if sale_row.get("customer_phone") else ""),
            normal,
        ),
        Spacer(1, 8),
    ]

    data = [["Item", "Qty", "Price", "Amount"]]
    for it in items:
        nm = str(it.get("name", "")) + (f" ({it.get('medicine_type')})" if it.get("medicine_type") else "")
        data.append([Paragraph(escape(nm), normal), str(it.get("qty", 0)), tk(it.get("unit_price", 0)), tk(it.get("subtotal", 0))])
    items_tbl = Table(data, colWidths=[58 * mm, 12 * mm, 25 * mm, 27 * mm], repeatRows=1)
    items_tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0f766e")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f1f5f9")]),
        ("LINEBELOW", (0, -1), (-1, -1), 0.6, colors.HexColor("#999999")),
    ]))
    story += [items_tbl, Spacer(1, 8)]

    pct = discount_pct_of(sale_row)
    tot = [
        ["Subtotal", tk(sale_row.get("subtotal", 0))],
        [f"Discount ({pct:g}%)", "- " + tk(sale_row.get("discount", 0))],
        ["Total Payable", tk(sale_row.get("total_amount", 0))],
        ["Paid Amount", tk(sale_row.get("paid_amount", 0))],
        ["Due Amount", tk(sale_row.get("due_amount", 0))],
    ]
    tot_tbl = Table(tot, colWidths=[95 * mm, 27 * mm], hAlign="RIGHT")
    tot_tbl.setStyle(TableStyle([
        ("FONTSIZE", (0, 0), (-1, -1), 9.5),
        ("ALIGN", (1, 0), (1, -1), "RIGHT"),
        ("FONTNAME", (0, 2), (-1, 2), "Helvetica-Bold"),
        ("LINEABOVE", (0, 2), (-1, 2), 0.8, colors.HexColor("#333333")),
    ]))
    story += [
        tot_tbl, Spacer(1, 8),
        Paragraph(f"<b>Payment Mode:</b> {escape(str(sale_row.get('payment_mode', '')))}", normal),
        Spacer(1, 14),
        Paragraph("Thank you for shopping with us!", center),
        Spacer(1, 4),
        Paragraph(CREDIT_TEXT, center),
    ]
    doc.build(story)
    return buf.getvalue()


def statement_table(party_type: str, party_key: str, party_name: str = ""):
    """Running-balance statement for one customer / supplier + PDF."""
    txn = fetch_txn(party_type, party_key)
    if txn.empty:
        st.info(
            "No history lines yet. History records from the time the ledger_txn table was created "
            "(older dues appear only in the balance)."
        )
        return
    show = txn.rename(columns={
        "txn_date": "Date", "txn_type": "Type", "ref_no": "Ref / Voucher",
        "due_added": "Due Added (TK)", "paid": "Paid (TK)", "balance_after": "Balance (TK)"})
    cols = ["Date", "Type", "Ref / Voucher", "Due Added (TK)", "Paid (TK)", "Balance (TK)"]
    st.dataframe(show[cols], hide_index=True, height=300, **STRETCH)
    pdf_rows = [[str(r["Date"]), str(r["Type"] or ""), str(r["Ref / Voucher"] or ""), tk(r["Due Added (TK)"]),
                 tk(r["Paid (TK)"]), tk(r["Balance (TK)"])] for _, r in show.iterrows()]
    pdf_button("🖨️ Statement PDF", f"pdf_stmt_{party_type}_{zlib.crc32(party_key.encode())}",
               f"statement_{party_type}.pdf", f"{party_type.title()} Statement - {party_name or party_key}",
               ["Date", "Type", "Ref / Voucher", "Due Added", "Paid", "Balance"], pdf_rows, [1.1, 1.5, 2.2, 1.3, 1.3, 1.3],
               right_cols=[3, 4, 5])


# =================================================================================
# HEADER + MENU
# =================================================================================
st.markdown(
    f"""
    <div class="erp-header">
        <h1>💊 {SHOP_NAME}</h1>
        <p>ERP System — Purchase, Sales, Stock &amp; Ledger Management</p>
    </div>
    """,
    unsafe_allow_html=True,
)

MENU = {
    "Sales": ["🛒 Sales / POS", "🧾 Sales History / Vouchers", "↩️ Sales Return"],
    "Purchase": ["➕ Purchase Entry", "📜 Purchase History", "↩️ Purchase Return"],
    "Stock": ["📥 Opening Stock Entry", "📦 Inventory Report", "⏰ Near Expiry Report"],
    "Ledger": ["📗 Customer Ledger", "🧾 Supplier Ledger"],
    "Finance": ["💸 Expenses", "📈 Profit & Loss"],
    "Reports": ["📑 Reports Hub"],
    "Settings": ["⚙️ Settings"],
}

with st.container(key="topnav"):
    group = st.radio("Menu", list(MENU.keys()), horizontal=True, label_visibility="collapsed", key="top_group")

st.sidebar.markdown("### 📋 Menu")
menu_query = st.sidebar.text_input(
    "Search by menu name", placeholder="🔍 Search by menu name", label_visibility="collapsed", key="menu_search"
)

if menu_query.strip():
    results = [p for pages in MENU.values() for p in pages if menu_query.strip().lower() in p.lower()]
    if results:
        page = st.sidebar.radio("Results", results, label_visibility="collapsed", key="sub_search")
    else:
        st.sidebar.warning("No menu found.")
        page = MENU[group][0]
else:
    page = st.sidebar.radio("Pages", MENU[group], label_visibility="collapsed", key=f"sub_{group}")

# ---- Expiry alerts in the sidebar (always visible) ----
try:
    _b = fetch_batches()
    if not _b.empty:
        _d = _b[_b["expiry_date"].notna()]
        _n_exp = int((_d["days_left"] < 0).sum())
        _days = get_alert_days()
        _n_near = int(((_d["days_left"] >= 0) & (_d["days_left"] <= _days)).sum())
        if _n_exp:
            st.sidebar.markdown(f'<div class="side-alert red">⛔ {_n_exp} batch EXPIRED</div>', unsafe_allow_html=True)
        if _n_near:
            st.sidebar.markdown(f'<div class="side-alert amber">⏰ {_n_near} batch expire within {_days} days</div>',
                                unsafe_allow_html=True)
except Exception:
    pass

st.sidebar.markdown("---")
if st.sidebar.button("🔄 Refresh Data", **STRETCH):
    clear_all_caches()
    st.rerun()
st.sidebar.caption(f"🕒 {datetime.now().strftime('%d %b %Y, %I:%M %p')}")
st.sidebar.markdown(f'<div class="arj-credit">{CREDIT_TEXT}</div>', unsafe_allow_html=True)


# =================================================================================
# SHARED WIDGETS: medicine picker, Box/Pcs row, Expiry row
# =================================================================================
def medicine_picker(prefix: str):
    """Explicit toggle between 'pick an existing medicine' and 'add a brand-new one' —
    this avoids the old confusing disabled/greyed 'New Medicine Name' box.
    Returns (is_new: bool, medicine_name or None, form_hint: str)."""
    mode = st.radio("Medicine source", ["🔍 Pick Existing Medicine", "🆕 Add a Brand-New Medicine"],
                    horizontal=True, key=f"{prefix}_mode", label_visibility="collapsed")
    is_new = mode.startswith("🆕")

    if is_new:
        name = st.text_input("New Medicine Name *", placeholder="Type the medicine name (e.g. Napa 500mg)",
                             key=f"{prefix}_newname")
        return True, (name.strip() or None), ""

    pick = get_pick_list()
    if pick.empty:
        st.error("⚠️ কোনো medicine data লোড হয়নি (আপনার stock-ও শূন্য দেখাচ্ছে, catalogue-ও)। "
                "এর মানে app আপনার Supabase থেকে ডেটা পড়তে পারছে না। সাধারণত এর কারণ **Row Level "
                "Security (RLS)** — Supabase-এ `medicines` টেবিলে SELECT policy না থাকলে ডেটা থাকলেও app "
                "কিছু দেখতে পায় না। Supabase → Authentication → Policies-এ গিয়ে `medicines` (ও অন্য টেবিল) "
                "টেবিলে read/write policy দিন, তারপর Sidebar থেকে 🔄 Refresh Data চাপুন।")
        return True, None, ""

    # A single selectbox with EVERY option loaded has its own built-in type-ahead search —
    # click it and type; it filters live, letter by letter, with no Enter key and no server round-trip.
    MAX_OPTIONS = 3000
    options_df = pick if len(pick) <= MAX_OPTIONS else pick.head(MAX_OPTIONS)
    if len(pick) > MAX_OPTIONS:
        st.caption(f"⚠️ {len(pick)}টা medicine আছে — dropdown-এ প্রথম {MAX_OPTIONS}টা লোড করা হয়েছে। "
                  "খুঁজে না পেলে '🆕 Add a Brand-New Medicine' ব্যবহার করুন।")

    labels = options_df["label"].tolist()
    sel = st.selectbox("Medicine *  (এখানে ক্লিক করে সরাসরি টাইপ করুন)", labels, index=None,
                       placeholder="টাইপ করুন… যেমন Ace, Napa, Zimax", key=f"{prefix}_sel")
    name_map = dict(zip(options_df["label"], options_df["name"]))
    form_map = dict(zip(options_df["label"], options_df["form"]))

    n_mine = int(pick["label"].str.endswith("[my stock]").sum())
    n_cat = len(pick) - n_mine
    st.caption(f"ℹ️ মোট {len(pick)}টা medicine লোড হয়েছে  —  নিজের stock: **{n_mine}**টা  ·  Catalogue: **{n_cat}**টা")
    if st.checkbox("👀 লোড হওয়া সব medicine-এর নাম দেখান (spelling মিলিয়ে দেখতে)", key=f"{prefix}_showall"):
        st.dataframe(pick[["name"]].rename(columns={"name": "Loaded Medicine Names"}), hide_index=True, height=220, **STRETCH)

    return False, name_map.get(sel), form_map.get(sel, "")


def qty_row(prefix: str):
    """One horizontal row: Unit | Boxes | Pcs per Box | Total Pcs. Returns (unit, boxes, ppb, total_pcs)."""
    c1, c2, c3, c4 = st.columns([1.3, 1, 1, 1.2])
    with c1:
        unit = st.selectbox("Unit *", [UNIT_BOX, UNIT_PCS], key=f"{prefix}_unit")
    is_box = unit == UNIT_BOX
    with c2:
        boxes = st.number_input("No. of Boxes", min_value=1, value=1, step=1, key=f"{prefix}_boxes", disabled=not is_box)
    with c3:
        ppb = st.number_input("Pcs per Box", min_value=1, value=10, step=1, key=f"{prefix}_ppb", disabled=not is_box)
    with c4:
        if is_box:
            total = int(boxes) * int(ppb)
            st.text_input("Total Pcs (auto)", value=f"{total}", disabled=True, key=f"{prefix}_tot_auto")
        else:
            total = int(st.number_input("Total Pcs *", min_value=1, value=10, step=1, key=f"{prefix}_tot_pcs"))
    if not is_box:
        return unit, None, None, total
    return unit, int(boxes), int(ppb), total


def expiry_row(prefix: str, with_sale_price: bool = False):
    """Batch no + Expiry (month/year, like the printed pack) + optional sale price.
    Returns (batch_no or None, expiry_date or None, sale_price)."""
    today = date.today()
    cols = st.columns([1.3, 0.9, 1, 1.1, 1.2] if with_sale_price else [1.3, 0.9, 1, 1.1])
    with cols[3]:
        spacer_line()
        no_exp = st.checkbox("No expiry", key=f"{prefix}_noexp", help="Tick if this item has no expiry date")
    with cols[0]:
        batch_no = st.text_input("Batch No", placeholder="optional", key=f"{prefix}_batch")
    with cols[1]:
        month = st.selectbox("Expiry Month *", list(range(1, 13)), index=today.month - 1,
                             format_func=lambda m: f"{m:02d}", key=f"{prefix}_expm", disabled=no_exp)
    with cols[2]:
        years = list(range(today.year - 1, today.year + 11))
        year = st.selectbox("Expiry Year *", years, index=years.index(today.year + 2), key=f"{prefix}_expy", disabled=no_exp)
    sale_price = 0.0
    if with_sale_price:
        with cols[4]:
            sale_price = st.number_input("Sale / pc (optional)", min_value=0.0, value=0.0, step=0.5, key=f"{prefix}_sp",
                                         help="Saved, and auto-filled in Sales / POS")
    exp = None if no_exp else month_end(year, month)
    if exp and exp < today:
        st.warning(f"⚠️ This expiry ({exp.strftime('%m/%Y')}) is already in the past.")
    return (batch_no.strip() or None), exp, float(sale_price)


def date_range_picker(prefix: str, default: str = "This Month"):
    """Daily / Monthly / Yearly / Custom date-range picker used by all reports.
    Returns (start_date, end_date, label_for_titles)."""
    today = date.today()
    presets = ["Today", "This Week", "This Month", "This Year", "Custom Range"]
    c1, c2, c3 = st.columns([1.3, 1, 1])
    with c1:
        preset = st.selectbox("Period", presets, index=presets.index(default), key=f"{prefix}_preset")

    if preset == "Today":
        start, end, label = today, today, today.strftime("%d %b %Y")
    elif preset == "This Week":
        start = today - timedelta(days=today.weekday())
        end, label = today, f"Week of {start.strftime('%d %b')} – {today.strftime('%d %b %Y')}"
    elif preset == "This Month":
        start, end, label = date(today.year, today.month, 1), today, today.strftime("%B %Y")
    elif preset == "This Year":
        start, end, label = date(today.year, 1, 1), today, str(today.year)
    else:
        with c2:
            start = st.date_input("From", value=date(today.year, today.month, 1), key=f"{prefix}_from")
        with c3:
            end = st.date_input("To", value=today, key=f"{prefix}_to")
        label = f"{start.strftime('%d %b %Y')} – {end.strftime('%d %b %Y')}"

    if start > end:
        st.error("⚠️ 'From' date cannot be after 'To' date.")
        st.stop()
    return start, end, label


# =================================================================================
# PAGE: PURCHASE ENTRY  (multi-item invoice — add many medicines, save once)
# =================================================================================
def generate_invoice_no() -> str:
    return f"PUR-{datetime.now().strftime('%Y%m%d-%H%M%S')}-{random.randint(100, 999)}"


def render_purchase():
    supplier_options = [NEW_SUPPLIER_LABEL] + fetch_supplier_names()
    if "purchase_cart" not in st.session_state:
        st.session_state.purchase_cart = []
    if "purchase_cart_ver" not in st.session_state:
        st.session_state.purchase_cart_ver = 0
    pv = st.session_state.purchase_cart_ver

    st.caption("This is one invoice: add every medicine you bought from this supplier below, "
              "then fill the supplier & payment details once and save the whole invoice together — "
              "exactly like a real purchase bill.")

    # ---------- Add item to cart ----------
    with st.container(border=True):
        sec("Add Medicine to This Invoice")
        r1 = st.columns([2.4, 1.6])
        with r1[0]:
            is_new, final_name, hint = medicine_picker(f"pur_add_{pv}")
        default_idx = next((i for i, t in enumerate(MEDICINE_TYPES) if hint and t.lower().startswith(hint.lower()[:4])), 0)
        with r1[1]:
            medicine_type = st.selectbox("Medicine Type *", MEDICINE_TYPES, index=default_idx,
                                         key=f"pur_add_type_{pv}_{zlib.crc32((final_name or '').encode())}")
        custom_type = ""
        if medicine_type == "Other":
            custom_type = st.text_input("Specify Type", key=f"pur_add_ctype_{pv}")

        unit, box_quantity, units_per_box, total_pcs = qty_row(f"pur_add_{pv}")
        batch_no, expiry, _ = expiry_row(f"pur_add_{pv}")

        r3 = st.columns([1, 1, 1])
        with r3[0]:
            unit_cost = st.number_input("Unit Cost / pc (TK) *", min_value=0.0, value=0.0, step=0.5, key=f"pur_add_cost_{pv}")
        with r3[1]:
            item_sale_price = st.number_input("Sale Price / pc (optional)", min_value=0.0, value=0.0, step=0.5,
                                              key=f"pur_add_sp_{pv}", help="Auto-fills the price in Sales / POS")
        with r3[2]:
            spacer_line()
            add_clicked = st.button("➕ Add to Invoice", type="primary", key=f"pur_add_btn_{pv}", **STRETCH)

        if add_clicked:
            final_type = custom_type.strip() if medicine_type == "Other" else medicine_type
            errs = []
            if not final_name:
                errs.append("Please select an existing medicine, or type a name under '🆕 Add a Brand-New Medicine'.")
            if medicine_type == "Other" and not final_type:
                errs.append("Please specify the custom medicine type.")
            if total_pcs <= 0:
                errs.append("Quantity must be greater than 0.")
            if unit_cost <= 0:
                errs.append("Unit cost must be greater than 0.")
            if errs:
                for e in errs:
                    st.error(f"⚠️ {e}")
            else:
                st.session_state.purchase_cart.append({
                    "name": final_name, "mtype": final_type, "unit": unit, "boxes": box_quantity, "ppb": units_per_box,
                    "qty": total_pcs, "batch_no": batch_no, "expiry": expiry.isoformat() if expiry else None,
                    "expiry_disp": expiry.strftime("%m/%Y") if expiry else "-",
                    "unit_cost": float(unit_cost), "sale_price": float(item_sale_price),
                    "line_total": round(total_pcs * unit_cost, 2)})
                st.session_state.purchase_cart_ver += 1
                st.rerun()

    if not st.session_state.purchase_cart:
        st.info("No items added yet. Add your first medicine above.")
        _recent_purchases_block()
        return

    cart = st.session_state.purchase_cart

    left, right = st.columns([1.6, 1])

    # ---------- Cart (editable) ----------
    with left:
        with st.container(border=True):
            sec(f"Invoice Items ({len(cart)})  —  edit Qty / Unit Cost directly, tick Remove to delete")
            view = pd.DataFrame(cart)[["name", "mtype", "qty", "expiry_disp", "unit_cost", "line_total"]].rename(
                columns={"name": "Medicine", "mtype": "Type", "qty": "Qty", "expiry_disp": "Expiry",
                        "unit_cost": "Unit Cost", "line_total": "Line Total"})
            view["Remove"] = False
            edited = st.data_editor(
                view, key=f"pur_cart_editor_{pv}", hide_index=True,
                disabled=["Medicine", "Type", "Expiry", "Line Total"], height=min(320, 90 + 36 * len(cart)),
                column_config={
                    "Qty": st.column_config.NumberColumn(min_value=1, step=1),
                    "Unit Cost": st.column_config.NumberColumn(min_value=0.0, step=0.5, format="%.2f"),
                    "Remove": st.column_config.CheckboxColumn("Remove"),
                }, **STRETCH)

            new_cart, removed = [], False
            for i, row in edited.iterrows():
                if bool(row["Remove"]):
                    removed = True
                    continue
                it = dict(cart[i])
                qty = max(int(row["Qty"]) if pd.notna(row["Qty"]) else it["qty"], 1)
                cost = float(row["Unit Cost"]) if pd.notna(row["Unit Cost"]) else it["unit_cost"]
                it.update(qty=qty, unit_cost=cost, line_total=round(qty * cost, 2))
                new_cart.append(it)
            st.session_state.purchase_cart = new_cart
            if removed:
                st.session_state.purchase_cart_ver += 1
                st.rerun()
            if st.button("🗑️ Clear All Items", key=f"pur_clear_{pv}"):
                st.session_state.purchase_cart = []
                st.session_state.purchase_cart_ver += 1
                st.rerun()

    cart = st.session_state.purchase_cart
    total_amount = round(sum(it["line_total"] for it in cart), 2)

    # ---------- Invoice-level supplier / payment / save ----------
    with right:
        with st.container(border=True):
            sec("Invoice Details")
            pur_date = st.date_input("Purchase Date", value=date.today(), key="pur_inv_date")
            sel_supplier = st.selectbox("Supplier / Company *", supplier_options, index=None,
                                        placeholder="Select supplier", key="pur_inv_supp")
            custom_supplier = ""
            if sel_supplier == NEW_SUPPLIER_LABEL:
                custom_supplier = st.text_input("New Supplier Name", key="pur_inv_custom_supp")
            payment_type = st.selectbox("Payment Type *", ["Cash", "Credit"], key="pur_inv_pay")

            bad_paid = False
            if payment_type == "Credit":
                paid_amount = st.number_input("Paid Now (TK)", min_value=0.0, value=0.0, step=1.0, key="pur_inv_paid")
                bad_paid = paid_amount > total_amount
            else:
                paid_amount = total_amount
                st.text_input("Paid Now (TK)", value=f"{total_amount:,.2f}", disabled=True, key="pur_inv_paid_ro")
            due_amount = max(round(total_amount - paid_amount, 2), 0.0)

            st.metric("Invoice Total", fmt_money(total_amount))
            st.metric("Due", fmt_money(due_amount))
            if bad_paid:
                st.error("⚠️ Paid amount cannot be more than the invoice total.")

            final_supplier = custom_supplier.strip() if sel_supplier == NEW_SUPPLIER_LABEL else (sel_supplier or "")
            ready = bool(final_supplier) and len(cart) > 0 and not bad_paid

            if st.button("💾 Save Purchase Invoice", type="primary", disabled=not ready, **STRETCH):
                invoice_no = generate_invoice_no()
                item_subtotals = [it["line_total"] for it in cart]
                remaining_paid = paid_amount if payment_type == "Credit" else total_amount
                rows, avg_cost_jobs, batch_jobs = [], [], []
                for i, it in enumerate(cart):
                    sub = item_subtotals[i]
                    if payment_type == "Credit":
                        item_paid = round(remaining_paid, 2) if i == len(cart) - 1 else (
                            round(sub / total_amount * paid_amount, 2) if total_amount > 0 else 0.0)
                        remaining_paid -= item_paid
                    else:
                        item_paid = sub
                    item_due = max(round(sub - item_paid, 2), 0.0)
                    rows.append({
                        "purchase_date": pur_date.isoformat(), "medicine_name": it["name"], "medicine_type": it["mtype"],
                        "supplier_name": final_supplier, "purchase_unit": "Box" if it["unit"] == UNIT_BOX else "Pcs",
                        "box_quantity": it["boxes"], "units_per_box": it["ppb"], "quantity": it["qty"],
                        "payment_type": payment_type, "total_amount": sub,
                        "paid_amount": item_paid if payment_type == "Credit" else sub,
                        "due_amount": item_due if payment_type == "Credit" else 0.0,
                        "expiry_date": it["expiry"], "batch_no": it["batch_no"], "invoice_no": invoice_no,
                    })
                purchase_ok = insert_purchase(rows)
                if not purchase_ok:
                    st.stop()

                stock_msgs = []
                for it in cart:
                    existing_before = find_medicine_by_name(it["name"])
                    prev_stock = int(existing_before["stock"]) if existing_before else 0
                    ok, msg = save_or_restock(it["name"], it["qty"], it["mtype"], it["sale_price"])
                    if ok:
                        update_avg_cost(it["name"], it["qty"], it["unit_cost"], prev_stock)
                        if it["expiry"]:
                            add_batches([batch_row(it["name"], it["mtype"], it["batch_no"],
                                                   datetime.fromisoformat(it["expiry"]).date(), it["qty"], "Purchase")])
                        stock_msgs.append(f"{it['name']} +{it['qty']}")

                ledger_ok = True
                if payment_type == "Credit" and due_amount > 0:
                    new_bal = upsert_supplier_due(final_supplier, due_amount)
                    ledger_ok = new_bal is not None
                    if ledger_ok:
                        log_txn("supplier", final_supplier, final_supplier, "Credit Purchase",
                                invoice_no, due_amount, 0, new_bal)

                if ledger_ok:
                    clear_data_caches()
                    st.session_state.purchase_cart = []
                    st.session_state.purchase_cart_ver += 1
                    st.session_state.purchase_flash = (
                        f"✅ Invoice {invoice_no} saved — {len(rows)} item(s), Total: {fmt_money(total_amount)}"
                        + (f" | Due: {fmt_money(due_amount)}" if payment_type == "Credit" and due_amount > 0 else "")
                    )
                    st.rerun()
                else:
                    st.error("❌ Invoice items were saved but the supplier ledger update had a problem. Please check Supplier Ledger.")

    if st.session_state.get("purchase_flash"):
        st.success(st.session_state.pop("purchase_flash"))

    _recent_purchases_block()


def _recent_purchases_block():
    with st.container(border=True):
        sec("Recent Purchases")
        purchases_df = fetch_purchases(limit=15)
        if purchases_df.empty:
            st.info("No purchase records yet.")
        else:
            st.dataframe(_purchase_display(purchases_df)[0], hide_index=True, height=280, **STRETCH)


def _purchase_display(df: pd.DataFrame):
    """Returns (display_df, columns_used)."""
    recent = df.copy()
    recent["purchase_date"] = recent["purchase_date"].dt.strftime("%Y-%m-%d")

    def describe_unit(row):
        if row.get("purchase_unit") == "Box" and pd.notna(row.get("box_quantity")) and pd.notna(row.get("units_per_box")):
            return f"{int(row['box_quantity'])} Box × {int(row['units_per_box'])}"
        return "Loose Pcs"

    recent["Purchased As"] = recent.apply(describe_unit, axis=1)
    recent["Expiry"] = recent["expiry_date"].apply(fmt_exp) if "expiry_date" in recent.columns else "-"
    recent["Invoice"] = recent["invoice_no"].fillna("-") if "invoice_no" in recent.columns else "-"
    recent = recent.rename(columns={
        "purchase_date": "Date", "medicine_name": "Medicine", "medicine_type": "Type", "supplier_name": "Supplier",
        "quantity": "Total Pcs", "payment_type": "Payment", "total_amount": "Total (TK)",
        "paid_amount": "Paid (TK)", "due_amount": "Due (TK)"})
    cols = ["Invoice", "Date", "Medicine", "Type", "Supplier", "Purchased As", "Expiry", "Total Pcs", "Payment",
            "Total (TK)", "Paid (TK)", "Due (TK)"]
    return recent[cols], cols


# =================================================================================
# PAGE: PURCHASE HISTORY (+ PDF)
# =================================================================================
def render_purchase_history():
    f1, f2, f3 = st.columns([1, 1, 2])
    with f1:
        start_date = st.date_input("From Date", value=date.today() - timedelta(days=30), key="ph_from")
    with f2:
        end_date = st.date_input("To Date", value=date.today(), key="ph_to")
    with f3:
        q = st.text_input("🔍 Search medicine / supplier", key="ph_q")

    df = fetch_purchases_range(start_date.isoformat(), end_date.isoformat())
    if df.empty:
        st.info("No purchases found in this date range.")
        return
    if q.strip():
        m = df["medicine_name"].astype(str).str.contains(q, case=False, regex=False) | \
            df["supplier_name"].astype(str).str.contains(q, case=False, regex=False)
        df = df[m]
        if df.empty:
            st.info("No matching purchases.")
            return

    m1, m2, m3 = st.columns(3)
    m1.metric("Total Purchase", fmt_money(df["total_amount"].sum()))
    m2.metric("Paid", fmt_money(df["paid_amount"].sum()))
    m3.metric("Due (Credit)", fmt_money(df["due_amount"].sum()))

    show, cols = _purchase_display(df)
    st.dataframe(show, hide_index=True, height=430, **STRETCH)
    pdf_rows = [[r["Date"], r["Medicine"], r["Supplier"], r["Expiry"], str(int(r["Total Pcs"])), r["Payment"],
                 tk(r["Total (TK)"]), tk(r["Due (TK)"])] for _, r in show.iterrows()]
    pdf_button("🖨️ Download Purchase Report (PDF)", "pdf_purchase_hist", f"purchases_{start_date}_{end_date}.pdf",
               "Purchase Report", ["Date", "Medicine", "Supplier", "Expiry", "Pcs", "Payment", "Total", "Due"], pdf_rows,
               [1.1, 2.6, 2, 0.9, 0.7, 0.9, 1.3, 1.2], subtitle=f"{start_date} to {end_date}", landscape_mode=True,
               summary=[f"Total: {tk(df['total_amount'].sum())}   Paid: {tk(df['paid_amount'].sum())}   Due: {tk(df['due_amount'].sum())}"],
               right_cols=[4, 6, 7])


# =================================================================================
# PAGE: OPENING STOCK ENTRY  (row style, Box / Pcs, expiry, no supplier effect)
# =================================================================================
def render_opening_stock():
    st.caption(
        "For stock you ALREADY have in the shop before starting this app. It only adds to inventory — "
        "NO purchase record, NO supplier due."
    )
    settings = fetch_settings()
    locked = settings.get("opening_locked") == "1"

    if locked:
        st.warning("🔒 Opening stock is LOCKED (entry finished). To unlock: Settings → Opening Stock Lock.")
    else:
        tab_single, tab_import = st.tabs(["✍️ Single Entry", "📄 Excel / CSV Import (many items)"])

        with tab_single:
            if "open_ver" not in st.session_state:
                st.session_state.open_ver = 0
            ov = st.session_state.open_ver

            with st.container(border=True):
                sec("Item Details")
                r1 = st.columns([2.4, 1.2, 1.2])
                with r1[0]:
                    is_new, name, hint = medicine_picker(f"op_{ov}")
                default_idx = next((i for i, t in enumerate(MEDICINE_TYPES) if hint and t.lower().startswith(hint.lower()[:4])), 0)
                with r1[1]:
                    mtype = st.selectbox("Medicine Type *", MEDICINE_TYPES, index=default_idx,
                                         key=f"op_type_{ov}_{zlib.crc32((name or '').encode())}")
                with r1[2]:
                    entry_date = st.date_input("Stock as of date", value=date.today(), key=f"op_date_{ov}")

                unit, boxes, ppb, total_pcs = qty_row(f"op_{ov}")
                batch_no, expiry, _ = expiry_row(f"op_{ov}")

                r3 = st.columns([1, 1, 1.4, 1, 1.6])
                with r3[0]:
                    cost = st.number_input("Cost / pc", min_value=0.0, value=0.0, step=0.5, key=f"op_cost_{ov}")
                with r3[1]:
                    sale_p = st.number_input("Sale / pc", min_value=0.0, value=0.0, step=0.5, key=f"op_sale_{ov}")
                with r3[2]:
                    spacer_line()
                    save_open = st.button("💾 Save Opening Stock", type="primary", key=f"op_save_{ov}", **STRETCH)
                with r3[3]:
                    spacer_line()
                    if st.button("↺ Clear", key=f"op_refresh_{ov}", **STRETCH):
                        st.session_state.open_ver += 1
                        st.rerun()

            if save_open:
                if not name:
                    st.error("⚠️ Please select an existing medicine, or type a name under '🆕 Add a Brand-New Medicine'.")
                else:
                    existing_before = find_medicine_by_name(name)
                    prev_stock = int(existing_before["stock"]) if existing_before else 0
                    ok, msg = save_or_restock(name, total_pcs, mtype, sale_p)
                    if ok:
                        update_avg_cost(name, total_pcs, cost, prev_stock)
                        insert_opening_rows([{
                            "entry_date": entry_date.isoformat(), "medicine_name": name, "medicine_type": mtype,
                            "quantity": total_pcs, "cost_price": cost, "sale_price": sale_p,
                            "purchase_unit": "Box" if unit == UNIT_BOX else "Pcs",
                            "box_quantity": boxes, "units_per_box": ppb,
                            "expiry_date": expiry.isoformat() if expiry else None, "batch_no": batch_no}])
                        if expiry:
                            add_batches([batch_row(name, mtype, batch_no, expiry, total_pcs, "Opening")])
                        clear_data_caches()
                        st.session_state.open_ver += 1
                        st.session_state.open_flash = f"✅ Opening stock saved. {msg}" + (
                            f" Expiry: {expiry.strftime('%m/%Y')}" if expiry else "")
                        st.rerun()

            if st.session_state.get("open_flash"):
                st.success(st.session_state.pop("open_flash"))

        with tab_import:
            st.caption(
                "Columns: **name** + either **quantity** (pcs) OR **boxes + pcs_per_box**. "
                "Optional: medicine_type, cost_price, sale_price, **expiry** (MM/YYYY or a date), batch_no."
            )
            template = pd.DataFrame([
                {"name": "Napa 500mg", "medicine_type": "Tablet", "boxes": 10, "pcs_per_box": 100, "quantity": "",
                 "cost_price": 1.2, "sale_price": 1.5, "expiry": "08/2027", "batch_no": "B123"},
                {"name": "Seclo 20mg", "medicine_type": "Capsule", "boxes": "", "pcs_per_box": "", "quantity": 350,
                 "cost_price": 4, "sale_price": 5, "expiry": "12/2026", "batch_no": ""},
            ])
            st.download_button("⬇️ Download CSV template", template.to_csv(index=False).encode("utf-8"),
                               "opening_stock_template.csv", "text/csv")
            up = st.file_uploader("Upload CSV or Excel", type=["csv", "xlsx"])
            if up is not None:
                try:
                    df = pd.read_csv(up, dtype=str) if up.name.lower().endswith(".csv") else pd.read_excel(up)
                    df.columns = [str(c).strip().lower() for c in df.columns]
                    if "name" not in df.columns:
                        st.error("⚠️ File must have a 'name' column.")
                    else:
                        df["name"] = df["name"].astype(str).str.strip()
                        for c in ("quantity", "boxes", "pcs_per_box", "cost_price", "sale_price"):
                            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0) if c in df.columns else 0
                        box_total = df["boxes"] * df["pcs_per_box"]
                        df["total_pcs"] = box_total.where(box_total > 0, df["quantity"]).astype(int)
                        df["unit"] = (box_total > 0).map({True: "Box", False: "Pcs"})
                        if "medicine_type" not in df.columns:
                            df["medicine_type"] = "Other"
                        df["medicine_type"] = df["medicine_type"].fillna("Other").astype(str)
                        raw_exp = df["expiry"] if "expiry" in df.columns else pd.Series([None] * len(df), index=df.index)
                        df["expiry_dt"] = raw_exp.apply(parse_expiry)
                        bad_exp = int((raw_exp.notna() & (raw_exp.astype(str).str.strip() != "") & df["expiry_dt"].isna()).sum())
                        if "batch_no" not in df.columns:
                            df["batch_no"] = ""
                        df["batch_no"] = df["batch_no"].fillna("").astype(str).str.strip()
                        df = df[(df["name"] != "") & (df["total_pcs"] > 0)]
                        st.write(f"**{len(df)} valid rows** ready to import "
                                 f"({int(df['expiry_dt'].notna().sum())} with expiry):")
                        if bad_exp:
                            st.warning(f"⚠️ {bad_exp} row(s) have an expiry value that could not be read — "
                                       "they will be imported WITHOUT expiry. Use MM/YYYY or YYYY-MM-DD.")
                        prev = df.copy()
                        prev["expiry"] = prev["expiry_dt"].apply(lambda d: d.strftime("%m/%Y") if d else "-")
                        st.dataframe(prev[["name", "medicine_type", "unit", "boxes", "pcs_per_box", "total_pcs",
                                           "expiry", "batch_no", "cost_price", "sale_price"]].head(50),
                                     hide_index=True, height=250, **STRETCH)
                        if st.button("✅ Import All Now", type="primary", **STRETCH):
                            with st.spinner("Importing…"):
                                items = [dict(name=r.name, qty=int(r.total_pcs), mtype=r.medicine_type,
                                              sale_price=float(r.sale_price)) for r in df.itertuples(index=False)]
                                done, failed = bulk_restock(items)
                                bulk_update_avg_cost([dict(name=r.name, added_qty=int(r.total_pcs),
                                                           unit_cost=float(r.cost_price))
                                                      for r in df.itertuples(index=False) if r.name in done])
                                log_rows, batch_rows = [], []
                                for r in df.itertuples(index=False):
                                    if r.name not in done:
                                        continue
                                    is_box = r.unit == "Box"
                                    exp = r.expiry_dt
                                    log_rows.append({
                                        "entry_date": date.today().isoformat(), "medicine_name": r.name,
                                        "medicine_type": r.medicine_type, "quantity": int(r.total_pcs),
                                        "cost_price": float(r.cost_price), "sale_price": float(r.sale_price),
                                        "purchase_unit": r.unit,
                                        "box_quantity": int(r.boxes) if is_box else None,
                                        "units_per_box": int(r.pcs_per_box) if is_box else None,
                                        "expiry_date": exp.isoformat() if exp else None,
                                        "batch_no": r.batch_no or None})
                                    if exp:
                                        batch_rows.append(batch_row(r.name, r.medicine_type, r.batch_no, exp,
                                                                    int(r.total_pcs), "Opening"))
                                insert_opening_rows(log_rows)
                                add_batches(batch_rows)
                                clear_data_caches()
                            st.success(f"✅ Imported {len(done)} items ({len(batch_rows)} expiry batches). Failed: {failed}")
                except Exception as e:
                    st.error(f"❌ Could not read file: {e}")

        if st.button("🔒 Finish & Lock Opening Stock"):
            if set_setting("opening_locked", "1"):
                st.rerun()

    with st.container(border=True):
        sec("Opening Stock Log")
        log = fetch_opening_stock()
        if log.empty:
            st.info("No opening stock entries yet.")
        else:
            def as_text(r):
                if r.get("purchase_unit") == "Box" and pd.notna(r.get("box_quantity")) and pd.notna(r.get("units_per_box")):
                    return f"{int(r['box_quantity'])} Box × {int(r['units_per_box'])}"
                return "Loose Pcs"

            log = log.copy()
            log["Entered As"] = log.apply(as_text, axis=1)
            log["Expiry"] = log["expiry_date"].apply(fmt_exp) if "expiry_date" in log.columns else "-"
            log = log.rename(columns={
                "entry_date": "As of", "medicine_name": "Medicine", "medicine_type": "Type",
                "quantity": "Total Pcs", "cost_price": "Cost/pc", "sale_price": "Sale/pc"})
            st.dataframe(log[["As of", "Medicine", "Type", "Entered As", "Expiry", "Total Pcs", "Cost/pc", "Sale/pc"]],
                         hide_index=True, height=260, **STRETCH)


# =================================================================================
# PAGE: SALES / POS  (left: editable cart, right: billing)
# =================================================================================
def render_sales():
    if "sale_cart" not in st.session_state:
        st.session_state.sale_cart = []
    if "last_voucher" not in st.session_state:
        st.session_state.last_voucher = None
    if "cart_ver" not in st.session_state:
        st.session_state.cart_ver = 0

    meds_df = fetch_medicines()
    in_stock_df = meds_df[meds_df["stock"] > 0] if not meds_df.empty else pd.DataFrame()
    if in_stock_df.empty:
        st.warning("⚠️ No medicines currently in stock. Add stock via Purchase Entry or Opening Stock Entry.")
        return

    nearest_valid, expired_qty = expiry_maps()

    # ---------- Row: search + product + qty + price + Add ----------
    with st.container(border=True):
        sec("Item Details")
        colp0, colp1, colp2, colp3, colp4 = st.columns([1.6, 3, 0.8, 1, 0.8])
        with colp0:
            search_term = st.text_input("Search product", placeholder="type name...")
        filtered_df = in_stock_df
        if search_term:
            filtered_df = in_stock_df[in_stock_df["name"].str.contains(search_term, case=False, na=False, regex=False)]

        if filtered_df.empty:
            st.warning("⚠️ No matching product found.")
        else:
            filtered_df = filtered_df.head(300)
            med_options = []
            for nm, tp, stk in zip(filtered_df["name"], filtered_df["medicine_type"], filtered_df["stock"]):
                lab = nm + (f"  ·  {tp}" if tp else "") + f"  ·  Avail: {int(stk)} pcs"
                if expired_qty.get(nm, 0) > 0:
                    lab += "  ·  ⛔ has expired stock"
                elif nm in nearest_valid:
                    lab += f"  ·  Exp {fmt_exp(nearest_valid[nm])}"
                med_options.append(lab)
            id_lookup = dict(zip(med_options, filtered_df["id"]))
            with colp1:
                selected_option = st.selectbox("Product", med_options)
            selected_id = id_lookup[selected_option]
            selected_row = filtered_df[filtered_df["id"] == selected_id].iloc[0]
            available_stock = int(selected_row["stock"])
            default_price = float(selected_row.get("sale_price", 0) or 0)
            with colp2:
                add_qty = st.number_input("Qty", min_value=1, max_value=available_stock, value=1, step=1,
                                          key=f"add_qty_{selected_id}")
            with colp3:
                add_price = st.number_input("Price / unit", min_value=0.0, value=default_price, step=0.5, format="%.2f",
                                            key=f"add_price_{selected_id}")
            with colp4:
                spacer_line()
                add_clicked = st.button("➕ Add", **STRETCH)

            n_exp = int(expired_qty.get(selected_row["name"], 0))
            if n_exp > 0:
                st.warning(f"⛔ {n_exp} pcs of **{selected_row['name']}** are EXPIRED. Don't sell them — "
                           "write them off from ⏰ Near Expiry Report.")

            if add_clicked:
                if add_price <= 0:
                    st.error("⚠️ Please enter a sales price greater than 0.")
                else:
                    cost_price = float(selected_row.get("avg_cost", 0) or 0)
                    idx = next((i for i, it in enumerate(st.session_state.sale_cart) if it["medicine_id"] == int(selected_id)), None)
                    if idx is not None:
                        new_qty = st.session_state.sale_cart[idx]["qty"] + int(add_qty)
                        if new_qty > available_stock:
                            st.error("⚠️ Total quantity in cart exceeds available stock.")
                        else:
                            st.session_state.sale_cart[idx].update(
                                qty=new_qty, unit_price=float(add_price), subtotal=new_qty * float(add_price), stock=available_stock)
                            st.session_state.cart_ver += 1
                            st.rerun()
                    else:
                        st.session_state.sale_cart.append({
                            "medicine_id": int(selected_id), "name": selected_row["name"],
                            "medicine_type": selected_row.get("medicine_type", ""), "qty": int(add_qty),
                            "unit_price": float(add_price), "subtotal": int(add_qty) * float(add_price),
                            "cost_price": cost_price, "stock": available_stock})
                        st.session_state.cart_ver += 1
                        st.rerun()

    if not st.session_state.sale_cart:
        st.info("Cart is empty. Search and add products above.")
        _show_last_voucher()
        return

    cart = st.session_state.sale_cart
    left, right = st.columns([1.5, 1])

    # ---------------- LEFT: editable cart ----------------
    with left:
        with st.container(border=True):
            sec("Cart  (edit Qty / Price directly · tick Remove to delete)")
            view = pd.DataFrame(cart).rename(columns={"name": "Product", "medicine_type": "Type", "qty": "Qty",
                                                      "unit_price": "Unit Price"})[["Product", "Type", "Qty", "Unit Price"]]
            view["Remove"] = False
            edited = st.data_editor(
                view, key=f"cart_editor_{st.session_state.cart_ver}", hide_index=True, disabled=["Product", "Type"],
                height=280, num_rows="fixed",
                column_config={
                    "Qty": st.column_config.NumberColumn(min_value=1, step=1),
                    "Unit Price": st.column_config.NumberColumn(min_value=0.0, step=0.5, format="%.2f"),
                    "Remove": st.column_config.CheckboxColumn("Remove"),
                }, **STRETCH)

            new_cart, removed, clamped = [], False, False
            for i, row in edited.iterrows():
                if bool(row["Remove"]):
                    removed = True
                    continue
                it = dict(cart[i])
                qty = int(row["Qty"]) if pd.notna(row["Qty"]) else it["qty"]
                cap = int(it.get("stock", qty))
                if qty > cap:
                    qty, clamped = cap, True
                qty = max(qty, 1)
                price = float(row["Unit Price"]) if pd.notna(row["Unit Price"]) else it["unit_price"]
                it.update(qty=qty, unit_price=price, subtotal=round(qty * price, 2))
                new_cart.append(it)
            st.session_state.sale_cart = new_cart
            if removed:
                st.session_state.cart_ver += 1
                st.rerun()
            if clamped:
                st.caption("ℹ️ Quantity was limited to the available stock.")
            if any(it["unit_price"] <= 0 for it in new_cart):
                st.warning("⚠️ Some items have price 0 — set a price before checkout.")
            if st.button("🗑️ Clear Cart"):
                st.session_state.sale_cart = []
                st.session_state.cart_ver += 1
                st.rerun()

    cart = st.session_state.sale_cart
    subtotal_amount = float(sum(it["subtotal"] for it in cart))
    zero_price = any(it["unit_price"] <= 0 for it in cart)

    # ---------------- RIGHT: billing ----------------
    with right:
        with st.container(border=True):
            sec("Billing")
            try:
                default_pct = float(fetch_settings().get("default_discount_pct", DEFAULT_DISCOUNT_PCT))
            except Exception:
                default_pct = DEFAULT_DISCOUNT_PCT

            cust_ledger = fetch_customer_ledger()
            cust_options = [NEW_CUSTOMER_LABEL]
            cust_lookup = {}
            if not cust_ledger.empty:
                for r in cust_ledger.itertuples(index=False):
                    lbl = f"{r.customer_name} ({r.customer_phone})  |  Due: {fmt_money(r.total_due)}"
                    cust_options.append(lbl)
                    cust_lookup[lbl] = (r.customer_name, r.customer_phone, float(r.total_due))

            b1, b2 = st.columns([1.5, 1])
            with b1:
                chosen_cust = st.selectbox("Customer", cust_options)
            with b2:
                discount_pct = st.number_input(
                    "Discount (%)", min_value=0.0, max_value=100.0, value=default_pct, step=0.5, format="%.2f",
                    key="sale_discount_pct", help="Default from Settings. Change for this sale if needed.")

            prev_due = 0.0
            if chosen_cust == NEW_CUSTOMER_LABEL:
                c1, c2 = st.columns(2)
                with c1:
                    customer_name = st.text_input("Customer Name", key="sale_cust_name")
                with c2:
                    customer_phone = st.text_input("Customer Phone", key="sale_cust_phone")
            else:
                customer_name, customer_phone, prev_due = cust_lookup[chosen_cust]
                c1, c2 = st.columns(2)
                with c1:
                    st.text_input("Customer Name", value=customer_name, disabled=True, key="sale_cust_name_ro")
                with c2:
                    st.text_input("Customer Phone", value=customer_phone, disabled=True, key="sale_cust_phone_ro")

            discount_amount = round(subtotal_amount * discount_pct / 100, 2)
            payable_amount = round(subtotal_amount - discount_amount, 2)

            p1, p2 = st.columns(2)
            with p1:
                payment_mode = st.selectbox("Payment Mode *", ["Cash", "Credit"])
            bad_paid = False
            with p2:
                if payment_mode == "Credit":
                    paid_amount = st.number_input("Paid Now (TK)", min_value=0.0, value=0.0, step=1.0)
                    bad_paid = paid_amount > payable_amount
                else:
                    paid_amount = payable_amount
                    st.text_input("Paid Now (TK)", value=f"{payable_amount:,.2f}", disabled=True, key="sale_paid_ro")
            due_amount = max(round(payable_amount - paid_amount, 2), 0.0)
            if bad_paid:
                st.error("⚠️ Paid amount is more than the total payable.")

            m1, m2 = st.columns(2)
            m1.metric("Subtotal", fmt_money(subtotal_amount))
            m2.metric(f"Discount ({discount_pct:g}%)", fmt_money(discount_amount))
            m3, m4 = st.columns(2)
            m3.metric("Total Payable", fmt_money(payable_amount))
            m4.metric("Due (this sale)", fmt_money(due_amount))

            if prev_due > 0:
                msg = f"📌 Previous due: **{fmt_money(prev_due)}**"
                if payment_mode == "Credit" and due_amount > 0:
                    msg += f"  →  New total due after this sale: **{fmt_money(prev_due + due_amount)}**"
                st.warning(msg)

            need_cust = payment_mode == "Credit" and due_amount > 0 and (not str(customer_name).strip() or not str(customer_phone).strip())
            if need_cust:
                st.warning("⚠️ Customer Name & Phone are required for Credit sales with a due amount.")

            if st.button("✅ Checkout & Generate Voucher", type="primary",
                         disabled=need_cust or bad_paid or zero_price, **STRETCH):
                fetch_medicines.clear()  # always check the freshest stock at checkout
                latest = fetch_medicines()
                stock_ok, updates = True, []
                for item in st.session_state.sale_cart:
                    row = latest[latest["id"] == item["medicine_id"]]
                    if row.empty:
                        st.error(f"⚠️ '{item['name']}' no longer exists in inventory.")
                        stock_ok = False
                        continue
                    new_stock = int(row.iloc[0]["stock"]) - item["qty"]
                    if new_stock < 0:
                        st.error(f"⚠️ Not enough stock for {item['name']} (available {int(row.iloc[0]['stock'])}).")
                        stock_ok = False
                        continue
                    updates.append((item["medicine_id"], new_stock, item["unit_price"], item["name"], item["qty"]))

                if stock_ok:
                    voucher_no = generate_voucher_no()
                    cname = str(customer_name).strip() or "Walk-in Customer"
                    cphone = str(customer_phone).strip()
                    sale_payload = {
                        "voucher_no": voucher_no,
                        "sale_date": date.today().isoformat(),
                        "customer_name": cname,
                        "customer_phone": cphone,
                        "items": json.dumps([{k: v for k, v in it.items() if k != "stock"} for it in st.session_state.sale_cart]),
                        "subtotal": subtotal_amount,
                        "discount": discount_amount,
                        "total_amount": payable_amount,
                        "payment_mode": payment_mode,
                        "paid_amount": paid_amount if payment_mode == "Credit" else payable_amount,
                        "due_amount": due_amount if payment_mode == "Credit" else 0.0,
                    }
                    # Sale record FIRST: if it fails, stock is untouched.
                    saved = insert_sale(sale_payload)
                    if not saved:
                        st.error("❌ Sale was NOT saved. Stock is unchanged — please try again.")
                    else:
                        for mid, ns, price, nm, qty in updates:
                            update_medicine_stock(mid, ns, sale_price=price)   # also remembers the price
                            consume_batches(nm, qty)
                        ledger_ok = True
                        if payment_mode == "Credit" and due_amount > 0:
                            new_bal = upsert_customer_due(cname, cphone, due_amount)   # same phone => due adds up
                            ledger_ok = new_bal is not None
                            if ledger_ok:
                                log_txn("customer", cphone, cname, "Credit Sale", voucher_no, due_amount, 0, new_bal)
                        clear_data_caches()
                        st.session_state.last_voucher = sale_payload
                        st.session_state.sale_cart = []
                        st.session_state.cart_ver += 1
                        if not ledger_ok:
                            st.session_state.last_voucher_warn = True
                        st.rerun()

    _show_last_voucher()


def _show_last_voucher():
    lv = st.session_state.get("last_voucher")
    if not lv:
        return
    with st.container(border=True):
        sec("Voucher / Receipt")
        st.success(f"✅ Sale completed! Voucher No: **{lv['voucher_no']}**")
        if st.session_state.pop("last_voucher_warn", False):
            st.error("⚠️ The sale was saved but the customer due could not be updated. Please check the Customer Ledger.")
        vc1, vc2 = st.columns([1.3, 1])
        with vc1:
            st.markdown(render_voucher_html(lv), unsafe_allow_html=True)
        with vc2:
            st.download_button(
                "🖨️ Download / Print Voucher (PDF)", data=build_voucher_pdf(lv), file_name=f"{lv['voucher_no']}.pdf",
                mime="application/pdf", type="primary", key="dl_last_voucher", **STRETCH)
            if st.button("✖️ Clear Voucher Preview", **STRETCH):
                st.session_state.last_voucher = None
                st.rerun()


# =================================================================================
# PAGE: SALES RETURN  (customer returns an item from a past voucher)
# =================================================================================
def render_sales_return():
    if not table_ready("sales_returns"):
        st.error("⚠️ The `sales_returns` table is not created yet. Run this once in **Supabase → SQL Editor**, "
                "then press 🔄 Refresh Data:")
        st.code(FULL_SQL, language="sql")
        return

    st.caption("Find the original voucher, choose which items and how many pieces are being returned, "
              "and the stock will be added back automatically.")
    f1, f2 = st.columns([1, 3])
    with f1:
        lookback = st.selectbox("Search vouchers from last", [30, 60, 90, 180, 365], index=1, key="sr_lookback")
    with f2:
        q = st.text_input("🔍 Voucher No / Customer Name / Phone", key="sr_q")

    sales_df = fetch_sales((date.today() - timedelta(days=lookback)).isoformat(), date.today().isoformat())
    if sales_df.empty:
        st.info("No sales found in this period.")
        return
    if q.strip():
        m = sales_df["voucher_no"].astype(str).str.contains(q, case=False, regex=False) | \
            sales_df["customer_name"].astype(str).str.contains(q, case=False, regex=False) | \
            sales_df["customer_phone"].astype(str).str.contains(q, case=False, regex=False)
        sales_df = sales_df[m]
    if sales_df.empty:
        st.info("No matching voucher.")
        return

    opts = sales_df.apply(lambda r: f"{r['voucher_no']}  ·  {r['customer_name']}  ·  {fmt_money(r['total_amount'])}"
                          f"  ·  {r['sale_date'].strftime('%d %b %Y')}", axis=1).tolist()
    pick = st.selectbox("Select voucher", opts, index=None, placeholder="Select a voucher to return items from", key="sr_pick")
    if not pick:
        return
    sale_row = sales_df.iloc[opts.index(pick)].to_dict()
    items = _parse_items(sale_row)
    already = returned_qty_for_voucher(sale_row["voucher_no"])

    with st.container(border=True):
        sec(f"Items in Voucher {sale_row['voucher_no']}")
        rows = []
        for it in items:
            sold = int(it.get("qty", 0))
            done = int(already.get(it["name"], 0))
            left = max(sold - done, 0)
            rows.append({"Item": it["name"], "Sold Qty": sold, "Already Returned": done,
                        "Returnable": left, "Unit Price": it.get("unit_price", 0)})
        info_df = pd.DataFrame(rows)
        st.dataframe(info_df, hide_index=True, height=min(220, 60 + 34 * len(rows)), **STRETCH)

        returnable_items = [r for r in rows if r["Returnable"] > 0]
        if not returnable_items:
            st.info("Everything on this voucher has already been returned.")
            return

        sec("Select Items to Return")
        cart_key = f"ret_cart_{sale_row['voucher_no']}"
        if cart_key not in st.session_state:
            st.session_state[cart_key] = []
        edit_df = pd.DataFrame([{"Item": r["Item"], "Max": r["Returnable"], "Unit Price": r["Unit Price"], "Return Qty": 0}
                               for r in returnable_items])
        edited = st.data_editor(edit_df, hide_index=True, disabled=["Item", "Max", "Unit Price"], key=f"sr_editor_{sale_row['voucher_no']}",
                                column_config={"Return Qty": st.column_config.NumberColumn(min_value=0, step=1)}, **STRETCH)

        return_lines = []
        for _, r in edited.iterrows():
            q_ret = int(r["Return Qty"] or 0)
            if q_ret > 0:
                q_ret = min(q_ret, int(r["Max"]))
                return_lines.append({"name": r["Item"], "qty": q_ret, "unit_price": float(r["Unit Price"]),
                                    "subtotal": round(q_ret * float(r["Unit Price"]), 2)})
        total_return = round(sum(l["subtotal"] for l in return_lines), 2)

        r1, r2, r3 = st.columns([1, 1, 2])
        with r1:
            st.metric("Return Amount", fmt_money(total_return))
        with r2:
            refund_mode = st.selectbox("Refund Mode", REFUND_MODES, key="sr_mode",
                                       help="'Adjust Against Due' reduces what the customer owes you.")
        with r3:
            note = st.text_input("Reason (optional)", placeholder="e.g. wrong item, damaged pack", key="sr_note")

        if st.button("↩️ Confirm Return", type="primary", disabled=not return_lines, **STRETCH):
            payload = {
                "return_date": date.today().isoformat(), "voucher_no": sale_row["voucher_no"],
                "customer_name": sale_row.get("customer_name"), "customer_phone": sale_row.get("customer_phone"),
                "items": json.dumps(return_lines), "total_amount": total_return, "refund_mode": refund_mode,
                "note": note.strip() or None,
            }
            if insert_sales_return(payload):
                for l in return_lines:
                    med = find_medicine_by_name(l["name"])
                    if med:
                        update_medicine_stock(med["id"], int(med["stock"]) + l["qty"])
                due_msg = ""
                if refund_mode == "Adjust Against Due":
                    phone = str(sale_row.get("customer_phone") or "").strip()
                    if phone:
                        new_bal = upsert_customer_due(sale_row.get("customer_name"), phone, -total_return)
                        if new_bal is not None:
                            log_txn("customer", phone, sale_row.get("customer_name"), "Sales Return",
                                    sale_row["voucher_no"], -total_return, 0, new_bal)
                            due_msg = f" Customer due reduced by {fmt_money(total_return)}."
                clear_data_caches()
                st.success(f"✅ Return recorded: {fmt_money(total_return)} for {len(return_lines)} item(s)."
                          f" Stock updated.{due_msg}")
                st.rerun()

    with st.container(border=True):
        sec("Recent Returns")
        recent = fetch_sales_returns((date.today() - timedelta(days=30)).isoformat(), date.today().isoformat())
        if recent.empty:
            st.info("No returns in the last 30 days.")
        else:
            show = recent.copy()
            show["return_date"] = show["return_date"].dt.strftime("%Y-%m-%d")
            show = show.rename(columns={"return_date": "Date", "voucher_no": "Voucher", "customer_name": "Customer",
                                        "total_amount": "Amount (TK)", "refund_mode": "Mode"})
            st.dataframe(show[["Date", "Voucher", "Customer", "Amount (TK)", "Mode"]], hide_index=True, height=220, **STRETCH)


# =================================================================================
# PAGE: PURCHASE RETURN  (return item to supplier)
# =================================================================================
def render_purchase_return():
    if not table_ready("purchase_returns"):
        st.error("⚠️ The `purchase_returns` table is not created yet. Run this once in **Supabase → SQL Editor**, "
                "then press 🔄 Refresh Data:")
        st.code(FULL_SQL, language="sql")
        return

    st.caption("For stock being sent back to the supplier (wrong item, damaged, expired, etc). "
              "This reduces your stock AND the supplier due (or is marked as a cash refund received).")
    meds = fetch_medicines()
    supplier_options = fetch_supplier_names()

    with st.container(border=True):
        sec("Return Details")
        r1 = st.columns([1, 1.6, 1.4, 1])
        with r1[0]:
            ret_date = st.date_input("Return Date", value=date.today(), key="prt_date")
        with r1[1]:
            supplier = st.selectbox("Supplier *", supplier_options, index=None, placeholder="Select supplier", key="prt_supp")
        with r1[2]:
            q = st.text_input("🔍 Search medicine", key="prt_q")
        pool = meds if not q.strip() else meds[meds["name"].str.contains(q, case=False, regex=False)]
        pool = pool.head(80)
        with r1[3]:
            med_opts = [f"{n}  ·  stock {int(s)}" for n, s in zip(pool["name"], pool["stock"])]
            med_pick = st.selectbox("Medicine *", med_opts, index=None, placeholder="Select medicine", key="prt_med")

        row = pool.iloc[med_opts.index(med_pick)] if med_pick else None
        if row is not None and int(row["stock"]) <= 0:
            st.warning(f"⚠️ {row['name']} currently has 0 pcs in stock — nothing to return.")
            row = None
        r2 = st.columns([1, 1, 1, 2])
        with r2[0]:
            max_q = int(row["stock"]) if row is not None else 1
            qty = st.number_input("Return Qty *", min_value=1, max_value=max(max_q, 1), value=1, step=1, key="prt_qty")
        with r2[1]:
            default_cost = float(row.get("avg_cost", 0) or 0) if row is not None else 0.0
            unit_cost = st.number_input("Unit Cost (TK) *", min_value=0.0, value=default_cost, step=0.5, key="prt_cost")
        with r2[2]:
            settle_mode = st.selectbox("Settle As", SETTLE_MODES, key="prt_settle")
        with r2[3]:
            note = st.text_input("Reason (optional)", placeholder="e.g. expired, wrong item", key="prt_note")

        total = round(qty * unit_cost, 2)
        st.info(f"📌 Return value: **{fmt_money(total)}**")

        if st.button("↩️ Confirm Supplier Return", type="primary",
                     disabled=not (supplier and row is not None and total > 0), **STRETCH):
            payload = {
                "return_date": ret_date.isoformat(), "supplier_name": supplier, "medicine_name": row["name"],
                "medicine_type": row.get("medicine_type", ""), "quantity": int(qty), "unit_cost": unit_cost,
                "total_amount": total, "settle_mode": settle_mode, "note": note.strip() or None,
            }
            if insert_purchase_return(payload):
                update_medicine_stock(int(row["id"]), max(int(row["stock"]) - int(qty), 0))
                due_msg = ""
                if settle_mode == "Reduce Payable Due":
                    new_bal = upsert_supplier_due(supplier, -total)
                    if new_bal is not None:
                        log_txn("supplier", supplier, supplier, "Purchase Return", row["name"], -total, 0, new_bal)
                        due_msg = f" Supplier due reduced by {fmt_money(total)}."
                clear_data_caches()
                st.success(f"✅ Returned {int(qty)} pcs of {row['name']} to {supplier} ({fmt_money(total)}).{due_msg}")
                st.rerun()

    with st.container(border=True):
        sec("Recent Supplier Returns")
        recent = fetch_purchase_returns((date.today() - timedelta(days=30)).isoformat(), date.today().isoformat())
        if recent.empty:
            st.info("No supplier returns in the last 30 days.")
        else:
            show = recent.copy()
            show["return_date"] = show["return_date"].dt.strftime("%Y-%m-%d")
            show = show.rename(columns={"return_date": "Date", "supplier_name": "Supplier", "medicine_name": "Medicine",
                                        "quantity": "Qty", "total_amount": "Amount (TK)", "settle_mode": "Settled As"})
            st.dataframe(show[["Date", "Supplier", "Medicine", "Qty", "Amount (TK)", "Settled As"]],
                         hide_index=True, height=220, **STRETCH)


# =================================================================================
# PAGE: EXPENSES  (business running costs: salary, rent, utility bills...)
# =================================================================================
def render_expenses():
    if not table_ready("expenses"):
        st.error("⚠️ The `expenses` table is not created yet. Run this once in **Supabase → SQL Editor**, "
                "then press 🔄 Refresh Data:")
        st.code(FULL_SQL, language="sql")
        return

    with st.container(border=True):
        sec("Add Expense")
        c1, c2, c3, c4 = st.columns([1, 1.3, 2, 1])
        with c1:
            exp_date = st.date_input("Date", value=date.today(), key="exp_date")
        with c2:
            category = st.selectbox("Category *", EXPENSE_CATEGORIES, key="exp_cat")
        with c3:
            description = st.text_input("Description", placeholder="optional note", key="exp_desc")
        with c4:
            amount = st.number_input("Amount (TK) *", min_value=0.0, value=0.0, step=50.0, key="exp_amt")
        if st.button("💾 Save Expense", type="primary", key="exp_save"):
            if amount <= 0:
                st.error("⚠️ Enter an amount greater than 0.")
            elif insert_expense({"exp_date": exp_date.isoformat(), "category": category,
                                 "description": description.strip() or None, "amount": amount}):
                fetch_expenses.clear()
                st.success(f"✅ {category} — {fmt_money(amount)} saved.")
                st.rerun()

    with st.container(border=True):
        sec("Expense History")
        start, end, label = date_range_picker("exp_list", default="This Month")
        df = fetch_expenses(start.isoformat(), end.isoformat())
        if df.empty:
            st.info(f"No expenses recorded for {label}.")
            return
        m1, m2 = st.columns(2)
        m1.metric("Total Expenses", fmt_money(df["amount"].sum()))
        m2.metric("Entries", f"{len(df)}")

        by_cat = df.groupby("category", as_index=False)["amount"].sum().sort_values("amount", ascending=False)
        cat_summary_df = by_cat.rename(columns={"category": "Category", "amount": "Amount"})
        st.dataframe(cat_summary_df, hide_index=True, height=min(250, 60 + 34 * len(by_cat)), **STRETCH)

        show = df.copy()
        show["exp_date"] = show["exp_date"].dt.strftime("%Y-%m-%d")
        show = show.rename(columns={"exp_date": "Date", "category": "Category", "description": "Description", "amount": "Amount (TK)"})
        st.dataframe(show[["Date", "Category", "Description", "Amount (TK)"]], hide_index=True, height=320, **STRETCH)

        export_buttons(
            "expenses", f"expenses_{start}_{end}",
            dict(title="Expense Report", subtitle=label,
                headers=["Date", "Category", "Description", "Amount"],
                rows=[[r["Date"], r["Category"], r["Description"] or "-", tk(r["Amount (TK)"])] for _, r in show.iterrows()],
                weights=[1.2, 1.8, 3, 1.3], summary=[f"Total: {tk(df['amount'].sum())}   Entries: {len(df)}"], right_cols=[3]),
            {"Expenses": show, "By Category": cat_summary_df})

        del_opts = show.apply(lambda r: f"{r['Date']} · {r['Category']} · {fmt_money(r['Amount (TK)'])}"
                              f"  [ID:{df.loc[r.name, 'id']}]", axis=1).tolist()
        with st.expander("🗑️ Delete a wrong entry"):
            pick = st.selectbox("Select entry", del_opts, index=None, placeholder="Select", key="exp_del_pick")
            if pick and st.button("Delete", key="exp_del_btn"):
                eid = int(pick.split("[ID:")[1].replace("]", ""))
                if delete_expense(eid):
                    fetch_expenses.clear()
                    st.success("✅ Deleted.")
                    st.rerun()


# =================================================================================
# PAGE: PROFIT & LOSS
# =================================================================================
def _compute_pl(start: date, end: date):
    """Returns a dict with every line of the P&L for the given range."""
    s, e = start.isoformat(), end.isoformat()
    sales = fetch_sales(s, e)
    s_returns = fetch_sales_returns(s, e)
    p_returns = fetch_purchase_returns(s, e)
    exp_df = fetch_expenses(s, e)
    avg_cost = get_avg_cost_map()

    revenue = float(sales["total_amount"].sum()) if not sales.empty else 0.0
    discount = float(sales["discount"].sum()) if not sales.empty else 0.0
    returns_amt = float(s_returns["total_amount"].sum()) if not s_returns.empty else 0.0

    cogs = 0.0
    cogs_missing = 0
    if not sales.empty:
        for items_json in sales["items"]:
            for it in _parse_items({"items": items_json}):
                qty = float(it.get("qty", 0))
                cp = it.get("cost_price")
                if cp is None:
                    cp = avg_cost.get(it.get("name"), 0)
                    if cp:
                        cogs_missing += 1
                cogs += qty * float(cp or 0)
    # returned goods reduce COGS too (approximate using current avg cost)
    cogs_returned = 0.0
    if not s_returns.empty:
        for items_json in s_returns["items"]:
            for it in json.loads(items_json or "[]"):
                cogs_returned += float(it.get("qty", 0)) * avg_cost.get(it.get("name"), 0)
    cogs_net = max(cogs - cogs_returned, 0.0)

    net_sales = revenue - returns_amt
    gross_profit = net_sales - cogs_net

    purchase_returns_amt = float(p_returns["total_amount"].sum()) if not p_returns.empty else 0.0
    total_expenses = float(exp_df["amount"].sum()) if not exp_df.empty else 0.0
    exp_by_cat = (exp_df.groupby("category", as_index=False)["amount"].sum().sort_values("amount", ascending=False)
                 if not exp_df.empty else pd.DataFrame(columns=["category", "amount"]))

    net_profit = gross_profit - total_expenses

    return dict(revenue=revenue, discount=discount, returns_amt=returns_amt, net_sales=net_sales,
               cogs=cogs_net, gross_profit=gross_profit, total_expenses=total_expenses,
               exp_by_cat=exp_by_cat, net_profit=net_profit, purchase_returns_amt=purchase_returns_amt,
               n_sales=len(sales), n_returns=len(s_returns), cogs_estimated=cogs_missing > 0)


def render_profit_loss():
    start, end, label = date_range_picker("pl", default="This Month")
    pl = _compute_pl(start, end)

    if pl["cogs_estimated"]:
        st.caption("ℹ️ Some older sales don't have a saved cost price, so their cost is estimated using "
                  "the medicine's current average cost. New sales made from today onward will be exact.")

    st.markdown(f"#### 📈 Profit & Loss — {label}")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Gross Sales", fmt_money(pl["revenue"]))
    m2.metric("Sales Returns", fmt_money(pl["returns_amt"]))
    m3.metric("Cost of Goods Sold", fmt_money(pl["cogs"]))
    m4.metric("Gross Profit", fmt_money(pl["gross_profit"]))

    m5, m6, m7 = st.columns(3)
    m5.metric("Operating Expenses", fmt_money(pl["total_expenses"]))
    m6.metric("Net Profit", fmt_money(pl["net_profit"]),
             delta=("Profit" if pl["net_profit"] >= 0 else "Loss"), delta_color="normal" if pl["net_profit"] >= 0 else "inverse")
    margin = (pl["net_profit"] / pl["net_sales"] * 100) if pl["net_sales"] else 0.0
    m7.metric("Net Margin", f"{margin:.1f}%")

    pl_rows = [
        ("Gross Sales", pl["revenue"]), ("(-) Discounts Given", -pl["discount"]),
        ("(-) Sales Returns", -pl["returns_amt"]), ("= Net Sales", pl["net_sales"]),
        ("(-) Cost of Goods Sold", -pl["cogs"]), ("= Gross Profit", pl["gross_profit"]),
        ("(-) Operating Expenses", -pl["total_expenses"]), ("= NET PROFIT", pl["net_profit"]),
    ]
    pl_df = pd.DataFrame(pl_rows, columns=["Line", "Amount (TK)"])
    with st.container(border=True):
        sec("Statement")
        st.dataframe(pl_df, hide_index=True, height=300, **STRETCH)

    if not pl["exp_by_cat"].empty:
        with st.container(border=True):
            sec("Expenses by Category")
            ecdf = pl["exp_by_cat"].rename(columns={"category": "Category", "amount": "Amount (TK)"})
            st.dataframe(ecdf, hide_index=True, height=min(260, 60 + 34 * len(ecdf)), **STRETCH)
            _bar_chart(ecdf)

    if pl["purchase_returns_amt"] > 0:
        st.caption(f"ℹ️ {fmt_money(pl['purchase_returns_amt'])} of stock was returned to suppliers in this period "
                  "(already reduces payable due; shown here for reference only).")

    export_buttons(
        "pl", f"profit_loss_{start}_{end}",
        dict(title="Profit & Loss Statement", subtitle=label, headers=["Line", "Amount"],
            rows=[[r["Line"], tk(r["Amount (TK)"])] for _, r in pl_df.iterrows()], weights=[3, 1.5],
            summary=[f"Period: {label}", f"Sales vouchers: {pl['n_sales']}   Returns: {pl['n_returns']}"], right_cols=[1]),
        {"Profit and Loss": pl_df, "Expenses by Category": pl["exp_by_cat"].rename(
            columns={"category": "Category", "amount": "Amount"})})


def _bar_chart(df):
    """Small inline bar preview for expense-by-category (kept local, no external tool dependency)."""
    try:
        st.bar_chart(df.set_index(df.columns[0]))
    except Exception:
        pass


# =================================================================================
# PAGE: REPORTS HUB  (one place: Sales / Purchase / Stock / Ledger / P&L, PDF+Excel)
# =================================================================================
def render_reports_hub():
    st.caption("Pick a report, set the date range if needed, and download it as PDF or Excel.")
    report = st.selectbox("Report", [
        "🧾 Sales Report", "➕ Purchase Report", "📦 Stock / Inventory Report",
        "📗 Customer Ledger (Dues)", "🧾 Supplier Ledger (Dues)", "↩️ Sales Returns",
        "↩️ Purchase Returns", "💸 Expenses", "📈 Profit & Loss",
    ], key="rh_report")

    st.markdown("---")

    if report == "🧾 Sales Report":
        start, end, label = date_range_picker("rh_sales")
        df = fetch_sales(start.isoformat(), end.isoformat())
        if df.empty:
            st.info("No sales in this period.")
            return
        show = df.copy()
        show["sale_date"] = show["sale_date"].dt.strftime("%Y-%m-%d")
        show = show.rename(columns={"voucher_no": "Voucher", "sale_date": "Date", "customer_name": "Customer",
                                    "payment_mode": "Payment", "total_amount": "Total", "paid_amount": "Paid", "due_amount": "Due"})
        cols = ["Voucher", "Date", "Customer", "Payment", "Total", "Paid", "Due"]
        st.dataframe(show[cols], hide_index=True, height=380, **STRETCH)
        st.caption(f"Total: {fmt_money(df['total_amount'].sum())}  ·  Collected: {fmt_money(df['paid_amount'].sum())}"
                  f"  ·  Due: {fmt_money(df['due_amount'].sum())}  ·  Vouchers: {len(df)}")
        export_buttons("rh_sales", f"sales_report_{start}_{end}",
                       dict(title="Sales Report", subtitle=label, headers=cols,
                           rows=[[r[c] if not isinstance(r[c], float) else tk(r[c]) for c in cols] for _, r in show[cols].iterrows()],
                           weights=[2, 1, 2, 1, 1.2, 1.2, 1.2], landscape_mode=True, right_cols=[4, 5, 6]),
                       {"Sales": show[cols]})

    elif report == "➕ Purchase Report":
        start, end, label = date_range_picker("rh_purch")
        df = fetch_purchases_range(start.isoformat(), end.isoformat())
        if df.empty:
            st.info("No purchases in this period.")
            return
        show, cols = _purchase_display(df)
        st.dataframe(show, hide_index=True, height=380, **STRETCH)
        st.caption(f"Total: {fmt_money(df['total_amount'].sum())}  ·  Paid: {fmt_money(df['paid_amount'].sum())}"
                  f"  ·  Due: {fmt_money(df['due_amount'].sum())}")
        export_buttons("rh_purch", f"purchase_report_{start}_{end}",
                       dict(title="Purchase Report", subtitle=label, headers=cols,
                           rows=[[str(v) for v in r] for r in show.itertuples(index=False)],
                           weights=[1.4, 1.1, 2.2, 1, 1.8, 1.4, 1, 0.9, 1, 1.1, 1, 1], landscape_mode=True),
                       {"Purchases": show})

    elif report == "📦 Stock / Inventory Report":
        meds = fetch_medicines()
        if meds.empty:
            st.info("No inventory yet.")
            return
        nearest_valid, expired_qty = expiry_maps()
        show = meds.rename(columns={"name": "Medicine", "medicine_type": "Type", "stock": "Stock"}).copy()
        show["Nearest Expiry"] = show["Medicine"].map(lambda n: fmt_exp(nearest_valid.get(n)))
        show["Avg Cost"] = show["avg_cost"].round(2)
        show["Stock Value"] = (show["Stock"] * show["avg_cost"]).round(2)
        cols = ["Medicine", "Type", "Stock", "Avg Cost", "Stock Value", "Nearest Expiry"]
        st.dataframe(show[cols], hide_index=True, height=380, **STRETCH)
        st.caption(f"Total stock value (at average cost): {fmt_money(show['Stock Value'].sum())}")
        export_buttons("rh_stock", "stock_report",
                       dict(title="Stock / Inventory Report", headers=cols,
                           rows=[[r["Medicine"], r["Type"], str(int(r["Stock"])), tk(r["Avg Cost"]), tk(r["Stock Value"]),
                                  r["Nearest Expiry"]] for _, r in show.iterrows()],
                           weights=[2.6, 1.2, 0.8, 1, 1.2, 1.2], right_cols=[2, 3, 4]),
                       {"Stock": show[cols]})

    elif report == "📗 Customer Ledger (Dues)":
        df = fetch_customer_ledger()
        if df.empty:
            st.info("No customer ledger records.")
            return
        show = df.rename(columns={"customer_name": "Customer", "customer_phone": "Phone", "total_due": "Due"})[
            ["Customer", "Phone", "Due"]]
        st.dataframe(show, hide_index=True, height=380, **STRETCH)
        st.caption(f"Total outstanding: {fmt_money(df['total_due'].sum())}")
        export_buttons("rh_custl", "customer_ledger",
                       dict(title="Customer Ledger (Dues)", headers=["Customer", "Phone", "Due"],
                           rows=[[r["Customer"], r["Phone"], tk(r["Due"])] for _, r in show.iterrows()],
                           weights=[3, 2, 1.5], right_cols=[2]),
                       {"Customer Ledger": show})

    elif report == "🧾 Supplier Ledger (Dues)":
        df = fetch_supplier_ledger()
        if df.empty:
            st.info("No supplier ledger records.")
            return
        show = df.rename(columns={"supplier_name": "Supplier", "total_due": "Due"})[["Supplier", "Due"]]
        st.dataframe(show, hide_index=True, height=380, **STRETCH)
        st.caption(f"Total payable: {fmt_money(df['total_due'].sum())}")
        export_buttons("rh_suppl", "supplier_ledger",
                       dict(title="Supplier Ledger (Dues)", headers=["Supplier", "Due"],
                           rows=[[r["Supplier"], tk(r["Due"])] for _, r in show.iterrows()],
                           weights=[3.5, 1.5], right_cols=[1]),
                       {"Supplier Ledger": show})

    elif report == "↩️ Sales Returns":
        start, end, label = date_range_picker("rh_sret")
        df = fetch_sales_returns(start.isoformat(), end.isoformat())
        if df.empty:
            st.info("No sales returns in this period.")
            return
        show = df.copy()
        show["return_date"] = show["return_date"].dt.strftime("%Y-%m-%d")
        show = show.rename(columns={"return_date": "Date", "voucher_no": "Voucher", "customer_name": "Customer",
                                    "total_amount": "Amount", "refund_mode": "Mode"})[
            ["Date", "Voucher", "Customer", "Amount", "Mode"]]
        st.dataframe(show, hide_index=True, height=380, **STRETCH)
        st.caption(f"Total returned: {fmt_money(df['total_amount'].sum())}")
        export_buttons("rh_sret", f"sales_returns_{start}_{end}",
                       dict(title="Sales Returns", subtitle=label, headers=list(show.columns),
                           rows=[[str(v) for v in r] for r in show.itertuples(index=False)], weights=[1.1, 2, 2, 1.2, 1.5]),
                       {"Sales Returns": show})

    elif report == "↩️ Purchase Returns":
        start, end, label = date_range_picker("rh_pret")
        df = fetch_purchase_returns(start.isoformat(), end.isoformat())
        if df.empty:
            st.info("No purchase returns in this period.")
            return
        show = df.copy()
        show["return_date"] = show["return_date"].dt.strftime("%Y-%m-%d")
        show = show.rename(columns={"return_date": "Date", "supplier_name": "Supplier", "medicine_name": "Medicine",
                                    "quantity": "Qty", "total_amount": "Amount", "settle_mode": "Settled As"})[
            ["Date", "Supplier", "Medicine", "Qty", "Amount", "Settled As"]]
        st.dataframe(show, hide_index=True, height=380, **STRETCH)
        st.caption(f"Total returned to suppliers: {fmt_money(df['total_amount'].sum())}")
        export_buttons("rh_pret", f"purchase_returns_{start}_{end}",
                       dict(title="Purchase Returns", subtitle=label, headers=list(show.columns),
                           rows=[[str(v) for v in r] for r in show.itertuples(index=False)], weights=[1.1, 1.8, 2.2, 0.8, 1.2, 1.6]),
                       {"Purchase Returns": show})

    elif report == "💸 Expenses":
        start, end, label = date_range_picker("rh_exp")
        df = fetch_expenses(start.isoformat(), end.isoformat())
        if df.empty:
            st.info("No expenses in this period.")
            return
        show = df.copy()
        show["exp_date"] = show["exp_date"].dt.strftime("%Y-%m-%d")
        show = show.rename(columns={"exp_date": "Date", "category": "Category", "description": "Description", "amount": "Amount"})[
            ["Date", "Category", "Description", "Amount"]]
        st.dataframe(show, hide_index=True, height=380, **STRETCH)
        st.caption(f"Total expenses: {fmt_money(df['amount'].sum())}")
        export_buttons("rh_exp", f"expenses_{start}_{end}",
                       dict(title="Expenses", subtitle=label, headers=list(show.columns),
                           rows=[[r["Date"], r["Category"], r["Description"] or "-", tk(r["Amount"])] for _, r in show.iterrows()],
                           weights=[1.2, 1.8, 3, 1.3], right_cols=[3]),
                       {"Expenses": show})

    else:  # Profit & Loss
        render_profit_loss()


# =================================================================================
# PAGE: INVENTORY REPORT  (+ nearest expiry + PDF)
# =================================================================================
def render_inventory_reports():
    meds_df = fetch_medicines()
    if meds_df.empty:
        st.info("ℹ️ No inventory records found yet.")
        return

    nearest_valid, expired_qty = expiry_maps()
    low_stock_df = meds_df[meds_df["stock"] < 10]
    n_exp_items = len(expired_qty)
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("💊 Total Medicine Types", f"{meds_df.shape[0]}")
    m2.metric("📦 Total Stock (Pcs)", f"{int(meds_df['stock'].sum())} pcs")
    m3.metric("⚠️ Low Stock Items (<10)", f"{low_stock_df.shape[0]}")
    m4.metric("⛔ Items with Expired Stock", f"{n_exp_items}")

    def prep(df):
        out = df.rename(columns={"name": "Medicine Name", "medicine_type": "Type", "stock": "Stock (Pcs)"}).copy()
        out["Nearest Expiry"] = out["Medicine Name"].map(lambda n: fmt_exp(nearest_valid.get(n)))
        out["Expired Pcs"] = out["Medicine Name"].map(lambda n: int(expired_qty.get(n, 0)))
        return out

    def pdf_for(df, key, title):
        rows = [[r["Medicine Name"], r["Type"], str(int(r["Stock (Pcs)"])), r["Nearest Expiry"], str(int(r["Expired Pcs"]))]
                for _, r in df.iterrows()]
        pdf_button("🖨️ Download / Print PDF", key, f"{key}.pdf", title,
                   ["Medicine", "Type", "Stock (Pcs)", "Nearest Expiry", "Expired Pcs"], rows, [3.4, 1.4, 1, 1.2, 1],
                   summary=[f"Items: {len(df)}   Total pcs: {int(df['Stock (Pcs)'].sum())}"], right_cols=[2, 4])

    tab_all, tab_low = st.tabs(["📦 All Stock", f"⚠️ Low Stock ({low_stock_df.shape[0]})"])
    with tab_all:
        f1, f2 = st.columns([2, 1])
        with f1:
            search_term = st.text_input("🔍 Search by Medicine Name")
        with f2:
            type_options = ["All Types"] + sorted([t for t in meds_df["medicine_type"].unique().tolist() if t])
            type_filter = st.selectbox("Filter by Type", type_options)

        display_df = meds_df
        if search_term:
            display_df = display_df[display_df["name"].str.contains(search_term, case=False, na=False, regex=False)]
        if type_filter != "All Types":
            display_df = display_df[display_df["medicine_type"] == type_filter]

        show_df = prep(display_df.sort_values("name"))
        st.dataframe(show_df[["Medicine Name", "Type", "Stock (Pcs)", "Nearest Expiry", "Expired Pcs"]],
                     hide_index=True, height=380, **STRETCH)
        pdf_for(show_df, "inventory_report", "Inventory / Stock Report")

    with tab_low:
        if low_stock_df.empty:
            st.success("✅ No low stock items.")
        else:
            low = prep(low_stock_df.sort_values("stock"))
            st.dataframe(low[["Medicine Name", "Type", "Stock (Pcs)", "Nearest Expiry"]], hide_index=True, height=380, **STRETCH)
            pdf_for(low, "low_stock_report", "Low Stock Report (below 10 pcs)")


# =================================================================================
# PAGE: NEAR EXPIRY REPORT
# =================================================================================
def _exp_status(d):
    if d < 0:
        return "EXPIRED"
    if d <= 30:
        return "Critical (≤30d)"
    return "Near expiry"


def _exp_bg(d):
    return "#fee2e2" if d < 0 else ("#ffedd5" if d <= 30 else "#fef9c3")


def _exp_display(df: pd.DataFrame) -> pd.DataFrame:
    out = pd.DataFrame({
        "Medicine": df["medicine_name"].values,
        "Type": df["medicine_type"].values,
        "Batch": df["batch_no"].values,
        "Expiry": df["expiry_date"].dt.strftime("%d %b %Y").values,
        "Days Left": df["days_left"].astype(int).values,
        "Qty (Pcs)": df["qty_left"].values,
        "Status": [_exp_status(d) for d in df["days_left"].astype(int)],
    })
    return out


def _exp_table(df: pd.DataFrame, key: str, title: str, subtitle: str):
    if df.empty:
        st.success("✅ Nothing here.")
        return
    disp = _exp_display(df)

    def colour(row):
        return [f"background-color:{_exp_bg(row['Days Left'])}; color:#111"] * len(row)

    st.dataframe(disp.style.apply(colour, axis=1), hide_index=True, height=360, **STRETCH)
    rows = [[r["Medicine"], r["Type"], r["Batch"], r["Expiry"], str(int(r["Days Left"])), str(int(r["Qty (Pcs)"])), r["Status"]]
            for _, r in disp.iterrows()]
    pdf_button("🖨️ Download / Print PDF", f"pdf_{key}", f"{key}.pdf", title,
               ["Medicine", "Type", "Batch", "Expiry", "Days Left", "Qty (Pcs)", "Status"], rows,
               [3.2, 1.3, 1.1, 1.3, 0.8, 0.9, 1.4], subtitle=subtitle, landscape_mode=True,
               summary=[f"Batches: {len(disp)}   Total pcs: {int(disp['Qty (Pcs)'].sum())}"],
               row_bg=[_exp_bg(d) for d in disp["Days Left"]], right_cols=[4, 5])


def render_expiry():
    if not batches_ready():
        st.error("⚠️ The expiry table is not created yet. Run this once in **Supabase → SQL Editor**, then press 🔄 Refresh Data:")
        st.code(FULL_SQL, language="sql")
        return

    batches = fetch_batches()
    today = date.today()
    default_days = get_alert_days()
    windows = sorted({30, 60, 90, 120, 180, 365, default_days})

    with st.container(border=True):
        sec("Filters")
        c1, c2 = st.columns([1.2, 3])
        with c1:
            window = st.selectbox("Show expiring within", windows, index=windows.index(default_days),
                                  format_func=lambda d: f"{d} days", key="exp_window")
        with c2:
            q = st.text_input("🔍 Search medicine", key="exp_q")

    dated = pd.DataFrame()
    if not batches.empty:
        dated = batches[batches["expiry_date"].notna()].copy()
        if q.strip():
            dated = dated[dated["medicine_name"].str.contains(q, case=False, regex=False)]

    if dated.empty:
        st.info("No batches with an expiry date yet. Enter the Expiry Month/Year in **Purchase Entry** or "
                "**Opening Stock Entry** — or add an expiry for your existing stock below.")
        expired = near = dated
    else:
        expired = dated[dated["days_left"] < 0].sort_values("days_left")
        near = dated[(dated["days_left"] >= 0) & (dated["days_left"] <= window)].sort_values("days_left")
        crit = near[near["days_left"] <= 30]
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("⛔ Expired batches", f"{len(expired)}", f"{int(expired['qty_left'].sum())} pcs", delta_color="off")
        m2.metric("🔴 Within 30 days", f"{len(crit)}", f"{int(crit['qty_left'].sum())} pcs", delta_color="off")
        m3.metric(f"⏰ Within {window} days", f"{len(near)}", f"{int(near['qty_left'].sum())} pcs", delta_color="off")
        m4.metric("📅 Tracked batches", f"{len(dated)}")

        tab_near, tab_exp, tab_all = st.tabs([f"⏰ Near Expiry ({len(near)})", f"⛔ Expired ({len(expired)})", "📋 All Batches"])
        with tab_near:
            _exp_table(near, f"near_expiry_{window}d", "Near Expiry Medicine List",
                       f"Expiring within {window} days (as of {today.strftime('%d %b %Y')})")
        with tab_exp:
            _exp_table(expired, "expired_medicines", "Expired Medicine List", f"As of {today.strftime('%d %b %Y')}")
        with tab_all:
            _exp_table(dated.sort_values("days_left"), "all_batches", "All Batches by Expiry",
                       f"As of {today.strftime('%d %b %Y')}")

    # ----- Write-off -----
    with st.expander("🗑️ Write-off / Return expired or damaged stock"):
        cand = pd.concat([expired, near]) if not dated.empty else pd.DataFrame()
        if cand.empty:
            st.info("No expired or near-expiry batches to act on.")
        else:
            labels = [f"{r.medicine_name} | Batch {r.batch_no or '-'} | Exp {fmt_exp(r.expiry_date)} | Qty {r.qty_left}  [B:{r.id}]"
                      for r in cand.itertuples(index=False)]
            pick = st.selectbox("Select batch", labels, key="wo_pick")
            bid = int(pick.split("[B:")[1].replace("]", ""))
            row = cand[cand["id"] == bid].iloc[0]
            w1, w2, w3 = st.columns([1, 1.4, 1])
            with w1:
                wo_qty = st.number_input("Qty to remove", min_value=1, max_value=int(row["qty_left"]),
                                         value=int(row["qty_left"]), step=1, key=f"wo_qty_{bid}")
            with w2:
                reason = st.selectbox("Reason", ["Expired", "Returned to supplier", "Damaged"], key="wo_reason")
            with w3:
                spacer_line()
                if st.button("🗑️ Remove from stock", type="primary", key="wo_btn", **STRETCH):
                    if write_off_batch(bid, str(row["medicine_name"]), int(wo_qty)):
                        clear_data_caches()
                        st.success(f"✅ {int(wo_qty)} pcs of {row['medicine_name']} removed ({reason}). Total stock updated.")
                        st.rerun()

    # ----- Add expiry for stock you already have -----
    with st.expander("➕ Add expiry for EXISTING stock (no stock change)"):
        st.caption("Use this for stock that was entered before expiry tracking. It only records the expiry batch; "
                   "the total stock stays the same.")
        meds = fetch_medicines()
        if meds.empty:
            st.info("No medicines in inventory.")
        else:
            q2 = st.text_input("🔍 Search inventory", key="ae_q")
            pool = meds if not q2.strip() else meds[meds["name"].str.contains(q2, case=False, regex=False)]
            pool = pool.head(80)
            opts = [f"{n}  ·  stock {int(s)}" for n, s in zip(pool["name"], pool["stock"])]
            sel = st.selectbox("Medicine", opts, index=None, placeholder="Select medicine", key="ae_sel")
            if sel:
                r = pool.iloc[opts.index(sel)]
                a1, a2 = st.columns([1, 3])
                with a1:
                    ae_qty = st.number_input("Qty in this batch", min_value=1, max_value=max(int(r["stock"]), 1),
                                             value=max(int(r["stock"]), 1), step=1, key=f"ae_qty_{r['id']}")
                bn, ex, _ = expiry_row(f"ae_{r['id']}")
                if st.button("💾 Save expiry batch", type="primary", key="ae_save"):
                    if ex is None:
                        st.error("⚠️ Please choose an expiry month/year.")
                    elif add_batches([batch_row(r["name"], r["medicine_type"], bn, ex, ae_qty, "Existing stock")]):
                        st.success(f"✅ Expiry {ex.strftime('%m/%Y')} saved for {r['name']}.")
                        st.rerun()


# =================================================================================
# PAGE: CUSTOMER LEDGER  (tabs: Due List & Collect | Statement)
# =================================================================================
def render_customer_ledger():
    ledger_df = fetch_customer_ledger()

    total_outstanding = 0.0 if ledger_df.empty else float(ledger_df["total_due"].sum())
    m1, m2 = st.columns(2)
    m1.metric("🙍 Customers with Due", f"{ledger_df[ledger_df['total_due'] > 0].shape[0] if not ledger_df.empty else 0}")
    m2.metric("💳 Total Outstanding Due", fmt_money(total_outstanding))

    tab_due, tab_stmt = st.tabs(["💵 Due List & Collect", "📜 Customer Statement (History)"])

    with tab_due:
        cl, cr = st.columns([1.4, 1])
        with cl:
            sec("Customers")
            search_term = st.text_input("Search", label_visibility="collapsed", placeholder="🔍 Search by Name or Phone")
            display_df = ledger_df
            if not display_df.empty and search_term:
                mask = display_df["customer_name"].astype(str).str.contains(search_term, case=False, na=False, regex=False) | \
                       display_df["customer_phone"].astype(str).str.contains(search_term, case=False, na=False, regex=False)
                display_df = display_df[mask]
            if display_df.empty:
                st.info("No customer ledger records found.")
            else:
                st.dataframe(display_df.rename(columns={"customer_name": "Customer", "customer_phone": "Phone", "total_due": "Due (TK)"})[
                    ["Customer", "Phone", "Due (TK)"]], hide_index=True, height=330, **STRETCH)
                due_only = display_df[display_df["total_due"] > 0]
                pdf_button("🖨️ Download Due List (PDF)", "pdf_cust_due", "customer_due_list.pdf", "Customer Due List",
                           ["Customer", "Phone", "Due"],
                           [[str(r.customer_name), str(r.customer_phone), tk(r.total_due)] for r in due_only.itertuples()],
                           [3, 2, 1.5], summary=[f"Customers with due: {len(due_only)}   Total: {tk(due_only['total_due'].sum())}"],
                           right_cols=[2])
        with cr:
            sec("Receive Payment")
            due_customers = ledger_df[ledger_df["total_due"] > 0] if not ledger_df.empty else pd.DataFrame()
            if due_customers.empty:
                st.success("✅ No pending dues from any customer.")
            else:
                options = due_customers.apply(
                    lambda r: f"{r['customer_name']} ({r['customer_phone']})  |  Due: {fmt_money(r['total_due'])}  [ID:{r['id']}]",
                    axis=1).tolist()
                selected = st.selectbox("Select Customer", options)
                selected_id = int(selected.split("[ID:")[1].replace("]", ""))
                selected_row = due_customers[due_customers["id"] == selected_id].iloc[0]
                st.metric("Current Due", fmt_money(selected_row["total_due"]))
                paid_now = st.number_input("Amount Received *", min_value=0.0, max_value=float(selected_row["total_due"]),
                                           step=1.0, key=f"cust_pay_{selected_id}")
                if st.button("💵 Collect Due", type="primary", **STRETCH):
                    if paid_now <= 0:
                        st.error("⚠️ Enter a valid amount greater than 0.")
                    else:
                        new_due = float(selected_row["total_due"]) - paid_now
                        if receive_customer_payment(selected_id, new_due):
                            log_txn("customer", str(selected_row["customer_phone"]), selected_row["customer_name"],
                                    "Payment Received", "", 0, paid_now, new_due)
                            clear_data_caches()
                            st.success(f"✅ Payment of {fmt_money(paid_now)} recorded! Remaining due: {fmt_money(new_due)}")

    with tab_stmt:
        if ledger_df.empty:
            st.info("No customers yet.")
        else:
            opts = ledger_df.apply(
                lambda r: f"{r['customer_name']} ({r['customer_phone']})  |  Balance: {fmt_money(r['total_due'])}", axis=1).tolist()
            phones = ledger_df["customer_phone"].astype(str).tolist()
            pick = st.selectbox("Select Customer", opts, key="stmt_cust")
            i = opts.index(pick)
            st.metric("Current Balance (Due)", fmt_money(ledger_df.iloc[i]["total_due"]))
            statement_table("customer", phones[i], str(ledger_df.iloc[i]["customer_name"]))


# =================================================================================
# PAGE: SUPPLIER LEDGER  (tabs)
# =================================================================================
def render_supplier_ledger():
    ledger_df = fetch_supplier_ledger()

    total_outstanding = 0.0 if ledger_df.empty else float(ledger_df["total_due"].sum())
    m1, m2 = st.columns(2)
    m1.metric("🏭 Suppliers with Credit Due", f"{ledger_df[ledger_df['total_due'] > 0].shape[0] if not ledger_df.empty else 0}")
    m2.metric("💳 Total Outstanding Credit", fmt_money(total_outstanding))

    tab_due, tab_stmt, tab_hist = st.tabs(["💵 Due List & Pay", "📜 Supplier Statement", "🛍️ Credit Purchase History"])

    with tab_due:
        cl, cr = st.columns([1.4, 1])
        with cl:
            sec("Suppliers")
            search_term = st.text_input("Search supplier", label_visibility="collapsed", placeholder="🔍 Search by Supplier Name")
            display_df = ledger_df
            if not display_df.empty and search_term:
                display_df = display_df[display_df["supplier_name"].astype(str).str.contains(search_term, case=False, na=False, regex=False)]
            if display_df.empty:
                st.info("No supplier ledger records found.")
            else:
                st.dataframe(display_df.rename(columns={"supplier_name": "Supplier", "total_due": "Outstanding Due (TK)"})[
                    ["Supplier", "Outstanding Due (TK)"]], hide_index=True, height=330, **STRETCH)
                due_only = display_df[display_df["total_due"] > 0]
                pdf_button("🖨️ Download Payable List (PDF)", "pdf_supp_due", "supplier_due_list.pdf", "Supplier Payable List",
                           ["Supplier", "Outstanding Due"],
                           [[str(r.supplier_name), tk(r.total_due)] for r in due_only.itertuples()],
                           [3.5, 1.5], summary=[f"Suppliers with due: {len(due_only)}   Total: {tk(due_only['total_due'].sum())}"],
                           right_cols=[1])
        with cr:
            sec("Pay Supplier (Settle Credit)")
            due_suppliers = ledger_df[ledger_df["total_due"] > 0] if not ledger_df.empty else pd.DataFrame()
            if due_suppliers.empty:
                st.success("✅ No outstanding credit with any supplier.")
            else:
                options = due_suppliers.apply(
                    lambda r: f"{r['supplier_name']}  |  Due: {fmt_money(r['total_due'])}  [ID:{r['id']}]", axis=1).tolist()
                selected = st.selectbox("Select Supplier", options)
                selected_id = int(selected.split("[ID:")[1].replace("]", ""))
                selected_row = due_suppliers[due_suppliers["id"] == selected_id].iloc[0]
                st.metric("Current Outstanding Due", fmt_money(selected_row["total_due"]))
                paid_now = st.number_input("Amount Paid Now *", min_value=0.0, max_value=float(selected_row["total_due"]),
                                           step=1.0, key=f"supp_pay_{selected_id}")
                if st.button("💵 Record Payment", type="primary", **STRETCH):
                    if paid_now <= 0:
                        st.error("⚠️ Enter a valid amount greater than 0.")
                    else:
                        new_due = float(selected_row["total_due"]) - paid_now
                        if pay_supplier(selected_id, new_due):
                            log_txn("supplier", selected_row["supplier_name"], selected_row["supplier_name"],
                                    "Payment Made", "", 0, paid_now, new_due)
                            clear_data_caches()
                            st.success(f"✅ Payment of {fmt_money(paid_now)} recorded! Remaining due: {fmt_money(new_due)}")

    with tab_stmt:
        if ledger_df.empty:
            st.info("No suppliers yet.")
        else:
            names = ledger_df["supplier_name"].tolist()
            pick = st.selectbox("Select Supplier", names, key="stmt_supp")
            bal = float(ledger_df[ledger_df["supplier_name"] == pick].iloc[0]["total_due"])
            st.metric("Current Balance (Due)", fmt_money(bal))
            statement_table("supplier", pick, pick)

    with tab_hist:
        credit_df = fetch_purchases(limit=200, credit_only=True)
        if credit_df.empty:
            st.info("No credit purchases recorded yet.")
        else:
            credit_df = credit_df.copy()
            credit_df["purchase_date"] = credit_df["purchase_date"].dt.strftime("%Y-%m-%d")
            credit_df = credit_df.rename(columns={
                "purchase_date": "Date", "medicine_name": "Medicine", "supplier_name": "Supplier", "quantity": "Qty",
                "total_amount": "Total (TK)", "paid_amount": "Paid (TK)", "due_amount": "Due (TK)"})
            st.dataframe(credit_df[["Date", "Medicine", "Supplier", "Qty", "Total (TK)", "Paid (TK)", "Due (TK)"]],
                         hide_index=True, height=330, **STRETCH)


# =================================================================================
# PAGE: SALES HISTORY / VOUCHERS  (table left, voucher right)
# =================================================================================
def render_sales_history():
    f1, f2, f3 = st.columns([1, 1, 2])
    with f1:
        start_date = st.date_input("From Date", value=date.today() - timedelta(days=30))
    with f2:
        end_date = st.date_input("To Date", value=date.today())
    with f3:
        q = st.text_input("🔍 Search voucher / customer / phone", key="sh_q")

    if start_date > end_date:
        st.error("⚠️ 'From Date' cannot be after 'To Date'.")
        return

    filtered_df = fetch_sales(start_date.isoformat(), end_date.isoformat())
    if filtered_df.empty:
        st.info("No sales found in the selected date range.")
        return
    if q.strip():
        m = filtered_df["voucher_no"].astype(str).str.contains(q, case=False, regex=False) | \
            filtered_df["customer_name"].astype(str).str.contains(q, case=False, regex=False) | \
            filtered_df["customer_phone"].astype(str).str.contains(q, case=False, regex=False)
        filtered_df = filtered_df[m]
        if filtered_df.empty:
            st.info("No matching sales.")
            return

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Sales", fmt_money(filtered_df["total_amount"].sum()))
    m2.metric("Total Collected", fmt_money(filtered_df["paid_amount"].sum()))
    m3.metric("Total Due", fmt_money(filtered_df["due_amount"].sum()))
    m4.metric("Vouchers", f"{len(filtered_df)}")

    left, right = st.columns([1.5, 1])
    with left:
        sec("Sales")
        show_df = filtered_df.copy()
        show_df["sale_date"] = show_df["sale_date"].dt.strftime("%Y-%m-%d")
        show_df = show_df.rename(columns={
            "voucher_no": "Voucher No", "sale_date": "Date", "customer_name": "Customer", "customer_phone": "Phone",
            "payment_mode": "Payment", "total_amount": "Total (TK)", "paid_amount": "Paid (TK)", "due_amount": "Due (TK)"})
        st.dataframe(show_df[["Voucher No", "Date", "Customer", "Phone", "Payment", "Total (TK)", "Paid (TK)", "Due (TK)"]],
                     hide_index=True, height=400, **STRETCH)
        pdf_button("🖨️ Download Sales Report (PDF)", "pdf_sales_hist", f"sales_{start_date}_{end_date}.pdf", "Sales Report",
                   ["Voucher", "Date", "Customer", "Payment", "Total", "Paid", "Due"],
                   [[r["Voucher No"], r["Date"], str(r["Customer"]), r["Payment"], tk(r["Total (TK)"]), tk(r["Paid (TK)"]),
                     tk(r["Due (TK)"])] for _, r in show_df.iterrows()],
                   [2.6, 1.1, 2, 0.9, 1.3, 1.3, 1.2], subtitle=f"{start_date} to {end_date}", landscape_mode=True,
                   summary=[f"Total sales: {tk(filtered_df['total_amount'].sum())}   Collected: {tk(filtered_df['paid_amount'].sum())}"
                            f"   Due: {tk(filtered_df['due_amount'].sum())}   Vouchers: {len(filtered_df)}"],
                   right_cols=[4, 5, 6])
    with right:
        sec("View / Reprint / Download Voucher")
        voucher_options = filtered_df.apply(
            lambda r: f"{r['voucher_no']}  ·  {r['customer_name']}  ·  {fmt_money(r['total_amount'])}", axis=1).tolist()
        voucher_lookup = dict(zip(voucher_options, filtered_df["voucher_no"]))
        selected_label = st.selectbox("Select a voucher", voucher_options, label_visibility="collapsed")
        voucher_row = filtered_df[filtered_df["voucher_no"] == voucher_lookup[selected_label]].iloc[0].to_dict()
        st.download_button(
            "🖨️ Download / Print Voucher (PDF)", data=build_voucher_pdf(voucher_row), file_name=f"{voucher_row['voucher_no']}.pdf",
            mime="application/pdf", type="primary", key="dl_hist_voucher", **STRETCH)
        st.markdown(render_voucher_html(voucher_row), unsafe_allow_html=True)


# =================================================================================
# PAGE: SETTINGS
# =================================================================================
def render_settings():
    settings = fetch_settings()
    c1, c2 = st.columns(2)

    with c1:
        with st.container(border=True):
            sec("Connection Status")
            try:
                supabase.table("medicines").select("id").limit(1).execute()
                st.success("✅ Connected to Supabase successfully.")
            except Exception as e:
                st.error(f"❌ Supabase connection issue: {e}")
            checks = [
                ("Expiry batches (medicine_batches)", batches_ready()),
                ("Sales Returns (sales_returns)", table_ready("sales_returns")),
                ("Purchase Returns (purchase_returns)", table_ready("purchase_returns")),
                ("Expenses (expenses)", table_ready("expenses")),
            ]
            missing = [name for name, ok in checks if not ok]
            for name, ok in checks:
                (st.success if ok else st.warning)(f"{'✅' if ok else '⚠️'} {name}")
            if missing:
                with st.expander(f"📋 Copy setup SQL for {len(missing)} missing table(s)"):
                    st.code(FULL_SQL, language="sql")

        with st.container(border=True):
            sec("Default Sales Discount")
            try:
                cur = float(settings.get("default_discount_pct", DEFAULT_DISCOUNT_PCT))
            except Exception:
                cur = DEFAULT_DISCOUNT_PCT
            new_pct = st.number_input("Default discount % on new sales", min_value=0.0, max_value=100.0, value=cur, step=0.5)
            if st.button("💾 Save Default Discount"):
                if set_setting("default_discount_pct", str(new_pct)):
                    st.success(f"✅ Default discount set to {new_pct:g}%")

        with st.container(border=True):
            sec("Near-Expiry Alert")
            days = st.number_input("Alert me for medicines expiring within (days)", min_value=1, max_value=730,
                                   value=get_alert_days(), step=15)
            if st.button("💾 Save Alert Days"):
                if set_setting("expiry_alert_days", str(int(days))):
                    st.success(f"✅ Alert set to {int(days)} days")
                    st.rerun()

    with c2:
        with st.container(border=True):
            sec("Opening Stock Lock")
            if settings.get("opening_locked") == "1":
                st.warning("Opening stock entry is currently LOCKED.")
                if st.button("🔓 Unlock Opening Stock Entry"):
                    if set_setting("opening_locked", "0"):
                        st.rerun()
            else:
                st.info("Opening stock entry is open. Lock it from Stock → Opening Stock Entry when finished.")

        with st.container(border=True):
            sec("Master Catalogue & Cache")
            master_df = fetch_master_medicines()
            st.write(f"Catalogue entries: **{0 if master_df.empty else master_df.shape[0]}**")
            if st.button("🔄 Clear Cache & Refresh Now"):
                clear_all_caches()
                st.success("✅ Cache cleared successfully.")
                st.rerun()

    st.caption(f"Med Life Pharmacy ERP · Built with Streamlit + Supabase · {CREDIT_TEXT}")


# =================================================================================
# ROUTER  (only the selected page runs = fast)
# =================================================================================
ROUTES = {
    "🛒 Sales / POS": render_sales,
    "🧾 Sales History / Vouchers": render_sales_history,
    "↩️ Sales Return": render_sales_return,
    "➕ Purchase Entry": render_purchase,
    "📜 Purchase History": render_purchase_history,
    "↩️ Purchase Return": render_purchase_return,
    "📥 Opening Stock Entry": render_opening_stock,
    "📦 Inventory Report": render_inventory_reports,
    "⏰ Near Expiry Report": render_expiry,
    "📗 Customer Ledger": render_customer_ledger,
    "🧾 Supplier Ledger": render_supplier_ledger,
    "💸 Expenses": render_expenses,
    "📈 Profit & Loss": render_profit_loss,
    "📑 Reports Hub": render_reports_hub,
    "⚙️ Settings": render_settings,
}
ROUTES[page]()
