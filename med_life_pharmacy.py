"""
=================================================================================
 MED LIFE PHARMACY - ERP System
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
                 sales, customer_ledger

NEW / CHANGED TABLES - run this ONCE in Supabase -> SQL Editor:

   -- (a) opening stock log
   create table if not exists opening_stock (
       id            bigint generated always as identity primary key,
       entry_date    date not null default current_date,
       medicine_name text not null,
       medicine_type text,
       quantity      integer not null,
       cost_price    numeric not null default 0,
       sale_price    numeric not null default 0,
       created_at    timestamptz default now()
   );
   -- (b) Box option for opening stock (safe to run even if (a) already existed)
   alter table opening_stock add column if not exists purchase_unit text;
   alter table opening_stock add column if not exists box_quantity integer;
   alter table opening_stock add column if not exists units_per_box integer;

   -- (c) app settings (default discount %, opening stock lock)
   create table if not exists app_settings (
       key   text primary key,
       value text
   );

   -- (d) NEW: ledger history (every credit sale / credit purchase / payment)
   create table if not exists ledger_txn (
       id            bigint generated always as identity primary key,
       txn_date      date not null default current_date,
       party_type    text not null,      -- 'customer' or 'supplier'
       party_key     text not null,      -- customer phone / supplier name
       party_name    text,
       txn_type      text,               -- 'Credit Sale', 'Payment Received', 'Credit Purchase', 'Payment Made'
       ref_no        text,
       due_added     numeric not null default 0,
       paid          numeric not null default 0,
       balance_after numeric not null default 0,
       created_at    timestamptz default now()
   );
   create index if not exists idx_ledger_txn_party on ledger_txn (party_type, party_key);

   -- (e) optional speed indexes
   create index if not exists idx_medicines_name on medicines (name);
   create index if not exists idx_sales_date on sales (sale_date);
   create index if not exists idx_purchases_date on purchases (purchase_date);

NOTE: If you use Row Level Security, allow your key to read/write the new tables.
=================================================================================
"""

import io
import json
import random
from datetime import date, datetime
from xml.sax.saxutils import escape

import pandas as pd
import streamlit as st
from supabase import create_client, Client

from reportlab.lib import colors
from reportlab.lib.pagesizes import A5
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

# =================================================================================
# CUSTOM CSS - ERP STYLE (top menu, green sidebar, compact rows)
# =================================================================================
st.markdown(
    """
    <style>
        .stApp { background-color: #f4f6f8; }
        .block-container { padding-top: 0.8rem; padding-bottom: 0.8rem; }
        div[data-testid="stVerticalBlock"] { gap: 0.55rem; }
        label[data-testid="stWidgetLabel"] p { font-size: 0.78rem; font-weight: 600; margin-bottom: 0; color: #1e3a8a; }

        .erp-header {
            background: linear-gradient(90deg, #0b3d33 0%, #0f766e 60%, #14b8a6 100%);
            padding: 10px 22px; border-radius: 10px 10px 0 0; color: white;
        }
        .erp-header h1 { margin: 0; font-size: 1.3rem; font-weight: 700; color: white; }
        .erp-header p { margin: 2px 0 0 0; font-size: 0.78rem; opacity: 0.9; }

        /* ERP section heading bar (like 'Requisition Reference' / 'Item Details') */
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
# DATA ACCESS - lazy per page, only needed columns, cached
# =================================================================================
@st.cache_data(ttl=600, show_spinner=False)
def fetch_master_medicines() -> pd.DataFrame:
    try:
        res = supabase.table("master_medicines").select("brand_name,company_name,form").order("brand_name").execute()
        return pd.DataFrame(res.data)
    except Exception as e:
        st.error(f"❌ Failed to fetch master medicine catalogue: {e}")
        return pd.DataFrame()


@st.cache_data(ttl=600, show_spinner=False)
def get_master_options():
    master_df = fetch_master_medicines()
    option_map, form_hint_map, display_options = {}, {}, [NEW_CUSTOM_LABEL]
    if not master_df.empty:
        for brand, company, form in zip(
            master_df["brand_name"].fillna(""), master_df["company_name"].fillna(""), master_df["form"].fillna("")
        ):
            brand, company, form = str(brand).strip(), str(company).strip(), str(form).strip()
            if not brand:
                continue
            label = f"{brand} ({form}) - {company}" if form or company else brand
            option_map[label] = brand
            form_hint_map[label] = form
            display_options.append(label)
    return display_options, option_map, form_hint_map


@st.cache_data(ttl=60, show_spinner=False)
def fetch_medicines() -> pd.DataFrame:
    try:
        res = supabase.table("medicines").select("id,name,medicine_type,stock,created_at").order("name").execute()
        df = pd.DataFrame(res.data)
        if not df.empty:
            df["stock"] = pd.to_numeric(df["stock"], errors="coerce").fillna(0).astype(int)
            df["medicine_type"] = df["medicine_type"].fillna("") if "medicine_type" in df.columns else ""
        return df
    except Exception as e:
        st.error(f"❌ Failed to fetch inventory: {e}")
        return pd.DataFrame()


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
        res = supabase.table("supplier_ledger").select("*").order("total_due", desc=True).execute()
        df = pd.DataFrame(res.data)
        if not df.empty:
            df["total_due"] = pd.to_numeric(df["total_due"], errors="coerce").fillna(0)
        return df
    except Exception as e:
        st.error(f"❌ Failed to fetch supplier ledger: {e}")
        return pd.DataFrame()


@st.cache_data(ttl=30, show_spinner=False)
def fetch_sales(days: int = 60) -> pd.DataFrame:
    try:
        since = (date.today() - pd.Timedelta(days=days)).isoformat()
        res = (
            supabase.table("sales").select("*").gte("sale_date", since)
            .order("created_at", desc=True).limit(1000).execute()
        )
        df = pd.DataFrame(res.data)
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
        res = supabase.table("customer_ledger").select("*").order("total_due", desc=True).execute()
        df = pd.DataFrame(res.data)
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


def clear_data_caches():
    for fn in (fetch_medicines, fetch_purchases, fetch_supplier_ledger, fetch_sales, fetch_customer_ledger,
               fetch_opening_stock, fetch_supplier_names, fetch_txn):
        fn.clear()


def clear_all_caches():
    clear_data_caches()
    fetch_master_medicines.clear()
    get_master_options.clear()
    fetch_settings.clear()


# ---------------------------------------------------------------------------------
# WRITE HELPERS
# ---------------------------------------------------------------------------------
def find_medicine_by_name(name: str):
    try:
        res = supabase.table("medicines").select("id,stock").eq("name", name).limit(1).execute()
        return res.data[0] if res.data else None
    except Exception as e:
        st.error(f"❌ Failed to look up medicine: {e}")
        return None


def insert_medicine(name: str, stock: int, medicine_type: str) -> bool:
    try:
        supabase.table("medicines").insert({"name": name, "stock": stock, "medicine_type": medicine_type}).execute()
        return True
    except Exception as e:
        st.error(f"❌ Failed to save new medicine: {e}")
        return False


def update_medicine_stock(med_id, new_stock: int, medicine_type: str = None) -> bool:
    try:
        payload = {"stock": new_stock}
        if medicine_type:
            payload["medicine_type"] = medicine_type
        supabase.table("medicines").update(payload).eq("id", med_id).execute()
        return True
    except Exception as e:
        st.error(f"❌ Failed to update stock: {e}")
        return False


def save_or_restock(name: str, added_qty: int, medicine_type: str) -> tuple:
    name = name.strip()
    if not name:
        return False, "Medicine name cannot be empty."
    existing = find_medicine_by_name(name)
    if existing:
        new_stock = int(existing["stock"]) + added_qty
        if update_medicine_stock(existing["id"], new_stock, medicine_type):
            return True, f"'{name}' already existed — stock increased to {new_stock} pcs."
        return False, "Failed to update existing medicine stock."
    if insert_medicine(name, added_qty, medicine_type):
        return True, f"'{name}' added as a new item with {added_qty} pcs in stock."
    return False, "Failed to save the new medicine."


def insert_purchase(payload: dict) -> bool:
    try:
        supabase.table("purchases").insert(payload).execute()
        return True
    except Exception as e:
        st.error(f"❌ Failed to record purchase transaction: {e}")
        return False


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
    """Log opening stock. If the Box columns aren't in the table yet, retry without them."""
    try:
        supabase.table("opening_stock").insert(rows).execute()
        return True
    except Exception:
        try:
            slim = [{k: v for k, v in r.items() if k not in ("purchase_unit", "box_quantity", "units_per_box")} for r in rows]
            supabase.table("opening_stock").insert(slim).execute()
            return True
        except Exception as e:
            st.error(f"❌ Could not log opening stock (is the `opening_stock` table created?): {e}")
            return False


# =================================================================================
# FORMAT + VOUCHER HELPERS
# =================================================================================
def fmt_money(val) -> str:
    try:
        return f"৳{float(val):,.2f}"
    except Exception:
        return "৳0.00"


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
        <div><b>Voucher No:</b> {sale_row.get('voucher_no', '')}</div>
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
    """Downloadable PDF voucher ('Tk' used because the default PDF font has no ৳ glyph)."""
    def tk(v):
        try:
            return f"Tk {float(v):,.2f}"
        except Exception:
            return "Tk 0.00"

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


def statement_table(party_type: str, party_key: str):
    """Show a running-balance statement for one customer / supplier."""
    txn = fetch_txn(party_type, party_key)
    if txn.empty:
        st.info(
            "No history lines yet. History starts recording from the time this update was installed "
            "(older dues appear only in the balance). Make sure the `ledger_txn` table is created."
        )
        return
    show = txn.rename(columns={
        "txn_date": "Date", "txn_type": "Type", "ref_no": "Ref / Voucher",
        "due_added": "Due Added (TK)", "paid": "Paid (TK)", "balance_after": "Balance (TK)"})
    st.dataframe(show[["Date", "Type", "Ref / Voucher", "Due Added (TK)", "Paid (TK)", "Balance (TK)"]],
                 hide_index=True, height=300, **STRETCH)


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
    "Sales": ["🛒 Sales / POS", "🧾 Sales History / Vouchers"],
    "Purchase": ["➕ Purchase Entry"],
    "Stock": ["📥 Opening Stock Entry", "📦 Inventory Report"],
    "Ledger": ["📗 Customer Ledger", "🧾 Supplier Ledger"],
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

st.sidebar.markdown("---")
if st.sidebar.button("🔄 Refresh Data", **STRETCH):
    clear_all_caches()
    st.rerun()
st.sidebar.caption(f"🕒 {datetime.now().strftime('%d %b %Y, %I:%M %p')}")
st.sidebar.markdown(f'<div class="arj-credit">{CREDIT_TEXT}</div>', unsafe_allow_html=True)


# =================================================================================
# SHARED: Box / Pcs quantity row (used by Purchase and Opening Stock)
# =================================================================================
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


# =================================================================================
# PAGE: PURCHASE ENTRY  (row style)
# =================================================================================
def render_purchase():
    display_options, option_map, form_hint_map = get_master_options()
    supplier_options = [NEW_SUPPLIER_LABEL] + fetch_supplier_names()

    if "purchase_form_version" not in st.session_state:
        st.session_state.purchase_form_version = 0
    v = st.session_state.purchase_form_version

    with st.container(border=True):
        sec("Purchase Reference")
        r1 = st.columns([1, 1.5, 1.5, 1, 1])
        with r1[0]:
            pur_date = st.date_input("Purchase Date", value=date.today(), key=f"pur_date_{v}")
        with r1[1]:
            sel_supplier = st.selectbox("Supplier / Company *", supplier_options, key=f"pur_supp_{v}")
        with r1[2]:
            custom_supplier = st.text_input(
                "New Supplier Name", placeholder="only if new supplier", key=f"pur_custom_supp_{v}",
                disabled=sel_supplier != NEW_SUPPLIER_LABEL)
        with r1[3]:
            payment_type = st.selectbox("Payment Type *", ["Cash", "Credit"], key=f"pur_pay_{v}")
        with r1[4]:
            total_amount = st.number_input("Total Amount (TK) *", min_value=0.0, value=0.0, step=1.0, key=f"pur_total_{v}")

        r2 = st.columns([1, 1, 3])
        with r2[0]:
            if payment_type == "Credit":
                paid_amount = st.number_input(
                    "Paid Now (TK)", min_value=0.0, max_value=float(total_amount), value=0.0, step=1.0, key=f"pur_paid_{v}")
            else:
                paid_amount = total_amount
                st.text_input("Paid Now (TK)", value=f"{total_amount:,.2f}", disabled=True, key=f"pur_paid_auto_{v}")
        due_amount = max(total_amount - paid_amount, 0.0)
        with r2[1]:
            st.text_input("Due (TK)", value=f"{due_amount:,.2f}", disabled=True, key=f"pur_due_auto_{v}")

    with st.container(border=True):
        sec("Item Details")
        r3 = st.columns([2, 1.6, 1.2, 1.4])
        with r3[0]:
            sel_med = st.selectbox("Medicine *", display_options, key=f"pur_med_{v}")
        with r3[1]:
            custom_name = st.text_input(
                "New Medicine Name", placeholder="only if new medicine", key=f"pur_custom_name_{v}",
                disabled=sel_med != NEW_CUSTOM_LABEL)
        hint = form_hint_map.get(sel_med, "")
        default_idx = next((i for i, t in enumerate(MEDICINE_TYPES) if hint and t.lower().startswith(hint.lower()[:4])), 0)
        with r3[2]:
            medicine_type = st.selectbox("Medicine Type *", MEDICINE_TYPES, index=default_idx, key=f"pur_type_{v}")
        with r3[3]:
            custom_type = st.text_input(
                "Specify Type", placeholder="if 'Other'", key=f"pur_custom_type_{v}", disabled=medicine_type != "Other")

        unit, box_quantity, units_per_box, total_pcs = qty_row(f"pur_{v}")

        b1, b2, b3 = st.columns([1.2, 1, 3])
        with b1:
            save_clicked = st.button("💾 Save Purchase", type="primary", key=f"pur_submit_{v}", **STRETCH)
        with b2:
            if st.button("↺ Refresh", key=f"pur_refresh_{v}", **STRETCH):
                st.session_state.purchase_form_version += 1
                st.rerun()
        with b3:
            if payment_type == "Credit" and due_amount > 0:
                st.info(f"📌 Credit due for this purchase: **{fmt_money(due_amount)}**")

    if save_clicked:
        final_name = custom_name.strip() if sel_med == NEW_CUSTOM_LABEL else option_map.get(sel_med, sel_med)
        final_type = custom_type.strip() if medicine_type == "Other" else medicine_type
        final_supplier = custom_supplier.strip() if sel_supplier == NEW_SUPPLIER_LABEL else sel_supplier

        errors = []
        if not final_name:
            errors.append("Medicine name is required.")
        if medicine_type == "Other" and not final_type:
            errors.append("Please specify the custom medicine type.")
        if not final_supplier:
            errors.append("Supplier / company name is required.")
        if total_pcs <= 0:
            errors.append("Quantity must be greater than 0.")
        if total_amount <= 0:
            errors.append("Total purchase amount must be greater than 0.")
        if errors:
            for err in errors:
                st.error(f"⚠️ {err}")
            st.stop()

        stock_ok, stock_msg = save_or_restock(final_name, total_pcs, final_type)
        purchase_ok = insert_purchase({
            "purchase_date": pur_date.isoformat(),
            "medicine_name": final_name,
            "medicine_type": final_type,
            "supplier_name": final_supplier,
            "purchase_unit": "Box" if unit == UNIT_BOX else "Pcs",
            "box_quantity": box_quantity,
            "units_per_box": units_per_box,
            "quantity": total_pcs,
            "payment_type": payment_type,
            "total_amount": total_amount,
            "paid_amount": paid_amount if payment_type == "Credit" else total_amount,
            "due_amount": due_amount if payment_type == "Credit" else 0.0,
        })
        ledger_ok = True
        if payment_type == "Credit" and due_amount > 0:
            new_bal = upsert_supplier_due(final_supplier, due_amount)
            ledger_ok = new_bal is not None
            if ledger_ok:
                log_txn("supplier", final_supplier, final_supplier, "Credit Purchase", final_name,
                        due_amount, 0, new_bal)

        if stock_ok and purchase_ok and ledger_ok:
            clear_data_caches()
            st.session_state.purchase_form_version += 1
            st.session_state.purchase_flash = (
                f"✅ {stock_msg}  |  🧾 {payment_type} — {total_pcs} pcs — Total: {fmt_money(total_amount)}"
                + (f" | Due: {fmt_money(due_amount)}" if payment_type == "Credit" and due_amount > 0 else "")
            )
            st.rerun()
        else:
            st.error("❌ Something went wrong while saving. Please check the errors above and try again.")

    if st.session_state.get("purchase_flash"):
        st.success(st.session_state.pop("purchase_flash"))

    with st.container(border=True):
        sec("Recent Purchases")
        purchases_df = fetch_purchases(limit=10)
        if purchases_df.empty:
            st.info("No purchase records yet.")
        else:
            recent = purchases_df.copy()
            recent["purchase_date"] = recent["purchase_date"].dt.strftime("%Y-%m-%d")

            def describe_unit(row):
                if row.get("purchase_unit") == "Box" and pd.notna(row.get("box_quantity")) and pd.notna(row.get("units_per_box")):
                    return f"{int(row['box_quantity'])} Box × {int(row['units_per_box'])}"
                return "Loose Pcs"

            recent["Purchased As"] = recent.apply(describe_unit, axis=1)
            recent = recent.rename(columns={
                "purchase_date": "Date", "medicine_name": "Medicine", "medicine_type": "Type", "supplier_name": "Supplier",
                "quantity": "Total Pcs", "payment_type": "Payment", "total_amount": "Total (TK)",
                "paid_amount": "Paid (TK)", "due_amount": "Due (TK)"})
            st.dataframe(
                recent[["Date", "Medicine", "Type", "Supplier", "Purchased As", "Total Pcs", "Payment", "Total (TK)", "Paid (TK)", "Due (TK)"]],
                hide_index=True, height=250, **STRETCH)


# =================================================================================
# PAGE: OPENING STOCK ENTRY  (row style, Box / Pcs, no supplier effect)
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
        display_options, option_map, form_hint_map = get_master_options()
        tab_single, tab_import = st.tabs(["✍️ Single Entry", "📄 Excel / CSV Import (many items)"])

        with tab_single:
            if "open_ver" not in st.session_state:
                st.session_state.open_ver = 0
            ov = st.session_state.open_ver

            with st.container(border=True):
                sec("Item Details")
                r1 = st.columns([2, 1.6, 1.2, 1.2])
                with r1[0]:
                    sel = st.selectbox("Medicine *", display_options, key=f"op_med_{ov}")
                with r1[1]:
                    custom = st.text_input("New Medicine Name", placeholder="only if new medicine",
                                           key=f"op_custom_{ov}", disabled=sel != NEW_CUSTOM_LABEL)
                with r1[2]:
                    mtype = st.selectbox("Medicine Type *", MEDICINE_TYPES, key=f"op_type_{ov}")
                with r1[3]:
                    entry_date = st.date_input("Stock as of date", value=date.today(), key=f"op_date_{ov}")

                unit, boxes, ppb, total_pcs = qty_row(f"op_{ov}")

                r3 = st.columns([1, 1, 1.2, 1, 2])
                with r3[0]:
                    cost = st.number_input("Cost / pc", min_value=0.0, value=0.0, step=0.5, key=f"op_cost_{ov}")
                with r3[1]:
                    sale_p = st.number_input("Sale / pc", min_value=0.0, value=0.0, step=0.5, key=f"op_sale_{ov}")
                with r3[2]:
                    save_open = st.button("💾 Save Opening Stock", type="primary", key=f"op_save_{ov}", **STRETCH)
                with r3[3]:
                    if st.button("↺ Refresh", key=f"op_refresh_{ov}", **STRETCH):
                        st.session_state.open_ver += 1
                        st.rerun()

            if save_open:
                name = custom.strip() if sel == NEW_CUSTOM_LABEL else option_map.get(sel, sel)
                if not name:
                    st.error("⚠️ Medicine name is required.")
                else:
                    ok, msg = save_or_restock(name, total_pcs, mtype)
                    if ok:
                        insert_opening_rows([{
                            "entry_date": entry_date.isoformat(), "medicine_name": name, "medicine_type": mtype,
                            "quantity": total_pcs, "cost_price": cost, "sale_price": sale_p,
                            "purchase_unit": "Box" if unit == UNIT_BOX else "Pcs",
                            "box_quantity": boxes, "units_per_box": ppb}])
                        clear_data_caches()
                        st.session_state.open_ver += 1
                        st.session_state.open_flash = f"✅ Opening stock saved. {msg}"
                        st.rerun()

            if st.session_state.get("open_flash"):
                st.success(st.session_state.pop("open_flash"))

        with tab_import:
            st.caption(
                "Columns: **name** + either **quantity** (pcs) OR **boxes + pcs_per_box**. "
                "Optional: medicine_type, cost_price, sale_price."
            )
            template = pd.DataFrame([
                {"name": "Napa 500mg", "medicine_type": "Tablet", "boxes": 10, "pcs_per_box": 100, "quantity": "", "cost_price": 1.2, "sale_price": 1.5},
                {"name": "Seclo 20mg", "medicine_type": "Capsule", "boxes": "", "pcs_per_box": "", "quantity": 350, "cost_price": 4, "sale_price": 5},
            ])
            st.download_button("⬇️ Download CSV template", template.to_csv(index=False).encode("utf-8"),
                               "opening_stock_template.csv", "text/csv")
            up = st.file_uploader("Upload CSV or Excel", type=["csv", "xlsx"])
            if up is not None:
                try:
                    df = pd.read_csv(up) if up.name.lower().endswith(".csv") else pd.read_excel(up)
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
                        df = df[(df["name"] != "") & (df["total_pcs"] > 0)]
                        st.write(f"**{len(df)} valid rows** ready to import:")
                        st.dataframe(df[["name", "medicine_type", "unit", "boxes", "pcs_per_box", "total_pcs", "cost_price", "sale_price"]].head(50),
                                     hide_index=True, height=250, **STRETCH)
                        if st.button("✅ Import All Now", type="primary", **STRETCH):
                            bar = st.progress(0.0)
                            log_rows, failed = [], 0
                            for i, r in enumerate(df.itertuples(index=False), start=1):
                                ok, _ = save_or_restock(r.name, int(r.total_pcs), r.medicine_type)
                                if ok:
                                    is_box = r.unit == "Box"
                                    log_rows.append({
                                        "entry_date": date.today().isoformat(), "medicine_name": r.name,
                                        "medicine_type": r.medicine_type, "quantity": int(r.total_pcs),
                                        "cost_price": float(r.cost_price), "sale_price": float(r.sale_price),
                                        "purchase_unit": r.unit,
                                        "box_quantity": int(r.boxes) if is_box else None,
                                        "units_per_box": int(r.pcs_per_box) if is_box else None})
                                else:
                                    failed += 1
                                bar.progress(i / len(df))
                            if log_rows:
                                insert_opening_rows(log_rows)
                            clear_data_caches()
                            st.success(f"✅ Imported {len(log_rows)} items. Failed: {failed}")
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
            log = log.rename(columns={
                "entry_date": "As of", "medicine_name": "Medicine", "medicine_type": "Type",
                "quantity": "Total Pcs", "cost_price": "Cost/pc", "sale_price": "Sale/pc"})
            st.dataframe(log[["As of", "Medicine", "Type", "Entered As", "Total Pcs", "Cost/pc", "Sale/pc"]],
                         hide_index=True, height=260, **STRETCH)


# =================================================================================
# PAGE: SALES / POS  (left: cart, right: billing - no scrolling)
# =================================================================================
def render_sales():
    if "sale_cart" not in st.session_state:
        st.session_state.sale_cart = []
    if "last_voucher" not in st.session_state:
        st.session_state.last_voucher = None

    meds_df = fetch_medicines()
    in_stock_df = meds_df[meds_df["stock"] > 0] if not meds_df.empty else pd.DataFrame()
    if in_stock_df.empty:
        st.warning("⚠️ No medicines currently in stock. Add stock via Purchase Entry or Opening Stock Entry.")
        return

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
            med_options = (
                filtered_df["name"]
                + filtered_df["medicine_type"].apply(lambda t: f"  ·  {t}" if t else "")
                + filtered_df["stock"].apply(lambda s: f"  ·  Avail: {int(s)} pcs")
            ).tolist()
            id_lookup = dict(zip(med_options, filtered_df["id"]))
            with colp1:
                selected_option = st.selectbox("Product", med_options)
            selected_id = id_lookup[selected_option]
            selected_row = filtered_df[filtered_df["id"] == selected_id].iloc[0]
            available_stock = int(selected_row["stock"])
            with colp2:
                add_qty = st.number_input("Qty", min_value=1, max_value=available_stock, value=1, step=1)
            with colp3:
                add_price = st.number_input("Price / unit", min_value=0.0, value=0.0, step=0.5, format="%.2f")
            with colp4:
                st.markdown("<div style='height:1.7rem'></div>", unsafe_allow_html=True)
                add_clicked = st.button("➕ Add", **STRETCH)

            if add_clicked:
                if add_price <= 0:
                    st.error("⚠️ Please enter a sales price greater than 0.")
                else:
                    idx = next((i for i, it in enumerate(st.session_state.sale_cart) if it["medicine_id"] == int(selected_id)), None)
                    if idx is not None:
                        new_qty = st.session_state.sale_cart[idx]["qty"] + int(add_qty)
                        if new_qty > available_stock:
                            st.error("⚠️ Total quantity in cart exceeds available stock.")
                        else:
                            st.session_state.sale_cart[idx].update(qty=new_qty, unit_price=float(add_price), subtotal=new_qty * float(add_price))
                            st.rerun()
                    else:
                        st.session_state.sale_cart.append({
                            "medicine_id": int(selected_id), "name": selected_row["name"],
                            "medicine_type": selected_row.get("medicine_type", ""), "qty": int(add_qty),
                            "unit_price": float(add_price), "subtotal": int(add_qty) * float(add_price)})
                        st.rerun()

    if not st.session_state.sale_cart:
        st.info("Cart is empty. Search and add products above.")
        _show_last_voucher()
        return

    cart_df = pd.DataFrame(st.session_state.sale_cart)
    subtotal_amount = float(cart_df["subtotal"].sum())

    left, right = st.columns([1.5, 1])

    # ---------------- LEFT: cart ----------------
    with left:
        with st.container(border=True):
            sec("Cart")
            st.dataframe(
                cart_df.rename(columns={"name": "Product", "medicine_type": "Type", "qty": "Qty",
                                        "unit_price": "Unit Price", "subtotal": "Subtotal"})[
                    ["Product", "Type", "Qty", "Unit Price", "Subtotal"]],
                hide_index=True, height=280, **STRETCH)
            rc1, rc2 = st.columns([3, 1])
            remove_options = [f"{it['name']} (Qty: {it['qty']})" for it in st.session_state.sale_cart]
            with rc1:
                item_to_remove = st.selectbox("Remove an item", remove_options)
            with rc2:
                st.markdown("<div style='height:1.7rem'></div>", unsafe_allow_html=True)
                if st.button("🗑️ Remove", **STRETCH):
                    st.session_state.sale_cart.pop(remove_options.index(item_to_remove))
                    st.rerun()

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
                    key="sale_discount_pct", help="Default 5%. Increase or decrease for this sale.")

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
                    st.text_input("Customer Name", value=customer_name, disabled=True)
                with c2:
                    st.text_input("Customer Phone", value=customer_phone, disabled=True)

            discount_amount = round(subtotal_amount * discount_pct / 100, 2)
            payable_amount = round(subtotal_amount - discount_amount, 2)

            p1, p2 = st.columns(2)
            with p1:
                payment_mode = st.selectbox("Payment Mode *", ["Cash", "Credit"])
            with p2:
                if payment_mode == "Credit":
                    paid_amount = st.number_input("Paid Now (TK)", min_value=0.0, max_value=float(payable_amount), value=0.0, step=1.0)
                else:
                    paid_amount = payable_amount
                    st.text_input("Paid Now (TK)", value=f"{payable_amount:,.2f}", disabled=True)
            due_amount = max(round(payable_amount - paid_amount, 2), 0.0)

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

            if st.button("✅ Checkout & Generate Voucher", type="primary", disabled=need_cust, **STRETCH):
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
                        st.error(f"⚠️ Not enough stock for {item['name']}.")
                        stock_ok = False
                        continue
                    updates.append((item["medicine_id"], new_stock))

                if stock_ok:
                    for mid, ns in updates:
                        update_medicine_stock(mid, ns)
                    voucher_no = generate_voucher_no()
                    cname = str(customer_name).strip() or "Walk-in Customer"
                    cphone = str(customer_phone).strip()
                    sale_payload = {
                        "voucher_no": voucher_no,
                        "sale_date": date.today().isoformat(),
                        "customer_name": cname,
                        "customer_phone": cphone,
                        "items": json.dumps(st.session_state.sale_cart),
                        "subtotal": subtotal_amount,
                        "discount": discount_amount,
                        "total_amount": payable_amount,
                        "payment_mode": payment_mode,
                        "paid_amount": paid_amount if payment_mode == "Credit" else payable_amount,
                        "due_amount": due_amount if payment_mode == "Credit" else 0.0,
                    }
                    saved = insert_sale(sale_payload)
                    ledger_ok = True
                    if payment_mode == "Credit" and due_amount > 0:
                        new_bal = upsert_customer_due(cname, cphone, due_amount)   # same phone => due adds up
                        ledger_ok = new_bal is not None
                        if ledger_ok:
                            log_txn("customer", cphone, cname, "Credit Sale", voucher_no, due_amount, 0, new_bal)
                    if saved and ledger_ok:
                        clear_data_caches()
                        st.session_state.last_voucher = sale_payload
                        st.session_state.sale_cart = []
                        st.rerun()
                    else:
                        st.error("❌ Sale could not be fully recorded. Please check the errors above.")

    _show_last_voucher()


def _show_last_voucher():
    lv = st.session_state.get("last_voucher")
    if not lv:
        return
    with st.container(border=True):
        sec("Voucher / Receipt")
        st.success(f"✅ Sale completed! Voucher No: **{lv['voucher_no']}**")
        vc1, vc2 = st.columns([1.3, 1])
        with vc1:
            st.markdown(render_voucher_html(lv), unsafe_allow_html=True)
        with vc2:
            st.download_button(
                "⬇️ Download Voucher (PDF)", data=build_voucher_pdf(lv), file_name=f"{lv['voucher_no']}.pdf",
                mime="application/pdf", type="primary", key="dl_last_voucher", **STRETCH)
            if st.button("✖️ Clear Voucher Preview", **STRETCH):
                st.session_state.last_voucher = None
                st.rerun()


# =================================================================================
# PAGE: INVENTORY REPORT
# =================================================================================
def render_inventory_reports():
    meds_df = fetch_medicines()
    if meds_df.empty:
        st.info("ℹ️ No inventory records found yet.")
        return

    low_stock_df = meds_df[meds_df["stock"] < 10]
    m1, m2, m3 = st.columns(3)
    m1.metric("💊 Total Medicine Types", f"{meds_df.shape[0]}")
    m2.metric("📦 Total Stock (Pcs)", f"{int(meds_df['stock'].sum())} pcs")
    m3.metric("⚠️ Low Stock Items (<10)", f"{low_stock_df.shape[0]}")

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

        show_df = display_df.rename(columns={"name": "Medicine Name", "medicine_type": "Type", "stock": "Stock (Pcs)"}).copy()
        cols = ["Medicine Name", "Type", "Stock (Pcs)"]
        if "created_at" in show_df.columns:
            show_df["Added On"] = pd.to_datetime(show_df["created_at"], errors="coerce").dt.strftime("%Y-%m-%d")
            cols.append("Added On")
        st.dataframe(show_df[cols].sort_values("Medicine Name"), hide_index=True, height=380, **STRETCH)

    with tab_low:
        if low_stock_df.empty:
            st.success("✅ No low stock items.")
        else:
            st.dataframe(
                low_stock_df.rename(columns={"name": "Medicine Name", "medicine_type": "Type", "stock": "Stock (Pcs)"})[
                    ["Medicine Name", "Type", "Stock (Pcs)"]].sort_values("Stock (Pcs)"),
                hide_index=True, height=380, **STRETCH)


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
                mask = display_df["customer_name"].str.contains(search_term, case=False, na=False, regex=False) | \
                       display_df["customer_phone"].str.contains(search_term, case=False, na=False, regex=False)
                display_df = display_df[mask]
            if display_df.empty:
                st.info("No customer ledger records found.")
            else:
                st.dataframe(display_df.rename(columns={"customer_name": "Customer", "customer_phone": "Phone", "total_due": "Due (TK)"})[
                    ["Customer", "Phone", "Due (TK)"]], hide_index=True, height=330, **STRETCH)
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
                paid_now = st.number_input("Amount Received *", min_value=0.0, max_value=float(selected_row["total_due"]), step=1.0)
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
            statement_table("customer", phones[i])


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
                display_df = display_df[display_df["supplier_name"].str.contains(search_term, case=False, na=False, regex=False)]
            if display_df.empty:
                st.info("No supplier ledger records found.")
            else:
                st.dataframe(display_df.rename(columns={"supplier_name": "Supplier", "total_due": "Outstanding Due (TK)"})[
                    ["Supplier", "Outstanding Due (TK)"]], hide_index=True, height=330, **STRETCH)
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
                paid_now = st.number_input("Amount Paid Now *", min_value=0.0, max_value=float(selected_row["total_due"]), step=1.0)
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
            statement_table("supplier", pick)

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
        start_date = st.date_input("From Date", value=date.today() - pd.Timedelta(days=30))
    with f2:
        end_date = st.date_input("To Date", value=date.today())

    sales_df = fetch_sales(days=max((date.today() - start_date).days + 1, 1))
    if sales_df.empty:
        st.info("No sales recorded yet.")
        return

    filtered_df = sales_df[(sales_df["sale_date"].dt.date >= start_date) & (sales_df["sale_date"].dt.date <= end_date)]
    if filtered_df.empty:
        st.info("No sales found in the selected date range.")
        return

    m1, m2, m3 = st.columns(3)
    m1.metric("Total Sales", fmt_money(filtered_df["total_amount"].sum()))
    m2.metric("Total Collected", fmt_money(filtered_df["paid_amount"].sum()))
    m3.metric("Total Due", fmt_money(filtered_df["due_amount"].sum()))

    left, right = st.columns([1.5, 1])
    with left:
        sec("Sales")
        show_df = filtered_df.copy()
        show_df["sale_date"] = show_df["sale_date"].dt.strftime("%Y-%m-%d")
        show_df = show_df.rename(columns={
            "voucher_no": "Voucher No", "sale_date": "Date", "customer_name": "Customer", "customer_phone": "Phone",
            "payment_mode": "Payment", "total_amount": "Total (TK)", "paid_amount": "Paid (TK)", "due_amount": "Due (TK)"})
        st.dataframe(show_df[["Voucher No", "Date", "Customer", "Phone", "Payment", "Total (TK)", "Paid (TK)", "Due (TK)"]],
                     hide_index=True, height=430, **STRETCH)
    with right:
        sec("View / Reprint / Download Voucher")
        voucher_options = filtered_df.apply(
            lambda r: f"{r['voucher_no']}  ·  {r['customer_name']}  ·  {fmt_money(r['total_amount'])}", axis=1).tolist()
        voucher_lookup = dict(zip(voucher_options, filtered_df["voucher_no"]))
        selected_label = st.selectbox("Select a voucher", voucher_options, label_visibility="collapsed")
        voucher_row = filtered_df[filtered_df["voucher_no"] == voucher_lookup[selected_label]].iloc[0].to_dict()
        st.download_button(
            "⬇️ Download Voucher (PDF)", data=build_voucher_pdf(voucher_row), file_name=f"{voucher_row['voucher_no']}.pdf",
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
    "➕ Purchase Entry": render_purchase,
    "📥 Opening Stock Entry": render_opening_stock,
    "📦 Inventory Report": render_inventory_reports,
    "📗 Customer Ledger": render_customer_ledger,
    "🧾 Supplier Ledger": render_supplier_ledger,
    "⚙️ Settings": render_settings,
}
ROUTES[page]()
