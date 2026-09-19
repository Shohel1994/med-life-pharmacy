"""
=================================================================================
 MED LIFE PHARMACY - ERP System
 Built with: Streamlit + Supabase (supabase-py)
 Apps developed by ARJ - ARJ Studio
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

EXISTING TABLES (unchanged): master_medicines, medicines, purchases,
                             supplier_ledger, sales, customer_ledger

TWO NEW TABLES (run this once in Supabase -> SQL Editor):

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

   create table if not exists app_settings (
       key   text primary key,
       value text
   );

   -- (optional, faster search / reports)
   create index if not exists idx_medicines_name on medicines (name);
   create index if not exists idx_sales_date on sales (sale_date);
   create index if not exists idx_purchases_date on purchases (purchase_date);

NOTE: If you use Row Level Security, add a policy allowing your key to
      read/write these two new tables.
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

# =================================================================================
# CUSTOM CSS - ERP STYLE (top menu bar + green left sidebar)
# =================================================================================
st.markdown(
    """
    <style>
        .stApp { background-color: #f4f6f8; }
        .block-container { padding-top: 1rem; padding-bottom: 1rem; }

        .erp-header {
            background: linear-gradient(90deg, #0b3d33 0%, #0f766e 60%, #14b8a6 100%);
            padding: 12px 22px;
            border-radius: 10px 10px 0 0;
            color: white;
        }
        .erp-header h1 { margin: 0; font-size: 1.35rem; font-weight: 700; color: white; }
        .erp-header p { margin: 2px 0 0 0; font-size: 0.8rem; opacity: 0.9; }

        /* ---- Top menu bar (like the Platform ERP) ---- */
        .st-key-topnav {
            background-color: #0f766e;
            padding: 4px 12px 6px 12px;
            border-radius: 0 0 10px 10px;
            margin-bottom: 1rem;
        }
        .st-key-topnav div[role="radiogroup"] { gap: 6px; }
        .st-key-topnav label { color: #ffffff !important; }
        .st-key-topnav label p { color: #ffffff !important; font-weight: 600; font-size: 0.95rem; }
        .st-key-topnav label > div:first-child { display: none; }  /* hide radio dot */
        .st-key-topnav label {
            padding: 6px 16px; border-radius: 6px; cursor: pointer;
        }
        .st-key-topnav label:has(input:checked) { background-color: #0b3d33; }
        .st-key-topnav label:hover { background-color: #14b8a6; }

        /* ---- Left green sidebar ---- */
        section[data-testid="stSidebar"] { background-color: #0b3d33; }
        section[data-testid="stSidebar"] * { color: #eafaf5 !important; }
        section[data-testid="stSidebar"] input { color: #111 !important; }
        section[data-testid="stSidebar"] .stRadio label { font-size: 0.95rem; }
        .arj-credit {
            margin-top: 18px; padding: 10px 6px; text-align: center;
            font-size: 0.78rem; border-top: 1px solid rgba(255,255,255,0.25);
            letter-spacing: 0.3px;
        }

        div[data-testid="stForm"] {
            background-color: white; padding: 20px; border-radius: 12px;
            border: 1px solid #e2e8f0; box-shadow: 0 1px 4px rgba(0,0,0,0.05);
        }
        div.stButton > button, div.stFormSubmitButton > button, div.stDownloadButton > button {
            border-radius: 8px; font-weight: 600; height: 2.8em;
        }
        div[data-testid="stMetric"] {
            background-color: white; padding: 14px; border-radius: 10px;
            border: 1px solid #e2e8f0; box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        }
        h2, h3 { color: #0b3d33; }

        .badge-cash { background-color: #dcfce7; color: #166534; padding: 3px 10px;
            border-radius: 20px; font-size: 0.8rem; font-weight: 700; }
        .badge-credit { background-color: #fef3c7; color: #92400e; padding: 3px 10px;
            border-radius: 20px; font-size: 0.8rem; font-weight: 700; }

        .voucher-box {
            background-color: white; border: 1.5px dashed #0f766e; border-radius: 10px;
            padding: 18px 20px; font-family: 'Courier New', monospace; color: #111827; margin-top: 0.6rem;
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
DEFAULT_DISCOUNT_PCT = 5.0

MEDICINE_TYPES = [
    "Tablet", "Capsule", "Syrup", "Suspension", "Injection", "Syringe",
    "Drops", "Ointment/Cream", "Inhaler", "Powder/Sachet", "IV Fluid/Saline", "Other",
]

SHOP_NAME = "Med Life Pharmacy"
SHOP_ADDRESS = "Chachkoir, Khalifa Para, Gurudaspur, Natore"


# =================================================================================
# SUPABASE CONNECTION (created once, reused - faster)
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
# DATA ACCESS - each page loads ONLY what it needs (lazy) + selected columns only
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
    """Build the auto-suggest list once and cache it (big catalogue = big speed win)."""
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
    """Clear transactional caches after a write. The big master catalogue is kept cached."""
    for fn in (fetch_medicines, fetch_purchases, fetch_supplier_ledger, fetch_sales,
               fetch_customer_ledger, fetch_opening_stock, fetch_supplier_names):
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
    """Add to stock if medicine exists, else create it. Returns (success, message)."""
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


def upsert_supplier_due(supplier_name: str, delta_due: float) -> bool:
    try:
        existing = supabase.table("supplier_ledger").select("total_due").eq("supplier_name", supplier_name).execute()
        now = datetime.now().isoformat()
        if existing.data:
            new_due = float(existing.data[0]["total_due"] or 0) + delta_due
            supabase.table("supplier_ledger").update({"total_due": new_due, "updated_at": now}).eq(
                "supplier_name", supplier_name).execute()
        else:
            supabase.table("supplier_ledger").insert(
                {"supplier_name": supplier_name, "total_due": delta_due, "updated_at": now}).execute()
        return True
    except Exception as e:
        st.error(f"❌ Failed to update supplier due: {e}")
        return False


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


def upsert_customer_due(name: str, phone: str, delta_due: float) -> bool:
    try:
        existing = supabase.table("customer_ledger").select("total_due").eq("customer_phone", phone).execute()
        now = datetime.now().isoformat()
        if existing.data:
            new_due = float(existing.data[0]["total_due"] or 0) + delta_due
            supabase.table("customer_ledger").update(
                {"total_due": new_due, "customer_name": name, "updated_at": now}).eq("customer_phone", phone).execute()
        else:
            supabase.table("customer_ledger").insert(
                {"customer_name": name, "customer_phone": phone, "total_due": delta_due, "updated_at": now}).execute()
        return True
    except Exception as e:
        st.error(f"❌ Failed to update customer due: {e}")
        return False


def receive_customer_payment(customer_id, new_due: float) -> bool:
    try:
        supabase.table("customer_ledger").update(
            {"total_due": new_due, "updated_at": datetime.now().isoformat()}).eq("id", customer_id).execute()
        return True
    except Exception as e:
        st.error(f"❌ Failed to record customer payment: {e}")
        return False


def insert_opening_rows(rows: list) -> bool:
    try:
        supabase.table("opening_stock").insert(rows).execute()
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
    """Create a downloadable PDF voucher. (PDF uses 'Tk' because the default PDF font has no ৳ glyph.)"""
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
        title=f"Voucher {sale_row.get('voucher_no', '')}", author="ARJ Studio",
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
        Paragraph("Apps developed by ARJ - ARJ Studio", center),
    ]
    doc.build(story)
    return buf.getvalue()


# =================================================================================
# HEADER
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

# =================================================================================
# MENU STRUCTURE  (top menu = groups, left sidebar = pages of that group)
# =================================================================================
MENU = {
    "Sales": ["🛒 Sales / POS", "🧾 Sales History / Vouchers"],
    "Purchase": ["➕ Purchase Entry"],
    "Stock": ["📥 Opening Stock Entry", "📦 Inventory Report"],
    "Ledger": ["🧾 Supplier Ledger", "📗 Customer Ledger"],
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
if st.sidebar.button("🔄 Refresh Data", use_container_width=True):
    clear_all_caches()
    st.rerun()
st.sidebar.caption(f"🕒 {datetime.now().strftime('%d %b %Y, %I:%M %p')}")
st.sidebar.markdown('<div class="arj-credit">Apps developed by ARJ<br><b>ARJ Studio</b></div>', unsafe_allow_html=True)


# =================================================================================
# PAGE: PURCHASE ENTRY
# =================================================================================
def render_purchase():
    st.subheader("➕ Purchase Entry")
    st.caption("Record NEW incoming stock with supplier and payment details. (For old/existing stock use Stock → Opening Stock Entry.)")

    display_options, option_map, form_hint_map = get_master_options()
    supplier_options = [NEW_SUPPLIER_LABEL] + fetch_supplier_names()

    if "purchase_form_version" not in st.session_state:
        st.session_state.purchase_form_version = 0
    v = st.session_state.purchase_form_version

    st.markdown("##### 💊 Medicine Details")
    col1, col2 = st.columns(2)
    with col1:
        selected_label = st.selectbox("Select Medicine", display_options, key=f"pur_med_{v}")
        custom_name = ""
        if selected_label == NEW_CUSTOM_LABEL:
            custom_name = st.text_input("Enter Medicine Name *", placeholder="e.g. Napa Extra 500mg", key=f"pur_custom_name_{v}")
    with col2:
        default_type_idx = 0
        hint = form_hint_map.get(selected_label, "")
        for i, t in enumerate(MEDICINE_TYPES):
            if hint and t.lower().startswith(hint.lower()[:4]):
                default_type_idx = i
                break
        medicine_type = st.selectbox("Medicine Type *", MEDICINE_TYPES, index=default_type_idx, key=f"pur_type_{v}")
        custom_type = ""
        if medicine_type == "Other":
            custom_type = st.text_input("Specify Medicine Type *", placeholder="e.g. Nebulizer Solution", key=f"pur_custom_type_{v}")

    st.markdown("---")
    st.markdown("##### 📦 Purchase Unit & Quantity")
    purchase_unit = st.radio("How was this purchased? *", ["Box / Carton", "Pcs (Loose Units)"], horizontal=True, key=f"pur_unit_{v}")

    box_quantity = None
    units_per_box = None
    if purchase_unit == "Box / Carton":
        colb1, colb2 = st.columns(2)
        with colb1:
            box_quantity = st.number_input("Number of Boxes *", min_value=1, value=1, step=1, key=f"pur_boxes_{v}")
        with colb2:
            units_per_box = st.number_input("Pcs per Box *", min_value=1, value=10, step=1, key=f"pur_ppb_{v}")
        total_pcs = int(box_quantity) * int(units_per_box)
        st.info(f"📦 {int(box_quantity)} Box × {int(units_per_box)} pcs/box = **{total_pcs} pcs** will be added to stock")
    else:
        total_pcs = int(st.number_input("Total Quantity in Pcs *", min_value=1, value=10, step=1, key=f"pur_pcs_{v}"))

    st.markdown("---")
    st.markdown("##### 🏭 Supplier & Payment Details")
    col3, col4 = st.columns(2)
    with col3:
        selected_supplier_label = st.selectbox("Purchased From (Supplier / Company) *", supplier_options, key=f"pur_supp_{v}")
        custom_supplier = ""
        if selected_supplier_label == NEW_SUPPLIER_LABEL:
            custom_supplier = st.text_input("Enter Supplier / Company Name *", placeholder="e.g. Square Pharmaceuticals", key=f"pur_custom_supp_{v}")
    with col4:
        payment_type = st.radio("Payment Type *", ["Cash", "Credit"], horizontal=True, key=f"pur_pay_{v}")

    total_amount = st.number_input("Total Purchase Amount (TK) *", min_value=0.0, value=0.0, step=1.0, key=f"pur_total_{v}")
    paid_amount = total_amount
    if payment_type == "Credit":
        paid_amount = st.number_input(
            "Amount Paid Now (TK)", min_value=0.0, max_value=float(total_amount), value=0.0, step=1.0, key=f"pur_paid_{v}")
    due_amount = max(total_amount - paid_amount, 0.0)
    if payment_type == "Credit":
        st.info(f"📌 Outstanding Credit for this purchase: **{fmt_money(due_amount)}**")

    st.markdown("---")
    if st.button("💾 Save Purchase & Update Stock", use_container_width=True, type="primary", key=f"pur_submit_{v}"):
        final_name = custom_name.strip() if selected_label == NEW_CUSTOM_LABEL else option_map.get(selected_label, selected_label)
        final_type = custom_type.strip() if medicine_type == "Other" else medicine_type
        final_supplier = custom_supplier.strip() if selected_supplier_label == NEW_SUPPLIER_LABEL else selected_supplier_label

        errors = []
        if not final_name:
            errors.append("Medicine name is required.")
        if medicine_type == "Other" and not final_type:
            errors.append("Please specify the custom medicine type.")
        if not final_supplier:
            errors.append("Supplier / company name is required.")
        if total_pcs <= 0:
            errors.append("Purchased quantity must be greater than 0.")
        if total_amount <= 0:
            errors.append("Total purchase amount must be greater than 0.")
        if errors:
            for err in errors:
                st.error(f"⚠️ {err}")
            st.stop()

        stock_ok, stock_msg = save_or_restock(final_name, total_pcs, final_type)
        purchase_ok = insert_purchase({
            "purchase_date": date.today().isoformat(),
            "medicine_name": final_name,
            "medicine_type": final_type,
            "supplier_name": final_supplier,
            "purchase_unit": "Box" if purchase_unit == "Box / Carton" else "Pcs",
            "box_quantity": int(box_quantity) if box_quantity is not None else None,
            "units_per_box": int(units_per_box) if units_per_box is not None else None,
            "quantity": total_pcs,
            "payment_type": payment_type,
            "total_amount": total_amount,
            "paid_amount": paid_amount if payment_type == "Credit" else total_amount,
            "due_amount": due_amount if payment_type == "Credit" else 0.0,
        })
        ledger_ok = True
        if payment_type == "Credit" and due_amount > 0:
            ledger_ok = upsert_supplier_due(final_supplier, due_amount)

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

    st.markdown("---")
    st.markdown("##### 🕘 Recent Purchases")
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
            use_container_width=True, hide_index=True)


# =================================================================================
# PAGE: OPENING STOCK ENTRY  (old/existing stock - NO supplier payable effect)
# =================================================================================
def render_opening_stock():
    st.subheader("📥 Opening Stock Entry")
    st.caption(
        "Use this for stock you ALREADY have in the shop before starting this app. "
        "It only adds to inventory — it does NOT create any purchase, supplier due or payment."
    )

    settings = fetch_settings()
    locked = settings.get("opening_locked") == "1"

    if locked:
        st.warning("🔒 Opening stock is LOCKED (entry finished). To unlock, go to Settings → Opening Stock Lock.")
    else:
        display_options, option_map, form_hint_map = get_master_options()
        tab_single, tab_import = st.tabs(["✍️ Single Entry", "📄 Excel / CSV Import (many items)"])

        with tab_single:
            if "open_ver" not in st.session_state:
                st.session_state.open_ver = 0
            ov = st.session_state.open_ver

            c1, c2 = st.columns(2)
            with c1:
                sel = st.selectbox("Select Medicine", display_options, key=f"op_med_{ov}")
                custom = ""
                if sel == NEW_CUSTOM_LABEL:
                    custom = st.text_input("Enter Medicine Name *", key=f"op_custom_{ov}")
            with c2:
                mtype = st.selectbox("Medicine Type *", MEDICINE_TYPES, key=f"op_type_{ov}")

            c3, c4, c5, c6 = st.columns(4)
            with c3:
                qty = st.number_input("Quantity (Pcs) *", min_value=1, value=1, step=1, key=f"op_qty_{ov}")
            with c4:
                cost = st.number_input("Cost price / pc", min_value=0.0, value=0.0, step=0.5, key=f"op_cost_{ov}")
            with c5:
                sale_p = st.number_input("Sale price / pc", min_value=0.0, value=0.0, step=0.5, key=f"op_sale_{ov}")
            with c6:
                entry_date = st.date_input("Stock as of date", value=date.today(), key=f"op_date_{ov}")

            if st.button("💾 Save Opening Stock", type="primary", use_container_width=True, key=f"op_save_{ov}"):
                name = custom.strip() if sel == NEW_CUSTOM_LABEL else option_map.get(sel, sel)
                if not name:
                    st.error("⚠️ Medicine name is required.")
                else:
                    ok, msg = save_or_restock(name, int(qty), mtype)
                    if ok:
                        insert_opening_rows([{
                            "entry_date": entry_date.isoformat(), "medicine_name": name, "medicine_type": mtype,
                            "quantity": int(qty), "cost_price": cost, "sale_price": sale_p}])
                        clear_data_caches()
                        st.session_state.open_ver += 1
                        st.session_state.open_flash = f"✅ Opening stock saved. {msg}"
                        st.rerun()

            if st.session_state.get("open_flash"):
                st.success(st.session_state.pop("open_flash"))

        with tab_import:
            st.caption("Columns needed: **name, quantity** — optional: medicine_type, cost_price, sale_price")
            template = pd.DataFrame(
                [{"name": "Napa 500mg", "medicine_type": "Tablet", "quantity": 500, "cost_price": 1.2, "sale_price": 1.5}])
            st.download_button("⬇️ Download CSV template", template.to_csv(index=False).encode("utf-8"),
                               "opening_stock_template.csv", "text/csv")
            up = st.file_uploader("Upload CSV or Excel", type=["csv", "xlsx"])
            if up is not None:
                try:
                    df = pd.read_csv(up) if up.name.lower().endswith(".csv") else pd.read_excel(up)
                    df.columns = [str(c).strip().lower() for c in df.columns]
                    if "name" not in df.columns or "quantity" not in df.columns:
                        st.error("⚠️ File must have 'name' and 'quantity' columns.")
                    else:
                        df["name"] = df["name"].astype(str).str.strip()
                        df["quantity"] = pd.to_numeric(df["quantity"], errors="coerce").fillna(0).astype(int)
                        df = df[(df["name"] != "") & (df["quantity"] > 0)]
                        for c, dflt in (("medicine_type", "Other"), ("cost_price", 0), ("sale_price", 0)):
                            if c not in df.columns:
                                df[c] = dflt
                        st.write(f"**{len(df)} valid rows** ready to import:")
                        st.dataframe(df.head(50), use_container_width=True, hide_index=True)
                        if st.button("✅ Import All Now", type="primary", use_container_width=True):
                            bar = st.progress(0.0)
                            log_rows, failed = [], 0
                            for i, r in enumerate(df.itertuples(index=False), start=1):
                                ok, _ = save_or_restock(r.name, int(r.quantity), str(r.medicine_type or "Other"))
                                if ok:
                                    log_rows.append({
                                        "entry_date": date.today().isoformat(), "medicine_name": r.name,
                                        "medicine_type": str(r.medicine_type or "Other"), "quantity": int(r.quantity),
                                        "cost_price": float(r.cost_price or 0), "sale_price": float(r.sale_price or 0)})
                                else:
                                    failed += 1
                                bar.progress(i / len(df))
                            if log_rows:
                                insert_opening_rows(log_rows)
                            clear_data_caches()
                            st.success(f"✅ Imported {len(log_rows)} items. Failed: {failed}")
                except Exception as e:
                    st.error(f"❌ Could not read file: {e}")

        st.markdown("---")
        if st.button("🔒 Finish & Lock Opening Stock", use_container_width=True):
            if set_setting("opening_locked", "1"):
                st.rerun()

    st.markdown("---")
    st.markdown("##### 🕘 Opening Stock Log")
    log = fetch_opening_stock()
    if log.empty:
        st.info("No opening stock entries yet.")
    else:
        log = log.rename(columns={
            "entry_date": "As of", "medicine_name": "Medicine", "medicine_type": "Type",
            "quantity": "Qty (Pcs)", "cost_price": "Cost/pc", "sale_price": "Sale/pc"})
        st.dataframe(log[["As of", "Medicine", "Type", "Qty (Pcs)", "Cost/pc", "Sale/pc"]],
                     use_container_width=True, hide_index=True)


# =================================================================================
# PAGE: SALES / POS  (discount in %, default 5%)
# =================================================================================
def render_sales():
    st.subheader("🛒 Sales / POS")

    if "sale_cart" not in st.session_state:
        st.session_state.sale_cart = []
    if "last_voucher" not in st.session_state:
        st.session_state.last_voucher = None

    meds_df = fetch_medicines()
    in_stock_df = meds_df[meds_df["stock"] > 0] if not meds_df.empty else pd.DataFrame()
    if in_stock_df.empty:
        st.warning("⚠️ No medicines currently in stock. Add stock via Purchase Entry or Opening Stock Entry.")
        return

    st.markdown("##### 🔎 Find & Add Product")
    search_term = st.text_input("Search", placeholder="Start typing a medicine name...", label_visibility="collapsed")

    filtered_df = in_stock_df
    if search_term:
        filtered_df = in_stock_df[in_stock_df["name"].str.contains(search_term, case=False, na=False, regex=False)]

    if filtered_df.empty:
        st.warning("⚠️ No matching product found.")
    else:
        filtered_df = filtered_df.head(300)  # keep dropdown light & fast
        med_options = (
            filtered_df["name"]
            + filtered_df["medicine_type"].apply(lambda t: f"  ·  {t}" if t else "")
            + filtered_df["stock"].apply(lambda s: f"  ·  Available: {int(s)} pcs")
        ).tolist()
        id_lookup = dict(zip(med_options, filtered_df["id"]))

        colp1, colp2, colp3, colp4 = st.columns([2.4, 0.9, 1.1, 0.8])
        with colp1:
            selected_option = st.selectbox("Product", med_options, label_visibility="collapsed")
        selected_id = id_lookup[selected_option]
        selected_row = filtered_df[filtered_df["id"] == selected_id].iloc[0]
        available_stock = int(selected_row["stock"])
        with colp2:
            add_qty = st.number_input("Qty", min_value=1, max_value=available_stock, value=1, step=1, label_visibility="collapsed")
        with colp3:
            add_price = st.number_input("Price/unit", min_value=0.0, value=0.0, step=0.5, format="%.2f", label_visibility="collapsed")
        with colp4:
            add_clicked = st.button("➕ Add", use_container_width=True)

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

    st.markdown("---")
    st.markdown("##### 🧺 Cart")
    if not st.session_state.sale_cart:
        st.info("Cart is empty. Search and add products above.")
        _show_last_voucher()
        return

    cart_df = pd.DataFrame(st.session_state.sale_cart)
    st.dataframe(
        cart_df.rename(columns={"name": "Product", "medicine_type": "Type", "qty": "Qty", "unit_price": "Unit Price", "subtotal": "Subtotal"})[
            ["Product", "Type", "Qty", "Unit Price", "Subtotal"]],
        use_container_width=True, hide_index=True)

    rcol1, rcol2 = st.columns([3, 1])
    remove_options = [f"{it['name']} (Qty: {it['qty']})" for it in st.session_state.sale_cart]
    with rcol1:
        item_to_remove = st.selectbox("Remove an item", remove_options, label_visibility="collapsed")
    with rcol2:
        if st.button("🗑️ Remove", use_container_width=True):
            st.session_state.sale_cart.pop(remove_options.index(item_to_remove))
            st.rerun()

    subtotal_amount = float(cart_df["subtotal"].sum())

    st.markdown("---")
    st.markdown("##### 💳 Billing Details")

    try:
        default_pct = float(fetch_settings().get("default_discount_pct", DEFAULT_DISCOUNT_PCT))
    except Exception:
        default_pct = DEFAULT_DISCOUNT_PCT
    discount_pct = st.number_input(
        "Discount (%)", min_value=0.0, max_value=100.0, value=default_pct, step=0.5, format="%.2f",
        key="sale_discount_pct", help="Default 5%. Increase or decrease for this sale.")
    discount_amount = round(subtotal_amount * discount_pct / 100, 2)
    payable_amount = round(subtotal_amount - discount_amount, 2)

    ccol1, ccol2 = st.columns(2)
    with ccol1:
        customer_name = st.text_input("Customer Name")
    with ccol2:
        customer_phone = st.text_input("Customer Phone")

    payment_mode = st.radio("Payment Mode *", ["Cash", "Credit"], horizontal=True)
    paid_amount = payable_amount
    if payment_mode == "Credit":
        paid_amount = st.number_input("Amount Paid Now (TK)", min_value=0.0, max_value=float(payable_amount), value=0.0, step=1.0)
    due_amount = max(round(payable_amount - paid_amount, 2), 0.0)

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Subtotal", fmt_money(subtotal_amount))
    m2.metric(f"Discount ({discount_pct:g}%)", fmt_money(discount_amount))
    m3.metric("Total Payable", fmt_money(payable_amount))
    m4.metric("Due Amount", fmt_money(due_amount))

    need_cust = payment_mode == "Credit" and due_amount > 0 and (not customer_name.strip() or not customer_phone.strip())
    if need_cust:
        st.warning("⚠️ Customer Name & Phone are required for Credit sales with a remaining due amount.")

    st.markdown("---")
    if st.button("✅ Checkout & Generate Voucher", use_container_width=True, type="primary", disabled=need_cust):
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
            sale_payload = {
                "voucher_no": voucher_no,
                "sale_date": date.today().isoformat(),
                "customer_name": customer_name.strip() or "Walk-in Customer",
                "customer_phone": customer_phone.strip(),
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
                ledger_ok = upsert_customer_due(sale_payload["customer_name"], customer_phone.strip(), due_amount)
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
    st.markdown("---")
    st.success(f"✅ Sale completed! Voucher No: **{lv['voucher_no']}**")
    st.markdown("##### 🧾 Voucher / Receipt")
    st.markdown(render_voucher_html(lv), unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        st.download_button(
            "⬇️ Download Voucher (PDF)", data=build_voucher_pdf(lv), file_name=f"{lv['voucher_no']}.pdf",
            mime="application/pdf", use_container_width=True, type="primary", key="dl_last_voucher")
    with c2:
        if st.button("✖️ Clear Voucher Preview", use_container_width=True):
            st.session_state.last_voucher = None
            st.rerun()


# =================================================================================
# PAGE: INVENTORY REPORT
# =================================================================================
def render_inventory_reports():
    st.subheader("📦 Inventory Report")
    meds_df = fetch_medicines()
    if meds_df.empty:
        st.info("ℹ️ No inventory records found yet.")
        return

    low_stock_df = meds_df[meds_df["stock"] < 10]
    m1, m2, m3 = st.columns(3)
    m1.metric("💊 Total Medicine Types", f"{meds_df.shape[0]}")
    m2.metric("📦 Total Stock (Pcs)", f"{int(meds_df['stock'].sum())} pcs")
    m3.metric("⚠️ Low Stock Items (<10)", f"{low_stock_df.shape[0]}")

    st.markdown("---")
    col1, col2 = st.columns([2, 1])
    with col1:
        search_term = st.text_input("🔍 Search by Medicine Name")
    with col2:
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
    st.dataframe(show_df[cols].sort_values("Medicine Name"), use_container_width=True, hide_index=True, height=420)

    if not low_stock_df.empty:
        st.markdown("##### ⚠️ Low Stock Alert (below 10 units)")
        st.dataframe(
            low_stock_df.rename(columns={"name": "Medicine Name", "medicine_type": "Type", "stock": "Stock (Pcs)"})[
                ["Medicine Name", "Type", "Stock (Pcs)"]].sort_values("Stock (Pcs)"),
            use_container_width=True, hide_index=True)


# =================================================================================
# PAGE: SUPPLIER LEDGER
# =================================================================================
def render_supplier_ledger():
    st.subheader("🧾 Supplier Ledger — Credit Due")
    ledger_df = fetch_supplier_ledger()

    total_outstanding = 0.0 if ledger_df.empty else float(ledger_df["total_due"].sum())
    m1, m2 = st.columns(2)
    m1.metric("🏭 Suppliers with Credit Due", f"{ledger_df[ledger_df['total_due'] > 0].shape[0] if not ledger_df.empty else 0}")
    m2.metric("💳 Total Outstanding Credit", fmt_money(total_outstanding))

    st.markdown("---")
    search_term = st.text_input("Search by Supplier Name", label_visibility="collapsed", placeholder="🔍 Search by Supplier Name")
    display_df = ledger_df
    if not display_df.empty and search_term:
        display_df = display_df[display_df["supplier_name"].str.contains(search_term, case=False, na=False, regex=False)]
    if display_df.empty:
        st.info("No supplier ledger records found.")
    else:
        st.dataframe(display_df.rename(columns={"supplier_name": "Supplier", "total_due": "Outstanding Due (TK)"})[
            ["Supplier", "Outstanding Due (TK)"]], use_container_width=True, hide_index=True)

    st.markdown("---")
    st.markdown("##### 💵 Pay Supplier (Settle Credit)")
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
        if st.button("💵 Record Payment to Supplier", use_container_width=True, type="primary"):
            if paid_now <= 0:
                st.error("⚠️ Enter a valid amount greater than 0.")
            else:
                new_due = float(selected_row["total_due"]) - paid_now
                if pay_supplier(selected_id, new_due):
                    clear_data_caches()
                    st.success(f"✅ Payment of {fmt_money(paid_now)} recorded! Remaining due: {fmt_money(new_due)}")

    st.markdown("---")
    st.markdown("##### 📜 Credit Purchase History")
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
                     use_container_width=True, hide_index=True)


# =================================================================================
# PAGE: CUSTOMER LEDGER
# =================================================================================
def render_customer_ledger():
    st.subheader("📗 Customer Ledger — Credit Due")
    ledger_df = fetch_customer_ledger()

    total_outstanding = 0.0 if ledger_df.empty else float(ledger_df["total_due"].sum())
    m1, m2 = st.columns(2)
    m1.metric("🙍 Customers with Due", f"{ledger_df[ledger_df['total_due'] > 0].shape[0] if not ledger_df.empty else 0}")
    m2.metric("💳 Total Outstanding Due", fmt_money(total_outstanding))

    st.markdown("---")
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
            ["Customer", "Phone", "Due (TK)"]], use_container_width=True, hide_index=True)

    st.markdown("---")
    st.markdown("##### 💵 Receive Payment")
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
        paid_now = st.number_input("Amount Paid Now *", min_value=0.0, max_value=float(selected_row["total_due"]), step=1.0)
        if st.button("💵 Collect Due", use_container_width=True, type="primary"):
            if paid_now <= 0:
                st.error("⚠️ Enter a valid amount greater than 0.")
            else:
                new_due = float(selected_row["total_due"]) - paid_now
                if receive_customer_payment(selected_id, new_due):
                    clear_data_caches()
                    st.success(f"✅ Payment of {fmt_money(paid_now)} recorded! Remaining due: {fmt_money(new_due)}")


# =================================================================================
# PAGE: SALES HISTORY / VOUCHERS  (with PDF download)
# =================================================================================
def render_sales_history():
    st.subheader("🧾 Sales History / Vouchers")

    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input("From Date", value=date.today() - pd.Timedelta(days=30))
    with col2:
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

    st.markdown("---")
    show_df = filtered_df.copy()
    show_df["sale_date"] = show_df["sale_date"].dt.strftime("%Y-%m-%d")
    show_df = show_df.rename(columns={
        "voucher_no": "Voucher No", "sale_date": "Date", "customer_name": "Customer", "customer_phone": "Phone",
        "payment_mode": "Payment", "total_amount": "Total (TK)", "paid_amount": "Paid (TK)", "due_amount": "Due (TK)"})
    st.dataframe(show_df[["Voucher No", "Date", "Customer", "Phone", "Payment", "Total (TK)", "Paid (TK)", "Due (TK)"]],
                 use_container_width=True, hide_index=True, height=320)

    st.markdown("---")
    st.markdown("##### 🧾 View / Reprint / Download a Voucher")
    voucher_options = filtered_df.apply(
        lambda r: f"{r['voucher_no']}  ·  {r['customer_name']}  ·  {fmt_money(r['total_amount'])}", axis=1).tolist()
    voucher_lookup = dict(zip(voucher_options, filtered_df["voucher_no"]))
    selected_label = st.selectbox("Select a voucher", voucher_options)
    voucher_row = filtered_df[filtered_df["voucher_no"] == voucher_lookup[selected_label]].iloc[0].to_dict()

    st.download_button(
        "⬇️ Download Voucher (PDF)", data=build_voucher_pdf(voucher_row), file_name=f"{voucher_row['voucher_no']}.pdf",
        mime="application/pdf", type="primary", use_container_width=True, key="dl_hist_voucher")
    st.markdown(render_voucher_html(voucher_row), unsafe_allow_html=True)


# =================================================================================
# PAGE: SETTINGS
# =================================================================================
def render_settings():
    st.subheader("⚙️ Settings")

    st.markdown("##### 🔌 Connection Status")
    try:
        supabase.table("medicines").select("id").limit(1).execute()
        st.success("✅ Connected to Supabase successfully.")
    except Exception as e:
        st.error(f"❌ Supabase connection issue: {e}")

    settings = fetch_settings()

    st.markdown("---")
    st.markdown("##### 🏷️ Default Sales Discount")
    try:
        cur = float(settings.get("default_discount_pct", DEFAULT_DISCOUNT_PCT))
    except Exception:
        cur = DEFAULT_DISCOUNT_PCT
    new_pct = st.number_input("Default discount % on new sales", min_value=0.0, max_value=100.0, value=cur, step=0.5)
    if st.button("💾 Save Default Discount"):
        if set_setting("default_discount_pct", str(new_pct)):
            st.success(f"✅ Default discount set to {new_pct:g}%")

    st.markdown("---")
    st.markdown("##### 🔒 Opening Stock Lock")
    if settings.get("opening_locked") == "1":
        st.warning("Opening stock entry is currently LOCKED.")
        if st.button("🔓 Unlock Opening Stock Entry"):
            if set_setting("opening_locked", "0"):
                st.rerun()
    else:
        st.info("Opening stock entry is open. Lock it from Stock → Opening Stock Entry when finished.")

    st.markdown("---")
    st.markdown("##### 🗂️ Master Medicine Catalogue")
    master_df = fetch_master_medicines()
    if master_df.empty:
        st.info("No records found in `master_medicines`.")
    else:
        st.write(f"Total catalogue entries: **{master_df.shape[0]}**")

    st.markdown("---")
    if st.button("🔄 Clear Cache & Refresh Now", use_container_width=True):
        clear_all_caches()
        st.success("✅ Cache cleared successfully.")
        st.rerun()

    st.markdown("---")
    st.caption("Med Life Pharmacy ERP · Built with Streamlit + Supabase · Apps developed by ARJ — ARJ Studio")


# =================================================================================
# ROUTER  (only the selected page runs = fast)
# =================================================================================
ROUTES = {
    "🛒 Sales / POS": render_sales,
    "🧾 Sales History / Vouchers": render_sales_history,
    "➕ Purchase Entry": render_purchase,
    "📥 Opening Stock Entry": render_opening_stock,
    "📦 Inventory Report": render_inventory_reports,
    "🧾 Supplier Ledger": render_supplier_ledger,
    "📗 Customer Ledger": render_customer_ledger,
    "⚙️ Settings": render_settings,
}
ROUTES[page]()
