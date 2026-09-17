"""
=================================================================================
 MED LIFE PHARMACY - ERP System (Purchase / Sales / Inventory / Supplier Ledger)
 Built with: Streamlit + Supabase (supabase-py)
=================================================================================

SUPABASE TABLE SCHEMA (must already exist in your Supabase project):

1) master_medicines   (a master catalogue used only for auto-suggest during purchase)
   - id                bigint, primary key, identity
   - brand_name         text, not null
   - company_name       text
   - generic_name       text
   - form               text            e.g. Tablet, Syrup, Capsule, Injection

2) medicines           (the live inventory table)
   - id                bigint, primary key, identity
   - name               text, not null, unique   (recommended: add a UNIQUE constraint)
   - medicine_type      text                       e.g. Tablet, Syrup, Injection...
   - stock              integer, not null, default 0
   - created_at         timestamptz, default now()

3) purchases           (every purchase/stock-in transaction, cash or credit)
   - id                bigint, primary key, identity
   - purchase_date      date, not null
   - medicine_name      text, not null
   - medicine_type      text
   - supplier_name      text, not null
   - purchase_unit      text            'Box' or 'Pcs' — how it was purchased
   - box_quantity        integer, nullable   number of boxes/cartons (only when purchase_unit = 'Box')
   - units_per_box       integer, nullable   pcs inside each box (only when purchase_unit = 'Box')
   - quantity           integer, not null    TOTAL PCS added to stock (Box × Units-per-Box, or direct pcs)
   - payment_type       text            'Cash' or 'Credit'
   - total_amount       numeric, not null, default 0
   - paid_amount        numeric, not null, default 0
   - due_amount         numeric, not null, default 0
   - created_at         timestamptz, default now()

   NOTE ON UNITS: Medicines are usually PURCHASED in Box/Carton but SOLD in individual Pcs.
   To keep inventory math simple and accurate, the `medicines.stock` column always stores
   the quantity in PCS. When you purchase by Box, the app asks for "Number of Boxes" and
   "Pcs per Box", multiplies them automatically, and adds that total to stock in pcs.

4) supplier_ledger     (running credit balance owed to each supplier/company)
   - id                bigint, primary key, identity
   - supplier_name      text, not null, unique
   - total_due          numeric, not null, default 0
   - updated_at         timestamptz, default now()

5) sales               (one row PER INVOICE/VOUCHER — a customer may buy many products in one sale)
   - id                bigint, primary key, identity
   - voucher_no         text, not null, unique   e.g. MLP-20260917-143205-482
   - sale_date          date, not null
   - customer_name      text
   - customer_phone     text
   - items              text, not null   JSON string: list of {medicine_id, name, medicine_type, qty, unit_price, subtotal}
   - subtotal           numeric, not null, default 0     sum of all line items before discount
   - discount           numeric, not null, default 0
   - total_amount       numeric, not null, default 0     subtotal - discount (the payable amount)
   - payment_mode       text            'Cash' or 'Credit'
   - paid_amount        numeric, not null, default 0
   - due_amount         numeric, not null, default 0
   - created_at         timestamptz, default now()

6) customer_ledger     (running credit balance a customer owes the pharmacy)
   - id                bigint, primary key, identity
   - customer_name      text, not null
   - customer_phone     text, not null, unique
   - total_due          numeric, not null, default 0
   - updated_at         timestamptz, default now()

STREAMLIT SECRETS (.streamlit/secrets.toml):
   SUPABASE_URL = "https://xxxxxxxxxxxx.supabase.co"
   SUPABASE_KEY = "your-supabase-anon-or-service-key"

PYTHON DEPENDENCIES:
   pip install streamlit supabase pandas
=================================================================================
"""

import json
import random
from datetime import date, datetime

import pandas as pd
import streamlit as st
from supabase import create_client, Client


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
# CUSTOM CSS - CORPORATE / ERP STYLE
# =================================================================================
st.markdown(
    """
    <style>
        .stApp { background-color: #f4f6f8; }

        .erp-header {
            background: linear-gradient(90deg, #0b3d33 0%, #0f766e 60%, #14b8a6 100%);
            padding: 20px 28px;
            border-radius: 14px;
            color: white;
            margin-bottom: 1.4rem;
            box-shadow: 0 4px 14px rgba(15, 118, 110, 0.25);
        }
        .erp-header h1 { margin: 0; font-size: 1.7rem; font-weight: 700; color: white; letter-spacing: 0.3px; }
        .erp-header p { margin: 4px 0 0 0; font-size: 0.9rem; opacity: 0.9; }

        section[data-testid="stSidebar"] { background-color: #0b3d33; }
        section[data-testid="stSidebar"] * { color: #eafaf5 !important; }
        section[data-testid="stSidebar"] .stRadio label { font-size: 1.0rem; }

        div[data-testid="stForm"] {
            background-color: white;
            padding: 20px;
            border-radius: 12px;
            border: 1px solid #e2e8f0;
            box-shadow: 0 1px 4px rgba(0,0,0,0.05);
        }

        div.stButton > button, div.stFormSubmitButton > button {
            border-radius: 8px;
            font-weight: 600;
            height: 2.8em;
        }

        div[data-testid="stMetric"] {
            background-color: white;
            padding: 14px;
            border-radius: 10px;
            border: 1px solid #e2e8f0;
            box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        }

        h2, h3 { color: #0b3d33; }

        .badge-cash {
            background-color: #dcfce7; color: #166534; padding: 3px 10px;
            border-radius: 20px; font-size: 0.8rem; font-weight: 700;
        }
        .badge-credit {
            background-color: #fef3c7; color: #92400e; padding: 3px 10px;
            border-radius: 20px; font-size: 0.8rem; font-weight: 700;
        }

        /* Supershop-style voucher/receipt */
        .voucher-box {
            background-color: white;
            border: 1.5px dashed #0f766e;
            border-radius: 10px;
            padding: 18px 20px;
            font-family: 'Courier New', monospace;
            color: #111827;
            margin-top: 0.6rem;
        }
        .voucher-box h3 {
            text-align: center;
            margin: 0 0 2px 0;
            color: #0b3d33;
        }
        .voucher-box .v-sub {
            text-align: center;
            font-size: 0.8rem;
            color: #444;
            margin-bottom: 10px;
        }
        .voucher-box hr {
            border: none;
            border-top: 1px dashed #999;
            margin: 8px 0;
        }
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

MEDICINE_TYPES = [
    "Tablet", "Capsule", "Syrup", "Suspension", "Injection", "Syringe",
    "Drops", "Ointment/Cream", "Inhaler", "Powder/Sachet", "IV Fluid/Saline", "Other",
]


# =================================================================================
# SUPABASE CONNECTION
# =================================================================================
@st.cache_resource
def init_connection() -> Client:
    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_KEY"]
    return create_client(url, key)


try:
    supabase: Client = init_connection()
except Exception as e:
    st.error(f"❌ Could not connect to Supabase. Please check your secrets.toml configuration. Details: {e}")
    st.stop()


# =================================================================================
# DATA ACCESS HELPERS (every Supabase call wrapped in try/except)
# =================================================================================
@st.cache_data(ttl=20, show_spinner=False)
def fetch_master_medicines() -> pd.DataFrame:
    try:
        res = supabase.table("master_medicines").select("*").order("brand_name").execute()
        return pd.DataFrame(res.data)
    except Exception as e:
        st.error(f"❌ Failed to fetch master medicine catalogue: {e}")
        return pd.DataFrame()


@st.cache_data(ttl=20, show_spinner=False)
def fetch_medicines() -> pd.DataFrame:
    try:
        res = supabase.table("medicines").select("*").order("name").execute()
        df = pd.DataFrame(res.data)
        if not df.empty:
            df["stock"] = pd.to_numeric(df["stock"], errors="coerce").fillna(0).astype(int)
            if "medicine_type" not in df.columns:
                df["medicine_type"] = ""
            df["medicine_type"] = df["medicine_type"].fillna("")
        return df
    except Exception as e:
        st.error(f"❌ Failed to fetch inventory: {e}")
        return pd.DataFrame()


@st.cache_data(ttl=20, show_spinner=False)
def fetch_purchases() -> pd.DataFrame:
    try:
        res = supabase.table("purchases").select("*").order("purchase_date", desc=True).execute()
        df = pd.DataFrame(res.data)
        if not df.empty:
            df["purchase_date"] = pd.to_datetime(df["purchase_date"], errors="coerce")
            for col in ["quantity", "total_amount", "paid_amount", "due_amount"]:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        return df
    except Exception as e:
        st.error(f"❌ Failed to fetch purchase history: {e}")
        return pd.DataFrame()


@st.cache_data(ttl=20, show_spinner=False)
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


@st.cache_data(ttl=15, show_spinner=False)
def fetch_sales() -> pd.DataFrame:
    try:
        res = supabase.table("sales").select("*").order("created_at", desc=True).execute()
        df = pd.DataFrame(res.data)
        if not df.empty:
            df["sale_date"] = pd.to_datetime(df["sale_date"], errors="coerce")
            for col in ["subtotal", "discount", "total_amount", "paid_amount", "due_amount"]:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        return df
    except Exception as e:
        st.error(f"❌ Failed to fetch sales history: {e}")
        return pd.DataFrame()


@st.cache_data(ttl=15, show_spinner=False)
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


def clear_all_caches():
    fetch_master_medicines.clear()
    fetch_medicines.clear()
    fetch_purchases.clear()
    fetch_supplier_ledger.clear()
    fetch_sales.clear()
    fetch_customer_ledger.clear()


def find_medicine_by_name(name: str):
    try:
        res = supabase.table("medicines").select("*").eq("name", name).execute()
        return res.data[0] if res.data else None
    except Exception as e:
        st.error(f"❌ Failed to look up medicine: {e}")
        return None


def insert_medicine(name: str, stock: int, medicine_type: str) -> bool:
    try:
        supabase.table("medicines").insert(
            {"name": name, "stock": stock, "medicine_type": medicine_type}
        ).execute()
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
    """If a medicine with this exact name already exists, add to its stock (restock) and
    refresh its type. Otherwise insert a brand new row. Returns (success, message)."""
    name = name.strip()
    if not name:
        return False, "Medicine name cannot be empty."

    existing = find_medicine_by_name(name)
    if existing:
        new_stock = int(existing["stock"]) + added_qty
        ok = update_medicine_stock(existing["id"], new_stock, medicine_type)
        if ok:
            return True, f"'{name}' already existed — stock increased to {new_stock} pcs."
        return False, "Failed to update existing medicine stock."
    else:
        ok = insert_medicine(name, added_qty, medicine_type)
        if ok:
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
    """Add delta_due to an existing supplier's total_due, or create a new ledger record."""
    try:
        existing = supabase.table("supplier_ledger").select("*").eq("supplier_name", supplier_name).execute()
        if existing.data:
            current_due = float(existing.data[0]["total_due"] or 0)
            new_due = current_due + delta_due
            supabase.table("supplier_ledger").update(
                {"total_due": new_due, "updated_at": datetime.now().isoformat()}
            ).eq("supplier_name", supplier_name).execute()
        else:
            supabase.table("supplier_ledger").insert(
                {"supplier_name": supplier_name, "total_due": delta_due, "updated_at": datetime.now().isoformat()}
            ).execute()
        return True
    except Exception as e:
        st.error(f"❌ Failed to update supplier due: {e}")
        return False


def pay_supplier(supplier_id, new_due: float) -> bool:
    try:
        supabase.table("supplier_ledger").update(
            {"total_due": new_due, "updated_at": datetime.now().isoformat()}
        ).eq("id", supplier_id).execute()
        return True
    except Exception as e:
        st.error(f"❌ Failed to record supplier payment: {e}")
        return False


def generate_voucher_no() -> str:
    """Unique, human-readable voucher/invoice number, e.g. MLP-20260917-143205-482."""
    return f"MLP-{datetime.now().strftime('%Y%m%d-%H%M%S')}-{random.randint(100, 999)}"


def insert_sale(payload: dict):
    """Insert a sale (invoice) row. Returns the inserted row (dict) on success, else None."""
    try:
        res = supabase.table("sales").insert(payload).execute()
        return res.data[0] if res.data else payload
    except Exception as e:
        st.error(f"❌ Failed to record sale: {e}")
        return None


def upsert_customer_due(name: str, phone: str, delta_due: float) -> bool:
    """Add delta_due to an existing customer's total_due, or create a new ledger record."""
    try:
        existing = supabase.table("customer_ledger").select("*").eq("customer_phone", phone).execute()
        if existing.data:
            current_due = float(existing.data[0]["total_due"] or 0)
            new_due = current_due + delta_due
            supabase.table("customer_ledger").update(
                {"total_due": new_due, "customer_name": name, "updated_at": datetime.now().isoformat()}
            ).eq("customer_phone", phone).execute()
        else:
            supabase.table("customer_ledger").insert(
                {
                    "customer_name": name,
                    "customer_phone": phone,
                    "total_due": delta_due,
                    "updated_at": datetime.now().isoformat(),
                }
            ).execute()
        return True
    except Exception as e:
        st.error(f"❌ Failed to update customer due: {e}")
        return False


def receive_customer_payment(customer_id, new_due: float) -> bool:
    try:
        supabase.table("customer_ledger").update(
            {"total_due": new_due, "updated_at": datetime.now().isoformat()}
        ).eq("id", customer_id).execute()
        return True
    except Exception as e:
        st.error(f"❌ Failed to record customer payment: {e}")
        return False


def fmt_money(val) -> str:
    try:
        return f"৳{float(val):,.2f}"
    except Exception:
        return "৳0.00"


def render_voucher_html(sale_row: dict) -> str:
    """Build the supershop-style printable voucher HTML for a completed sale."""
    try:
        items = json.loads(sale_row.get("items") or "[]")
    except Exception:
        items = []

    sale_date = sale_row.get("sale_date", "")
    if hasattr(sale_date, "strftime"):
        sale_date = sale_date.strftime("%Y-%m-%d")

    rows_html = ""
    for it in items:
        rows_html += (
            f"<tr><td>{it.get('name', '')}"
            f"{' (' + it.get('medicine_type') + ')' if it.get('medicine_type') else ''}</td>"
            f"<td class='v-right'>{it.get('qty', 0)}</td>"
            f"<td class='v-right'>{fmt_money(it.get('unit_price', 0))}</td>"
            f"<td class='v-right'>{fmt_money(it.get('subtotal', 0))}</td></tr>"
        )

    payment_mode = sale_row.get("payment_mode", "")
    badge_class = "badge-cash" if payment_mode == "Cash" else "badge-credit"

    html = f"""
    <div class="voucher-box">
        <h3>💊 Med Life Pharmacy</h3>
        <div class="v-sub">Chachkoir, Khalifa Para, Gurudaspur, Natore</div>
        <hr>
        <div><b>Voucher No:</b> {sale_row.get('voucher_no', '')}</div>
        <div><b>Date:</b> {sale_date}</div>
        <div><b>Customer:</b> {sale_row.get('customer_name') or 'Walk-in Customer'}
             {('· ' + sale_row.get('customer_phone')) if sale_row.get('customer_phone') else ''}</div>
        <hr>
        <table>
            <tr><th>Item</th><th class="v-right">Qty</th><th class="v-right">Price</th><th class="v-right">Amount</th></tr>
            {rows_html}
        </table>
        <hr>
        <table>
            <tr><td>Subtotal</td><td class="v-right">{fmt_money(sale_row.get('subtotal', 0))}</td></tr>
            <tr><td>Discount</td><td class="v-right">- {fmt_money(sale_row.get('discount', 0))}</td></tr>
            <tr class="v-total-row"><td>Total Payable</td><td class="v-right">{fmt_money(sale_row.get('total_amount', 0))}</td></tr>
            <tr><td>Paid Amount</td><td class="v-right">{fmt_money(sale_row.get('paid_amount', 0))}</td></tr>
            <tr><td>Due Amount</td><td class="v-right">{fmt_money(sale_row.get('due_amount', 0))}</td></tr>
        </table>
        <hr>
        <div>Payment Mode: <span class="{badge_class}">{payment_mode}</span></div>
        <div class="v-sub" style="margin-top:10px;">Thank you for shopping with us!</div>
    </div>
    """
    return html


# =================================================================================
# HEADER
# =================================================================================
st.markdown(
    """
    <div class="erp-header">
        <h1>💊 Med Life Pharmacy</h1>
        <p>ERP System — Purchase, Multi-Product Sales, Inventory &amp; Ledger Management</p>
    </div>
    """,
    unsafe_allow_html=True,
)

# =================================================================================
# SIDEBAR NAVIGATION
# =================================================================================
st.sidebar.markdown("## 📋 ERP Modules")
page = st.sidebar.radio(
    "Navigate",
    [
        "➕ Purchase / Stock Entry",
        "🛒 Sales Module",
        "📦 Inventory Reports",
        "🧾 Supplier Ledger (Credit Due)",
        "📗 Customer Ledger (Credit Due)",
        "🧾 Sales History / Vouchers",
        "⚙️ Settings",
    ],
    label_visibility="collapsed",
)

st.sidebar.markdown("---")
if st.sidebar.button("🔄 Refresh Data", use_container_width=True):
    clear_all_caches()
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.caption(f"🕒 {datetime.now().strftime('%d %b %Y, %I:%M %p')}")


# =================================================================================
# MODULE 1: PURCHASE / STOCK ENTRY
# =================================================================================
def render_purchase():
    st.subheader("➕ Purchase / Stock Entry")
    st.caption("Record incoming stock along with medicine type, supplier and payment details (Cash or Credit).")

    master_df = fetch_master_medicines()
    purchases_df = fetch_purchases()

    # ---------- Build medicine auto-suggest options ----------
    option_map = {}
    display_options = [NEW_CUSTOM_LABEL]
    form_hint_map = {}

    if not master_df.empty:
        for _, row in master_df.iterrows():
            brand = str(row.get("brand_name") or "").strip()
            company = str(row.get("company_name") or "").strip()
            form = str(row.get("form") or "").strip()
            if not brand:
                continue
            label = f"{brand} ({form}) - {company}" if form or company else brand
            option_map[label] = brand
            form_hint_map[label] = form
            display_options.append(label)
    else:
        st.info("ℹ️ No items found in the master medicine catalogue yet. You can still add a custom medicine below.")

    # ---------- Build supplier auto-suggest options ----------
    supplier_options = [NEW_SUPPLIER_LABEL]
    if not purchases_df.empty and "supplier_name" in purchases_df.columns:
        known_suppliers = sorted(purchases_df["supplier_name"].dropna().unique().tolist())
        supplier_options += known_suppliers

    # A version counter is used as a key-suffix so that after a successful save we can
    # reset every input back to its default (st.form's clear_on_submit can't be used here
    # because we need fields to react live — e.g. Box vs Pcs — before the user submits).
    if "purchase_form_version" not in st.session_state:
        st.session_state.purchase_form_version = 0
    v = st.session_state.purchase_form_version

    st.markdown("##### 💊 Medicine Details")
    col1, col2 = st.columns(2)
    with col1:
        selected_label = st.selectbox("Select Medicine", display_options, key=f"pur_med_{v}")
        custom_name = ""
        if selected_label == NEW_CUSTOM_LABEL:
            custom_name = st.text_input(
                "Enter Medicine Name *", placeholder="e.g. Napa Extra 500mg", key=f"pur_custom_name_{v}"
            )

    with col2:
        default_type_idx = 0
        hint = form_hint_map.get(selected_label, "")
        for i, t in enumerate(MEDICINE_TYPES):
            if hint and t.lower().startswith(hint.lower()[:4]):
                default_type_idx = i
                break
        medicine_type = st.selectbox(
            "Medicine Type *", MEDICINE_TYPES, index=default_type_idx, key=f"pur_type_{v}"
        )
        custom_type = ""
        if medicine_type == "Other":
            custom_type = st.text_input(
                "Specify Medicine Type *", placeholder="e.g. Nebulizer Solution", key=f"pur_custom_type_{v}"
            )

    st.markdown("---")
    st.markdown("##### 📦 Purchase Unit & Quantity")
    st.caption("Medicines are usually purchased by Box/Carton but sold individually in Pcs — the app converts automatically.")

    purchase_unit = st.radio(
        "How was this purchased? *", ["Box / Carton", "Pcs (Loose Units)"], horizontal=True, key=f"pur_unit_{v}"
    )

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
        total_pcs = st.number_input(
            "Total Quantity in Pcs *", min_value=1, value=10, step=1, key=f"pur_pcs_{v}"
        )
        total_pcs = int(total_pcs)

    st.markdown("---")
    st.markdown("##### 🏭 Supplier & Payment Details")
    col3, col4 = st.columns(2)
    with col3:
        selected_supplier_label = st.selectbox(
            "Purchased From (Supplier / Company) *", supplier_options, key=f"pur_supp_{v}"
        )
        custom_supplier = ""
        if selected_supplier_label == NEW_SUPPLIER_LABEL:
            custom_supplier = st.text_input(
                "Enter Supplier / Company Name *", placeholder="e.g. Square Pharmaceuticals", key=f"pur_custom_supp_{v}"
            )
    with col4:
        payment_type = st.radio("Payment Type *", ["Cash", "Credit"], horizontal=True, key=f"pur_pay_{v}")

    total_amount = st.number_input(
        "Total Purchase Amount (TK) *", min_value=0.0, value=0.0, step=1.0, key=f"pur_total_{v}"
    )

    paid_amount = total_amount
    if payment_type == "Credit":
        paid_amount = st.number_input(
            "Amount Paid Now (TK)", min_value=0.0, max_value=float(total_amount), value=0.0, step=1.0,
            key=f"pur_paid_{v}",
        )

    due_amount = max(total_amount - paid_amount, 0.0)
    if payment_type == "Credit":
        st.info(f"📌 Outstanding Credit for this purchase: **{fmt_money(due_amount)}**")

    st.markdown("---")
    submitted = st.button(
        "💾 Save Purchase & Update Stock", use_container_width=True, type="primary", key=f"pur_submit_{v}"
    )

    if submitted:
        # Resolve medicine name
        if selected_label == NEW_CUSTOM_LABEL:
            final_name = custom_name.strip()
        else:
            final_name = option_map.get(selected_label, selected_label)

        # Resolve medicine type
        final_type = custom_type.strip() if medicine_type == "Other" else medicine_type

        # Resolve supplier
        if selected_supplier_label == NEW_SUPPLIER_LABEL:
            final_supplier = custom_supplier.strip()
        else:
            final_supplier = selected_supplier_label

        # ---------- Validation ----------
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
        if payment_type == "Credit" and paid_amount > total_amount:
            errors.append("Paid amount cannot exceed the total purchase amount.")

        if errors:
            for err in errors:
                st.error(f"⚠️ {err}")
            st.stop()

        # ---------- Save ----------
        stock_ok, stock_msg = save_or_restock(final_name, total_pcs, final_type)

        purchase_payload = {
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
        }
        purchase_ok = insert_purchase(purchase_payload)

        ledger_ok = True
        if payment_type == "Credit" and due_amount > 0:
            ledger_ok = upsert_supplier_due(final_supplier, due_amount)

        if stock_ok and purchase_ok and ledger_ok:
            clear_all_caches()
            st.success(f"✅ {stock_msg}")
            st.success(
                f"🧾 Purchase recorded — {payment_type} — {total_pcs} pcs — Total: {fmt_money(total_amount)}"
                + (f" | Due: {fmt_money(due_amount)}" if payment_type == "Credit" and due_amount > 0 else "")
            )
            st.balloons()
            # Reset all inputs back to their defaults for the next entry
            st.session_state.purchase_form_version += 1
            st.rerun()
        else:
            st.error("❌ Something went wrong while saving. Please check the errors above and try again.")

    # ---------- Recent purchases preview ----------
    st.markdown("---")
    st.markdown("##### 🕘 Recent Purchases")
    if purchases_df.empty:
        st.info("No purchase records yet.")
    else:
        recent = purchases_df.head(10).copy()
        recent["purchase_date"] = recent["purchase_date"].dt.strftime("%Y-%m-%d")

        def describe_unit(row):
            if row.get("purchase_unit") == "Box" and pd.notna(row.get("box_quantity")) and pd.notna(row.get("units_per_box")):
                return f"{int(row['box_quantity'])} Box × {int(row['units_per_box'])}"
            return "Loose Pcs"

        recent["Purchased As"] = recent.apply(describe_unit, axis=1)
        recent = recent.rename(
            columns={
                "purchase_date": "Date",
                "medicine_name": "Medicine",
                "medicine_type": "Type",
                "supplier_name": "Supplier",
                "quantity": "Total Pcs",
                "payment_type": "Payment",
                "total_amount": "Total (TK)",
                "paid_amount": "Paid (TK)",
                "due_amount": "Due (TK)",
            }
        )
        st.dataframe(
            recent[
                ["Date", "Medicine", "Type", "Supplier", "Purchased As", "Total Pcs",
                 "Payment", "Total (TK)", "Paid (TK)", "Due (TK)"]
            ],
            use_container_width=True,
            hide_index=True,
        )


# =================================================================================
# MODULE 2: SALES MODULE  (multi-product cart, voucher, discount, Cash/Credit)
# =================================================================================
def render_sales():
    st.subheader("🛒 Sales Module")
    st.caption("Add one or more products to the cart, then checkout to generate a voucher — just like a supershop bill.")

    if "sale_cart" not in st.session_state:
        st.session_state.sale_cart = []  # list of dicts: medicine_id, name, medicine_type, qty, unit_price, subtotal
    if "last_voucher" not in st.session_state:
        st.session_state.last_voucher = None

    meds_df = fetch_medicines()
    in_stock_df = meds_df[meds_df["stock"] > 0] if not meds_df.empty else pd.DataFrame()

    if in_stock_df.empty:
        st.warning("⚠️ No medicines currently in stock. Please add stock from the Purchase module first.")
        return

    # ---------- Add Product to Cart ----------
    st.markdown("##### 🔎 Find & Add Product")
    search_term = st.text_input(
        "Type to search product name", placeholder="Start typing a medicine name...", label_visibility="collapsed"
    )

    filtered_df = in_stock_df.copy()
    if search_term:
        filtered_df = filtered_df[filtered_df["name"].str.contains(search_term, case=False, na=False)]

    if filtered_df.empty:
        st.warning("⚠️ No matching product found for that search term.")
    else:
        med_options = filtered_df.apply(
            lambda r: f"{r['name']}"
                      + (f"  ·  {r['medicine_type']}" if r.get("medicine_type") else "")
                      + f"  ·  Available: {int(r['stock'])} pcs",
            axis=1,
        ).tolist()
        id_lookup = dict(zip(med_options, filtered_df["id"]))

        colp1, colp2, colp3, colp4 = st.columns([2.4, 0.9, 1.1, 0.8])
        with colp1:
            selected_option = st.selectbox("Product", med_options, label_visibility="collapsed")
        selected_id = id_lookup[selected_option]
        selected_row = filtered_df[filtered_df["id"] == selected_id].iloc[0]
        available_stock = int(selected_row["stock"])

        with colp2:
            add_qty = st.number_input(
                "Qty", min_value=1, max_value=available_stock, value=1, step=1, label_visibility="collapsed"
            )
        with colp3:
            add_price = st.number_input(
                "Price/unit", min_value=0.0, value=0.0, step=0.5, format="%.2f", label_visibility="collapsed"
            )
        with colp4:
            add_clicked = st.button("➕ Add", use_container_width=True)

        if add_clicked:
            if add_price <= 0:
                st.error("⚠️ Please enter a sales price greater than 0.")
            elif add_qty > available_stock:
                st.error("⚠️ Quantity exceeds available stock.")
            else:
                existing_idx = next(
                    (i for i, it in enumerate(st.session_state.sale_cart) if it["medicine_id"] == selected_id), None
                )
                if existing_idx is not None:
                    new_qty = st.session_state.sale_cart[existing_idx]["qty"] + add_qty
                    if new_qty > available_stock:
                        st.error("⚠️ Total quantity in cart exceeds available stock.")
                    else:
                        st.session_state.sale_cart[existing_idx]["qty"] = new_qty
                        st.session_state.sale_cart[existing_idx]["unit_price"] = add_price
                        st.session_state.sale_cart[existing_idx]["subtotal"] = new_qty * add_price
                else:
                    st.session_state.sale_cart.append(
                        {
                            "medicine_id": int(selected_id),
                            "name": selected_row["name"],
                            "medicine_type": selected_row.get("medicine_type", ""),
                            "qty": int(add_qty),
                            "unit_price": float(add_price),
                            "subtotal": int(add_qty) * float(add_price),
                        }
                    )
                st.rerun()

    # ---------- Cart ----------
    st.markdown("---")
    st.markdown("##### 🧺 Cart")

    if not st.session_state.sale_cart:
        st.info("Cart is empty. Search and add products above.")
        return

    cart_df = pd.DataFrame(st.session_state.sale_cart)
    cart_display = cart_df.rename(
        columns={"name": "Product", "medicine_type": "Type", "qty": "Qty", "unit_price": "Unit Price", "subtotal": "Subtotal"}
    )[["Product", "Type", "Qty", "Unit Price", "Subtotal"]]
    st.dataframe(cart_display, use_container_width=True, hide_index=True)

    rcol1, rcol2 = st.columns([3, 1])
    remove_options = [f"{it['name']} (Qty: {it['qty']})" for it in st.session_state.sale_cart]
    with rcol1:
        item_to_remove = st.selectbox("Remove an item", remove_options, label_visibility="collapsed")
    with rcol2:
        if st.button("🗑️ Remove", use_container_width=True):
            idx = remove_options.index(item_to_remove)
            st.session_state.sale_cart.pop(idx)
            st.rerun()

    subtotal_amount = float(cart_df["subtotal"].sum())

    # ---------- Discount & Customer ----------
    st.markdown("---")
    st.markdown("##### 💳 Billing Details")

    discount = st.number_input("Discount (Flat Amount TK)", min_value=0.0, max_value=subtotal_amount, step=1.0)
    payable_amount = subtotal_amount - discount

    ccol1, ccol2 = st.columns(2)
    with ccol1:
        customer_name = st.text_input("Customer Name")
    with ccol2:
        customer_phone = st.text_input("Customer Phone")

    payment_mode = st.radio("Payment Mode *", ["Cash", "Credit"], horizontal=True)

    paid_amount = payable_amount
    if payment_mode == "Credit":
        paid_amount = st.number_input(
            "Amount Paid Now (TK)", min_value=0.0, max_value=float(payable_amount), value=0.0, step=1.0
        )
    due_amount = max(payable_amount - paid_amount, 0.0)

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Subtotal", fmt_money(subtotal_amount))
    m2.metric("Discount", fmt_money(discount))
    m3.metric("Total Payable", fmt_money(payable_amount))
    m4.metric("Due Amount", fmt_money(due_amount))

    if payment_mode == "Credit" and due_amount > 0 and (not customer_name.strip() or not customer_phone.strip()):
        st.warning("⚠️ Customer Name & Phone are required for Credit sales with a remaining due amount.")

    checkout_disabled = payment_mode == "Credit" and due_amount > 0 and (
        not customer_name.strip() or not customer_phone.strip()
    )

    st.markdown("---")
    if st.button("✅ Checkout & Generate Voucher", use_container_width=True, type="primary", disabled=checkout_disabled):
        # 1. Re-validate & deduct stock for every cart item
        stock_ok = True
        latest_meds_df = fetch_medicines()
        for item in st.session_state.sale_cart:
            current_row = latest_meds_df[latest_meds_df["id"] == item["medicine_id"]]
            if current_row.empty:
                st.error(f"⚠️ '{item['name']}' no longer exists in inventory.")
                stock_ok = False
                continue
            current_stock = int(current_row.iloc[0]["stock"])
            new_stock = current_stock - item["qty"]
            if new_stock < 0:
                st.error(f"⚠️ Not enough stock for {item['name']}.")
                stock_ok = False
                continue
            update_medicine_stock(item["medicine_id"], new_stock)

        if stock_ok:
            voucher_no = generate_voucher_no()
            sale_payload = {
                "voucher_no": voucher_no,
                "sale_date": date.today().isoformat(),
                "customer_name": customer_name.strip() if customer_name.strip() else "Walk-in Customer",
                "customer_phone": customer_phone.strip(),
                "items": json.dumps(st.session_state.sale_cart),
                "subtotal": subtotal_amount,
                "discount": discount,
                "total_amount": payable_amount,
                "payment_mode": payment_mode,
                "paid_amount": paid_amount if payment_mode == "Credit" else payable_amount,
                "due_amount": due_amount if payment_mode == "Credit" else 0.0,
            }
            saved_sale = insert_sale(sale_payload)

            ledger_ok = True
            if payment_mode == "Credit" and due_amount > 0:
                ledger_ok = upsert_customer_due(sale_payload["customer_name"], customer_phone.strip(), due_amount)

            if saved_sale and ledger_ok:
                clear_all_caches()
                st.session_state.last_voucher = sale_payload
                st.session_state.sale_cart = []
                st.success(f"✅ Sale completed! Voucher No: **{voucher_no}**")
                st.balloons()
                st.rerun()
            else:
                st.error("❌ Sale could not be fully recorded. Please check the errors above.")

    # ---------- Show the last generated voucher ----------
    if st.session_state.last_voucher:
        st.markdown("---")
        st.markdown("##### 🧾 Voucher / Receipt")
        st.markdown(render_voucher_html(st.session_state.last_voucher), unsafe_allow_html=True)
        if st.button("✖️ Clear Voucher Preview", use_container_width=True):
            st.session_state.last_voucher = None
            st.rerun()


# =================================================================================
# MODULE 3: INVENTORY REPORTS
# =================================================================================
def render_inventory_reports():
    st.subheader("📦 Inventory Reports")

    meds_df = fetch_medicines()

    if meds_df.empty:
        st.info("ℹ️ No inventory records found yet. Add stock from the Purchase module first.")
        return

    total_items = meds_df.shape[0]
    total_units = int(meds_df["stock"].sum())
    low_stock_df = meds_df[meds_df["stock"] < 10]

    m1, m2, m3 = st.columns(3)
    m1.metric("💊 Total Medicine Types", f"{total_items}")
    m2.metric("📦 Total Stock (Pcs)", f"{total_units} pcs")
    m3.metric("⚠️ Low Stock Items (<10)", f"{low_stock_df.shape[0]}")

    st.markdown("---")

    col1, col2 = st.columns([2, 1])
    with col1:
        search_term = st.text_input("🔍 Search by Medicine Name")
    with col2:
        type_options = ["All Types"] + sorted([t for t in meds_df["medicine_type"].unique().tolist() if t])
        type_filter = st.selectbox("Filter by Type", type_options)

    display_df = meds_df.copy()
    if search_term:
        display_df = display_df[display_df["name"].str.contains(search_term, case=False, na=False)]
    if type_filter != "All Types":
        display_df = display_df[display_df["medicine_type"] == type_filter]

    show_df = display_df.rename(
        columns={"name": "Medicine Name", "medicine_type": "Type", "stock": "Stock (Pcs)"}
    )
    columns_to_show = ["Medicine Name", "Type", "Stock (Pcs)"]
    if "created_at" in display_df.columns:
        show_df["created_at"] = pd.to_datetime(display_df["created_at"], errors="coerce")
        show_df["Added On"] = show_df["created_at"].dt.strftime("%Y-%m-%d")
        columns_to_show.append("Added On")

    st.dataframe(
        show_df[columns_to_show].sort_values("Medicine Name"),
        use_container_width=True,
        hide_index=True,
    )

    if not low_stock_df.empty:
        st.markdown("---")
        st.markdown("##### ⚠️ Low Stock Alert (below 10 units)")
        st.dataframe(
            low_stock_df.rename(columns={"name": "Medicine Name", "medicine_type": "Type", "stock": "Stock (Pcs)"})[
                ["Medicine Name", "Type", "Stock (Pcs)"]
            ].sort_values("Stock (Pcs)"),
            use_container_width=True,
            hide_index=True,
        )


# =================================================================================
# MODULE 4: SUPPLIER LEDGER (CREDIT DUE)
# =================================================================================
def render_supplier_ledger():
    st.subheader("🧾 Supplier Ledger — Credit Due")
    st.caption("Track how much your pharmacy owes each supplier for credit purchases, and settle payments.")

    ledger_df = fetch_supplier_ledger()
    purchases_df = fetch_purchases()

    total_outstanding = 0.0 if ledger_df.empty else float(ledger_df["total_due"].sum())
    m1, m2 = st.columns(2)
    m1.metric("🏭 Suppliers with Credit Due", f"{ledger_df[ledger_df['total_due'] > 0].shape[0] if not ledger_df.empty else 0}")
    m2.metric("💳 Total Outstanding Credit", fmt_money(total_outstanding))

    st.markdown("---")
    st.markdown("##### 🔍 Search Supplier")
    search_term = st.text_input(
        "Search by Supplier Name", label_visibility="collapsed", placeholder="Search by Supplier Name"
    )

    display_df = ledger_df.copy()
    if not display_df.empty and search_term:
        display_df = display_df[display_df["supplier_name"].str.contains(search_term, case=False, na=False)]

    if display_df.empty:
        st.info("No supplier ledger records found.")
    else:
        show_df = display_df.rename(columns={"supplier_name": "Supplier", "total_due": "Outstanding Due (TK)"})
        st.dataframe(show_df[["Supplier", "Outstanding Due (TK)"]], use_container_width=True, hide_index=True)

    st.markdown("---")
    st.markdown("##### 💵 Pay Supplier (Settle Credit)")

    due_suppliers = ledger_df[ledger_df["total_due"] > 0] if not ledger_df.empty else pd.DataFrame()

    if due_suppliers.empty:
        st.success("✅ No outstanding credit with any supplier.")
    else:
        options = due_suppliers.apply(
            lambda r: f"{r['supplier_name']}  |  Due: {fmt_money(r['total_due'])}  [ID:{r['id']}]", axis=1
        ).tolist()
        selected = st.selectbox("Select Supplier", options)
        selected_id = int(selected.split("[ID:")[1].replace("]", ""))
        selected_row = due_suppliers[due_suppliers["id"] == selected_id].iloc[0]

        st.metric("Current Outstanding Due", fmt_money(selected_row["total_due"]))

        paid_now = st.number_input(
            "Amount Paid Now *", min_value=0.0, max_value=float(selected_row["total_due"]), step=1.0
        )

        if st.button("💵 Record Payment to Supplier", use_container_width=True, type="primary"):
            if paid_now <= 0:
                st.error("⚠️ Enter a valid amount greater than 0.")
            else:
                new_due = float(selected_row["total_due"]) - paid_now
                if pay_supplier(selected_id, new_due):
                    clear_all_caches()
                    st.success(f"✅ Payment of {fmt_money(paid_now)} recorded! Remaining due: {fmt_money(new_due)}")
                    st.balloons()

    st.markdown("---")
    st.markdown("##### 📜 Credit Purchase History")
    if purchases_df.empty:
        st.info("No purchase records yet.")
    else:
        credit_df = purchases_df[purchases_df["payment_type"] == "Credit"].copy()
        if credit_df.empty:
            st.info("No credit purchases recorded yet.")
        else:
            credit_df["purchase_date"] = credit_df["purchase_date"].dt.strftime("%Y-%m-%d")
            credit_df = credit_df.rename(
                columns={
                    "purchase_date": "Date",
                    "medicine_name": "Medicine",
                    "supplier_name": "Supplier",
                    "quantity": "Qty",
                    "total_amount": "Total (TK)",
                    "paid_amount": "Paid (TK)",
                    "due_amount": "Due (TK)",
                }
            )
            st.dataframe(
                credit_df[["Date", "Medicine", "Supplier", "Qty", "Total (TK)", "Paid (TK)", "Due (TK)"]],
                use_container_width=True,
                hide_index=True,
            )


# =================================================================================
# MODULE 5: CUSTOMER LEDGER (CREDIT DUE)
# =================================================================================
def render_customer_ledger():
    st.subheader("📗 Customer Ledger — Credit Due")
    st.caption("Track how much each customer owes the pharmacy for credit sales, and collect payments.")

    ledger_df = fetch_customer_ledger()

    total_outstanding = 0.0 if ledger_df.empty else float(ledger_df["total_due"].sum())
    m1, m2 = st.columns(2)
    m1.metric(
        "🙍 Customers with Due",
        f"{ledger_df[ledger_df['total_due'] > 0].shape[0] if not ledger_df.empty else 0}",
    )
    m2.metric("💳 Total Outstanding Due", fmt_money(total_outstanding))

    st.markdown("---")
    st.markdown("##### 🔍 Search Customer")
    search_term = st.text_input(
        "Search by Name or Phone", label_visibility="collapsed", placeholder="Search by Name or Phone"
    )

    display_df = ledger_df.copy()
    if not display_df.empty and search_term:
        mask = display_df["customer_name"].str.contains(search_term, case=False, na=False) | display_df[
            "customer_phone"
        ].str.contains(search_term, case=False, na=False)
        display_df = display_df[mask]

    if display_df.empty:
        st.info("No customer ledger records found.")
    else:
        show_df = display_df.rename(
            columns={"customer_name": "Customer", "customer_phone": "Phone", "total_due": "Due (TK)"}
        )
        st.dataframe(show_df[["Customer", "Phone", "Due (TK)"]], use_container_width=True, hide_index=True)

    st.markdown("---")
    st.markdown("##### 💵 Receive Payment")

    due_customers = ledger_df[ledger_df["total_due"] > 0] if not ledger_df.empty else pd.DataFrame()

    if due_customers.empty:
        st.success("✅ No pending dues from any customer.")
    else:
        options = due_customers.apply(
            lambda r: f"{r['customer_name']} ({r['customer_phone']})  |  Due: {fmt_money(r['total_due'])}  [ID:{r['id']}]",
            axis=1,
        ).tolist()
        selected = st.selectbox("Select Customer", options)
        selected_id = int(selected.split("[ID:")[1].replace("]", ""))
        selected_row = due_customers[due_customers["id"] == selected_id].iloc[0]

        st.metric("Current Due", fmt_money(selected_row["total_due"]))

        paid_now = st.number_input(
            "Amount Paid Now *", min_value=0.0, max_value=float(selected_row["total_due"]), step=1.0
        )

        if st.button("💵 Collect Due", use_container_width=True, type="primary"):
            if paid_now <= 0:
                st.error("⚠️ Enter a valid amount greater than 0.")
            else:
                new_due = float(selected_row["total_due"]) - paid_now
                if receive_customer_payment(selected_id, new_due):
                    clear_all_caches()
                    st.success(f"✅ Payment of {fmt_money(paid_now)} recorded! Remaining due: {fmt_money(new_due)}")
                    st.balloons()


# =================================================================================
# MODULE 6: SALES HISTORY / VOUCHERS
# =================================================================================
def render_sales_history():
    st.subheader("🧾 Sales History / Vouchers")

    sales_df = fetch_sales()
    if sales_df.empty:
        st.info("No sales recorded yet.")
        return

    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input("From Date", value=date.today() - pd.Timedelta(days=30))
    with col2:
        end_date = st.date_input("To Date", value=date.today())

    filtered_df = sales_df[
        (sales_df["sale_date"].dt.date >= start_date) & (sales_df["sale_date"].dt.date <= end_date)
    ]

    if filtered_df.empty:
        st.info("No sales found in the selected date range.")
        return

    total_sales = float(filtered_df["total_amount"].sum())
    total_paid = float(filtered_df["paid_amount"].sum())
    total_due = float(filtered_df["due_amount"].sum())

    m1, m2, m3 = st.columns(3)
    m1.metric("Total Sales", fmt_money(total_sales))
    m2.metric("Total Collected", fmt_money(total_paid))
    m3.metric("Total Due", fmt_money(total_due))

    st.markdown("---")

    show_df = filtered_df.copy()
    show_df["sale_date"] = show_df["sale_date"].dt.strftime("%Y-%m-%d")
    show_df = show_df.rename(
        columns={
            "voucher_no": "Voucher No",
            "sale_date": "Date",
            "customer_name": "Customer",
            "customer_phone": "Phone",
            "payment_mode": "Payment",
            "total_amount": "Total (TK)",
            "paid_amount": "Paid (TK)",
            "due_amount": "Due (TK)",
        }
    )
    st.dataframe(
        show_df[["Voucher No", "Date", "Customer", "Phone", "Payment", "Total (TK)", "Paid (TK)", "Due (TK)"]],
        use_container_width=True,
        hide_index=True,
    )

    st.markdown("---")
    st.markdown("##### 🧾 View / Reprint a Voucher")
    voucher_options = filtered_df.apply(
        lambda r: f"{r['voucher_no']}  ·  {r['customer_name']}  ·  {fmt_money(r['total_amount'])}", axis=1
    ).tolist()
    voucher_lookup = dict(zip(voucher_options, filtered_df["voucher_no"]))
    selected_voucher_label = st.selectbox("Select a voucher", voucher_options)
    selected_voucher_no = voucher_lookup[selected_voucher_label]
    voucher_row = filtered_df[filtered_df["voucher_no"] == selected_voucher_no].iloc[0].to_dict()

    st.markdown(render_voucher_html(voucher_row), unsafe_allow_html=True)



# =================================================================================
# MODULE 7: SETTINGS
# =================================================================================
def render_settings():
    st.subheader("⚙️ Settings")

    st.markdown("##### 🔌 Connection Status")
    try:
        supabase.table("medicines").select("id").limit(1).execute()
        st.success("✅ Connected to Supabase successfully.")
    except Exception as e:
        st.error(f"❌ Supabase connection issue: {e}")

    st.markdown("---")
    st.markdown("##### 🗂️ Master Medicine Catalogue")
    master_df = fetch_master_medicines()
    if master_df.empty:
        st.info("No records found in `master_medicines`.")
    else:
        st.write(f"Total catalogue entries: **{master_df.shape[0]}**")
        with st.expander("View Master Catalogue"):
            cols = [c for c in ["brand_name", "generic_name", "company_name", "form"] if c in master_df.columns]
            st.dataframe(master_df[cols], use_container_width=True, hide_index=True)

    st.markdown("---")
    st.markdown("##### 🧹 Cache Management")
    st.caption("Data is cached for 20 seconds for faster performance. Force a refresh if you just updated data elsewhere.")
    if st.button("🔄 Clear Cache & Refresh Now", use_container_width=True):
        clear_all_caches()
        st.success("✅ Cache cleared successfully.")
        st.rerun()

    st.markdown("---")
    st.markdown("##### ℹ️ About")
    st.caption("Med Life Pharmacy ERP · Built with Streamlit + Supabase")


# =================================================================================
# ROUTER
# =================================================================================
if page == "➕ Purchase / Stock Entry":
    render_purchase()
elif page == "🛒 Sales Module":
    render_sales()
elif page == "📦 Inventory Reports":
    render_inventory_reports()
elif page == "🧾 Supplier Ledger (Credit Due)":
    render_supplier_ledger()
elif page == "📗 Customer Ledger (Credit Due)":
    render_customer_ledger()
elif page == "🧾 Sales History / Vouchers":
    render_sales_history()
elif page == "⚙️ Settings":
    render_settings()
