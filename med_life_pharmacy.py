"""
=================================================================================
 MED LIFE PHARMACY - Pharmacy Management System
 Location: Chachkoir, Khalifa Para, Gurudaspur, Natore
 Built with: Streamlit + Supabase (supabase-py)
=================================================================================

SUPABASE TABLE SCHEMA (create these tables in your Supabase project before running):

1) medicines
   - id                bigint, primary key, identity
   - name               text, not null
   - generic_name       text
   - shelf_no           text
   - purchase_price     numeric, not null, default 0
   - selling_price      numeric, not null, default 0
   - stock_quantity     numeric, not null, default 0
   - expiry_date        date
   - created_at         timestamptz, default now()

2) sales
   - id                bigint, primary key, identity
   - sale_date          date, not null
   - customer_name      text
   - customer_phone     text
   - items              text   (JSON string of cart items)
   - total_amount       numeric, not null
   - discount           numeric, default 0
   - paid_amount        numeric, not null
   - due_amount         numeric, default 0
   - created_at         timestamptz, default now()

3) customer_ledger
   - id                bigint, primary key, identity
   - customer_name      text, not null
   - customer_phone     text, not null, unique
   - total_due          numeric, not null, default 0
   - updated_at         timestamptz, default now()

STREAMLIT SECRETS (.streamlit/secrets.toml):
   SUPABASE_URL = "https://xxxxxxxxxxxx.supabase.co"
   SUPABASE_KEY = "your-supabase-anon-or-service-key"
=================================================================================
"""

import json
from datetime import date, datetime, timedelta

import pandas as pd
import streamlit as st
from supabase import create_client, Client


# =================================================================================
# PAGE CONFIG - MUST BE FIRST STREAMLIT COMMAND
# =================================================================================
st.set_page_config(
    page_title="Med Life Pharmacy",
    page_icon="💊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# =================================================================================
# MOBILE-FRIENDLY CSS
# =================================================================================
st.markdown(
    """
    <style>
        /* Bigger, touch friendly buttons */
        div.stButton > button, div.stFormSubmitButton > button {
            height: 3em;
            font-size: 1.05rem;
            font-weight: 600;
            border-radius: 10px;
        }
        /* Reduce top padding for more usable space on small screens */
        .block-container {
            padding-top: 1.2rem;
            padding-bottom: 3rem;
        }
        /* Readable metric labels */
        div[data-testid="stMetricValue"] {
            font-size: 1.5rem;
        }
        /* Header banner */
        .pharmacy-header {
            background: linear-gradient(90deg, #0f766e, #14b8a6);
            padding: 14px 18px;
            border-radius: 12px;
            color: white;
            margin-bottom: 1rem;
        }
        .pharmacy-header h1 {
            margin: 0;
            font-size: 1.5rem;
            color: white;
        }
        .pharmacy-header p {
            margin: 0;
            font-size: 0.9rem;
            opacity: 0.9;
        }
        /* Make tabs bigger for touch */
        button[data-baseweb="tab"] {
            font-size: 1rem;
            padding: 10px 14px;
        }
        @media (max-width: 640px) {
            div[data-testid="stMetricValue"] { font-size: 1.15rem; }
            .pharmacy-header h1 { font-size: 1.2rem; }
        }
    </style>
    """,
    unsafe_allow_html=True,
)

CURRENCY = "৳"


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
    st.error(f"❌ Could not connect to Supabase. Check your secrets.toml. Details: {e}")
    st.stop()


# =================================================================================
# DATA ACCESS HELPERS (all wrapped in try/except)
# =================================================================================
@st.cache_data(ttl=15, show_spinner=False)
def fetch_medicines() -> pd.DataFrame:
    try:
        res = supabase.table("medicines").select("*").order("name").execute()
        df = pd.DataFrame(res.data)
        if not df.empty:
            df["expiry_date"] = pd.to_datetime(df["expiry_date"], errors="coerce")
            for col in ["purchase_price", "selling_price", "stock_quantity"]:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        return df
    except Exception as e:
        st.error(f"❌ Failed to fetch medicines: {e}")
        return pd.DataFrame()


@st.cache_data(ttl=15, show_spinner=False)
def fetch_sales() -> pd.DataFrame:
    try:
        res = supabase.table("sales").select("*").order("sale_date", desc=True).execute()
        df = pd.DataFrame(res.data)
        if not df.empty:
            df["sale_date"] = pd.to_datetime(df["sale_date"], errors="coerce")
            for col in ["total_amount", "discount", "paid_amount", "due_amount"]:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        return df
    except Exception as e:
        st.error(f"❌ Failed to fetch sales history: {e}")
        return pd.DataFrame()


@st.cache_data(ttl=15, show_spinner=False)
def fetch_customers() -> pd.DataFrame:
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
    fetch_medicines.clear()
    fetch_sales.clear()
    fetch_customers.clear()


def insert_medicine(payload: dict) -> bool:
    try:
        supabase.table("medicines").insert(payload).execute()
        return True
    except Exception as e:
        st.error(f"❌ Failed to add medicine: {e}")
        return False


def update_medicine(med_id, payload: dict) -> bool:
    try:
        supabase.table("medicines").update(payload).eq("id", med_id).execute()
        return True
    except Exception as e:
        st.error(f"❌ Failed to update medicine: {e}")
        return False


def delete_medicine(med_id) -> bool:
    try:
        supabase.table("medicines").delete().eq("id", med_id).execute()
        return True
    except Exception as e:
        st.error(f"❌ Failed to delete medicine: {e}")
        return False


def insert_sale(payload: dict) -> bool:
    try:
        supabase.table("sales").insert(payload).execute()
        return True
    except Exception as e:
        st.error(f"❌ Failed to record sale: {e}")
        return False


def upsert_customer_due(name: str, phone: str, delta_due: float) -> bool:
    """Add delta_due to an existing customer's total_due, or create a new record."""
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


def receive_payment(customer_id, new_due: float) -> bool:
    try:
        supabase.table("customer_ledger").update(
            {"total_due": new_due, "updated_at": datetime.now().isoformat()}
        ).eq("id", customer_id).execute()
        return True
    except Exception as e:
        st.error(f"❌ Failed to record payment: {e}")
        return False


def fmt_money(val) -> str:
    try:
        return f"{CURRENCY}{float(val):,.2f}"
    except Exception:
        return f"{CURRENCY}0.00"


# =================================================================================
# HEADER
# =================================================================================
st.markdown(
    """
    <div class="pharmacy-header">
        <h1>💊 Med Life Pharmacy</h1>
        <p>📍 Chachkoir, Khalifa Para, Gurudaspur, Natore</p>
    </div>
    """,
    unsafe_allow_html=True,
)

# =================================================================================
# SESSION STATE INIT
# =================================================================================
if "cart" not in st.session_state:
    st.session_state.cart = []  # list of dicts: id, name, price, qty, subtotal

# =================================================================================
# SIDEBAR NAVIGATION
# =================================================================================
st.sidebar.title("📋 Menu")
page = st.sidebar.radio(
    "Navigate",
    [
        "📊 Dashboard (ড্যাশবোর্ড)",
        "💊 Inventory & Purchase (ইনভেন্টরি ও প্রোডাক্ট এন্ট্রি)",
        "🛒 POS / Billing (বিক্রি ও বিল)",
        "📗 Customer Due Ledger (বাকির হিসাব)",
        "📈 Sales History (বিক্রির রিপোর্ট)",
    ],
    label_visibility="collapsed",
)

if st.sidebar.button("🔄 Refresh Data", use_container_width=True):
    clear_all_caches()
    st.rerun()


# =================================================================================
# PAGE 1: DASHBOARD
# =================================================================================
def render_dashboard():
    st.subheader("📊 Dashboard")

    meds_df = fetch_medicines()
    sales_df = fetch_sales()
    cust_df = fetch_customers()

    total_medicines = 0 if meds_df.empty else meds_df.shape[0]
    total_due = 0.0 if cust_df.empty else float(cust_df["total_due"].sum())

    today = date.today()
    if not sales_df.empty:
        today_sales_df = sales_df[sales_df["sale_date"].dt.date == today]
        today_sales = float(today_sales_df["total_amount"].sum()) if not today_sales_df.empty else 0.0
    else:
        today_sales = 0.0

    c1, c2, c3 = st.columns(3)
    c1.metric("💊 Total Active Medicines", f"{total_medicines}")
    c2.metric("📗 Total Dues Receivable", fmt_money(total_due))
    c3.metric("🛒 Today's Total Sales", fmt_money(today_sales))

    st.markdown("---")

    col_a, col_b = st.columns(2)

    with col_a:
        st.markdown("##### ⚠️ Low Stock Alert (Qty < 10)")
        if not meds_df.empty:
            low_stock = meds_df[meds_df["stock_quantity"] < 10][
                ["name", "generic_name", "stock_quantity", "shelf_no"]
            ].sort_values("stock_quantity")
            if not low_stock.empty:
                st.dataframe(low_stock, use_container_width=True, hide_index=True)
            else:
                st.success("✅ No low stock items.")
        else:
            st.info("No medicine data available yet.")

    with col_b:
        st.markdown("##### ⏳ Expiring Soon (within 30 days)")
        if not meds_df.empty:
            cutoff = pd.Timestamp(today + timedelta(days=30))
            expiring = meds_df[
                (meds_df["expiry_date"].notna()) & (meds_df["expiry_date"] <= cutoff)
            ][["name", "generic_name", "expiry_date", "stock_quantity"]].sort_values("expiry_date")
            if not expiring.empty:
                expiring = expiring.copy()
                expiring["expiry_date"] = expiring["expiry_date"].dt.strftime("%Y-%m-%d")
                st.dataframe(expiring, use_container_width=True, hide_index=True)
            else:
                st.success("✅ No medicines expiring soon.")
        else:
            st.info("No medicine data available yet.")


# =================================================================================
# PAGE 2: INVENTORY & PURCHASE
# =================================================================================
def render_inventory():
    st.subheader("💊 Inventory & Purchase Management")

    tab1, tab2, tab3 = st.tabs(
        ["➕ Add New Product", "📦 Restock Existing", "📋 Stock View & Search"]
    )

    # ---------------- TAB 1: ADD NEW PRODUCT ----------------
    with tab1:
        st.markdown("##### নতুন ওষুধ এন্ট্রি")
        with st.form("add_new_product_form", clear_on_submit=True):
            name = st.text_input("Medicine Name *")
            generic_name = st.text_input("Generic Name")
            shelf_no = st.text_input("Shelf / Rack No")

            col1, col2 = st.columns(2)
            with col1:
                purchase_price = st.number_input("Purchase Price (ক্রয়মূল্য) *", min_value=0.0, step=0.5)
                purchased_qty = st.number_input("Purchased Quantity *", min_value=0.0, step=1.0)
            with col2:
                selling_price = st.number_input("Selling Price (বিক্রয়মূল্য) *", min_value=0.0, step=0.5)
                expiry_date_val = st.date_input("Expiry Date *", value=date.today() + timedelta(days=180))

            submitted = st.form_submit_button("💾 Save Product", use_container_width=True, type="primary")

            if submitted:
                if not name.strip():
                    st.error("⚠️ Medicine Name is required.")
                elif purchased_qty <= 0:
                    st.error("⚠️ Purchased Quantity must be greater than 0.")
                else:
                    payload = {
                        "name": name.strip(),
                        "generic_name": generic_name.strip(),
                        "shelf_no": shelf_no.strip(),
                        "purchase_price": purchase_price,
                        "selling_price": selling_price,
                        "stock_quantity": purchased_qty,
                        "expiry_date": expiry_date_val.isoformat(),
                    }
                    if insert_medicine(payload):
                        clear_all_caches()
                        st.success(f"✅ '{name}' added to inventory successfully!")
                        st.rerun()

    # ---------------- TAB 2: RESTOCK EXISTING ----------------
    with tab2:
        st.markdown("##### বিদ্যমান ওষুধের স্টক বাড়ানো")
        meds_df = fetch_medicines()

        if meds_df.empty:
            st.info("No medicines found. Please add a product first.")
        else:
            meds_df = meds_df.sort_values("name")
            options = meds_df.apply(
                lambda r: f"{r['name']} (Stock: {int(r['stock_quantity'])})  [ID:{r['id']}]", axis=1
            ).tolist()
            selected_option = st.selectbox("Select Medicine to Restock *", options)

            selected_id = int(selected_option.split("[ID:")[1].replace("]", ""))
            selected_row = meds_df[meds_df["id"] == selected_id].iloc[0]

            info_col1, info_col2, info_col3 = st.columns(3)
            info_col1.metric("Current Stock", f"{int(selected_row['stock_quantity'])}")
            info_col2.metric("Current Purchase Price", fmt_money(selected_row["purchase_price"]))
            info_col3.metric("Current Selling Price", fmt_money(selected_row["selling_price"]))

            with st.form("restock_form", clear_on_submit=True):
                additional_qty = st.number_input("Additional Quantity Purchased *", min_value=0.0, step=1.0)

                col1, col2 = st.columns(2)
                with col1:
                    updated_purchase_price = st.number_input(
                        "Updated Purchase Price", min_value=0.0, value=float(selected_row["purchase_price"]), step=0.5
                    )
                with col2:
                    updated_selling_price = st.number_input(
                        "Updated Selling Price", min_value=0.0, value=float(selected_row["selling_price"]), step=0.5
                    )

                update_expiry = st.checkbox("Update Expiry Date?")
                new_expiry = None
                if update_expiry:
                    default_expiry = (
                        selected_row["expiry_date"].date()
                        if pd.notna(selected_row["expiry_date"])
                        else date.today()
                    )
                    new_expiry = st.date_input("New Expiry Date", value=default_expiry)

                restock_submit = st.form_submit_button("📦 Update Stock", use_container_width=True, type="primary")

                if restock_submit:
                    if additional_qty <= 0:
                        st.error("⚠️ Additional Quantity must be greater than 0.")
                    else:
                        new_stock = float(selected_row["stock_quantity"]) + additional_qty
                        payload = {
                            "stock_quantity": new_stock,
                            "purchase_price": updated_purchase_price,
                            "selling_price": updated_selling_price,
                        }
                        if update_expiry and new_expiry:
                            payload["expiry_date"] = new_expiry.isoformat()

                        if update_medicine(selected_id, payload):
                            clear_all_caches()
                            st.success(f"✅ Stock updated! New quantity: {int(new_stock)}")
                            st.rerun()

    # ---------------- TAB 3: STOCK VIEW & SEARCH ----------------
    with tab3:
        st.markdown("##### ইনভেন্টরি দেখা ও সংশোধন")
        meds_df = fetch_medicines()

        if meds_df.empty:
            st.info("No medicines in inventory yet.")
        else:
            search_term = st.text_input("🔍 Search by Medicine Name or Generic Name")
            display_df = meds_df.copy()
            if search_term:
                mask = display_df["name"].str.contains(search_term, case=False, na=False) | display_df[
                    "generic_name"
                ].str.contains(search_term, case=False, na=False)
                display_df = display_df[mask]

            show_df = display_df.copy()
            show_df["expiry_date"] = show_df["expiry_date"].dt.strftime("%Y-%m-%d")
            st.dataframe(
                show_df[
                    ["id", "name", "generic_name", "shelf_no", "purchase_price", "selling_price",
                     "stock_quantity", "expiry_date"]
                ],
                use_container_width=True,
                hide_index=True,
            )

            st.markdown("---")
            st.markdown("##### ✏️ Edit / 🗑️ Delete a Medicine")

            if not display_df.empty:
                edit_options = display_df.apply(lambda r: f"{r['name']}  [ID:{r['id']}]", axis=1).tolist()
                edit_selected = st.selectbox("Select Medicine to Edit or Delete", edit_options, key="edit_select")
                edit_id = int(edit_selected.split("[ID:")[1].replace("]", ""))
                edit_row = display_df[display_df["id"] == edit_id].iloc[0]

                with st.form("edit_medicine_form"):
                    e_name = st.text_input("Medicine Name", value=edit_row["name"])
                    e_generic = st.text_input("Generic Name", value=edit_row["generic_name"] or "")
                    e_shelf = st.text_input("Shelf / Rack No", value=edit_row["shelf_no"] or "")

                    ecol1, ecol2 = st.columns(2)
                    with ecol1:
                        e_purchase = st.number_input(
                            "Purchase Price", min_value=0.0, value=float(edit_row["purchase_price"]), step=0.5
                        )
                        e_qty = st.number_input(
                            "Stock Quantity", min_value=0.0, value=float(edit_row["stock_quantity"]), step=1.0
                        )
                    with ecol2:
                        e_selling = st.number_input(
                            "Selling Price", min_value=0.0, value=float(edit_row["selling_price"]), step=0.5
                        )
                        e_expiry_default = (
                            edit_row["expiry_date"].date() if pd.notna(edit_row["expiry_date"]) else date.today()
                        )
                        e_expiry = st.date_input("Expiry Date", value=e_expiry_default)

                    ecol_btn1, ecol_btn2 = st.columns(2)
                    with ecol_btn1:
                        update_btn = st.form_submit_button("💾 Update", use_container_width=True, type="primary")
                    with ecol_btn2:
                        delete_btn = st.form_submit_button("🗑️ Delete", use_container_width=True)

                    if update_btn:
                        payload = {
                            "name": e_name.strip(),
                            "generic_name": e_generic.strip(),
                            "shelf_no": e_shelf.strip(),
                            "purchase_price": e_purchase,
                            "selling_price": e_selling,
                            "stock_quantity": e_qty,
                            "expiry_date": e_expiry.isoformat(),
                        }
                        if update_medicine(edit_id, payload):
                            clear_all_caches()
                            st.success("✅ Medicine updated successfully!")
                            st.rerun()

                    if delete_btn:
                        if delete_medicine(edit_id):
                            clear_all_caches()
                            st.success("✅ Medicine deleted successfully!")
                            st.rerun()


# =================================================================================
# PAGE 3: POS / BILLING
# =================================================================================
def render_pos():
    st.subheader("🛒 POS / Billing")

    meds_df = fetch_medicines()

    if meds_df.empty:
        st.info("No medicines available. Please add products in the Inventory module first.")
        return

    in_stock_df = meds_df[meds_df["stock_quantity"] > 0].sort_values("name")
    if in_stock_df.empty:
        st.warning("⚠️ All medicines are out of stock.")
        return

    st.markdown("##### ➕ Add Item to Cart")
    med_options = in_stock_df.apply(
        lambda r: f"{r['name']}  |  Stock: {int(r['stock_quantity'])}  |  Price: {fmt_money(r['selling_price'])}  [ID:{r['id']}]",
        axis=1,
    ).tolist()

    col1, col2, col3 = st.columns([3, 1.3, 1])
    with col1:
        med_choice = st.selectbox("Select Medicine", med_options, label_visibility="collapsed")
    med_id = int(med_choice.split("[ID:")[1].replace("]", ""))
    med_row = in_stock_df[in_stock_df["id"] == med_id].iloc[0]

    with col2:
        sell_qty = st.number_input(
            "Qty", min_value=1.0, max_value=float(med_row["stock_quantity"]), value=1.0, step=1.0,
            label_visibility="collapsed",
        )
    with col3:
        add_clicked = st.button("➕ Add", use_container_width=True)

    if add_clicked:
        if sell_qty > med_row["stock_quantity"]:
            st.error("⚠️ Not enough stock available.")
        else:
            existing_idx = next(
                (i for i, item in enumerate(st.session_state.cart) if item["id"] == med_id), None
            )
            if existing_idx is not None:
                new_qty = st.session_state.cart[existing_idx]["qty"] + sell_qty
                if new_qty > med_row["stock_quantity"]:
                    st.error("⚠️ Total quantity exceeds available stock.")
                else:
                    st.session_state.cart[existing_idx]["qty"] = new_qty
                    st.session_state.cart[existing_idx]["subtotal"] = new_qty * med_row["selling_price"]
            else:
                st.session_state.cart.append(
                    {
                        "id": int(med_row["id"]),
                        "name": med_row["name"],
                        "price": float(med_row["selling_price"]),
                        "qty": float(sell_qty),
                        "subtotal": float(sell_qty) * float(med_row["selling_price"]),
                    }
                )
            st.rerun()

    st.markdown("---")
    st.markdown("##### 🧾 Current Cart")

    if not st.session_state.cart:
        st.info("Cart is empty. Add medicines above.")
    else:
        cart_df = pd.DataFrame(st.session_state.cart)
        cart_display = cart_df.rename(
            columns={"name": "Medicine", "price": "Unit Price", "qty": "Qty", "subtotal": "Subtotal"}
        )[["Medicine", "Unit Price", "Qty", "Subtotal"]]
        st.dataframe(cart_display, use_container_width=True, hide_index=True)

        remove_options = [f"{item['name']} (Qty: {item['qty']})" for item in st.session_state.cart]
        rcol1, rcol2 = st.columns([3, 1])
        with rcol1:
            item_to_remove = st.selectbox("Remove an item", remove_options, label_visibility="collapsed")
        with rcol2:
            if st.button("🗑️ Remove", use_container_width=True):
                idx = remove_options.index(item_to_remove)
                st.session_state.cart.pop(idx)
                st.rerun()

        subtotal_amount = float(cart_df["subtotal"].sum())

        st.markdown("---")
        st.markdown("##### 💳 Payment Details")

        discount = st.number_input("Discount (Flat Amount TK)", min_value=0.0, max_value=subtotal_amount, step=1.0)
        payable_amount = subtotal_amount - discount

        cust_col1, cust_col2 = st.columns(2)
        with cust_col1:
            customer_name = st.text_input("Customer Name")
        with cust_col2:
            customer_phone = st.text_input("Customer Phone")

        paid_amount = st.number_input("Amount Paid *", min_value=0.0, value=float(payable_amount), step=1.0)
        due_amount = max(payable_amount - paid_amount, 0.0)

        m1, m2, m3 = st.columns(3)
        m1.metric("Total Payable", fmt_money(payable_amount))
        m2.metric("Amount Paid", fmt_money(paid_amount))
        m3.metric("Remaining Due", fmt_money(due_amount))

        if due_amount > 0 and (not customer_name.strip() or not customer_phone.strip()):
            st.warning("⚠️ Customer Name & Phone are required when there is a remaining due amount.")

        checkout_disabled = due_amount > 0 and (not customer_name.strip() or not customer_phone.strip())

        if st.button("✅ Checkout & Save Bill", use_container_width=True, type="primary", disabled=checkout_disabled):
            # 1. Deduct stock for each item
            stock_ok = True
            for item in st.session_state.cart:
                current_row = meds_df[meds_df["id"] == item["id"]]
                if current_row.empty:
                    stock_ok = False
                    continue
                current_stock = float(current_row.iloc[0]["stock_quantity"])
                new_stock = current_stock - item["qty"]
                if new_stock < 0:
                    st.error(f"⚠️ Not enough stock for {item['name']}.")
                    stock_ok = False
                    continue
                update_medicine(item["id"], {"stock_quantity": new_stock})

            if stock_ok:
                # 2. Insert sales record
                sale_payload = {
                    "sale_date": date.today().isoformat(),
                    "customer_name": customer_name.strip() if customer_name else "Walk-in Customer",
                    "customer_phone": customer_phone.strip() if customer_phone else "",
                    "items": json.dumps(st.session_state.cart),
                    "total_amount": payable_amount,
                    "discount": discount,
                    "paid_amount": paid_amount,
                    "due_amount": due_amount,
                }
                if insert_sale(sale_payload):
                    # 3. Update customer ledger if there is a due
                    if due_amount > 0:
                        upsert_customer_due(customer_name.strip(), customer_phone.strip(), due_amount)

                    clear_all_caches()
                    st.success("✅ Sale recorded successfully!")
                    st.session_state.cart = []
                    st.rerun()


# =================================================================================
# PAGE 4: CUSTOMER DUE LEDGER
# =================================================================================
def render_ledger():
    st.subheader("📗 Customer Due Ledger")

    cust_df = fetch_customers()

    st.markdown("##### 🔍 Search Customer")
    search_term = st.text_input("Search by Name or Phone", label_visibility="collapsed", placeholder="Search by Name or Phone")

    display_df = cust_df.copy()
    if not display_df.empty and search_term:
        mask = display_df["customer_name"].str.contains(search_term, case=False, na=False) | display_df[
            "customer_phone"
        ].str.contains(search_term, case=False, na=False)
        display_df = display_df[mask]

    if display_df.empty:
        st.info("No customer due records found.")
    else:
        show_df = display_df.rename(
            columns={"customer_name": "Customer Name", "customer_phone": "Phone", "total_due": "Total Due (TK)"}
        )[["Customer Name", "Phone", "Total Due (TK)"]]
        st.dataframe(show_df, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.markdown("##### 💰 Receive Payment")

    due_customers = cust_df[cust_df["total_due"] > 0] if not cust_df.empty else pd.DataFrame()

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
                if receive_payment(selected_id, new_due):
                    clear_all_caches()
                    st.success(f"✅ Payment of {fmt_money(paid_now)} recorded! Remaining due: {fmt_money(new_due)}")
                    st.rerun()


# =================================================================================
# PAGE 5: SALES HISTORY
# =================================================================================
def render_sales_history():
    st.subheader("📈 Sales History")

    sales_df = fetch_sales()

    if sales_df.empty:
        st.info("No sales recorded yet.")
        return

    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input("From Date", value=date.today() - timedelta(days=30))
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
            "sale_date": "Date",
            "customer_name": "Customer",
            "customer_phone": "Phone",
            "total_amount": "Total (TK)",
            "discount": "Discount (TK)",
            "paid_amount": "Paid (TK)",
            "due_amount": "Due (TK)",
        }
    )
    st.dataframe(
        show_df[["Date", "Customer", "Phone", "Total (TK)", "Discount (TK)", "Paid (TK)", "Due (TK)"]],
        use_container_width=True,
        hide_index=True,
    )

    st.markdown("---")
    st.markdown("##### 🧾 View Bill Details")
    bill_options = filtered_df.apply(
        lambda r: f"{r['sale_date'].strftime('%Y-%m-%d')} - {r['customer_name']} - {fmt_money(r['total_amount'])}  [ID:{r['id']}]",
        axis=1,
    ).tolist()
    bill_selected = st.selectbox("Select a bill to view items", bill_options)
    bill_id = int(bill_selected.split("[ID:")[1].replace("]", ""))
    bill_row = filtered_df[filtered_df["id"] == bill_id].iloc[0]

    try:
        items = json.loads(bill_row["items"]) if bill_row["items"] else []
        if items:
            items_df = pd.DataFrame(items).rename(
                columns={"name": "Medicine", "price": "Unit Price", "qty": "Qty", "subtotal": "Subtotal"}
            )[["Medicine", "Unit Price", "Qty", "Subtotal"]]
            st.dataframe(items_df, use_container_width=True, hide_index=True)
        else:
            st.info("No item details available for this bill.")
    except Exception:
        st.info("Could not parse item details for this bill.")


# =================================================================================
# ROUTER
# =================================================================================
if page.startswith("📊"):
    render_dashboard()
elif page.startswith("💊"):
    render_inventory()
elif page.startswith("🛒"):
    render_pos()
elif page.startswith("📗"):
    render_ledger()
elif page.startswith("📈"):
    render_sales_history()