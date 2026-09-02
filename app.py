import os
from datetime import datetime, timedelta
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import streamlit as st
import streamlit_authenticator as stauth

DB_FILE = "CakeHutERP_System.xlsx"
BACKUP_DIR = "backups"
PDF_DIR = "pdf_prints"

st.set_page_config(page_title="Cake Hut ERP", page_icon="🎂", layout="wide")

# -----------------------------------------------------------------------------
# DYNAMIC CREDENTIALS & SESSION STATE INITIALIZATION (Username: admin | Password: 12345)
# -----------------------------------------------------------------------------
if 'credentials' not in st.session_state:
    hashed_password = stauth.Hasher.hash('12345')
    st.session_state.credentials = {
        'usernames': {
            'admin': {
                'name': 'Bakery Owner',
                'email': 'admin@cakehut.com',
                'password': hashed_password
            }
        }
    }

authenticator = stauth.Authenticate(
    st.session_state.credentials,
    'cake_hut_cookie_name',
    'cake_hut_signature_key',
    cookie_expiry_days=30
)

authenticator.login('main', 'Login to Cake Hut ERP')

name = st.session_state.get('name')
authentication_status = st.session_state.get('authentication_status')
username = st.session_state.get('username')

if authentication_status == False:
    st.error('Username/password is incorrect')
elif authentication_status == None:
    st.warning('Please enter your username and password (Default Username: **admin** | Password: **12345**)')
elif authentication_status:
    authenticator.logout('Logout', 'sidebar')

    # -----------------------------------------------------------------------------
    # DIRECTORY INITIALIZATION
    # -----------------------------------------------------------------------------
    def init_directories():
        if not os.path.exists(BACKUP_DIR):
            os.makedirs(BACKUP_DIR)
        if not os.path.exists(PDF_DIR):
            os.makedirs(PDF_DIR)

    # -----------------------------------------------------------------------------
    # UNIT CONVERSION UTILITIES
    # -----------------------------------------------------------------------------
    def convert_to_base(qty, unit):
        u = str(unit).strip().lower()
        if u in ['kg', 'kilogram', 'kilograms']:
            return qty * 1000.0, 'grams'
        elif u in ['g', 'gram', 'grams']:
            return qty * 1.0, 'grams'
        elif u in ['l', 'liter', 'liters', 'litre', 'litres']:
            return qty * 1000.0, 'ml'
        elif u in ['ml', 'milliliter', 'milliliters']:
            return qty * 1.0, 'ml'
        elif u in ['dozen', 'dozens']:
            return qty * 12.0, 'pcs'
        else:
            return qty * 1.0, 'pcs'

    # -----------------------------------------------------------------------------
    # DATABASE INITIALIZATION, LOADING & AUTOMATED REPLACEMENT BACKUP
    # -----------------------------------------------------------------------------
    def init_db():
        init_directories()
        if not os.path.exists(DB_FILE):
            df_orders = pd.DataFrame(columns=[
                "Order ID", "Order Date", "Deadline Date", "Customer Name", "Address", "Phone Number", 
                "Cake Type", "Weight (Pounds)", "Delivery Charges", "Price Charged (Income)", 
                "Total Collection", "Notes", "Ingredients Cost", "Net Profit", "Status"
            ])
            
            df_purchases = pd.DataFrame(columns=[
                "Purchase ID", "Date", "Ingredient Name", "Quantity Bought", "Purchase Unit", 
                "Base Qty Bought", "Base Unit", "Total Price paid"
            ])
            
            df_recipes = pd.DataFrame(columns=[
                "Order ID Link", "Cake Type & Weight", "Ingredient Used", "Qty Used", 
                "Consumption Unit", "Base Qty Used", "Est. Cost of Qty"
            ])
            
            df_inventory = pd.DataFrame(columns=[
                "Ingredient Name", "Base Unit", "Total Purchased", "Total Utilized", "Current Stock Left"
            ])

            initial_ingredients = [
                ("Flour", "grams"), ("Sugar", "grams"), ("Eggs", "pcs"), ("Oil", "ml"), 
                ("Fondant", "grams"), ("Colors", "ml"), ("Cream", "ml"), ("Fox Balls", "pcs")
            ]
            for ing, b_unit in initial_ingredients:
                df_inventory = pd.concat([df_inventory, pd.DataFrame([{
                    "Ingredient Name": ing,
                    "Base Unit": b_unit,
                    "Total Purchased": 0.0,
                    "Total Utilized": 0.0,
                    "Current Stock Left": 0.0
                }])], ignore_index=True)

            with pd.ExcelWriter(DB_FILE, engine='openpyxl') as writer:
                df_orders.to_excel(writer, sheet_name="Orders", index=False)
                df_purchases.to_excel(writer, sheet_name="Material_Purchases", index=False)
                df_recipes.to_excel(writer, sheet_name="Recipes", index=False)
                df_inventory.to_excel(writer, sheet_name="Inventory", index=False)

    def check_and_replace_weekly_backup():
        init_directories()
        current_week_tag = datetime.now().strftime('%Y_%W')
        backup_filename = os.path.join(BACKUP_DIR, f"Backup_{current_week_tag}.xlsx")
        
        if os.path.exists(DB_FILE) and not os.path.exists(backup_filename):
            for f in os.listdir(BACKUP_DIR):
                if f.endswith(".xlsx"):
                    try:
                        os.remove(os.path.join(BACKUP_DIR, f))
                    except:
                        pass
            with open(DB_FILE, "rb") as src, open(backup_filename, "wb") as dst:
                dst.write(src.read())

    def load_data():
        init_db()
        check_and_replace_weekly_backup()
        
        try:
            orders = pd.read_excel(DB_FILE, sheet_name="Orders")
            purchases = pd.read_excel(DB_FILE, sheet_name="Material_Purchases")
            recipes = pd.read_excel(DB_FILE, sheet_name="Recipes")
            inventory = pd.read_excel(DB_FILE, sheet_name="Inventory")
        except Exception:
            init_db()
            orders = pd.read_excel(DB_FILE, sheet_name="Orders")
            purchases = pd.read_excel(DB_FILE, sheet_name="Material_Purchases")
            recipes = pd.read_excel(DB_FILE, sheet_name="Recipes")
            inventory = pd.read_excel(DB_FILE, sheet_name="Inventory")
        
        if "Address" not in orders.columns:
            orders["Address"] = "N/A"
        if "Delivery Charges" not in orders.columns:
            orders["Delivery Charges"] = 0.0
        if "Total Collection" not in orders.columns:
            orders["Total Collection"] = orders["Price Charged (Income)"] if "Price Charged (Income)" in orders.columns else 0.0
        if "Notes" not in orders.columns:
            orders["Notes"] = "None"
        if "Deadline Date" not in orders.columns:
            orders["Deadline Date"] = datetime.now().strftime("%Y-%m-%d")

        if "Base Qty Bought" not in purchases.columns:
            purchases["Base Qty Bought"] = purchases["Quantity Bought"] if "Quantity Bought" in purchases.columns else 0.0
        if "Base Unit" not in purchases.columns:
            purchases["Base Unit"] = purchases["Purchase Unit"] if "Purchase Unit" in purchases.columns else "grams"

        if "Base Qty Used" not in recipes.columns:
            recipes["Base Qty Used"] = recipes["Qty Used"] if "Qty Used" in recipes.columns else 0.0
        if "Consumption Unit" not in recipes.columns:
            recipes["Consumption Unit"] = "grams"

        if "Base Unit" not in inventory.columns:
            inventory["Base Unit"] = "grams"

        for idx, row in inventory.iterrows():
            ing = row["Ingredient Name"]
            tot_purchased = purchases[purchases["Ingredient Name"] == ing]["Base Qty Bought"].sum() if not purchases.empty else 0.0
            tot_utilized = recipes[recipes["Ingredient Used"] == ing]["Base Qty Used"].sum() if not recipes.empty else 0.0
            inventory.at[idx, "Total Purchased"] = tot_purchased
            inventory.at[idx, "Total Utilized"] = tot_utilized
            inventory.at[idx, "Current Stock Left"] = tot_purchased - tot_utilized

        return orders, purchases, recipes, inventory

    def save_data(orders, purchases, recipes, inventory):
        with pd.ExcelWriter(DB_FILE, engine='openpyxl') as writer:
            orders.to_excel(writer, sheet_name="Orders", index=False)
            purchases.to_excel(writer, sheet_name="Material_Purchases", index=False)
            recipes.to_excel(writer, sheet_name="Recipes", index=False)
            inventory.to_excel(writer, sheet_name="Inventory", index=False)

    def erase_entire_database():
        if os.path.exists(DB_FILE):
            try:
                os.remove(DB_FILE)
            except:
                pass
        init_db()

    # -----------------------------------------------------------------------------
    # PDF REPORT & UNIFIED INVOICE UTILITIES (Saved directly in pdf_prints/)
    # -----------------------------------------------------------------------------
    def generate_invoice_pdf(order_row, recipe_rows):
        init_directories()
        pdf_path = os.path.join(PDF_DIR, f"Invoice_{order_row['Order ID']}.pdf")
        doc = SimpleDocTemplate(pdf_path, pagesize=letter)
        styles = getSampleStyleSheet()
        story = []

        deadline_style = ParagraphStyle(
            'DeadlineTopStyle', 
            parent=styles['Heading2'], 
            fontName='Helvetica-Bold', 
            fontSize=13, 
            textColor=colors.HexColor('#C0392B'), 
            spaceAfter=10
        )
        story.append(Paragraph(f"⏰ DEADLINE / DELIVERY DUE DATE: {order_row.get('Deadline Date', 'N/A')}", deadline_style))

        title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontName='Helvetica-Bold', fontSize=18, textColor=colors.HexColor('#9B59B6'), spaceAfter=2)
        story.append(Paragraph("CAKE HUT HARIPUR", title_style))
        story.append(Paragraph("<b>Contact:</b> 0310586903", styles['Normal']))
        story.append(Paragraph("<b>OFFICIAL CUSTOMER & RIDER COLLECTION INVOICE SLIP</b>", styles['Heading3']))
        story.append(Spacer(1, 10))

        meta_data = [
            [Paragraph(f"<b>Order ID:</b> {order_row['Order ID']}", styles['Normal']), Paragraph(f"<b>Order Date:</b> {order_row['Order Date']}", styles['Normal'])],
            [Paragraph(f"<b>Customer Name:</b> {order_row['Customer Name']}", styles['Normal']), Paragraph(f"<b>Phone:</b> {order_row['Phone Number']}", styles['Normal'])],
            [Paragraph(f"<b>Delivery Address:</b> {order_row.get('Address', 'N/A')}", styles['Normal']), Paragraph(f"<b>Status:</b> {order_row['Status']}", styles['Normal'])],
            [Paragraph(f"<b>Special Notes:</b> {order_row.get('Notes', 'None')}", styles['Normal']), Paragraph("", styles['Normal'])]
        ]
        t_meta = Table(meta_data, colWidths=[260, 240])
        t_meta.setStyle(TableStyle([('BOTTOMPADDING', (0,0), (-1,-1), 4)]))
        story.append(t_meta)
        story.append(Spacer(1, 15))

        story.append(Paragraph("<b>Order Specifications</b>", styles['Heading4']))
        item_data = [
            ["Cake Flavor / Type", "Weight", "Subtotal Price"],
            [str(order_row['Cake Type']), str(order_row['Weight (Pounds)']), f"Rs. {order_row['Price Charged (Income)']:,.2f}"]
        ]
        t_item = Table(item_data, colWidths=[220, 130, 150])
        t_item.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#F0F2F5')),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey),
            ('BOTTOMPADDING', (0,0), (-1,0), 6)
        ]))
        story.append(t_item)
        story.append(Spacer(1, 15))

        collection_data = [
            [Paragraph("<b>Cake Price:</b>", styles['Normal']), Paragraph(f"Rs. {order_row['Price Charged (Income)']:,.2f}", styles['Normal'])],
            [Paragraph("<b>Delivery Charges:</b>", styles['Normal']), Paragraph(f"Rs. {order_row.get('Delivery Charges', 0.0):,.2f}", styles['Normal'])],
            [Paragraph("<b>TOTAL AMOUNT TO BE COLLECTED BY RIDER FROM CUSTOMER:</b>", styles['Heading4']), Paragraph(f"<b>Rs. {order_row.get('Total Collection', order_row['Price Charged (Income)']):,.2f}</b>", styles['Heading4'])]
        ]
        t_coll = Table(collection_data, colWidths=[280, 220])
        t_coll.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#FFF0F5')),
            ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey),
            ('BOTTOMPADDING', (0,0), (-1,-1), 6)
        ]))
        story.append(t_coll)
        story.append(Spacer(1, 20))
        story.append(Paragraph("Thank you for choosing Cake Hut Haripur! For queries call 0310586903.", styles['Normal']))

        doc.build(story)
        return pdf_path

    # -----------------------------------------------------------------------------
    # STREAMLIT UI & SINGLE-CLICK BUTTON NAVIGATION ARCHITECTURE
    # -----------------------------------------------------------------------------
    orders_df, purchases_df, recipes_df, inventory_df = load_data()

    if "nav_mode" not in st.session_state:
        st.session_state.nav_mode = "Dashboard Overview"

    st.sidebar.markdown("<h2 style='color: #9B59B6; text-align: center;'>🎂 Cake Hut ERP</h2>", unsafe_allow_html=True)
    st.sidebar.markdown("<p style='font-size: small; color: gray; text-align: center;'>Bakery Enterprise System</p>", unsafe_allow_html=True)
    st.sidebar.divider()

    st.sidebar.markdown("### 🔄 Sequential Workflow")
    if st.sidebar.button("📊 Dashboard Overview", use_container_width=True):
        st.session_state.nav_mode = "Dashboard Overview"
    if st.sidebar.button("1️⃣ Take Customer Order", use_container_width=True):
        st.session_state.nav_mode = "1. Log Customer Order"
    if st.sidebar.button("2️⃣ Log Kitchen Consumption", use_container_width=True):
        st.session_state.nav_mode = "2. Log Consumption / Recipe"
    if st.sidebar.button("3️⃣ Register Material Expense", use_container_width=True):
        st.session_state.nav_mode = "3. Log Material Expense"

    st.sidebar.markdown("### 📁 ERP Ledgers & Reports")
    if st.sidebar.button("⏳ Pending Orders Control Panel", use_container_width=True):
        st.session_state.nav_mode = "Pending Orders"
    if st.sidebar.button("📋 All Orders & Invoices", use_container_width=True):
        st.session_state.nav_mode = "Orders & Invoices"
    if st.sidebar.button("📦 Inventory Stock Ledger", use_container_width=True):
        st.session_state.nav_mode = "Inventory & Shopping"
    if st.sidebar.button("📈 Financial Analytics", use_container_width=True):
        st.session_state.nav_mode = "Analytics & Reports"
    if st.sidebar.button("⚙️ System Administration & DB", use_container_width=True):
        st.session_state.nav_mode = "System Settings"

    app_mode = st.session_state.nav_mode

    st.sidebar.divider()
    st.sidebar.markdown("### ⏰ 24-Hr Deadline Alerts")
    now = datetime.now()
    if not orders_df.empty:
        pending_check = orders_df[orders_df["Status"].astype(str).str.strip().str.lower() == "pending"]
        urgent_count = 0
        for _, r in pending_check.iterrows():
            try:
                deadline_dt = datetime.strptime(str(r.get("Deadline Date", "")), "%Y-%m-%d")
                if timedelta(0) <= (deadline_dt - now) <= timedelta(hours=24):
                    urgent_count += 1
                    st.sidebar.warning(f"⚠️ **{r['Order ID']}** ({r['Customer Name']}) due soon!")
            except:
                pass
        if urgent_count == 0:
            st.sidebar.success("All pending orders on schedule.")

    # =============================================================================
    # MODULE 1: DASHBOARD OVERVIEW
    # =============================================================================
    if app_mode == "Dashboard Overview":
        st.markdown("<h1 style='text-align: center; color: #9B59B6;'>🎂 CAKE HUT HARIPUR - ERP CONTROL CENTER 🎂</h1>", unsafe_allow_html=True)
        st.markdown("<p style='text-align: center; color: gray;'>End-to-End Bakery Operations & Inventory Management</p>", unsafe_allow_html=True)
        st.divider()

        low_stock = inventory_df[inventory_df["Current Stock Left"] <= 500.0]["Ingredient Name"].tolist()
        if low_stock:
            st.warning(f"⚠️ **Inventory Warning:** Stock is running low for: {', '.join(low_stock)}")

        total_income = orders_df["Price Charged (Income)"].sum() if not orders_df.empty else 0.0
        total_expenses = purchases_df["Total Price paid"].sum() if not purchases_df.empty else 0.0
        net_profit = total_income - total_expenses
        total_pending = len(orders_df[orders_df["Status"].astype(str).str.strip().str.lower() == "pending"]) if not orders_df.empty else 0

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("TOTAL REVENUE", f"Rs. {total_income:,.2f}")
        c2.metric("TOTAL EXPENSES", f"Rs. {total_expenses:,.2f}")
        c3.metric("NET EARNINGS", f"Rs. {net_profit:,.2f}")
        c4.metric("PENDING ORDERS", total_pending)

        st.divider()
        st.markdown("### 🚀 Disciplined ERP Workflow Quick Actions")
        col_a, col_b, col_c = st.columns(3)
        
        with col_a:
            if st.button("1️⃣ Take Customer Order", use_container_width=True, type="primary", key="btn_dash_1"):
                st.session_state.nav_mode = "1. Log Customer Order"
                st.rerun()
        with col_b:
            if st.button("2️⃣ Log Kitchen Consumption", use_container_width=True, type="primary", key="btn_dash_2"):
                st.session_state.nav_mode = "2. Log Consumption / Recipe"
                st.rerun()
        with col_c:
            if st.button("3️⃣ Register Material Expense", use_container_width=True, type="primary", key="btn_dash_3"):
                st.session_state.nav_mode = "3. Log Material Expense"
                st.rerun()

    # =============================================================================
    # MODULE 2: STEP 1 - LOG CUSTOMER ORDER
    # =============================================================================
    elif app_mode == "1. Log Customer Order":
        st.header("1️⃣ Step 1: Take & Log Customer Order")
        st.markdown("Record incoming customer specifications, delivery guidelines, and deadlines.")

        with st.form("erp_order_form"):
            c_name = st.text_input("Customer Name")
            c_phone = st.text_input("Phone Number")
            c_address = st.text_area("Delivery Address")
            c_type = st.text_input("Cake Flavor / Theme (e.g. Chocolate Truffle Fondant)")
            c_weight = st.text_input("Weight (e.g. 2 lbs)")
            c_deadline = st.date_input("Deadline Date", value=datetime.now() + timedelta(days=1))
            c_price = st.text_input("Cake Price (Income in Rs.)", value="0")
            c_delivery = st.text_input("Delivery Charges (Rs.)", value="0")
            c_notes = st.text_area("Special Instructions / Design Notes")
            c_status = st.selectbox("Order Status", ["Pending", "Completed"])
            
            if st.form_submit_button("Save Customer Order to ERP", type="primary"):
                if not c_name or not c_type:
                    st.error("Please enter Customer Name and Cake Type/Theme!")
                else:
                    try:
                        price_val = float(c_price)
                        delivery_val = float(c_delivery)
                    except ValueError:
                        st.error("Please enter valid numeric figures for Price and Delivery Charges!")
                        st.stop()

                    order_id = f"ORD-{1001 + len(orders_df)}"
                    total_collection = price_val + delivery_val
                    
                    new_order = pd.DataFrame([{
                        "Order ID": order_id,
                        "Order Date": datetime.now().strftime("%Y-%m-%d"),
                        "Deadline Date": c_deadline.strftime("%Y-%m-%d"),
                        "Customer Name": c_name,
                        "Address": c_address,
                        "Phone Number": str(c_phone),
                        "Cake Type": c_type,
                        "Weight (Pounds)": c_weight,
                        "Delivery Charges": delivery_val,
                        "Price Charged (Income)": price_val,
                        "Total Collection": total_collection,
                        "Notes": c_notes,
                        "Ingredients Cost": 0.0,
                        "Net Profit": price_val,
                        "Status": c_status
                    }])
                    orders_df = pd.concat([orders_df, new_order], ignore_index=True)
                    save_data(orders_df, purchases_df, recipes_df, inventory_df)
                    st.success(f"Order **{order_id}** saved successfully!")

    # =============================================================================
    # MODULE 3: STEP 2 - LOG KITCHEN CONSUMPTION
    # =============================================================================
    elif app_mode == "2. Log Consumption / Recipe":
        st.header("2️⃣ Step 2: Log Kitchen Consumption Against Specific Order")
        st.markdown("Select a baked order and record exact ingredient consumption to automatically deduct inventory.")

        if orders_df.empty:
            st.warning("No orders found in database. Complete Step 1 first.")
        else:
            order_options = {f"{row['Order ID']} - {row['Customer Name']} ({row['Cake Type']} - {row['Weight (Pounds)']})": row["Order ID"] for _, row in orders_df.iterrows()}
            selected_label = st.selectbox("Select Specific Order:", list(order_options.keys()))
            target_id = order_options[selected_label]
            
            match_order = orders_df[orders_df["Order ID"] == target_id].iloc[0]
            cake_desc = f"{match_order['Cake Type']} ({match_order['Weight (Pounds)']})"

            with st.form("erp_consumption_form"):
                available_ingredients = inventory_df["Ingredient Name"].tolist() if not inventory_df.empty else []
                
                if "erp_cons_rows" not in st.session_state:
                    st.session_state.erp_cons_rows = 4

                consumed_items = []
                for i in range(st.session_state.erp_cons_rows):
                    col1, col2, col3 = st.columns([2, 1, 1])
                    with col1:
                        ing_sel = st.selectbox(f"Ingredient #{i+1}", ["-- Select --"] + available_ingredients, key=f"erp_ing_{i}")
                    with col2:
                        qty_sel = st.number_input(f"Qty #{i+1}", min_value=0.0, value=0.0, step=1.0, key=f"erp_qty_{i}")
                    with col3:
                        unit_sel = st.selectbox(f"Unit #{i+1}", ["grams", "kg", "ml", "liter", "pcs", "dozen"], key=f"erp_unit_{i}")
                    
                    if ing_sel != "-- Select --" and qty_sel > 0:
                        consumed_items.append((ing_sel, qty_sel, unit_sel))

                if st.form_submit_button("➕ Add More Rows"):
                    st.session_state.erp_cons_rows += 1
                    st.rerun()

                if st.form_submit_button("Save Consumption & Deduct Inventory", type="primary"):
                    if not consumed_items:
                        st.error("Please add at least one valid ingredient item and quantity.")
                    else:
                        total_recipe_cost = 0.0
                        for ing, qty, unit in consumed_items:
                            base_qty, base_unit = convert_to_base(qty, unit)
                            
                            ing_purchases = purchases_df[purchases_df["Ingredient Name"] == ing]
                            tot_p_cost = ing_purchases["Total Price paid"].sum() if not purchases_df.empty else 0.0
                            tot_p_qty = ing_purchases["Base Qty Bought"].sum() if not purchases_df.empty else 0.0
                            unit_cost = (tot_p_cost / tot_p_qty) if tot_p_qty > 0 else 0.0
                            est_cost = base_qty * unit_cost
                            total_recipe_cost += est_cost
                            
                            new_recipe = pd.DataFrame([{
                                "Order ID Link": target_id,
                                "Cake Type & Weight": cake_desc,
                                "Ingredient Used": ing,
                                "Qty Used": qty,
                                "Consumption Unit": unit,
                                "Base Qty Used": base_qty,
                                "Est. Cost of Qty": est_cost
                            }])
                            recipes_df = pd.concat([recipes_df, new_recipe], ignore_index=True)

                        current_ing_cost = orders_df.loc[orders_df["Order ID"] == target_id, "Ingredients Cost"].values[0]
                        new_total_cost = (current_ing_cost if pd.notna(current_ing_cost) else 0.0) + total_recipe_cost
                        price_charged = orders_df.loc[orders_df["Order ID"] == target_id, "Price Charged (Income)"].values[0]

                        orders_df.loc[orders_df["Order ID"] == target_id, "Ingredients Cost"] = new_total_cost
                        orders_df.loc[orders_df["Order ID"] == target_id, "Net Profit"] = price_charged - new_total_cost

                        save_data(orders_df, purchases_df, recipes_df, inventory_df)
                        st.success(f"Kitchen consumption logged successfully for order **{target_id}**!")

    # =============================================================================
    # MODULE 4: STEP 3 - REGISTER MATERIAL EXPENSE
    # =============================================================================
    elif app_mode == "3. Log Material Expense":
        st.header("3️⃣ Step 3: Register Material Purchase & Expense")
        st.markdown("Record market purchases to restock inventory and track operating expenses.")

        with st.form("erp_purchase_form"):
            mat_name = st.text_input("Ingredient Name (e.g. Sugar, Flour, Butter)")
            qty_bought = st.number_input("Quantity Bought", min_value=0.0, step=1.0)
            p_unit = st.selectbox("Purchase Unit", ["kg", "grams", "liter", "ml", "pcs", "dozen", "packet"])
            price_paid = st.number_input("Total Price Paid (Rs.)", min_value=0.0, step=50.0)
            
            if st.form_submit_button("Register Purchase & Restock Inventory", type="primary"):
                if not mat_name or qty_bought <= 0:
                    st.error("Please enter a valid ingredient name and quantity.")
                else:
                    base_qty, base_unit = convert_to_base(qty_bought, p_unit)
                    exp_id = f"EXP-{1001 + len(purchases_df)}"
                    
                    new_exp = pd.DataFrame([{
                        "Purchase ID": exp_id,
                        "Date": datetime.now().strftime("%Y-%m-%d"),
                        "Ingredient Name": mat_name,
                        "Quantity Bought": qty_bought,
                        "Purchase Unit": p_unit,
                        "Base Qty Bought": base_qty,
                        "Base Unit": base_unit,
                        "Total Price paid": price_paid
                    }])
                    purchases_df = pd.concat([purchases_df, new_exp], ignore_index=True)
                    
                    if mat_name not in inventory_df["Ingredient Name"].values:
                        new_inv = pd.DataFrame([{
                            "Ingredient Name": mat_name,
                            "Base Unit": base_unit,
                            "Total Purchased": base_qty,
                            "Total Utilized": 0.0,
                            "Current Stock Left": base_qty
                        }])
                        inventory_df = pd.concat([inventory_df, new_inv], ignore_index=True)
                        
                    save_data(orders_df, purchases_df, recipes_df, inventory_df)
                    st.success(f"Material expense registered and inventory restocked for **{mat_name}**!")

    # =============================================================================
    # MODULE 5: DEDICATED PENDING ORDERS MANAGEMENT
    # =============================================================================
    elif app_mode == "Pending Orders":
        st.header("⏳ Pending Orders Control Panel")
        st.markdown("View active pending orders, review delivery details, and mark them as complete.")

        if orders_df.empty:
            st.info("No orders registered in the system.")
        else:
            pending_df = orders_df[orders_df["Status"].astype(str).str.strip().str.lower() == "pending"]
            
            if pending_df.empty:
                st.success("🎉 Zero pending orders! All orders are completed.")
            else:
                st.markdown(f"### Total Pending Orders: **{len(pending_df)}**")
                
                for _, row in pending_df.iterrows():
                    with st.expander(f"📦 Order ID: {row['Order ID']} | Customer: {row['Customer Name']} | Due Deadline: {row.get('Deadline Date', 'N/A')}"):
                        col_d1, col_d2 = st.columns(2)
                        with col_d1:
                            st.write(f"**Customer Name:** {row['Customer Name']}")
                            st.write(f"**Phone Number:** {row['Phone Number']}")
                            st.write(f"**Delivery Address:** {row.get('Address', 'N/A')}")
                            st.write(f"**Deadline / Due Date:** {row.get('Deadline Date', 'N/A')}")
                        with col_d2:
                            st.write(f"**Cake Type:** {row['Cake Type']}")
                            st.write(f"**Weight:** {row['Weight (Pounds)']}")
                            st.write(f"**Total Collection Required:** Rs. {row.get('Total Collection', row['Price Charged (Income)']):,.2f}")
                            st.write(f"**Notes:** {row.get('Notes', 'None')}")
                        
                        st.markdown("---")
                        col_b1, col_b2, col_b3 = st.columns(3)
                        with col_b1:
                            if st.button(f"✅ Mark as Complete", key=f"complete_pending_{row['Order ID']}", type="primary"):
                                orders_df.loc[orders_df["Order ID"] == row["Order ID"], "Status"] = "Completed"
                                save_data(orders_df, purchases_df, recipes_df, inventory_df)
                                st.success(f"Order {row['Order ID']} successfully marked as Completed!")
                                st.rerun()
                        with col_b2:
                            recs = recipes_df[recipes_df["Order ID Link"] == row["Order ID"]]
                            pdf_path = generate_invoice_pdf(row, recs)
                            with open(pdf_path, "rb") as f:
                                st.download_button(f"📥 Download Invoice", f, file_name=os.path.basename(pdf_path), mime="application/pdf", key=f"dl_pend_{row['Order ID']}")
                        with col_b3:
                            if st.button(f"🗑️ Delete Order", key=f"del_pend_{row['Order ID']}"):
                                orders_df = orders_df[orders_df["Order ID"] != row["Order ID"]]
                                recipes_df = recipes_df[recipes_df["Order ID Link"] != row["Order ID"]]
                                save_data(orders_df, purchases_df, recipes_df, inventory_df)
                                st.success(f"Order {row['Order ID']} deleted successfully!")
                                st.rerun()

    # =============================================================================
    # MODULE 6: ORDERS, INVOICES & ROW-LEVEL DELETION
    # =============================================================================
    elif app_mode == "Orders & Invoices":
        st.header("📋 All Orders Ledger, Invoices & Row Management")
        
        tab_log, tab_invoice, tab_delete = st.tabs(["All Orders Ledger", "Download Invoice & Complete", "🗑️ Delete Individual Rows"])
        
        with tab_log:
            st.dataframe(orders_df, use_container_width=True)
            
        with tab_invoice:
            col_i1, col_i2 = st.columns(2)
            with col_i1:
                st.subheader("Download Unified Invoice (Customer & Rider)")
                if not orders_df.empty:
                    sel_order = st.selectbox("Select Order ID for Invoice", orders_df["Order ID"].tolist())
                    if st.button("Generate Invoice PDF"):
                        row = orders_df[orders_df["Order ID"] == sel_order].iloc[0]
                        recs = recipes_df[recipes_df["Order ID Link"] == sel_order]
                        pdf_path = generate_invoice_pdf(row, recs)
                        with open(pdf_path, "rb") as f:
                            st.download_button("📥 Download PDF from Folder", f, file_name=os.path.basename(pdf_path), mime="application/pdf")
            with col_i2:
                st.subheader("Mark Pending Order as Completed")
                pending_list = orders_df[orders_df["Status"].astype(str).str.strip().str.lower() == "pending"]
                if pending_list.empty:
                    st.info("No pending orders to complete.")
                else:
                    p_opts = {f"{r['Order ID']} - {r['Customer Name']} ({r['Cake Type']})": r["Order ID"] for _, r in pending_list.iterrows()}
                    p_sel = st.selectbox("Select Pending Order to Fulfill", list(p_opts.keys()), key="inv_tab_sel")
                    if st.button("✅ Mark Order as Completed", type="primary", key="inv_tab_btn"):
                        target_id = p_opts[p_sel]
                        orders_df.loc[orders_df["Order ID"] == target_id, "Status"] = "Completed"
                        save_data(orders_df, purchases_df, recipes_df, inventory_df)
                        st.success(f"Order **{target_id}** marked as Completed!")
                        st.rerun()
                        
        with tab_delete:
            st.subheader("🗑️ Delete Specific Order Record")
            if orders_df.empty:
                st.info("No order records available for deletion.")
            else:
                del_opts = {f"{r['Order ID']} - {r['Customer Name']} ({r['Cake Type']})": r["Order ID"] for _, r in orders_df.iterrows()}
                del_sel = st.selectbox("Select Order Row to Delete", list(del_opts.keys()), key="del_order_sel")
                if st.button("🗑️ Permanently Delete Selected Order Row", type="primary", key="btn_del_order"):
                    target_del_id = del_opts[del_sel]
                    orders_df = orders_df[orders_df["Order ID"] != target_del_id]
                    recipes_df = recipes_df[recipes_df["Order ID Link"] != target_del_id]
                    save_data(orders_df, purchases_df, recipes_df, inventory_df)
                    st.success(f"Order row **{target_del_id}** deleted successfully!")
                    st.rerun()

    # =============================================================================
    # MODULE 7: INVENTORY STOCK LEDGER & ROW DELETION
    # =============================================================================
    elif app_mode == "Inventory & Shopping":
        st.header("📦 Inventory Stock, Consumption Ledger & Management")
        
        t1, t2, t3 = st.tabs(["Current Stock Levels", "Kitchen Recipes Ledger", "Market Purchases & Row Deletion"])
        with t1:
            st.dataframe(inventory_df, use_container_width=True)
        with t2:
            st.dataframe(recipes_df, use_container_width=True)
        with t3:
            st.dataframe(purchases_df, use_container_width=True)
            st.markdown("---")
            st.subheader("🗑️ Delete Specific Purchase Row")
            if purchases_df.empty:
                st.info("No purchase records found.")
            else:
                p_del_opts = {f"{r['Purchase ID']} - {r['Ingredient Name']} ({r['Quantity Bought']} {r['Purchase Unit']})": r["Purchase ID"] for _, r in purchases_df.iterrows()}
                p_del_sel = st.selectbox("Select Purchase Expense Row to Delete", list(p_del_opts.keys()), key="del_purch_sel")
                if st.button("🗑️ Delete Selected Purchase Expense Row", type="primary", key="btn_del_purch"):
                    target_exp_id = p_del_opts[p_del_sel]
                    purchases_df = purchases_df[purchases_df["Purchase ID"] != target_exp_id]
                    save_data(orders_df, purchases_df, recipes_df, inventory_df)
                    st.success(f"Purchase expense row **{target_exp_id}** deleted successfully!")
                    st.rerun()

    # =============================================================================
    # MODULE 8: FINANCIAL ANALYTICS
    # =============================================================================
    elif app_mode == "Analytics & Reports":
        st.header("📈 Financial & Operations Analytics")
        
        total_income = orders_df["Price Charged (Income)"].sum() if not orders_df.empty else 0.0
        total_expenses = purchases_df["Total Price paid"].sum() if not purchases_df.empty else 0.0
        net_profit = total_income - total_expenses

        c1, c2, c3 = st.columns(3)
        c1.metric("Total Revenue", f"Rs. {total_income:,.2f}")
        c2.metric("Total Purchases", f"Rs. {total_expenses:,.2f}")
        c3.metric("Net Profit", f"Rs. {net_profit:,.2f}")

        st.divider()
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Inventory Stock Balances")
            if not inventory_df.empty:
                st.bar_chart(inventory_df.set_index("Ingredient Name")["Current Stock Left"])
        with col2:
            st.subheader("Financial Comparison")
            fin_df = pd.DataFrame({"Metric": ["Revenue", "Expenses", "Net Profit"], "Amount": [total_income, total_expenses, net_profit]})
            st.bar_chart(fin_df.set_index("Metric"))

    # =============================================================================
    # MODULE 9: SYSTEM ADMINISTRATION, CREDENTIAL SETTINGS & DATABASE MANAGEMENT
    # =============================================================================
    elif app_mode == "System Settings":
        st.header("⚙️ ERP System Administration, Credentials & Database Management")
        
        st.subheader("🔐 Change Login Username & Password")
        st.markdown("Update your system credentials dynamically right from the application interface.")
        
        with st.form("change_credentials_form"):
            current_user = username
            new_username = st.text_input("New Username", value=current_user)
            new_name = st.text_input("Display Name", value=st.session_state.credentials['usernames'][current_user]['name'])
            new_email = st.text_input("Email Address", value=st.session_state.credentials['usernames'][current_user]['email'])
            new_password = st.text_input("New Password", type="password")
            confirm_password = st.text_input("Confirm New Password", type="password")
            
            if st.form_submit_button("Update Credentials", type="primary"):
                if not new_username:
                    st.error("Username cannot be empty!")
                elif new_password and new_password != confirm_password:
                    st.error("New passwords do not match!")
                else:
                    try:
                        hashed_pw = stauth.Hasher.hash(new_password) if new_password else st.session_state.credentials['usernames'][current_user]['password']
                        
                        updated_creds = {
                            'usernames': {
                                new_username: {
                                    'name': new_name,
                                    'email': new_email,
                                    'password': hashed_pw
                                }
                            }
                        }
                        st.session_state.credentials = updated_creds
                        st.success("Credentials updated successfully! Please log out and log back in with your new credentials.")
                    except Exception as e:
                        st.error(f"Error updating credentials: {e}")

        st.divider()
        st.subheader("📥 Excel Backup Management & Restoration")
        init_directories()
        backup_files = [f for f in os.listdir(BACKUP_DIR) if f.endswith(".xlsx")]
        
        col_b1, col_b2 = st.columns(2)
        
        with col_b1:
            st.markdown("#### Download Weekly Excel Backup")
            if backup_files:
                for b_file in backup_files:
                    b_path = os.path.join(BACKUP_DIR, b_file)
                    with open(b_path, "rb") as bf:
                        st.download_button(f"📥 Download {b_file}", bf, file_name=b_file, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                st.info("No weekly backup generated yet.")

        with col_b2:
            st.markdown("#### Restore System From Backup (.xlsx)")
            uploaded_backup = st.file_uploader("Upload backup Excel file to restore data", type=["xlsx"])
            if uploaded_backup is not None:
                if st.button("🔄 Overwrite & Restore System From Uploaded Backup", type="primary"):
                    with open(DB_FILE, "wb") as f:
                        f.write(uploaded_backup.getbuffer())
                    st.success("System successfully restored from backup! Refreshing app...")
                    st.rerun()

        st.divider()
        st.markdown("#### Master Database Export")
        if os.path.exists(DB_FILE):
            with open(DB_FILE, "rb") as f:
                st.download_button("📥 Download Live Master ERP Database (.xlsx)", f, file_name=f"CakeHut_LiveMaster_{datetime.now().strftime('%Y-%m-%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        st.divider()
        st.markdown("### ⚠️ Danger Zone: Erase Entire Database")
        st.error("Permanently wipe out and reset all tables (Orders, Purchases, Recipes, and Inventory Stock) back to an empty state.")
        
        if "confirm_erase" not in st.session_state:
            st.session_state.confirm_erase = False

        if not st.session_state.confirm_erase:
            if st.button("🗑️ Erase Entire Database", type="secondary"):
                st.session_state.confirm_erase = True
                st.rerun()
        else:
            st.warning("⚠️ **Are you absolutely sure?** This action cannot be undone unless you have a downloaded backup file.")
            col_e1, col_e2 = st.columns(2)
            with col_e1:
                if st.button("🚨 Yes, Permanently Erase Everything Now!", type="primary"):
                    erase_entire_database()
                    st.session_state.confirm_erase = False
                    st.success("Entire database has been completely erased and reset! Refreshing app...")
                    st.rerun()
            with col_e2:
                if st.button("❌ Cancel Erasure"):
                    st.session_state.confirm_erase = False
                    st.rerun()