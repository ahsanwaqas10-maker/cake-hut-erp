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
# DYNAMIC CREDENTIALS & ETERNAL DEFAULT ADMIN CONFIGURATION
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

            df_recipe_templates = pd.DataFrame(columns=[
                "Recipe Name", "Ingredient Name", "Default Qty", "Unit"
            ])

            default_templates = [
                ("Standard Sponge (1 Pound)", "Flour", 250.0, "grams"),
                ("Standard Sponge (1 Pound)", "Sugar", 200.0, "grams"),
                ("Standard Sponge (1 Pound)", "Eggs", 4.0, "pcs"),
                ("Standard Sponge (1 Pound)", "Oil", 100.0, "ml"),
                ("Chocolate Sponge (1 Pound)", "Flour", 220.0, "grams"),
                ("Chocolate Sponge (1 Pound)", "Cocoa Powder", 30.0, "grams"),
                ("Chocolate Sponge (1 Pound)", "Sugar", 200.0, "grams"),
                ("Chocolate Sponge (1 Pound)", "Eggs", 4.0, "pcs"),
                ("Chocolate Sponge (1 Pound)", "Oil", 100.0, "ml"),
            ]
            for r_name, i_name, q_val, u_val in default_templates:
                df_recipe_templates = pd.concat([df_recipe_templates, pd.DataFrame([{
                    "Recipe Name": r_name,
                    "Ingredient Name": i_name,
                    "Default Qty": q_val,
                    "Unit": u_val
                }])], ignore_index=True)

            initial_ingredients = [
                ("Flour", "grams"), ("Sugar", "grams"), ("Eggs", "pcs"), ("Oil", "ml"), 
                ("Cocoa Powder", "grams"), ("Fondant", "grams"), ("Colors", "ml"), ("Cream", "ml"), ("Fox Balls", "pcs")
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
                df_recipe_templates.to_excel(writer, sheet_name="Recipe_Templates", index=False)

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
            try:
                recipe_templates = pd.read_excel(DB_FILE, sheet_name="Recipe_Templates")
            except:
                recipe_templates = pd.DataFrame(columns=["Recipe Name", "Ingredient Name", "Default Qty", "Unit"])
        except Exception:
            init_db()
            orders = pd.read_excel(DB_FILE, sheet_name="Orders")
            purchases = pd.read_excel(DB_FILE, sheet_name="Material_Purchases")
            recipes = pd.read_excel(DB_FILE, sheet_name="Recipes")
            inventory = pd.read_excel(DB_FILE, sheet_name="Inventory")
            recipe_templates = pd.DataFrame(columns=["Recipe Name", "Ingredient Name", "Default Qty", "Unit"])
        
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

        return orders, purchases, recipes, inventory, recipe_templates

    def save_data(orders, purchases, recipes, inventory, recipe_templates):
        with pd.ExcelWriter(DB_FILE, engine='openpyxl') as writer:
            orders.to_excel(writer, sheet_name="Orders", index=False)
            purchases.to_excel(writer, sheet_name="Material_Purchases", index=False)
            recipes.to_excel(writer, sheet_name="Recipes", index=False)
            inventory.to_excel(writer, sheet_name="Inventory", index=False)
            recipe_templates.to_excel(writer, sheet_name="Recipe_Templates", index=False)

    def erase_entire_database():
        if os.path.exists(DB_FILE):
            try:
                os.remove(DB_FILE)
            except:
                pass
        init_db()

    # -----------------------------------------------------------------------------
    # PDF REPORT & UNIFIED INVOICE UTILITIES
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
        story.append(Paragraph("<b>Contact:</b> 03151593937", styles['Normal']))
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
        story.append(Paragraph("Thank you for choosing Cake Hut Haripur! For queries call 03151593937.", styles['Normal']))

        doc.build(story)
        return pdf_path

    # -----------------------------------------------------------------------------
    # STREAMLIT UI & CONSISTENT NAVIGATION ARCHITECTURE
    # -----------------------------------------------------------------------------
    orders_df, purchases_df, recipes_df, inventory_df, recipe_templates_df = load_data()

    if "nav_mode" not in st.session_state:
        st.session_state.nav_mode = "Dashboard Overview"

    st.sidebar.markdown("<h2 style='color: #9B59B6; text-align: center;'>🎂 Cake Hut ERP</h2>", unsafe_allow_html=True)
    st.sidebar.markdown("<p style='font-size: small; color: gray; text-align: center;'>Bakery Enterprise System</p>", unsafe_allow_html=True)
    st.sidebar.divider()

    def sidebar_nav_button(label, target_mode):
        if st.sidebar.button(label, use_container_width=True):
            st.session_state.nav_mode = target_mode
            st.rerun()

    st.sidebar.markdown("### 🔄 Sequential Workflow")
    sidebar_nav_button("📊 Dashboard Overview", "Dashboard Overview")
    sidebar_nav_button("⏳ Pending Orders Control Panel", "Pending Orders")
    sidebar_nav_button("📦 Pending Inventory Dashboard", "Inventory & Shopping")
    sidebar_nav_button("📋 Pending Orders & Invoices", "Orders & Invoices")
    sidebar_nav_button("1️⃣ Take Customer Order", "1. Log Customer Order")
    sidebar_nav_button("2️⃣ Log Kitchen Consumption", "2. Log Consumption / Recipe")
    sidebar_nav_button("3️⃣ Register Material Expense", "3. Log Material Expense")

    st.sidebar.markdown("### 📁 ERP Ledgers & Reports")
    sidebar_nav_button("📖 Recipe Templates Manager", "Recipe Templates")
    sidebar_nav_button("📦 Inventory Stock Ledger", "Inventory & Shopping")
    sidebar_nav_button("📈 Financial Analytics", "Analytics & Reports")
    sidebar_nav_button("⚙️ System Administration & DB", "System Settings")

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
        st.markdown("### 🚀 Quick Access Shortcuts & Management")
        dash_b1, dash_b2 = st.columns(2)
        
        with dash_b1:
            if st.button("⏳ Pending Orders Panel", use_container_width=True, type="primary", key="dash_btn_pending"):
                st.session_state.nav_mode = "Pending Orders"
                st.rerun()
        with dash_b2:
            if st.button("📦 Pending Inventory", use_container_width=True, type="primary", key="dash_btn_inv"):
                st.session_state.nav_mode = "Inventory & Shopping"
                st.rerun()

        st.divider()
        st.markdown("### ⚡ Disciplined ERP Workflow Quick Actions")
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

        if st.session_state.get('last_order_saved'):
            st.success(f"✅ Success! Customer order **{st.session_state.last_order_saved}** has been successfully saved to the ERP system!")
            st.session_state.pop('last_order_saved', None)

        with st.form("erp_order_form", clear_on_submit=True):
            c_name = st.text_input("Customer Name")
            c_phone = st.text_input("Phone Number", value="03151593937")
            c_address = st.text_area("Delivery Address")
            c_type = st.text_input("Cake Flavor / Theme (e.g. Chocolate Truffle Fondant)")
            c_weight = st.text_input("Weight (e.g. 2 lbs)")
            c_deadline = st.date_input("Deadline Date", value=datetime.now() + timedelta(days=1))
            c_price = st.text_input("Cake Price (Income in Rs.)", value="0")
            c_delivery = st.text_input("Delivery Charges (Rs.)", value="0")
            c_notes = st.text_area("Special Instructions / Design Notes")
            c_status = st.selectbox("Order Status", ["Pending", "Completed"])
            
            submitted_order = st.form_submit_button("Save Customer Order to ERP", type="primary")
            if submitted_order:
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
                    save_data(orders_df, purchases_df, recipes_df, inventory_df, recipe_templates_df)
                    
                    st.session_state.last_order_saved = order_id
                    st.balloons()
                    st.rerun()

    # =============================================================================
    # MODULE 3: STEP 2 - LOG KITCHEN CONSUMPTION & WHOLE RECIPES / MATERIALS
    # =============================================================================
    elif app_mode == "2. Log Consumption / Recipe":
        st.header("2️⃣ Step 2: Log Kitchen Consumption & Apply Recipes")
        st.markdown("Choose a complete master recipe template or add specific variable/extra kitchen materials (like cream, fondant, colors) to deduct inventory automatically.")

        if st.session_state.get('last_consumption_saved'):
            st.success(f"✅ Success! Kitchen consumption successfully logged and inventory deducted for order **{st.session_state.last_consumption_saved}**.")
            st.session_state.pop('last_consumption_saved', None)

        if orders_df.empty:
            st.warning("No orders found in database. Complete Step 1 first.")
        else:
            order_options = {f"{row['Order ID']} - {row['Customer Name']} ({row['Cake Type']} - {row['Weight (Pounds)']})": row["Order ID"] for _, row in orders_df.iterrows()}
            selected_label = st.selectbox("Select Specific Order:", list(order_options.keys()))
            target_id = order_options[selected_label]
            
            match_order = orders_df[orders_df["Order ID"] == target_id].iloc[0]
            cake_desc = f"{match_order['Cake Type']} ({match_order['Weight (Pounds)']})"

            st.markdown("### 🍰 Choose Complete Master Recipe Template")
            unique_recipes = recipe_templates_df["Recipe Name"].unique().tolist() if not recipe_templates_df.empty else []
            selected_recipe = st.selectbox("Select Master Recipe Template", ["-- Select Complete Recipe Template --"] + unique_recipes)

            recipe_template_items = []
            if selected_recipe != "-- Select Complete Recipe Template --":
                sub_t = recipe_templates_df[recipe_templates_df["Recipe Name"] == selected_recipe]
                for _, tr in sub_t.iterrows():
                    recipe_template_items.append((tr["Ingredient Name"], tr["Default Qty"], tr["Unit"]))
                st.info(f"Loaded whole recipe **{selected_recipe}** comprising {len(recipe_template_items)} material items.")

            available_ingredients = inventory_df["Ingredient Name"].tolist() if not inventory_df.empty else []
            
            if "erp_cons_rows" not in st.session_state:
                st.session_state.erp_cons_rows = 3

            with st.form("erp_consumption_form"):
                st.markdown("#### Review / Customize Recipe & Kitchen Consumption Materials:")
                
                template_items_entered = []
                if recipe_template_items:
                    st.markdown("**Included Master Recipe Materials:**")
                    for idx, (t_ing, t_qty, t_unit) in enumerate(recipe_template_items):
                        col_t1, col_t2, col_t3 = st.columns([2, 1, 1])
                        with col_t1:
                            st.text(f"Material: {t_ing}")
                        with col_t2:
                            t_q_val = st.number_input(f"Qty {t_ing}", min_value=0.0, value=float(t_qty), step=1.0, key=f"t_qty_{idx}")
                        with col_t3:
                            t_u_val = st.text_input(f"Unit {t_ing}", value=str(t_unit), disabled=True, key=f"t_unit_{idx}")
                        if t_q_val > 0:
                            template_items_entered.append((t_ing, t_q_val, str(t_unit)))

                st.divider()
                st.markdown("#### Additional Kitchen Consumption / Extra Materials (Cream, Colors, Fondant, etc.):")
                variable_items_entered = []
                for i in range(st.session_state.erp_cons_rows):
                    col1, col2, col3 = st.columns([2, 1, 1])
                    with col1:
                        ing_sel = st.selectbox(f"Material Item #{i+1}", ["-- Select --"] + available_ingredients, key=f"erp_ing_{i}")
                    with col2:
                        qty_sel = st.number_input(f"Qty #{i+1}", min_value=0.0, value=0.0, step=1.0, key=f"erp_qty_{i}")
                    with col3:
                        unit_sel = st.selectbox(f"Unit #{i+1}", ["grams", "kg", "ml", "liter", "pcs", "dozen"], key=f"erp_unit_{i}")
                    
                    if ing_sel != "-- Select --" and qty_sel > 0:
                        variable_items_entered.append((ing_sel, qty_sel, unit_sel))

                if st.form_submit_button("➕ Add More Material Rows"):
                    st.session_state.erp_cons_rows += 1
                    st.rerun()

                submitted_consumption = st.form_submit_button("Save Kitchen Consumption & Deduct Inventory", type="primary")
                if submitted_consumption:
                    all_consumed = template_items_entered + variable_items_entered
                    if not all_consumed:
                        st.error("No valid materials or ingredients recorded for kitchen consumption.")
                    else:
                        total_recipe_cost = 0.0
                        for ing, qty, unit in all_consumed:
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

                        save_data(orders_df, purchases_df, recipes_df, inventory_df, recipe_templates_df)
                        
                        st.session_state.last_consumption_saved = target_id
                        st.balloons()
                        st.rerun()

    # =============================================================================
    # MODULE 4: STEP 3 - REGISTER MATERIAL EXPENSE
    # =============================================================================
    elif app_mode == "3. Log Material Expense":
        st.header("3️⃣ Step 3: Register Material Purchase & Expense")
        st.markdown("Record market purchases to restock inventory and track operating expenses.")

        if st.session_state.get('last_purchase_saved'):
            st.success(f"✅ Success! Material expense **{st.session_state.last_purchase_saved}** has been successfully registered and inventory restocked.")
            st.session_state.pop('last_purchase_saved', None)

        with st.form("erp_purchase_form", clear_on_submit=True):
            mat_name = st.text_input("Ingredient/Material Name (e.g. Sugar, Flour, Cream, Butter)")
            qty_bought = st.number_input("Quantity Bought", min_value=0.0, step=1.0)
            p_unit = st.selectbox("Purchase Unit", ["kg", "grams", "liter", "ml", "pcs", "dozen", "packet"])
            price_paid = st.number_input("Total Price Paid (Rs.)", min_value=0.0, step=50.0)
            
            submitted_purchase = st.form_submit_button("Register Purchase & Restock Inventory", type="primary")
            if submitted_purchase:
                if not mat_name or qty_bought <= 0:
                    st.error("Please enter a valid material name and quantity.")
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
                        
                    save_data(orders_df, purchases_df, recipes_df, inventory_df, recipe_templates_df)
                    
                    st.session_state.last_purchase_saved = exp_id
                    st.balloons()
                    st.rerun()

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
                                save_data(orders_df, purchases_df, recipes_df, inventory_df, recipe_templates_df)
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
                                save_data(orders_df, purchases_df, recipes_df, inventory_df, recipe_templates_df)
                                st.success(f"Order {row['Order ID']} deleted successfully!")
                                st.rerun()

    # =============================================================================
    # MODULE 6: WHOLE RECIPE TEMPLATES MANAGER (MULTI-MATERIAL PROVISION)
    # =============================================================================
    elif app_mode == "Recipe Templates":
        st.header("📖 Recipe Templates Manager")
        st.markdown("Create and manage master recipes comprising multiple materials and exact quantities.")

        if st.session_state.get('last_template_saved'):
            st.success(f"✅ Success! Complete recipe template has been successfully saved.")
            st.session_state.pop('last_template_saved', None)

        available_ingredients = inventory_df["Ingredient Name"].tolist() if not inventory_df.empty else []

        if "recipe_rows_count" not in st.session_state:
            st.session_state.recipe_rows_count = 4

        with st.form("whole_recipe_template_form"):
            master_recipe_name = st.text_input("Master Recipe Name (e.g. Premium Chocolate Cake Recipe 1lb)")
            
            st.markdown("#### Define All Materials & Quantities for this Recipe:")
            recipe_items_data = []
            for i in range(st.session_state.recipe_rows_count):
                rc1, rc2, rc3 = st.columns([2, 1, 1])
                with rc1:
                    r_ing = st.selectbox(f"Material Item #{i+1}", ["-- Select Material --"] + available_ingredients, key=f"rt_ing_{i}")
                with rc2:
                    r_qty = st.number_input(f"Quantity #{i+1}", min_value=0.0, value=0.0, step=1.0, key=f"rt_qty_{i}")
                with rc3:
                    r_unit = st.selectbox(f"Unit #{i+1}", ["grams", "ml", "pcs"], key=f"rt_unit_{i}")
                
                if r_ing != "-- Select Material --" and r_qty > 0:
                    recipe_items_data.append((r_ing, r_qty, r_unit))

            if st.form_submit_button("➕ Add More Material Rows"):
                st.session_state.recipe_rows_count += 1
                st.rerun()

            submitted_whole_recipe = st.form_submit_button("Save Complete Master Recipe Template", type="primary")
            if submitted_whole_recipe:
                if not master_recipe_name or not recipe_items_data:
                    st.error("Please enter a recipe name and at least one material item with a valid quantity.")
                else:
                    for ing_n, qty_n, unit_n in recipe_items_data:
                        new_t_row = pd.DataFrame([{
                            "Recipe Name": master_recipe_name,
                            "Ingredient Name": ing_n,
                            "Default Qty": qty_n,
                            "Unit": unit_n
                        }])
                        recipe_templates_df = pd.concat([recipe_templates_df, new_t_row], ignore_index=True)
                    
                    save_data(orders_df, purchases_df, recipes_df, inventory_df, recipe_templates_df)
                    st.session_state.last_template_saved = True
                    st.rerun()

        st.divider()
        st.subheader("Existing Master Recipe Templates")
        if not recipe_templates_df.empty:
            st.dataframe(recipe_templates_df, use_container_width=True)
        else:
            st.info("No recipe templates found.")

    # =============================================================================
    # MODULE 7: ORDERS, INVOICES & ROW-LEVEL DELETION
    # =============================================================================
    elif app_mode == "Orders & Invoices":
        st.header("📋 Pending Orders, All Invoices & Row Management")
        
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
                        save_data(orders_df, purchases_df, recipes_df, inventory_df, recipe_templates_df)
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
                    save_data(orders_df, purchases_df, recipes_df, inventory_df, recipe_templates_df)
                    st.success(f"Order row **{target_del_id}** deleted successfully!")
                    st.rerun()

    # =============================================================================
    # MODULE 8: INVENTORY STOCK LEDGER & ROW DELETION
    # =============================================================================
    elif app_mode == "Inventory & Shopping":
        st.header("📦 Pending Inventory Stock & Shopping Dashboard")
        
        t1, t2, t3 = st.tabs(["Current Stock Levels (Pending Inventory)", "Kitchen Recipes Ledger", "Market Purchases & Row Deletion"])
        with t1:
            st.subheader("Pending Inventory Stock Left")
            st.dataframe(inventory_df, use_container_width=True)
            
            inv_csv = inventory_df.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Download Pending Inventory Report (.csv)",
                data=inv_csv,
                file_name=f"Pending_Inventory_{datetime.now().strftime('%Y-%m-%d')}.csv",
                mime="text/csv",
                type="primary"
            )
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
                    save_data(orders_df, purchases_df, recipes_df, inventory_df, recipe_templates_df)
                    st.success(f"Purchase expense row **{target_exp_id}** deleted successfully!")
                    st.rerun()

    # =============================================================================
    # MODULE 9: FINANCIAL ANALYTICS
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
    # MODULE 10: SYSTEM ADMINISTRATION & PERMANENT ETERNAL CREDENTIALS
    # =============================================================================
    elif app_mode == "System Settings":
        st.header("⚙️ ERP System Administration & Database Management")
        
        st.subheader("🔐 System Credentials Info")
        st.info("The default user (**admin**) and eternal password (**12345**) remain permanently configured as requested. Internal credential modification is disabled to ensure system stability.")

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
        st.error("Permanently wipe out and reset all tables (Orders, Purchases, Recipes, Inventory Stock, and Recipe Templates) back to an empty state.")
        
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
