import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
import sqlite3
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# Initialize Database
def init_db():
    conn = sqlite3.connect("projects.db")
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS projects (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_name TEXT,
            client_name TEXT,
            firm_name TEXT DEFAULT 'Globe Consultants',
            tax_rate REAL DEFAULT 18.0
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS boq_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER,
            name TEXT,
            length REAL,
            breadth REAL,
            height REAL,
            unit TEXT,
            unit_rate REAL,
            wastage REAL,
            FOREIGN KEY (project_id) REFERENCES projects (id)
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS vendor_quotes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            item_id INTEGER,
            supplier_1 TEXT,
            rate_1 REAL,
            supplier_2 TEXT,
            rate_2 REAL,
            supplier_3 TEXT,
            rate_3 REAL,
            FOREIGN KEY (item_id) REFERENCES boq_items (id)
        )
    ''')
    conn.commit()
    conn.close()

init_db()

st.set_page_config(page_title="ArchBuild SaaS - Enterprise Suite", page_icon="🏗️", layout="wide")

st.title("🏗️ ArchBuild Enterprise BOQ & Procurement Suite")
st.markdown("### Multi-Project Management, Vendor Comparison, Purchase Orders & Client PDF Proposals")

# Sidebar for Project Management
st.sidebar.header("📁 Project Management")

conn = sqlite3.connect("projects.db")
cursor = conn.cursor()
cursor.execute("SELECT id, project_name FROM projects")
projects = cursor.fetchall()
project_dict = {p[1]: p[0] for p in projects}

selected_project = st.sidebar.selectbox("Switch Project", options=["-- Create New Project --"] + list(project_dict.keys()))

if selected_project == "-- Create New Project --":
    st.sidebar.markdown("---")
    st.sidebar.subheader("Create New Project")
    new_proj_name = st.sidebar.text_input("Project Name", "Commercial Complex Fitout")
    new_client_name = st.sidebar.text_input("Client Name", "Global Corp")
    new_firm_name = st.sidebar.text_input("Your Firm Name", "Globe Consultants")
    new_tax = st.sidebar.number_input("GST / Tax Rate (%)", value=18.0)
    
    if st.sidebar.button("💾 Save New Project"):
        if new_proj_name:
            cursor.execute("INSERT INTO projects (project_name, client_name, firm_name, tax_rate) VALUES (?, ?, ?, ?)", 
                           (new_proj_name, new_client_name, new_firm_name, new_tax))
            conn.commit()
            st.sidebar.success(f"Project '{new_proj_name}' created!")
            st.rerun()
        else:
            st.sidebar.error("Project name cannot be empty.")
    project_id = None
    client_name = "N/A"
    project_name = "New Project"
    firm_name = "Globe Consultants"
    tax_rate = 18.0
else:
    project_id = project_dict[selected_project]
    cursor.execute("SELECT client_name, firm_name, tax_rate FROM projects WHERE id = ?", (project_id,))
    p_data = cursor.fetchone()
    client_name = p_data[0]
    firm_name = p_data[1]
    tax_rate = p_data[2]
    project_name = selected_project
    st.sidebar.markdown(f"**Active Client:** {client_name}")
    st.sidebar.markdown(f"**Firm:** {firm_name}")

conn.close()

st.markdown("---")

if project_id:
    # Use tabs to organize BOQ vs Vendor Comparison vs PO vs PDF Quotation
    tab1, tab2, tab3, tab4 = st.tabs([
        "📋 BOQ Builder & Summary", 
        "🤝 Vendor Price Comparison", 
        "📄 Purchase Order (PO)", 
        "🖨️ Client PDF Proposal"
    ])
    
    with tab1:
        st.subheader(f"BOQ Builder for: {project_name}")
        
        # Form to add new item to database
        with st.form("add_item_form"):
            col1, col2, col3, col4, col5 = st.columns(5)
            with col1:
                item_name = st.text_input("Item Description", "Vitrified Tile Flooring")
            with col2:
                length = st.number_input("Length (m)", min_value=0.0, value=10.0, step=0.5)
            with col3:
                breadth = st.number_input("Breadth / Width (m)", min_value=0.0, value=5.0, step=0.5)
            with col4:
                height = st.number_input("Height / Depth (m)", min_value=0.0, value=0.0, step=0.1)
            with col5:
                unit_rate = st.number_input("Unit Rate (₹)", min_value=0.0, value=1200.0, step=100.0)
                
            col_a, col_b = st.columns(2)
            with col_a:
                unit = st.selectbox("Unit", ["sq.m", "cu.m", "running m", "nos"])
            with col_b:
                wastage = st.slider("Item Wastage (%)", 0, 25, 5) / 100.0
                
            submitted = st.form_submit_button("➕ Add Item to Database")
            if submitted:
                conn = sqlite3.connect("projects.db")
                cursor = conn.cursor()
                cursor.execute('''
                    INSERT INTO boq_items (project_id, name, length, breadth, height, unit, unit_rate, wastage)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                ''', (project_id, item_name, length, breadth, height, unit, unit_rate, wastage))
                item_id = cursor.lastrowid
                # Initialize default vendor quote entry
                cursor.execute('''
                    INSERT INTO vendor_quotes (item_id, supplier_1, rate_1, supplier_2, rate_2, supplier_3, rate_3)
                    VALUES (?, 'Supplier Alpha', ?, 'Supplier Beta', ?, 'Supplier Gamma', ?)
                ''', (item_id, unit_rate * 1.05, unit_rate, unit_rate * 0.95))
                conn.commit()
                conn.close()
                st.success(f"Added '{item_name}' successfully!")
                st.rerun()

        # Fetch items for active project from database
        conn = sqlite3.connect("projects.db")
        df_items = pd.read_sql_query("SELECT id, name, length, breadth, height, unit, unit_rate, wastage FROM boq_items WHERE project_id = ?", conn, params=(project_id,))
        conn.close()

        if not df_items.empty:
            boq_rows = []
            for idx, row in df_items.iterrows():
                length = row["length"]
                breadth = row["breadth"]
                height = row["height"]
                
                if row["unit"] == "cu.m" and height > 0:
                    net_qty = length * breadth * height
                elif row["unit"] == "running m":
                    net_qty = length
                else:
                    net_qty = length * breadth if height == 0 else length * breadth * height
                    
                total_qty = net_qty * (1 + row["wastage"])
                total_cost = total_qty * row["unit_rate"]
                
                boq_rows.append({
                    "DB_ID": row["id"],
                    "ID": idx + 1,
                    "Description": row["name"],
                    "Dimensions (L x B x H)": f"{length} x {breadth} x {height}" if height > 0 else f"{length} x {breadth}",
                    "Net Qty": round(net_qty, 2),
                    "Unit": row["unit"],
                    "Wastage": f"{int(row['wastage']*100)}%",
                    "Total Qty (w/ Waste)": round(total_qty, 2),
                    "Unit Rate (₹)": row["unit_rate"],
                    "Total Cost (₹)": round(total_cost, 2)
                })

            df_boq = pd.DataFrame(boq_rows)

            st.markdown("### 📊 Project Cost Summary")
            st.dataframe(df_boq.drop(columns=["DB_ID"]), use_container_width=True)
            
            subtotal = df_boq["Total Cost (₹)"].sum()
            tax_amount = subtotal * (tax_rate / 100.0)
            grand_total = subtotal + tax_amount
            
            col_m1, col_m2, col_m3 = st.columns(3)
            with col_m1:
                st.metric(label="Subtotal", value=f"₹{subtotal:,.2f}")
            with col_m2:
                st.metric(label=f"GST / Tax ({tax_rate}%)", value=f"₹{tax_amount:,.2f}")
            with col_m3:
                st.metric(label="💰 Grand Total (Incl. Tax)", value=f"₹{grand_total:,.2f}")

            # Cost Breakdown Chart
            st.markdown("### 📈 Cost Breakdown by Item")
            chart_data = df_boq.set_index("Description")["Total Cost (₹)"]
            st.bar_chart(chart_data)
            
            # Delete item option
            item_to_delete = st.selectbox("Select Item to Remove", options=["-- Select --"] + list(df_boq["Description"]))
            if item_to_delete != "-- Select --":
                if st.button("🗑️ Delete Selected Item"):
                    item_row = df_boq[df_boq["Description"] == item_to_delete].iloc[0]
                    db_id_to_del = item_row["DB_ID"]
                    conn = sqlite3.connect("projects.db")
                    cursor = conn.cursor()
                    cursor.execute("DELETE FROM boq_items WHERE id = ?", (db_id_to_del,))
                    cursor.execute("DELETE FROM vendor_quotes WHERE item_id = ?", (db_id_to_del,))
                    conn.commit()
                    conn.close()
                    st.success(f"Deleted '{item_to_delete}'")
                    st.rerun()

            # Excel Export generator using openpyxl
            def generate_excel():
                output = BytesIO()
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "BOQ Summary"
                
                ws.append([project_name.upper()])
                ws.append([f"Client: {client_name} | Firm: {firm_name}"])
                ws.append([])
                
                headers = ["ID", "Description", "Dimensions", "Net Qty", "Unit", "Wastage", "Total Qty", "Unit Rate (₹)", "Total Cost (₹)"]
                ws.append(headers)
                
                for row in boq_rows:
                    ws.append([
                        row["ID"], row["Description"], row["Dimensions (L x B x H)"], 
                        row["Net Qty"], row["Unit"], row["Wastage"], 
                        row["Total Qty (w/ Waste)"], row["Unit Rate (₹)"], row["Total Cost (₹)"]
                    ])
                    
                ws.append([])
                ws.append(["", "", "", "", "", "", "", "Subtotal:", subtotal])
                ws.append(["", "", "", "", "", "", "", f"GST ({tax_rate}%):", tax_amount])
                ws.append(["", "", "", "", "", "", "", "Grand Total:", grand_total])
                
                wb.save(output)
                output.seek(0)
                return output

            excel_data = generate_excel()
            st.download_button(
                label="📥 Download Professional Excel BOQ",
                data=excel_data,
                file_name=f"{project_name.replace(' ', '_')}_BOQ.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.info("No items added to this project yet. Use the BOQ builder above to add items.")

    with tab2:
        st.subheader("🤝 Multi-Vendor Price Comparison Engine")
        st.write("Compare rates from 3 different suppliers for each BOQ item and automatically identify the most cost-effective procurement option.")
        
        conn = sqlite3.connect("projects.db")
        df_items_vc = pd.read_sql_query("SELECT id, name FROM boq_items WHERE project_id = ?", conn, params=(project_id,))
        conn.close()
        
        if not df_items_vc.empty:
            for idx, r in df_items_vc.iterrows():
                item_id = r["id"]
                item_desc = r["name"]
                
                st.markdown(f"#### 🔹 Item: **{item_desc}**")
                
                conn = sqlite3.connect("projects.db")
                cursor = conn.cursor()
                cursor.execute("SELECT supplier_1, rate_1, supplier_2, rate_2, supplier_3, rate_3 FROM vendor_quotes WHERE item_id = ?", (item_id,))
                q_data = cursor.fetchone()
                conn.close()
                
                if q_data:
                    s1, r1, s2, r2, s3, r3 = q_data
                else:
                    s1, r1, s2, r2, s3, r3 = "Supplier A", 100.0, "Supplier B", 105.0, "Supplier C", 95.0
                
                with st.form(f"vendor_form_{item_id}"):
                    vc1, vc2, vc3 = st.columns(3)
                    with vc1:
                        sup_1 = st.text_input(f"Supplier 1 Name", s1, key=f"s1_{item_id}")
                        rate_1 = st.number_input(f"Rate 1 (₹)", min_value=0.0, value=float(r1), key=f"r1_{item_id}")
                    with vc2:
                        sup_2 = st.text_input(f"Supplier 2 Name", s2, key=f"s2_{item_id}")
                        rate_2 = st.number_input(f"Rate 2 (₹)", min_value=0.0, value=float(r2), key=f"r2_{item_id}")
                    with vc3:
                        sup_3 = st.text_input(f"Supplier 3 Name", s3, key=f"s3_{item_id}")
                        rate_3 = st.number_input(f"Rate 3 (₹)", min_value=0.0, value=float(r3), key=f"r3_{item_id}")
                        
                    save_quotes = st.form_submit_button("💾 Update Supplier Quotes")
                    if save_quotes:
                        conn = sqlite3.connect("projects.db")
                        cursor = conn.cursor()
                        cursor.execute('''
                            UPDATE vendor_quotes 
                            SET supplier_1=?, rate_1=?, supplier_2=?, rate_2=?, supplier_3=?, rate_3=?
                            WHERE item_id=?
                        ''', (sup_1, rate_1, sup_2, rate_2, sup_3, rate_3, item_id))
                        conn.commit()
                        conn.close()
                        st.success(f"Quotes updated for {item_desc}!")
                        st.rerun()
                
                # Determine lowest rate
                rates = {sup_1: rate_1, sup_2: rate_2, sup_3: rate_3}
                lowest_sup = min(rates, key=rates.get)
                lowest_rate = rates[lowest_sup]
                st.success(f"💡 **Recommended Vendor:** **{lowest_sup}** at **₹{lowest_rate:,.2f}** per unit.")
                st.markdown("---")
        else:
            st.info("Please add BOQ items in the first tab before comparing vendor rates.")

    with tab3:
        st.subheader("📄 Purchase Order (PO) Generator")
        st.write("Generate official purchase orders based on recommended lowest vendor rates.")
        
        conn = sqlite3.connect("projects.db")
        df_po_items = pd.read_sql_query('''
            SELECT b.name, b.length, b.breadth, b.height, b.unit, b.wastage,
                   v.supplier_1, v.rate_1, v.supplier_2, v.rate_2, v.supplier_3, v.rate_3
            FROM boq_items b
            LEFT JOIN vendor_quotes v ON b.id = v.item_id
            WHERE b.project_id = ?
        ''', conn, params=(project_id,))
        conn.close()
        
        if not df_po_items.empty:
            po_rows = []
            for idx, row in df_po_items.iterrows():
                length = row["length"]
                breadth = row["breadth"]
                height = row["height"]
                
                if row["unit"] == "cu.m" and height > 0:
                    net_qty = length * breadth * height
                elif row["unit"] == "running m":
                    net_qty = length
                else:
                    net_qty = length * breadth if height == 0 else length * breadth * height
                    
                total_qty = net_qty * (1 + row["wastage"])
                
                rates = {row["supplier_1"]: row["rate_1"], row["supplier_2"]: row["rate_2"], row["supplier_3"]: row["rate_3"]}
                best_sup = min(rates, key=rates.get)
                best_rate = rates[best_sup]
                total_po_cost = total_qty * best_rate
                
                po_rows.append({
                    "Item": row["name"],
                    "Quantity": round(total_qty, 2),
                    "Unit": row["unit"],
                    "Selected Vendor": best_sup,
                    "Best Unit Rate (₹)": best_rate,
                    "Total PO Cost (₹)": round(total_po_cost, 2)
                })
                
            df_po_final = pd.DataFrame(po_rows)
            st.dataframe(df_po_final, use_container_width=True)
            
            po_grand_total = df_po_final["Total PO Cost (₹)"].sum()
            st.metric(label="📝 Total Purchase Order Value", value=f"₹{po_grand_total:,.2f}")
            
            def generate_po_excel():
                output = BytesIO()
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Purchase Order"
                
                ws.append([f"PURCHASE ORDER - {project_name.upper()}"])
                ws.append([f"Client: {client_name} | Firm: {firm_name}"])
                ws.append([])
                
                headers = ["Item Description", "Quantity", "Unit", "Selected Vendor", "Best Unit Rate (₹)", "Total Cost (₹)"]
                ws.append(headers)
                
                for r in po_rows:
                    ws.append([r["Item"], r["Quantity"], r["Unit"], r["Selected Vendor"], r["Best Unit Rate (₹)"], r["Total PO Cost (₹)"]])
                    
                ws.append([])
                ws.append(["", "", "", "", "Grand Total:", po_grand_total])
                
                wb.save(output)
                output.seek(0)
                return output
                
            po_excel_data = generate_po_excel()
            st.download_button(
                label="📥 Download Purchase Order Excel",
                data=po_excel_data,
                file_name=f"{project_name.replace(' ', '_')}_Purchase_Order.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.info("No items available to generate Purchase Orders.")

    with tab4:
        st.subheader("🖨️ Professional Client PDF Proposal Generator")
        st.write("Generate a formal, presentation-ready PDF quotation including your firm's branding, itemized BOQ, tax breakdown, and standard commercial terms.")
        
        if not df_items.empty:
            def generate_pdf_proposal():
                pdf_buffer = BytesIO()
                doc = SimpleDocTemplate(pdf_buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
                story = []
                
                styles = getSampleStyleSheet()
                title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=20, textColor=colors.HexColor('#1E3A8A'), spaceAfter=4)
                subtitle_style = ParagraphStyle('SubTitleStyle', parent=styles['Normal'], fontSize=11, textColor=colors.HexColor('#4B5563'), spaceAfter=15)
                heading_style = ParagraphStyle('HeadingStyle', parent=styles['Heading2'], fontSize=13, textColor=colors.HexColor('#1F2937'), spaceBefore=10, spaceAfter=6)
                body_style = ParagraphStyle('BodyStyle', parent=styles['Normal'], fontSize=9, textColor=colors.HexColor('#374151'), spaceAfter=6)
                
                # Header Block
                story.append(Paragraph(f"<b>{firm_name.upper()}</b>", title_style))
                story.append(Paragraph(f"<b>Professional Architectural & Engineering Quotation</b><br/>Project: {project_name} | Client: {client_name}", subtitle_style))
                story.append(Spacer(1, 10))
                
                # BOQ Table for PDF
                story.append(Paragraph("<b>1. Itemized Bill of Quantities (BOQ)</b>", heading_style))
                
                table_data = [["Item Description", "Dimensions", "Qty", "Unit", "Rate (₹)", "Total (₹)"]]
                for row in boq_rows:
                    table_data.append([
                        row["Description"],
                        row["Dimensions (L x B x H)"],
                        str(row["Total Qty (w/ Waste)"]),
                        row["Unit"],
                        f"₹{row['Unit Rate (₹)']:,.2f}",
                        f"₹{row['Total Cost (₹)']:,.2f}"
                    ])
                
                table_data.append(["", "", "", "", "Subtotal:", f"₹{subtotal:,.2f}"])
                table_data.append(["", "", "", "", f"GST ({tax_rate}%):", f"₹{tax_amount:,.2f}"])
                table_data.append(["", "", "", "", "Grand Total:", f"₹{grand_total:,.2f}"])
                
                t = Table(table_data, colWidths=[160, 90, 50, 50, 80, 110])
                t.setStyle(TableStyle([
                    ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1E3A8A')),
                    ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
                    ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                    ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0,0), (-1,0), 9),
                    ('BOTTOMPADDING', (0,0), (-1,0), 6),
                    ('GRID', (0,0), (-1,-4), 0.5, colors.HexColor('#D1D5DB')),
                    ('BACKGROUND', (0,-3), (-1,-1), colors.HexColor('#F3F4F6')),
                    ('FONTNAME', (0,-3), (-1,-1), 'Helvetica-Bold'),
                    ('ALIGN', (4,-3), (4,-1), 'RIGHT'),
                ]))
                
                story.append(t)
                story.append(Spacer(1, 15))
                
                # Terms and Conditions
                story.append(Paragraph("<b>2. Commercial Terms & Conditions</b>", heading_style))
                terms_text = "1. <b>Validity:</b> This quotation is valid for 30 days from issuance.<br/>2. <b>Payment Schedule:</b> 30% advance upon sign-off, 40% upon milestone completion, 30% upon final handover.<br/>3. <b>Modifications:</b> Scope changes outside this BOQ require a formal change order.<br/>4. <b>Compliance:</b> Workmanship complies with standard engineering tolerances."
                story.append(Paragraph(terms_text, body_style))
                story.append(Spacer(1, 30))
                
                # Signatures
                sig_data = [
                    ["__________________________", "__________________________"],
                    ["Authorized Signatory", "Client Acceptance"],
                    [f"{firm_name}", f"{client_name}"]
                ]
                t_sig = Table(sig_data, colWidths=[250, 250])
                t_sig.setStyle(TableStyle([
                    ('ALIGN', (0,0), (-1,-1), 'LEFT'),
                    ('FONTSIZE', (0,0), (-1,-1), 9),
                    ('TEXTCOLOR', (0,0), (-1,-1), colors.HexColor('#4B5563')),
                    ('TOPPADDING', (0,0), (-1,-1), 2),
                ]))
                story.append(t_sig)
                
                doc.build(story)
                pdf_buffer.seek(0)
                return pdf_buffer

            pdf_data = generate_pdf_proposal()
            st.download_button(
                label="📥 Download Official Client Proposal PDF",
                data=pdf_data,
                file_name=f"{project_name.replace(' ', '_')}_Proposal.pdf",
                mime="application/pdf"
            )
        else:
            st.info("Please add BOQ items in Tab 1 before generating the client PDF proposal.")
else:
    st.info("👈 Please create a new project or select an existing one from the sidebar to get started.")
