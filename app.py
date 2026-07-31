import streamlit as st
import pandas as pd
import openpyxl
from datetime import datetime
import io
import os
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable

st.set_page_config(page_title="Screen Printing R&D Portal", layout="wide", page_icon="🎨")

EXCEL_FILE = "Screen_Printing_RND_Technical_Library_Workbook.xlsx"

# Initial Chemical/Ink Library Data
INITIAL_INK_LIBRARY = [
    {"Product Name": "Silicone Base Clear", "Supplier Code": "SIL-BASE-90", "Role": "Base Transparent", "Notes": "High elasticity silicone base"},
    {"Product Name": "Silicone White Undercoat", "Supplier Code": "SIL-WHT-10", "Role": "Base Opaque / White", "Notes": "Provides opacity & bleed barrier"},
    {"Product Name": "High Fastness Black Pigment", "Supplier Code": "PIG-BLK-05", "Role": "Pigment Colorant", "Notes": "Color shade matching"},
    {"Product Name": "High Fastness Red Pigment", "Supplier Code": "PIG-RED-02", "Role": "Pigment Colorant", "Notes": "Vibrant red shade matching"},
    {"Product Name": "Silicone Platinum Catalyst", "Supplier Code": "CAT-SIL-02", "Role": "Catalyst / Hardener", "Notes": "Mix thoroughly before printing"},
    {"Product Name": "Anti-Fading Crosslinker", "Supplier Code": "XL-MOD-01", "Role": "Crosslinker / Fixer", "Notes": "Enhances wash fastness (20+ cycles)"},
    {"Product Name": "Water-Based Elastic Clear Base", "Supplier Code": "WB-BASE-01", "Role": "Base Transparent", "Notes": "Eco-friendly soft hand base"},
    {"Product Name": "Plastisol High-Opacity White", "Supplier Code": "PL-WHT-99", "Role": "Base Opaque / White", "Notes": "Heavy coverage underbase"}
]

# Initialize Session State
if "ink_library" not in st.session_state:
    st.session_state.ink_library = pd.DataFrame(INITIAL_INK_LIBRARY)

if "formulation_df" not in st.session_state:
    st.session_state.formulation_df = pd.DataFrame([
        {"Delete": False, "Role": "Base Transparent", "Product Name": "Silicone Base Clear", "Code": "SIL-BASE-90", "Percentage (%)": 70.0, "Mixing Notes": "High elasticity base"},
        {"Delete": False, "Role": "Base Opaque / White", "Product Name": "Silicone White Undercoat", "Code": "SIL-WHT-10", "Percentage (%)": 20.0, "Mixing Notes": "Provides opacity & bleed barrier"},
        {"Delete": False, "Role": "Pigment Colorant", "Product Name": "High Fastness Black Pigment", "Code": "PIG-BLK-05", "Percentage (%)": 5.0, "Mixing Notes": "Color shade matching"},
        {"Delete": False, "Role": "Catalyst / Hardener", "Product Name": "Silicone Platinum Catalyst", "Code": "CAT-SIL-02", "Percentage (%)": 2.0, "Mixing Notes": "Mix thoroughly before printing"},
        {"Delete": False, "Role": "Crosslinker / Fixer", "Product Name": "Anti-Fading Crosslinker", "Code": "XL-MOD-01", "Percentage (%)": 3.0, "Mixing Notes": "Enhances wash fastness (20+ cycles)"}
    ])

# Helper Function: Load Data
def load_data():
    if not os.path.exists(EXCEL_FILE):
        return pd.DataFrame(), pd.DataFrame()
    try:
        xls = pd.ExcelFile(EXCEL_FILE)
        trials_df = pd.read_excel(xls, "R&D Master Trial Log", skiprows=2) if "R&D Master Trial Log" in xls.sheet_names else pd.DataFrame()
        matrix_df = pd.read_excel(xls, "Fabric Compatibility Matrix", skiprows=2) if "Fabric Compatibility Matrix" in xls.sheet_names else pd.DataFrame()
        return trials_df, matrix_df
    except Exception as e:
        st.error(f"Error loading workbook: {e}")
        return pd.DataFrame(), pd.DataFrame()

# Helper Function: Save Trial to Excel
def save_trial_to_excel(trial_data):
    try:
        if not os.path.exists(EXCEL_FILE):
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "R&D Master Trial Log"
            ws.append([])
            ws.append([])
            ws.append(["Trial ID", "Date", "Style / Reference", "Technique", "Fabric Type", "Objective", "Recipe Variation", "Wash Result", "Crocking Result", "Status"])
        else:
            wb = openpyxl.load_workbook(EXCEL_FILE)
            if "R&D Master Trial Log" in wb.sheetnames:
                ws = wb["R&D Master Trial Log"]
            else:
                ws = wb.create_sheet("R&D Master Trial Log")
                ws.append([])
                ws.append([])
                ws.append(["Trial ID", "Date", "Style / Reference", "Technique", "Fabric Type", "Objective", "Recipe Variation", "Wash Result", "Crocking Result", "Status"])
        
        ws.append(list(trial_data.values()))
        wb.save(EXCEL_FILE)
        return True
    except Exception as e:
        st.error(f"Failed to save to Excel file: {e}")
        return False

trials_df, matrix_df = load_data()

# PDF Generator Function
def generate_pdf_report(data):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle(
        'HeaderTitle', parent=styles['Heading1'], fontName='Helvetica-Bold', fontSize=15, leading=18,
        textColor=colors.HexColor('#1E293B'), alignment=1
    )
    subtitle_style = ParagraphStyle(
        'HeaderSub', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, leading=13,
        textColor=colors.HexColor('#2563EB'), alignment=1
    )
    sec_style = ParagraphStyle(
        'SecTitle', parent=styles['Heading2'], fontName='Helvetica-Bold', fontSize=11, leading=14,
        textColor=colors.HexColor('#2563EB'), spaceBefore=8, spaceAfter=4
    )
    p_style = ParagraphStyle(
        'BodyCustom', parent=styles['Normal'], fontName='Helvetica', fontSize=8.5, leading=11
    )

    elements = []
    
    elements.append(Paragraph("JK GARMENT SCREEN PRINTING R&D", title_style))
    elements.append(Paragraph("ADVANCED TECHNICAL RECIPE SPECIFICATION REPORT", subtitle_style))
    elements.append(Spacer(1, 6))
    elements.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor('#2563EB'), spaceAfter=8))
    
    # 1. Job Header
    h_data = [
        [Paragraph("<b>Recipe ID:</b>", p_style), Paragraph(str(data['recipe_id']), p_style), Paragraph("<b>Date:</b>", p_style), Paragraph(str(data['date']), p_style)],
        [Paragraph("<b>Style Name/No:</b>", p_style), Paragraph(str(data['style_name']), p_style), Paragraph("<b>Print Technique:</b>", p_style), Paragraph(str(data['print_tech']), p_style)],
        [Paragraph("<b>Developer / Exec:</b>", p_style), Paragraph(str(data['developer']), p_style), Paragraph("<b>Revision No:</b>", p_style), Paragraph(str(data['revision']), p_style)],
    ]
    t_h = Table(h_data, colWidths=[95, 170, 95, 170])
    t_h.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F8FAFC')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
        ('PADDING', (0,0), (-1,-1), 4),
    ]))
    elements.append(t_h)
    
    # 2. Fabric Spec
    elements.append(Paragraph("1. SUBSTRATE & FABRIC SPECIFICATIONS", sec_style))
    f_data = [
        [Paragraph("<b>Fabric Composition:</b>", p_style), Paragraph(str(data['fabric_comp']), p_style), Paragraph("<b>Fabric Color / CW:</b>", p_style), Paragraph(str(data['fabric_color']), p_style)],
        [Paragraph("<b>Fabric Construction:</b>", p_style), Paragraph(str(data['fabric_const']), p_style), Paragraph("<b>GSM:</b>", p_style), Paragraph(str(data['gsm']), p_style)],
        [Paragraph("<b>Dye Migration Risk:</b>", p_style), Paragraph(str(data['dye_risk']), p_style), Paragraph("<b>Undercoat Required:</b>", p_style), Paragraph(str(data['undercoat']), p_style)],
    ]
    t_f = Table(f_data, colWidths=[110, 155, 110, 155])
    t_f.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
        ('PADDING', (0,0), (-1,-1), 4),
    ]))
    elements.append(t_f)
    
    # 3. Ink Formulation
    elements.append(Paragraph(f"2. INK FORMULATION & CHEMICAL RECIPE (Target Batch: {data['batch_size']} g)", sec_style))
    ink_headers = ["Component Role", "Chemical / Ink Product Name", "Code", "Ratio (%)", "Weight (g)", "Mixing Notes"]
    i_table = [[Paragraph(f"<b>{h}</b>", ParagraphStyle('TH', parent=p_style, textColor=colors.white)) for h in ink_headers]]
    
    for row in data['formulation']:
        i_table.append([
            Paragraph(str(row.get('Role', '')), p_style),
            Paragraph(str(row.get('Product Name', '')), p_style),
            Paragraph(str(row.get('Code', '')), p_style),
            Paragraph(f"{float(row.get('Percentage (%)', 0)):.1f}%", p_style),
            Paragraph(f"{float(row.get('Weight (g)', 0)):.1f}g", p_style),
            Paragraph(str(row.get('Mixing Notes', '')), p_style),
        ])
    t_i = Table(i_table, colWidths=[90, 130, 60, 50, 55, 145])
    t_i.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1E293B')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
        ('PADDING', (0,0), (-1,-1), 4),
    ]))
    elements.append(t_i)
    
    # 4. Technical Machine Parameters
    elements.append(Paragraph("3. TECHNICAL PRINTING & MACHINE SETUP PARAMETERS", sec_style))
    m_data = [
        [Paragraph("<b>Mesh Count:</b>", p_style), Paragraph(str(data['mesh']), p_style), Paragraph("<b>Flash Cure Temp / Time:</b>", p_style), Paragraph(str(data['flash_cure']), p_style)],
        [Paragraph("<b>Squeegee Durometer:</b>", p_style), Paragraph(str(data['squeegee_duro']), p_style), Paragraph("<b>Main Curing Temp:</b>", p_style), Paragraph(str(data['main_cure']), p_style)],
        [Paragraph("<b>Squeegee Angle/Speed:</b>", p_style), Paragraph(str(data['squeegee_angle']), p_style), Paragraph("<b>Drying Belt Speed/Time:</b>", p_style), Paragraph(str(data['belt_speed']), p_style)],
        [Paragraph("<b>Off-Contact Distance:</b>", p_style), Paragraph(str(data['off_contact']), p_style), Paragraph("<b>Number of Passes:</b>", p_style), Paragraph(str(data['passes']), p_style)],
    ]
    t_m = Table(m_data, colWidths=[110, 155, 110, 155])
    t_m.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
        ('PADDING', (0,0), (-1,-1), 4),
    ]))
    elements.append(t_m)

    # 5. Sign-off
    elements.append(Paragraph("4. TECHNICAL APPROVAL & SIGN-OFF", sec_style))
    s_data = [
        [Paragraph("<b>R&D Exec:</b>", p_style), Paragraph(str(data['sig_rd']), p_style), Paragraph("<b>Quality Dept:</b>", p_style), Paragraph(str(data['sig_qa']), p_style)],
        [Paragraph("<b>Sample Dev Head:</b>", p_style), Paragraph(str(data['sig_sample']), p_style), Paragraph("<b>Production Manager:</b>", p_style), Paragraph(str(data['sig_prod']), p_style)],
    ]
    t_s = Table(s_data, colWidths=[100, 165, 100, 165])
    t_s.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F8FAFC')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
        ('PADDING', (0,0), (-1,-1), 4),
    ]))
    elements.append(t_s)

    doc.build(elements)
    buffer.seek(0)
    return buffer.getvalue()

st.title("🎨 Screen Printing R&D - Technical Portal")

menu = st.sidebar.radio(
    "Select Action", 
    [
        "🎨 Technical Recipe Builder & Report Generator", 
        "📚 Chemical & Ink Library Manager",
        "🧪 Log R&D Trial Result", 
        "📊 View Master Trial Log", 
        "🧩 Fabric Compatibility Matrix"
    ]
)

if menu == "🎨 Technical Recipe Builder & Report Generator":
    st.header("1. Job Header & Substrate Specifications")
    
    col1, col2 = st.columns(2)
    with col1:
        recipe_id = st.text_input("Recipe ID", value="RND-REC-2026-001")
        style_name = st.text_input("Style Name / No", value="ST-2026-GYMSHARK-01")
        developer = st.text_input("Developer / Exec", value="Durability Lab Exec")
        fabric_comp = st.text_input("Fabric Composition", value="95% Cotton / 5% Elastane")
        fabric_const = st.text_input("Fabric Construction", value="Single Jersey (Knitted)")
        dye_migration = st.selectbox("Dye Migration Risk", ["Low", "Medium", "High", "Critical"], index=1)
    
    with col2:
        rec_date = st.date_input("Date", value=datetime(2026, 8, 1))
        print_tech = st.text_input("Print Technique", value="Silicone Rubber")
        revision = st.text_input("Revision No", value="v2.1")
        fabric_color = st.text_input("Fabric Color / CW", value="Charcoal Dark Grey")
        gsm = st.text_input("GSM", value="220 GSM")
        undercoat = st.text_input("Undercoat Required", value="Yes (Anti-Bleed Barrier)")

    st.divider()
    st.header("2. Ink Formulation & Chemical Recipe")
    batch_size = st.number_input("Target Batch Size (Grams)", value=1000, step=100)
    
    st.caption("💡 **To Add Rows:** Click `+ Add row` at the bottom of the table. **To Delete Rows:** Check the `Delete` box and click **'🗑️ Delete Selected Rows'** below.")

    # Interactive Spreadsheet Table
    edited_df = st.data_editor(
        st.session_state.formulation_df,
        num_rows="dynamic",
        column_config={
            "Delete": st.column_config.CheckboxColumn("Delete?", default=False),
            "Role": st.column_config.SelectboxColumn(
                "Component Role",
                options=[
                    "Base Transparent",
                    "Base Opaque / White",
                    "Pigment Colorant",
                    "Catalyst / Hardener",
                    "Crosslinker / Fixer",
                    "Additive / Retarder",
                    "Thinner / Reducer",
                    "Thickeners",
                    "Other"
                ],
                required=True
            ),
            "Product Name": st.column_config.TextColumn("Product Name", required=True),
            "Code": st.column_config.TextColumn("Code"),
            "Percentage (%)": st.column_config.NumberColumn("Ratio (%)", min_value=0.0, max_value=100.0, step=0.5, format="%.1f%%"),
            "Mixing Notes": st.column_config.TextColumn("Mixing Notes")
        },
        use_container_width=True,
        key="formulation_editor"
    )

    # Save current edit state
    st.session_state.formulation_df = edited_df

    # Explicit Delete Button Action
    if st.button("🗑️ Delete Selected Rows"):
        if "Delete" in edited_df.columns:
            # Filter out checked rows
            filtered_df = edited_df[edited_df["Delete"] == False].reset_index(drop=True)
            st.session_state.formulation_df = filtered_df
            st.rerun()

    # Calculate Weights dynamically based on batch_size
    calc_df = st.session_state.formulation_df.copy()
    calc_df["Percentage (%)"] = pd.to_numeric(calc_df["Percentage (%)"], errors="coerce").fillna(0.0)
    calc_df["Weight (g)"] = (calc_df["Percentage (%)"] / 100.0) * batch_size
    
    total_pct = calc_df["Percentage (%)"].sum()
    total_weight = calc_df["Weight (g)"].sum()
    
    m_col1, m_col2 = st.columns(2)
    m_col1.metric("Total Ratio (%)", f"{total_pct:.1f}%")
    m_col2.metric("Total Weight (g)", f"{total_weight:.1f} g")

    st.divider()
    st.header("3. Technical Printing & Machine Setup Parameters")
    m_col1, m_col2 = st.columns(2)
    with m_col1:
        mesh = st.text_input("Mesh Count (T/Inch)", value="120T (305 Mesh)")
        squeegee_duro = st.text_input("Squeegee Durometer", value="70/90/70 Triple")
        squeegee_angle = st.text_input("Squeegee Angle / Speed", value="75° / Medium Speed")
        off_contact = st.text_input("Off-Contact Distance", value="2.5 mm")
    with m_col2:
        flash_cure = st.text_input("Flash Cure Temp / Time", value="110°C / 5 Seconds")
        main_cure = st.text_input("Main Curing Temp", value="100°C")
        belt_speed = st.text_input("Drying Belt Speed / Time", value="90sec")
        passes = st.text_input("Number of Passes / Strokes", value="2 Print - 1 Flash - 2 Print")

    st.divider()
    st.header("4. Sign-off Status")
    s_col1, s_col2 = st.columns(2)
    with s_col1:
        sig_rd = st.selectbox("R&D Exec Approval", ["Pending", "Approved", "Rejected"], index=1)
        sig_sample = st.selectbox("Sample Dev Head Approval", ["Pending", "Approved", "Rejected"], index=1)
    with s_col2:
        sig_qa = st.selectbox("Quality Dept Approval", ["Pending", "Approved", "Rejected"], index=1)
        sig_prod = st.selectbox("Production Manager Approval", ["Pending", "Approved", "Rejected"], index=1)

    st.divider()
    
    recipe_summary = {
        'recipe_id': recipe_id, 'date': str(rec_date), 'style_name': style_name,
        'print_tech': print_tech, 'developer': developer, 'revision': revision,
        'fabric_comp': fabric_comp, 'fabric_color': fabric_color, 'fabric_const': fabric_const,
        'gsm': gsm, 'dye_risk': dye_migration, 'undercoat': undercoat,
        'batch_size': batch_size, 'formulation': calc_df.to_dict(orient="records"),
        'mesh': mesh, 'squeegee_duro': squeegee_duro, 'squeegee_angle': squeegee_angle,
        'off_contact': off_contact, 'flash_cure': flash_cure, 'main_cure': main_cure,
        'belt_speed': belt_speed, 'passes': passes, 'sig_rd': sig_rd, 'sig_qa': sig_qa,
        'sig_sample': sig_sample, 'sig_prod': sig_prod
    }

    if st.button("Generate Technical Specification Report (PDF)"):
        pdf_bytes = generate_pdf_report(recipe_summary)
        st.success("✅ Technical Recipe Report generated successfully!")
        st.download_button(
            label="💾 Download PDF Specification Report",
            data=pdf_bytes,
            file_name=f"Recipe_Report_{recipe_id}.pdf",
            mime="application/pdf"
        )

elif menu == "📚 Chemical & Ink Library Manager":
    st.header("📚 Chemical & Advanced Ink Library")
    
    with st.expander("➕ Add New Chemical Product to Library"):
        with st.form("add_chem_form"):
            new_pname = st.text_input("Product Name")
            new_code = st.text_input("Supplier / Product Code")
            new_role = st.selectbox("Component Role", ["Base Transparent", "Base Opaque / White", "Pigment Colorant", "Catalyst / Hardener", "Crosslinker / Fixer", "Additive / Retarder", "Thinner / Reducer"])
            new_notes = st.text_area("Notes & Properties")
            
            if st.form_submit_button("Add to Library"):
                if new_pname and new_code:
                    new_item = pd.DataFrame([{"Product Name": new_pname, "Supplier Code": new_code, "Role": new_role, "Notes": new_notes}])
                    st.session_state.ink_library = pd.concat([st.session_state.ink_library, new_item], ignore_index=True)
                    st.success(f"Added '{new_pname}' to chemical library!")
                else:
                    st.warning("Please provide both Product Name and Code.")

    st.dataframe(st.session_state.ink_library, use_container_width=True)

elif menu == "🧪 Log R&D Trial Result":
    st.header("Log Experimental & Durability Trial Result")
    with st.form("log_trial_form"):
        col1, col2 = st.columns(2)
        with col1:
            trial_id = st.text_input("Trial ID", value="TR-2026-006")
            style_ref = st.text_input("Style / Reference", value="ST-2026-GYM-01")
            fabric_type = st.text_input("Fabric Type", value="95% Ctn / 5% Elastane")
            recipe_var = st.text_area("Recipe / Parameter Variation", value="Added 3% Crosslinker XL-MOD-01")
            crocking_res = st.text_input("Crocking Result", value="Grade 4.5")
        with col2:
            trial_date = st.date_input("Date", value=datetime(2026, 8, 1))
            technique = st.selectbox("Technique", ["Silicone", "Rubber Print", "High Density", "Flock Print", "Glitter Print"])
            objective = st.text_area("Trial Objective", value="Improve wash fastness past 20 cycles")
            wash_res = st.text_input("Wash Result (20 Cycles)", value="Grade 4.5 (Pass)")
            status = st.selectbox("Overall Status", ["APPROVED", "REJECTED", "PENDING"])
        
        if st.form_submit_button("Save Trial Log"):
            trial_data = {
                "Trial ID": trial_id,
                "Date": str(trial_date),
                "Style / Reference": style_ref,
                "Technique": technique,
                "Fabric Type": fabric_type,
                "Objective": objective,
                "Recipe Variation": recipe_var,
                "Wash Result": wash_res,
                "Crocking Result": crocking_res,
                "Status": status
            }
            if save_trial_to_excel(trial_data):
                st.success(f"Trial {trial_id} logged and saved to Excel successfully!")

elif menu == "📊 View Master Trial Log":
    st.header("R&D Master Experimental & Durability Trial Log")
    fresh_trials, _ = load_data()
    if not fresh_trials.empty:
        st.dataframe(fresh_trials, use_container_width=True)
    else:
        st.info("No trial records found in the Excel workbook yet.")

elif menu == "🧩 Fabric Compatibility Matrix":
    st.header("Print Technique vs. Fabric Substrate Compatibility Matrix")
    _, fresh_matrix = load_data()
    if not fresh_matrix.empty:
        st.dataframe(fresh_matrix, use_container_width=True)
    else:
        st.info("No compatibility matrix sheet found in the Excel workbook.") 