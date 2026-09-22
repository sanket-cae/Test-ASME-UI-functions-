import streamlit as st
import pandas as pd
import numpy as np
import datetime

# Page configuration
st.set_page_config(
    page_title="ASME Sec II Part D Property Viewer",
    page_icon="⚙️",
    layout="wide"
)

st.title("⚙️ ASME Section II, Part D Material Property Database")
st.markdown("Interactive web application replicating your Excel workbook structure for material property lookups and temperature-wise evaluations.")

# Load Excel Database (Cleaning dates and ensuring string formats)
@st.cache_data
def load_data(file_path):
    xls = pd.ExcelFile(file_path)
    db_as = pd.read_excel(file_path, sheet_name='DB_AS', dtype=str)
    db_y1 = pd.read_excel(file_path, sheet_name='DB_Y1', dtype=str)
    db_te = pd.read_excel(file_path, sheet_name='DB_TE', dtype=str)
    db_tcd = pd.read_excel(file_path, sheet_name='DB_TCD', dtype=str)
    db_tm = pd.read_excel(file_path, sheet_name='DB_TM', dtype=str)
    
    # Fix DB_MAP headers from row 0
    db_map_raw = pd.read_excel(file_path, sheet_name='DB_MAP', dtype=str)
    db_map_raw.columns = db_map_raw.iloc[0]
    db_map = db_map_raw[1:].reset_index(drop=True)
    
    # Clean any datetime conversion artifacts in Type/Grade columns
    for df in [db_as, db_y1, db_map]:
        if 'Type/Grade' in df.columns:
            df['Type/Grade'] = df['Type/Grade'].apply(lambda x: x.strftime('%b-%d').upper() if isinstance(x, (datetime.datetime, datetime.date, pd.Timestamp)) else str(x).strip())
        if 'Type/\nGrade' in df.columns:
            df['Type/\nGrade'] = df['Type/\nGrade'].apply(lambda x: x.strftime('%b-%d').upper() if isinstance(x, (datetime.datetime, datetime.date, pd.Timestamp)) else str(x).strip())

    return db_as, db_y1, db_te, db_tcd, db_tm, db_map

excel_file = 'ASME SecII Part D.xlsx'

try:
    db_as, db_y1, db_te, db_tcd, db_tm, db_map = load_data(excel_file)
except Exception as e:
    st.error(f"Error loading Excel file: {e}. Please ensure 'ASME SecII Part D.xlsx' is uploaded in your repository.")
    st.stop()

# Initialize session state for persistence
if 'evaluated' not in st.session_state:
    st.session_state.evaluated = False

# --- SIDEBAR INPUTS ---
st.sidebar.header("📝 Material & Condition Inputs")

# 1. Applicability (Code Section Dropdown mapped to exact DB_AS column names)
code_section_map = {
    'I': 'I',
    'III': 'III',
    'VIII-1': 'VIII-1',
    'VIII-2 Cl. 2': 'VIII-2\nCl. 2',
    'XII': 'XII'
}
selected_section_label = st.sidebar.selectbox("1. Applicability (Code Section)", list(code_section_map.keys()))
db_column_name = code_section_map[selected_section_label]

# 2. Spec No. selection
unique_specs = sorted([str(s).strip() for s in db_as['Spec No.'].dropna().unique() if str(s).strip() != 'nan'])
selected_spec = st.sidebar.selectbox("2. Specification Number (Spec No.)", unique_specs)

# Filter Type/Grades based on Spec No.
filtered_grades = sorted([str(g).strip() for g in db_as[db_as['Spec No.'].str.strip() == selected_spec]['Type/Grade'].dropna().unique() if str(g).strip() != 'nan'])
selected_grade = st.sidebar.selectbox("3. Type / Grade", filtered_grades)

# Filter Alloy / UNS based on Spec and Grade
filtered_uns = sorted([str(u).strip() for u in db_as[(db_as['Spec No.'].str.strip() == selected_spec) & (db_as['Type/Grade'].str.strip() == selected_grade)]['Alloy Desig./\nUNS No.'].dropna().unique() if str(u).strip() != 'nan' and str(u).strip() != '…'])
selected_uns = st.sidebar.selectbox("4. Alloy Designation / UNS No. (Optional)", ['All'] + filtered_uns)

# Filter matching rows for variants
variant_df = db_as[(db_as['Spec No.'].str.strip() == selected_spec) & (db_as['Type/Grade'].str.strip() == selected_grade)].copy()
if selected_uns != 'All':
    variant_df = variant_df[variant_df['Alloy Desig./\nUNS No.'].str.strip() == selected_uns]

# Filter variants to only include rows where the selected Code Section is valid (not 'NP')
if db_column_name in variant_df.columns:
    variant_df = variant_df[variant_df[db_column_name].str.strip().str.upper() != 'NP']

variant_df = variant_df.reset_index(drop=True)
variant_df['Variant_No'] = variant_df.index + 1

# 5. Max Allowable Stress Variants / Variant Selection
if len(variant_df) > 1:
    selected_variant = st.sidebar.selectbox(
        "5. Max. Allow. Stress Variant", 
        variant_df['Variant_No'].tolist(), 
        format_func=lambda x: f"Variant {x} (Min Tensile: {variant_df[variant_df['Variant_No']==x]['Min. Tensile\nStrength,\nMpa'].values[0]} MPa)"
    )
elif len(variant_df) == 1:
    selected_variant = 1
    st.sidebar.info("Automatically selected Variant 1 (Only 1 valid variant found for this section)")
else:
    st.warning("No permitted material records found for this Code Section applicability.")
    st.stop()

# Target Temperature for Single-Temp Evaluation
target_temp = st.sidebar.number_input("Target Temperature (°C) for Single-Point Evaluation", value=20.0, step=25.0)

# Find Properties Button
if st.sidebar.button("🔍 Find Properties", type="primary"):
    st.session_state.evaluated = True

# --- DOWNLOAD UPLOADED EXCEL FILE BUTTON ---
st.sidebar.markdown("---")
st.sidebar.subheader("📥 Download Excel Database")
try:
    with open(excel_file, "rb") as f:
        excel_bytes = f.read()
    st.sidebar.download_button(
        label="Download Uploaded Excel File",
        data=excel_bytes,
        file_name="ASME SecII Part D.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
except Exception as e:
    st.sidebar.info("Excel download button unavailable.")

# --- MAIN EXECUTION ---
if st.session_state.evaluated:
    record = variant_df[variant_df['Variant_No'] == selected_variant].iloc[0]

    # Extract Additional Material Details
    nom_comp = record.get('Nominal\nComposition', '-')
    prod_form = record.get('Product\nForm', '-')
    class_cond = record.get('Class/\nCondition/\nTemper', '-')
    size_thick = record.get('Size/\nThickness \n(mm)', '-')
    min_tensile = record.get('Min. Tensile\nStrength,\nMpa', '-')
    min_yield = record.get('Min. Yield\nStrength,\nMpa', '-')
    raw_temp_limit = record.get(db_column_name, 'NP')
    max_temp_limit = f"{raw_temp_limit} °C" if pd.notnull(raw_temp_limit) and str(raw_temp_limit).strip() not in ['NP', '-', ''] else str(raw_temp_limit)
    ext_chart = record.get('Ext. Pressure \nChart No.', '-')
    notes = record.get('Notes', '-')

    # Look up TE, TCD, TM groups, Poisson's Ratio, and Density from DB_MAP
    te_group, tcd_group, tm_group, poisson, density = 'Group 1', 'Group A', 'C<=0.3%', 0.3, 7850
    if 'Spec No.' in db_map.columns:
        matched_map = db_map[
            (db_map['Spec No.'].astype(str).str.strip() == str(selected_spec).strip()) & 
            (db_map['Type/\nGrade'].astype(str).str.strip() == str(selected_grade).strip())
        ]
        if matched_map.empty:
            matched_map = db_map[db_map['Spec No.'].astype(str).str.strip() == str(selected_spec).strip()]

        if not matched_map.empty:
            te_group = matched_map.iloc[0].get('Table TE Group', 'Group 1')
            tcd_group = matched_map.iloc[0].get('Table TCD Group', 'Group A')
            tm_group = matched_map.iloc[0].get('Table TM Group', 'C<=0.3%')
            poisson = float(matched_map.iloc[0].get("Poisson's\nRatio", 0.3)) if pd.notnull(matched_map.iloc[0].get("Poisson's\nRatio")) else 0.3
            density = float(matched_map.iloc[0].get("Density\nkg/m3", 7850)) if pd.notnull(matched_map.iloc[0].get("Density\nkg/m3")) else 7850

    # Similar Properties with Nominal Composition
    similar_compositions = []
    if 'Table TE Group' in db_map.columns and 'Nominal composition' in db_map.columns:
        filtered_map = db_map[
            (db_map['Table TE Group'].astype(str).str.strip() == str(te_group).strip()) & 
            (db_map['Table TCD Group'].astype(str).str.strip() == str(tcd_group).strip()) & 
            (db_map['Table TM Group'].astype(str).str.strip() == str(tm_group).strip())
        ]
        raw_comps = filtered_map['Nominal composition'].dropna().astype(str).str.strip()
        similar_compositions = sorted(list(raw_comps.unique()))

    # --- SECTION 1: ADDITIONAL MATERIAL DETAILS & GROUPS ---
    st.subheader("📋 Additional Material Details & Property Groups")
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Nominal Composition", str(nom_comp))
        st.metric("Product Form", str(prod_form))
        st.metric("TE Group", str(te_group))
    with col2:
        st.metric("Class / Condition", str(class_cond))
        st.metric("Size / Thickness", str(size_thick))
        st.metric("TCD Group", str(tcd_group))
    with col3:
        st.metric("Min. Tensile Strength", f"{float(min_tensile):.3f}" if pd.notnull(min_tensile) and str(min_tensile).replace('.','',1).isdigit() else str(min_tensile))
        st.metric("Min. Yield Strength", f"{float(min_yield):.3f}" if pd.notnull(min_yield) and str(min_yield).replace('.','',1).isdigit() else str(min_yield))
        st.metric("TM Group", str(tm_group))
    with col4:
        st.metric(f"Max Temp Limit ({selected_section_label})", str(max_temp_limit))
        st.metric("Ext. Pressure Chart", str(ext_chart))
        st.metric("Notes", str(notes) if pd.notnull(notes) else "-")

    if similar_compositions:
        st.info(f"**Similar Properties with Nominal Composition (matching TE, TCD, & TM Groups):** " + ", ".join(similar_compositions))

    st.markdown("---")

    # --- INTERPOLATION FUNCTION ---
    def interpolate_prop(df, temp_col, val_col, target, group_col=None, group_val=None):
        sub_df = df
        if group_col and group_val:
            sub_df = df[df[group_col].astype(str).str.strip() == str(group_val).strip()].copy()
        if sub_df.empty:
            return '-'
        
        sub_df[temp_col] = pd.to_numeric(sub_df[temp_col], errors='coerce')
        sub_df[val_col] = pd.to_numeric(sub_df[val_col], errors='coerce')
        sub_df = sub_df.dropna(subset=[temp_col, val_col]).sort_values(by=temp_col)
        
        if sub_df.empty:
            return '-'
        
        temps = sub_df[temp_col].values
        vals = sub_df[val_col].values

        if target in temps:
            idx = np.where(temps == target)[0][0]
            return vals[idx]
        elif target < temps[0] or target > temps[-1]:
            return '-'
        else:
            return np.interp(target, temps, vals)

    # --- SECTION 2: SINGLE TEMPERATURE EVALUATION PROPERTIES ---
    st.subheader(f"🎯 Evaluated Properties at Target Temperature: {target_temp} °C")

    te_a_val = interpolate_prop(db_te, 'T (˚C)', 'A', target_temp, 'TE GROUP', te_group)
    te_b_val = interpolate_prop(db_te, 'T (˚C)', 'B', target_temp, 'TE GROUP', te_group)
    tc_val = interpolate_prop(db_tcd, 'T (˚C)', 'TC', target_temp, 'TCD GR.', tcd_group)
    td_val = interpolate_prop(db_tcd, 'T (˚C)', 'TD', target_temp, 'TCD GR.', tcd_group)
    e_val = interpolate_prop(db_tm, 'T (˚C)', 'E', target_temp, 'TM GR.', tm_group)

    cp_val = '-'
    try:
        if isinstance(tc_val, (int, float)) and isinstance(td_val, (int, float)) and td_val > 0 and density:
            cp_val = (tc_val * (10**6)) / (float(density) * td_val)
    except Exception:
        cp_val = '-'

    temp_cols = [c for c in db_as.columns if str(c).strip().replace('.', '', 1).isdigit()]

    def get_row_stress_at_temp(row_data, t_cols, target):
        t_vals, s_vals = [], []
        for t in t_cols:
            val = row_data.get(t, 'NP')
            if pd.notnull(val):
                try:
                    f_val = float(val)
                    t_vals.append(float(t))
                    s_vals.append(f_val)
                except ValueError:
                    continue
        if not t_vals:
            return 'NP'
        
        sorted_indices = np.argsort(t_vals)
        t_vals = np.array(t_vals)[sorted_indices]
        s_vals = np.array(s_vals)[sorted_indices]

        if target in t_vals:
            return s_vals[np.where(t_vals == target)[0][0]]
        elif target < t_vals[0] or target > t_vals[-1]:
            return 'NP'
        else:
            return np.interp(target, t_vals, s_vals)

    allow_stress_val = get_row_stress_at_temp(record, temp_cols, target_temp)
    
    y1_row = db_y1[(db_y1['Spec No.'].astype(str).str.strip() == str(selected_spec).strip()) & (db_y1['Type/Grade'].astype(str).str.strip() == str(selected_grade).strip())]
    yield_stress_val = '-'
    if not y1_row.empty:
        yield_stress_val = get_row_stress_at_temp(y1_row.iloc[0], temp_cols, target_temp)

    # Display Property Cards (Formatted to max 3 decimal places)
    p_col1, p_col2, p_col3, p_col4 = st.columns(4)
    with p_col1:
        st.metric("Allowable Stress", f"{allow_stress_val:.3f} MPa" if isinstance(allow_stress_val, (int, float)) else str(allow_stress_val))
        st.metric("Yield Strength", f"{yield_stress_val:.3f} MPa" if isinstance(yield_stress_val, (int, float)) else str(yield_stress_val))
        st.metric("Modulus of Elasticity (E)", f"{e_val:.3f} GPa" if isinstance(e_val, (int, float)) else str(e_val))
    with p_col2:
        st.metric("Thermal Exp. (A)", f"{te_a_val:.3e}" if isinstance(te_a_val, (int, float)) else str(te_a_val))
        st.metric("Thermal Exp. (B)", f"{te_b_val:.3e}" if isinstance(te_b_val, (int, float)) else str(te_b_val))
        st.metric("Specific Heat (Cp)", f"{cp_val:.3f} J/kg·°C" if isinstance(cp_val, (int, float)) else str(cp_val))
    with p_col3:
        st.metric("Thermal Cond. (TC)", f"{tc_val:.3f} W/m·°C" if isinstance(tc_val, (int, float)) else str(tc_val))
        st.metric("Thermal Diff. (TD)", f"{td_val:.3f} 10⁻⁶ m²/s" if isinstance(td_val, (int, float)) else str(td_val))
    with p_col4:
        st.metric("Poisson's Ratio (μ)", f"{float(poisson):.3f}" if isinstance(poisson, (int, float)) else str(poisson))
        st.metric("Density (ρ)", f"{float(density):.3f} kg/m³" if isinstance(density, (int, float)) else str(density))

    st.markdown("---")

    # --- SECTION 3: FULL TEMPERATURE-WISE RESULTS TABLE ---
    st.subheader("📊 Full Temperature-Wise Results Breakdown")

    if temp_cols:
        stress_values = []
        for t in temp_cols:
            val = record.get(t, 'NP')
            try:
                stress_values.append(round(float(val), 3) if pd.notnull(val) and str(val).strip() not in ['NP', '-', ''] else val)
            except ValueError:
                stress_values.append(val)

        yield_values = []
        for t in temp_cols:
            val = y1_row.iloc[0].get(t, '-') if not y1_row.empty else '-'
            try:
                yield_values.append(round(float(val), 3) if pd.notnull(val) and str(val).strip() not in ['-', ''] else val)
            except ValueError:
                yield_values.append(val)

        te_a_vals, te_b_vals, tc_vals, td_vals, cp_vals, e_vals = [], [], [], [], [], []

        for t_str in temp_cols:
            t_num = float(t_str)
            
            a_val = interpolate_prop(db_te, 'T (˚C)', 'A', t_num, 'TE GROUP', te_group)
            b_val = interpolate_prop(db_te, 'T (˚C)', 'B', t_num, 'TE GROUP', te_group)
            te_a_vals.append(round(a_val, 6) if isinstance(a_val, (int, float)) else a_val)
            te_b_vals.append(round(b_val, 6) if isinstance(b_val, (int, float)) else b_val)

            tc_i = interpolate_prop(db_tcd, 'T (˚C)', 'TC', t_num, 'TCD GR.', tcd_group)
            td_i = interpolate_prop(db_tcd, 'T (˚C)', 'TD', t_num, 'TCD GR.', tcd_group)
            tc_vals.append(round(tc_i, 3) if isinstance(tc_i, (int, float)) else tc_i)
            td_vals.append(round(td_i, 3) if isinstance(td_i, (int, float)) else td_i)

            try:
                if isinstance(tc_i, (int, float)) and isinstance(td_i, (int, float)) and td_i > 0 and density:
                    cp_vals.append(round((tc_i * (10**6)) / (float(density) * td_i), 3))
                else:
                    cp_vals.append('-')
            except Exception:
                cp_vals.append('-')

            e_i = interpolate_prop(db_tm, 'T (˚C)', 'E', t_num, 'TM GR.', tm_group)
            e_vals.append(round(e_i, 3) if isinstance(e_i, (int, float)) else e_i)

        results_df = pd.DataFrame({
            'Temperature (°C)': temp_cols,
            'Allowable Stress (MPa)': stress_values,
            'Yield Strength (MPa)': yield_values,
            'Modulus E (GPa)': e_vals,
            'Thermal Cond. TC (W/m·°C)': tc_vals,
            'Thermal Diff. TD (10⁻⁶ m²/s)': td_vals,
            'Specific Heat Cp (J/kg·°C)': cp_vals,
            'Thermal Exp. A': te_a_vals,
            'Thermal Exp. B': te_b_vals,
        })

        st.dataframe(results_df, use_container_width=True)

        # Plot Trend (Temperature on X-axis)
        st.subheader("📈 Property vs. Temperature Trend")
        chart_metric = st.selectbox("Select Property to Plot:", ['Allowable Stress (MPa)', 'Yield Strength (MPa)', 'Modulus E (GPa)', 'Thermal Cond. TC (W/m·°C)', 'Specific Heat Cp (J/kg·°C)'])
        
        plot_df = results_df[['Temperature (°C)', chart_metric]].copy()
        plot_df['Temperature (°C)'] = pd.to_numeric(plot_df['Temperature (°C)'], errors='coerce')
        plot_df[chart_metric] = pd.to_numeric(plot_df[chart_metric], errors='coerce')
        plot_df = plot_df.dropna().sort_values('Temperature (°C)')

        if not plot_df.empty:
            st.line_chart(plot_df.set_index('Temperature (°C)')[chart_metric])
        else:
            st.info("No numeric data available to plot for the selected property across temperatures.")
    else:
        st.warning("No temperature columns detected in the database.")
else:
    st.info("👈 Please set your inputs in the sidebar and click **'Find Properties'** to view material details and property evaluations.")
        
