import streamlit as st
import pandas as pd
import math
import warnings
import time
import json
import re
import uuid

import google.generativeai as genai
from streamlit_gsheets import GSheetsConnection
from PIL import Image

warnings.filterwarnings('ignore')

st.set_page_config(page_title="Aryan Express Portal", page_icon="🚚", layout="wide")

st.markdown("""
    <style>
    [data-testid="stForm"] {
        background: var(--background-secondary, rgba(30, 41, 59, 0.4));
        border-top: 4px solid #E67E22;
        border-radius: 12px;
        padding: 2rem;
        box-shadow: 0 10px 25px -5px rgba(0,0,0,0.1);
    }
    .winner-card {
        background: linear-gradient(135deg, #27AE60 0%, #2ECC71 100%);
        color: white;
        padding: 2rem;
        border-radius: 12px;
        box-shadow: 0 10px 30px rgba(46, 204, 113, 0.4);
        text-align: center;
        margin-bottom: 1.5rem;
    }
    .winner-card h1 { color: white !important; font-size: 3rem; margin: 0; font-weight: 800; }
    .winner-card h3 { color: rgba(255,255,255,0.9) !important; margin: 0; font-size: 1.5rem; text-transform: uppercase; letter-spacing: 1px;}
    .winner-card p { font-size: 1.1rem; margin-top: 10px; opacity: 0.9; }
    
    .stButton>button {
        background: linear-gradient(135deg, #F39C12 0%, #E67E22 100%);
        color: white !important;
        font-weight: 600;
        border-radius: 8px;
        padding: 0.6rem 1.2rem;
        width: 100%;
        border: none;
    }
    .processing-animation {
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        padding: 2rem;
        animation: pulse 1.5s infinite;
        color: #E67E22;
        font-weight: 600;
    }
    @keyframes pulse { 0% { opacity: 0.6; } 50% { opacity: 1; } 100% { opacity: 0.6; } }
    </style>
""", unsafe_allow_html=True)

if "authenticated" not in st.session_state: 
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.image("https://img.icons8.com/color/96/000000/in-transit--v1.png", width=80)
        st.title("Aryan Express Secure Portal")
        pin_entered = st.text_input("4-Digit PIN", type="password")
        if st.button("Secure Login ➔"):
            if pin_entered == "2026":
                st.session_state.authenticated = True
                st.rerun()
            else: 
                st.error("❌ Incorrect PIN.")
    st.stop()

try:
    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
    model = genai.GenerativeModel("gemini-3.5-flash")
    conn = st.connection("gsheets", type=GSheetsConnection)
    cloud_connected = True
except Exception as e:
    cloud_connected = False
    st.error(f"⚠️ Cloud connection initialization failed: {e}")

@st.cache_data(ttl=43200) 
def fetch_fuel_surcharges():
    try:
        url = "https://www.thepostalconnect.com/rw/fuel-surcharge.asp"
        tables = pd.read_html(url)
        dhl_fs, fedex_fs = 0.4800, 0.5125 
        for df in tables:
            if 'DHL' in df.columns: dhl_fs = float(str(df['DHL'].iloc[0]).replace('%', '')) / 100.0
            if 'FEDEX' in df.columns: fedex_fs = float(str(df['FEDEX'].iloc[0]).replace('%', '')) / 100.0
        return {"DHL": dhl_fs, "FEDEX": fedex_fs}
    except: 
        return {"DHL": 0.4800, "FEDEX": 0.5125}

@st.cache_data(show_spinner="Parsing Master Pricing Files...")
def load_and_parse_data():
    def clean_name(series): 
        return series.astype(str).str.strip().str.title().str.replace(r' \(.*\)', '', regex=True).str.replace(r'\s+', ' ', regex=True)
    
    citi_raw = pd.read_excel('CITI DHL_2.xls', sheet_name='DHL RATES', header=None)
    citi_doc = citi_raw.iloc[6:10, [0] + list(range(2, 16))]
    citi_doc.columns = ['Weight'] + [f'Zone {i}' for i in range(1, 15)]
    citi_nondoc = citi_raw.iloc[13:60, [0] + list(range(2, 16))]
    citi_nondoc.columns = ['Weight'] + [f'Zone {i}' for i in range(1, 15)]
    citi_zones = pd.read_excel('CITI DHL_2.xls', sheet_name='DHL ZONE LIST')[['LIST', 'Zone']].dropna()
    citi_zones.columns = ['Country', 'Zone']
    citi_zones['Country'] = clean_name(citi_zones['Country'])
    
    fedex_raw = pd.read_excel('FFEDEX EXPORT revised_2.xls', sheet_name='fedex', header=None)
    fedex_doc = fedex_raw.iloc[4:9, 0:18]
    fedex_doc.columns = ['Weight', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q']
    fedex_nondoc = fedex_raw.iloc[10:70, 0:18]
    fedex_nondoc.columns = ['Weight', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q']
    fedex_zones = pd.read_excel('FFEDEX EXPORT revised_2.xls', sheet_name='fedex zone list', header=2)[['Country/Territory', 'IP']].dropna()
    fedex_zones.columns = ['Country', 'Zone']
    fedex_zones['Country'] = clean_name(fedex_zones['Country'])
    
    rati_raw = pd.read_excel('RATI DHL_2.xlsx', sheet_name='2026 pricing', header=None)
    rati_doc = rati_raw.iloc[12:17, 0:11]
    rati_doc.columns = ['Weight', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']
    rati_nondoc = rati_raw.iloc[19:70, 0:11]
    rati_nondoc.columns = ['Weight', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']
    rati_zones = pd.read_excel('RATI DHL_2.xlsx', sheet_name='Zone List')[['Countries', 'Zone']].dropna()
    rati_zones.columns = ['Country', 'Zone']
    rati_zones['Country'] = clean_name(rati_zones['Country'])
    
    master_countries = sorted(pd.concat([citi_zones['Country'], fedex_zones['Country'], rati_zones['Country']]).unique())
    
    return {'CITI': {'doc': citi_doc, 'nondoc': citi_nondoc, 'zones': citi_zones}, 
            'FEDEX': {'doc': fedex_doc, 'nondoc': fedex_nondoc, 'zones': fedex_zones}, 
            'RATI': {'doc': rati_doc, 'nondoc': rati_nondoc, 'zones': rati_zones}, 
            'master_countries': master_countries}

try:
    data = load_and_parse_data()
except Exception as e:
    st.error(f"Error loading local Excel rate files: {e}")
    st.stop()

def get_rate(vendor, is_doc, weight, country, data):
    zones_df = data[vendor]['zones']
    match = zones_df[zones_df['Country'].str.lower() == country.strip().lower()]
    if match.empty: return None
    zone = match['Zone'].values[0]
    df = data[vendor]['doc'].copy() if is_doc else data[vendor]['nondoc'].copy()
    df = df.dropna(subset=['Weight'])
    df['Weight'] = pd.to_numeric(df['Weight'], errors='coerce')
    rounded_weight = math.ceil(weight * 2) / 2
    col_name = f"Zone {int(float(zone))}" if vendor == 'CITI' else str(zone).strip().upper()
    if col_name not in df.columns: return None
    rate_row = df[df['Weight'] == rounded_weight]
    if rate_row.empty: rate_row = df[df['Weight'] >= weight].head(1)
    if rate_row.empty: return None
    try: return float(rate_row[col_name].values[0])
    except: return None

def get_compliance_rules(country, material_type, is_doc):
    docs = ["KYC of Sender", "KYC of Receiver", "Commercial/Proforma Invoice", "Detailed Packing List"]
    restricted = ["Counterfeit currency/goods", "Hazardous waste/chemicals", "Live animals/endangered species"]
    if "Liquid" in material_type:
        docs.extend(["MSDS (Material Safety Data Sheet)", "Non-DG Declaration"])
        restricted.append("Flammable liquids/corrosives without limits")
    elif "Sample" in material_type: docs.append("Declaration of 'No Commercial Value'")
    target = country.lower()
    if any(c in target for c in ["united states", "usa"]):
        docs.append("US Customs Form 5106 (if >$2,500)")
        restricted.append("FDA restricted foods/unauthorized plants")
    elif any(c in target for c in ["united kingdom", "uk"]):
        docs.append("EORI Number (commercial imports)")
        restricted.append("Uncertified meat/dairy")
    return docs, restricted

header_col1, header_col2, header_col3 = st.columns([1, 15, 2])
with header_col1: st.image("https://img.icons8.com/color/96/000000/in-transit--v1.png", width=60)
with header_col2: st.markdown("<h2 style='margin-top: 0px;'>Aryan Express Operations</h2>", unsafe_allow_html=True)
with header_col3: 
    if st.button("Lock 🔒"): 
        st.session_state.authenticated = False
        st.rerun()

st.markdown("---")
tab_route, tab_crm = st.tabs(["🚀 Routing Engine", "📊 CRM & Operations Hub"])

with tab_route:
    col_form, col_results = st.columns([4, 6], gap="large")

    with col_form:
        with st.form("dispatch_query"):
            st.markdown("### 🧮 Shipment Parameters")
            country = st.selectbox("Destination", options=data['master_countries'], index=None)
            actual_weight = st.number_input("Actual Weight (Kgs)", min_value=0.1, max_value=70.0, step=0.5, value=1.0)
            material_type = st.selectbox("Material Type", ["General Solid", "Samples", "Liquid/Powder"])
            
            doc_status = st.radio("Classification", ["Non-Document (Package)", "Document (PAK)"], horizontal=True)
            can_split = st.radio("Allow asymmetrical package splitting for cost optimization?", ["No", "Yes"], horizontal=True)
            
            st.markdown("#### 📦 Dimensions (cm)")
            d1, d2, d3 = st.columns(3)
            with d1: length = st.number_input("L", min_value=1.0, value=20.0, step=1.0)
            with d2: width = st.number_input("W", min_value=1.0, value=20.0, step=1.0)
            with d3: height = st.number_input("H", min_value=1.0, value=20.0, step=1.0)
            
            st.markdown("<br>", unsafe_allow_html=True)
            submit = st.form_submit_button("Calculate Route 🚀")

    with col_results:
        if submit:
            is_doc = doc_status == "Document (PAK)"
            volumetric_weight = (length * width * height) / 5000.0
            chargeable_weight = max(actual_weight, volumetric_weight)
            rounded_chargeable = math.ceil(chargeable_weight * 2) / 2
            
            if is_doc and rounded_chargeable > 2.5:
                st.error("❌ Document cannot exceed 2.5 Kg. Please select Non-Document (Package).")
            elif not country:
                st.error("⚠️ Please select a destination country.")
            else:
                with st.spinner(""):
                    placeholder = st.empty()
                    placeholder.markdown('<div class="processing-animation">Analyzing Slabs, Dynamic Combinations & Surcharges...</div>', unsafe_allow_html=True)
                    time.sleep(0.8)
                    placeholder.empty()

                raw_rates = {
                    "CITI DHL": get_rate('CITI', is_doc, chargeable_weight, country, data),
                    "FEDEX": get_rate('FEDEX', is_doc, chargeable_weight, country, data),
                    "RATI DHL": get_rate('RATI', is_doc, chargeable_weight, country, data)
                }
                valid_rates = {k: v for k, v in raw_rates.items() if v is not None and not math.isnan(v)}
                
                if valid_rates:
                    fs_rates = fetch_fuel_surcharges()
                    final_pricing = []
                    for vendor, base in valid_rates.items():
                        fs_pct = fs_rates["FEDEX"] if "FEDEX" in vendor else fs_rates["DHL"]
                        fs_amt = base * fs_pct
                        final_pricing.append({"Vendor": vendor, "Base": base, "FS_Pct": fs_pct, "Total": base + fs_amt})
                        
                    df_rates = pd.DataFrame(final_pricing).sort_values("Total")
                    winner = df_rates.iloc[0]
                    
                    st.markdown(f"""
                        <div class="winner-card">
                            <h3>Primary Recommendation</h3>
                            <h1>{winner['Vendor']}</h1>
                            <h2>₹ {winner['Total']:,.2f}</h2>
                            <p>Includes {winner['FS_Pct']*100:.2f}% Fuel Surcharge &nbsp;|&nbsp; ⚖️ Chargeable: {rounded_chargeable}kg</p>
                        </div>
                    """, unsafe_allow_html=True)
                    
                    if volumetric_weight > actual_weight:
                        st.warning(f"⚖️ **Volumetric Weight Applied:** {volumetric_weight:.2f}kg > Actual {actual_weight}kg.")
                    
                    st.markdown("#### 📊 Vendor Comparison")
                    st.dataframe(df_rates[['Vendor', 'Base', 'Total']].style.format({"Base": "₹{:,.2f}", "Total": "₹{:,.2f}"}), use_container_width=True, hide_index=True)
                    
                    if can_split == "Yes" and rounded_chargeable > 1.0:
                        best_split_cost, best_split_details = float('inf'), None
                        steps = int(rounded_chargeable / 0.5)
                        for i in range(1, steps):
                            w1, w2 = i * 0.5, rounded_chargeable - (i * 0.5)
                            if w1 < w2: continue 
                            w1_best_cost, w1_best_vendor = float('inf'), ""
                            w2_best_cost, w2_best_vendor = float('inf'), ""
                            for vendor in ['CITI', 'FEDEX', 'RATI']:
                                r1 = get_rate(vendor, is_doc, w1, country, data)
                                if r1 is not None and not math.isnan(r1):
                                    pct = fs_rates["FEDEX"] if vendor == "FEDEX" else fs_rates["DHL"]
                                    if r1 * (1 + pct) < w1_best_cost:
                                        w1_best_cost, w1_best_vendor = r1 * (1 + pct), vendor
                            for vendor in ['CITI', 'FEDEX', 'RATI']:
                                r2 = get_rate(vendor, is_doc, w2, country, data)
                                if r2 is not None and not math.isnan(r2):
                                    pct = fs_rates["FEDEX"] if vendor == "FEDEX" else fs_rates["DHL"]
                                    if r2 * (1 + pct) < w2_best_cost:
                                        w2_best_cost, w2_best_vendor = r2 * (1 + pct), vendor
                                        
                            total_split_cost = w1_best_cost + w2_best_cost
                            if total_split_cost < best_split_cost:
                                best_split_cost = total_split_cost
                                best_split_details = {"w1": w1, "v1": w1_best_vendor, "c1": w1_best_cost, "w2": w2, "v2": w2_best_vendor, "c2": w2_best_cost}

                        if best_split_details and best_split_cost < winner['Total']:
                            sd = best_split_details
                            st.success(f"💡 **Dynamic Split Optimization:** Splitting this shipment into **{sd['w1']}kg** (via {sd['v1']} for ₹{sd['c1']:,.2f}) and **{sd['w2']}kg** (via {sd['v2']} for ₹{sd['c2']:,.2f}) totals **₹{best_split_cost:,.2f}**, saving you an extra **₹{winner['Total'] - best_split_cost:,.2f}** over the best single-box rate!")

                    docs, restricted = get_compliance_rules(country, material_type, is_doc)
                    st.markdown(f"#### 🛡️ Compliance: {country}")
                    c1, c2 = st.columns(2)
                    with c1:
                        st.markdown("**✅ Mandatory Docs**")
                        for d in docs: st.caption(f"• {d}")
                    with c2:
                        st.markdown("**⛔ Restricted**")
                        for r in restricted: st.caption(f"• {r}")
                else:
                    st.error("Pricing unavailable for this specific weight and country combination.")
        else:
            st.info("👈 Enter your shipment details and click Calculate to view routing options.")

with tab_crm:
    st.markdown("### 📸 Intelligent AI Document & Invoice Digitizer")
    
    col_upload, col_dispatch = st.columns([1, 1], gap="large")
    
    with col_upload:
        st.info("Upload documents using industry-standard upload options below.")
        
        doc_type = st.selectbox("Document Classification", ["Commercial Invoice (Sender to Recipient)", "Courier Company Invoice / AWB"])
        upload_mode = st.selectbox("Upload Method", ["📁 File Upload", "📂 Explore / Local Storage", "🖼️ Photos Gallery", "📷 Live Camera Capture"])
        
        doc_input = None
        if upload_mode in ["📁 File Upload", "📂 Explore / Local Storage"]:
            doc_input = st.file_uploader("Select or Browse Document", type=["jpg", "png", "jpeg", "pdf"], key="doc_uploader_primary")
        elif upload_mode == "🖼️ Photos Gallery":
            doc_input = st.file_uploader("Select Image from Gallery", type=["jpg", "png", "jpeg"], key="doc_uploader_gallery")
        else:
            doc_input = st.camera_input("Snap Document via Camera")
            
        if doc_input is not None and cloud_connected:
            if hasattr(doc_input, "type") and doc_input.type != "application/pdf":
                st.image(doc_input, caption="Uploaded Document Preview", use_container_width=True)
            else:
                st.success("📄 Document loaded successfully.")
                
            if st.button("Extract Data & Check Duplicates ✨"):
                with st.spinner("Extracting with Gemini Vision..."):
                    try:
                        bytes_data = doc_input.getvalue()
                        mime_type = getattr(doc_input, "type", "image/jpeg")
                        
                        prompt = """
                        Extract details from this shipping invoice/document and return ONLY a valid JSON object. 
                        Keys must be exactly: 
                        'Invoice_No', 'Invoice_Date', 
                        'Shipper_Name', 'Shipper_Address', 'Shipper_Phone', 'Shipper_Email', 
                        'Receiver_Name', 'Destination', 'Weight_kg', 'Cost_INR', 'AWB_Number'. 
                        Format date as YYYY-MM-DD. Leave missing fields empty string. No markdown formatting.
                        """
                        response = model.generate_content([prompt, {"mime_type": mime_type, "data": bytes_data}])
                        raw_text = response.text
                        cleaned = re.sub(r"```json", "", raw_text)
                        cleaned = re.sub(r"```", "", cleaned).strip()
                        extracted = json.loads(cleaned)
                        
                        # Intelligent Duplicate Check
                        matched_id = ""
                        is_duplicate = False
                        if cloud_connected:
                            try:
                                existing_db = conn.read(spreadsheet=st.secrets["connections"]["gsheets"]["spreadsheet"], worksheet="Sheet1", ttl=0)
                                if not existing_db.empty and "Unique_ID" in existing_db.columns:
                                    inv_check = str(extracted.get("Invoice_No", "")).strip()
                                    awb_check = str(extracted.get("AWB_Number", "")).strip()
                                    
                                    for _, row in existing_db.iterrows():
                                        if inv_check and str(row.get("Invoice_No", "")).strip() == inv_check:
                                            is_duplicate, matched_id = True, row["Unique_ID"]
                                            break
                                        if awb_check and awb_check != "PENDING" and str(row.get("AWB_Number", "")).strip() == awb_check:
                                            is_duplicate, matched_id = True, row["Unique_ID"]
                                            break
                            except: pass
                            
                        st.session_state["extracted_data"] = extracted
                        st.session_state["doc_type_loaded"] = doc_type
                        
                        if is_duplicate:
                            st.session_state["target_id"] = matched_id
                            st.session_state["duplicate_warning"] = True
                        else:
                            st.session_state["target_id"] = f"AX-{uuid.uuid4().hex[:8].upper()}"
                            st.session_state["duplicate_warning"] = False
                            
                        # Clear uploader field state for next transaction
                        for k in ["doc_uploader_primary", "doc_uploader_gallery"]:
                            if k in st.session_state: del st.session_state[k]
                            
                        st.success("Extraction Complete! Ready for commit.")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Failed to extract data: {e}")

    with col_dispatch:
        st.markdown("### 📝 Operations Commit & Audit Controls")
        
        if st.session_state.get("duplicate_warning", False):
            st.warning(f"⚠️ **Duplicate Entry Detected!** This matches existing Unique ID: `{st.session_state.get('target_id')}`. Submitting will record an official amendment.")
            
        ext = st.session_state.get("extracted_data", {})
        loaded_doc_type = st.session_state.get("doc_type_loaded", "Commercial Invoice (Sender to Recipient)")
        active_unique_id = st.session_state.get("target_id", f"AX-{uuid.uuid4().hex[:8].upper()}")
        
        existing_row = {}
        is_frozen = False
        if st.session_state.get("target_id") and cloud_connected:
            try:
                db_full = conn.read(spreadsheet=st.secrets["connections"]["gsheets"]["spreadsheet"], worksheet="Sheet1", ttl=0)
                if not db_full.empty and "Unique_ID" in db_full.columns:
                    match_row = db_full[db_full["Unique_ID"] == st.session_state["target_id"]]
                    if not match_row.empty:
                        existing_row = match_row.iloc[0].to_dict()
                        if loaded_doc_type == "Commercial Invoice (Sender to Recipient)" and existing_row.get("Commercial_Invoice_Locked", "No") == "Yes":
                            is_frozen = True
            except: pass

        amendment_author = st.text_input("Logged-in Operator / Officer Name", value="Admin")
        
        with st.form("dispatch_form"):
            st.markdown(f"**Active Unique ID:** `{active_unique_id}`")
            if is_frozen:
                st.warning("🔒 Commercial Invoice locked. Updating dispatch metrics only.")
                
            invoice_no = st.text_input("Invoice Number", value=existing_row.get("Invoice_No", ext.get("Invoice_No", "")), disabled=is_frozen)
            awb = st.text_input("AWB / Tracking Number", value=existing_row.get("AWB_Number", ext.get("AWB_Number", "")))
            forwarding_no = st.text_input("Forwarding No (Carrier Tracking Ref)", value=existing_row.get("Forwarding_No", ""))
            forwarder = st.selectbox("Forwarder", ["CITI DHL", "FEDEX", "RATI DHL"], index=0)
            
            d1, d2 = st.columns(2)
            with d1:
                inv_val = existing_row.get("Invoice_Date", ext.get("Invoice_Date", ""))
                try: default_inv = pd.to_datetime(inv_val).date() if inv_val else pd.Timestamp.now().date()
                except: default_inv = pd.Timestamp.now().date()
                invoice_date = st.date_input("Invoice Date", value=default_inv)
            with d2:
                dispatch_date = st.date_input("Dispatch Date", value=pd.Timestamp.now().date())
            
            st.markdown("#### 📤 Sender Details")
            shipper = st.text_input("Sender Name", value=existing_row.get("Shipper_Name", ext.get("Shipper_Name", "")), disabled=is_frozen)
            shipper_address = st.text_area("Sender Address", value=existing_row.get("Shipper_Address", ext.get("Shipper_Address", "")), disabled=is_frozen)
            
            sc1, sc2 = st.columns(2)
            with sc1:
                shipper_phone = st.text_input("Sender Phone No", value=existing_row.get("Shipper_Phone", ext.get("Shipper_Phone", "")), disabled=is_frozen)
            with sc2:
                shipper_email = st.text_input("Sender Email / ID", value=existing_row.get("Shipper_Email", ext.get("Shipper_Email", "")), disabled=is_frozen)
            
            st.markdown("#### 📥 Recipient & Shipment Parameters")
            c1, c2 = st.columns(2)
            with c1:
                receiver = st.text_input("Receiver Name", value=existing_row.get("Receiver_Name", ext.get("Receiver_Name", "")), disabled=is_frozen)
            with c2:
                destination = st.text_input("Destination", value=existing_row.get("Destination", ext.get("Destination", "")), disabled=is_frozen)
                
            c3, c4 = st.columns(2)
            with c3:
                weight = st.text_input("Weight (kg)", value=str(existing_row.get("Weight_kg", ext.get("Weight_kg", ""))), disabled=is_frozen)
                cost = st.text_input("Courier Billed Cost (₹)", value=str(existing_row.get("Cost_INR", ext.get("Cost_INR", ""))))
            with c4:
                client_price = st.number_input("Client Charged Price (₹)", min_value=0.0, step=100.0, value=float(existing_row.get("Client_Price", 0.0)))
                
            status = st.selectbox("Status", ["Pending AWB", "Dispatched", "In Transit", "Customs Hold", "Delivered"])
            
            submit_db = st.form_submit_button("💾 Save Entry & Log Audit Trail")
            
            if submit_db:
                if cloud_connected:
                    try:
                        final_awb = awb if awb else "PENDING"
                        final_fwd_no = forwarding_no if forwarding_no else ""
                        
                        # Tracking links as per specs: CITI uses citinetwork.in for 1st level, DHL/Fedex for forwarding no, RATI uses DHL
                        if "CITI" in forwarder:
                            link = f"https://www.citinetwork.in/" if not final_fwd_no else f"https://www.dhl.com/in-en/home/tracking/tracking-express.html?submit=1&tracking-id={final_fwd_no}"
                        elif "FEDEX" in forwarder:
                            link = f"https://www.fedex.com/fedextrack/?trknbr={final_awb}" if not final_fwd_no else f"https://www.fedex.com/fedextrack/?trknbr={final_fwd_no}"
                        else:
                            link = f"https://www.dhl.com/in-en/home/tracking/tracking-express.html?submit=1&tracking-id={final_awb}"
                        
                        existing_df = conn.read(spreadsheet=st.secrets["connections"]["gsheets"]["spreadsheet"], worksheet="Sheet1", ttl=0)
                        if existing_df.empty or "Unique_ID" not in existing_df.columns:
                            existing_df = pd.DataFrame(columns=[
                                "Unique_ID", "Invoice_No", "Invoice_Date", "Dispatch_Date", "AWB_Number", "Forwarding_No", "Forwarder", 
                                "Shipper_Name", "Shipper_Address", "Shipper_Phone", "Shipper_Email",
                                "Receiver_Name", "Destination", "Weight_kg", "Cost_INR", 
                                "Client_Price", "Status", "Forwarder_Status", "Tracking_Link", "Commercial_Invoice_Locked"
                            ])
                        
                        is_existing_id = active_unique_id in existing_df["Unique_ID"].values
                        action_type = "AMENDMENT" if is_existing_id else "NEW_ENTRY"
                        
                        if is_existing_id:
                            old_row = existing_df[existing_df["Unique_ID"] == active_unique_id].iloc[0].to_dict()
                            existing_df = existing_df[existing_df["Unique_ID"] != active_unique_id]
                        else:
                            old_row = {}

                        # Pack changes into a single cell json string for the AuditLogs sheet
                        detailed_change_payload = json.dumps({
                            "Previous": old_row,
                            "Updated": {
                                "Invoice_No": invoice_no, "AWB_Number": final_awb, "Forwarding_No": final_fwd_no,
                                "Shipper_Name": shipper, "Receiver_Name": receiver, "Destination": destination, 
                                "Cost_INR": cost, "Status": status
                            }
                        })

                        audit_row = pd.DataFrame([{
                            "Unique_ID": active_unique_id,
                            "Timestamp": str(pd.Timestamp.now()),
                            "Operator": amendment_author,
                            "Action_Type": action_type,
                            "Detailed_Changes_JSON": detailed_change_payload
                        }])
                        
                        try:
                            existing_audit = conn.read(spreadsheet=st.secrets["connections"]["gsheets"]["spreadsheet"], worksheet="AuditLogs", ttl=0)
                            updated_audit = pd.concat([existing_audit, audit_row], ignore_index=True)
                            conn.update(spreadsheet=st.secrets["connections"]["gsheets"]["spreadsheet"], worksheet="AuditLogs", data=updated_audit)
                        except:
                            conn.update(spreadsheet=st.secrets["connections"]["gsheets"]["spreadsheet"], worksheet="AuditLogs", data=audit_row)

                        lock_status = "Yes" if loaded_doc_type == "Commercial Invoice (Sender to Recipient)" else existing_row.get("Commercial_Invoice_Locked", "No")
                        
                        new_record = pd.DataFrame([{
                            "Unique_ID": active_unique_id,
                            "Invoice_No": invoice_no,
                            "Invoice_Date": str(invoice_date),
                            "Dispatch_Date": str(dispatch_date),
                            "AWB_Number": final_awb,
                            "Forwarding_No": final_fwd_no,
                            "Forwarder": forwarder,
                            "Shipper_Name": shipper,
                            "Shipper_Address": shipper_address,
                            "Shipper_Phone": shipper_phone,
                            "Shipper_Email": shipper_email,
                            "Receiver_Name": receiver,
                            "Destination": destination,
                            "Weight_kg": weight,
                            "Cost_INR": cost,
                            "Client_Price": client_price,
                            "Status": status,
                            "Forwarder_Status": "Pending Automated Sync",
                            "Tracking_Link": link if final_awb != "PENDING" else "",
                            "Commercial_Invoice_Locked": lock_status
                        }])
                        
                        updated_df = pd.concat([existing_df, new_record], ignore_index=True)
                        conn.update(spreadsheet=st.secrets["connections"]["gsheets"]["spreadsheet"], worksheet="Sheet1", data=updated_df)
                        
                        st.success(f"✅ Record [{active_unique_id}] saved and audited successfully!")
                        st.balloons()
                        
                        for k in ["extracted_data", "target_id", "duplicate_warning"]:
                            if k in st.session_state: del st.session_state[k]
                            
                    except Exception as e:
                        st.error(f"Failed to save record: {e}")
                else:
                    st.error("Database connection inactive.")
                        
    st.markdown("---")
    st.markdown("### 📊 Live Analytics & Master Roster")
    
    if st.button("🔄 Refresh Master Database") and cloud_connected:
        st.rerun()
        
    if cloud_connected:
        try:
            db_data = conn.read(spreadsheet=st.secrets["connections"]["gsheets"]["spreadsheet"], worksheet="Sheet1", ttl=0)
            
            if not db_data.empty and len(db_data) > 0 and pd.notna(db_data.iloc[0]["Unique_ID"]):
                f1, f2 = st.columns(2)
                with f1:
                    sel_status = st.multiselect("Filter by Status", db_data["Status"].dropna().unique())
                with f2:
                    sel_fwd = st.multiselect("Filter by Forwarder", db_data["Forwarder"].dropna().unique())
                    
                filtered = db_data.copy()
                if sel_status: filtered = filtered[filtered["Status"].isin(sel_status)]
                if sel_fwd: filtered = filtered[filtered["Forwarder"].isin(sel_fwd)]
                
                def make_clickable(val):
                    if pd.isna(val) or not str(val).startswith("http"): return val
                    return f'<a target="_blank" href="{val}" style="color:#E67E22; font-weight:bold;">Track ↗</a>'
                    
                html_table = filtered.to_html(escape=False, formatters={'Tracking_Link': make_clickable}, index=False)
                st.markdown(html_table, unsafe_allow_html=True)
                
                st.markdown("<br>", unsafe_allow_html=True)
                csv = filtered.to_csv(index=False).encode('utf-8')
                st.download_button("📥 Export Roster to CSV", csv, "aryan_express_operations.csv", "text/csv")
            else:
                st.info("The database is currently empty. Add a shipment above to see analytics.")
        except Exception as e:
            st.error(f"⚠ Connection Diagnostic Error [{type(e).__name__}]: {repr(e)}")
