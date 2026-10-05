import streamlit as st
import pandas as pd
import math
import warnings
import time
import json
import re

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
    model = genai.GenerativeModel("gemini-1.5-flash")
    conn = st.connection("gsheets", type=GSheetsConnection)
    cloud_connected = True
except Exception as e:
    cloud_connected = False
    st.error("⚠️ Cloud connection failed. Please ensure secrets.toml is set up correctly.")

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

@st.cache_data
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

data = load_and_parse_data()

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
    st.markdown("### 📸 AI Waybill & Invoice Digitizer")
    
    col_upload, col_dispatch = st.columns([1, 1], gap="large")
    
    with col_upload:
        st.info("Upload a Commercial Invoice to auto-extract customer details.")
        uploaded_file = st.file_uploader("Upload Document (JPG/PNG)", type=["jpg", "png", "jpeg"])
        
        if uploaded_file is not None and cloud_connected:
            st.image(uploaded_file, caption="Uploaded Document", use_container_width=True)
            if st.button("Extract Data with AI ✨"):
                with st.spinner("Extracting with Gemini Vision..."):
                    try:
                        img = Image.open(uploaded_file)
                        prompt = """
                        Extract the following details from this shipping invoice/document and return ONLY a valid JSON object. 
                        Keys must be exactly: 'Invoice_Date', 'Shipper_Name', 'Receiver_Name', 'Destination', 'Weight_kg', 'Cost_INR'. 
                        Format the date as YYYY-MM-DD. If a field is not found, leave it as an empty string. Do not use Markdown block syntax, just raw JSON.
                        """
                        response = model.generate_content([prompt, img])
                        
                        raw_text = response.text
                        cleaned = re.sub(r"```json", "", raw_text)
                        cleaned = re.sub(r"```", "", cleaned).strip()
                        extracted = json.loads(cleaned)
                        
                        st.session_state["extracted_data"] = extracted
                        st.success("Extraction Complete! Form pre-filled.")
                    except Exception as e:
                        st.error("Failed to extract data. The image might be blurry or the AI hit a temporary snag.")

    with col_dispatch:
        st.markdown("### 📝 Dispatch & Commit")
        
        ext = st.session_state.get("extracted_data", {})
        
        with st.form("dispatch_form"):
            st.info("💡 You can leave the AWB blank if you are only logging the Invoice right now.")
            awb = st.text_input("AWB / Tracking Number (Leave blank if pending)")
            forwarder = st.selectbox("Forwarder", ["CITI DHL", "FEDEX", "RATI DHL"])
            
            d1, d2 = st.columns(2)
            with d1:
                inv_val = ext.get("Invoice_Date", "")
                try: default_inv = pd.to_datetime(inv_val).date() if inv_val else pd.Timestamp.now().date()
                except: default_inv = pd.Timestamp.now().date()
                invoice_date = st.date_input("Invoice Date", value=default_inv)
                
            with d2:
                dispatch_date = st.date_input("Dispatch Date", value=pd.Timestamp.now().date())
            
            c1, c2 = st.columns(2)
            with c1:
                shipper = st.text_input("Shipper Name", value=ext.get("Shipper_Name", ""))
                receiver = st.text_input("Receiver Name", value=ext.get("Receiver_Name", ""))
            with c2:
                destination = st.text_input("Destination", value=ext.get("Destination", ""))
                weight = st.text_input("Weight (kg)", value=ext.get("Weight_kg", ""))
                
            c3, c4 = st.columns(2)
            with c3:
                cost = st.text_input("Billed Cost (₹)", value=ext.get("Cost_INR", ""))
            with c4:
                client_price = st.number_input("Client Charged Price (₹)", min_value=0.0, step=100.0)
                
            status = st.selectbox("Status", ["Pending AWB", "Dispatched", "In Transit", "Customs Hold", "Delivered"])
            
            submit_db = st.form_submit_button("💾 Save to Cloud Database")
            
            if submit_db:
                if cloud_connected:
                    try:
                        final_awb = awb if awb else "PENDING"
                        link = ""
                        
                        if final_awb != "PENDING":
                            if "FEDEX" in forwarder:
                                link = f"https://www.fedex.com/fedextrack/?trknbr={final_awb}"
                            else:
                                link = f"https://www.dhl.com/in-en/home/tracking/tracking-express.html?submit=1&tracking-id={final_awb}"
                        
                        new_data = pd.DataFrame([{
                            "Invoice_Date": str(invoice_date),
                            "Dispatch_Date": str(dispatch_date),
                            "AWB_Number": final_awb,
                            "Forwarder": forwarder,
                            "Shipper_Name": shipper,
                            "Receiver_Name": receiver,
                            "Destination": destination,
                            "Weight_kg": weight,
                            "Cost_INR": cost,
                            "Client_Price": client_price,
                            "Status": status,
                            "Tracking_Link": link
                        }])
                        
                        existing_df = conn.read(worksheet="Sheet1", usecols=list(range(12)))
                        if existing_df.empty or "AWB_Number" not in existing_df.columns:
                            existing_df = pd.DataFrame(columns=["Invoice_Date", "Dispatch_Date", "AWB_Number", "Forwarder", "Shipper_Name", "Receiver_Name", "Destination", "Weight_kg", "Cost_INR", "Client_Price", "Status", "Tracking_Link"])
                            
                        updated_df = pd.concat([existing_df, new_data], ignore_index=True)
                        conn.update(worksheet="Sheet1", data=updated_df)
                        
                        if final_awb == "PENDING":
                            st.success(f"✅ Invoice logged successfully! (Pending AWB generation)")
                        else:
                            st.success(f"✅ Shipment {final_awb} successfully logged!")
                            st.balloons()
                        
                        if "extracted_data" in st.session_state:
                            del st.session_state["extracted_data"]
                            
                    except Exception as e:
                        st.error(f"Failed to save to database. Error: {e}")
                else:
                    st.error("Cannot save. Database connection is inactive.")
                        
    st.markdown("---")
    st.markdown("### 📊 Live Analytics & Master Roster")
    
    if st.button("🔄 Refresh Master Database") and cloud_connected:
        st.rerun()
        
    if cloud_connected:
        try:
            db_data = conn.read(worksheet="Sheet1", usecols=list(range(12)))
            
            if not db_data.empty and len(db_data) > 0 and pd.notna(db_data.iloc[0]["AWB_Number"]):
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
            st.warning("Database headers mismatch. Ensure your Google Sheet 'Sheet1' top row exactly matches the 12 columns created in Step 1.")
