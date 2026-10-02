import streamlit as st
import pandas as pd
import math
import warnings

warnings.filterwarnings('ignore')

# --- UI CONFIGURATION ---
st.set_page_config(page_title="Aryan Express Courier", page_icon="🚚", layout="wide")

# Inject Custom CSS
st.markdown("""
    <style>
    .stButton>button {
        background-color: #E67E22; 
        color: white !important;
        font-weight: 600;
        border-radius: 6px;
        padding: 0.6rem 1.2rem;
        width: 100%;
        border: none;
        box-shadow: 0 4px 6px rgba(0,0,0,0.2);
        transition: all 0.3s ease;
    }
    .stButton>button:hover { background-color: #D35400; box-shadow: 0 6px 8px rgba(0,0,0,0.3); }
    [data-testid="stForm"] { border-top: 4px solid #E67E22; border-radius: 8px; box-shadow: 0 4px 15px rgba(0,0,0,0.1); }
    </style>
""", unsafe_allow_html=True)

# --- SECURITY / LOGIN SCREEN ---
# This checks if the user is logged in. If not, it stops the app and shows the PIN screen.
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.image("https://img.icons8.com/color/96/000000/in-transit--v1.png", width=80)
        st.title("Aryan Express Secure Portal")
        st.markdown("Please enter your employee PIN to access the routing engine.")
        
        # YOU CAN CHANGE YOUR PASSWORD HERE (Currently set to "2026")
        pin_entered = st.text_input("4-Digit PIN", type="password")
        
        if st.button("Login"):
            if pin_entered == "2026":
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("❌ Incorrect PIN. Access Denied.")
    st.stop() # This entirely hides the rest of the app until the PIN is correct

# --- APP DATA & ENGINE (Only runs if authenticated) ---
@st.cache_data
def load_and_parse_data():
    citi_raw = pd.read_excel('CITI DHL_2.xls', sheet_name='DHL RATES', header=None)
    citi_doc = citi_raw.iloc[6:10, [0] + list(range(2, 16))]
    citi_doc.columns = ['Weight'] + [f'Zone {i}' for i in range(1, 15)]
    citi_nondoc = citi_raw.iloc[13:60, [0] + list(range(2, 16))]
    citi_nondoc.columns = ['Weight'] + [f'Zone {i}' for i in range(1, 15)]
    citi_zones = pd.read_excel('CITI DHL_2.xls', sheet_name='DHL ZONE LIST')[['LIST', 'Zone']].dropna()
    citi_zones.columns = ['Country', 'Zone']
    
    fedex_raw = pd.read_excel('FFEDEX EXPORT revised_2.xls', sheet_name='fedex', header=None)
    fedex_doc = fedex_raw.iloc[4:9, 0:18]
    fedex_doc.columns = ['Weight', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q']
    fedex_nondoc = fedex_raw.iloc[10:70, 0:18]
    fedex_nondoc.columns = ['Weight', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q']
    fedex_zones = pd.read_excel('FFEDEX EXPORT revised_2.xls', sheet_name='fedex zone list', header=2)[['Country/Territory', 'IP']].dropna()
    fedex_zones.columns = ['Country', 'Zone']
    
    rati_raw = pd.read_excel('RATI DHL_2.xlsx', sheet_name='2026 pricing', header=None)
    rati_doc = rati_raw.iloc[12:17, 0:11]
    rati_doc.columns = ['Weight', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']
    rati_nondoc = rati_raw.iloc[19:70, 0:11]
    rati_nondoc.columns = ['Weight', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']
    rati_zones = pd.read_excel('RATI DHL_2.xlsx', sheet_name='Zone List')[['Countries', 'Zone']].dropna()
    rati_zones.columns = ['Country', 'Zone']
    
    def clean_name(series): return series.astype(str).str.strip().str.title().str.replace(r' \(.*\)', '', regex=True)
    c1, c2, c3 = clean_name(citi_zones['Country']), clean_name(fedex_zones['Country']), clean_name(rati_zones['Country'])
    master_countries = sorted(pd.concat([c1, c2, c3]).unique())
    
    return {
        'CITI': {'doc': citi_doc, 'nondoc': citi_nondoc, 'zones': citi_zones},
        'FEDEX': {'doc': fedex_doc, 'nondoc': fedex_nondoc, 'zones': fedex_zones},
        'RATI': {'doc': rati_doc, 'nondoc': rati_nondoc, 'zones': rati_zones},
        'master_countries': master_countries
    }

def get_rate(vendor, is_doc, weight, country, data):
    zones_df = data[vendor]['zones']
    target = country.strip().lower()
    match = zones_df[zones_df['Country'].astype(str).str.strip().str.lower() == target]
    if match.empty:
        match = zones_df[zones_df['Country'].astype(str).str.strip().str.lower().str.startswith(target)]
    if match.empty: return None
    
    zone = match['Zone'].values[0]
    df = data[vendor]['doc'].copy() if is_doc else data[vendor]['nondoc'].copy()
    df = df.dropna(subset=['Weight'])
    df['Weight'] = pd.to_numeric(df['Weight'], errors='coerce')
    
    rounded_weight = math.ceil(weight * 2) / 2
    col_name = f"Zone {int(zone)}" if vendor == 'CITI' else str(zone).strip()
    
    if col_name not in df.columns: return None
    
    rate_row = df[df['Weight'] == rounded_weight]
    if rate_row.empty: 
        rate_row = df[df['Weight'] >= weight].head(1)
        if rate_row.empty: return None
    try: return float(rate_row[col_name].values[0])
    except: return None

data = load_and_parse_data()

# --- HEADER SECTION ---
header_col1, header_col2 = st.columns([1, 20])
with header_col1:
    st.image("https://img.icons8.com/color/96/000000/in-transit--v1.png", width=70)
with header_col2:
    st.markdown("<h1 style='margin-top: -15px;'>Aryan Express Courier Operations & Rate Routing</h1>", unsafe_allow_html=True)
    
# Logout Button
if st.button("Lock System 🔒", type="secondary"):
    st.session_state.authenticated = False
    st.rerun()

st.markdown("---")

# --- UI QUESTIONNAIRE ---
with st.form("dispatch_query"):
    col1, col2 = st.columns(2)
    with col1:
        country = st.selectbox("Destination Country", options=data['master_countries'], index=None, placeholder="Type to search (e.g., Bosnia And Herzegovina)...")
        weight = st.number_input("Total Weight (Kgs)", min_value=0.5, max_value=30.0, step=0.5)
        material_type = st.selectbox("Category of Material", ["General Solid (Apparel, Books, Machinery)", "Samples", "Liquid/Powder (Chemicals, Perfumes, Meds)"])
    with col2:
        doc_status = st.radio("Courier Classification", ["Document (PAK)", "Non-Document (Package)"])
        can_split = st.radio("Can this package be split to leverage weight slabs?", ["No", "Yes"])
    
    st.markdown("<br>", unsafe_allow_html=True)
    submit = st.form_submit_button("Calculate Best Route 🚀")

# --- EXECUTION LOGIC ---
if submit:
    is_doc = doc_status == "Document (PAK)"
    
    if is_doc and weight > 2.5:
        st.error("❌ **Error:** A Document cannot be more than 2.5 Kg. Please change the Courier Classification to 'Non-Document (Package)' or reduce the total weight.")
    elif not country:
        st.error("⚠️ Please select a destination country from the dropdown before calculating.")
    else:
        st.subheader(f"Routing Analysis: {weight}kg to {country}")
        
        rates = {
            "CITI DHL": get_rate('CITI', is_doc, weight, country, data),
            "FEDEX": get_rate('FEDEX', is_doc, weight, country, data),
            "RATI DHL": get_rate('RATI', is_doc, weight, country, data)
        }
        valid_rates = {k: v for k, v in rates.items() if v is not None and not math.isnan(v)}
        
        if valid_rates:
            best_vendor = min(valid_rates, key=valid_rates.get)
            st.success(f"🏆 Primary Recommendation: **{best_vendor}** at **₹{valid_rates[best_vendor]:,.2f}**")
            
            comparison_df = pd.DataFrame(list(valid_rates.items()), columns=["Vendor", "Rate (₹)"]).sort_values("Rate (₹)")
            st.dataframe(comparison_df, use_container_width=True, hide_index=True)
            
            if can_split == "Yes" and weight > 1.0:
                raw_half = weight / 2
                half_weight = math.ceil(raw_half * 2) / 2
                split_rates = {
                    "CITI DHL": get_rate('CITI', is_doc, half_weight, country, data),
                    "FEDEX": get_rate('FEDEX', is_doc, half_weight, country, data),
                    "RATI DHL": get_rate('RATI', is_doc, half_weight, country, data)
                }
                valid_split = {k: v for k, v in split_rates.items() if v is not None and not math.isnan(v)}
                if valid_split:
                    best_split_vendor = min(valid_split, key=valid_split.get)
                    split_price = valid_split[best_split_vendor] * 2
                    if split_price < valid_rates[best_vendor]:
                        st.success(f"💡 **Optimization Identified:** By splitting your {weight}kg shipment into two {half_weight}kg packages via {best_split_vendor}, the total cost drops to **₹{split_price:,.2f}**, saving you **₹{valid_rates[best_vendor] - split_price:,.2f}**!")
                    else:
                        st.info(f"📊 **Split Test Checked:** Sending two {half_weight}kg packages would cost ₹{split_price:,.2f}. It remains cheaper to send as a single {weight}kg package.")
                        
            if "Liquid" in material_type:
                st.warning("⚠ **Compliance Flag:** Material classified as Liquid/Powder/Chemical. Ensure the MSDS (Material Safety Data Sheet) is attached prior to waybill generation to avoid dangerous goods surcharges.")
        else:
            st.error(f"Pricing data unavailable for '{country}'. Selected weight combination is missing from the vendor rate sheets.")
