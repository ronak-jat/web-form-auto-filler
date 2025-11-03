import streamlit as st
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import time
import io

# --- Page configuration ---
st.set_page_config(
    page_title="Microsoft Forms Auto-Filler",
    page_icon="🤖",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- Custom CSS for stunning styling ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;600;700&display=swap');

    .main {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        font-family: 'Poppins', sans-serif;
    }
h2 {
        color: #667eea !important;
        font-weight: 600 !important;
        padding-bottom: 10px;
        margin-top: 20px !important;
    }

    h3 {
        color: #667eea !important;
        font-weight: 600 !important;
        padding-bottom: 10px;
        margin-top: 20px !important;
        border-bottom: none !important;
        border: none !important;
    }
            
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #667eea 0%, #764ba2 100%);
    }
    [data-testid="stSidebar"] * { color: white !important; }

    [data-testid="stVerticalBlock"] > [style*="flex-direction: column;"] > [data-testid="stVerticalBlock"] {
        background: rgba(255, 255, 255, 0.95);
        border-radius: 20px;
        padding: 30px;
        box-shadow: 0 10px 40px rgba(0, 0, 0, 0.3);
        backdrop-filter: blur(10px);
    }

    h1 {
        color: white !important;
        font-weight: 700 !important;
        text-align: center;
        text-shadow: 2px 2px 8px rgba(0,0,0,0.3);
        font-size: 3.5em !important;
        margin-bottom: 10px !important;
    }

    h2, h3 {
        color: #667eea !important;
        font-weight: 600 !important;
        border-bottom: 3px solid #667eea;
        padding-bottom: 10px;
        margin-top: 20px !important;
    }

    .stTextInput > div > div > input,
    .stNumberInput > div > div > input,
    .stSelectbox > div > div > select {
        border: 2px solid #e0e0e0 !important;
        border-radius: 10px !important;
        font-size: 1em !important;
        padding: 10px !important;
        transition: all 0.3s ease !important;
    }

    .stTextInput > div > div > input:focus,
    .stNumberInput > div > div > input:focus,
    .stSelectbox > div > div > select:focus {
        border-color: #667eea !important;
        box-shadow: 0 0 0 3px rgba(102, 126, 234, 0.2) !important;
    }

    .stButton > button {
        width: 100%;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%) !important;
        color: white !important;
        border: none !important;
        border-radius: 15px !important;
        padding: 15px !important;
        font-size: 1.2em !important;
        font-weight: 700 !important;
        box-shadow: 0 8px 20px rgba(102, 126, 234, 0.4) !important;
        transition: all 0.3s ease !important;
        margin-top: 20px !important;
    }

    .stButton > button:hover {
        transform: translateY(-3px) !important;
        box-shadow: 0 12px 30px rgba(102, 126, 234, 0.6) !important;
    }

    .stAlert {
        border-radius: 15px !important;
        border: none !important;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1) !important;
        font-weight: 500 !important;
    }

    .stProgress > div > div > div {
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%) !important;
        border-radius: 10px !important;
        height: 20px !important;
    }

    [data-testid="stFileUploader"] {
        background: rgba(102, 126, 234, 0.05);
        border: 2px dashed #667eea;
        border-radius: 10px;
        padding: 15px;
        transition: all 0.3s ease;
    }

    [data-testid="stFileUploader"]:hover {
        background: rgba(102, 126, 234, 0.1);
        border-color: #764ba2;
    }

    .browser-mode-card {
        background: linear-gradient(135deg, rgba(102, 126, 234, 0.1) 0%, rgba(118, 75, 162, 0.1) 100%);
        border: 2px solid #667eea;
        border-radius: 15px;
        padding: 20px;
        margin: 10px 0;
        transition: all 0.3s ease;
    }

    .browser-mode-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 8px 20px rgba(102, 126, 234, 0.3);
    }

    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    </style>
""", unsafe_allow_html=True)

# Initialize session state
if 'df' not in st.session_state:
    st.session_state.df = None
if 'input_mapping' not in st.session_state:
    st.session_state.input_mapping = {}
if 'choice_mapping' not in st.session_state:
    st.session_state.choice_mapping = {}
if 'excel_sheets' not in st.session_state:
    st.session_state.excel_sheets = None
if 'uploaded_file_content' not in st.session_state:
    st.session_state.uploaded_file_content = None
if 'browser_mode' not in st.session_state:
    st.session_state.browser_mode = "visible"

# --- Header with animation ---
st.markdown('<h1 style="color: #cbaacb;"> Microsoft Forms Auto-Filler</h1>', unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# --- Step 1: Form URL ---
st.markdown("###  Enter Your Microsoft Forms URL")
form_url = st.text_input(
    "Form URL",
    placeholder="https://forms.office.com/...",
    help="Paste your Microsoft Forms URL here",
    label_visibility="collapsed"
)

# --- Step 2: Upload Data File ---
st.markdown("###  Upload Your Data File")
uploaded_file = st.file_uploader(
    "Upload CSV or Excel file",
    type=["csv", "xlsx", "xls"],
    help="Upload your data file containing form responses",
    label_visibility="collapsed"
)

if uploaded_file:
    # Store file content for multiple reads
    if st.session_state.uploaded_file_content is None or uploaded_file.name != st.session_state.get('uploaded_file_name'):
        st.session_state.uploaded_file_content = uploaded_file.read()
        st.session_state.uploaded_file_name = uploaded_file.name
        uploaded_file.seek(0)
    
    # Handle CSV files
    if uploaded_file.name.endswith('.csv'):
        st.session_state.df = pd.read_csv(io.BytesIO(st.session_state.uploaded_file_content))
        st.session_state.excel_sheets = None
        st.success("✅ CSV file loaded successfully!")
    
    # Handle Excel files
    elif uploaded_file.name.endswith(('.xlsx', '.xls')):
        try:
            st.session_state.df = pd.read_excel(io.BytesIO(st.session_state.uploaded_file_content))
            st.session_state.excel_sheets = None
            st.success("✅ Excel file loaded successfully!")
                
        except Exception as e:
            st.error(f"❌ Error reading Excel file: {e}")
            st.session_state.df = None

    # Display data preview
    if st.session_state.df is not None:
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("####  Data Preview")
        
        col1, col2, col3 = st.columns([2, 1, 3])
        with col1:
            preview_rows = st.slider("Preview rows", min_value=3, max_value=min(50, len(st.session_state.df)), value=min(5, len(st.session_state.df)))
        
        st.dataframe(st.session_state.df.head(preview_rows), use_container_width=True)
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        # --- Browser Mode Selection ---
        st.markdown("###  Browser Mode Selection")
        
        col1, col2 = st.columns(2, gap="large")
        
        with col1:
            st.markdown("""
                <div class='browser-mode-card'>
                    <h4 style='color: #667eea; margin-top: 0;'>👁️ Visible Mode</h4>
                    <p><strong>Watch the automation in action!</strong></p>
                    <ul>
                        <li>See forms being filled in real-time</li>
                        <li>Better for debugging and verification</li>
                    </ul>
                </div>
            """, unsafe_allow_html=True)
            if st.button(" Use Visible Mode", key="visible_btn", use_container_width=True):
                st.session_state.browser_mode = "visible"
                st.success("✅ Visible mode selected!")
        
        with col2:
            st.markdown("""
                <div class='browser-mode-card'>
                    <h4 style='color: #764ba2; margin-top: 0;'>⚡ Headless Mode</h4>
                    <p><strong>Fast background processing!</strong></p>
                    <ul>
                        <li>Runs in the background (no browser window)</li>
                        <li>Faster automation speed</li>
                    </ul>
                </div>
            """, unsafe_allow_html=True)
            if st.button(" Use Headless Mode", key="headless_btn", use_container_width=True):
                st.session_state.browser_mode = "headless"
                st.success("✅ Headless mode selected!")
        
        # Display current mode
        mode_emoji = "👁️" if st.session_state.browser_mode == "visible" else "⚡"
        mode_text = "Visible" if st.session_state.browser_mode == "visible" else "Headless"
        st.info(f"{mode_emoji} **Current Mode:** {mode_text} Mode")
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        # --- Step 3: Field Configuration ---
        st.markdown("###  Configure Field Mappings")
        
        col1, col2 = st.columns(2, gap="large")
        
        with col1:
            st.markdown("####  Text Input Fields")
            st.info("Map columns to text input boxes (Name, Age, Email, etc.)")
            num_input_fields = st.number_input("Number of text fields", min_value=0, max_value=20, value=0, step=1)
            input_mapping = {}
            for i in range(int(num_input_fields)):
                with st.container():
                    st.markdown(f"<b>Input Field {i+1}</b></div>", unsafe_allow_html=True)
                    csv_col = st.selectbox(
                        f"Column",
                        options=[""] + list(st.session_state.df.columns),
                        key=f"input_csv_{i}",
                        label_visibility="collapsed"
                    )
                    if csv_col:
                        input_mapping[csv_col] = {"order": i+1, "type": "input"}
            st.session_state.input_mapping = input_mapping

        with col2:
            st.markdown("####  Choice Fields")
            st.info("Map columns to radio buttons or checkboxes (Gender, Status, etc.)")
            num_choice_fields = st.number_input("Number of choice fields", min_value=0, max_value=20, value=0, step=1)
            choice_mapping = {}
            for i in range(int(num_choice_fields)):
                with st.container():
                    st.markdown(f"<b>Choice Field {i+1}</b></div>", unsafe_allow_html=True)
                    csv_col = st.selectbox(
                        f"Column",
                        options=[""] + list(st.session_state.df.columns),
                        key=f"choice_csv_{i}",
                        label_visibility="collapsed"
                    )
                    if csv_col:
                        choice_mapping[csv_col] = {"order": i+1, "type": "choice"}
            st.session_state.choice_mapping = choice_mapping

        st.markdown("<br>", unsafe_allow_html=True)
        
        # --- Step 4: Summary and Start ---
        
        st.markdown("###  Review Configuration & Start")
        
        if st.session_state.input_mapping or st.session_state.choice_mapping:
            col1, col2 = st.columns(2, gap="large")
            
            with col1:
                st.markdown("####  Text Input Mappings")
                sorted_inputs = sorted(st.session_state.input_mapping.items(), key=lambda x: x[1]['order'])
                for idx, (col, info) in enumerate(sorted_inputs, 1):
                    st.markdown(f"**{idx}.** `{col}`")
            
            with col2:
                st.markdown("####  Choice Field Mappings")
                sorted_choices = sorted(st.session_state.choice_mapping.items(), key=lambda x: x[1]['order'])
                for idx, (col, info) in enumerate(sorted_choices, 1):
                    st.markdown(f"**{idx}.** `{col}`")
            
            # Start Button
            mode_emoji = "👁️" if st.session_state.browser_mode == "visible" else "⚡"
            button_text = f"{mode_emoji} START AUTOMATION ({st.session_state.browser_mode.upper()} MODE)"
            
            if st.button(button_text, type="primary"):
                if not form_url:
                    st.error("❌ Please enter the form URL first.")
                else:
                    progress = st.progress(0)
                    status_text = st.empty()
                    log_container = st.container()
                    
                    try:
                        # Configure Chrome options based on mode
                        chrome_options = Options()
                        
                        if st.session_state.browser_mode == "headless":
                            chrome_options.add_argument("--headless")
                            chrome_options.add_argument("--disable-gpu")
                            chrome_options.add_argument("--no-sandbox")
                            chrome_options.add_argument("--disable-dev-shm-usage")
                            chrome_options.add_argument("--window-size=1920,1080")
                            with st.spinner('⚡ Initializing headless browser...'):
                                driver = webdriver.Chrome(options=chrome_options)
                            status_text.info(f"⚡ Running in headless mode (background)")
                        else:
                            with st.spinner('🔄 Initializing visible browser...'):
                                driver = webdriver.Chrome(options=chrome_options)
                                driver.maximize_window()
                            status_text.info(f"👁️ Running in visible mode (you can watch the automation)")
                        
                        total = len(st.session_state.df)
                        
                        sorted_inputs = sorted(st.session_state.input_mapping.items(), key=lambda x: x[1]['order'])
                        sorted_choices = sorted(st.session_state.choice_mapping.items(), key=lambda x: x[1]['order'])
                        
                        automation_stopped = False
                        for i, row in st.session_state.df.iterrows():
                            with log_container:
                                st.markdown(f"<b>📝 Processing Entry {i+1}/{total}</b></div>", unsafe_allow_html=True)
                            
                            try:
                                driver.get(form_url)
                                time.sleep(3)
                                
                                # Fill text inputs
                                text_inputs = driver.find_elements(By.CSS_SELECTOR, "input[data-automation-id='textInput']")
                                for idx, (csv_col, info) in enumerate(sorted_inputs):
                                    if idx < len(text_inputs):
                                        value = str(row[csv_col]) if pd.notna(row[csv_col]) else ""
                                        text_inputs[idx].clear()
                                        text_inputs[idx].send_keys(value)
                                        
                                        time.sleep(0.5)
                                
                                # Fill choice fields
                                for idx, (csv_col, info) in enumerate(sorted_choices):
                                    if pd.notna(row[csv_col]):
                                        choice_value = str(row[csv_col]).strip().capitalize()
                                        try:
                                            choice_element = driver.find_element(
                                                By.XPATH,
                                                f"//div[contains(@role,'radio') or contains(@role,'checkbox')]//span[normalize-space()='{choice_value}']"
                                            )
                                            driver.execute_script("arguments[0].scrollIntoView(true);", choice_element)
                                            time.sleep(0.3)
                                            choice_element.click()
                                           
                                            time.sleep(0.5)
                                        except Exception as e:
                                            with log_container:
                                                st.warning(f"⚠️ Could not select '{choice_value}' for {csv_col}")
                                
                                # Submit form
                                try:
                                    submit_btn = WebDriverWait(driver, 5).until(
                                        EC.element_to_be_clickable((By.CSS_SELECTOR, "button[data-automation-id='submitButton']"))
                                    )
                                    submit_btn.click()
                                    with log_container:
                                        st.markdown(f"✅ Submitted Entry {i+1}</div>", unsafe_allow_html=True)
                                    time.sleep(2)
                                except Exception as e:
                                    error_msg = str(e).lower()
                                    if 'invalid session' in error_msg or 'session deleted' in error_msg:
                                        with log_container:
                                            st.error(f"🛑 Browser was closed manually - Automation stopped at entry {i+1}")
                                        automation_stopped = True
                                        break
                                    else:
                                        with log_container:
                                            st.error(f"❌ Submit failed for entry {i+1}")
                                        
                            except Exception as e:
                                error_msg = str(e).lower()
                                if 'invalid session' in error_msg or 'session deleted' in error_msg or 'disconnected' in error_msg:
                                    with log_container:
                                        st.error(f"🛑 Browser was closed manually - Automation stopped at entry {i+1}")
                                    automation_stopped = True
                                    break
                                else:
                                    with log_container:
                                        st.error(f"❌ Error processing entry {i+1}")
                            
                            progress.progress((i + 1) / total)
                        
                        try:
                            driver.quit()
                        except:
                            pass
                        
                        if automation_stopped:
                            status_text.warning("⚠️ Automation stopped - Browser was closed manually")
                        else:
                            status_text.markdown("🎉 AUTOMATION COMPLETED!</div>", unsafe_allow_html=True)
                            st.balloons()
                        
                    except Exception as e:
                        st.error(f"❌ Fatal error: {e}")
                        try:
                            driver.quit()
                        except:
                            pass
        else:
            st.warning("⚠️ Please configure at least one field mapping in Step 3.")
else:
    st.info("👆 Upload a CSV or Excel file to begin configuration.")

# --- Footer ---
st.markdown("---")
st.markdown("""
    <div style='text-align: center; color: white; padding: 30px; font-size: 0.9em;'>
        <p style='font-weight: 600; font-size: 1.1em;'> Microsoft Forms Auto-Filler</p>
        <p>Built with ❤️ using Streamlit & Selenium</p>
    </div>
            
""", unsafe_allow_html=True)

