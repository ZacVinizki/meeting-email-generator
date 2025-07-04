"""
Meeting Follow-up Email Generator
=================================

Internal tool for wealth management team to convert client meeting recordings
into professional follow-up emails.

Setup Instructions:
1. Install dependencies: pip install -r requirements.txt
2. Configure Streamlit secrets
3. Run the app: streamlit run app.py
4. Upload audio files and generate professional follow-up emails

Author: Investment Fund AI Team
"""

import streamlit as st
import os
import datetime
import json
from pathlib import Path
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import openai
import uuid
import time
import msal
import requests
import pandas as pd

# PASSWORD PROTECTION
def check_password():
    """Returns True if the user entered the correct password."""
    
    def password_entered():
        """Checks whether a password entered by the user is correct."""
        entered_password = st.session_state["password"].strip()
        
        # Allow different capitalizations of "Morris Ewing"
        correct_passwords = [
            "morris ewing",
            "Morris Ewing", 
            "Morris ewing",
            "morris Ewing",
            "MORRIS EWING"
        ]
        
        if entered_password in correct_passwords:
            st.session_state["password_correct"] = True
            del st.session_state["password"]  # Don't store the password
        else:
            st.session_state["password_correct"] = False

    # Return True if password already verified
    if st.session_state.get("password_correct", False):
        return True

    # Show login form
    st.markdown("### 🔒 Enter Passcode to Access Tool")
    st.text_input(
        "Passcode", 
        type="password", 
        on_change=password_entered, 
        key="password",
        placeholder="Enter passcode..."
    )
    
    if "password_correct" in st.session_state:
        if not st.session_state["password_correct"]:
            st.error("😞 Incorrect passcode. Please try again.")
    
    return False

class ExcelOnlineManager:
    def __init__(self):
        self.client_id = os.getenv("MICROSOFT_CLIENT_ID")
        self.client_secret = os.getenv("MICROSOFT_CLIENT_SECRET")
        self.tenant_id = os.getenv("MICROSOFT_TENANT_ID")
        self.excel_file_id = os.getenv("EXCEL_FILE_ID")
        
        self.authority = f"https://login.microsoftonline.com/{self.tenant_id}"
        self.scope = ["https://graph.microsoft.com/.default"]  # Application scope
        self.graph_url = "https://graph.microsoft.com/v1.0"
        
    def get_app_token(self):
        """Get application token - NO USER INTERACTION NEEDED"""
        try:
            app = msal.ConfidentialClientApplication(
                client_id=self.client_id,
                client_credential=self.client_secret,
                authority=self.authority
            )
            
            # Get token using client credentials (autonomous)
            result = app.acquire_token_for_client(scopes=self.scope)
            
            if "access_token" in result:
                return result["access_token"]
            else:
                st.error(f"Token failed: {result.get('error_description', 'Unknown error')}")
                return None
        except Exception as e:
            st.error(f"Auth error: {str(e)}")
            return None
    
    def add_tasks_to_excel(self, client_name: str, tasks: list) -> bool:
        """Add tasks autonomously - NO USER INTERACTION"""
        
        # Get app token automatically
        token = self.get_app_token()
        if not token:
            return False
        
        headers = {
            'Authorization': f'Bearer {token}',
            'Content-Type': 'application/json'
        }
        
        try:
            # Get next empty row
            url = f"{self.graph_url}/drives/b!{self.excel_file_id}/root/workbook/worksheets/Sheet1/usedRange"
            response = requests.get(url, headers=headers)
            
            next_row = 2
            if response.status_code == 200:
                used_range = response.json()
                if 'rowCount' in used_range:
                    next_row = used_range['rowCount'] + 1
            
            # Prepare tasks
            current_date = datetime.datetime.now().strftime('%Y-%m-%d')
            values = []
            for task in tasks:
                row_data = [
                    client_name, task, current_date, "Pending", 
                    "Meeting Follow-up", "Medium", "James", ""
                ]
                values.append(row_data)
            
            # Add to Excel
            start_row = next_row
            end_row = next_row + len(tasks) - 1
            range_address = f"A{start_row}:H{end_row}"
            
            body = {"values": values}
            
            url = f"{self.graph_url}/drives/b!{self.excel_file_id}/root/workbook/worksheets/Sheet1/range(address='{range_address}')"
            response = requests.patch(url, headers=headers, json=body)
            
            return response.status_code == 200
                
        except Exception as e:
            st.error(f"Excel error: {str(e)}")
            return False
            
def test_excel_connection():
    """Test Excel connection with user authentication"""
    excel_manager = ExcelOnlineManager()
    
    if 'excel_access_token' in st.session_state:
        st.success("✅ Already authenticated! Excel connection ready.")
        return True
    else:
        st.info("🔐 Authentication required for Excel access")
        return excel_manager.authenticate_user()
        
# Configure page
st.set_page_config(
    page_title="Meeting Follow-up Generator",
    page_icon="🤖",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS styling
st.markdown("""
<style>
    /* NUCLEAR DARK THEME - OVERRIDE EVERYTHING */

    /* Force entire app to be black */
    .stApp, .stApp > div, .main, .block-container, .element-container {
        background: #000000 !important;
        color: #ffffff !important;
    }

    /* App background */
    .stApp {
        background: linear-gradient(135deg, #000000 0%, #111111 50%, #000000 100%) !important;
    }

    /* Main content area background */
    .main .block-container {
        background: rgba(0, 0, 0, 0.95) !important;
        border-radius: 20px;
        border: 1px solid #333;
        backdrop-filter: blur(10px);
        box-shadow: 0 8px 32px rgba(0, 0, 0, 0.4);
        color: #ffffff !important;
        padding: 2rem;
    }
    /* SIDEBAR - FORCE PURE BLACK */
    .css-1d391kg, .css-1d391kg > div, .css-1d391kg section, .css-1d391kg .stMarkdown,
    .css-1d391kg .element-container, .css-1d391kg .block-container,
    [data-testid="stSidebar"], [data-testid="stSidebar"] > div, 
    [data-testid="stSidebar"] section, [data-testid="stSidebar"] .stMarkdown {
        background: #000000 !important;
        background-color: #000000 !important;
        color: #ffffff !important;
        border-right: 3px solid #00d4ff !important;
        box-shadow: 5px 0 20px rgba(0, 212, 255, 0.3);
    }

    /* Force sidebar elements to be white text on black */
    .css-1d391kg *, [data-testid="stSidebar"] * {
        color: #ffffff !important;
        background: transparent !important;
    }

    /* Headers */
    h1 {
        color: #00d4ff !important;
        text-align: center;
        font-weight: 700;
        font-size: 3rem !important;
        margin-bottom: 0.5rem !important;
        text-shadow: 0 0 30px rgba(0, 212, 255, 0.8);
    }

    h2 {
        color: #ffffff !important;
        font-weight: 600;
        border-bottom: 2px solid #00d4ff;
        padding-bottom: 0.5rem;
        text-shadow: 0 0 10px rgba(255, 255, 255, 0.3);
    }

    h3 {
        color: #00d4ff !important;
        font-weight: 500;
        text-shadow: 0 0 10px rgba(0, 212, 255, 0.5);
    }

    /* FORCE ALL TEXT WHITE */
    .stMarkdown, .stText, p, span, div, label, .stMarkdown p, .stMarkdown div {
        color: #ffffff !important;
    }

    /* Labels bright white */
    .stTextInput label, .stFileUploader label, .stCheckbox label, 
    .stSelectbox label, .stSlider label, .stRadio label {
        color: #ffffff !important;
        font-weight: 600;
        text-shadow: 0 0 5px rgba(255, 255, 255, 0.3);
    }

    /* Help text */
    .stTextInput .help, .stFileUploader .help, .stCheckbox .help {
        color: #cccccc !important;
    }

    /* ULTRA SEXY BUTTONS */
    .stButton > button {
        background: linear-gradient(135deg, #00d4ff 0%, #0099cc 100%) !important;
        color: #000000 !important;
        font-weight: 700;
        border: none;
        border-radius: 15px;
        padding: 1rem 2.5rem;
        font-size: 1.1rem;
        transition: all 0.4s ease;
        box-shadow: 0 8px 25px rgba(0, 212, 255, 0.4);
        text-transform: uppercase;
        letter-spacing: 1px;
    }

    .stButton > button:hover {
        background: linear-gradient(135deg, #ffffff 0%, #e6e6e6 100%) !important;
        box-shadow: 0 12px 35px rgba(0, 212, 255, 0.7);
        transform: translateY(-3px) scale(1.02);
        color: #000000 !important;
    }

    /* Form submit buttons */
    .stFormSubmitButton > button {
        background: linear-gradient(135deg, #00d4ff 0%, #0099cc 100%) !important;
        color: #000000 !important;
        font-weight: 700;
        border: none;
        border-radius: 15px;
        padding: 1rem 2rem;
        font-size: 1.1rem;
        width: 100%;
        text-transform: uppercase;
        letter-spacing: 1px;
        box-shadow: 0 8px 25px rgba(0, 212, 255, 0.4);
        transition: all 0.4s ease;
    }

    /* INPUT FIELDS - PURE BLACK */
    .stTextInput > div > div > input {
        background: #000000 !important;
        color: #ffffff !important;
        border: 2px solid #333 !important;
        border-radius: 15px;
        padding: 1rem;
        font-size: 1rem;
        transition: all 0.3s ease;
        box-shadow: inset 0 2px 5px rgba(0, 0, 0, 0.3);
    }

    .stTextInput > div > div > input:focus {
        border-color: #00d4ff !important;
        box-shadow: 0 0 20px rgba(0, 212, 255, 0.5), inset 0 2px 5px rgba(0, 0, 0, 0.3);
        background: #000000 !important;
        color: #ffffff !important;
    }

    .stTextInput input::placeholder {
        color: #888888 !important;
    }

 /* FILE UPLOADER - CLEAN SINGLE BORDER */
.stFileUploader > div {
    background: #000000 !important;
    border: 3px dashed #00d4ff !important;
    border-radius: 20px;
    padding: 3rem 2rem;
    text-align: center;
    color: #ffffff !important;
    transition: all 0.3s ease;
    box-shadow: inset 0 4px 10px rgba(0, 0, 0, 0.5);
}

.stFileUploader > div:hover {
    border-color: #ffffff !important;
    box-shadow: 0 0 30px rgba(0, 212, 255, 0.3), inset 0 4px 10px rgba(0, 0, 0, 0.5);
}

/* File uploader content - no additional borders */
.stFileUploader div[data-testid="stFileUploadDropzone"] {
    background: transparent !important;
    border: none !important;
    padding: 0;
}

    .stFileUploader div[data-testid="stFileUploadDropzone"]:hover {
        border-color: #ffffff !important;
        box-shadow: 0 0 30px rgba(0, 212, 255, 0.3), inset 0 4px 10px rgba(0, 0, 0, 0.5);
    }

    /* Force file uploader text white */
    .stFileUploader label, .stFileUploader p, .stFileUploader span,
    .stFileUploader div[data-testid="stFileUploadDropzone"] *,
    .stFileUploader * {
        color: #ffffff !important;
        background: transparent !important;
        font-weight: 500;
    }

    /* Text areas */
    .stTextArea > div > div > textarea {
        background: #000000 !important;
        color: #ffffff !important;
        border: 2px solid #333 !important;
        border-radius: 15px;
        font-family: 'Courier New', monospace;
        box-shadow: inset 0 2px 5px rgba(0, 0, 0, 0.3);
    }

    /* CHECKBOX - GLOWING WHITE OUTLINE */
.stCheckbox label, .stCheckbox span, .stCheckbox div {
    color: #ffffff !important;
    font-weight: 500;
}

/* Checkbox input styling */
.stCheckbox input[type="checkbox"] {
    appearance: none;
    width: 20px;
    height: 20px;
    border: 2px solid #ffffff;
    border-radius: 4px;
    background: transparent;
    position: relative;
    cursor: pointer;
    transition: all 0.3s ease;
    box-shadow: 0 0 10px rgba(255, 255, 255, 0.3);
}

.stCheckbox input[type="checkbox"]:hover {
    border-color: #00d4ff;
    box-shadow: 0 0 15px rgba(0, 212, 255, 0.5);
    background: rgba(0, 212, 255, 0.1);
}

.stCheckbox input[type="checkbox"]:checked {
    background: linear-gradient(135deg, #00d4ff 0%, #0099cc 100%);
    border-color: #00d4ff;
    box-shadow: 0 0 20px rgba(0, 212, 255, 0.6);
}

.stCheckbox input[type="checkbox"]:checked::after {
    content: "✓";
    position: absolute;
    top: 50%;
    left: 50%;
    transform: translate(-50%, -50%);
    color: #000000;
    font-weight: bold;
    font-size: 14px;
}

/* Any other input elements that might be hard to see */
.stSelectbox > div > div {
    background: rgba(0, 0, 0, 0.8) !important;
    border: 2px solid #ffffff !important;
    border-radius: 12px;
    box-shadow: 0 0 10px rgba(255, 255, 255, 0.2);
    color: #ffffff !important;
}

.stSelectbox > div > div:hover {
    border-color: #00d4ff !important;
    box-shadow: 0 0 15px rgba(0, 212, 255, 0.4);
}

/* Radio buttons */
.stRadio input[type="radio"] {
    appearance: none;
    width: 18px;
    height: 18px;
    border: 2px solid #ffffff;
    border-radius: 50%;
    background: transparent;
    cursor: pointer;
    transition: all 0.3s ease;
    box-shadow: 0 0 8px rgba(255, 255, 255, 0.3);
}

.stRadio input[type="radio"]:hover {
    border-color: #00d4ff;
    box-shadow: 0 0 12px rgba(0, 212, 255, 0.5);
}

.stRadio input[type="radio"]:checked {
    background: linear-gradient(135deg, #00d4ff 0%, #0099cc 100%);
    border-color: #00d4ff;
    box-shadow: 0 0 15px rgba(0, 212, 255, 0.6);
}

.stRadio input[type="radio"]:checked::after {
    content: "";
    position: absolute;
    top: 50%;
    left: 50%;
    transform: translate(-50%, -50%);
    width: 8px;
    height: 8px;
    border-radius: 50%;
    background: #000000;
}

/* Slider styling */
.stSlider > div > div > div {
    background: #333 !important;
}

.stSlider > div > div > div > div {
    background: linear-gradient(135deg, #00d4ff 0%, #0099cc 100%) !important;
    box-shadow: 0 0 10px rgba(0, 212, 255, 0.4);
}

/* Number input */
.stNumberInput > div > div > input {
    background: rgba(0, 0, 0, 0.8) !important;
    color: #ffffff !important;
    border: 2px solid #ffffff !important;
    border-radius: 12px;
    box-shadow: 0 0 10px rgba(255, 255, 255, 0.2);
}

.stNumberInput > div > div > input:focus {
    border-color: #00d4ff !important;
    box-shadow: 0 0 15px rgba(0, 212, 255, 0.4);
}

    /* SUCCESS/ERROR MESSAGES */
    .stSuccess {
        background: linear-gradient(135deg, #00ff88 0%, #00cc6a 100%) !important;
        color: #000000 !important;
        border-radius: 15px;
        border: none;
        box-shadow: 0 8px 25px rgba(0, 255, 136, 0.4);
        font-weight: 600;
    }

    .stError {
        background: linear-gradient(135deg, #ff4757 0%, #ff3742 100%) !important;
        color: #ffffff !important;
        border-radius: 15px;
        border: none;
        box-shadow: 0 8px 25px rgba(255, 71, 87, 0.4);
        font-weight: 600;
    }

    .stInfo {
        background: linear-gradient(135deg, #00d4ff 0%, #0099cc 100%) !important;
        color: #000000 !important;
        border-radius: 15px;
        border: none;
        box-shadow: 0 8px 25px rgba(0, 212, 255, 0.4);
        font-weight: 600;
    }

    .stWarning {
        background: linear-gradient(135deg, #ffa502 0%, #ff8c00 100%) !important;
        color: #000000 !important;
        border-radius: 15px;
        border: none;
        box-shadow: 0 8px 25px rgba(255, 165, 2, 0.4);
        font-weight: 600;
    }

    /* TABS - FIXED TEXT VISIBILITY */
.stTabs [data-baseweb="tab-list"] {
    background: #000000 !important;
    border-radius: 15px;
    gap: 10px;
    padding: 0.5rem;
    border: 1px solid #333;
}

.stTabs [data-baseweb="tab"] {
    background: #1a1a1a !important;
    color: #ffffff !important;
    border-radius: 12px;
    padding: 1rem 2rem;
    font-weight: 600;
    transition: all 0.3s ease;
    border: 1px solid #333;
}

.stTabs [aria-selected="true"] {
    background: linear-gradient(135deg, #00d4ff 0%, #0099cc 100%) !important;
    color: #ffffff !important;
    box-shadow: 0 5px 15px rgba(0, 212, 255, 0.5);
    transform: translateY(-2px);
    text-shadow: 0 0 5px rgba(255, 255, 255, 0.8);
}

/* Force tab content text to be white */
.stTabs div[data-baseweb="tab-panel"] {
    color: #ffffff !important;
}

.stTabs div[data-baseweb="tab-panel"] * {
    color: #ffffff !important;
}

    /* DOWNLOAD BUTTONS */
    .stDownloadButton > button {
        background: linear-gradient(135deg, #333333 0%, #1a1a1a 100%) !important;
        color: #ffffff !important;
        border: 2px solid #00d4ff;
        border-radius: 15px;
        padding: 1rem 2rem;
        font-weight: 600;
        transition: all 0.3s ease;
        text-transform: uppercase;
        letter-spacing: 1px;
    }

    .stDownloadButton > button:hover {
        background: linear-gradient(135deg, #00d4ff 0%, #0099cc 100%) !important;
        color: #000000 !important;
        box-shadow: 0 8px 25px rgba(0, 212, 255, 0.5);
        transform: translateY(-2px);
    }

    /* CUSTOM CONTAINERS */
    .email-container {
        background: linear-gradient(135deg, rgba(0, 0, 0, 0.95) 0%, rgba(26, 26, 26, 0.95) 100%) !important;
        border: 2px solid #00d4ff;
        border-radius: 20px;
        padding: 2.5rem;
        margin: 1.5rem 0;
        box-shadow: 0 15px 50px rgba(0, 212, 255, 0.2);
        backdrop-filter: blur(15px);
        color: #ffffff !important;
    }

    .login-container {
        background: linear-gradient(135deg, rgba(0, 0, 0, 0.95) 0%, rgba(26, 26, 26, 0.95) 100%) !important;
        border: 3px solid #00d4ff;
        border-radius: 25px;
        padding: 4rem;
        margin: 3rem auto;
        max-width: 500px;
        box-shadow: 0 20px 60px rgba(0, 212, 255, 0.3);
        text-align: center;
        color: #ffffff !important;
        backdrop-filter: blur(20px);
    }

    /* FOOTER */
    .footer {
        background: linear-gradient(135deg, rgba(0, 0, 0, 0.95) 0%, rgba(26, 26, 26, 0.9) 100%) !important;
        border-top: 3px solid #00d4ff;
        margin-top: 3rem;
        padding: 2.5rem;
        border-radius: 20px 20px 0 0;
        color: #ffffff !important;
        box-shadow: 0 -10px 30px rgba(0, 212, 255, 0.2);
    }

    /* SCROLLBAR */
    ::-webkit-scrollbar {
        width: 12px;
    }

    ::-webkit-scrollbar-track {
        background: #000000;
        border-radius: 6px;
    }

    ::-webkit-scrollbar-thumb {
        background: linear-gradient(135deg, #00d4ff 0%, #0099cc 100%);
        border-radius: 6px;
        box-shadow: 0 2px 10px rgba(0, 212, 255, 0.3);
    }

    ::-webkit-scrollbar-thumb:hover {
        background: linear-gradient(135deg, #ffffff 0%, #00d4ff 100%);
        box-shadow: 0 4px 15px rgba(0, 212, 255, 0.5);
    }

    /* FORCE OVERRIDE ANY REMAINING WHITE BACKGROUNDS */
    .stApp *, .main *, .css-1d391kg *, [data-testid="stSidebar"] * {
        background-color: transparent !important;
    }

    /* Ensure no white anywhere */
    .stApp > div, .main > div, .block-container > div,
    .element-container > div, .stMarkdown > div {
        background: transparent !important;
        color: #ffffff !important;
    }

    /* Final nuclear option - force everything */
    * {
        color: inherit !important;
    }

    .stApp, .main, .css-1d391kg, [data-testid="stSidebar"] {
        background: #000000 !important;
    }

    /* File uploader icons */
    .stFileUploader svg {
        color: #00d4ff !important;
        fill: #00d4ff !important;
    }

    /* Spinner */
    .stSpinner > div {
        border-top-color: #00d4ff !important;
        border-width: 4px !important;
    }

    .stSpinner + div {
        color: #ffffff !important;
        font-weight: 500;
    }
    /* FIX TOP RIGHT MENU DROPDOWN */
.stApp > header {
    background: transparent !important;
}

/* Main menu button */
[data-testid="stHeader"] {
    background: transparent !important;
}

[data-testid="stHeader"] button {
    color: #ffffff !important;
    background: rgba(0, 0, 0, 0.8) !important;
    border: 1px solid #00d4ff !important;
    border-radius: 8px;
}

[data-testid="stHeader"] button:hover {
    background: rgba(0, 212, 255, 0.2) !important;
    box-shadow: 0 0 10px rgba(0, 212, 255, 0.5);
}

/* Dropdown menu background */
[data-baseweb="popover"] {
    background: #000000 !important;
    border: 2px solid #00d4ff !important;
    border-radius: 15px !important;
    box-shadow: 0 10px 30px rgba(0, 212, 255, 0.3) !important;
}

[data-baseweb="popover"] > div {
    background: #000000 !important;
    color: #ffffff !important;
}

/* Menu items */
[data-baseweb="popover"] button, 
[data-baseweb="popover"] a,
[data-baseweb="popover"] div {
    background: #000000 !important;
    color: #ffffff !important;
    border-radius: 8px !important;
    margin: 0.2rem 0 !important;
    padding: 0.5rem 1rem !important;
    transition: all 0.3s ease !important;
}

[data-baseweb="popover"] button:hover,
[data-baseweb="popover"] a:hover {
    background: rgba(0, 212, 255, 0.2) !important;
    color: #ffffff !important;
    box-shadow: 0 0 10px rgba(0, 212, 255, 0.3) !important;
}

/* Settings text and icons */
[data-baseweb="popover"] span,
[data-baseweb="popover"] svg {
    color: #ffffff !important;
    fill: #ffffff !important;
}

/* Force menu text visibility */
[data-testid="stMainMenu"] {
    background: #000000 !important;
    color: #ffffff !important;
}

[data-testid="stMainMenu"] * {
    color: #ffffff !important;
    background: transparent !important;
}
</style>
""", unsafe_allow_html=True)

# Initialize directories
AUDIO_DIR = Path("audio_files")
EMAILS_DIR = Path("emails")
AUDIO_DIR.mkdir(exist_ok=True)
EMAILS_DIR.mkdir(exist_ok=True)

# Configure OpenAI
openai.api_key = os.getenv("OPENAI_API_KEY")

class MeetingEmailGenerator:
    def __init__(self):
        self.client = openai.OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

    def transcribe_audio(self, audio_file_path):
        """Transcribe audio using OpenAI Whisper API"""
        try:
            with open(audio_file_path, "rb") as audio_file:
                transcript = self.client.audio.transcriptions.create(
                    model="whisper-1",
                    file=audio_file,
                    response_format="text"
                )
            return transcript
        except Exception as e:
            st.error(f"Transcription failed: {str(e)}")
            return None

    def generate_email(self, transcript, recipient_name=""):
        """Generate professional follow-up email using GPT-3.5-turbo"""

        prompt = f"""
        You are James Telfser, a professional wealth management advisor at Ewing Morris. Based on the meeting transcript below, generate a professional follow-up email that matches your established tone and format.

        Your email style characteristics:
        - Warm, professional greeting
        - Brief friendly opener acknowledging the meeting
        - Clear structure with sections like "Key Takeaways" and "Next Steps"
        - Use bullet points and numbered lists for clarity
        - Bold formatting for important sections (**text**)
        - Professional but personable tone
        - Include specific details and action items
        - End with warm sign-off like "All the best, James" or "Warm regards, James"

        Example structure to follow:
        - Greeting with recipient name
        - Brief intro acknowledging the meeting
        - **Participants** section
        - **Objective** section  
        - **Key Takeaways** with numbered points and sub-bullets
        - **Next Steps** or **Action Items**
        - Professional closing

        Meeting Transcript:
        {transcript}

        Generate a professional follow-up email that summarizes the key discussion points, action items, and next steps. Make it client-friendly and well-formatted.Below are a few examples of emails I have already sent. Try and replicate this style to your best ability:
        
        Hi Aidan,

 

Thanks for the time today – exciting times for you!  Keep it going.

 

I thought it would be helpful to follow up with some notes to help with your understanding and to make sure we are on the same page.  

 

Please review and provide feedback, as necessary. 

 

-------------

Participants: Aidan Harradence and James Telfser via Teams Meeting.

Objective: 1) Get caught up ; 2) Review Portfolio; 3) Discuss changes to the portfolio based on market conditions & current lifestyle.

 

Key Takeaways:

Investment Accounts:
Performance has improved this year with accounts up ~20% after fees.  
It's been a great year for small/mid cap stocks, and we were waiting on this rebound to rebalance the account which is great (ie. we did not sell at the lows).
We have been raising some cash ($20k) throughout the year to diversify the account into less volatile investments.
We talked about ETFs and how we will be utilizing these more going forward to take advantage of "time arbitrage".
Planning & Personal:
You are currently living in the UK and are pursuing an entrepreneurial venture in the film industry (very cool!).  
You currently have ~5-6 months of living expenses covered assuming you do take on any part time work.
You do not expect to require any of these savings, but would like to keep some in cash as a back up (~$20k).
You reiterated that you are still a beginner when it comes to investing and would like some book recommendations (we have these all in the office next time you are here:
A Random Walk Down Wall Street
Rich Dad Poor Dad
The Psychology of Money
Reporting:
You would like a login for NDEX/NBIN+
You would like to set up a recurring quarterly check in, especially given that you are trying to learn and may require additional income.
Next Steps:

Action Items:
You will review what your CRA details say online about what you can contribute to your TFSA.
James to send an email with an updated asset allocation and changes to the accounts incorporating your updated information and desire for less volatility.
Katherine will send over login details for our systems so you can track the accounts.
 

All the best,
James

Heres another-

Hi everyone,


It was great to see you all yesterday.  Please see below for minutes and some additional details. 

 

As always, please let me know if I have missed anything or if you have any questions.


All the best,
James

 

Participants: Brad Jarvis (BJ), Faye Wardrop (FW), Steve Barlow (SB) and James Telfser (JT) – Victoria in person

Objective: 1) Quarterly Account Review;  2) Market Outlook

 

Key Takeaways:

Quarterly Review:
We discussed performance, asset allocation and recent changes.  Overall, Jarvis Group of accounts was +1% in Q1-25.
Major contributors were gold, global equities ex-US, Canadian large cap equities and preferred shares.
Weakness was driven by the large cap US Equity ETFs as technology/consumer discretionary underperformed.
We remained within the Target Asset Allocation ranges as per our recent IPS.
We discussed the changes to the Aventine Funds – Selling down the US Equity Fund (for more passive and fixed income exposure) & an eventual merger of the Aventine Canadian Equity Fund.
We currently have orders in to sell $200k ACE Fund and $450k of the US Equity Fund at April 30th, representing 4.5% additional cash. 
We discussed the NorthStream Credit Fund sell which settles early May ($425,000 or ~3% of the account)
We will hold a higher than normal cash position in the account, however we will add to fixed income & infrastructure (KKR) to round out the asset allocation. 
We discussed continuing to hold a higher percentage of cash/fixed income/gold which we all agreed was a good idea as we deal with the uncertainty at the moment.
We agreed that holding cash/money market, despite lower rates in CAD was the right decision given the volatility in currencies and goal of capital preservation, especially for downpayment cash.
IPS/KYC:
BJ has an offer on Saturna ($1.34mm less expenses) and is working towards a close in the coming months.
BJ also has an offer on one of the lots for $300k less expenses.
It is currently expected that BJ will require $1-$1.2mm in cash for downpayment on Salt Spring in addition to cash from the properties above (sometime in next 12 months)
JT to continue to be conservative with cash and keep enough liquidity for downpayment (currently $2.8mm in cash). 
Upon the withdrawal of the funds for downpayment, JT to rebalance to ensure that we remain in line with IPS targets and current market outlook.
Expenses for BJ seem to be a little higher than originally projected in our IPS from a couple years ago.
Next Steps:
JT to make asset allocation changes in Q2 once aforementioned trades settle in early May. 
JT to continue to build up the passive side of the accounts with proceeds from active mandates, however, will be more conservative than normal given current expectation for volatility and potential cash needs.
JT to include FW on any asset allocation or special investment opportunities communications. 
FW will ensure that the appropriate accounting/legal team is in the loop.
JT to follow up with BJ/FW regarding an updated IPS to incorporate new and more realistic spending targets. 
JT to offer up financial planning exercise to ensure current spending is sustainable under conservative assumptions.
BJ/FW to keep JT informed of any structural changes to the accounts or changes to investment criteria or eligibility.
Have a great day ahead!

James

this was more how he would write the email if it was to a 3rd party or a larg client not an individual investor:

heres a final example :

Hi Yvonne,


It was great to see you in person yesterday!  Let's make sure to do that at least once per year 😊.  Look forward to hearing about your 4 generation(!) trip to NFLD.

 

Please see below for minutes and some additional details. 

 

As always, please let me know if I have missed anything or if you have any questions.


All the best,
James

 

Participants: Yvonne Taylor (YT), James Telfser (JT) – In Person

Objective: 1) YTD Account Review;  2) KYC Check up

 

Key Takeaways:

YTD Review:
We discussed performance, asset allocation and recent changes to the accounts and fund structures. Overall, your accounts are up ~1% in 2025, up just shy of 6% per year after fees since inception.  
Major contributors were preferred shares, Canadian large cap equities and global equities (IQLT ETF).  
Weakness was primarily driven by Apple (-20% YTD + USD weakness) and Transforce (-36% YTD).
We remained within the Target Asset Allocation ranges of 0-30% Cash,  15-35% Fixed Income, and 30-70% Equity.  We've kept it simple without allocations to our alternatives sleeve.
We discussed the changes to the Aventine Funds – Selling down the US Equity Fund (for more passive and fixed income exposure) & an eventual merger of the Aventine Canadian Equity Fund.
We currently have orders in to sell all but ~$1,000 of the ACE Fund at the end of June.  
We discussed the EM Flexible Fixed Income Fund and the other EM products.  JT made it clear we will not add additional Ewing Morris products.
We also walked through the conflicts of interest document.
We walked through the various ETFs and the benefits of holdings them (QQQ/HXQ = Tech w/o concentration risk, XIC/TSX = Broad Canada, resource exposure w/o having to pick, etc). 
Finally, we discussed the various fixed income ETFs and why we hold them.  Reasons include higher yield and view that the USD will appreciate vs CAD so could make 4% income and an additional 2-3% in currency appreciation.
IPS/KYC:
No change in financial position or asset allocation, currently receiving enough cash from the monthly RRIF payments.
Continue to manage account considering time horizon and avoid any illiquid investments.
No major purchases/withdrawals expected. 
Other/Next Steps:
JT to enter sell order for ACE/Darkhorse
JT reviewed TFSA contribution room, and you CAN contribute the $85,000 (deregistered in 2024) in 2025.
JT to contribute the AirBoss and Preferred Share position to shelter from taxes going forward.
JT to send Addepar login and upload presentation
Completed…please let me know if you have trouble logging in.  
Have a great day ahead!

James

        {"If you can identify the client's name from the transcript, use it in the greeting. Otherwise, use a general greeting." if not recipient_name else f"Address the email to {recipient_name}."}
        """

        try:
            response = self.client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "system",
                     "content": "You are a professional wealth management advisor who writes clear, structured follow-up emails after client meetings."},
                    {"role": "user", "content": prompt}
                ],
                max_tokens=1500,
                temperature=0.7
            )
            return response.choices[0].message.content
        except Exception as e:
            st.error(f"Email generation failed: {str(e)}")
            return None

    def send_email(self, recipient_email, subject, email_body, transcript=None):
        """Send email via SMTP"""
        try:
            # Email configuration
            smtp_server = os.getenv("SMTP_SERVER", "smtp.gmail.com")
            smtp_port = int(os.getenv("SMTP_PORT", "587"))
            sender_email = os.getenv("SENDER_EMAIL")
            sender_password = os.getenv("SENDER_PASSWORD")

            if not all([sender_email, sender_password]):
                st.error("Email credentials not configured. Please check your secrets.")
                return False

            # Create message
            msg = MIMEMultipart()
            msg['From'] = sender_email
            msg['To'] = recipient_email
            msg['Subject'] = subject

            # Add email body
            msg.attach(MIMEText(email_body, 'plain'))

            # Optionally attach transcript
            if transcript and st.session_state.get('include_transcript', False):
                transcript_attachment = MIMEText(transcript, 'plain')
                transcript_attachment.add_header('Content-Disposition', 'attachment', filename='meeting_transcript.txt')
                msg.attach(transcript_attachment)

            # Send email
            server = smtplib.SMTP(smtp_server, smtp_port)
            server.starttls()
            server.login(sender_email, sender_password)
            text = msg.as_string()
            server.sendmail(sender_email, recipient_email, text)
            server.quit()

            return True

        except Exception as e:
            st.error(f"Failed to send email: {str(e)}")
            return False

    def save_email_record(self, recipient_email, subject, email_body, transcript, audio_filename):
        """Save email record for future reference"""
        record = {
            "timestamp": datetime.datetime.now().isoformat(),
            "recipient_email": recipient_email,
            "subject": subject,
            "email_body": email_body,
            "transcript": transcript,
            "audio_filename": audio_filename,
            "id": str(uuid.uuid4())
        }

        filename = f"email_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        filepath = EMAILS_DIR / filename

        with open(filepath, 'w') as f:
            json.dump(record, f, indent=2)

        return filepath

def main():
    if not check_password():
        st.stop()

    # Hero section with new styling
    st.markdown("""
        <div style="text-align: center; margin-bottom: 2rem;">
            <h1 style="color: #00d4ff; text-shadow: 0 0 20px rgba(0, 212, 255, 0.5); margin-bottom: 0.5rem;">
                🤖 Ewing Morris AI Assistant
            </h1>
            <p style="color: #cccccc; font-size: 1.2rem; font-weight: 300; margin-bottom: 0;">
                Transform meeting recordings into professional follow-up emails
            </p>
            <div style="width: 100px; height: 3px; background: linear-gradient(90deg, #00d4ff, #ffffff); margin: 1rem auto;"></div>
        </div>
    """, unsafe_allow_html=True)

    # Initialize session state
    if 'generator' not in st.session_state:
        st.session_state.generator = MeetingEmailGenerator()

    # Sidebar for settings and history
    with st.sidebar:
        st.markdown("""
            <div style="text-align: center; margin-bottom: 2rem;">
                <h2 style="color: #00d4ff; font-size: 1.5rem; margin-bottom: 1rem;">⚙️ Control Panel</h2>
            </div>
        """, unsafe_allow_html=True)

        # API Key check with new styling
        if not os.getenv("OPENAI_API_KEY"):
            st.markdown("""
                <div style="background: linear-gradient(135deg, #ff4757, #ff3742); color: white; padding: 1rem; border-radius: 10px; margin: 1rem 0; text-align: center;">
                    <strong>⚠️ OpenAI API Not Configured</strong><br>
                    <small>Please configure your secrets</small>
                </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown("""
                <div style="background: linear-gradient(135deg, #00ff88, #00cc6a); color: black; padding: 1rem; border-radius: 10px; margin: 1rem 0; text-align: center;">
                    <strong>✅ OpenAI API Active</strong><br>
                    <small>Ready for transcription</small>
                </div>
            """, unsafe_allow_html=True)

        if not all([os.getenv("SENDER_EMAIL"), os.getenv("SENDER_PASSWORD")]):
            st.markdown("""
                <div style="background: linear-gradient(135deg, #ffa502, #ff8c00); color: white; padding: 1rem; border-radius: 10px; margin: 1rem 0; text-align: center;">
                    <strong>⚠️ Email Not Configured</strong><br>
                    <small>Email sending disabled</small>
                </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown("""
                <div style="background: linear-gradient(135deg, #00ff88, #00cc6a); color: black; padding: 1rem; border-radius: 10px; margin: 1rem 0; text-align: center;">
                    <strong>✅ Email System Active</strong><br>
                    <small>Ready to send emails</small>
                </div>
            """, unsafe_allow_html=True)

        # Additional status boxes
        st.markdown("""
            <div style="background: linear-gradient(135deg, #00ff88, #00cc6a); color: black; padding: 1rem; border-radius: 10px; margin: 1rem 0; text-align: center;">
                <strong>✅ Audio Processing Active</strong><br>
                <small>MP3, WAV, M4A ready</small>
            </div>
        """, unsafe_allow_html=True)

        st.markdown("""
            <div style="background: linear-gradient(135deg, #00ff88, #00cc6a); color: black; padding: 1rem; border-radius: 10px; margin: 1rem 0; text-align: center;">
                <strong>✅ Security Active</strong><br>
                <small>Authentication enabled</small>
            </div>
        """, unsafe_allow_html=True)

        st.markdown("---")
        
        # Excel Integration Status
        st.markdown("### 📊 Excel Integration")
        
        # Check Excel configuration
        excel_configured = all([
            os.getenv("MICROSOFT_CLIENT_ID"),
            os.getenv("MICROSOFT_CLIENT_SECRET"), 
            os.getenv("MICROSOFT_TENANT_ID"),
            os.getenv("EXCEL_FILE_ID")
        ])
        
        if excel_configured:
            st.markdown("""
                <div style="background: linear-gradient(135deg, #00ff88, #00cc6a); color: black; padding: 1rem; border-radius: 10px; margin: 1rem 0; text-align: center;">
                    <strong>✅ Excel API Configured</strong><br>
                    <small>Ready to sync tasks</small>
                </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown("""
                <div style="background: linear-gradient(135deg, #ffa502, #ff8c00); color: white; padding: 1rem; border-radius: 10px; margin: 1rem 0; text-align: center;">
                    <strong>⚠️ Excel API Not Configured</strong><br>
                    <small>Task sync disabled</small>
                </div>
            """, unsafe_allow_html=True)
        
        # Test connection button
        if st.button("🔗 Test Excel Connection"):
            with st.spinner("Testing connection..."):
                test_excel_connection()
        
        # Show file ID if configured
        excel_file_id = os.getenv("EXCEL_FILE_ID", "")
        if excel_file_id:
            st.text(f"📄 File ID: {excel_file_id[:8]}...")

        st.markdown("---")

        # Recent emails - COLLAPSIBLE CLEAN DESIGN
        email_files = list(EMAILS_DIR.glob("*.json"))
        email_count = len(email_files)

        # Email count display
        st.markdown(f"""
               <div style="text-align: center; margin-bottom: 1rem;">
                   <span style="color: #666; font-size: 0.8rem;">📊 Total Emails Sent: </span>
                   <span style="color: #00d4ff; font-weight: 600; font-size: 1.1rem;">{email_count}</span>
               </div>
           """, unsafe_allow_html=True)

        # Collapsible email history
        if email_files:
            with st.expander("📁 Email History", expanded=False):
                email_files.sort(key=lambda x: x.stat().st_mtime, reverse=True)

                for email_file in email_files:
                    try:
                        with open(email_file, 'r') as f:
                            record = json.load(f)

                        # Extract email name (before @) and truncate if needed
                        email_name = record['recipient_email'].split('@')[0]
                        if len(email_name) > 15:
                            email_name = email_name[:15] + "..."

                        # Format timestamp
                        timestamp = record['timestamp'][:16].replace('T', ' ')
                        date_part = timestamp.split(' ')[0]
                        time_part = timestamp.split(' ')[1]

                        # Create each email entry
                        st.markdown(f"""
                               <div style="
                                   background: linear-gradient(135deg, rgba(26, 26, 26, 0.8) 0%, rgba(20, 20, 20, 0.8) 100%); 
                                   padding: 1rem; 
                                   border-radius: 12px; 
                                   margin: 0.5rem 0; 
                                   border-left: 4px solid #00d4ff;
                                   border: 1px solid #333;
                                   transition: all 0.3s ease;
                                   position: relative;
                                   overflow: hidden;
                               ">
                                   <div style="display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 0.5rem;">
                                       <div style="flex: 1;">
                                           <div style="color: #00d4ff; font-weight: 700; font-size: 1rem; margin-bottom: 0.3rem; display: flex; align-items: center;">
                                               <span style="margin-right: 0.5rem;">👤</span>
                                               {email_name}
                                           </div>
                                           <div style="color: #cccccc; font-size: 0.85rem; margin-bottom: 0.3rem;">
                                               📧 {record['recipient_email']}
                                           </div>
                                           <div style="display: flex; gap: 1rem; font-size: 0.75rem;">
                                               <span style="color: #888; display: flex; align-items: center;">
                                                   <span style="margin-right: 0.3rem;">📅</span>
                                                   {date_part}
                                               </span>
                                               <span style="color: #888; display: flex; align-items: center;">
                                                   <span style="margin-right: 0.3rem;">🕒</span>
                                                   {time_part}
                                               </span>
                                           </div>
                                       </div>
                                       <div style="display: flex; flex-direction: column; align-items: center; margin-left: 1rem;">
                                           <div style="color: #00ff88; font-size: 1.5rem; margin-bottom: 0.2rem;">
                                               ✓
                                           </div>
                                           <div style="color: #00ff88; font-size: 0.7rem; font-weight: 600;">
                                               SENT
                                           </div>
                                       </div>
                                   </div>
                               </div>
                           """, unsafe_allow_html=True)

                    except (json.JSONDecodeError, KeyError):
                        continue

                # Summary at bottom of expanded section
                st.markdown(f"""
                       <div style="
                           text-align: center; 
                           margin-top: 1.5rem; 
                           padding: 1rem;
                           background: rgba(0, 212, 255, 0.05);
                           border-radius: 10px;
                           border: 1px solid rgba(0, 212, 255, 0.2);
                       ">
                           <span style="color: #00d4ff; font-weight: 600;">📈 Total Communications: {email_count}</span>
                       </div>
                   """, unsafe_allow_html=True)

        else:
            st.markdown("""
                   <div style="
                       text-align: center; 
                       color: #666; 
                       padding: 2rem 1rem;
                       background: rgba(0, 0, 0, 0.3);
                       border-radius: 15px;
                       border: 1px dashed #333;
                       margin-top: 1rem;
                   ">
                       <div style="font-size: 3rem; margin-bottom: 1rem;">📭</div>
                       <p style="color: #888; margin: 0; font-weight: 500;">No emails sent yet</p>
                       <p style="color: #666; font-size: 0.8rem; margin: 0.5rem 0 0 0;">Your email history will appear here</p>
                   </div>
               """, unsafe_allow_html=True)
    
    # Main content area
    col1, col2 = st.columns([1, 1])

    with col1:
        st.markdown("""
            <div style="margin-bottom: 2rem;">
                <h2 style="color: #00d4ff; border-bottom: 2px solid #00d4ff; padding-bottom: 0.5rem;">🎤 Upload Audio</h2>
            </div>
        """, unsafe_allow_html=True)

        # File uploader
        uploaded_file = st.file_uploader(
            "Choose an audio file",
            type=['mp3', 'wav', 'm4a', 'flac'],
            help="Upload your client meeting recording"
        )

        # Recipient email input
        recipient_email = st.text_input(
            "Recipient Email",
            placeholder="client@example.com",
            help="Email address to send the follow-up to"
        )

        # Optional recipient name
        recipient_name = st.text_input(
            "Client Name (optional)",
            placeholder="e.g., John Smith",
            help="Client name for personalized greeting"
        )

        # Generate button
        generate_button = st.button(
            "🔄 Generate Email",
            type="primary",
            disabled=not uploaded_file or not recipient_email
        )

    with col2:
        st.markdown("""
            <div style="margin-bottom: 2rem;">
                <h2 style="color: #00d4ff; border-bottom: 2px solid #00d4ff; padding-bottom: 0.5rem;">📄 Results</h2>
            </div>
        """, unsafe_allow_html=True)

        if generate_button and uploaded_file and recipient_email:
            with st.spinner("Processing audio and generating email..."):
                # Save uploaded file
                audio_filename = f"{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}_{uploaded_file.name}"
                audio_path = AUDIO_DIR / audio_filename

                with open(audio_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())

                # Transcribe audio
                st.info("🎧 Transcribing audio...")
                transcript = st.session_state.generator.transcribe_audio(audio_path)

                if transcript:
                    st.success("✅ Transcription completed!")

                    # Generate email
                    st.info("📝 Generating professional email...")
                    email_body = st.session_state.generator.generate_email(transcript, recipient_name)

                    if email_body:
                        st.success("✅ Email generated!")

                        # Store results in session state
                        st.session_state.current_transcript = transcript
                        st.session_state.current_email = email_body
                        st.session_state.current_recipient = recipient_email
                        st.session_state.current_recipient_name = recipient_name
                        st.session_state.current_audio_filename = audio_filename
                        st.session_state.current_audio_path = audio_path

    # Display results if available
    if hasattr(st.session_state, 'current_transcript'):
        st.markdown("---")

        # Create tabs for transcript and email
        tab1, tab2 = st.tabs(["📝 Generated Email", "📄 Full Transcript"])

        with tab1:
            st.markdown("""
                <div style="text-align: center; margin-bottom: 2rem;">
                    <h2 style="color: #00d4ff;">📧 Generated Email</h2>
                    <p style="color: #cccccc;">Professional follow-up ready to send</p>
                </div>
            """, unsafe_allow_html=True)

            # Display email in a styled container
            st.markdown(f"""
                <div class="email-container">
                    <div style="background: rgba(0, 212, 255, 0.1); border-left: 4px solid #00d4ff; padding: 1.5rem; border-radius: 0 10px 10px 0; margin-bottom: 1rem;">
                        <h4 style="color: #00d4ff; margin: 0 0 1rem 0;">📩 Email Preview</h4>
                        <pre style="white-space: pre-wrap; font-family: 'Arial', sans-serif; line-height: 1.6; color: #ffffff; background: transparent; border: none; padding: 0; margin: 0; font-size: 0.95rem;">{st.session_state.current_email}</pre>
                    </div>
                </div>
            """, unsafe_allow_html=True)

            # Send email section
            st.markdown("---")
            col1, col2, col3 = st.columns([1, 1, 1])

            with col1:
                if st.button("📧 Send Email", type="primary"):
                    subject = "Follow-Up from Our Recent Meeting"
                    with st.spinner("Sending email..."):
                        success = st.session_state.generator.send_email(
                            st.session_state.current_recipient,
                            subject,
                            st.session_state.current_email,
                            st.session_state.current_transcript
                        )
                    if success:
                        st.success(f"✅ Email sent successfully to {st.session_state.current_recipient}!")
                        # Save record
                        record_path = st.session_state.generator.save_email_record(
                            st.session_state.current_recipient,
                            subject,
                            st.session_state.current_email,
                            st.session_state.current_transcript,
                            st.session_state.current_audio_filename
                        )
                        st.info(f"📁 Email record saved: {record_path.name}")
                        # Clean up audio file
                        if st.session_state.current_audio_path.exists():
                            os.remove(st.session_state.current_audio_path)
                            st.info("🗑️ Audio file cleaned up")
with col2:
            if st.button("🚀 Add Tasks to Excel", type="secondary"):
                client_name = st.session_state.get('current_recipient_name', '') or st.session_state.current_recipient.split('@')[0]
                
                # Extract tasks (same logic as before)
                email_text = st.session_state.current_email
                extracted_tasks = []
                lines = email_text.split('\n')
                in_next_steps = False
                
                for line in lines:
                    line_clean = line.strip()
                    if 'next steps:' in line_clean.lower() or 'action items:' in line_clean.lower():
                        in_next_steps = True
                        continue
                    if (line_clean.lower().startswith(('warm regards', 'all the best', 'sincerely', 'should you have')) and in_next_steps):
                        break
                    if in_next_steps and line_clean:
                        if (line_clean.startswith(('○', '•', '-', '*', '1.', '2.', '3.', '4.', '5.', '6.', '7.', '8.', '9.')) and len(line_clean) > 3):
                            task = line_clean
                            for prefix in ['○ ', '• ', '- ', '* ', '1. ', '2. ', '3. ', '4. ', '5. ', '6. ', '7. ', '8. ', '9. ']:
                                if task.startswith(prefix):
                                    task = task[len(prefix):].strip()
                                    break
                            if task and len(task) > 5:
                                extracted_tasks.append(task)
                
                if extracted_tasks:
                    excel_manager = ExcelOnlineManager()
                    if excel_manager.add_tasks_to_excel(client_name, extracted_tasks):
                        st.success(f"✅ Added {len(extracted_tasks)} tasks to Excel!")
                        with st.expander("📋 Tasks Added"):
                            for i, task in enumerate(extracted_tasks, 1):
                                st.write(f"{i}. **{client_name}:** {task}")
                    else:
                        st.error("❌ Failed to add tasks")
                else:
                    st.warning("⚠️ No tasks found")
                        
            # Direct Excel access button
            excel_file_id = os.getenv('EXCEL_FILE_ID')
            excel_url = f"https://office.live.com/start/Excel.aspx?omkt=en-US&ui=en-US&rs=US&WOPISrc=https%3A//graph.microsoft.com/v1.0/me/drive/items/{excel_file_id}"
            st.markdown(f"""
            <a href="{excel_url}" target="_blank">
                <button style="
                    background: linear-gradient(135deg, #00d4ff 0%, #0099cc 100%);
                    color: #000000;
                    font-weight: 700;
                    border: none;
                    border-radius: 15px;
                    padding: 1rem 2rem;
                    font-size: 1rem;
                    width: 100%;
                    cursor: pointer;
                    text-transform: uppercase;
                    letter-spacing: 1px;
                    margin-top: 0.5rem;
                ">
                    📊 Open Task Master Excel
                </button>
            </a>
            """, unsafe_allow_html=True)

        with col3:
            # Download email as text file
            email_text = f"Subject: Follow-Up from Our Recent Meeting\n\n{st.session_state.current_email}"
            st.download_button(
                "💾 Download Email",
                email_text,
                file_name=f"email_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
                mime="text/plain"
            )


        with tab2:
            st.markdown("""
                <div style="text-align: center; margin-bottom: 2rem;">
                    <h2 style="color: #00d4ff;">📄 Meeting Transcript</h2>
                    <p style="color: #cccccc;">Full audio transcription</p>
                </div>
            """, unsafe_allow_html=True)

            # Create a styled transcript container
            st.markdown(f"""
                <div style="
                    background: rgba(0, 0, 0, 0.8);
                    color: #ffffff;
                    padding: 1.5rem;
                    border-radius: 15px;
                    border: 2px solid #333;
                    font-family: 'Courier New', monospace;
                    line-height: 1.6;
                    max-height: 400px;
                    overflow-y: auto;
                    white-space: pre-wrap;
                    box-shadow: inset 0 2px 5px rgba(0, 0, 0, 0.3);
                ">
{st.session_state.current_transcript}
                </div>
            """, unsafe_allow_html=True)

            # Download transcript
            st.download_button(
                "💾 Download Transcript",
                st.session_state.current_transcript,
                file_name=f"transcript_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
                mime="text/plain"
            )

    # Footer with new styling
    st.markdown("""
        <div class="footer">
            <div style="text-align: center;">
                <h3 style="color: #00d4ff; margin-bottom: 1rem;">🤖 Ewing Morris AI Assistant</h3>
                <p style="color: #cccccc; margin-bottom: 0.5rem;">Enterprise-grade • Secure • Intelligent</p>
                <div style="display: flex; justify-content: center; gap: 2rem; margin: 1rem 0;">
                    <span style="color: #666;">🔒 Encrypted</span>
                    <span style="color: #666;">⚡ Real-time</span>
                    <span style="color: #666;">🎯 Professional</span>
                </div>
                <p style="color: #444; font-size: 0.85rem;">Powered by OpenAI • Built for Investment Professionals</p>
            </div>
        </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
