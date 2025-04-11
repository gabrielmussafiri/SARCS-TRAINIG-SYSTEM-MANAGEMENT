import streamlit as st
import pandas as pd
import qrcode
import os
import base64
from io import BytesIO
from google.oauth2 import service_account
from googleapiclient.discovery import build
from PIL import Image
import time
import json
import re
from datetime import datetime
import PyPDF2
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
import tempfile
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

from dotenv import load_dotenv , dotenv_values
import google.auth

load_dotenv()

# Page configuration
st.set_page_config(
    page_title="Red Cross Trainers Management",
    page_icon="⛑️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #e63946;
    }
    .subheader {
        font-size: 1.5rem;
        color: #457b9d;
    }
    .stButton>button {
        background-color: #e63946;
        color: white;
    }
    .stButton>button:hover {
        background-color: #d62828;
        color: white;
    }
    .warning {
        color: #d62828;
        font-weight: bold;
    }
    .success {
        color: #2a9d8f;
        font-weight: bold;
    }
    .sidebar .sidebar-content {
        background-color: #f1faee;
    }
    .finished {
        background-color: #d8f3dc;
    }
    .not-finished {
        background-color: #ffccd5;
    }
    .status-badge {
        padding: 3px 10px;
        border-radius: 10px;
        font-weight: bold;
    }
    .status-finished {
        background-color: #2a9d8f;
        color: white;
    }
    .status-pending {
        background-color: #e9c46a;
        color: white;
    }
    .action-button {
        border: none;
        background: none;
        color: #457b9d;
        cursor: pointer;
        margin: 0 5px;
    }
    .action-button:hover {
        color: #1d3557;
    }
    .delete-button {
        color: #e63946;
    }
    .delete-button:hover {
        color: #d62828;
    }
    .edit-button {
        color: #2a9d8f;
    }
    .edit-button:hover {
        color: #1d3557;
    }
    .trainer-row {
        cursor: pointer;
    }
    .trainer-row:hover {
        background-color: #f8f9fa;
    }
    .stDataFrame {
        border: 1px solid #e6e6e6;
        border-radius: 5px;
    }
    /* Custom styling for the action buttons in the table */
    .stDataFrame [data-testid="stHorizontalBlock"] {
        gap: 5px !important;
    }
    /* Modal-like container for edit form */
    .edit-container {
        border: 1px solid #e6e6e6;
        border-radius: 10px;
        padding: 20px;
        background-color: white;
        box_shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        margin-bottom: 20px;
    }
    
    /* Enhanced sidebar styling */
    .sidebar-header {
        margin-top: 20px;
        margin-bottom: 20px;
        text-align: center;
        color: #e63946;
    }
    .sidebar-nav {
        margin-top: 30px;
        margin-bottom: 30px;
    }
    .nav-item {
        display: flex;
        align-items: center;
        padding: 10px;
        margin-bottom: 10px;
        border-radius: 5px;
        transition: background-color 0.3s;
    }
    .nav-item:hover {
        background-color: #f8f9fa;
    }
    .nav-item-active {
        background-color: #e63946;
        color: white;
    }
    .nav-icon {
        margin-right: 10px;
        font-size: 1.2rem;
    }
    .sidebar-footer {
        position: absolute;
        bottom: 20px;
        width: 100%;
        text-align: center;
        font-size: 0.8rem;
        color: #6c757d;
    }
    .sidebar-divider {
        margin-top: 20px;
        margin-bottom: 20px;
        border-top: 1px solid #e6e6e6;
    }
    .sidebar-button {
        width: 100%;
        margin-top: 10px;
        margin-bottom: 10px;
    }
    /* Add this CSS for the certificate download button (in the custom CSS section) */
    .certificate-download {
        display: inline-block;
        background-color: #e63946;
        color: white;
        padding: 10px 20px;
        text-decoration: none;
        border-radius: 5px;
        font-weight: bold;
        margin-top: 10px;
        text-align: center;
    }
    .certificate-download:hover {
        background-color: #d62828;
    }
    .certificate-container {
        margin-top: 20px;
        padding: 20px;
        border: 1px solid #e6e6e6;
        border-radius: 10px;
        background-color: #f8f9fa;
    }
    /* New styles for sheet management */
    .sheet-manager {
        background-color: #f8f9fa;
        padding: 20px;
        border-radius: 10px;
        margin-bottom: 20px;
    }
    .sheet-card {
        background-color: white;
        padding: 15px;
        border-radius: 8px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        margin-bottom: 10px;
        cursor: pointer;
        transition: transform 0.2s;
    }
    .sheet-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 8px rgba(0,0,0,0.15);
    }
    .sheet-card.active {
        border-left: 4px solid #e63946;
    }
    .sheet-title {
        font-weight: bold;
        color: #1d3557;
    }
    .sheet-info {
        color: #6c757d;
        font-size: 0.9rem;
    }
    .sheet-count {
        background-color: #e9ecef;
        padding: 2px 8px;
        border-radius: 10px;
        font-size: 0.8rem;
        color: #495057;
    }
</style>
""", unsafe_allow_html=True)

# Initialize session state for caching and UI state
if 'finished_data' not in st.session_state:
    st.session_state.finished_data = None
if 'all_trainers_data' not in st.session_state:
    st.session_state.all_trainers_data = None
if 'last_refresh' not in st.session_state:
    st.session_state.last_refresh = 0
if 'qr_image' not in st.session_state:
    st.session_state.qr_image = None
if 'form_data' not in st.session_state:
    st.session_state.form_data = {}
if 'selected_sheet' not in st.session_state:
    st.session_state.selected_sheet = "NAMESTHATFINISHED"
if 'editing' not in st.session_state:
    st.session_state.editing = False
if 'edit_index' not in st.session_state:
    st.session_state.edit_index = -1
if 'edit_data' not in st.session_state:
    st.session_state.edit_data = {}
if 'show_delete_confirm' not in st.session_state:
    st.session_state.show_delete_confirm = False
if 'delete_index' not in st.session_state:
    st.session_state.delete_index = -1
if 'edit_data' not in st.session_state:
    st.session_state.edit_data = {}
if 'show_delete_confirm' not in st.session_state:
    st.session_state.show_delete_confirm = False
if 'delete_index' not in st.session_state:
    st.session_state.delete_index = -1
if 'sheet_metadata' not in st.session_state:
    st.session_state.sheet_metadata = {}
if 'show_create_sheet' not in st.session_state:
    st.session_state.show_create_sheet = False
if 'available_sheets' not in st.session_state:
    st.session_state.available_sheets = []
if 'sheet_data' not in st.session_state:
    st.session_state.sheet_data = {}

# Constants
# SPREADSHEET_ID = config["SPREADSHEET_ID"]
config = dotenv_values(".env")

# Constants
# Try to get SPREADSHEET_ID from config or environment variables

if "SPREADSHEET_ID" in config:
    SPREADSHEET_ID = config["SPREADSHEET_ID"]
elif "SPREADSHEET_ID" in os.environ:
    SPREADSHEET_ID = os.environ["SPREADSHEET_ID"]
else:
    # If not found, show an error
    st.error("SPREADSHEET_ID not found in configuration or environment variables")
    st.stop()
FINISHED_SHEET = "NAMESTHATFINISHED"
METADATA_SHEET = "SHEET_METADATA"  # New sheet to store metadata about other sheets

# Helper function to format date
def format_date_to_dd_mm_yyyy(date_value):
    """Format a date to dd/mm/yyyy format."""
    if not date_value:
        return ""
        
    try:
        if isinstance(date_value, str):
            # Try to parse the date string
            date_obj = datetime.strptime(date_value, '%Y-%m-%d')
            return date_obj.strftime('%d/%m/%Y')
        elif hasattr(date_value, 'strftime'):
            # If it's already a date object
            return date_value.strftime('%d/%m/%Y')
        else:
            return str(date_value)
    except Exception:
        # Return original if parsing fails
        return str(date_value)

def sanitize_sheet_name(name):
    """Sanitize sheet name to avoid issues with special characters"""
    # Replace problematic characters with underscores
    sanitized = re.sub(r'[&\s\+\-$$$$\[\]\{\}\.\,\;\:\'\"\!\@\#\$\%\^\*\=\<\>\?\/\\\|]', '_', name)
    return sanitized


# Function to generate QR code data
def generate_qr_code_data(full_name, id_number, gender, cert_type, cert_no, issue_date):
    """Generate consistent QR code data string."""
    formatted_date = format_date_to_dd_mm_yyyy(issue_date)
     # Convert list to clean string if necessary
    if isinstance(cert_type, list):
        cert_type = ", ".join(cert_type)
    
    return (f"Full Name: {full_name} -- ID No: {id_number} -- Gender: {gender} -- "
            f"Course: {cert_type} -- Cert. No: {cert_no} -- Issue Date: {formatted_date} -- "
            f"Issued by the South African Red Cross Society on the Bearer's Fulfillment of all Requirements. "
            f"For Further Enquiries concerning {full_name} Contact: trainingadmin.wc@redcross.org.za")

# Google Sheets API Setup
@st.cache_resource
def get_google_sheets_credentials():
    """Get Google Sheets API credentials, preferring st.secrets but falling back to keys.json."""

    try:
        # 1. Check if secrets file exists (safe for local dev)
        secrets_path = os.path.join(os.getcwd(), ".streamlit", "secrets.toml")
        if os.path.exists(secrets_path):
            try:
                creds_json = json.loads(st.secrets["google_credentials"])
                creds = service_account.Credentials.from_service_account_info(
                    creds_json,
                    scopes=["https://www.googleapis.com/auth/spreadsheets"]
                )
                return creds
            except Exception as e:
                st.warning(f"Found secrets.toml but failed to load: {e}")

        # 2. Fallback to local keys.json
        if os.path.exists("keys.json"):
            creds = service_account.Credentials.from_service_account_file(
                "keys.json",
                scopes=["https://www.googleapis.com/auth/spreadsheets"]
            )
            return creds

        # 3. No credentials found
        st.error("No credentials found. Please add keys.json or use .streamlit/secrets.toml")
        return None

    except Exception as e:
        st.error(f"Error loading credentials: {e}")
        return None

    """Get Google Sheets API credentials."""
    try:
        # 1. Use secrets if running on Streamlit Cloud
        if "google_credentials" in st.secrets:
            creds_json = json.loads(st.secrets["google_credentials"])
            creds = service_account.Credentials.from_service_account_info(
                creds_json,
                scopes=["https://www.googleapis.com/auth/spreadsheets"]
            )
            return creds

        # 2. Use local keys.json for development
        if os.path.exists("keys.json"):
            creds = service_account.Credentials.from_service_account_file(
                "keys.json",
                scopes=["https://www.googleapis.com/auth/spreadsheets"]
            )
            return creds

        st.error("No credentials found. Please provide either st.secrets['google_credentials'] or keys.json.")
        return None

    except Exception as e:
        st.error(f"Error loading credentials: {e}")
        return None

    """Securely load Google Sheets API credentials."""
    try:
        # 1. Load from Streamlit secrets (recommended for deployment)
        if "google_credentials" in st.secrets:
            creds_json = json.loads(st.secrets["google_credentials"])
            creds = service_account.Credentials.from_service_account_info(
                creds_json,
                scopes=["https://www.googleapis.com/auth/spreadsheets"]
            )
            return creds

        # 2. (Optional) Fallback to local keys.json for local dev
        if os.path.exists('keys.json'):
            creds = service_account.Credentials.from_service_account_file(
                'keys.json',
                scopes=["https://www.googleapis.com/auth/spreadsheets"]
            )
            return creds

        st.error("Google credentials not found in secrets or keys.json")
        return None

    except Exception as e:
        st.error(f"Error loading credentials: {e}")
        return None

creds = get_google_sheets_credentials()

if creds is None:
    st.error("Failed to load Google API credentials. Please check your configuration.")
    st.stop()

# Build the service
try:
    service = build("sheets", "v4", credentials=creds)
    sheet = service.spreadsheets()
    st.success("Successfully connected to Google Sheets API")
except Exception as e:
    st.error(f"Error connecting to Google Sheets API: {e}")
    import traceback
    st.error(traceback.format_exc())
    st.stop()

# Function to get all sheets in the spreadsheet
def get_all_sheets():
    try:
        sheet_metadata = sheet.get(spreadsheetId=SPREADSHEET_ID).execute()
        sheets = sheet_metadata.get('sheets', '')
        sheet_names = [s['properties']['title'] for s in sheets]
        return sheet_names
    except Exception as e:
        st.error(f"Error getting sheets: {e}")
        return []

# Function to create the metadata sheet if it doesn't exist
def create_metadata_sheet_if_not_exists():
    try:
        # Check if the metadata sheet exists
        all_sheets = get_all_sheets()
        if METADATA_SHEET not in all_sheets:
            # Create the metadata sheet
            request_body = {
                'requests': [{
                    'addSheet': {
                        'properties': {
                            'title': METADATA_SHEET
                        }
                    }
                }]
            }
            
            sheet.batchUpdate(
                spreadsheetId=SPREADSHEET_ID,
                body=request_body
            ).execute()
            
            # Add headers to the metadata sheet
            headers = [
                "Sheet Name", "Display Name", "Certificate Format", "Last Certificate Number", "Creation Date"
            ]
            
            sheet.values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=f"{METADATA_SHEET}!A1:E1",
                valueInputOption="USER_ENTERED",
                body={"values": [headers]}
            ).execute()
            
            st.success(f"Created metadata sheet: {METADATA_SHEET}")
            return True
        
        return True
    except Exception as e:
        st.error(f"Error creating metadata sheet: {e}")
        return False

# Function to load sheet metadata
def load_sheet_metadata(force_refresh=False):
    if not st.session_state.sheet_metadata or force_refresh:
        try:
            # Ensure metadata sheet exists
            create_metadata_sheet_if_not_exists()
            
            # Get metadata
            result = sheet.values().get(
                spreadsheetId=SPREADSHEET_ID, 
                range=f"{METADATA_SHEET}!A2:E"
            ).execute()
            
            values = result.get("values", [])
            
            # Convert to dictionary
            metadata = {}
            for row in values:
                if len(row) >= 3:  # Ensure we have at least sheet name and certificate format
                    sheet_name = row[0]
                    metadata[sheet_name] = {
                        "display_name": row[1] if len(row) > 1 else sheet_name,
                        "cert_format": row[2] if len(row) > 2 else "",
                        "last_cert_number": row[3] if len(row) > 3 else "0",
                        "creation_date": row[4] if len(row) > 4 else datetime.now().strftime('%Y-%m-%d')
                    }
            
            st.session_state.sheet_metadata = metadata
            
            # Update available sheets
            all_sheets = get_all_sheets()
            # Filter out system sheets
            available_sheets = [s for s in all_sheets if s != METADATA_SHEET and s != FINISHED_SHEET]
            st.session_state.available_sheets = available_sheets
            
            # Initialize sheet_data for each available sheet
            if 'sheet_data' not in st.session_state:
                st.session_state.sheet_data = {}
            
            return metadata
        except Exception as e:
            st.error(f"Error loading sheet metadata: {e}")
            return {}
    
    return st.session_state.sheet_metadata

# Function to create a new sheet
def create_new_sheet(sheet_name, display_name, cert_format):
    try:
        # Sanitize the sheet name
        sheet_name = sanitize_sheet_name(sheet_name)
        
        # Check if the sheet already exists
        all_sheets = get_all_sheets()
        if sheet_name in all_sheets:
            st.error(f"Sheet '{sheet_name}' already exists")
            return False
        
        # Create the new sheet
        request_body = {
            'requests': [{
                'addSheet': {
                    'properties': {
                        'title': sheet_name
                    }
                }
            }]
        }
        
        sheet.batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body=request_body
        ).execute()
        
        # Add headers to the new sheet
        headers = [
            "Name(s)", "Surname", "Full Name", "ID number", "Certificate No.", 
            "Issue Date", "Contact No.", "Province", "Branch", 
            "Type of Certification", "Gender", "Home Address", "Nationality", 
            "QR-Picture", "QR Code", "Candidate Full Description", "Training Status"
        ]
        
        sheet.values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{sheet_name}!A1:Q1",
            valueInputOption="USER_ENTERED",
            body={"values": [headers]}
        ).execute()
        
        # Add metadata for the new sheet
        metadata = load_sheet_metadata()
        
        # Append to metadata sheet
        new_metadata = [
            sheet_name, 
            display_name, 
            cert_format, 
            "0",  # Last certificate number
            datetime.now().strftime('%Y-%m-%d')  # Creation date
        ]
        
        sheet.values().append(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{METADATA_SHEET}!A1",
            valueInputOption="USER_ENTERED",
            body={"values": [new_metadata]}
        ).execute()
        
        # Update local metadata
        metadata[sheet_name] = {
            "display_name": display_name,
            "cert_format": cert_format,
            "last_cert_number": "0",
            "creation_date": datetime.now().strftime('%Y-%m-%d')
        }
        
        st.session_state.sheet_metadata = metadata
        
        # Update available sheets
        all_sheets = get_all_sheets()
        available_sheets = [s for s in all_sheets if s != METADATA_SHEET and s != FINISHED_SHEET]
        st.session_state.available_sheets = available_sheets
        
        # Initialize the sheet data cache
        if 'sheet_data' not in st.session_state:
            st.session_state.sheet_data = {}
        st.session_state.sheet_data[sheet_name] = pd.DataFrame(columns=headers)
        
        st.success(f"Created new sheet: {display_name}")
        return True
    except Exception as e:
        st.error(f"Error creating sheet: {e}")
        return False

# Function to update sheet metadata
def update_sheet_metadata(sheet_name, field, value):
    try:
        metadata = load_sheet_metadata()
        
        if sheet_name not in metadata:
            st.error(f"Sheet '{sheet_name}' not found in metadata")
            return False
        
        # Update the field
        metadata[sheet_name][field] = value
        
        # Find the row in the metadata sheet
        result = sheet.values().get(
            spreadsheetId=SPREADSHEET_ID, 
            range=f"{METADATA_SHEET}!A2:A"
        ).execute()
        
        values = result.get("values", [])
        row_index = None
        
        for i, row in enumerate(values):
            if row[0] == sheet_name:
                row_index = i + 2  # +2 because we start at A2
                break
        
        if row_index is None:
            st.error(f"Sheet '{sheet_name}' not found in metadata sheet")
            return False
        
        # Update the specific field
        if field == "display_name":
            col = "B"
        elif field == "cert_format":
            col = "C"
        elif field == "last_cert_number":
            col = "D"
        elif field == "creation_date":
            col = "E"
        else:
            st.error(f"Invalid metadata field: {field}")
            return False
        
        sheet.values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{METADATA_SHEET}!{col}{row_index}",
            valueInputOption="USER_ENTERED",
            body={"values": [[value]]}
        ).execute()
        
        # Update local metadata
        st.session_state.sheet_metadata = metadata
        
        return True
    except Exception as e:
        st.error(f"Error updating sheet metadata: {e}")
        return False

# Function to get the next certificate number for a sheet
def get_next_certificate_number(sheet_name):
    try:
        metadata = load_sheet_metadata()
        
        if sheet_name not in metadata:
            st.error(f"Sheet '{sheet_name}' not found in metadata")
            return ""
        
        cert_format = metadata[sheet_name]["cert_format"]
        last_number = metadata[sheet_name]["last_cert_number"]
        
        # Check if the format contains any placeholders
        if "{number}" not in cert_format and "{####}" not in cert_format and "{###}" not in cert_format and "{##}" not in cert_format:
            # No placeholders - use the format as-is with an incrementing number at the end
            next_number = int(last_number) + 1
            new_cert_number = f"{cert_format}{next_number}"
        else:
            # Find the numeric part in the format
            # The format should contain a placeholder like {number} or {####}
            if "{number}" in cert_format:
                # Simple incrementing number
                next_number = int(last_number) + 1
                new_cert_number = cert_format.replace("{number}", str(next_number))
            elif "{####}" in cert_format:
                # 4-digit number with leading zeros
                prefix = cert_format.split("{####}")[0]  # Get the prefix (non-numeric part)
                numeric_part = int(last_number.split(prefix)[-1])  # Extract the numeric part
                next_number = numeric_part + 1
                new_cert_number = f"{prefix}{next_number:04d}"  # Keep leading zeros
            elif "{###}" in cert_format:
                # 3-digit number with leading zeros
                prefix = cert_format.split("{###}")[0]  # Get the prefix (non-numeric part)
                numeric_part = int(last_number.split(prefix)[-1])  # Extract the numeric part
                next_number = numeric_part + 1
                new_cert_number = f"{prefix}{next_number:03d}"  # Keep leading zeros
            elif "{##}" in cert_format:
                # 2-digit number with leading zeros
                prefix = cert_format.split("{##}")[0]  # Get the prefix (non-numeric part)
                numeric_part = int(last_number.split(prefix)[-1])  # Extract the numeric part
                next_number = numeric_part + 1
                new_cert_number = f"{prefix}{next_number:02d}"  # Keep leading zeros
        
        # Update the last certificate number in metadata
        update_sheet_metadata(sheet_name, "last_cert_number", str(next_number))
        
        return new_cert_number
    except Exception as e:
        st.error(f"Error generating certificate number: {e}")
        return ""

    try:
        metadata = load_sheet_metadata()
        
        if sheet_name not in metadata:
            st.error(f"Sheet '{sheet_name}' not found in metadata")
            return ""
        
        cert_format = metadata[sheet_name]["cert_format"]
        last_number = metadata[sheet_name]["last_cert_number"]
        
        # Check if the format contains any placeholders
        if "{number}" not in cert_format and "{####}" not in cert_format and "{###}" not in cert_format and "{##}" not in cert_format:
            # No placeholders - use the format as-is with an incrementing number at the end
            next_number = int(last_number) + 1
            new_cert_number = f"{cert_format}{next_number}"
        else:
            # Find the numeric part in the format
            # The format should contain a placeholder like {number} or {####}
            if "{number}" in cert_format:
                # Simple incrementing number
                next_number = int(last_number) + 1
                new_cert_number = cert_format.replace("{number}", str(next_number))
            elif "{####}" in cert_format:
                # 4-digit number with leading zeros
                next_number = int(last_number) + 1
                new_cert_number = cert_format.replace("{####}", f"{next_number:04d}")
            elif "{###}" in cert_format:
                # 3-digit number with leading zeros
                next_number = int(last_number) + 1
                new_cert_number = cert_format.replace("{###}", f"{next_number:03d}")
            elif "{##}" in cert_format:
                # 2-digit number with leading zeros
                next_number = int(last_number) + 1
                new_cert_number = cert_format.replace("{##}", f"{next_number:02d}")
        
        # Update the last certificate number in metadata
        update_sheet_metadata(sheet_name, "last_cert_number", str(next_number))
        
        return new_cert_number
    except Exception as e:
        st.error(f"Error generating certificate number: {e}")
        return ""

# Add this function after the refresh_all_data() function to specifically handle the HOME_BASE_LEVEL2&3 sheet

def force_refresh_specific_sheet(sheet_name):
    """Force refresh data for a specific sheet and ensure it's properly loaded"""
    try:
        st.info(f"Refreshing data for {sheet_name}...")
        
        # Direct API call to get the data
        result = sheet.values().get(
            spreadsheetId=SPREADSHEET_ID, 
            range=f"{sheet_name}!A1:Q"
        ).execute()
        
        values = result.get("values", [])
        
        if not values:
            st.warning(f"No data found in {sheet_name}")
            return pd.DataFrame()
        
        # Ensure all rows have the same length
        max_cols = 17  # A to Q
        values = [row + [""] * (max_cols - len(row)) for row in values]
        
        # Create DataFrame
        df = pd.DataFrame(values[1:], columns=values[0])
        
        # Update the cache
        if 'sheet_data' not in st.session_state:
            st.session_state.sheet_data = {}
        
        st.session_state.sheet_data[sheet_name] = df
        st.session_state.last_refresh = time.time()
        
        st.success(f"Successfully refreshed {sheet_name} with {len(df)} records")
        return df
    except Exception as e:
        st.error(f"Error refreshing {sheet_name}: {e}")
        return pd.DataFrame()

# Modify the get_data function to better handle special sheet names
def get_data(sheet_name, force_refresh=False):
    """
    Fetch data from a Google Sheet with caching.
    """
    # Special handling for HOME_BASE_LEVEL2&3 sheet
    if sheet_name == "HOME_BASE_LEVEL2&3" and (force_refresh or 'HOME_BASE_LEVEL2&3' not in st.session_state.get('sheet_data', {})):
        return force_refresh_specific_sheet(sheet_name)
    
    # Determine which session state variable to use
    if sheet_name == FINISHED_SHEET:
        data_key = 'finished_data'
    else:
        # Use a dictionary to store data for different sheets
        if 'sheet_data' not in st.session_state:
            st.session_state.sheet_data = {}
        data_key = 'sheet_data'
    
    # Check if we need to refresh the data
    current_time = time.time()
    should_refresh = (
        (sheet_name == FINISHED_SHEET and (st.session_state[data_key] is None or force_refresh)) or
        (sheet_name != FINISHED_SHEET and (sheet_name not in st.session_state[data_key] or force_refresh)) or
        (current_time - st.session_state.last_refresh) > 300  # Refresh every 5 minutes
    )
    
    if should_refresh:
        with st.spinner(f"Loading data from {sheet_name}..."):
            try:
                result = sheet.values().get(
                    spreadsheetId=SPREADSHEET_ID, 
                    range=f"{sheet_name}!A1:Q"
                ).execute()
                
                values = result.get("values", [])
                
                if not values:
                    empty_df = pd.DataFrame(columns=[
                        "Name(s)", "Surname", "Full Name", "ID number", "Certificate No.", 
                        "Issue Date", "Contact No.", "Province", "Branch", 
                        "Type of Certification", "Gender", "Home Address", "Nationality", 
                        "QR-Picture", "QR Code", "Candidate Full Description", "Training Status"
                    ])
                    
                    if sheet_name == FINISHED_SHEET:
                        st.session_state[data_key] = empty_df
                    else:
                        st.session_state[data_key][sheet_name] = empty_df
                else:
                    # Ensure all rows have the same length
                    max_cols = 17  # A to Q
                    values = [row + [""] * (max_cols - len(row)) for row in values]
                    df = pd.DataFrame(values[1:], columns=values[0])
                    
                    if sheet_name == FINISHED_SHEET:
                        st.session_state[data_key] = df
                    else:
                        st.session_state[data_key][sheet_name] = df
                
                st.session_state.last_refresh = current_time
            except Exception as e:
                st.error(f"Error fetching data from {sheet_name}: {e}")
                if sheet_name == FINISHED_SHEET and st.session_state[data_key] is None:
                    st.session_state[data_key] = pd.DataFrame()
                elif sheet_name != FINISHED_SHEET and sheet_name not in st.session_state[data_key]:
                    st.session_state[data_key][sheet_name] = pd.DataFrame()
    
    # Return the data
    if sheet_name == FINISHED_SHEET:
        return st.session_state[data_key]
    else:
        if sheet_name not in st.session_state[data_key]:
            # Initialize with empty DataFrame if not exists
            st.session_state[data_key][sheet_name] = pd.DataFrame(columns=[
                "Name(s)", "Surname", "Full Name", "ID number", "Certificate No.", 
                "Issue Date", "Contact No.", "Province", "Branch", 
                "Type of Certification", "Gender", "Home Address", "Nationality", 
                "QR-Picture", "QR Code", "Candidate Full Description", "Training Status"
            ])
        return st.session_state[data_key][sheet_name]

# Function to Add Data to a specific sheet
def add_data(sheet_name, row):
    # Basic input validation
    if not isinstance(row, list):
        st.error("Invalid data format: row must be a list")
        return False

    # Sanitize inputs to prevent potential injection
    sanitized_row = []
    for item in row:
        # Convert lists (like ['FIRST AID LEVEL 1']) to strings
        if isinstance(item, list):
            item = ", ".join(map(str, item))
            
        if isinstance(item, str):
            # Remove any potentially harmful characters
            sanitized_item = item.replace('=', '').replace('+', '').replace('-', '')
            sanitized_row.append(sanitized_item)
        else:
            sanitized_row.append(item)
    row = sanitized_row
    try:
        # Ensure row has the right number of elements
        if sheet_name != FINISHED_SHEET and len(row) == 16:
            # Add Training Status column for non-FINISHED sheets
            row.append("Pending")
        
        # Ensure row has 17 elements (including Training Status)
        row = row + [""] * (17 - len(row))
        
        sheet.values().append(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{sheet_name}!A1",
            valueInputOption="USER_ENTERED",
            body={"values": [row]}
        ).execute()
        
        # Force refresh data to update the cache
        get_data(sheet_name, force_refresh=True)
        
        return True
    except Exception as e:
        st.error(f"Error adding data to {sheet_name}: {e}")
        return False

# Function to Update Data in a specific sheet
def update_data(sheet_name, index, updated_row):
    """
    Update data in a Google Sheet.
    """
    try:
        data = get_data(sheet_name)
        
        # Verify index is valid
        if index < 0 or index >= len(data):
            st.error("Invalid index selected")
            return False
        
        # Ensure row has the right number of elements
        if len(updated_row) < 17:  # Including Training Status
            updated_row = updated_row + [""] * (17 - len(updated_row))
        
        # Convert any date objects to strings
        for i, val in enumerate(updated_row):
            if isinstance(val, (pd.Timestamp, pd.DatetimeIndex, pd.DatetimeTZDtype)):
                updated_row[i] = val.strftime('%Y-%m-%d')
            elif hasattr(val, 'strftime'):  # For datetime.date objects
                updated_row[i] = val.strftime('%Y-%m-%d')
        
        # Update the data in the DataFrame
        for i, col in enumerate(data.columns):
            if i < len(updated_row):
                data.at[index, col] = updated_row[i]
        
        # Clear the sheet (except headers)
        sheet.values().clear(
            spreadsheetId=SPREADSHEET_ID, 
            range=f"{sheet_name}!A2:Q"
        ).execute()
        
        # Update with new data
        if not data.empty:
            # Convert DataFrame to list of lists for Google Sheets API
            values_to_update = []
            for _, row in data.iterrows():
                row_values = []
                for val in row:
                    # Convert any date objects to strings
                    if isinstance(val, (pd.Timestamp, pd.DatetimeIndex, pd.DatetimeTZDtype)):
                        row_values.append(val.strftime('%Y-%m-%d'))
                    elif hasattr(val, 'strftime'):  # For datetime.date objects
                        row_values.append(val.strftime('%Y-%m-%d'))
                    else:
                        row_values.append(val)
                values_to_update.append(row_values)
            
            sheet.values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=f"{sheet_name}!A2",
                valueInputOption="USER_ENTERED",
                body={"values": values_to_update}
            ).execute()
        
        # Update the cache
        if sheet_name == FINISHED_SHEET:
            st.session_state.finished_data = data
        elif 'sheet_data' in st.session_state:
            st.session_state.sheet_data[sheet_name] = data
        
        return True
    except Exception as e:
        st.error(f"Error updating data in {sheet_name}: {e}")
        return False

# Function to Delete Data from a specific sheet
def delete_data(sheet_name, index):
    try:
        data = get_data(sheet_name)
        
        # Verify index is valid
        if index < 0 or index >= len(data):
            st.error("Invalid index selected")
            return False
        
        # Create a copy to avoid modifying during iteration
        data_copy = data.copy()
        data_copy = data_copy.drop(index)
        
        # Clear the sheet (except headers)
        sheet.values().clear(
            spreadsheetId=SPREADSHEET_ID, 
            range=f"{sheet_name}!A2:Q"
        ).execute()
        
        # Update with new data
        if not data_copy.empty:
            sheet.values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=f"{sheet_name}!A2",
                valueInputOption="USER_ENTERED",
                body={"values": data_copy.values.tolist()}
            ).execute()
        
        # Update the cache
        if sheet_name == FINISHED_SHEET:
            st.session_state.finished_data = data_copy
        elif 'sheet_data' in st.session_state:
            st.session_state.sheet_data[sheet_name] = data_copy
        
        return True
    except Exception as e:
        st.error(f"Error deleting data from {sheet_name}: {e}")
        return False

# Function to mark a trainer as finished
def mark_trainer_as_finished(sheet_name, index):
    try:
        data = get_data(sheet_name)
        
        # Verify index is valid
        if index < 0 or index >= len(data):
            st.error("Invalid index selected")
            return False
        
        # Get the trainer's data
        trainer_data = data.iloc[index].tolist()
        
        # Update status in the source sheet
        data.at[index, "Training Status"] = "Finished"
        
        # Update the QR code data with the new format
        trainer_full_name = data.at[index, "Full Name"]
        trainer_id = data.at[index, "ID number"]
        trainer_gender = data.at[index, "Gender"]
        trainer_cert_type = data.at[index, "Type of Certification"]
        trainer_cert_no = data.at[index, "Certificate No."]
        trainer_issue_date = data.at[index, "Issue Date"]

        # Format the issue date to dd/mm/yyyy
        formatted_date = format_date_to_dd_mm_yyyy(trainer_issue_date)
            
        # Generate the QR code data
        new_qr_data = generate_qr_code_data(
            trainer_full_name, trainer_id, trainer_gender, 
            trainer_cert_type, trainer_cert_no, trainer_issue_date
        )
        data.at[index, "QR Code"] = new_qr_data
        
        # Clear the sheet (except headers)
        sheet.values().clear(
            spreadsheetId=SPREADSHEET_ID, 
            range=f"{sheet_name}!A2:Q"
        ).execute()
        
        # Update sheet with new data
        if not data.empty:
            sheet.values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=f"{sheet_name}!A2",
                valueInputOption="USER_ENTERED",
                body={"values": data.values.tolist()}
            ).execute()
        
        # Add the trainer to the FINISHED_SHEET
        # Remove the Training Status column for the FINISHED_SHEET
        finished_data = trainer_data[:-1]  # Remove the last element (Training Status)
        
        # Check if trainer already exists in FINISHED_SHEET
        finished_trainers = get_data(FINISHED_SHEET)
        id_number = trainer_data[3]  # ID number is at index 3
        
        # Check if this ID already exists in finished trainers
        if not finished_trainers.empty and 'ID number' in finished_trainers.columns:
            existing = finished_trainers[finished_trainers['ID number'] == id_number]
            if not existing.empty:
                # Trainer already exists in FINISHED_SHEET, update the record
                finished_index = existing.index[0]
                update_data(FINISHED_SHEET, finished_index, finished_data)
            else:
                # Add to FINISHED_SHEET
                add_data(FINISHED_SHEET, finished_data)
        else:
            add_data(FINISHED_SHEET, finished_data)
        
        # Force refresh data
        get_data(sheet_name, force_refresh=True)
        get_data(FINISHED_SHEET, force_refresh=True)
        
        # Update the sheet_data cache
        if 'sheet_data' in st.session_state and sheet_name in st.session_state.sheet_data:
            st.session_state.sheet_data[sheet_name] = data
        
        return True
    except Exception as e:
        st.error(f"Error marking trainer as finished: {e}")
        return False

# Function to check if Certificate Number exists in any sheet
def check_certificate_exists(certificate_no, exclude_sheet=None):
    try:
        # Check FINISHED_SHEET first
        finished_data = get_data(FINISHED_SHEET)
        if not finished_data.empty and 'Certificate No.' in finished_data.columns:
            if certificate_no in finished_data['Certificate No.'].values:
                return True, FINISHED_SHEET
        
        # Check all other sheets
        for sheet_name in st.session_state.available_sheets:
            if sheet_name == exclude_sheet or sheet_name == METADATA_SHEET:
                continue
                
            # Skip sheets that don't exist or can't be loaded
            try:
                sheet_data = get_data(sheet_name)
                if not sheet_data.empty and 'Certificate No.' in sheet_data.columns:
                    if id_number in sheet_data['Certificate No.'].values:
                        return True, sheet_name
            except Exception as e:
                st.warning(f"Could not check Certificate's in sheet {sheet_name}: {e}")
                continue
        
        return False, None
    except Exception as e:
        st.error(f"Error checking Certificate existence: {e}")
        return False, None

# Function to Generate QR Code
def generate_qr(data, box_size=10):
    """Generate a QR code from the provided data."""
    if not data:
        st.warning("No data provided for QR code generation")
        return None
        
    try:
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_L,
            box_size=box_size,
            border=4,
        )
        qr.add_data(data)
        qr.make(fit=True)
        
        img = qr.make_image(fill_color="black", back_color="white")
        
        # Create a BytesIO object for the QR code
        buf = BytesIO()
        img.save(buf, format="PNG")
        buf.seek(0)
        return buf
    except Exception as e:
        st.error(f"Error generating QR code: {e}")
        return None

# Function to get download link for QR code
def get_qr_download_link(qr_buf, filename="qrcode.png"):
    try:
        b64 = base64.b64encode(qr_buf.getvalue()).decode()
        href = f'<a href="data:image/png;base64,{b64}" download="{filename}">Download QR Code</a>'
        return href
    except Exception as e:
        st.error(f"Error creating download link: {e}")
        return ""

# Function to validate form data
def validate_form_data(form_data):
    required_fields = ["Name(s)", "Surname", "Full Name", "Certificate No."]
    errors = []
    
    for field in required_fields:
        if not form_data.get(field):
            errors.append(f"{field} is required")
    
    # Check if Certificate No already exists in any sheet
    certificate_no = form_data.get("Certificate No.")
    if certificate_no:
        exists, sheet_name = check_certificate_exists(certificate_no)
        if exists:
            errors.append(f"Certificate No. {certificate_no} already exists in sheet {sheet_name}")
    
    return errors

# Function to start editing a trainer
def start_editing(index, sheet_name):
    data = get_data(sheet_name)
    if index >= 0 and index < len(data):
        st.session_state.editing = True
        st.session_state.edit_index = index
        st.session_state.edit_data = data.iloc[index].to_dict()
        st.session_state.edit_sheet = sheet_name

# Function to cancel editing
def cancel_editing():
    st.session_state.editing = False
    st.session_state.edit_index = -1
    st.session_state.edit_data = {}

# Function to save edited trainer
def save_edited_trainer():
    if st.session_state.edit_index >= 0:
        sheet_name = st.session_state.edit_sheet
        index = st.session_state.edit_index
        
        # Get the current data to understand its structure
        data = get_data(sheet_name)
        
        # Check if ID has changed and if it already exists
        original_id = data.iloc[index]["ID number"]
        new_id = st.session_state.edit_data.get("ID number")
        
        if original_id != new_id:
            exists, existing_sheet = check_certificate_exists(new_id, exclude_sheet=sheet_name)
            if exists:
                st.error(f"ID number {new_id} already exists in sheet {existing_sheet}")
                return False
        
        # Initialize updated_row with the correct number of columns
        updated_row = [""] * len(data.columns)
        
        # Convert edit_data dict to a row list in the correct order
        for i, col in enumerate(data.columns):
            if col in st.session_state.edit_data:
                value = st.session_state.edit_data.get(col, "")
                # Convert date objects to strings
                if isinstance(value, (pd.Timestamp, pd.DatetimeIndex, pd.DatetimeTZDtype)):
                    value = value.strftime('%Y-%m-%d')
                elif hasattr(value, 'strftime'):  # For datetime.date objects
                    value = value.strftime('%Y-%m-%d')
                updated_row[i] = value
        
        # Update the data
        if update_data(sheet_name, index, updated_row):
            st.success("Trainer information updated successfully!")
            
            # If status changed to Finished, update FINISHED_SHEET too
            if st.session_state.edit_data.get("Training Status") == "Finished":
                # Check if we need to update or add to FINISHED_SHEET
                finished_data = get_data(FINISHED_SHEET)
                id_number = st.session_state.edit_data.get("ID number")
                
                if not finished_data.empty and 'ID number' in finished_data.columns:
                    # Check if this ID exists in FINISHED_SHEET
                    matches = finished_data[finished_data['ID number'] == id_number]
                    if not matches.empty:
                        # Update the existing record
                        finished_index = matches.index[0]
                        # Remove Training Status for FINISHED_SHEET
                        finished_row = updated_row[:-1]
                        update_data(FINISHED_SHEET, finished_index, finished_row)
                    else:
                        # Add to FINISHED_SHEET
                        finished_row = updated_row[:-1]  # Remove Training Status
                        add_data(FINISHED_SHEET, finished_row)
            
            # Reset editing state
            cancel_editing()
            # Refresh the data
            get_data(sheet_name, force_refresh=True)
            get_data(FINISHED_SHEET, force_refresh=True)
            time.sleep(1)
            st.rerun()
        else:
            st.error("Failed to update trainer information.")

# Function to confirm deletion
def confirm_delete(index, sheet_name):
    st.session_state.show_delete_confirm = True
    st.session_state.delete_index = index
    st.session_state.delete_sheet = sheet_name

# Function to cancel deletion
def cancel_delete():
    st.session_state.show_delete_confirm = False
    st.session_state.delete_index = -1

# Function to execute deletion
def execute_delete():
    if st.session_state.delete_index >= 0:
        sheet_name = st.session_state.delete_sheet
        index = st.session_state.delete_index
        
        # Get the ID number before deleting
        data = get_data(sheet_name)
        if index >= 0 and index < len(data):
            id_number = data.iloc[index]["ID number"]
            st.session_state.deleted_id = id_number
        
        if delete_data(sheet_name, index):
            st.success("Trainer deleted successfully!")
            
            # Check if we need to delete from FINISHED_SHEET too
            if 'deleted_id' in st.session_state:
                id_number = st.session_state.deleted_id
                finished_data = get_data(FINISHED_SHEET)
                
                # Check if this ID exists in FINISHED_SHEET
                if not finished_data.empty and 'ID number' in finished_data.columns:
                    matches = finished_data[finished_data['ID number'] == id_number]
                    if not matches.empty:
                        # Delete from FINISHED_SHEET too
                        finished_index = matches.index[0]
                        delete_data(FINISHED_SHEET, finished_index)
            
            # Reset deletion state
            cancel_delete()
            # Refresh the data
            get_data(sheet_name, force_refresh=True)
            get_data(FINISHED_SHEET, force_refresh=True)
            time.sleep(1)
            st.rerun()
        else:
            st.error("Failed to delete trainer.")

# Function to generate certificate
def generate_certificate(full_name, certificate_no, issue_date, id_number, certification_type, qr_code_buf=None):
    try:
        # Path to your existing PDF template
        template_path = "certificate_template.pdf"
        
        # Check if the template file exists
        if not os.path.exists(template_path):
            st.error(f"Certificate template not found at {template_path}")
            return None
        
        # Create temporary files for the output and overlay
        output_pdf_path = "output_certificate.pdf"
        overlay_pdf_path = "overlay.pdf"
        
        # Register Times New Roman font if available
        try:
            # Try to register Times New Roman font
            pdfmetrics.registerFont(TTFont('Times-New-Roman', 'times.ttf'))
            times_font = 'Times-New-Roman'
        except:
            # Fall back to Helvetica if Times New Roman is not available
            times_font = 'Helvetica'
        
        # Create the canvas for adding text
        c = canvas.Canvas(overlay_pdf_path, pagesize=A4)
        
        # Calculate font size based on name length to maintain consistent positioning
        name_length = len(full_name)
        if name_length <= 10:
            font_size = 24  # Larger font for short names
        elif name_length <= 20:
            font_size = 20  # Medium font for medium names
        elif name_length <= 30:
            font_size = 16  # Smaller font for longer names
        else:
            font_size = 14  # Very small font for very long names
        
        # Set font and size for the name
        c.setFont(f"{times_font}-Bold", font_size)
        
        # Position for the full name - keep this position consistent regardless of name length
        c.drawString(155, 420, full_name)
        
        # Add ID number with smaller font - now bold
        c.setFont(f"{times_font}-Bold", 12)
        c.drawString(250, 405, f"{id_number}")
        
        # Add Type of Certification - now bold
        c.setFont(f"{times_font}-Bold", 16)
        c.drawString(175, 530, f"{certification_type}")
        
        # Position for certificate number (adjust as needed)
        c.setFont(f"{times_font}", 12)
        c.drawString(350, 215, f"{certificate_no}")
        
        # Position for issue date (adjust as needed)
        c.drawString(150, 215, f"{issue_date}")
        
        # Add QR code if provided
        if qr_code_buf:
            # Create a temporary file for the QR code
            qr_temp_path = "temp_qr.png"
            with open(qr_temp_path, 'wb') as qr_file:
                qr_file.write(qr_code_buf.getvalue())
            
            # Add QR code to the PDF (adjust position and size as needed)
            c.drawImage(qr_temp_path, 450, 30, width=100, height=100)
        
        # Save the canvas
        c.save()
        
        # Open the template PDF and overlay PDF
        with open(template_path, "rb") as template_file, open(overlay_pdf_path, "rb") as overlay_file:
            template_pdf = PyPDF2.PdfReader(template_file)
            overlay_pdf = PyPDF2.PdfReader(overlay_file)
            
            # Create a PDF writer for the output
            output_writer = PyPDF2.PdfWriter()
            
            # Get the first page from each PDF
            template_page = template_pdf.pages[0]
            overlay_page = overlay_pdf.pages[0]
            
            # Merge the template with the overlay
            template_page.merge_page(overlay_page)
            
            # Add the merged page to the output
            output_writer.add_page(template_page)
            
            # Write the output to the file
            with open(output_pdf_path, "wb") as output_file:
                output_writer.write(output_file)
        
        # Read the output file into a BytesIO object for download
        with open(output_pdf_path, "rb") as f:
            pdf_bytes = BytesIO(f.read())
        
        # Clean up temporary files
        if os.path.exists(overlay_pdf_path):
            os.remove(overlay_pdf_path)
        if os.path.exists(output_pdf_path):
            os.remove(output_pdf_path)
        if qr_code_buf and os.path.exists("temp_qr.png"):
            os.remove("temp_qr.png")
        
        return pdf_bytes
    except Exception as e:
        st.error(f"Error generating certificate: {e}")
        import traceback
        st.error(traceback.format_exc())
        return None

# Function to get download link for certificate
def get_certificate_download_link(certificate_buf, filename="certificate.pdf"):
    try:
        b64 = base64.b64encode(certificate_buf.getvalue()).decode()
        href = f'<a href="data:application/pdf;base64,{b64}" download="{filename}" class="certificate-download">Download Certificate</a>'
        return href
    except Exception as e:
        st.error(f"Error creating certificate download link: {e}")
        return ""

# Main UI
st.markdown('<h1 class="main-header">⛑️ The SARCS Western Cape - Trainers Management System</h1>', unsafe_allow_html=True)

# Sidebar
with st.sidebar:
    st.image("https://i0.wp.com/redcross.org.za/wp-content/uploads/2020/06/HEADER-LOGO-ONLY.png?fit=3726%2C998&ssl=1", width=300)
    
    st.markdown('<div class="sidebar-header"><h2>Trainers Management</h2></div>', unsafe_allow_html=True)
    
    st.markdown('<div class="sidebar-nav">', unsafe_allow_html=True)
    
    # Navigation options with icons
    option = st.radio(
        "",
        ["View Trainers", "Add Trainer", "Manage Sheets", "About"],
        format_func=lambda x: f"👥 {x}" if x == "View Trainers" else 
                             (f"➕ {x}" if x == "Add Trainer" else 
                             (f"📋 {x}" if x == "Manage Sheets" else f"ℹ️ {x}"))
    )
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown('<div class="sidebar-divider"></div>', unsafe_allow_html=True)
    
    # Additional options based on current view
    if option == "View Trainers":
        st.markdown('<h3 style="color: #457b9d;">View Options</h3>', unsafe_allow_html=True)
        
        # Load sheet metadata
        metadata = load_sheet_metadata()
        
        # Create a list of available sheets with their display names
        sheet_options = [(FINISHED_SHEET, "🎓 Finished Trainers")]
        
        for sheet_name in st.session_state.available_sheets:
            if sheet_name != METADATA_SHEET:
                display_name = metadata.get(sheet_name, {}).get("display_name", sheet_name)
                sheet_options.append((sheet_name, f"📚 {display_name}"))
        
        # Create a radio button for sheet selection
        selected_option = st.radio(
            "Select Sheet", 
            [option[1] for option in sheet_options],
            key="view_options"
        )
        
        # Find the selected sheet name
        for sheet_name, display_name in sheet_options:
            if display_name == selected_option:
                st.session_state.selected_sheet = sheet_name
                break
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown('<div class="sidebar-divider"></div>', unsafe_allow_html=True)
    
    # Action buttons
    st.markdown('<h3 style="color: #457b9d;">Actions</h3>', unsafe_allow_html=True)
    
    if st.button("🔄 Refresh Data", key="refresh_button"):
        if refresh_all_data():
            st.success("All data refreshed successfully!")
        else:
            st.error("Error refreshing data. Please try again.")
    
    if st.button("📊 Export to CSV", key="export_button"):
        try:
            # Create a BytesIO object to store the CSV
            csv_buffer = BytesIO()
            
            # Get the data from the selected sheet
            data = get_data(st.session_state.selected_sheet)
            
            # Get display name for the sheet
            metadata = load_sheet_metadata()
            if st.session_state.selected_sheet == FINISHED_SHEET:
                display_name = "finished_trainers"
            else:
                display_name = metadata.get(st.session_state.selected_sheet, {}).get("display_name", st.session_state.selected_sheet)
                display_name = display_name.lower().replace(" ", "_")
            
            filename = f"{display_name}.csv"
            
            # Convert to CSV
            data.to_csv(csv_buffer, index=False)
            csv_buffer.seek(0)
            
            # Create download link
            b64 = base64.b64encode(csv_buffer.getvalue()).decode()
            href = f'<a href="data:file/csv;base64,{b64}" download="{filename}">Download {filename}</a>'
            
            st.markdown(href, unsafe_allow_html=True)
            st.success(f"Data ready for download!")
        except Exception as e:
            st.error(f"Error exporting data: {e}")
    
    # System info at the bottom
    st.markdown('<div class="sidebar-divider"></div>', unsafe_allow_html=True)
    st.markdown('<div class="sidebar-footer">', unsafe_allow_html=True)
    st.write(f"Last refresh: {time.strftime('%H:%M:%S', time.localtime(st.session_state.last_refresh))}")
    
    # Show stats in the sidebar
    finished_trainers = get_data(FINISHED_SHEET)
    
    # Count total trainers across all sheets
    total_trainers = len(finished_trainers)
    for sheet_name in st.session_state.available_sheets:
        if sheet_name != METADATA_SHEET:
            sheet_data = get_data(sheet_name)
            total_trainers += len(sheet_data)
    
    col1, col2 = st.columns(2)
    with col1:
        st.metric("Total", total_trainers)
    with col2:
        st.metric("Finished", len(finished_trainers))
    
    st.markdown('</div>', unsafe_allow_html=True)

# Show delete confirmation dialog if needed
if st.session_state.show_delete_confirm:
    sheet_name = st.session_state.delete_sheet
    index = st.session_state.delete_index
    
    # Get trainer name for confirmation message
    data = get_data(sheet_name)
    if index >= 0 and index < len(data):
        trainer_name = data.iloc[index]["Full Name"]
        
        # Store ID for potential cross-sheet deletion
        st.session_state.deleted_id = data.iloc[index]["ID number"]
        
        st.warning(f"Are you sure you want to delete {trainer_name}?")
        col1, col2 = st.columns(2)
        with col1:
            if st.button("Yes, Delete"):
                execute_delete()
        with col2:
            if st.button("Cancel"):
                cancel_delete()

# Show edit form if editing
if st.session_state.editing:
    st.markdown('<div class="edit-container">', unsafe_allow_html=True)
    st.markdown("### Edit Trainer Information")
    
    # Create form for editing
    with st.form("edit_trainer_form"):
        edit_data = st.session_state.edit_data
        
        st.markdown("### Personal Information")
        col1, col2 = st.columns(2)
        
        with col1:
            edit_data["Name(s)"] = st.text_input("Name(s)", value=edit_data.get("Name(s)", ""))
            edit_data["ID number"] = st.text_input("ID Number", value=edit_data.get("ID number", ""))
            edit_data["Contact No."] = st.text_input("Contact No.", value=edit_data.get("Contact No.", ""))
            edit_data["Nationality"] = st.text_input("Nationality", value=edit_data.get("Nationality", "SOUTH AFRICAN"))
        
        with col2:
            edit_data["Surname"] = st.text_input("Surname", value=edit_data.get("Surname", ""))
            edit_data["Gender"] = st.selectbox(
                "Gender", 
                ["Male", "Female", "Other"], 
                index=["Male", "Female", "Other"].index(edit_data.get("Gender", "Male"))
            )
            edit_data["Province"] = st.text_input("Province", value=edit_data.get("Province", ""))
            edit_data["Branch"] = st.text_input("Branch", value=edit_data.get("Branch", ""))
        
        # Auto-generate full name
        edit_data["Full Name"] = f"{edit_data.get('Name(s)', '')} {edit_data.get('Surname', '')}".strip()
        st.text_input("Full Name", value=edit_data["Full Name"], disabled=True)
        
        edit_data["Home Address"] = st.text_area("Home Address", value=edit_data.get("Home Address", ""))
        
        st.markdown("### Training Details")
        col1, col2 = st.columns(2)
        
        with col1:
            # Use dropdown for certification type
            edit_data["Type of Certification"] = st.text_input("Type of Certification", value=edit_data.get("Type of Certification", ""))
            # edit_data["Type of Certification"] = st.selectbox(
            #     "Type of Certification", 
            #     ["HOME BASED CARE LEVEL ONE", "HOME BASED CARE LEVEL 2 & 3", "FIRST AID LEVEL ONE","FIRST AID LEVEL 2 & 3" , "FUTHER EDUCATION TRAINING CERTIFICATE COMMUNITY HEALTH WORK NQF LEVEL 4"], 
            #     index=["HOME BASED CARE LEVEL ONE", "HOME BASED CARE LEVEL 2 & 3", "FIRST AID LEVEL ONE","FIRST AID LEVEL 2 & 3" , "FUTHER EDUCATION TRAINING CERTIFICATE COMMUNITY HEALTH WORK NQF LEVEL 4"]
            # )
            
            # # Certificate number
            # current_cert_no = edit_data.get("Certificate No.", "")
            # edit_data["Certificate No."] = st.text_input("Certificate No.", value=current_cert_no)
        
        with col2:
            # Handle date input - convert string to date if needed
            issue_date_str = edit_data.get("Issue Date", "")
            try:
                if issue_date_str:
                    issue_date = pd.to_datetime(issue_date_str).date()
                else:
                    issue_date = pd.Timestamp.now().date()
            except:
                issue_date = pd.Timestamp.now().date()
                
            edit_data["Issue Date"] = st.date_input("Issue Date", value=issue_date)
            edit_data["Candidate Full Description"] = st.text_area(
                "Candidate Full Description", 
                value=edit_data.get("Candidate Full Description", "")
            )
        
        # Training status
        current_status = edit_data.get("Training Status", "Pending")
        edit_data["Training Status"] = st.radio(
            "Training Status",
            ["Pending", "Finished"],
            index=0 if current_status == "Pending" else 1,
            help="Select 'Finished' if the trainer has completed training"
        )
        
        # Format the issue date to dd/mm/yyyy
        formatted_date = format_date_to_dd_mm_yyyy(edit_data.get('Issue Date', ''))
            
        # Generate the QR code data
        qr_data = generate_qr_code_data(
            edit_data.get('Full Name', ''), 
            edit_data.get('ID number', ''), 
            edit_data.get('Gender', ''), 
            edit_data.get('Type of Certification', ''), 
            edit_data.get('Certificate No.', ''), 
            edit_data.get('Issue Date', '')
        )

        # Store the QR code data in edit_data for saving
        edit_data["QR Code"] = qr_data
        
        # Form buttons
        col1, col2, col3 = st.columns(3)
        with col1:
            preview_qr = st.form_submit_button("Preview QR Code")
        with col2:
            save = st.form_submit_button("Save Changes")
        with col3:
            cancel = st.form_submit_button("Cancel")

    # Generate Certificate button (outside the form)
    if "Certificate No." in edit_data and edit_data["Certificate No."]:
        if st.button("Generate Certificate", key="edit_gen_cert"):
            # Get the necessary information
            full_name = edit_data["Full Name"]
            certificate_no = edit_data["Certificate No."]
            id_number = edit_data.get("ID number", "")
            certification_type = edit_data.get("Type of Certification", "")
            
            # Format the issue date
            issue_date = edit_data.get("Issue Date", "")
            if isinstance(issue_date, (pd.Timestamp, pd.DatetimeIndex)):
                issue_date = issue_date.strftime('%Y-%m-%d')
            elif hasattr(issue_date, 'strftime'):  # For datetime.date objects
                issue_date = issue_date.strftime('%Y-%m-%d')
            
            # Generate QR code
            qr_data = edit_data.get("QR Code", "")
            qr_buf = generate_qr(qr_data) if qr_data else None
            
            # Generate certificate
            certificate_buf = generate_certificate(full_name, certificate_no, issue_date, id_number, certification_type, qr_buf)
            
            if certificate_buf:
                st.markdown('<div class="certificate-container">', unsafe_allow_html=True)
                st.success("Certificate generated successfully!")
                st.markdown(get_certificate_download_link(certificate_buf, f"{full_name}_Certificate.pdf"), unsafe_allow_html=True)
                st.markdown('</div>', unsafe_allow_html=True)
    
    # Handle form actions
    if preview_qr:
        qr_buf = generate_qr(qr_data)
        if qr_buf:
            st.image(qr_buf, width=300)
            st.markdown(get_qr_download_link(qr_buf, f"{edit_data['Full Name']}_QR.png"), unsafe_allow_html=True)
    
    if save:
        # Update the session state
        st.session_state.edit_data = edit_data
        # Save the changes
        save_edited_trainer()
    
    if cancel:
        cancel_editing()
        st.rerun()
    
    st.markdown('</div>', unsafe_allow_html=True)

# Main content based on selected option
if option == "View Trainers" and not st.session_state.editing and not st.session_state.show_delete_confirm:
    sheet_name = st.session_state.selected_sheet
    
    # Get display name for the sheet
    metadata = load_sheet_metadata()
    if sheet_name == FINISHED_SHEET:
        sheet_title = "Finished Trainers"
    else:
        sheet_title = metadata.get(sheet_name, {}).get("display_name", sheet_name)
    
    st.markdown(f'<h2 class="subheader">{sheet_title}</h2>', unsafe_allow_html=True)
    
    # Add a force refresh button for the current sheet
    if st.button("🔄 Force Refresh This Sheet", key="force_refresh_sheet"):
        data = force_refresh_specific_sheet(sheet_name)
        st.success(f"Refreshed {sheet_title} sheet with {len(data)} records")
        st.rerun()
    
    # Search and filter options
    col1, col2 = st.columns([3, 1])
    with col1:
        search_query = st.text_input("Search by Name, ID Number, or Certificate No.")
    with col2:
        filter_options = ["All", "Province", "Branch", "Gender", "Type of Certification"]
        if sheet_name != FINISHED_SHEET:
            filter_options.append("Training Status")
        filter_option = st.selectbox("Filter by", filter_options)
    
    if filter_option != "All":
        data = get_data(sheet_name)
        if not data.empty and filter_option in data.columns:
            filter_values = ["All"] + sorted(data[filter_option].dropna().unique().tolist())
            filter_value = st.selectbox(f"Select {filter_option}", filter_values)
    
    # Get and filter data
    data = get_data(sheet_name)
    
    # Apply search filter
    if search_query:
        data = data[data.apply(lambda row: search_query.lower() in str(row.values).lower(), axis=1)]
    
    # Apply dropdown filter
    if filter_option != "All" and 'filter_value' in locals() and filter_value != "All" and filter_option in data.columns:
        data = data[data[filter_option] == filter_value]
    
    # Display record count
    st.write(f"Showing {len(data)} records")
    
    # Display data with action buttons
    if not data.empty:
        # Create a custom dataframe with action buttons
        st.markdown("### Trainers List")
        
        # Create a container for the table
        table_container = st.container()
        
        with table_container:
            # Display each row with action buttons
            for index, row in data.iterrows():
                col1, col2, col3, col4, col5, col6, col7 = st.columns([3, 2, 2, 2, 1, 1, 1])
                
                with col1:
                    st.write(f"**{row['Full Name']}**")
                with col2:
                    st.write(f"ID: {row['ID number']}")
                with col3:
                    st.write(f"Type: {row['Type of Certification']}")
                with col4:
                    st.write(f"Cert: {row['Certificate No.']}")
                    if sheet_name != FINISHED_SHEET and "Training Status" in row:
                        status_class = "status-finished" if row["Training Status"] == "Finished" else "status-pending"
                        st.markdown(f"<span class='status-badge {status_class}'>{row['Training Status']}</span>", unsafe_allow_html=True)
                
                # Action buttons
                with col5:
                    # View button
                    if st.button("👁️", key=f"view_{index}", help="View Details"):
                        # Always get the latest data from the sheet to ensure QR code is up-to-date
                        latest_data = get_data(sheet_name, force_refresh=True)
                        # Find the record with matching ID number to ensure we have the latest data
                        for idx, latest_row in latest_data.iterrows():
                            if latest_row["ID number"] == row["ID number"]:
                                st.session_state.selected_record = latest_row
                                break
                        else:
                            # Fallback to the current row if not found
                            st.session_state.selected_record = row
                        st.rerun()
                
                with col6:
                    # Edit button
                    if st.button("✏️", key=f"edit_{index}", help="Edit"):
                        start_editing(index, sheet_name)
                        st.rerun()
                
                with col7:
                    # Delete button
                    if st.button("🗑️", key=f"delete_{index}", help="Delete"):
                        confirm_delete(index, sheet_name)
                        st.rerun()
                
                # Add a separator line
                st.markdown("---")
        
        # Display selected record details if any
        if 'selected_record' in st.session_state:
            st.markdown("### Trainer Details")
            record = st.session_state.selected_record
            
            col1, col2 = st.columns(2)
            with col1:
                st.markdown(f"### {record['Full Name']}")
                
                # Display record details in a more organized way
                st.markdown("**Personal Information**")
                st.write(f"ID Number: {record['ID number']}")
                st.write(f"Gender: {record['Gender']}")
                st.write(f"Nationality: {record['Nationality']}")
                st.write(f"Contact: {record['Contact No.']}")
                st.write(f"Address: {record['Home Address']}")
            
            with col2:
                st.markdown("**Training Details**")
                if "Certificate No." in record and record["Certificate No."]:
                    st.write(f"Certificate No: {record['Certificate No.']}")
                if "Issue Date" in record and record["Issue Date"]:
                    st.write(f"Issue Date: {record['Issue Date']}")
                if "Type of Certification" in record and record["Type of Certification"]:
                    st.write(f"Type: {record['Type of Certification']}")
                st.write(f"Province: {record['Province']}")
                st.write(f"Branch: {record['Branch']}")
                
                if sheet_name != FINISHED_SHEET and "Training Status" in record:
                    status_class = "status-finished" if record["Training Status"] == "Finished" else "status-pending"
                    st.markdown(f"**Status:** <span class='status-badge {status_class}'>{record['Training Status']}</span>", unsafe_allow_html=True)
            
            # Display QR code if available
            if "QR Code" in record and record['QR Code']:
                st.markdown("**QR Code**")
                qr_buf = generate_qr(record['QR Code'])
                if qr_buf:
                    st.image(qr_buf, width=200)
                    st.markdown(get_qr_download_link(qr_buf, f"{record['Full Name']}_QR.png"), unsafe_allow_html=True)
            
            # Generate Certificate button
            if "Certificate No." in record and record["Certificate No."]:
                if st.button("Generate Certificate", key="gen_cert"):
                    # Get the necessary information
                    full_name = record["Full Name"]
                    certificate_no = record["Certificate No."]
                    id_number = record.get("ID number", "")
                    certification_type = record.get("Type of Certification", "")
                    issue_date = record.get("Issue Date", datetime.now().strftime('%Y-%m-%d'))
                    
                    # Generate QR code if not already available
                    if 'qr_buf' not in locals():
                        qr_buf = generate_qr(record['QR Code']) if "QR Code" in record else None
                    
                    # Generate certificate
                    certificate_buf = generate_certificate(full_name, certificate_no, issue_date, id_number, certification_type, qr_buf)
                    
                    if certificate_buf:
                        st.markdown('<div class="certificate-container">', unsafe_allow_html=True)
                        st.success("Certificate generated successfully!")
                        st.markdown(get_certificate_download_link(certificate_buf, f"{full_name}_Certificate.pdf"), unsafe_allow_html=True)
                        st.markdown('</div>', unsafe_allow_html=True)
            
            # Action buttons for the selected record
            col1, col2, col3 = st.columns(3)
            with col1:
                if st.button("Edit This Trainer"):
                    # Find the index of this record
                    for idx, row in data.iterrows():
                        if row["ID number"] == record["ID number"]:
                            start_editing(idx, sheet_name)
                            st.rerun()
                            break
            
            with col2:
                if sheet_name != FINISHED_SHEET and record.get("Training Status") != "Finished":
                    col2a, col2b = st.columns(2)
                    with col2a:
                        if st.button("Mark as Finished"):
                            # Find the index of this record
                            for idx, row in data.iterrows():
                                if row["ID number"] == record["ID number"]:
                                    if mark_trainer_as_finished(sheet_name, idx):
                                        st.success("Trainer marked as finished!")
                                        time.sleep(1)
                                        st.rerun()
                                    break
                    with col2b:
                        # Update the call in the "Mark & Generate Certificate" button
                        if "Certificate No." in record and record["Certificate No."]:
                            if st.button("Mark & Generate Certificate"):
                                # Find the index of this record
                                for idx, row in data.iterrows():
                                    if row["ID number"] == record["ID number"]:
                                        if mark_trainer_as_finished(sheet_name, idx):
                                            # Get the necessary information
                                            full_name = record["Full Name"]
                                            certificate_no = record["Certificate No."]
                                            id_number = record.get("ID number", "")
                                            certification_type = record.get("Type of Certification", "")
                                            issue_date = record.get("Issue Date", datetime.now().strftime('%Y-%m-%d'))
                                            
                                            # Generate QR code
                                            qr_buf = generate_qr(record['QR Code']) if "QR Code" in record else None
                                            
                                            # Generate certificate
                                            certificate_buf = generate_certificate(full_name, certificate_no, issue_date, id_number, certification_type, qr_buf)
                                            
                                            if certificate_buf:
                                                st.markdown('<div class="certificate-container">', unsafe_allow_html=True)
                                                st.success("Trainer marked as finished and certificate generated!")
                                                st.markdown(get_certificate_download_link(certificate_buf, f"{full_name}_Certificate.pdf"), unsafe_allow_html=True)
                                                st.markdown('</div>', unsafe_allow_html=True)
                                        break
            
            with col3:
                if st.button("Close Details"):
                    del st.session_state.selected_record
                    st.rerun()
    else:
        st.info("No records found matching your search criteria.")

elif option == "Add Trainer" and not st.session_state.editing and not st.session_state.show_delete_confirm:
    st.markdown('<h2 class="subheader">Add New Trainer</h2>', unsafe_allow_html=True)
    
    # Load sheet metadata
    metadata = load_sheet_metadata()
    
    # Create a list of available sheets with their display names
    sheet_options = []
    for sheet_name in st.session_state.available_sheets:
        if sheet_name != METADATA_SHEET and sheet_name != FINISHED_SHEET:
            display_name = metadata.get(sheet_name, {}).get("display_name", sheet_name)
            sheet_options.append((sheet_name, display_name))
    
    # If no sheets available, prompt to create one
    if not sheet_options:
        st.warning("No certification sheets available. Please create a sheet first in the 'Manage Sheets' section.")
        if st.button("Go to Manage Sheets"):
            st.session_state.option = "Manage Sheets"
            st.rerun()
    else:
        # Continue with the rest of the code for adding trainers
        # Select sheet for adding trainer
        selected_sheet_name = None
        selected_sheet_display = st.selectbox(
            "Select Certification Type",
            [option[1] for option in sheet_options]
        )
        
        # Find the selected sheet name
        for sheet_name, display_name in sheet_options:
            if display_name == selected_sheet_display:
                selected_sheet_name = sheet_name
                break
        
        if not selected_sheet_name:
            st.error("Please select a valid certification type")
        else:
            # Initialize form data from session state
            form_data = st.session_state.form_data
            
            # Create form
            with st.form("add_trainer_form"):
                st.markdown("### Personal Information")
                col1, col2 = st.columns(2)
                
                with col1:
                    name = st.text_input("Name(s)", value=form_data.get("Name(s)", ""))
                    id_number = st.text_input("ID Number", value=form_data.get("ID number", ""))
                    contact_no = st.text_input("Contact No.", value=form_data.get("Contact No.", ""))
                    nationality = st.text_input("Nationality", value=form_data.get("Nationality", "South African"))
                
                with col2:
                    surname = st.text_input("Surname", value=form_data.get("Surname", ""))
                    gender = st.selectbox("Gender", ["Male", "Female", "Other"], index=["Male", "Female", "Other"].index(form_data.get("Gender", "Male")))
                    province = st.text_input("Province", value=form_data.get("Province", ""))
                    branch = st.text_input("Branch", value=form_data.get("Branch", ""))
                
                # Auto-generate full name
                full_name = f"{name} {surname}".strip()
                st.text_input("Full Name", value=full_name, disabled=True)
                
                home_address = st.text_area("Home Address", value=form_data.get("Home Address", ""))
                
                st.markdown("### Training Details")
                col1, col2 = st.columns(2)
                
                with col1:
                    certificate_options = ["FIRST AID LEVEL 1", "FIRST AID LEVEL 2 & 3", "HOME BASED CARE LEVEL 1", "HOME BASED CARE LEVEL 2 & 3"]
                    # Let the user select one or more certificate types
                    certification_type = st.multiselect("Select Certification Type(s)", certificate_options)
                    if certification_type:
                        formatted_type = ", ".join(certification_type)  # Convert list tostring
                        st.write(f"Type of Certification: **{formatted_type}**")
                    
                    
                    
                    # Certificate number - can be auto-generated or manually entered
                    suggested_cert_no = get_next_certificate_number(selected_sheet_name)
                    use_suggested = st.checkbox("Use auto-generated certificate number", value=False)
                    if use_suggested:
                        certificate_no = suggested_cert_no
                        st.write(f"Certificate Number: {suggested_cert_no}")
                    else:
                        certificate_no = st.text_input("Certificate No.", value=form_data.get("Certificate No.", ""))
                
                with col2:
                    issue_date = st.date_input("Issue Date (if already completed)")
                    candidate_description = st.text_area("Candidate Full Description", value=form_data.get("Candidate Full Description", ""))
                
                # Training status
                training_status = st.radio(
                    "Training Status",
                    ["Pending", "Finished"],
                    index=0,
                    help="Select 'Finished' if the trainer has completed training"
                )
                
                # Format the issue date to dd/mm/yyyy
                formatted_date = format_date_to_dd_mm_yyyy(issue_date)
                
                # Generate the QR code data
                qr_data = generate_qr_code_data(
                    full_name, id_number, gender, certification_type, certificate_no, issue_date
                )
                
                # Form buttons
                col1, col2 = st.columns(2)
                with col1:
                    preview_qr = st.form_submit_button("Preview QR Code")
                with col2:
                    submitted = st.form_submit_button("Add Trainer")
            
            # Save form data to session state
            st.session_state.form_data = {
                "Name(s)": name,
                "Surname": surname,
                "Full Name": full_name,
                "ID number": id_number,
                "Certificate No.": certificate_no,
                "Issue Date": issue_date,
                "Contact No.": contact_no,
                "Province": province,
                "Branch": branch,
                "Type of Certification": certification_type,
                "Gender": gender,
                "Home Address": home_address,
                "Nationality": nationality,
                "Candidate Full Description": candidate_description,
                "Training Status": training_status
            }
            
            # Handle QR code preview
            if preview_qr or ('qr_image' in st.session_state and st.session_state.qr_image is not None):
                qr_buf = generate_qr(qr_data)
                if qr_buf:
                    st.session_state.qr_image = qr_buf
                    st.markdown("### QR Code Preview")
                    st.image(qr_buf, width=300)
                    st.markdown(get_qr_download_link(qr_buf), unsafe_allow_html=True)
            
            # Add a Clear Form button
            if st.button("Clear Form"):
                st.session_state.form_data = {}
                st.session_state.qr_image = None
                st.rerun()
        
            # Handle form submission
            if submitted:
                # Validate form data
                errors = validate_form_data(st.session_state.form_data)
                
                if errors:
                    for error in errors:
                        st.error(error)
                else:
                    # Add record to the appropriate sheet(s)
                    row_data = [
                        name, surname, full_name, id_number, certificate_no, 
                        issue_date.strftime('%Y-%m-%d'), contact_no, province, 
                        branch, certification_type, gender, home_address, 
                        nationality, "", qr_data, candidate_description
                    ]
                    
                    # Always add to the selected sheet
                    success = add_data(selected_sheet_name, row_data)
                    
                    # If marked as finished, also add to FINISHED_SHEET
                    if training_status == "Finished":
                        # Add to FINISHED_SHEET
                        finished_row = row_data + ["Finished"]
                        add_data(FINISHED_SHEET, finished_row[:-1])  # Remove Training Status for FINISHED_SHEET
                        
                        success_message = f"Trainer added successfully to {certification_type} and Finished Trainers lists!"
                    else:
                        success_message = f"Trainer added successfully to {certification_type} list!"
                    
                    if success:
                        st.success(success_message)
                        # Clear form data
                        st.session_state.form_data = {}
                        st.session_state.qr_image = None
                        time.sleep(1)
                        st.rerun()

elif option == "Manage Sheets" and not st.session_state.editing and not st.session_state.show_delete_confirm:
    st.markdown('<h2 class="subheader">Manage Certification Sheets</h2>', unsafe_allow_html=True)
    
    # Load sheet metadata
    metadata = load_sheet_metadata(force_refresh=True)
    
    # Display existing sheets
    st.markdown("### Existing Certification Types")
    
    if not st.session_state.available_sheets or all(s == METADATA_SHEET or s == FINISHED_SHEET for s in st.session_state.available_sheets):
        st.info("No certification sheets created yet. Use the form below to create your first sheet.")
    else:
        # Create a grid of sheet cards
        cols = st.columns(3)
        col_index = 0
        
        for sheet_name in st.session_state.available_sheets:
            if sheet_name != METADATA_SHEET and sheet_name != FINISHED_SHEET:
                sheet_info = metadata.get(sheet_name, {})
                display_name = sheet_info.get("display_name", sheet_name)
                cert_format = sheet_info.get("cert_format", "")
                last_cert = sheet_info.get("last_cert_number", "0")
                creation_date = sheet_info.get("creation_date", "")
                
                # Get record count
                sheet_data = get_data(sheet_name)
                record_count = len(sheet_data)
                
                with cols[col_index % 3]:
                    st.markdown(f"""
                    <div class="sheet-card">
                        <div class="sheet-title">{display_name}</div>
                        <div class="sheet-info">Format: {cert_format}</div>
                        <div class="sheet-info">Last Certificate: {last_cert}</div>
                        <div class="sheet-info">Created: {creation_date}</div>
                        <div class="sheet-info">Records: <span class="sheet-count">{record_count}</span></div>
                    </div>
                    """, unsafe_allow_html=True)
                
                col_index += 1
    
    # Create new sheet form
    st.markdown("### Create New Certification Type")
    
    with st.form("create_sheet_form"):
        col1, col2 = st.columns(2)
        
        with col1:
            sheet_name = st.text_input("Sheet Name (internal ID, no spaces)", 
                                      help="This is used internally. Use only letters, numbers, and underscores.")
            display_name = st.text_input("Display Name", 
                                        help="This is what users will see when selecting the certification type.")
        
        with col2:
            cert_format = st.text_input("Certificate Number Format", 
                                       help="Enter the format exactly as you want it to appear, with or without placeholders.")
            st.write("Options:")
            st.write("1. Static prefix with auto-incrementing number (e.g., 'C/25/FA' will produce C/25/FA1, C/25/FA2, etc.)")
            st.write("2. With placeholders:")
            st.write("   - C/25/FA{####} → C/25/FA0001, C/25/FA0002, etc.")
            st.write("   - C/HCB/{###} → C/HCB/001, C/HCB/002, etc.")
        
        create_sheet = st.form_submit_button("Create New Sheet")
    
    if create_sheet:
        # Validate inputs
        if not sheet_name:
            st.error("Sheet Name is required")
        elif not display_name:
            st.error("Display Name is required")
        elif not cert_format:
            st.error("Certificate Number Format is required")
        elif " " in sheet_name:
            st.error("Sheet Name cannot contain spaces")
        elif sheet_name in st.session_state.available_sheets or sheet_name == METADATA_SHEET or sheet_name == FINISHED_SHEET:
            st.error(f"Sheet '{sheet_name}' already exists")
        else:
            # Create the new sheet
            if create_new_sheet(sheet_name, display_name, cert_format):
                st.success(f"Created new certification type: {display_name}")
                time.sleep(1)
                st.rerun()

elif option == "About" and not st.session_state.editing and not st.session_state.show_delete_confirm:
    st.markdown('<h2 class="subheader">About This Application</h2>', unsafe_allow_html=True)
    
    st.markdown("""
    ### Red Cross Society - Trainers Management System
    
    This application helps the Red Cross Society manage trainer records efficiently, tracking both in-progress and completed trainings across multiple certification types.
    
    #### Features:
    - Create and manage multiple certification types with custom certificate numbering
    - View trainers by certification type or view all completed certifications
    - Mark trainers as "Finished" when they complete training
    - Add new trainers to the system with the appropriate certification type
    - Edit trainer information
    - Delete trainers from the system
    - Generate QR codes for each trainer
    - Search and filter trainer records
    - Auto-generate sequential certificate numbers based on custom formats
    - Ensure trainers are unique based on ID number
    
    #### How to Use:
    1. **Manage Sheets**: Create certification types with custom certificate number formats
    2. **Add Trainer**: Enter details for a new trainer, selecting the appropriate certification type
    3. **View Trainers**: Browse trainers by certification type or view all finished trainers
    4. **Edit Trainer**: Click the edit icon next to a trainer to modify their information
    5. **Delete Trainer**: Click the delete icon next to a trainer to remove them
    6. **Mark as Finished**: Select a trainer and mark them as finished when they complete training
    
    #### Certificate Numbering System:
    - Each certification type can have its own numbering format
    - Use placeholders like {####} for 4-digit numbers with leading zeros
    - Examples:
      - First Aid: C/25/FA0034
      - Home Based Care: C/HCB/085
    
    #### Need Help?
    Contact the system administrator for assistance.
    """)
    
    st.markdown("### System Information")
    st.write(f"Last data refresh: {time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(st.session_state.last_refresh))}")
    
    # Count total trainers across all sheets
    finished_trainers = get_data(FINISHED_SHEET)
    total_trainers = len(finished_trainers)
    
    for sheet_name in st.session_state.available_sheets:
        if sheet_name != METADATA_SHEET and sheet_name != FINISHED_SHEET:
            sheet_data = get_data(sheet_name)
            total_trainers += len(sheet_data)
    
    col1, col2 = st.columns(2)
    with col1:
        st.metric("Total Trainers", total_trainers)
    with col2:
        st.metric("Completed Training", len(finished_trainers))
    
    # Calculate percentage
    if total_trainers > 0:
        completion_rate = (len(finished_trainers) / total_trainers) * 100
        # Ensure progress value is between 0 and 1
        progress_value = min(1.0, max(0.0, completion_rate / 100))
        st.progress(progress_value)
        st.write(f"Completion Rate: {completion_rate:.1f}%")
    
    if st.button("Check Google Sheets Connection"):
        try:
            # Check metadata sheet
            create_metadata_sheet_if_not_exists()
            
            # Check finished sheet
            get_data(FINISHED_SHEET, force_refresh=True)
            
            # Check all other sheets
            for sheet_name in st.session_state.available_sheets:
                if sheet_name != METADATA_SHEET:
                    get_data(sheet_name, force_refresh=True)
            
            st.success("Connection successful! All sheets are accessible.")
        except Exception as e:
            st.error(f"Connection failed: {e}")

# Fix 9: Add a function to sanitize sheet names for safer handling

# Fix 11: Add a function to force refresh all data
def refresh_all_data():
    """Force refresh all data from all sheets"""
    try:
        # Refresh metadata
        metadata = load_sheet_metadata(force_refresh=True)
        
        # Refresh FINISHED_SHEET
        get_data(FINISHED_SHEET, force_refresh=True)
        
        # Refresh all other sheets
        for sheet_name in st.session_state.available_sheets:
            if sheet_name != METADATA_SHEET:
                get_data(sheet_name, force_refresh=True)
        
        st.session_state.last_refresh = time.time()
        return True
    except Exception as e:
        st.error(f"Error refreshing data: {e}")
        return False

