import streamlit as st
from datetime import datetime
import os
import base64
import pandas as pd
import io
from PIL import Image
import openpyxl
from openpyxl.drawing.image import Image as XLImage
import io as io_lib
import tempfile
import subprocess
import webbrowser
import platform
import time

# ==============================================================================
# 1. PAGE CONFIGURATION
# ==============================================================================
st.set_page_config(
    page_title="CITIC Dicastal Portal",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Navigation management via session
if "page" not in st.session_state:
    st.session_state.page = "menu"
    st.session_state.logged_in = False
    st.session_state.login_success = False
    st.session_state.file_modified = False

# Function to encode local image to Base64
def get_base64_image(image_path):
    if os.path.exists(image_path):
        with open(image_path, "rb") as img_file:
            return base64.b64encode(img_file.read()).decode()
    return None

def get_base64_from_bytes(image_bytes):
    """Convert image bytes to base64"""
    return base64.b64encode(image_bytes).decode()

# ==============================================================================
# DATA INTEGRATION - LECTURE DIRECTE DEPUIS LE FICHIER EXCEL + IMAGES
# ==============================================================================
def load_project_data():
    """Load project data and images directly from the Excel file"""
    excel_path = "Projects -PN (1).xlsx"
    
    try:
        if not os.path.exists(excel_path):
            return pd.DataFrame(), {}
        
        df = pd.read_excel(excel_path, header=0)
        df = df.dropna(how='all')
        
        if len(df.columns) >= 4:
            df_clean = df.iloc[:, 0:4].copy()
            df_clean.columns = ['Program', 'Project', 'Diversity', 'Representative_PN']
            df_clean['Program'] = df_clean['Program'].ffill()
            df_clean['Project'] = df_clean['Project'].ffill()
            df_clean = df_clean.dropna(subset=['Project'])
            df_clean = df_clean[df_clean['Project'] != 'Project']
            
            df_clean['Program'] = df_clean['Program'].astype(str)
            df_clean['Program'] = df_clean['Program'].str.replace(r'\.0$', '', regex=True)
            df_clean['Program'] = df_clean['Program'].str.replace(r'\.0', '', regex=False)
            
            images_dict = {}
            try:
                wb = openpyxl.load_workbook(excel_path)
                ws = wb.active
                for img in ws._images:
                    if isinstance(img, XLImage):
                        anchor = img.anchor
                        if hasattr(anchor, '_from'):
                            row_idx = anchor._from.row + 1
                            col_idx = anchor._from.col + 1
                            if col_idx == 5:
                                image_data = img._data()
                                if image_data and row_idx <= len(df_clean):
                                    project_name = df_clean.iloc[row_idx - 1]['Project'] if row_idx - 1 < len(df_clean) else None
                                    if project_name:
                                        images_dict[project_name] = image_data
            except Exception as e:
                pass
            
            return df_clean, images_dict
        else:
            return df, {}
    except Exception as e:
        return pd.DataFrame(), {}

# ==============================================================================
# FONCTION POUR OUVRIR LE FICHIER EXCEL DIRECTEMENT
# ==============================================================================
def open_excel_file():
    """Ouvrir le fichier Excel directement avec l'application par défaut"""
    excel_path = os.path.abspath("Projects -PN (1).xlsx")
    
    if not os.path.exists(excel_path):
        return False, "File not found"
    
    try:
        system = platform.system()
        
        if system == "Windows":
            os.startfile(excel_path)
        elif system == "Darwin":  # macOS
            subprocess.run(["open", excel_path])
        else:  # Linux
            subprocess.run(["xdg-open", excel_path])
        
        return True, "File opened successfully"
    except Exception as e:
        return False, str(e)

# ==============================================================================
# FONCTION POUR VERIFIER SI LE FICHIER A ETE MODIFIE
# ==============================================================================
def get_file_modification_time():
    """Retourne la date de modification du fichier Excel"""
    excel_path = "Projects -PN (1).xlsx"
    if os.path.exists(excel_path):
        return os.path.getmtime(excel_path)
    return None

def check_file_modified(initial_time):
    """Vérifie si le fichier a été modifié depuis initial_time"""
    current_time = get_file_modification_time()
    if current_time and initial_time:
        return current_time > initial_time
    return False

# Load data
if 'project_data' not in st.session_state:
    data, images = load_project_data()
    if not data.empty:
        st.session_state.project_data = data
        st.session_state.images_data = images
    else:
        data = {
            'Program': ['1232301', '1232301'],
            'Project': ['TESLA OPAL FRONT KNUCKLE BASE', 'TESLA OPAL FRONT KNUCKLE BASE'],
            'Diversity': ['RH', 'LH'],
            'Representative_PN': ['2188316-00-B', '2188311-00-B']
        }
        df = pd.DataFrame(data)
        st.session_state.project_data = df
        st.session_state.images_data = {}

# Initialiser le timestamp du fichier
if 'file_timestamp' not in st.session_state:
    st.session_state.file_timestamp = get_file_modification_time()

image_path = "usine.png" 
img_base64 = get_base64_image(image_path)

if img_base64:
    bg_style = f"background-image: url('data:image/png;base64,{img_base64}');"
else:
    bg_style = "background: linear-gradient(135deg, #1a0000 0%, #2d0a0a 50%, #1a0000 100%);" 

# ==============================================================================
# 2. GLOBAL CSS DESIGN - RED THEME
# ==============================================================================
st.markdown(f"""
    <style>
        #MainMenu, footer, header {{visibility: hidden;}}
        .block-container {{padding: 0rem 0rem 0rem 0rem !important; max-width: 100% !important;}}
        
        .stApp {{
            background-color: #f8f5f5; 
            font-family: 'Inter', 'Segoe UI', system-ui, sans-serif;
        }} 

        .left-panel-corporate {{
            position: fixed;
            left: 0; top: 0; width: 30%;
            height: calc(100vh - 45px);
            {bg_style}
            background-size: cover; background-position: center;
            display: flex; flex-direction: column; justify-content: center;
            padding-left: 40px;
            z-index: 10;
            box-shadow: 4px 0 30px rgba(0,0,0,0.2);
            overflow: hidden;
        }}

        .left-panel-corporate::before {{
            content: ''; position: absolute; top: 0; left: 0; right: 0; bottom: 0;
            background: linear-gradient(135deg, rgba(20, 0, 0, 0.92) 0%, rgba(60, 10, 10, 0.88) 40%, rgba(100, 20, 20, 0.80) 100%);
            z-index: 1;
        }}

        .left-content-corp {{ position: relative; z-index: 2; }}
        
        .brand-title-main {{ 
            font-size: 48px !important;
            font-weight: 700 !important; 
            margin: 0 !important; 
            text-transform: uppercase; 
            letter-spacing: 4px;
            background: linear-gradient(135deg, #ffffff 0%, #ff6b6b 40%, #ff3d3d 70%, #ffffff 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            background-clip: text;
        }}

        .brand-subtitle-sub {{ 
            font-size: 20px !important;
            font-weight: 300 !important; 
            margin: 8px 0 0 0 !important; 
            text-transform: uppercase; 
            letter-spacing: 6px;
            color: rgba(255,200,200,0.7);
        }}

        .brand-divider {{
            width: 60px;
            height: 3px;
            background: linear-gradient(90deg, #e3051c, #ff6b6b, #e3051c);
            margin: 15px 0;
            border-radius: 2px;
        }}

        .right-corp-wrapper {{ padding: 20px 30px 5px 30px; text-align: center; }}
        .menu-title-corp {{ 
            font-size: 32px; 
            color: #1a0000; 
            font-weight: 300; 
            text-transform: uppercase; 
            letter-spacing: 6px; 
            margin: 0; 
        }}
        .menu-title-corp span {{ 
            color: #e3051c; 
            font-weight: 700; 
        }}

        .corp-line {{ 
            width: 60px; 
            height: 2px; 
            background: linear-gradient(90deg, #e3051c, #ff6b6b); 
            margin: 10px auto; 
            border-radius: 2px; 
        }}
        .corp-subtitle {{ 
            font-size: 14px; 
            color: #6b4a4a; 
            letter-spacing: 3px; 
            font-weight: 400;
            margin-bottom: 30px; 
        }}

        .custom-menu-card {{
            background: #ffffff;
            border: 1px solid rgba(227, 5, 28, 0.08);
            border-radius: 16px;
            padding: 40px 25px;
            min-height: 220px;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            box-shadow: 0 4px 20px rgba(20, 0, 0, 0.04);
            transition: all 0.4s ease;
            cursor: pointer;
            text-align: center;
        }}

        .custom-menu-card:hover {{
            border-color: rgba(227, 5, 28, 0.2);
            box-shadow: 0 20px 60px rgba(227, 5, 28, 0.10);
            transform: translateY(-8px);
        }}

        .custom-menu-card .card-icon {{ font-size: 40px; margin-bottom: 12px; }}
        .custom-menu-card .card-title {{ 
            color: #1a0000; 
            font-size: 20px; 
            font-weight: 600; 
            margin: 6px 0; 
            padding-bottom: 10px; 
            border-bottom: 2px solid #f0ecec; 
            width: 100%; 
            text-transform: uppercase; 
            letter-spacing: 2px; 
        }}
        .custom-menu-card .card-desc {{ 
            color: #6b4a4a; 
            font-size: 13px; 
            line-height: 1.5; 
            margin: 4px 0 0 0; 
        }}

        .info-header {{
            display: none !important;
        }}

        .doc-btn-group {{
            margin: 15px 0 5px 0;
            padding: 20px 15px 15px 15px;
            background: linear-gradient(135deg, #ffffff, #f8f5f5);
            border-radius: 16px;
            border: 2px solid #e3051c;
            box-shadow: 0 4px 25px rgba(227, 5, 28, 0.12);
        }}
        
        .doc-btn-title {{
            text-align: center;
            color: #1a0000;
            font-size: 20px;
            font-weight: 700;
            margin: 0 0 5px 0;
            letter-spacing: 3px;
            text-transform: uppercase;
        }}
        
        .doc-btn-title span {{
            color: #e3051c;
        }}
        
        .doc-btn-subtitle {{
            text-align: center;
            color: #6b4a4a;
            font-size: 13px;
            margin: 0 0 15px 0;
            letter-spacing: 0.5px;
        }}
        
        .doc-category-label {{
            font-size: 11px !important;
            font-weight: 600 !important;
            color: #888888 !important;
            padding: 4px 12px !important;
            background: #f5f5f5 !important;
            border-radius: 12px !important;
            display: inline-block !important;
            margin: 8px 0 6px 0 !important;
            border-left: 3px solid #e3051c !important;
            letter-spacing: 1.5px !important;
            text-transform: uppercase !important;
        }}
        
        .stButton > button {{
            border-radius: 12px !important;
            font-weight: 700 !important;
            font-size: 15px !important;
            padding: 14px 12px !important;
            width: 100% !important;
            transition: all 0.3s ease !important;
            min-height: 55px !important;
            letter-spacing: 0.5px !important;
            border: 2px solid !important;
            box-shadow: 0 4px 15px rgba(0,0,0,0.10) !important;
        }}
        
        .stButton > button:hover {{
            transform: translateY(-3px) scale(1.02) !important;
            box-shadow: 0 8px 30px rgba(0,0,0,0.20) !important;
        }}
        
        .stButton > button.doc-2d {{
            background: linear-gradient(135deg, #0d47a1, #1565c0) !important;
            color: white !important;
            border-color: #0d47a1 !important;
            box-shadow: 0 4px 15px rgba(13, 71, 161, 0.4) !important;
        }}
        .stButton > button.doc-2d:hover {{
            background: linear-gradient(135deg, #1565c0, #0d47a1) !important;
            box-shadow: 0 8px 30px rgba(13, 71, 161, 0.6) !important;
        }}
        
        .stButton > button.doc-3d {{
            background: linear-gradient(135deg, #1b5e20, #2e7d32) !important;
            color: white !important;
            border-color: #1b5e20 !important;
            box-shadow: 0 4px 15px rgba(27, 94, 32, 0.4) !important;
        }}
        .stButton > button.doc-3d:hover {{
            background: linear-gradient(135deg, #2e7d32, #1b5e20) !important;
            box-shadow: 0 8px 30px rgba(27, 94, 32, 0.6) !important;
        }}
        
        .stButton > button.doc-mold {{
            background: linear-gradient(135deg, #e65100, #f57c00) !important;
            color: white !important;
            border-color: #e65100 !important;
            box-shadow: 0 4px 15px rgba(230, 81, 0, 0.4) !important;
        }}
        .stButton > button.doc-mold:hover {{
            background: linear-gradient(135deg, #f57c00, #e65100) !important;
            box-shadow: 0 8px 30px rgba(230, 81, 0, 0.6) !important;
        }}
        
        .stButton > button.doc-standard {{
            background: linear-gradient(135deg, #4a148c, #6a1b9a) !important;
            color: white !important;
            border-color: #4a148c !important;
            box-shadow: 0 4px 15px rgba(74, 20, 140, 0.4) !important;
        }}
        .stButton > button.doc-standard:hover {{
            background: linear-gradient(135deg, #6a1b9a, #4a148c) !important;
            box-shadow: 0 8px 30px rgba(74, 20, 140, 0.6) !important;
        }}

        .stButton button {{
            border-radius: 10px !important;
            font-weight: 600 !important;
            letter-spacing: 0.5px;
            transition: all 0.3s ease !important;
            font-size: 13px !important;
            padding: 6px 14px !important;
            background: transparent !important;
            border: 1px solid rgba(227, 5, 28, 0.3) !important;
            color: #e3051c !important;
            width: 100% !important;
            margin-top: 10px !important;
        }}

        .stButton button:hover {{
            background: #e3051c !important;
            color: #ffffff !important;
            border: 1px solid #e3051c !important;
        }}

        .stButton button[kind="primary"] {{
            background: linear-gradient(135deg, #e3051c, #cc0418) !important;
            border: none !important;
            color: #ffffff !important;
        }}

        .stButton button[kind="primary"]:hover {{
            background: linear-gradient(135deg, #cc0418, #b00314) !important;
        }}

        .stButton button[kind="secondary"] {{
            background: #ffffff !important;
            border: 2px solid #e3051c !important;
            color: #e3051c !important;
            padding: 4px 14px !important;
            font-size: 13px !important;
        }}

        .stButton button[kind="secondary"]:hover {{
            background: #e3051c !important;
            color: #ffffff !important;
        }}

        .corp-footer {{
            position: fixed; 
            bottom: 0; 
            left: 0; 
            width: 100%; 
            height: 40px; 
            background: linear-gradient(90deg, #1a0000, #2d0a0a, #3d0d0d);
            display: flex; 
            align-items: center; 
            justify-content: space-between; 
            padding: 0 40px;
            color: rgba(255,200,200,0.6); 
            font-size: 12px; 
            z-index: 100;
            border-top: 1px solid rgba(255,100,100,0.05);
        }}

        .corp-footer span:first-child {{
            font-weight: 500;
            color: rgba(255,200,200,0.9);
            letter-spacing: 2px;
        }}

        .corp-footer span:last-child {{
            color: rgba(255,200,200,0.4);
            font-weight: 300;
        }}

        .result-with-image {{
            display: flex;
            gap: 20px;
            align-items: center;
            background: #faf8f8;
            border-left: 4px solid #e3051c;
            padding: 12px 15px;
            border-radius: 8px;
            margin-bottom: 10px;
            height: 100%;
        }}
        .result-with-image .result-info {{
            flex: 1;
        }}
        .result-with-image .result-image {{
            flex: 0 0 180px;
            max-width: 180px;
            display: flex;
            align-items: center;
            justify-content: center;
        }}
        .result-with-image .result-image img {{
            width: 100%;
            height: auto;
            max-height: 150px;
            object-fit: contain;
            border-radius: 8px;
            border: 1px solid #ddd;
            background: white;
            padding: 5px;
        }}
        .result-with-image .result-image .no-image {{
            width: 100%;
            height: 100px;
            display: flex;
            align-items: center;
            justify-content: center;
            background: #f0ecec;
            border-radius: 8px;
            color: #999;
            font-size: 12px;
            text-align: center;
            border: 1px dashed #ccc;
            padding: 10px;
        }}
        .result-with-image p {{
            margin: 3px 0;
            font-size: 14px;
        }}
        .result-with-image b {{
            color: #1a0000;
        }}
        .result-with-image span {{
            color: #4a2a2a;
        }}

        .result-card {{
            background: #faf8f8; 
            border-left: 4px solid #e3051c; 
            padding: 8px 15px; 
            border-radius: 8px; 
            margin-bottom: 5px;
            height: 100%;
        }}
        .result-card p {{
            margin: 2px 0;
            font-size: 14px;
        }}
        .result-card b {{
            color: #1a0000;
        }}
        .result-card span {{
            color: #4a2a2a;
        }}

        .stat-card {{
            background: linear-gradient(135deg, #ffffff, #faf8f8); 
            border: 1px solid rgba(227,5,28,0.15); 
            border-radius: 12px; 
            padding: 8px 16px; 
            flex: 1; 
            min-width: 90px; 
            text-align: center;
            box-shadow: 0 2px 8px rgba(0,0,0,0.04);
        }}
        .stat-card .stat-number {{
            font-size: 22px; 
            font-weight: 700; 
            color: #e3051c;
        }}
        .stat-card .stat-label {{
            font-size: 11px; 
            color: #6b4a4a; 
            letter-spacing: 1.5px; 
            text-transform: uppercase;
            margin-top: 1px;
        }}

        .stSelectbox label {{
            font-size: 14px !important;
            font-weight: 500 !important;
        }}
        .stSelectbox div[data-baseweb="select"] {{
            margin-top: 1px !important;
        }}

        .element-container {{
            margin-bottom: 3px !important;
        }}

        .block-container {{
            padding-top: 0rem !important;
            padding-bottom: 0rem !important;
        }}

        .row-widget.stColumns {{
            gap: 10px !important;
        }}

        .page-title-container {{
            text-align: center;
            margin: 5px 0 8px 0;
        }}
        .page-title {{
            font-size: 40px;
            font-weight: 700;
            color: #e3051c;
            margin: 0;
            letter-spacing: 4px;
            text-transform: uppercase;
        }}
        .page-title span {{
            color: #1a0000;
        }}
        .page-subtitle {{
            font-size: 16px;
            color: #6b4a4a;
            margin: 4px 0 0 0;
            letter-spacing: 3px;
        }}
        
        .search-title {{
            font-size: 19px;
            font-weight: 600;
            color: #1a0000;
            margin: 8px 0 5px 0;
        }}
        
        .refresh-btn {{
            background: transparent !important;
            border: 1px solid #2196F3 !important;
            color: #2196F3 !important;
        }}
        .refresh-btn:hover {{
            background: #2196F3 !important;
            color: #ffffff !important;
        }}
        
        /* Styles pour la page Company Info */
        .company-card {{
            background: white;
            border-radius: 16px;
            padding: 30px 35px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.05);
            border: 1px solid #f0ecec;
        }}
        
        .company-header {{
            text-align: center;
            margin-bottom: 25px;
        }}
        
        .company-name {{
            font-size: 42px;
            font-weight: 700;
            color: #e3051c;
            letter-spacing: 8px;
            margin: 0;
            text-transform: uppercase;
        }}
        
        .company-subtitle {{
            font-size: 18px;
            font-weight: 300;
            color: #1a0000;
            letter-spacing: 6px;
            margin: 2px 0;
            text-transform: uppercase;
        }}
        
        .company-divider {{
            width: 60px;
            height: 3px;
            background: #e3051c;
            margin: 12px auto 8px auto;
            border-radius: 2px;
        }}
        
        .company-tag {{
            color: #888;
            font-size: 13px;
            letter-spacing: 3px;
        }}
        
        .info-row {{
            display: flex;
            padding: 10px 0;
            border-bottom: 1px solid #f5f0f0;
            align-items: center;
        }}
        
        .info-row:last-child {{
            border-bottom: none;
        }}
        
        .info-label {{
            font-weight: 600;
            color: #1a0000;
            min-width: 120px;
            font-size: 14px;
        }}
        
        .info-value {{
            color: #1a0000;
            font-size: 14px;
        }}
        
        .info-value.highlight {{
            color: #e3051c;
            font-weight: 600;
        }}
        
        .company-footer {{
            text-align: center;
            color: #aaa;
            font-size: 12px;
            letter-spacing: 2px;
            padding: 15px 0 5px 0;
        }}
        
        /* Styles pour le bouton Open Excel */
        .open-excel-btn {{
            background: linear-gradient(135deg, #1a73e8, #0d47a1) !important;
            color: white !important;
            border: none !important;
            padding: 16px 32px !important;
            border-radius: 12px !important;
            font-size: 18px !important;
            font-weight: 600 !important;
            cursor: pointer !important;
            transition: all 0.3s ease !important;
            width: 100% !important;
            box-shadow: 0 4px 20px rgba(26, 115, 232, 0.3) !important;
        }}
        .open-excel-btn:hover {{
            transform: translateY(-2px) !important;
            box-shadow: 0 8px 35px rgba(26, 115, 232, 0.5) !important;
        }}
        
        .edit-container {{
            background: white;
            border-radius: 12px;
            padding: 20px;
            border: 1px solid #e0e0e0;
            margin: 10px 0;
        }}
        
        .success-box {{
            background: #e8f5e9;
            border: 2px solid #4caf50;
            border-radius: 10px;
            padding: 15px;
            margin: 10px 0;
            text-align: center;
        }}
        
        .info-box {{
            background: #e3f2fd;
            border: 1px solid #2196f3;
            border-radius: 8px;
            padding: 12px 15px;
            margin: 10px 0;
        }}
    </style>
""", unsafe_allow_html=True)

# ==============================================================================
# 3. LEFT PANEL
# ==============================================================================
st.markdown("""
    <div class="left-panel-corporate">
        <div class="left-content-corp">
            <h1 class="brand-title-main">Dicastal</h1>
            <div class="brand-divider"></div>
            <h2 class="brand-subtitle-sub">Morocco Casting</h2>
        </div>
    </div>
""", unsafe_allow_html=True)

col_left_void, col_right_main = st.columns([3, 7])

# ==============================================================================
# 4. RIGHT NAVIGATION STRUCTURE
# ==============================================================================
with col_right_main:
    
    # PAGE: MAIN MENU
    if st.session_state.page == "menu":
        st.markdown("""
            <div class="right-corp-wrapper">
                <h2 class="menu-title-corp">Main <span>Menu</span></h2>
                <div class="corp-line"></div>
                <div class="corp-subtitle">Integrated Production Management System</div>
            </div>
        """, unsafe_allow_html=True)
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        _, central_grid, _ = st.columns([0.05, 0.9, 0.05])
        
        with central_grid:
            card_col1, card_col2, card_col3 = st.columns(3, gap="large")
            
            with card_col1:
                st.markdown("""
                    <div class="custom-menu-card card-1">
                        <div class="card-icon">📊</div>
                        <div class="card-title">Consult</div>
                        <div class="card-desc">Real-time and historical manufacturing flow visualization.</div>
                    </div>
                """, unsafe_allow_html=True)
                if st.button("🔍 Open Consult", key="btn_consult", use_container_width=True):
                    data, images = load_project_data()
                    st.session_state.project_data = data
                    st.session_state.images_data = images
                    st.session_state.page = "consult"
                    st.rerun()
                
            with card_col2:
                st.markdown("""
                    <div class="custom-menu-card card-2">
                        <div class="card-icon">📝</div>
                        <div class="card-title">Modify</div>
                        <div class="card-desc">Editing, updating and entry of industrial data sheets.</div>
                    </div>
                """, unsafe_allow_html=True)
                if st.button("✏️ Open Modify", key="btn_modify", use_container_width=True):
                    st.session_state.page = "login"
                    st.rerun()

            with card_col3:
                st.markdown("""
                    <div class="custom-menu-card card-3">
                        <div class="card-icon">🏢</div>
                        <div class="card-title">Company Info</div>
                        <div class="card-desc">Corporate information and key details about Dicastal DMC.</div>
                    </div>
                """, unsafe_allow_html=True)
                if st.button("📄 Open Info", key="btn_info", use_container_width=True):
                    st.session_state.page = "company_info"
                    st.rerun()

    # PAGE: CONSULT REGISTRY
    elif st.session_state.page == "consult":
        col_back, col_mid, col_refresh = st.columns([0.12, 0.76, 0.12])
        with col_back:
            if st.button("← Back", type="secondary", key="return_from_consult"):
                st.session_state.page = "menu"
                st.rerun()
        with col_refresh:
            if st.button("🔄 Refresh", key="refresh_consult", help="Recharger les données depuis Excel"):
                data, images = load_project_data()
                st.session_state.project_data = data
                st.session_state.images_data = images
                st.rerun()
        
        st.markdown("""
            <div class="page-title-container">
                <div class="page-title">Project <span>Consultation</span></div>
                <div class="page-subtitle">Search and view project information</div>
            </div>
        """, unsafe_allow_html=True)
        
        df = st.session_state.project_data
        images_dict = st.session_state.images_data
        
        if not df.empty:
            total_rows = len(df)
            total_projects = len(df['Program'].unique())
            total_products = len(df['Project'].unique())
            
            st.markdown(f"""
                <div style="display: flex; gap: 20px; margin: 10px 20px 12px 20px; flex-wrap: wrap; justify-content: center;">
                    <div class="stat-card">
                        <div class="stat-number">{total_projects}</div>
                        <div class="stat-label">Projects</div>
                    </div>
                    <div class="stat-card">
                        <div class="stat-number">{total_rows}</div>
                        <div class="stat-label">References</div>
                    </div>
                    <div class="stat-card">
                        <div class="stat-number">{total_products}</div>
                        <div class="stat-label">Products</div>
                    </div>
                </div>
            """, unsafe_allow_html=True)
        
        st.markdown('<p class="search-title">🔍 Search Projects</p>', unsafe_allow_html=True)
        
        unique_projects = df['Project'].dropna().unique().tolist()
        unique_projects = [p for p in unique_projects if p != '']
        unique_projects.sort()
        
        col_search, col_btn = st.columns([4, 1])
        with col_search:
            search_value = st.selectbox(
                "Select a Project Name",
                [""] + unique_projects,
                key="project_search",
                label_visibility="visible"
            )
        with col_btn:
            st.write("")
            search_clicked = st.button("🔍 Search", key="search_project", type="primary")
        
        if search_clicked and search_value:
            result_df = df[df['Project'] == search_value]
            
            if not result_df.empty:
                rh_data = result_df[result_df['Diversity'] == 'RH']
                lh_data = result_df[result_df['Diversity'] == 'LH']
                
                col1, col2 = st.columns(2)
                
                with col1:
                    if not rh_data.empty:
                        for idx, row in rh_data.iterrows():
                            rep_pn = row['Representative_PN']
                            if pd.isna(rep_pn) or rep_pn == '' or rep_pn is None:
                                rep_pn_display = 'N/A'
                            else:
                                rep_pn_display = rep_pn
                            
                            project_name = row['Project']
                            image_html = ""
                            if project_name in images_dict:
                                img_b64 = get_base64_from_bytes(images_dict[project_name])
                                image_html = f'<img src="data:image/png;base64,{img_b64}" alt="Product Image">'
                            else:
                                image_html = '<div class="no-image">📷 Pas d\'image</div>'
                            
                            st.markdown(f"""
                                <div class="result-with-image">
                                    <div class="result-info">
                                        <p><b>Project Name:</b> <span>{row['Project']}</span></p>
                                        <p><b>Dicastal PN:</b> <span>{row['Program']}</span></p>
                                        <p><b>Diversity:</b> <span style="color: #e3051c; font-weight: 600;">{row['Diversity']}</span></p>
                                        <p><b>Representative Part Number:</b> <span>{rep_pn_display}</span></p>
                                    </div>
                                    <div class="result-image">
                                        {image_html}
                                    </div>
                                </div>
                            """, unsafe_allow_html=True)
                
                with col2:
                    if not lh_data.empty:
                        for idx, row in lh_data.iterrows():
                            rep_pn = row['Representative_PN']
                            if pd.isna(rep_pn) or rep_pn == '' or rep_pn is None:
                                rep_pn_display = 'N/A'
                            else:
                                rep_pn_display = rep_pn
                            
                            project_name = row['Project']
                            image_html = ""
                            if project_name in images_dict:
                                img_b64 = get_base64_from_bytes(images_dict[project_name])
                                image_html = f'<img src="data:image/png;base64,{img_b64}" alt="Product Image">'
                            else:
                                image_html = '<div class="no-image">📷 Pas d\'image</div>'
                            
                            st.markdown(f"""
                                <div class="result-with-image">
                                    <div class="result-info">
                                        <p><b>Project Name:</b> <span>{row['Project']}</span></p>
                                        <p><b>Dicastal PN:</b> <span>{row['Program']}</span></p>
                                        <p><b>Diversity:</b> <span style="color: #e3051c; font-weight: 600;">{row['Diversity']}</span></p>
                                        <p><b>Representative Part Number:</b> <span>{rep_pn_display}</span></p>
                                    </div>
                                    <div class="result-image">
                                        {image_html}
                                    </div>
                                </div>
                            """, unsafe_allow_html=True)
                
                st.markdown("""
                    <div class="doc-btn-group">
                        <div class="doc-btn-title">📄 Available <span>Documents</span></div>
                        <div class="doc-btn-subtitle">Click on a document type to view or download</div>
                    </div>
                """, unsafe_allow_html=True)
                
                st.markdown('<div class="doc-category-label">📐 2D Drawings</div>', unsafe_allow_html=True)
                col_2d_1, col_2d_2, col_2d_3 = st.columns(3)
                with col_2d_1:
                    if st.button("📐 2D Raw Part", key=f"2d_raw_{search_value}", use_container_width=True):
                        st.success(f"✅ 2D Raw Part Drawing for {search_value}")
                with col_2d_2:
                    if st.button("📐 2D Machining", key=f"2d_mach_{search_value}", use_container_width=True):
                        st.success(f"✅ 2D Machining Part Drawing for {search_value}")
                with col_2d_3:
                    if st.button("📐 2D Assembly", key=f"2d_assy_{search_value}", use_container_width=True):
                        st.success(f"✅ 2D Assembly Drawing for {search_value}")
                
                st.markdown('<div class="doc-category-label">📐 3D Drawings</div>', unsafe_allow_html=True)
                col_3d_1, col_3d_2, col_3d_3 = st.columns(3)
                with col_3d_1:
                    if st.button("📐 3D Raw Part", key=f"3d_raw_{search_value}", use_container_width=True):
                        st.success(f"✅ 3D Raw Part Model for {search_value}")
                with col_3d_2:
                    if st.button("📐 3D Machining", key=f"3d_mach_{search_value}", use_container_width=True):
                        st.success(f"✅ 3D Machining Part Model for {search_value}")
                with col_3d_3:
                    if st.button("📐 3D Assembly", key=f"3d_assy_{search_value}", use_container_width=True):
                        st.success(f"✅ 3D Assembly Model for {search_value}")
                
                st.markdown('<div class="doc-category-label">🏗️ Mold & Tooling</div>', unsafe_allow_html=True)
                col_mold_1, col_mold_2 = st.columns(2)
                with col_mold_1:
                    if st.button("🏗️ Mold Drawing", key=f"mold_{search_value}", use_container_width=True):
                        st.success(f"✅ Mold Drawing for {search_value}")
                with col_mold_2:
                    if st.button("🔧 Tooling Drawing", key=f"tooling_{search_value}", use_container_width=True):
                        st.success(f"✅ Tooling Drawing for {search_value}")
                
                st.markdown('<div class="doc-category-label">✅ Standards & Bushings</div>', unsafe_allow_html=True)
                col_std_1, col_std_2 = st.columns(2)
                with col_std_1:
                    if st.button("✅ Product Standard", key=f"standard_{search_value}", use_container_width=True):
                        st.success(f"✅ Product Standard for {search_value}")
                with col_std_2:
                    if st.button("🔩 Bushings Drawing", key=f"bushings_{search_value}", use_container_width=True):
                        st.success(f"✅ Bushings Drawing for {search_value}")
            else:
                st.warning("No results found")
        elif search_clicked and not search_value:
            st.info("Please select a project name")
        else:
            st.info("Select a project and click Search to see results")

    # PAGE: LOGIN FORM - AVEC OUVERTURE DIRECTE ET REFRESH AUTOMATIQUE
    elif st.session_state.page == "login":
        st.markdown("""
            <div class="page-title-container">
                <div class="page-title">Restricted <span>Area</span></div>
                <div class="page-subtitle">Authentication required for modification access</div>
            </div>
        """, unsafe_allow_html=True)
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        space_l, login_grid, space_r = st.columns([0.2, 0.6, 0.2])
        
        with login_grid:
            username = st.text_input("Username / ID", placeholder="Enter your identifier")
            password = st.text_input("Password", type="password", placeholder="Enter your access key")
            
            st.markdown("<br>", unsafe_allow_html=True)
            
            col1, col2 = st.columns(2)
            with col1:
                if st.button("🔐 Login", type="primary", key="login_btn", use_container_width=True):
                    if username == "DMC" and password == "dmc1122":
                        st.session_state.logged_in = True
                        st.session_state.login_success = True
                        # Sauvegarder le timestamp du fichier
                        st.session_state.file_timestamp = get_file_modification_time()
                        st.rerun()
                    else:
                        st.error("❌ Authentication failed")
                        st.session_state.login_success = False
                
            with col2:
                if st.button("← Back", type="secondary", key="back_from_login_btn", use_container_width=True):
                    st.session_state.page = "menu"
                    st.rerun()
            
            if st.session_state.get("login_success", False):
                st.markdown("---")
                
                st.markdown("""
                    <div class="success-box">
                        <p style="margin: 0; font-size: 16px; font-weight: 700; color: #2e7d32;">
                            ✅ Login successful!
                        </p>
                    </div>
                """, unsafe_allow_html=True)
                
                excel_path = "Projects -PN (1).xlsx"
                if os.path.exists(excel_path):
                    
                    # Message d'information
                    st.markdown("""
                        <div class="info-box">
                            <p style="margin: 0; font-size: 14px; color: #0d47a1; font-weight: 500;">
                                📝 <b>Edit Excel File</b>
                            </p>
                            <p style="margin: 5px 0 0 0; font-size: 13px; color: #1565c0;">
                                1. Click <b>"Open Excel File"</b> to edit directly in your spreadsheet application<br>
                                2. Make your changes and <b>save the file</b><br>
                                3. Click <b>"Refresh Data"</b> to load the changes
                            </p>
                        </div>
                    """, unsafe_allow_html=True)
                    
                    # Bouton Open Excel
                    col_open, col_refresh_btn = st.columns(2)
                    with col_open:
                        if st.button("📂 Open Excel File", type="primary", use_container_width=True):
                            success, message = open_excel_file()
                            if success:
                                st.success("✅ File opened successfully in your spreadsheet application!")
                                # Sauvegarder le timestamp pour détecter les modifications
                                st.session_state.file_timestamp = get_file_modification_time()
                            else:
                                st.error(f"❌ Error: {message}")
                    
                    with col_refresh_btn:
                        # Bouton Refresh avec détection automatique des modifications
                        if st.button("🔄 Refresh Data", type="primary", use_container_width=True):
                            current_time = get_file_modification_time()
                            if st.session_state.file_timestamp:
                                if current_time and current_time > st.session_state.file_timestamp:
                                    # Le fichier a été modifié
                                    data, images = load_project_data()
                                    st.session_state.project_data = data
                                    st.session_state.images_data = images
                                    st.session_state.file_timestamp = current_time
                                    st.success("✅ Data refreshed successfully! Changes loaded.")
                                    st.balloons()
                                else:
                                    st.info("ℹ️ No changes detected in the file.")
                            else:
                                data, images = load_project_data()
                                st.session_state.project_data = data
                                st.session_state.images_data = images
                                st.session_state.file_timestamp = get_file_modification_time()
                                st.success("✅ Data refreshed successfully!")
                    
                    # Détection automatique des modifications (refresh automatique)
                    st.markdown("---")
                    
                    # Vérification automatique
                    current_time = get_file_modification_time()
                    if st.session_state.file_timestamp and current_time:
                        if current_time > st.session_state.file_timestamp and not st.session_state.get('auto_refreshed', False):
                            # Le fichier a été modifié - refresh automatique
                            data, images = load_project_data()
                            st.session_state.project_data = data
                            st.session_state.images_data = images
                            st.session_state.file_timestamp = current_time
                            st.session_state.auto_refreshed = True
                            st.rerun()
                    
                    # Afficher le statut du fichier
                    st.markdown("""
                        <div style="background: #f5f5f5; border-radius: 8px; padding: 10px 15px; margin: 10px 0;">
                            <p style="margin: 0; font-size: 12px; color: #666;">
                                📄 <b>File status:</b> 
                    """, unsafe_allow_html=True)
                    
                    if st.session_state.file_timestamp:
                        last_modified = datetime.fromtimestamp(st.session_state.file_timestamp).strftime("%H:%M:%S")
                        st.markdown(f"""
                            <span style="color: #4caf50;">✓ Last modified: {last_modified}</span>
                            </p>
                        </div>
                        """, unsafe_allow_html=True)
                    else:
                        st.markdown("""
                            <span style="color: #ff9800;">⚠ No modification detected</span>
                            </p>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    # Go to Consult
                    st.markdown("<br>", unsafe_allow_html=True)
                    if st.button("📊 Go to Consult", type="primary", use_container_width=True):
                        st.session_state.page = "consult"
                        st.rerun()
                                
                else:
                    st.error(f"❌ File not found: {excel_path}")

    # PAGE: COMPANY INFORMATION
    elif st.session_state.page == "company_info":
        col_back, col_spacer = st.columns([0.12, 0.88])
        with col_back:
            if st.button("← Back", type="secondary", key="return_from_info"):
                st.session_state.page = "menu"
                st.rerun()

        st.markdown("<br>", unsafe_allow_html=True)
        
        col_left, col_center, col_right = st.columns([1, 1.8, 1])
        
        with col_center:
            st.markdown("""
            <div class="company-card">
                <div class="company-header">
                    <div class="company-name">DICASTAL</div>
                    <div class="company-subtitle">MOROCCO CASTING</div>
                    <div class="company-divider"></div>
                    <div class="company-tag">CITIC Dicastal Morocco Casting · DMC</div>
                </div>
            """, unsafe_allow_html=True)
            
            # Informations
            items = [
                ("📍 Address", "Atlantic Free Zone, Kénitra, Morocco", False),
                ("📞 Contact", "+212 5 37 00 00 00", False),
                ("📧 Email", "info@dicastal.ma", False),
                ("🏢 Activity", "Automotive Aluminum Casting", True),
                ("📅 Founded", "2019", False)
            ]
            
            for i, (label, value, highlight) in enumerate(items):
                highlight_class = "info-value highlight" if highlight else "info-value"
                st.markdown(f"""
                <div class="info-row">
                    <div class="info-label">{label}</div>
                    <div class="{highlight_class}">{value}</div>
                </div>
                """, unsafe_allow_html=True)
            
            st.markdown("""
                <div class="company-footer">🌍 Global Leader in Aluminum Casting Solutions</div>
            </div>
            """, unsafe_allow_html=True)

# ==============================================================================
# 5. FOOTER
# ==============================================================================
date_today = datetime.now().strftime("%d/%m/%Y")
st.markdown(f"""
    <div class="corp-footer">
        <span>CITIC DICASTAL</span>
        <span>{date_today}</span>
    </div>
""", unsafe_allow_html=True)