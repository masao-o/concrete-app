import streamlit as st
import google.generativeai as genai
from PIL import Image
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.drawing.image import Image as ExcelImage
import io
import os

# --- 1. ページ設定とCSS（デザイン・視認性確保） ---
st.set_page_config(page_title="T&N コンクリート劣化診断 AI Suite Pro", layout="wide")
st.markdown("""
<style>
.main { background-color: #0F172A; color: #FFFFFF; }
.stApp { background-color: #0F172A; }
section[data-testid="stSidebar"] { background-color: #1E293B !important; border-right: 1px solid #334155; }
h1, h2, h3, h4, h5, h6, p, span, label, .stMarkdown,
[data-testid="stSidebar"] label, [data-testid="stSidebar"] p, [data-testid="stSidebar"] span,
.stCheckbox label, div[data-testid="stMarkdownContainer"] p {
    color: #FFFFFF !important; font-family: 'Helvetica Neue', Arial, sans-serif; font-weight: bold !important;
}
input, textarea { color: #0F172A !important; font-weight: bold !important; }
input::placeholder, textarea::placeholder { color: #64748B !important; opacity: 1 !important; }
div[data-testid="stFileUploader"] section { background-color: #F8FAFC !important; border: 2px dashed #94A3B8 !important; }
div[data-testid="stFileUploader"] section div, div[data-testid="stFileUploader"] section p,
div[data-testid="stFileUploader"] section span, div[data-testid="stFileUploader"] section small {
    color: #475569 !important; font-weight: bold !important;
}
.stExpander div, .stExpander p, .stExpander span { color: #FFFFFF !important; }
.stButton>button {
    background-color: #0284C7; color: #FFFFFF; border: 2px solid #38BDF8; border-radius: 12px;
    padding: 14px 28px; font-weight: bold; width: 100%; font-size: 18px;
}
.stButton>button:hover { background-color: #38BDF8; box-shadow: 0 0 20px #38BDF8; }
.status-card { padding: 25px; background-color: #0F172A; border-radius: 16px; border-left: 8px solid #38BDF8; margin-bottom: 20px; border-top: 1px solid #334155; border-right: 1px solid #334155; border-bottom: 1px solid #334155; }
</style>
""", unsafe_allow_html=True)

# --- 2. パスワード認証 ---
if 'authenticated' not in st.session_state:
    st.session_state['authenticated'] = False

def check_password():
    def password_entered():
        if st.session_state["password"] == "tn0000":
            st.session_state["authenticated"] = True
            del st.session_state["password"]
        else:
            st.sidebar.error("❌ パスワードが違います")
            
    if not st.session_state["authenticated"]:
        if os.path.exists("logo.png"): 
            st.image("logo.png", width=250)
        st.markdown("<h2 style='text-align: center; color: white;'>🔒 閉域環境・コンクリート劣化診断 AI Suite Pro</h2>", unsafe_allow_html=True)
        st.text_input("アクセスパスワード（担当者専用）", type="password", on_change=password_entered, key="password")
        return False
    return True

if check_password():
    if os.path.exists("logo.png"): 
        st.image("logo.png", width=220)
        
    st.markdown("<h1 style='color: white;'>🚗 AI Suite Pro - 実務特化型コンクリート高精密診断システム</h1>", unsafe_allow_html=True)
    st.markdown("---")
    
    # APIキーの安全な取得
    api_key = st.secrets.get("GEMINI_API_KEY", "")
    
    # --- 3. サイドバー：プロ診断士用 環境条件設定 ---
    st.sidebar.markdown("## 🛠️ プロ診断士用 条件設定")
    struct_type = st.sidebar.selectbox("① 構造物の種類", ["（未選択・写真から自動判定）", "橋梁（上部工/下部工）", "ボックスカルバート", "擁壁", "トンネル覆工", "港湾・河川構造物", "建築物基礎・柱・壁"])
    
    env_location = st.sidebar.multiselect("② 設置環境・
