import streamlit as st
import google.generativeai as genai
from PIL import Image
import openpyxl
from openpyxl.drawing.image import Image as ExcelImage
import io
import os
import json

# 1. ページ設定（iPhoneやカーナビのようなスタイリッシュなダークテーマUI）
st.set_page_config(page_title="T&N コンクリート劣化診断 AI Suite Pro", layout="wide")

# スタイリッシュなデザインを適用するカスタムCSS
st.markdown("""
    <style>
    .main { background-color: #0F172A; color: #F8FAFC; }
    .stApp { background-color: #0F172A; }
    h1, h2, h3, h4 { color: #FFFFFF; font-family: 'Helvetica Neue', Arial, sans-serif; font-weight: 700; }
    .stButton>button {
        background-color: #1E293B; color: #38BDF8; 
        border: 2px solid #38BDF8; border-radius: 12px;
        padding: 12px 24px; font-weight: bold; width: 100%;
        transition: all 0.3s;
    }
    .stButton>button:hover { background-color: #38BDF8; color: #0F172A; box-shadow: 0 0 15px #38BDF8; }
    .status-card { padding: 20px; background-color: #1E293B; border-radius: 16px; border-left: 6px solid #38BDF8; margin-bottom: 20px; }
    .report-title { font-size: 24px; color: #38BDF8; margin-bottom: 10px; }
    </style>
""", unsafe_allow_html=True)

# 2. パスワードセッション管理（前回の仕組みを完全キープ）
if 'authenticated' not in st.session_state:
    st.session_state['authenticated'] = False

def check_password():
    def password_entered():
        if st.session_state["password"] == "masao0605":
            st.session_state["authenticated"] = True
            del st.session_state["password"]
        else:
            st.sidebar.error("❌ パスワードが違います")
            
    if not st.session_state["authenticated"]:
        st.markdown("<div style='text-align: center; margin-top: 50px;'>", unsafe_allow_html=True)
        # ログイン画面にもロゴを配置
        if os.path.exists("logo.png"):
            st.image("logo.png", width=250, use_container_width=False)
        st.markdown("<h2 style='text-align: center;'>🔒 コンクリート劣化診断 AI Suite Pro</h2>", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)
        
        st.text_input("関係者限定アクセスパスワード：", type="password", on_change=password_entered, key="password")
        return False
    return True

if check_password():
    # 3. アプリ本編（ホーム画面）
    # ヘッダー部分に「Ｔ＆日本メンテ開発㈱」ロゴを配置
    if os.path.exists("logo.png"):
        st.image("logo.png", width=220)
        
    st.title("🚗 AI Suite Pro - 高精密コンクリート劣化診断システム")
    st.markdown("実務特化型：プロの診断士の視点で、微細なひび割れや劣化原因を高精密に解析します。")
    st.markdown("---")

    # カーナビ/スマホ風の左右2カラムレイアウト
    col1, col2 = st.columns([1, 1])

    with col1:
        st.markdown("### 📷 診断写真の選択")
        uploaded_file = st.file_uploader("コンクリート構造物のアップロード (.jpg, .jpeg, .png)", type=["jpg", "jpeg", "png"])
        
        # APIキーの取得（前回の環境変数の仕組みを完全継承）
        api_key = os.environ.get("GEMINI_API_KEY")
        if not api_key:
            # 開発・テスト用のバックアップ設定
            api_key = st.secrets.get("GEMINI_API_KEY", "")

        if uploaded_file is not None:
            image = Image.open(uploaded_file)
            st.image(image, caption="対象構造物写真", use_container_width=True)
            
            st.markdown("---")
            execute_analysis = st.button("🚀 高精密AI解析を実行する")

    with col2:
        st.markdown("### 📊 高精密診断レポート")
        
        if uploaded_file is not None and execute_analysis:
            if not api_key:
                st.error("APIキーが設定されていません。StreamlitのSettingsを確認してください。")
            else:
                with st.spinner("🔍 プロの診断士AIが画像をセル単位で超精密解析中..."):
                    try:
                        # 4. 高精密AI診断（Gemini 1.5 Proによる徹底解析プロンプト）
                        genai.configure(api_key=api_key)
                        model = genai.GenerativeModel('gemini-1.5-pro')
                        
                        prompt = """
                        あなたはコンクリート診断士の最高峰プロフェッショナルです。
                        提供された写真のコンクリート構造物を、コンサルタントや役所へ提出するレベルで精密に診断してください。
                        
                        必ず以下の4つの項目を特定し、プログラムが処理できるように正確なJSONフォーマットのみで出力してください。
                        特に「ひび割れ幅」と「ひび割れ長さ」は、画像内のスケールやクラックスケールを模して、実務上想定される最も正確な数値を推測して出力してください。
                        
```json
                        {
                          "width": 0.15,
                          "length": 12.5,
                          "reason": "ここに詳細な劣化原因の推測を記述（乾燥収縮、中性化、負荷など）",
                          "solution": "ここに具体的な補修・対策案を記述（エポキシ樹脂注入、経過観察など）"
                        }
