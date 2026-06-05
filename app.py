import streamlit as st
import google.generativeai as genai
from PIL import Image
import openpyxl
from openpyxl.drawing.image import Image as ExcelImage
import io
import os
import json
from datetime import datetime

# 1. ページ設定
st.set_page_config(page_title="T&N コンクリート劣化診断 AI Suite Pro", layout="wide")

# カスタムCSS
st.markdown("""
    <style>
    .main { background-color: #0F172A; color: #FFFFFF; }
    .stApp { background-color: #0F172A; }
    h1, h2, h3, h4, h5, h6, p, span, label, .stMarkdown { color: #FFFFFF !important; font-family: 'Helvetica Neue', Arial, sans-serif; }
    .stWidgetFormLabel, p, label { font-size: 16px !important; font-weight: bold !important; color: #FFFFFF !important; }
    .stButton>button {
        background-color: #0284C7; color: #FFFFFF; border: 2px solid #38BDF8; border-radius: 12px;
        padding: 14px 28px; font-weight: bold; width: 100%; font-size: 18px; transition: all 0.3s;
    }
    .stButton>button:hover { background-color: #38BDF8; box-shadow: 0 0 20px #38BDF8; }
    .status-card { padding: 25px; background-color: #1E293B; border-radius: 16px; border-left: 8px solid #38BDF8; margin-bottom: 20px; }
    </style>
""", unsafe_allow_html=True)

# 2. パスワードセッション管理
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
        st.markdown("<h2 style='text-align: center; color: white;'>🔒 コンクリート劣化診断 AI Suite Pro</h2>", unsafe_allow_html=True)
        st.text_input("パスワードを入力してください", type="password", on_change=password_entered, key="password")
        return False
    return True

if check_password():
    if os.path.exists("logo.png"):
        st.image("logo.png", width=220)
    st.markdown("<h1 style='color: white;'>🚗 AI Suite Pro - 実務特化型コンクリート高精密診断システム</h1>", unsafe_allow_html=True)
    st.markdown("---")

    # 本物の最新の鍵をSecretsから読み込みます
    api_key = st.secrets.get("GEMINI_API_KEY", "")

    # 🛠️ 左側サイドバー：環境選択メニュー
    st.sidebar.markdown("<h2 style='color: white;'>🛠️ プロ診断士用 環境条件設定</h2>", unsafe_allow_html=True)
    struct_type = st.sidebar.selectbox("① 構造物の種類", ["（未選択・写真から自動判定）", "橋梁（上部工/下部工）", "ボックスカルバート", "擁壁（重力式/もたれ式）", "トンネル覆工", "港湾・河川構造物", "建築物基礎・柱・壁"])
    env_location = st.sidebar.selectbox("② 設置環境・大分類", ["（未選択・写真から自動判定）", "一般地域（屋外・雨掛かり）", "一般地域（日陰・軒下）", "塩害警戒地域（海岸付近）", "寒冷地・凍枯地域", "屋内（常時乾燥）"])
    wet_status = st.sidebar.selectbox("③ コンクリートの湿潤状態", ["（未選択・写真から自動判定）", "常時乾燥状態", "乾湿の繰り返し（ひび割れが進展しやすい）", "常時湿潤状態（漏水・滞水あり）"])
    cement_type = st.sidebar.selectbox("④ 使用セメントの種類（推測）", ["（未選択）", "普通ポルトランドセメント", "高炉セメント（B種など）", "早強ポルトランドセメント", "不明"])
    elapsed_years = st.sidebar.selectbox("⑤ 供用年数（経過年数）", ["（未選択）", "5年未満（初期欠陥の可能性）", "5年以上〜20年未満", "20年以上〜50年未満", "50年以上（高経年化）"])
    crack_type = st.sidebar.selectbox("⑥ 目視での主たる劣化症状", ["（未選択・写真から自動判定）", "ひび割れ（単一・規則性）", "亀甲状のひび割れ（ASRなどの疑い）", "エフロレッセンス（白華）の析出伴う", "コンクリートの剥離・鉄筋露出（爆裂現象）", "漏水・遊離石灰を伴う錆汁"])

    # 左右の2カラムレイアウト
    col1, col2 = st.columns([1, 1])

    with col1:
        st.markdown("<h3 style='color: white;'>🏢 提出用 業務情報入力（空欄でもOK）</h3>", unsafe_allow_html=True)
        project_name = st.text_input("項目A：物件名（工事名・業務名）", placeholder="（例：〇〇高架橋修繕工事に伴う劣化調査）")
        location_name = st.text_input("項目B：調査位置・測定箇所", placeholder="（例：A1橋台 正面左側中央部）")
        inspector_name = st.text_input("項目C：調査担当者（コンクリート診断士名）", placeholder="（例：太田 正雄）")
        st.markdown("---")
        
        st.markdown("<h3 style='color: white;'>🔧 人間による補足情報入力（重複なし）</h3>", unsafe_allow_html=True)
        st.markdown("**【環境要因】**")
        cb_salt = st.checkbox("海岸線から2km以内（塩害リスク）")
        cb_freeze = st.checkbox("寒冷地・凍結防止剤の散布地域（凍害リスク）")
        cb_wet = st.checkbox("常時湿潤・漏水・滞水環境（アルカリ骨材反応・溶出リスク）")
        cb_traffic = st.checkbox("交通量が極めて多い（排気ガスによる中性化加速）")
        st.markdown("**【施工・初期欠陥要因】**")
        cb_janka = st.checkbox("ジャンカ・初期ひび割れの目視確認あり")
        cb_joint = st.checkbox("施工目地・コールドジョイント部")
        cb_cover = st.checkbox("設計かぶり厚の不足が疑われる・または既知")
        
        selected_human_factors = []
        if cb_salt: selected_human_factors.append("海岸線から2km以内（塩害リスク）")
        if cb_freeze: selected_human_factors.append("寒冷地・凍結防止剤の散布地域（凍害リスク）")
        if cb_wet: selected_human_factors.append("常時湿潤・漏水・滞水環境（アルカリ骨材反応・溶出リスク）")
        if cb_traffic: selected_human_factors.append("交通量が極めて多い（排気ガスによる中性化加速）")
        if cb_janka: selected_human_factors.append("ジャンカ・初期ひび割れの目視確認あり")
        if cb_joint: selected_human_factors.append("施工目地・コールドジョイント部")
        if cb_cover: selected_human_factors.append("設計かぶり厚の不足が疑われる・または既知")
        human_factors_text = "、".join(selected_human_factors) if selected_human_factors else "特なし"

        st.markdown("---")
        st.markdown("<h3 style='color: white;'>📷 診断写真の選択（必須）</h3>", unsafe_allow_html=True)
        uploaded_file = st.file_uploader("ここにコンクリート構造物の写真をアップロードしてください", type=["jpg", "jpeg", "png"])
        if uploaded_file is not None:
            image = Image.open(uploaded_file)
            st.image(image, caption="診断対象のコンクリート写真", use_container_width=True)
            st.markdown("---")
            execute_analysis = st.button("🚀 この内容で高精密AI解析を実行する")

    with col2:
        st.markdown("<h3 style='color: white;'>📊 高精密診断レポート</h3>", unsafe_allow_html=True)
        if uploaded_file is not None and 'execute_analysis' in locals() and execute_analysis:
            if not api_key
