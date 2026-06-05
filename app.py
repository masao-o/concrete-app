import streamlit as st
import google.generativeai as genai
from PIL import Image
import openpyxl
from openpyxl.drawing.image import Image as ExcelImage
import io
import os
import json
from datetime import datetime

# 1. ページ設定（スタイリッシュかつ超高コントラストなUI）
st.set_page_config(page_title="T&N コンクリート劣化診断 AI Suite Pro", layout="wide")

# 文字をくっきり白く、見やすくするためのカスタムCSS
st.markdown("""
    <style>
    .main { background-color: #0F172A; color: #FFFFFF; }
    .stApp { background-color: #0F172A; }
    
    /* すべての見出し・テキストを純白にして視認性を最優先に */
    h1, h2, h3, h4, h5, h6, p, span, label, .stMarkdown { 
        color: #FFFFFF !important; 
        font-family: 'Helvetica Neue', Arial, sans-serif;
    }
    
    /* アップローダーや入力欄のラベル文字を大きく太く */
    .stWidgetFormLabel, p, label {
        font-size: 16px !important;
        font-weight: bold !important;
        color: #FFFFFF !important;
    }
    
    .stButton>button {
        background-color: #0284C7; color: #FFFFFF; 
        border: 2px solid #38BDF8; border-radius: 12px;
        padding: 14px 28px; font-weight: bold; width: 100%;
        font-size: 18px;
        transition: all 0.3s;
    }
    .stButton>button:hover { background-color: #38BDF8; box-shadow: 0 0 20px #38BDF8; }
    .status-card { padding: 25px; background-color: #1E293B; border-radius: 16px; border-left: 8px solid #38BDF8; margin-bottom: 20px; }
    </style>
""", unsafe_allow_html=True)

# 2. パスワードセッション管理（ご要望の「tn0000」に変更しました！）
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
    # 3. ホーム画面ヘッダー（ロゴ配置）
    if os.path.exists("logo.png"):
        st.image("logo.png", width=220)
        
    st.markdown("<h1 style='color: white;'>🚗 AI Suite Pro - 実務特化型コンクリート高精密診断システム</h1>", unsafe_allow_html=True)
    st.markdown("---")

    # 📘 取扱説明書（マニュアル）
    with st.expander("📘 本アプリの取扱説明書（マニュアル）を開く", expanded=False):
        st.markdown("""
        ### 【アプリの使い方】
        1. **実務情報の入力：** 画面左側の入力欄に、物件名や調査位置などの情報を入力してください。
        2. **プロ診断士用の環境条件選択：** 左側のサイドバーから、構造物の種類やセメントの種類、置かれている詳細な環境を設定します。
        3. **写真のアップロード＆AI診断：** 診断写真をアップロードし、「高精密AI解析を実行する」ボタンを押すと、すべての条件を考慮したプロレベルの診断書データと提出用Excelが作成されます。
        """)

    # 🔑 セキュリティエラーを完全に防止する「自動鍵読み込みの仕組み」に修正
    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        api_key = st.secrets.get("GEMINI_API_KEY", "")

    # 🛠️ プロのコンクリート診断士視点で必要な「環境・条件設定項目」
    st.sidebar.markdown("<h2 style='color: white;'>🛠️ プロ診断士用 環境条件設定</h2>", unsafe_allow_html=True)
    
    struct_type = st.sidebar.selectbox(
        "① 構造物の種類",
        ["橋梁（上部工/下部工）", "ボックスカルバート", "擁壁（重力式/もたれ式）", "トンネル覆工", "港湾・河川構造物", "建築物基礎・柱・壁"]
    )
    env_location = st.sidebar.selectbox(
        "② 設置環境・地域",
        ["一般地域（屋外・雨掛かり）", "一般地域（日陰・軒下）", "海岸付近（塩害地域）", "寒冷地（凍結融解の恐れ）", "温泉・工場地帯（酸性水・化学的腐食）", "屋内（常時乾燥）"]
    )
    wet_status = st.sidebar.selectbox(
        "③ コンクリートの湿潤状態",
        ["常時乾燥状態", "乾湿の繰り返し（最もひび割れが進展しやすい）", "常時湿潤状態（漏水・滞水あり）"]
    )
    cement_type = st.sidebar.selectbox(
        "④ 使用セメントの種類（推測でも可）",
        ["普通ポルトランドセメント", "高炉セメント（B種など）", "早強ポルトランドセメント", "不明"]
    )
    elapsed_years = st.sidebar.selectbox(
        "⑤ 供用年数（経過年数）",
        ["5年未満（初期欠陥の可能性）", "5年以上〜20年未満", "20年以上〜50年未満", "50年以上（高経年化）"]
    )
    crack_type = st.sidebar.selectbox(
        "⑥ 目視での主たる劣化症状",
        ["ひび割れ（単一・規則性）", "亀甲状のひび割れ（ASRなどの疑い）", "エフロレッセンス（白華）の析出伴う", "コンクリートの剥離・鉄筋露出（爆裂現象）", "漏水・遊離石灰を伴う錆汁"]
    )

    # 左右の2カラムレイアウト（スマホ・カーナビ風UI）
    col1, col2 = st.columns([1, 1])

    with col1:
        # 🏢 役所・コンサル・顧客提出用「実務書類情報入力欄」
        st.markdown("<h3 style='color: white;'>🏢 提出用 業務情報入力</h3>", unsafe_allow_html=True)
        project_name = st.text_input("項目A：物件名（工事名・業務名）", placeholder="〇〇高架橋修繕工事に伴う劣化調査")
        location_name = st.text_input("項目B：調査位置・測定箇所", placeholder="A1橋台 正面左側中央部")
        inspector_name = st.text_input("項目C：調査担当者（コンクリート診断士名）", placeholder="太田 正雄")
        
        st.markdown("---")
        st.markdown("<h3 style='color: white;'>📷 診断写真の選択</h3>", unsafe_allow_html=True)
        uploaded_file = st.file_uploader("ここにコンクリート構造物の写真をアップロードしてください", type=["jpg", "jpeg", "png"])
        
        if uploaded_file is not None:
            image = Image.open(uploaded_file)
            st.image(image, caption="診断対象のコンクリート写真", use_container_width=True)
            st.markdown("---")
            execute_analysis = st.button("🚀 この内容で高精密AI解析を実行する")

    with col2:
        st.markdown("<h3 style='color: white;'>📊 高精密診断レポート</h3>", unsafe_allow_html=True)
        
        if uploaded_file is not None and 'execute_analysis' in locals() and execute_analysis:
            if not api_key:
                st.error("管理画面のSecretsにAPIキーが保存されていません。")
            else:
                with st.spinner("🔍 プロの診断士AIが、環境条件と写真を総合的にマトリクス解析中..."):
                    try:
                        genai.configure(api_key=api_key)
                        model = genai.GenerativeModel('gemini-1.5-pro')
                        
                        prompt = f"""
                        あなたは日本コンクリート工学会認定の「コンクリート診断士」における、国内最高峰の有識者です。
                        以下の【実務書類情報】、【プロ診断士用環境条件】、および【写真】を重ね合わせ、役所や技術コンサルタントに提出できる、精密な診断を行ってください。
                        
                        【プロ診断士用環境条件】
                        ・構造物種別: {struct_type}
                        ・設置環境: {env_location}
                        ・湿潤状態: {wet_status}
                        
                        必ず以下の4つの項目を特定し、正確なJSONフォーマットのみで出力してください。
                        ひび割れ幅（width）とひび割れ長さ（length）は、実務上想定される最も現実的な数値を推測して算出してください。
                        
                        ```json
                        {{
                          "width": 0.15,
                          "length": 18.3,
                          "reason": "科学的根拠に基づく詳細な劣化原因の推測を記述",
                          "solution": "具体的な補修工法と、今後の点検計画に関する提案を記述"
                        }}
                        ```
                        余計な挨拶や説明文は絶対に省き、上記のJSONのみを返してください。
