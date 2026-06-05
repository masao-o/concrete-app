import streamlit as st
import google.generativeai as genai
from PIL import Image
import openpyxl
from openpyxl.drawing.image import Image as ExcelImage
import io

# ページ設定（スタイリッシュなダークテーマ風のUI）
st.set_page_config(page_title="T&N コンクリート劣化診断 AI Suite Pro", layout="wide")

# カスタムCSSでiPhone/カーナビ風のデザインを適用
st.markdown("""
    <style>
    .main { background-color: #121212; color: #E0E0E0; }
    h1, h2, h3 { color: #FFFFFF; font-family: 'Helvetica Neue', Arial, sans-serif; }
    .stButton>button {
        background-color: #1F1F1F; color: #00A3E0; 
        border: 2px solid #00A3E0; border-radius: 10px;
        padding: 10px 24px; font-weight: bold; width: 100%;
    }
    .stButton>button:hover { background-color: #00A3E0; color: #FFFFFF; }
    .report-box { padding: 20px; background-color: #1E1E1E; border-radius: 15px; border-left: 5px solid #00A3E0; }
    </style>
""", unsafe_allow_html=True)

# パスワードセッション管理
if 'authenticated' not in st.session_state:
    st.session_state['authenticated'] = False

def check_password():
    def password_entered():
        if st.session_state["password"] == "masao0605":  # 以前設定したパスワード
            st.session_state["authenticated"] = True
            del st.session_state["password"]
        else:
            st.st.error("パスワードが違います")
    
    if not st.session_state["authenticated"]:
        st.markdown("<h2 style='text-align: center;'>🔒 コンクリート劣化診断 AI Suite Pro</h2>", unsafe_allow_html=True)
        st.text_input("パスワードを入力してください", type="password", on_change=password_entered, key="password")
        return False
    return True

if check_password():
    # ホーム画面ヘッダー（ロゴ配置）
    try:
        logo = Image.open("logo.png")
        st.image(logo, width=200)
    except:
        pass
        
    st.title("🚗 AI Suite Pro - 高精密コンクリート診断システム")
    st.markdown("---")
    
    # 2カラムレイアウト（ナビ風）
    col1, col2 = st.columns([1, 2])
    
    with col1:
        st.markdown("### 📷 診断メニュー")
        uploaded_file = st.file_uploader("コンクリートの構造物写真をアップロードしてください", type=["jpg", "png", "jpeg"])
        
    with col2:
        if uploaded_file is not None:
            image = Image.open(uploaded_file)
            st.image(image, caption="アップロードされた写真", use_container_width=True)
            
            if st.button("🚀 高精密AI診断を実行"):
                with st.spinner("プロの診断士AIが解析中..."):
                    # 高精密プロンプト設定
                    model = genai.GenerativeModel('gemini-1.5-pro')
                    prompt = """
                    コンクリート診断士の専門視点で、写真のコンクリート劣化（ひび割れ、エフロレッセンス、剥離など）を高精密に解析してください。
                    特に「ひび割れ」がある場合は、以下の2つのデータ（数値のみ、または推測値）を必ず見つけ出し、指定の形式で出力してください。
                    
                    【出力形式】
                    ■ひび割れ幅: [数値]mm
                    ■ひび割れ長さ: [数値]cm
                    ■劣化原因の推測:
                    ■対策案:
                    """
                    
                    # AI解析ダミー（実際はAPIキー設定後に動作）
                    # 実務用にモックデータを配置して表示を確認可能にします
                    width_val = 0.15  # サンプル判定用
                    length_val = 12.5
                    
                    st.markdown("### 📊 AI高精密診断結果")
                    
                    # 幅による色判定ルールの適用 (0.2mm以下: 赤 / 0.2mm以上: 黄色)
                    if width_val <= 0.2:
                        color_code = "#FF4B4B"  # 赤
                        status_text = f"【警告】ひび割れ幅 {width_val}mm (0.2mm以下のため赤色表示)"
                    else:
                        color_code = "#FFEB3B"  # 黄色
                        status_text = f"【注意】ひび割れ幅 {width_val}mm (0.2mm以上のため黄色表示)"
                        
                    st.markdown(f"<h4 style='color:{color_code};'>{status_text}</h4>", unsafe_allow_html=True)
                    st.markdown(f"📐 **検出されたひびの長さ:** {length_val} cm")
                    
                    # Excel報告書作成
                    wb = openpyxl.Workbook()
                    ws = wb.active
                    ws.title = "劣化診断報告書"
                    
                    # 役所・コンサル提出用書式レイアウト
                    ws.merge_cells("A1:E1")
                    ws["A1"] = "コンクリート構造物 劣化診断報告書"
                    ws["A1"].font = openpyxl.styles.Font(size=16, bold=True)
                    
                    ws["A3"] = "調査会社"
                    ws["B3"] = "Ｔ＆日本メンテ開発株式会社"
                    ws["A4"] = "判定結果"
                    ws["B4"] = status_text
                    ws["A5"] = "ひび割れ長さ"
                    ws["B5"] = f"{length_val} cm"
                    
                    # Excelにロゴと写真を埋め込み
                    try:
                        ws.add_image(ExcelImage("logo.png"), "E3")
                    except:
                        pass
                        
                    output = io.BytesIO()
                    wb.save(output)
                    processed_data = output.getvalue()
                    
                    st.markdown("---")
                    st.download_button(
                        label="📥 役所・顧客提出用 Excel報告書をダウンロード",
                        data=processed_data,
                        file_name="コンクリート劣化診断報告書.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
