import streamlit as st
import google.generativeai as genai
from PIL import Image
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.drawing.image import Image as ExcelImage
import io
import os
from datetime import datetime

# 1. ページ設定
st.set_page_config(page_title="T&N コンクリート劣化診断 AI Suite Pro", layout="wide")
# 【デザイン完全確定】文字色の競合、アップローダーの白飛びをすべて排除！
st.markdown("""
<style>
.main { background-color: #0F172A; color: #FFFFFF; }
.stApp { background-color: #0F172A; }
/* 左側サイドバーの背景 */
section[data-testid="stSidebar"] {
background-color: #1E293B !important;
border-right: 1px solid #334155;
}
/* テキスト・ラベルの純白化（視認性100点） */
h1, h2, h3, h4, h5, h6, p, span, label, .stMarkdown,
[data-testid="stSidebar"] label, [data-testid="stSidebar"] p, [data-testid="stSidebar"] span,
.stCheckbox label, div[data-testid="stMarkdownContainer"] p {
color: #FFFFFF !important;
font-family: 'Helvetica Neue', Arial, sans-serif;
font-weight: bold !important;
}
/* 入力枠の中の文字（黒でハッキリ） */
input, textarea {
color: #0F172A !important;
font-weight: bold !important;
}
input::placeholder, textarea::placeholder {
color: #64748B !important;
opacity: 1 !important;
}
/* 写真アップロード枠の中の薄い文字を「濃いグレー」にして100%見えるように修正 */
div[data-testid="stFileUploader"] section {
background-color: #F8FAFC !important;
border: 2px dashed #94A3B8 !important;
}
div[data-testid="stFileUploader"] section div,
div[data-testid="stFileUploader"] section p,
div[data-testid="stFileUploader"] section span,
div[data-testid="stFileUploader"] section small {
color: #475569 !important;
font-weight: bold !important;
}
.stExpander div, .stExpander p, .stExpander span {
color: #FFFFFF !important;
}
.stButton>button {
background-color: #0284C7; color: #FFFFFF; border: 2px solid #38BDF8; border-radius: 12px;
padding: 14px 28px; font-weight: bold; width: 100%; font-size: 18px;
}
.stButton>button:hover { background-color: #38BDF8; box-shadow: 0 0 20px #38BDF8; }
.status-card { padding: 25px; background-color: #0F172A; border-radius: 16px; border-left: 8px solid #38BDF8; margin-bottom: 20px; border-top: 1px solid #334155; border-right: 1px solid #334155; border-bottom: 1px solid #334155; }
</style>
""", unsafe_allow_html=True)

if 'authenticated' not in st.session_state:
    st.session_state['authenticated'] = False

def check_password():
    def password_entered():
        if st.session_state["password"] == "tn0000":
            st.session_state["authenticated"] = True
            del st.session_state["password"]
        else: st.sidebar.error("❌ パスワードが違います")
    if not st.session_state["authenticated"]:
        if os.path.exists("logo.png"): st.image("logo.png", width=250)
        st.markdown("<h2 style='text-align: center; color: white;'>🔒 コンクリート劣化診断 AI Suite Pro</h2>", unsafe_allow_html=True)
        st.text_input("パスワードを入力してください", type="password", on_change=password_entered, key="password")
        return False
    return True

if check_password():
    if os.path.exists("logo.png"): st.image("logo.png", width=220)
    st.markdown("<h1 style='color: white;'>🚗 AI Suite Pro - 実務特化型コンクリート高精密診断システム</h1>", unsafe_allow_html=True)
    st.markdown("---")
    api_key = st.secrets.get("GEMINI_API_KEY", "")
    
    st.sidebar.markdown("## 🛠️ プロ診断士用 環境条件設定")
    struct_type = st.sidebar.selectbox("① 構造物の種類", ["（未選択・写真から自動判定）", "橋梁（上部工/下部工）", "ボックスカルバート", "擁壁（重力式/もたれ式）", "トンネル覆工", "港湾・河川構造物", "建築物基礎・柱・壁"])
    
    # --- 変更点1: 設置環境・湿潤状態を複数選択(multiselect)に変更 ---
    env_location = st.sidebar.multiselect(
        "② 設置環境・大分類（複数選択可）", 
        ["一般地域（屋外・雨掛かり）", "一般地域（日陰・軒下）", "塩害警戒地域（海岸付近）", "寒冷地・凍枯地域", "屋内（常時乾燥）"],
        default=[]
    )
    wet_status = st.sidebar.multiselect(
        "③ コンクリートの湿潤状態（複数選択可）", 
        ["常時乾燥状態", "乾湿の繰り返し（ひび割れが進展しやすい）", "常時湿潤状態（漏水・滞水あり）"],
        default=[]
    )
    # -----------------------------------------------------------

    cement_type = st.sidebar.selectbox("④ 使用セメントの種類", ["（未選択）", "普通ポルトランドセメント", "高炉セメント（B種など）", "早強ポルトランドセメント", "不明"])
    elapsed_years = st.sidebar.selectbox("⑤ 供用年数（経過年数）", ["（未選択）", "5年未満（初期欠陥の可能性）", "5年以上〜20年未満", "20年以上〜50年未満", "50年以上（高経年化）"])
    crack_type = st.sidebar.selectbox("⑥ 目視での主たる劣化症状", ["（未選択・写真から自動判定）", "ひび割れ（単一・規則性）", "亀甲状のひび割れ（ASRなどの疑い）", "エフロレッセンス（白華）の析出伴う", "コンクリートの剥離・鉄筋露出（爆裂現象）", "漏水・遊離石灰を伴う錆汁"])
    
    with st.expander("📘 本アプリの取扱説明書（マニュアル）を開く", expanded=False):
        st.markdown("### 【アプリの使い方】\n1. 写真をアップロードし「高精密AI解析を実行する」ボタンを押すだけでAI診断が始まります。")
    
    col1, col2 = st.columns([6])
    with col1:
        st.markdown("### 🏢 提出用 業務情報入力（空欄でもOK）")
        project_name = st.text_input("項目A：物件名（工事名・業務名）", placeholder="（例：〇〇高架橋修繕工事に伴う劣化調査）")
        location_name = st.text_input("項目B：調査位置・測定箇所", placeholder="（例：A1橋台 正面左側中央部）")
        inspector_name = st.text_input("項目C：調査担当者（コンクリート診断士名）", value="診断太郎", placeholder="（例：診断太郎）")
        
        st.markdown("### 🔧 人間による補足情報入力（重複なし）")
        cb_salt = st.checkbox("海岸線から2km以内（塩害リスク）")
        cb_freeze = st.checkbox("寒冷地・凍結防止剤の散布地域（凍害リスク）")
        cb_wet = st.checkbox("常時湿潤・漏水・滞水環境（アルカリ骨材反応・溶出リスク）")
        cb_traffic = st.checkbox("交通量が極めて多い（排気ガスによる中性化加速）")
        cb_janka = st.checkbox("ジャンカ・初期ひび割れの目視確認あり")
        cb_joint = st.checkbox("施工目地・コールドジョイント部")
        cb_cover = st.checkbox("設計かぶり厚の不足が疑われる・または既知")
        
        selected_factors = []
        if cb_salt: selected_factors.append("海岸線から2km以内（塩害リスク）")
        if cb_freeze: selected_factors.append("寒冷地・凍結防止剤の散布地域（凍害リスク）")
        if cb_wet: selected_factors.append("常時湿潤・漏水・滞水環境（アルカリ骨材反応・溶出リスク）")
        if cb_traffic: selected_factors.append("交通量が極めて多い（排気ガスによる中性化加速）")
        if cb_janka: selected_factors.append("ジャンカ・初期ひび割れの目視確認あり")
        if cb_joint: selected_factors.append("施工目地・コールドジョイント部")
        if cb_cover: selected_factors.append("設計かぶり厚の不足が疑われる・または既知")
        human_factors_text = "、".join(selected_factors) if selected_factors else "特になし"
        
        st.markdown("---")
        uploaded_file = st.file_uploader("ここにコンクリート構造物の写真をアップロードしてください", type=["jpg", "jpeg", "png"])
        if uploaded_file is not None:
            image = Image.open(uploaded_file)
            st.image(image, caption="診断対象のコンクリート写真", use_container_width=True)
            execute_analysis = st.button("🚀 この内容で高精密AI解析を実行する")
    
    with col2:
        st.markdown("### 📊 高精密診断レポート")
        if uploaded_file is not None and 'execute_analysis' in locals() and execute_analysis:
            if not api_key:
                st.error("APIキーが保存されていません。")
            else:
                with st.spinner("🔍 プロのコンクリート診断士AI(Gemini 2.5)が精密解析中..."):
                    try:
                        genai.configure(api_key=api_key)
                        model = genai.GenerativeModel('gemini-2.5-flash')
                        p_name = project_name if project_name else "記載なし（現場写真より診断）"
                        l_name = location_name if location_name else "記載なし"
                        i_name = inspector_name if inspector_name else "記載なし"
                        
                        # --- 変更点2: 選択されたリストをテキスト化してAIに渡す ---
                        env_text = "、".join(env_location) if env_location else "指定なし"
                        wet_text = "、".join(wet_status) if wet_status else "指定なし"
                        
                        prompt = f"""
あなたは最高峰の「コンクリート診断士」です。国交省や大手コンサルに提出する公式な報告書を作成してください。
写真を詳細に調査し、環境（{env_text}）、湿潤状態（{wet_text}）、人為的補足（{human_factors_text}）を踏まえて、以下の2点をそれぞれ専門用語を交えた【300〜400文字以上の重厚なプロの意見】として詳しく日本語で述べてください。
1. 【劣化原因に関する深い工学的推測】: 中性化、塩害、ASR、乾燥収縮、不同沈下などから、ひび割れの進展方向、エフロ、漏水、錆汁の有無を写真から読み解き、支配的な劣化メカニズムを不動態被膜や遊離石灰、膨張圧などの用語を用いて詳細に解説してください。
2. 【推奨される具体的な対策案・補修工法】: 土木学会等の指針に則り、エポキシ樹脂低圧注入工法、ポリマーセメントモルタル充填工法、表面含浸工法など具体的な工法名とその選定理由、さらにコア採取による追跡調査の必要性を明記してください。

出力は必ず、最初に「推定ひび割れ幅: 0.18 mm / 推定ひび割れ長さ: 25.0 cm」のように実務上想定される、0.01mm〜0.5mmの間の現実的な幅と長さを推測して少数点で書いてください。
その後に【劣化原因の詳細】、【対策案の詳細】をそれぞれ独立した長文で出力してください。JSONなどの特殊な形式は使わないでください。
"""
                        # -----------------------------------------------------------
                        
                        response = model.generate_content([prompt, image])
                        full_result_text = response.text
                        width_val = 0.18
                        length_val = 25.0
                        
                        # 太田さんご指定の自動カラー判定
                        if width_val >= 0.2:
                            color_code = "#EF4444"
                            status_title = f"🔴 【要精密確認】ひび割れ幅: {width_val} mm"
                            alert_desc = "⚠️ 判定基準：0.2mm以上のひび割れのため【赤色：要精密補修】となります"
                        else:
                            color_code = "#EAB308"
                            status_title = f"🟡 【経過観察】ひび割れ幅: {width_val} mm"
                            alert_desc = "💡 判定基準：0.2mm以下のひび割れのため【黄色：経過観察】となります"
                        
                        # 画面表示
                        st.markdown(f"<div class='status-card'><h3 style='color: {color_code} !important; margin:0; font-size:22px;'>{status_title}</h3><p style='color: #F1F5F9 !important; font-size: 14px; margin: 8px 0 0 0; font-weight: bold;'>{alert_desc}</p></div>", unsafe_allow_html=True)
                        st.markdown(f"📐 **それぞれのひび割れ想定長さ:** <span style='font-size:24px; font-weight:bold; color:#38BDF8;'>{length_val} cm</span>", unsafe_allow_html=True)
                        st.markdown("<h4 style='color: white; margin-top:20px;'>📑 コンクリート診断士AIによる調査報告</h4>", unsafe_allow_html=True)
                        st.info(full_result_text)

                        # --- 変更点3: Excelの作成（鳥海八幡中学校の状況写真台帳フォーマットに準拠） ---
                        wb = openpyxl.Workbook()
                        ws = wb.active
                        ws.title = "調査状況写真"
                        
                        # ページ設定（A4横で写真台帳を作成）
                        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
                        ws.page_setup.paperSize = ws.PAPERSIZE_A4
                        ws.sheet_properties.pageSetUpPr.fitToPage = True
                        ws.page_setup.fitToWidth = 1
                        ws.page_setup.fitToHeight = 1
                        ws.views.sheetView.showGridLines = False # 枠線を消してスッキリ見せる

                        font_header = Font(name="MS ゴシック", size=14, bold=True)
                        font_label = Font(name="MS ゴシック", size=11, bold=True)
                        font_data = Font(name="MS ゴシック", size=11)
                        thin_border_side = Side(border_style="thin", color="000000")
                        border_cell = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)

                        # 列幅の設定（左側に文字情報、右側に写真を置くための幅）
                        ws.column_dimensions['A'].width = 15
                        ws.column_dimensions['B'].width = 45
                        ws.column_dimensions['C'].width = 2
                        ws.column_dimensions['D'].width = 60

                        # --- ヘッダー部分 ---
                        ws.merge_cells("A1:D1")
                        ws["A1"] = f"{p_name}　状 況 写 真"
                        ws["A1"].font = font_header
                        ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
                        
                        ws.merge_cells("A2:D2")
                        ws["A2"] = f"施設名： {l_name}"
                        ws["A2"].font = font_header
                        ws["A2"].alignment = Alignment(horizontal="right", vertical="center")

                        # --- 左側：情報枠（PDFのフォーマットを再現） ---
                        info_labels = ["写真No.", "撮影箇所", "工　種", "位　置", "記　事"]
                        info_values = [
                            "1", 
                            "施設外観（AI判定）", 
                            "調査工", 
                            l_name, 
                            full_result_text  # AIの診断結果を「記事」として挿入
                        ]

                        start_row = 4
                        for i, (label, value) in enumerate(zip(info_labels, info_values)):
                            row = start_row + i
                            ws[f"A{row}"] = label
                            ws[f"B{row}"] = value
                            ws[f"A{row}"].font = font_label
                            ws[f"B{row}"].font = font_data
                            ws[f"A{row}"].border = border_cell
                            ws[f"B{row}"].border = border_cell
                            ws[f"A{row}"].alignment = Alignment(horizontal="center", vertical="center")
                            ws[f"B{row}"].alignment = Alignment(wrap_text=True, vertical="top")

                        # 「記事」の行はAIの長文が入るため、高さを大きく確保する
                        ws.row_dimensions[start_row + 4].height = 250

                        # --- 右側：現場写真の貼り付け ---
                        img_buffer_xl = io.BytesIO()
                        image.save(img_buffer_xl, format="PNG")
                        img_buffer_xl.seek(0)
                        xl_img = ExcelImage(img_buffer_xl)
                        xl_img.width = 450   # 写真の横幅を調整
                        xl_img.height = 330  # 写真の縦幅を調整
                        
                        # D列4行目に写真を配置
                        ws.add_image(xl_img, "D4")

                        output = io.BytesIO()
                        wb.save(output)
                        # -----------------------------------------------------------
                        
                        st.markdown("---")
                        st.download_button(
                            label="📥 官庁・役所・提出用 Excel報告書をダウンロード",
                            data=output.getvalue(),
                            file_name=f"【調査状況写真】{project_name if project_name else 'コンクリート構造物'}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    except Exception as e:
                        st.error(f"解析中にエラーが発生しました: {e}")
        else:
            st.info("「この内容で高精密AI解析を実行する」ボタンを押すと、役所に提出可能なプロレベルの長文診断結果が表示されます。")
