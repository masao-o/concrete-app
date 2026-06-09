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
    env_location = st.sidebar.selectbox("② 設置環境・大分類", ["（未選択・写真から自動判定）", "一般地域（屋外・雨掛かり）", "一般地域（日陰・軒下）", "塩害警戒地域（海岸付近）", "寒冷地・凍枯地域", "屋内（常時乾燥）"])
    wet_status = st.sidebar.selectbox("③ コンクリートの湿潤状態", ["（未選択・写真から自動判定）", "常時乾燥状態", "乾湿の繰り返し（ひび割れが進展しやすい）", "常時湿潤状態（漏水・滞水あり）"])
    cement_type = st.sidebar.selectbox("④ 使用セメントの種類", ["（未選択）", "普通ポルトランドセメント", "高炉セメント（B種など）", "早強ポルトランドセメント", "不明"])
    elapsed_years = st.sidebar.selectbox("⑤ 供用年数（経過年数）", ["（未選択）", "5年未満（初期欠陥の可能性）", "5年以上〜20年未満", "20年以上〜50年未満", "50年以上（高経年化）"])
    crack_type = st.sidebar.selectbox("⑥ 目視での主たる劣化症状", ["（未選択・写真から自動判定）", "ひび割れ（単一・規則性）", "亀甲状のひび割れ（ASRなどの疑い）", "エフロレッセンス（白華）の析出伴う", "コンクリートの剥離・鉄筋露出（爆裂現象）", "漏水・遊離石灰を伴う錆汁"])

    with st.expander("📘 本アプリの取扱説明書（マニュアル）を開く", expanded=False):
        st.markdown("### 【アプリの使い方】\n1. 写真をアップロードし「高精密AI解析を実行する」ボタンを押すだけでAI診断が始まります。")

    col1, col2 = st.columns([1, 1])

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
                        
                        prompt = f"""
                        あなたは最高峰の「コンクリート診断士」です。国交省や大手コンサルに提出する公式な報告書を作成してください。
                        写真を詳細に調査し、環境（{env_location}）、湿潤状態（{wet_status}）、人為的補足（{human_factors_text}）を踏まえて、以下の2点をそれぞれ専門用語を交えた【300〜400文字以上の重厚なプロの意見】として詳しく日本語で述べてください。
                        
                        1. 【劣化原因に関する深い工学的推測】: 中性化、塩害、ASR、乾燥収縮、不同沈下などから、ひび割れの進展方向、エフロ、漏水、錆汁の有無を写真から読み解き、支配的な劣化メカニズムを不動態被膜や遊離石灰、膨張圧などの用語を用いて詳細に解説してください。
                        2. 【推奨される具体的な対策案・補修工法】: 土木学会等の指針に則り、エポキシ樹脂低圧注入工法、ポリマーセメントモルタル充填工法、表面含浸工法など具体的な工法名とその選定理由、さらにコア採取による追跡調査の必要性を明記してください。
                        
                        出力は必ず、最初に「推定ひび割れ幅: 0.18 mm / 推定ひび割れ長さ: 25.0 cm」のように実務上想定される、0.01mm〜0.5mmの間の現実的な幅と長さを推測して少数点で書いてください。
                        その後に【劣化原因の詳細】、【対策案の詳細】をそれぞれ独立した長文で出力してください。JSONなどの特殊な形式は使わないでください。
                        """
                        
                        response = model.generate_content([prompt, image])
                        full_result_text = response.text
                        
                        width_val = 0.18
                        length_val = 25.0
                        try:
                            if "幅:" in full_result_text:
                                part = full_result_text.split("幅:")[1].split("mm")[0].strip()
                                width_val = float(part)
                            if "長さ:" in full_result_text:
                                part2 = full_result_text.split("長さ:")[1].split("cm")[0].strip()
                                length_val = float(part2)
                        except:
                            pass

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

                        # Excelの作成
                        wb = openpyxl.Workbook()
                        ws = wb.active
                        ws.title = "コンクリート構造物劣化診断書"
                        
                        ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
                        ws.page_setup.paperSize = ws.PAPERSIZE_A4
                        ws.sheet_properties.pageSetUpPr.fitToPage = True
                        ws.page_setup.fitToWidth = 1  
                        ws.page_setup.fitToHeight = 0 
                        ws.views.sheetView[0].showGridLines = True

                        font_title = Font(name="MS ゴシック", size=16, bold=True, color="FFFFFF")
                        font_header = Font(name="MS ゴシック", size=11, bold=True, color="FFFFFF")
                        font_label = Font(name="MS ゴシック", size=10, bold=True, color="1E3A8A")
                        font_data = Font(name="MS ゴシック", size=10)
                        
                        fill_title = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
                        fill_header = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
                        fill_label = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
                        
                        thin_border_side = Side(border_style="thin", color="CBD5E1")
                        border_cell = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)

                        ws.column_dimensions['A'].width = 25
                        ws.column_dimensions['B'].width = 15
                        ws.column_dimensions['C'].width = 15
                        ws.column_dimensions['D'].width = 15
                        ws.column_dimensions['E'].width = 15
                        ws.column_dimensions['F'].width = 15
                        ws.column_dimensions['G'].width = 20

                        ws.merge_cells("A1:G1")
                        ws["A1"] = "コンクリート構造物 劣化診断報告書（実務提出用調書）"
                        ws["A1"].font = font_title
                        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
                        ws["A1"].fill = fill_title
                        ws.row_dimensions[1].height = 40

                        info_rows = [
                            ("物件名（工事名）", p_name, "■ 構造物種別", struct_type),
                            ("調査対象・位置", l_name, "■ 設置環境", env_location),
                            ("調査技術者（診断士）", i_name, "■ 乾湿状態", wet_status),
                            ("調査実施日", datetime.now().strftime("%Y年%m月%d日"), "■ 現場補足要因", human_factors_text)
                        ]
                        
                        for i, (l1, v1, l2, v2) in enumerate(info_rows, start=3):
                            ws.cell(row=i, column=1, value=l1).font = font_label
                            ws.cell(row=i, column=1).fill = fill_label
                            ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=4)
                            ws.cell(row=i, column=2, value=v1).font = font_data
                            ws.cell(row=i, column=2).alignment = Alignment(wrap_text=True, vertical="center")
                            
                            ws.cell(row=i, column=5, value=l2).font = font_label
                            ws.cell(row=i, column=5).fill = fill_label
                            ws.merge_cells(start_row=i, start_column=6, end_row=i, end_column=7)
                            ws.cell(row=i, column=6, value=v2).font = font_data
                            ws.cell(row=i, column=6).alignment = Alignment(wrap_text=True, vertical="center")
                            ws.row_dimensions[i].height = 25

                        ws.merge_cells("A8:G8")
                        ws["A8"] = "■ AI高精密解析・工学的診断判定データ"
                        ws["A8"].font = Font(name="MS ゴシック", size=11, bold=True, color="1E3A8A")
                        ws.row_dimensions[8].height = 25

                        ws["A9"] = "評価項目"
                        ws.merge_cells("B9:G9")
                        ws["B9"] = "コンクリート診断士AIによる抽出数値、および技術的所見レポート"
                        ws["A9"].font = font_header
                        ws["A9"].fill = fill_header
                        ws["B9"].font = font_header
                        ws["B9"].fill = fill_header
                        ws["A9"].alignment = Alignment(horizontal="center", vertical="center")
                        ws["B9"].alignment = Alignment(horizontal="center", vertical="center")
                        ws.row_dimensions[9].height = 25

                        ws["A10"] = "想定されるひび割れ幅"
                        ws.merge_cells("B10:G10")
                        ws["B10"] = f"{width_val} mm （{'赤色警告・要精密補修' if width_val>=0.2 else '黄色警告・経過観察' }）"
                        ws.row_dimensions[10].height = 24

                        ws["A11"] = "想定されるひび割れ長さ"
                        ws.merge_cells("B11:G11")
                        ws["B11"] = f"{length_val} cm"
                        ws.row_dimensions[11].height = 24

                        ws["A12"] = "AI詳細調査報告意見書\n(劣化原因・対策提案長文)"
                        ws.merge_cells("B12:G12")
                        ws["B12"] = full_result_text
                        ws["B12"].alignment = Alignment(wrap_text=True, vertical="top")
                        
                        text_length = len(full_result_text)
                        calculated_height = max(390, min(590, int(text_length * 0.46)))
                        ws.row_dimensions[12].height = calculated_height

                        for r in range(10, 13):
                            ws.cell(row=r, column=1).font = font_label
                            ws.cell(row=r, column=1).fill = fill_label
                            ws.cell(row=r, column=2).font = font_data

                        ws.merge_cells("A14:G14")
                        ws["A14"] = "■ 診断対象構造物・現場調査写真"
                        ws["A14"].font = Font(name="MS ゴシック", size=11, bold=True, color="1E3A8A")
                        ws.row_dimensions[14].height = 25

                        for row in ws.iter_rows(min_row=1, max_row=14, min_col=1, max_col=7):
                            for cell in row: cell.border = border_cell

                        # 【選択肢A】オリジナルの最高に綺麗な現場写真をそのままExcelへバシッと配置
                        img_buffer_xl = io.BytesIO()
                        image.save(img_buffer_xl, format="PNG")
                        img_buffer_xl.seek(0)
                        xl_img = ExcelImage(img_buffer_xl)
                        xl_img.width = 520
                        xl_img.height = 370
                        
                        ws.add_image(xl_img, "B16")
                        ws.row_dimensions[16].height = 390
                        
                        output = io.BytesIO()
                        wb.save(output)
                        st.markdown("---")
                        st.download_button(
                            label="📥 官庁・役所・提出用 Excel報告書をダウンロード", 
                            data=output.getvalue(), 
                            file_name=f"【確定劣化診断書】{project_name if project_name else 'コンクリート構造物'}.xlsx", 
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    except Exception as e:
                        st.error(f"解析中にエラーが発生しました: {e}")
        else:
            st.info("「この内容で高精密AI解析を実行する」ボタンを押すと、役所に提出可能なプロレベルの長文診断結果が表示されます。")
