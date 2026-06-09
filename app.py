import streamlit as st
import google.generativeai as genai
from PIL import Image
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.drawing.image import Image as ExcelImage
import io
import os
import re
from datetime import datetime

# --- 1. ページ設定とUI ---
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
.status-card { padding: 25px; background-color: #1E293B; border-radius: 16px; border-left: 8px solid #38BDF8; margin-bottom: 20px; border-top: 1px solid #334155; border-right: 1px solid #334155; border-bottom: 1px solid #334155; }
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
    
    api_key = st.secrets.get("GEMINI_API_KEY", "")
    
    # --- 3. サイドバー設定 ---
    st.sidebar.markdown("## 🛠️ プロ診断士用 条件設定")
    struct_type = st.sidebar.selectbox("① 構造物の種類", ["（未選択・写真から自動判定）", "橋梁（上部工/下部工）", "ボックスカルバート", "擁壁", "トンネル覆工", "港湾・河川構造物", "建築物基礎・柱・壁"])
    env_location = st.sidebar.multiselect("② 設置環境・大分類（複数選択可）", ["一般地域（屋外・雨掛かり）", "一般地域（日陰・軒下）", "塩害警戒地域（海岸付近）", "寒冷地・凍枯地域", "屋内（常時乾燥）"], default=[])
    wet_status = st.sidebar.multiselect("③ 湿潤状態（複数選択可）", ["常時乾燥状態", "乾湿の繰り返し（ひび割れ進展）", "常時湿潤状態（漏水・滞水）"], default=[])
    cement_type = st.sidebar.selectbox("④ 使用セメントの種類", ["（未選択）", "普通ポルトランドセメント", "高炉セメント（B種など）", "早強ポルトランドセメント", "不明"])
    elapsed_years = st.sidebar.selectbox("⑤ 供用年数（経過年数）", ["（未選択）", "5年未満（初期欠陥の可能性）", "5年以上〜20年未満", "20年以上〜50年未満", "50年以上（高経年化）"])
    crack_type = st.sidebar.selectbox("⑥ 目視での主たる劣化症状", ["（未選択・写真から自動判定）", "ひび割れ（単一・規則性）", "亀甲状のひび割れ（ASRなどの疑い）", "エフロレッセンス（白華）の析出伴う", "コンクリートの剥離・鉄筋露出（爆裂現象）", "漏水・遊離石灰を伴う錆汁"])
    
    st.sidebar.markdown("### 🌦️ 気象・地域特有の環境入力")
    region_info = st.sidebar.text_area("⑦ 地域・気象特記事項", placeholder="例: 冬季の凍結融解サイクルが多い地域、海岸から近く飛来塩分が多い等")

    # --- 4. メイン画面 ---
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("### 🏢 業務情報と補足")
        project_name = st.text_input("項目A：物件名（工事名・業務名）", placeholder="（例：塩竈清掃工場 躯体調査）")
        location_name = st.text_input("項目B：調査位置・測定箇所", placeholder="（例：沈殿池 南面壁）")
        inspector_name = st.text_input("項目C：調査担当者（コンクリート診断士名）", value="診断太郎", placeholder="（例：診断太郎）")
        
        st.markdown("### 🔧 人間による補足情報入力")
        cb_salt = st.checkbox("海岸線から2km以内（塩害）")
        cb_freeze = st.checkbox("寒冷地・凍結防止剤散布（凍害）")
        cb_wet = st.checkbox("常時湿潤・漏水（ASR・溶出）")
        cb_shear = st.checkbox("X状のせん断ひび割れ疑い")
        cb_janka = st.checkbox("ジャンカ・初期ひび割れの目視確認あり")
        cb_joint = st.checkbox("施工目地・コールドジョイント部")
        cb_cover = st.checkbox("設計かぶり厚の不足が疑われる・または既知")
        
        selected_factors = []
        if cb_salt: selected_factors.append("海岸線から2km以内（塩害リスク）")
        if cb_freeze: selected_factors.append("寒冷地・凍害リスク")
        if cb_wet: selected_factors.append("常時湿潤・漏水・ASRリスク")
        if cb_shear: selected_factors.append("地震等によるせん断応力の疑い（X状クラック）")
        if cb_janka: selected_factors.append("ジャンカ・初期ひび割れの目視確認あり")
        if cb_joint: selected_factors.append("施工目地・コールドジョイント部")
        if cb_cover: selected_factors.append("設計かぶり厚の不足が疑われる・または既知")
        human_factors_text = "、".join(selected_factors) if selected_factors else "特になし"

        st.markdown("### 📏 【重要】寸法の手動上書き指定")
        manual_width = st.number_input("実測ひび割れ幅 (mm)", min_value=0.0, step=0.05, value=0.0)
        manual_length = st.number_input("実測ひび割れ長さ (cm)", min_value=0.0, step=1.0, value=0.0)

        st.markdown("---")
        st.markdown("### 📸 現場写真アップロード（最大6枚）")
        uploaded_files = st.file_uploader("写真をアップロードしてください", type=["jpg", "jpeg", "png"], accept_multiple_files=True)
        
        images = []
        photo_comments = []
        
        if uploaded_files:
            if len(uploaded_files) > 6:
                st.warning("⚠️ 最初の6枚のみを処理します。")
                uploaded_files = uploaded_files[:6]
                
            for idx, file in enumerate(uploaded_files):
                img = Image.open(file)
                images.append(img)
                st.image(img, caption=f"Photo No.{idx+1}", width=250)
                comment = st.text_input(f"Photo No.{idx+1} の補足コメント", key=f"comment_{idx}")
                photo_comments.append(f"【Photo No.{idx+1}】: {comment if comment else '特記事項なし'}")
                
            execute_analysis = st.button("🚀 環境情報・各写真データを統合して高精密AI解析を実行")

    # --- 5. AI解析処理 ---
    with col2:
        st.markdown("### 📊 高精密診断レポート")
        if uploaded_files and 'execute_analysis' in locals() and execute_analysis:
            if not api_key:
                st.error("APIキーが設定されていません。")
            else:
                with st.spinner("🔍 熟練コンクリート診断士AI(Gemini 2.5)が解析中..."):
                    try:
                        genai.configure(api_key=api_key)
                        model = genai.GenerativeModel('gemini-2.5-flash')
                        
                        env_text = "、".join(env_location) if env_location else "指定なし"
                        wet_text = "、".join(wet_status) if wet_status else "指定なし"
                        photo_comments_text = "\n".join(photo_comments)
                        
                        dim_info = "手動指定なし（スケールが無ければ測定不可として逆質問してください）"
                        if manual_width > 0 or manual_length > 0:
                            dim_info = f"ひび割れ幅: {manual_width} mm, 長さ: {manual_length} cm"

                        prompt = f"コンクリート診断士として報告書を作ってください。構造物:{struct_type}, 環境:{env_text}, 湿潤:{wet_text}, セメント:{cement_type}, 年数:{elapsed_years}, 症状:{crack_type}, 地域:{region_info}, 補足:{human_factors_text}, 寸法:{dim_info}, 写真コメント:{photo_comments_text}。捏造は禁止しスケールがなければ逆質問してください。0.2mm未満は経過観察、以上は注入、1mm以上は充填工法。最初に「確定ひび割れ幅: 〇.〇 mm」と出力し、その後詳細を述べてください。"
                        
                        request_contents = [prompt] + images
                        response = model.generate_content(request_contents)
                        full_result_text = response.text
                        
                        final_width = manual_width
                        if final_width == 0:
                            try:
                                match = re.search(r"確定ひび割れ幅:\s*([0-9.]+)", full_result_text)
                                if match:
                                    final_width = float(match.group(1))
                            except Exception:
                                final_width = 0.0

                        if final_width >= 0.2:
                            color_code = "#EF4444"
                            status_title = f"🔴 【要精密補修】確定ひび割れ幅: {final_width} mm"
                            alert_desc = "⚠️ 判定基準：0.2mm以上のひび割れのため「注入工法」等の検討が必要です。"
                        elif final_width > 0:
                            color_code = "#EAB308"
                            status_title = f"🟡 【経過観察】確定ひび割れ幅: {final_width} mm"
                            alert_desc = "💡 判定基準：0.2mm未満のため、経過観察または表面含浸に該当します。"
                        else:
                            color_code = "#3B82F6"
                            status_title = "🔵 【寸法判定保留・逆質問あり】"
                            alert_desc = "ℹ️ スケールが不明なため、AIは数値を推測せず保留しています。実測値を確認してください。"
                        
                        st.markdown(f"<div class='status-card'><h3 style='color: {color_code} !important; margin:0; font-size:22px;'>{status_title}</h3><p style='color: #F1F5F9 !important; font-size: 14px; margin: 8px 0 0 0; font-weight: bold;'>{alert_desc}</p></div>", unsafe_allow_html=True)
                        st.markdown("<h4 style='color: white; margin-top:20px;'>📑 AI Suite Pro 統合解析レポート</h4>", unsafe_allow_html=True)
                        st.info(full_result_text)

                        # --- 6. Excel出力 ---
                        wb = openpyxl.Workbook()
                        ws = wb.active
                        ws.title = "調査状況写真台帳"
                        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
                        ws.page_setup.paperSize = ws.PAPERSIZE_A4
                        ws.views.sheetView[0].showGridLines = False

                        font_header = Font(name="MS ゴシック", size=14, bold=True)
                        font_label = Font(name="MS ゴシック", size=11, bold=True)
                        font_data = Font(name="MS ゴシック", size=11)
                        
                        thin_side = Side(border_style="thin", color="000000")
                        border_cell = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
                        fill_label = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

                        ws.column_dimensions['A'].width = 15
                        ws.column_dimensions['B'].width = 45
                        ws.column_dimensions['C'].width = 2
                        ws.column_dimensions['D'].width = 60
                        
                        p_name = project_name if project_name else "コンクリート構造物調査"
                        l_name = location_name if location_name else "現場写真"

                        start_row = 1
                        for idx, img in enumerate(images):
                            ws.merge_cells(f"A{start_row}:D{start_row}")
                            ws[f"A{start_row}"] = f"■ {p_name} 状況写真"
                            ws[f"A{start_row}"].font = font_header
                            ws.merge_cells(f"A{start_row+1}:D{start_row+1}")
                            ws[f"A{start_row+1}"] = f"位置： {l_name}"
                            ws[f"A{start_row+1}"].font = font_label
                            ws[f"A{start_row+1}"].alignment = Alignment(horizontal="right")

                            info_labels = ["写真No.", "撮影箇所", "工種", "位置", "AI判定・コメント"]
                            article_text = full_result_text if idx == 0 else photo_comments[idx]
                            info_values = [f"No.{idx+1}", f"現場写真 {idx+1}", "劣化調査", l_name, article_text]

                            for i, (label, value) in enumerate(zip(info_labels, info_values)):
                                r = start_row + 3 + i
                                ws[f"A{r}"] = label
                                ws[f"B{r}"] = value
                                ws[f"A{r}"].font = font_label
                                ws[f"A{r}"].fill = fill_label
                                ws[f"B{r}"].font = font_data
                                ws[f"A{r}"].border = border_cell
                                ws[f"B{r}"].border = border_cell
                                ws[f"A{r}"].alignment = Alignment(horizontal="center", vertical="center")
                                ws[f"B{r}"].alignment = Alignment(wrap_text=True, vertical="top")
                                
                                if label == "AI判定・コメント":
                                    ws.row_dimensions[r].height = 150
                                else:
                                    ws.row_dimensions[r].height = 24

                            img_buffer = io.BytesIO()
                            img.save(img_buffer, format="PNG")
                            img_buffer.seek(0)
                            xl_img = ExcelImage(img_buffer)
                            xl_img.width, xl_img.height = 420, 310
                            ws.add_image(xl_img, f"D{start_row + 3}")

                            start_row += 12

                        output = io.BytesIO()
                        wb.save(output)
                        
                        st.markdown("---")
                        st.download_button(
                            label="📥 Excel写真台帳をダウンロード",
                            data=output.getvalue(),
                            file_name=f"【写真台帳】{p_name}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    except Exception as e:
                        st.error(f"エラーが発生しました: {e}")
