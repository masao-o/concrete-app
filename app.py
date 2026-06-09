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

# --- 2. 高セキュリティ：パスワード認証 ---
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
        st.markdown("<h2 style='text-align: center; color: white;'>🔒 閉域環境・コンクリート劣化診断 AI Suite Pro</h2>", unsafe_allow_html=True)
        st.text_input("アクセスパスワード（担当者専用）", type="password", on_change=password_entered, key="password")
        return False
    return True

if check_password():
    st.markdown("<h1 style='color: white;'>🚗 AI Suite Pro - 実務特化型コンクリート高精密診断システム</h1>", unsafe_allow_html=True)
    st.markdown("---")
    
    # APIキーの安全な取得
    try:
        api_key = st.secrets["GEMINI_API_KEY"]
    except Exception:
        api_key = ""
    
    # --- 3. サイドバー：プロ診断士用 環境条件設定 ---
    st.sidebar.markdown("## 🛠️ プロ診断士用 条件設定")
    struct_type = st.sidebar.selectbox("① 構造物の種類", ["（未選択・写真から自動判定）", "橋梁（上部工/下部工）", "ボックスカルバート", "擁壁", "トンネル覆工", "港湾・河川構造物", "建築物基礎・柱・壁"])
    
    env_location = st.sidebar.multiselect("② 設置環境・大分類（複数選択可）", ["一般地域（屋外・雨掛かり）", "一般地域（日陰・軒下）", "塩害警戒地域（海岸付近）", "寒冷地・凍枯地域", "屋内（常時乾燥）"], default=[])
    wet_status = st.sidebar.multiselect("③ 湿潤状態（複数選択可）", ["常時乾燥状態", "乾湿の繰り返し（ひび割れ進展）", "常時湿潤状態（漏水・滞水）"], default=[])
    
    cement_type = st.sidebar.selectbox("④ 使用セメントの種類", ["（未選択）", "普通ポルトランドセメント", "高炉セメント（B種など）", "早強ポルトランドセメント", "不明"])
    elapsed_years = st.sidebar.selectbox("⑤ 供用年数（経過年数）", ["（未選択）", "5年未満（初期欠陥の可能性）", "5年以上〜20年未満", "20年以上〜50年未満", "50年以上（高経年化）"])
    crack_type = st.sidebar.selectbox("⑥ 目視での主たる劣化症状", ["（未選択・写真から自動判定）", "ひび割れ（単一・規則性）", "亀甲状のひび割れ（ASRなどの疑い）", "エフロレッセンス（白華）の析出伴う", "コンクリートの剥離・鉄筋露出（爆裂現象）", "漏水・遊離石灰を伴う錆汁"])
    
    st.sidebar.markdown("### 🌦️ 気象・地域特有の環境入力")
    region_info = st.sidebar.text_area("⑦ 地域・気象特記事項", placeholder="例: 冬季の凍結融解サイクルが多い地域、海岸から近く飛来塩分が多い等")

    # --- 4. メイン画面：業務情報・写真アップロード ---
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("### 🏢 業務情報と補足")
        project_name = st.text_input("項目A：物件名（工事名・業務名）", placeholder="（例：塩竈清掃工場 躯体調査）")
        location_name = st.text_input("項目B：調査位置・測定箇所", placeholder="（例：沈殿池 南面壁）")
        inspector_name = st.text_input("項目C：調査担当者（コンクリート診断士名）", value="診断太郎", placeholder="（例：診断太郎）")
        
        st.markdown("### 🔧 人間による補足情報入力（重複なし）")
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
        st.info("写真にクラックスケールが無い場合、AIのハルシネーションを防ぐため実測値を入力してください。未入力の場合、AIは勝手な推測をせずユーザーに質問します。")
        manual_width = st.number_input("実測ひび割れ幅 (mm)", min_value=0.0, step=0.05, value=0.0)
        manual_length = st.number_input("実測ひび割れ長さ (cm)", min_value=0.0, step=1.0, value=0.0)

        st.markdown("---")
        st.markdown("### 📸 現場写真アップロード（最大6枚）")
        uploaded_files = st.file_uploader("写真をアップロード", type=["jpg", "jpeg", "png"], accept_multiple_files=True)
        
        images = []
        photo_comments = []
        
        if uploaded_files:
            if len(uploaded_files) > 6:
                st.warning("⚠️ 写真は最大6枚までです。最初の6枚のみを処理します。")
                uploaded_files = uploaded_files[:6]
                
            for idx, file in enumerate(uploaded_files):
                img = Image.open(file)
                images.append(img)
                st.image(img, caption=f"Photo No.{idx+1}", width=250)
                comment = st.text_input(f"Photo No.{idx+1} の補足コメント（任意）", key=f"comment_{idx}")
                photo_comments.append(f"【Photo No.{idx+1}】: {comment if comment else '特記事項なし'}")
                
            execute_analysis = st.button("🚀 環境情報・各写真データを統合して高精密AI解析を実行")

    # --- 5. AI解析処理 ---
    with col2:
        st.markdown("### 📊 高精密診断レポート")
        if uploaded_files and 'execute_analysis' in locals() and execute_analysis:
            if not api_key:
                st.error("APIキーが設定されていません。StreamlitのSecrets設定を確認してください。")
            else:
                with st.spinner("🔍 熟練コンクリート診断士AI(Gemini)が過去の実績に基づき統合解析中..."):
                    try:
                        genai.configure(api_key=api_key)
                        # ★修正箇所：環境依存のエラーを回避するため「-latest」を付与したモデル名に変更
                        model = genai.GenerativeModel('gemini-1.5-flash-latest')
                        
                        env_text = "、".join(env_location) if env_location else "指定なし"
                        wet_text = "、".join(wet_status) if wet_status else "指定なし"
                        photo_comments_text = "\n".join(photo_comments)
                        
                        # 手動入力があればそれを適用
                        dim_info = "手動指定なし（写真内にスケールが無ければ測定不可として質問してください）"
                        if manual_width > 0 or manual_length > 0:
                            dim_info = f"【診断士による実測確定値】ひび割れ幅: {manual_width} mm, 長さ: {manual_length} cm"

                        # 超強力なプロンプトの構築
                        prompt = f"""
あなたは最高峰の「コンクリート診断士」です。官公庁や大手コンサルタントに提出する公式な報告書を作成してください。
以下の環境条件、入力情報、および各写真（Photo No.1〜）とそのコメントを踏まえて、テンプレートを排した完全オーダーメイドの重厚な工学的推測文（400文字以上）を出力してください。

【環境・現場入力情報】
- 構造物: {struct_type}
- 環境・湿潤: {env_text} / {wet_text}
- セメント種類: {cement_type}
- 供用年数: {elapsed_years}
- 主たる劣化症状: {crack_type}
- 気象・地域特有の環境: {region_info}
- 人為的補足: {human_factors_text}
- 寸法情報: {dim_info}
- 写真ごとのコメント:
{photo_comments_text}

【絶対厳守命令（ハルシネーション・寸法捏造の完全禁止）】
1. 手動指定寸法が無く、かつ写真内に「クラックスケール」や明確な寸法基準が確認できない場合、絶対に寸法を捏造しないでください。必ず文章の冒頭で「写真から正確な寸法を測定するための基準が確認できないため、勝手に判断せず保留します。正確な測定のために実測値または縮尺基準の提供を求めます」と回答・逆質問してください。
2. テンプレート回答を避け、塩害、凍害、中性化、ASR、不同沈下、せん断応力などの支配的メカニズムを、不動態被膜、遊離石灰、膨張圧などの専門用語を用いて、対象環境に合わせた推測をしてください。

【補修工法および追跡調査の選定基準】
・ひび割れ幅が0.2mm未満（または軽度）の場合は「劣化度Ⅰ」とし表面含浸工法等の予防保全や経過観察とする。
・0.2mm以上1.0mm未満の場合は「注入工法（低圧エポキシ樹脂注入など）」を提案する。
・1.0mm以上の場合は「充填工法（ポリマーセメントモルタル充填など）」を提案する。
・さらに、コンクリート内部の劣化度を確認するため、シュミットハンマーによる反発硬度試験、コア採取による圧縮強度試験、ドリル削孔による中性化深さ試験などの詳細調査の必要性を必ずプロの視点で明記すること。

出力は、確認できた（または手動入力された）「確定ひび割れ幅: 〇〇 mm」を冒頭に示し、その後【劣化原因の詳細（気象条件の考慮含む）】、【写真ごとの個別見解】、【対策案および詳細調査の推奨】を長文で出力してください。JSONは不要です。
"""
                        request_contents = [prompt] + images
                        response = model.generate_content(request_contents)
                        full_result_text = response.text
                        
                        # アラートカラーの判定（手動入力優先、なければ0）
                        final_width = manual_width
                        if final_width == 0:
                            try:
                                if "確定ひび割れ幅:" in full_result_text:
                                    val_str = full_result_text.split("確定ひび割れ幅:")[1].split("mm").strip()
                                    final_width = float(val_str)
                            except Exception:
                                final_width = 0.0

                        if final_width >= 0.2:
                            color_code = "#EF4444"
                            status_title = f"🔴 【要精密補修】確定/推定ひび割れ幅: {final_width} mm"
                            alert_desc = "⚠️ 判定基準：0.2mm以上のひび割れのため、指針に基づく「注入工法」等の検討が必要です。"
                        elif final_width > 0:
                            color_code = "#EAB308"
                            status_title = f"🟡 【経過観察】確定/推定ひび割れ幅: {final_width} mm"
                            alert_desc = "💡 判定基準：0.2mm未満のひび割れのため、表面含浸や経過観察（劣化度Ⅰ）に該当します。"
                        else:
                            color_code = "#3B82F6"
                            status_title = "🔵 【寸法未確定・質問あり】"
                            alert_desc = "ℹ️ スケールが不明、または寸法が入力されていないため、実測値の確認が求められています。"
                        
                        st.markdown(f"<div class='status-card'><h3 style='color: {color_code} !important; margin:0; font-size:22px;'>{status_title}</h3><p style='color: #F1F5F9 !important; font-size: 14px; margin: 8px 0 0 0; font-weight: bold;'>{alert_desc}</p></div>", unsafe_allow_html=True)
                        st.markdown("<h4 style='color: white; margin-top:20px;'>📑 AI Suite Pro 統合解析レポート</h4>", unsafe_allow_html=True)
                        st.info(full_result_text)

                        # --- 6. 複数枚対応 Excel出力（調査状況写真台帳フォーマット） ---
                        wb = openpyxl.Workbook()
                        ws = wb.active
                        ws.title = "調査状況写真台帳"
                        
                        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
                        ws.page_setup.paperSize = ws.PAPERSIZE_A4
                        ws.views.sheetView.showGridLines = False

                        font_header = Font(name="MS ゴシック", size=14, bold=True)
                        font_label = Font(name="MS ゴシック", size=11, bold=True)
                        font_data = Font(name="MS ゴシック", size=11)
                        border_cell = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))

                        ws.column_dimensions['A'].width = 15
                        ws.column_dimensions['B'].width = 45
                        ws.column_dimensions['C'].width = 2
                        ws.column_dimensions['D'].width = 60
                        
                        p_name = project_name if project_name else "コンクリート構造物躯体調査"
                        l_name = location_name if location_name else "現場写真"

                        start_row = 1
                        for idx, img in enumerate(images):
                            ws.merge_cells(f"A{start_row}:D{start_row}")
                            ws[f"A{start_row}"] = f"{p_name}　状 況 写 真"
                            ws[f"A{start_row}"].font = font_header
                            ws.merge_cells(f"A{start_row+1}:D{start_row+1}")
                            ws[f"A{start_row+1}"] = f"施設名： {l_name}"
                            ws[f"A{start_row+1}"].font = font_header
                            ws[f"A{start_row+1}"].alignment = Alignment(horizontal="right")

                            info_labels = ["写真No.", "撮影箇所", "工　種", "位　置", "記　事（AI見解）"]
                            article_text = full_result_text if idx == 0 else photo_comments[idx]
                            info_values = [str(idx + 1), f"現場撮影写真 {idx+1}", "劣化状況調査", l_name, article_text]

                            for i, (label, value) in enumerate(zip(info_labels, info_values)):
                                r = start_row + 3 + i
                                ws[f"A{r}"] = label
                                ws[f"B{r}"] = value
                                ws[f"A{r}"].font = font_label
                                ws[f"B{r}"].font = font_data
                                ws[f"A{r}"].border = border_cell
                                ws[f"B{r}"].border = border_cell
                                ws[f"A{r}"].alignment = Alignment(horizontal="center", vertical="center")
                                ws[f"B{r}"].alignment = Alignment(wrap_text=True, vertical="top")

                            ws.row_dimensions[start_row + 7].height = 200

                            img_buffer = io.BytesIO()
                            img.save(img_buffer, format="PNG")
                            img_buffer.seek(0)
                            xl_img = ExcelImage(img_buffer)
                            xl_img.width, xl_img.height = 420, 310
                            ws.add_image(xl_img, f"D{start_row + 3}")

                            start_row += 35 

                        output = io.BytesIO()
                        wb.save(output)
                        
                        st.markdown("---")
                        st.download_button(
                            label="📥 官庁・役所・提出用 Excel写真台帳をダウンロード",
                            data=output.getvalue(),
                            file_name=f"【調査状況写真】{project_name if project_name else 'コンクリート調査'}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    except Exception as e:
                        st.error(f"解析中にエラーが発生しました: {e}")
