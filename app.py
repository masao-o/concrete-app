import streamlit as st
import google.generativeai as genai
from PIL import Image
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.drawing.image import Image as ExcelImage
import io
import os
import re
import time
from datetime import datetime

# --- 1. ページ設定とUIデザイン ---
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

.report-text-box {
    background-color: #1E293B !important;
    color: #FFFFFF !important;
    border: 1px solid #475569;
    border-radius: 12px;
    padding: 25px;
    font-size: 16px;
    line-height: 1.8;
    white-space: pre-wrap;
    font-family: 'Helvetica Neue', Arial, sans-serif;
}
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
    
    # --- 3. サイドバー設定（住所から自動環境解析モジュール） ---
    st.sidebar.markdown("## 📍 調査対象の所在地（自動解析用）")
    address_input = st.sidebar.text_input("構造物の設置住所・路線名を入力", placeholder="例：山形県酒田市大浜 国道112号")
    
    auto_freeze_info = "判定不能（住所未入力）"
    auto_salt_info = "判定不能（住所未入力）"
    auto_agent_info = "判定不能（住所未入力）"
    auto_weather_summary = "特記事項なし"
    
    if address_input:
        cold_regions = ["北海道", "青森", "岩手", "秋田", "山形", "宮城", "福島", "新潟", "富山", "石川", "福井", "長野", "岐阜", "群馬", "山梨"]
        is_cold = any(reg in address_input for reg in cold_regions)
        
        if is_cold:
            auto_freeze_info = "【激甚環境】過去数十年の気象統計より、冬季の凍結融解サイクルが年平均45〜60回以上発生する凍害ハイリスク地域。コンクリート内部の水分膨張圧によるスケーリングや微細組織の破壊リスクが極めて高い。"
            auto_agent_info = "【散布確率：大】自治体道路雪氷対策計画の重点路線エリア。冬季の路面凍結防止剤（塩化ナトリウム・塩化カルシウム）が大量に散布され、飛沫（しぶき）による中性化加速および二次的塩害を強く受ける環境。"
        else:
            auto_freeze_info = "【一般環境】温暖地域、または冬季の凍結融解サイクルが年平均15回未満の低凍害リスクエリア。"
            auto_agent_info = "【散布確率：小】定期的な凍結防止剤の大量散布路線には該当しない可能性が高い。"
            
        salt_keywords = ["浜", "海岸", "港", "湾", "岬", "磯", "シーサイド", "大浜", "臨海", "塩", "浦", "津"]
        is_coast = any(kw in address_input for kw in salt_keywords)
        
        if is_coast:
            auto_salt_info = "【重塩害警戒】海岸線からの直線距離が極めて近く（推定500m〜2km圏内）、通年の強風（日本海海上の季節風や台風等）により、高濃度の飛来塩分が構造物表面に定着しやすい環境。鉄筋の不動態被膜破壊が早期に進行するリスクがある。"
        else:
            auto_salt_info = "【一般地域】塩害警戒地域（沿岸近傍）の直接的な飛来塩分の影響は少ないエリア（内陸部）。"
            
        if is_cold and is_coast:
            auto_weather_summary = "日本海側あるいは過酷な沿岸寒冷地に位置。過去数十年の複合的な気象要因（乾湿の繰り返し、激しい寒暖差、塩風による化学的侵食、強風による微粒子の衝突に起因する物理的摩耗・風化）が重なる極めて厳しい環境。"
        elif is_cold:
            auto_weather_summary = "内陸寒冷地または山間部に位置。積雪による常時湿潤環境と乾燥の繰り返し、および凍結防止剤の塩化物侵食が支配的な要因。"
        else:
            auto_weather_summary = "比較的温暖な一般環境。ただし大気中の中性化（二酸化炭素）および降雨による溶出が主たる経年要因。"

    st.sidebar.markdown("---")
    st.sidebar.markdown("## 🛠️ プロ診断士用 条件設定")
    struct_type = st.sidebar.selectbox("① 構造物の種類", ["（未選択・写真から自動判定）", "橋梁（上部工/下部工）", "ボックスカルバート", "擁壁", "トンネル覆工", "港湾・河川構造物", "建築物基礎・柱・壁"])
    env_location = st.sidebar.multiselect("② 設置環境・大分類（複数選択可）", ["一般地域（屋外・雨掛かり）", "一般地域（日陰・軒下）", "塩害警戒地域（海岸付近）", "寒冷地・凍枯地域", "屋内（常時乾燥）"], default=[])
    wet_status = st.sidebar.multiselect("③ 湿潤状態（複数選択可）", ["常時乾燥状態", "乾湿の繰り返し（ひび割れ進展）", "常時湿潤状態（漏水・滞水）"], default=[])
    cement_type = st.sidebar.selectbox("④ 使用セメントの種類", ["（未選択）", "普通ポルトランドセメント", "高炉セメント（B種など）", "早強ポルトランドセメント", "不明"])
    elapsed_years = st.sidebar.selectbox("⑤ 供用年数（経過年数）", ["（未選択）", "5年未満（初期欠陥の可能性）", "5年以上〜20年未満", "20年以上〜50年未満", "50年以上（高経年化）"])
    crack_type = st.sidebar.selectbox("⑥ 目視での主たる劣化症状", ["（未選択・写真から自動判定）", "ひび割れ（単一・規則性）", "亀甲状のひび割れ（ASRなどの疑い）", "エフロレッセンス（白華）の析出伴う", "コンクリートの剥離・鉄筋露出（爆裂現象）", "漏水・遊離石灰を伴う錆汁"])
    
    st.sidebar.markdown("### 🌦️ 気象・地域特有の環境入力")
    region_info = st.sidebar.text_area("⑦ 地域・気象特記事項（手動補足用）", placeholder="例: 特に無し（上部の住所自動解析が優先されます）")

    # --- 4. メイン画面 ---
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("### 🏢 業務情報と補足")
        project_name = st.text_input("項目A：物件名（工事名・業務名）", placeholder="（例：塩竈清掃工場 躯体調査）")
        location_name = st.text_input("項目B：調査位置・測定箇所", placeholder="（例：沈殿池 南面壁）")
        inspector_name = st.text_input("項目C：調査担当者（コンクリート診断士名）", value="診断太郎", placeholder="（例：診断太郎）")
        
        st.markdown("### 🔧 人間による補足情報入力")
        cb_salt = st.checkbox("海岸線から2km以内（自動判定の補強）")
        cb_freeze = st.checkbox("寒冷地・融雪剤散布地域（自動判定の補強）")
        cb_wet = st.checkbox("常時湿潤・漏水（ASR・溶出リスク）")
        cb_shear = st.checkbox("X状のせん断ひび割れ疑い（構造要因）")
        cb_janka = st.checkbox("ジャンカ・初期ひび割れの目視確認あり")
        cb_joint = st.checkbox("施工目地・コールドジョイント部")
        cb_cover = st.checkbox("設計かぶり厚の不足が疑われる・または既知")
        
        selected_factors = []
        if cb_salt: selected_factors.append("海岸線から2km以内（塩害）")
        if cb_freeze: selected_factors.append("寒冷地・凍害・凍結防止剤散布地域")
        if cb_wet: selected_factors.append("常時湿潤・漏水・ASRリスクあり")
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
                with st.spinner("🔍 熟練コンクリート診断士AI(Gemini 2.5)が、所在地環境データを裏側で読み解きながら解析中..."):
                    
                    # 【改良点】429クォータエラーが出た場合に、自動で裏でリトライするロジック
                    full_result_text = None
                    max_retries = 3
                    
                    for attempt in range(max_retries):
                        try:
                            genai.configure(api_key=api_key)
                            model = genai.GenerativeModel('gemini-2.5-flash')
                            
                            env_text = "、".join(env_location) if env_location else "指定なし"
                            wet_text = "、".join(wet_status) if wet_status else "指定なし"
                            photo_comments_text = "\n".join(photo_comments)
                            
                            dim_info = "手動指定なし（写真内にスケールが無ければ測定不可として厳格に質問してください）"
                            if manual_width > 0 or manual_length > 0:
                                dim_info = f"【診断士実測確定値】ひび割れ幅: {manual_width} mm, 長さ: {manual_length} cm"

                            prompt = f"""
あなたは日本最高峰の「コンクリート診断士」です。国交省や大手建設コンサルタントに提出する公式な報告書を作成してください。
テンプレート回答や定型句を完全に排除し、他者の著作権を一切侵害しない完全オーダーメイドの重厚な工学的推測文（各項目400文字以上）を論理的に出力してください。

【対象構造物の所在地（システム自動算出の過酷環境データ）】
- 入力住所: {address_input if address_input else '未指定（写真より追測）'}
- 地域凍害・凍結融解リスク: {auto_freeze_info}
- 地域飛来塩分リスク: {auto_salt_info}
- 融雪剤（凍結防止剤）影響度: {auto_agent_info}
- 気象・摩耗・風化サマリー: {auto_weather_summary}

【手動条件設定項目】
- 構造物の種類: {struct_type}
- 設置環境・湿潤状態: {env_text} / {wet_text}
- 使用セメント・経過年数: {cement_type} / {elapsed_years}
- 主たる劣化症状: {crack_type}
- 人為的補足要因: {human_factors_text}
- 寸法情報に関する条件: {dim_info}
- 写真ごとのコメント:
{photo_comments_text}

【絶対厳守命令（ハルシネーション・寸法捏造の完全禁止）】
1. 手動指定寸法が無く、かつ写真内に「クラックスケール」や明確な寸法基準が確認できない場合、絶対に寸法を捏造しないでください。必ず文章の冒頭で「写真から正確な寸法を測定するための基準が確認できないため、勝手に判断せず保留します。正確な測定のために実測値または縮尺基準の提供を求めます」と逆質問・要請してください。
2. 上記の【所在地（システム自動算出の過酷環境データ）】を深く考察に組み込み、塩害、凍害、中性化、アルカリ骨材反応（ASR）、不同沈下、構造要因、風化摩耗などの支配的劣化メカニズムを、不動態被膜の破壊、遊離石灰、膨張圧、微細クラックの進展方向などの専門用語を用いて、この現場特有の深い推測をしてください。

【補修工法および追跡詳細調査の選定基準】
・ひび割れ幅が0.2mm未満の場合は「劣化度Ⅰ」とし表面含浸工法等の予防保全や経過観察とする。
・0.2mm以上1.0mm未満の場合は「注入工法（低圧エポキシ樹脂注入など）」を提案する。
・1.0mm以上の場合は「充填工法（ポリマーセメントモルタル充填など）」を提案する。
・コンクリート内部の実際の劣化度を確認するため、シュミットハンマーによる反発硬度試験、コア採取による圧縮強度試験、ドリル削孔による中性化深さ試験などの詳細な追跡調査の必要性を必ずプロの視点で明記すること。

出力は、確認できた場合のみ「確定ひび割れ幅: 〇.〇 mm」を必ず冒頭の1行目に示し、その後【劣化原因の詳細（気象条件および自動算出した地域要因含む）】、【写真ごとの個別見解】、【対策案および詳細調査の推奨】を、役所に提出できるレベルの非常に濃い長文で出力してください。JSONは不要です。
"""
                            request_contents = [prompt] + images
                            response = model.generate_content(request_contents)
                            full_result_text = response.text
                            break # 成功したらループを抜ける
                            
                        except Exception as e:
                            # 429エラーの場合は2.5秒待って再試行
                            if "429" in str(e) or "quota" in str(e).lower():
                                if attempt < max_retries - 1:
                                    time.sleep(2.5)
                                    continue
                            # それ以外のエラーはそのまま上に投げる
                            raise e
                    
                    # 最終的な判定・表示処理
                    if full_result_text:
                        try:
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
                                alert_desc = "⚠️ 判定基準：0.2mm以上のひび割れのため、指針に基づく「注入・充填工法」の設計検討が必要です。"
                            elif final_width > 0:
                                color_code = "#EAB308"
                                status_title = f"🟡 【経過観察】確定ひび割れ幅: {final_width} mm"
                                alert_desc = "💡 判定基準：0.2mm未満のため、表面含浸工法による予防保全または経過観察に該当します。"
                            else:
                                color_code = "#3B82F6"
                                status_title = "🔵 【寸法判定保留・逆質問あり】"
                                alert_desc = "ℹ️ スケールが不明なため数値を推測せず保留しています。実測値または縮尺基準を確認してください。"
                            
                            st.markdown(f"<div class='status-card'><h3 style='color: {color_code} !important; margin:0; font-size:22px;'>{status_title}</h3><p style='color: #F1F5F9 !important; font-size: 14px; margin: 8px 0 0 0; font-weight: bold;'>{alert_desc}</p></div>", unsafe_allow_html=True)
                            st.markdown("<h4 style='color: white; margin-top:20px;'>📑 AI Suite Pro 高精密統合解析レポート（気象データ連動）</h4>", unsafe_allow_html=True)
                            st.markdown(f"<div class='report-text-box'>{full_result_text}</div>", unsafe_allow_html=True)

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

                            ws.column_dimensions['A'].width = 18
                            ws.column_dimensions['B'].width = 52  
                            ws.column_dimensions['C'].width = 2
                            ws.column_dimensions['D'].width = 68  
                            
                            p_name = project_name if project_name else "コンクリート構造物劣化状況調査業務"
                            l_name = location_name if location_name else "現場調査測定箇所"
                            if address_input:
                                l_name = f"{address_input} ({l_name})"

                            start_row = 1
                            for idx, img in enumerate(images):
                                ws.merge_cells(f"A{start_row}:D{start_row}")
                                ws[f"A{start_row}"] = f"■ {p_name} 状況写真台帳"
                                ws[f"A{start_row}"].font = font_header
                                ws.merge_cells(f"A{start_row+1}:D{start_row+1}")
                                ws[f"A{start_row+1}"] = f"施設・位置： {l_name} (No.{idx+1})"
                                ws[f"A{start_row+1}"].font = font_label
                                ws[f"A{start_row+1}"].alignment = Alignment(horizontal="right")

                                info_labels = ["写真No.", "撮影箇所", "工種・項目", "位置・部材", "AI工学所見・記事"]
                                article_text = full_result_text if idx == 0 else photo_comments[idx]
                                info_values = [f"Photo No.{idx+1}", f"構造物近景劣化状況写真 ({idx+1})", "コンクリート劣化度目視調査", l_name, article_text]

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
                                    
                                    if label == "AI工学所見・記事":
                                        text_len = len(str(value))
                                        dynamic_height = max(180, min(450, int(text_len * 0.45)))
                                        ws.row_dimensions[r].height = dynamic_height
                                    else:
                                        ws.row_dimensions[r].height = 26

                                img_buffer = io.BytesIO()
                                img.save(img_buffer, format="PNG")
                                img_buffer.seek(0)
                                xl_img = ExcelImage(img_buffer)
                                xl_img.width, xl_img.height = 450, 320
                                ws.add_image(xl_img, f"D{start_row + 3}")

                                start_row += 12

                            output = io.BytesIO()
                            wb.save(output)
                            
                            st.markdown("---")
                            st.download_button(
                                label="📥 官庁・役所・提出用 高精密Excel写真台帳をダウンロード",
                                data=output.getvalue(),
                                file_name=f"【確定写真台帳】{p_name}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                        except Exception as excel_err:
                            st.error(f"Excelの生成中にエラーが発生しました: {excel_err}")
                    else:
                        st.error("⚠️ アクセスが非常に集中しているため、AIからの応答を一時的に受信できませんでした。2〜3分置いて再度お試しいただくか、Googleの有料プラン（秒間アクセス上限緩和）への移行をご検討ください。")
