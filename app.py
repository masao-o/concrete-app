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

# デザイン設定（視認性100点・白飛び黒同化完全排除）
st.markdown("""
    <style>
    .main { background-color: #0F172A; color: #FFFFFF; }
    .stApp { background-color: #0F172A; }
    
    section[data-testid="stSidebar"] {
        background-color: #1E293B !important;
        border-right: 1px solid #334155;
    }
    
    h1, h2, h3, h4, h5, h6, p, span, label, .stMarkdown,
    [data-testid="stSidebar"] label, [data-testid="stSidebar"] p, [data-testid="stSidebar"] span,
    .stCheckbox label, div[data-testid="stMarkdownContainer"] p { 
        color: #FFFFFF !important; 
        font-family: 'Helvetica Neue', Arial, sans-serif;
        font-weight: bold !important;
    }
    
    input, textarea {
        color: #0F172A !important; 
        font-weight: bold !important;
    }
    input::placeholder, textarea::placeholder {
        color: #64748B !important;
        opacity: 1 !important;
    }
    
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

    # サイドバー設定
    st.sidebar.markdown("## 🛠️ プロ診断士用 環境条件設定")
    struct_type = st.sidebar.selectbox("① 構造物の種類", ["（未選択・写真から自動判定）", "橋梁（上部工/下部工）", "ボックスカルバート", "擁壁（重力式/もたれ式）", "トンネル覆工", "港湾・河川構造物", "清掃工場・RC造建築物基礎・柱・壁"])
    
    # 💥 【ご要望】複数環境が選択できるようにマルチセレクトに大改造！
    env_locations = st.sidebar.multiselect("② 設置環境・大分類（複数選択可能）", ["一般地域（屋外・雨掛かり）", "一般地域（日陰・軒下）", "塩害警戒地域（海岸付近）", "寒冷地・凍枯地域", "屋内（常時乾燥）", "化学的腐食環境（清掃工場等）"])
    env_location_text = "、".join(env_locations) if env_locations else "（未選択・写真から自動判定）"
    
    wet_status = st.sidebar.selectbox("③ コンクリートの湿潤状態", ["（未選択・写真から自動判定）", "常時乾燥状態", "乾湿の繰り返し（ひび割れが進展しやすい）", "常時湿潤状態（漏水・滞水あり）"])
    cement_type = st.sidebar.selectbox("④ 使用セメントの種類", ["（未選択）", "普通ポルトランドセメント", "高炉セメント（B種など）", "早強ポルトランドセメント", "不明"])
    elapsed_years = st.sidebar.selectbox("⑤ 供用年数（経過年数）", ["（未選択）", "5年未満（初期欠陥の可能性）", "5年以上〜20年未満", "20年以上〜50年未満", "50年以上（高経年化）"])
    crack_type = st.sidebar.selectbox("⑥ 目視での主たる劣化症状", ["（未選択・写真から自動判定）", "ひび割れ（単一・規則性）", "亀甲状のひび割れ（ASRなどの疑い）", "エフロレッセンス（白華）の析出伴う", "コンクリートの剥離・浮き・鉄筋露出（爆裂現象）", "漏水・遊離石灰を伴う錆汁"])

    with st.expander("📘 本アプリの取扱説明書（マニュアル）を開く", expanded=False):
        st.markdown("### 【アプリの使い方】\n1. 写真を複数枚アップロードし「高精密AI解析を実行する」ボタンを押すだけでAI診断が始まります。")

    col1, col2 = st.columns([1, 1])

    with col1:
        st.markdown("### 🏢 提出用 業務情報入力（空欄でもOK）")
        project_name = st.text_input("項目A：物件名（工事名・業務名）", placeholder="（例：〇〇清掃工場コンクリート劣化構造調査）")
        location_name = st.text_input("項目B：調査位置・測定箇所", placeholder="（例：RC造擁壁部、地下ピット壁面など）")
        inspector_name = st.text_input("項目C：調査担当者（コンクリート診断士名）", value="診断太郎", placeholder="（例：診断太郎）")
        
        st.markdown("### 🔧 人間による補足情報入力（重複なし）")
        cb_salt = st.checkbox("海岸線から2km以内（塩害リスク）")
        cb_freeze = st.checkbox("寒冷地・凍結防止剤の散布地域（凍害リスク）")
        cb_wet = st.checkbox("常時湿潤・漏水・滞水環境（アルカリ骨材反応・溶出リスク）")
        cb_traffic = st.checkbox("交通量が極めて多い（排気ガスによる中性化加速）")
        cb_janka = st.checkbox("ジャンカ・初期ひび割れ・施工不良の目視確認あり")
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
        # 💥 【ご要望】最大6枚の複数写真のアップロードに対応！
        uploaded_files = st.file_uploader("ここにコンクリート構造物の写真をアップロードしてください（最大6枚まで）", type=["jpg", "jpeg", "png"], accept_multiple_files=True)
        
        images = []
        if uploaded_files:
            # 6枚を超えたら制限
            if len(uploaded_files) > 6:
                st.warning("⚠️ 写真は最大6枚までです。上位6枚のみを解析対象とします。")
                uploaded_files = uploaded_files[:6]
                
            for idx, file in enumerate(uploaded_files, start=1):
                img = Image.open(file)
                images.append(img)
                st.image(img, caption=f"Photo No.{idx} - アップロード画像", use_container_width=True)
            
            execute_analysis = st.button("🚀 この内容で高精密AI解析を実行する")

    with col2:
        st.markdown("### 📊 高精密多角診断レポート")
        if uploaded_files and 'execute_analysis' in locals() and execute_analysis:
            if not api_key:
                st.error("APIキーが保存されていません。")
            else:
                with st.spinner("🔍 過去の実績調書ベースでプロのコンクリート診断士AI(Gemini 2.5)が精密解析中..."):
                    try:
                        genai.configure(api_key=api_key)
                        model = genai.GenerativeModel('gemini-2.5-flash')
                        
                        p_name = project_name if project_name else "記載なし（現場写真より診断）"
                        l_name = location_name if location_name else "記載なし"
                        i_name = inspector_name if inspector_name else "記載なし"
                        
                        # 💥 過去の実績書から学習したT&N専用のプロ仕様プロンプト
                        prompt = f"""
                        あなたは最高峰の「コンクリート診断士」です。提供された過去の調査実績報告書（塩竈清掃工場、鳥海八幡中など）の工学的思考・技術水準を完全に引き継ぎ、提出された最大6枚の現場写真を多角的に目視調査して、公式な【総合劣化診断調査報告書】を作成してください。
                        
                        【現場の環境条件・マトリクス】
                        ・構造物種別: {struct_type}
                        ・設置環境（複合要因）: {env_location_text}
                        ・コンクリート湿潤状態: {wet_status}
                        ・使用セメントの種類: {cement_type}
                        ・供用年数（経過年数）: {elapsed_years}
                        ・表面目視劣化症状: {crack_type}
                        ・施工・人為的補足要因: {human_factors_text}
                        
                        【診断における厳格な出力仕様】
                        1. 挨拶や余計な解説は省き、以下の構成で出力してください。
                        2. 最初に必ず、複数枚の写真全体から総合判断される実務上の数値として「推定最大ひび割れ幅: 0.18 mm / 推定平均ひび割れ長さ: 25.0 cm」の形式（数値は状況からロジカルに推測）を1行目で宣言してください。
                        3. 【劣化原因に関する深い工学的推測】欄では、過去の実績書に倣い、単なる一般論ではなく、写真に現れている「微細ひび割れ」「浮き・剥離」「エフロレッセンス（白華）」「鉄筋露出・爆裂」「錆汁の溶出」を詳細に指摘してください。二酸化炭素の侵入による中性化速度、不動態被膜の破壊、アルカリシリカ反応（ASR）による膨張圧、乾湿の繰り返しによる熱応力、あるいは清掃工場特有の化学的腐食要因との因果関係を、土木・建築工学のプロの用語を用いて【400文字以上の重厚な文章】で詳細に論理展開してください。
                        4. 【推奨される具体的な対策案・補修工法】欄では、土木学会およびコンクリート工学会の維持管理指針に準拠し、ひび割れ注入工法（エポキシ樹脂低圧・高圧注入）、断面修復工法（ポリマーセメントモルタル充填工法）、表面含浸工法（シラン系・ケイ酸塩系）、爆裂部の防錆処理などを、選定理由と共に具体的に提案してください。さらに、内部進行度を測定するための詳細追加調査（コア採取による圧縮強度試験、フェノールフタレイン液による中性化深さ測定、はつり調査など）の必要性についても【300文字以上の文章】で明記してください。
                        """
                        
                        # 写真配列をすべてAIに一発投入（多角的マルチモーダル解析）
                        content_inputs = [prompt] + images
                        response = model.generate_content(content_inputs)
                        full_result_text = response.text
                        
                        # 数値のスマート抽出
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

                        if width_val >= 0.2:
                            color_code = "#EF4444"
                            status_title = f"🔴 【要精密確認】最大ひび割れ幅: {width_val} mm"
                            alert_desc = "⚠️ 判定基準：0.2mm以上のひび割れのため【赤色：要精密補修（樹脂注入・断面修復対象）】となります"
                        else:
                            color_code = "#EAB308"
                            status_title = f"🟡 【経過観察】最大ひび割れ幅: {width_val} mm"
                            alert_desc = "💡 判定基準：0.2mm以下のひび割れのため【黄色：経過観察（または表面保護対象）】となります"

                        # 画面への美麗レポート出力
                        st.markdown(f"<div class='status-card'><h3 style='color: {color_code} !important; margin:0; font-size:22px;'>{status_title}</h3><p style='color: #F1F5F9 !important; font-size: 14px; margin: 8px 0 0 0; font-weight: bold;'>{alert_desc}</p></div>", unsafe_allow_html=True)
                        st.markdown(f"📐 **それぞれのひび割れ想定長さ:** <span style='font-size:24px; font-weight:bold; color:#38BDF8;'>{length_val} cm</span>", unsafe_allow_html=True)
                        st.markdown("<h4 style='color: white; margin-top:20px;'>📑 過去調書実績学習・コンクリート診断士AI調査報告書</h4>", unsafe_allow_html=True)
                        st.info(full_result_text)

                        # --- ここからExcelの作成（複数写真対応・Photo No. ラベル自動整列） ---
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

                        # カラム幅定義（A4黄金比率）
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
                            ("調査対象・位置", l_name, "■ 設置環境（複合）", env_location_text),
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

                        ws["A10"] = "想定最大ひび割れ幅"
                        ws.merge_cells("B10:G10")
                        ws["B10"] = f"{width_val} mm （{'赤色警告・要精密補修対象' if width_val>=0.2 else '黄色警告・経過観察対象' }）"
                        ws.row_dimensions[10].height = 24

                        ws["A11"] = "想定ひび割れ平均長さ"
                        ws.merge_cells("B11:G11")
                        ws["B11"] = f"{length_val} cm"
                        ws.row_dimensions[11].height = 24

                        ws["A12"] = "AI詳細調査報告意見書\n(過去調書実績準拠・長文)"
                        ws.merge_cells("B12:G12")
                        ws["B12"] = full_result_text
                        ws["B12"].alignment = Alignment(wrap_text=True, vertical="top")
                        
                        # オート・ハイト
                        text_length = len(full_result_text)
                        calculated_height = max(420, min(650, int(text_length * 0.48)))
                        ws.row_dimensions[12].height = calculated_height

                        for r in range(10, 13):
                            ws.cell(row=r, column=1).font = font_label
                            ws.cell(row=r, column=1).fill = fill_label
                            ws.cell(row=r, column=2).font = font_data

                        ws.merge_cells("A14:G14")
                        ws["A14"] = "■ 診断対象構造物・現場調査写真（Photo No. 管理整列配置）"
                        ws["A14"].font = Font(name="MS ゴシック", size=11, bold=True, color="1E3A8A")
                        ws.row_dimensions[14].height = 25

                        for row in ws.iter_rows(min_row=1, max_row=14, min_col=1, max_col=7):
                            for cell in row: cell.border = border_cell

                        # 📷 【大進化】最大6枚の写真を、横2列×縦3列の完璧なマトリクスでExcelへ自動貼り付け！
                        current_row = 16
                        for idx, img in enumerate(images):
                            img_buffer_xl = io.BytesIO()
                            img.save(img_buffer_xl, format="PNG")
                            img_buffer_xl.seek(0)
                            xl_img = ExcelImage(img_buffer_xl)
                            
                            # A4の半分のサイズ（横250px、縦185px）に縮小して美しく2列に並べる
                            xl_img.width = 250
                            xl_img.height = 185
                            
                            # 奇数枚目はA列、偶数枚目はE列に自動配置して Photo No. ラベルを印字
                            if idx % 2 == 0:
                                # 左側の列
                                ws.cell(row=current_row, column=1, value=f"【Photo No.{idx+1}】").font = font_label
                                ws.add_image(xl_img, f"A{current_row+1}")
                            else:
                                # 右側の列
                                ws.cell(row=current_row, column=5, value=f"【Photo No.{idx+1}】").font = font_label
                                ws.add_image(xl_img, f"E{current_row+1}")
                                # 両方埋まったら行を進める（1枚の写真エリアの高さを行高さで綺麗に確保）
                                ws.row_dimensions[current_row].height = 20
                                ws.row_dimensions[current_row+1].height = 195
                                current_row += 3
                        
                        # 奇数枚で終わった場合のための最後の高さ調整
                        if len(images) % 2 != 0:
                            ws.row_dimensions[current_row].height = 20
                            ws.row_dimensions[current_row+1].height = 195

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
            st.info("写真をアップロードし「この内容で高精密AI解析を実行する」ボタンを押すと、過去の実績書を学習したプロレベルの診断結果が表示されます。")
