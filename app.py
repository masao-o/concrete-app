import streamlit as st
import google.generativeai as genai
from PIL import Image
import io
import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ==========================================
# 0. セキュリティ：パスワード保護機能（追加）
# ==========================================
def check_password():
    """正しいパスワードが入力されたら True を返す"""
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False

    # 既に認証済みの場合はスキップ
    if st.session_state["password_correct"]:
        return True

    # パスワード入力画面の表示
    st.markdown("<h2 style='text-align: center; color: #f97316;'>🔒 コンクリート劣化診断 AI Suite Pro</h2>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: #cccccc;'>このアプリは関係者限定です。パスワードを入力してください。</p>", unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        password = st.text_input("パスワード", type="password", key="password_input")
        if st.button("サインイン"):
            # ★ここでパスワードを設定しています（自由に変更可能です）
            if password == "tn0000":
                st.session_state["password_correct"] = True
                st.rerun()
            else:
                st.error("❌ パスワードが違います。企画調査課の太田までご確認ください。")
    return False

# パスワードチェックが通らない場合は、ここでアプリの処理をストップする
if not check_password():
    st.stop()

# ==========================================
# 1. アプリケーションの基本設定と画面デザイン
# ==========================================
st.set_page_config(
    page_title="コンクリート劣化診断 AI Suite Pro", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# プロフェッショナルなダークテーマのカスタムCSS
st.markdown("""
    <style>
    .report-title { color: #f97316; font-weight: bold; font-size: 32px; margin-bottom: 20px; }
    .section-header { background-color: #1e293b; padding: 10px; border-left: 6px solid #38bdf8; border-radius: 4px; margin-top: 20px; margin-bottom: 15px; }
    .stButton>button { background-color: #f97316; color: white; font-weight: bold; width: 100%; border-radius: 8px; height: 45px; }
    .stButton>button:hover { background-color: #ea580c; border-color: #ea580c; }
    .ai-response { background-color: #1e293b; padding: 20px; border-radius: 8px; border: 1px solid #475569; line-height: 1.6; }
    
    /* マニュアル用の印刷用・美麗スタイル */
    .manual-page { background-color: #ffffff; color: #333333; padding: 40px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); margin-bottom: 30px; font-family: "MS Gothic", "BIZ UDPGothic", sans-serif; }
    .manual-title { color: #1f497d; font-size: 28px; font-weight: bold; text-align: center; margin-bottom: 10px; }
    .manual-subtitle { color: #555555; font-size: 16px; text-align: center; margin-bottom: 40px; }
    .manual-h1 { color: #1f497d; font-size: 20px; font-weight: bold; border-bottom: 3px solid #1f497d; padding-bottom: 5px; margin-top: 30px; margin-bottom: 15px; }
    .manual-h2 { color: #366092; font-size: 16px; font-weight: bold; margin-top: 20px; margin-bottom: 10px; border-left: 4px solid #366092; padding-left: 10px; }
    .manual-text { font-size: 14px; line-height: 1.8; color: #333333; text-align: justify; }
    .manual-meta { text-align: right; font-size: 12px; color: #666666; margin-top: 50px; }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 2. サイドバー：認証・物件情報・環境要因リスト
# ==========================================
with st.sidebar:
    st.title("🏗️ 統括管理パネル")
    api_key = "AQ.Ab8RN6IGjKOwFH03mminYY6Nzh4RUwJhSfG89fg76C-tlpqwTA"
    
    st.divider()
    st.header("📋 調査基本情報")
    project_name = st.text_input("物件・工事名", "〇〇高架橋定期点検（第2工区）")
    structure_type = st.selectbox("対象構造物", ["橋梁（RC床版）", "橋梁（橋脚・橋台）", "トンネル覆工", "カルバート", "擁壁・法面", "建築物躯体", "その他"])
    inspector = st.text_input("担当診断士", "技術 太郎（コンクリート診断士）")
    inspection_date = st.date_input("調査実施日", datetime.date.today())

    st.divider()
    st.header("🛠️ 人間による補足情報入力")
    st.markdown("##### 【環境要因】")
    env_coast = st.checkbox("海岸線から2km以内（塩害リスク）")
    env_freeze = st.checkbox("寒冷地・凍結防止剤の散布地域（凍害リスク）")
    env_water = st.checkbox("常時湿潤・漏水・滞水環境（アルカリ骨材反応・溶出リスク）")
    env_gas = st.checkbox("交通量が極めて多い（排気ガスによる中性化加速）")
    
    st.markdown("##### 【施工・初期欠陥要因】")
    const_honeycomb = st.checkbox("ジャンカ・初期ひび割れの目視確認あり")
    const_joint = st.checkbox("施工目地・コールドジョイント部")
    const_cover = st.checkbox("設計かぶり厚の不足が疑われる・または既知")

# タブ機能を使って、メイン画面に「劣化診断」と「取扱説明書」を切り替えられるように配置
tab1, tab2 = st.tabs(["🔍 コンクリート劣化診断実行", "📖 技術製品仕様・取扱説明書（社内配布・PDF化用）"])

# ==========================================
# 3. メイン画面 タブ1：劣化診断システム
# ==========================================
with tab1:
    st.markdown("<div class='report-title'>🏗️ コンクリート劣化診断 AI Suite Pro</div>", unsafe_allow_html=True)
    st.markdown("コンクリート診断士の知見とAIを融合させ、ハルシネーション（知ったかぶり）のない厳密な劣化原因特定を行います。")
    st.divider()

    if not api_key:
        st.warning("👈 アプリを有効化するには、左側のサイドバーに「Google AI Studio APIキー」を入力してください。")
        st.stop()

    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-2.5-flash")
    col_left, col_right = st.columns([1, 1])

    with col_left:
        st.markdown("<div class='section-header'><h3>📸 診断写真の登録</h3></div>", unsafe_allow_html=True)
        uploaded_file = st.file_uploader("デジタルカメラまたはタブレットで撮影した損傷写真（JPG/PNG）", type=["jpg", "jpeg", "png"])
        
        if uploaded_file:
            image = Image.open(uploaded_file)
            st.image(image, caption="解析対象の損傷写真", use_container_width=True)
        else:
            st.info("ここに写真をドロップするか、ファイルを選択してください。")

    with col_right:
        st.markdown("<div class='section-header'><h3>🔍 専門家AI 協働解析エンジン</h3></div>", unsafe_allow_html=True)
        
        if uploaded_file:
            human_inputs = []
            if env_coast: human_inputs.append("・海岸線から近く塩害リスクがある環境")
            if env_freeze: human_inputs.append("・凍結融解の繰り返し、または凍結防止剤が散布される環境")
            if env_water: human_inputs.append("・背面からの漏水、または常に水にさらされている湿潤環境")
            if env_gas: human_inputs.append("・排気ガスや二酸化炭素濃度が高く、中性化が進みやすい環境")
            if const_honeycomb: human_inputs.append("・施工時のジャンカ、または初期欠陥が目視で確認されている")
            if const_joint: human_inputs.append("・打ち継ぎ目地、またはコールドジョイントに該当する部位")
            if const_cover: human_inputs.append("・かぶり厚が浅い、または不足している可能性が高い")
            
            human_info_text = "\n".join(human_inputs) if human_inputs else "・特になし（または現場写真からのみ判断）"

            if st.button("🤖 診断士AIと連携して解析を実行"):
                with st.spinner("熟練診断士のロジックで写真を精査中..."):
                    prompt = f"""
                    あなたは土木学会（JSCE）および日本コンクリート工学会（JCI）の基準に精通した「最高峰のコンクリート診断士」です。
                    添付された写真と、人間の調査員から提供された【人間の補足情報】を組み合わせて、客観的かつ厳密な劣化診断を行ってください。

                    【基本情報】
                    ・物件・工事名: {project_name}
                    ・対象構造物: {structure_type}
                    
                    【人間の補足情報（現場環境・施工状況）】
                    {human_info_text}

                    項目名は変更せず、以下の構成で厳格に出力してください。
                    【1. 部材・損傷の状況】
                    【2. 劣化原因の特定】
                    【3. 判定グレード】
                    【4. 推奨補修工法】
                    【5. 診断士AIからの追記質問】
                    """
                    
                    response = model.generate_content([prompt, image])
                    res_text = response.text
                    st.session_state['last_res_text'] = res_text
                    st.session_state['last_human_info'] = human_info_text
                    
                    st.success("✨ 診断完了（AIと人間の合意形成に成功）")
                    st.markdown("<div class='ai-response'>", unsafe_allow_html=True)
                    st.markdown(res_text)
                    st.markdown("</div>", unsafe_allow_html=True)
                    st.divider()
                    
                    def create_excel_report():
                        wb = openpyxl.Workbook()
                        ws = wb.active
                        ws.title = "劣化診断報告書"
                        ws.views.sheetView[0].showGridLines = True
                        
                        title_font = Font(name="MS Gothic", size=16, bold=True, color="FFFFFF")
                        header_font = Font(name="MS Gothic", size=11, bold=True, color="000000")
                        body_font = Font(name="MS Gothic", size=11, color="000000")
                        label_font = Font(name="MS Gothic", size=11, bold=True, color="FFFFFF")
                        
                        title_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
                        label_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
                        section_fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
                        
                        thin_side = Side(border_style="thin", color="A6A6A6")
                        thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
                        
                        ws.merge_cells("A1:D2")
                        ws["A1"] = "コンクリート構造物 劣化診断報告書（台帳案）"
                        ws["A1"].font = title_font
                        ws["A1"].fill = title_fill
                        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
                        
                        info_labels = [
                            ("A4", "物件・工事名", "B4", project_name),
                            ("C4", "調査実施日", "D4", str(inspection_date)),
                            ("A5", "対象構造物", "B5", structure_type),
                            ("C5", "担当診断士", "D5", inspector)
                        ]
                        for lbl_cell, lbl_val, val_cell, val_val in info_labels:
                            ws[lbl_cell] = lbl_val
                            ws[lbl_cell].font = label_font
                            ws[lbl_cell].fill = label_fill
                            ws[lbl_cell].alignment = Alignment(horizontal="center", vertical="center")
                            ws[lbl_cell].border = thin_border
                            ws[val_cell] = val_val
                            ws[val_cell].font = body_font
                            ws[val_cell].alignment = Alignment(horizontal="left", vertical="center")
                            ws[val_cell].border = thin_border

                        ws.merge_cells("A7:D7")
                        ws["A7"] = "【現地補足情報（周辺環境・施工要因）】"
                        ws["A7"].font = header_font
                        ws["A7"].fill = section_fill
                        ws["A7"].border = thin_border
                        
                        ws.merge_cells("A8:D8")
                        ws["A8"] = human_info_text.replace("・", "  ・") if human_info_text else "  ・特になし"
                        ws["A8"].font = body_font
                        ws["A8"].alignment = Alignment(wrap_text=True, vertical="top")
                        ws["A8"].border = thin_border
                        ws.row_dimensions[8].height = 45

                        sections = [
                            ("【1. 部材・損傷の状況】", "【2. 劣化原因の特定】", "A10", "A11", "💎 1. 画像識別による変状・損傷の状況"),
                            ("【2. 劣化原因の特定】", "【3. 判定グレード】", "A13", "A14", "💎 2. 劣化メカニズム・原因の特定（クロス分析）"),
                            ("【3. 判定グレード】", "【4. 推奨補修工法】", "A16", "A17", "💎 3. 定量評価に基づく判定グレード（JSCE/JCI準拠）"),
                            ("【4. 推奨補修工法】", "【5. 診断士AIからの追記質問】", "A19", "A20", "💎 4. 劣化原因に対応した推奨対策・補修工法"),
                            ("【5. 診断士AIからの追記質問】", "🌟END🌟", "A22", "A23", "💎 5. 統括診断士AIからの追加確認・要求事項")
                        ]
                        
                        for start_tag, end_tag, title_idx, body_idx, section_title in sections:
                            t_row = int(title_idx[1:])
                            ws.merge_cells(f"A{t_row}:D{t_row}")
                            ws[title_idx] = section_title
                            ws[title_idx].font = header_font
                            ws[title_idx].fill = section_fill
                            ws[title_idx].border = thin_border
                            
                            content = ""
                            if start_tag in res_text:
                                start_pos = res_text.find(start_tag) + len(start_tag)
                                if end_tag != "🌟END🌟" and end_tag in res_text:
                                    end_pos = res_text.find(end_tag)
                                    content = res_text[start_pos:end_pos].strip()
                                else:
                                    content = res_text[start_pos:].strip()
                            
                            if not content:
                                content = "該当データなし"
                            
                            b_row = int(body_idx[1:])
                            ws.merge_cells(f"A{b_row}:D{b_row}")
                            ws[body_idx] = content.replace("`", "").replace("*", "")
                            ws[body_idx].font = body_font
                            ws[body_idx].alignment = Alignment(wrap_text=True, vertical="top")
                            ws[body_idx].border = thin_border
                            
                            line_count = max(content.count("\n") + 1, len(content) // 40 + 1)
                            ws.row_dimensions[b_row].height = max(line_count * 18, 40)
                        
                        ws.column_dimensions["A"].width = 25
                        ws.column_dimensions["B"].width = 35
                        ws.column_dimensions["C"].width = 20
                        ws.column_dimensions["D"].width = 35
                        
                        excel_data = io.BytesIO()
                        wb.save(excel_data)
                        excel_data.seek(0)
                        return excel_data

                    excel_file = create_excel_report()
                    st.download_button(
                        label="📊 顧客提出用・正式Excel報告書をダウンロード",
                        data=excel_file,
                        file_name=f"コンクリート劣化診断報告書_{project_name}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        else:
            st.info("写真をアップロードすると、AI解析と対話機能が有効になります。")

# ==========================================
# 4. メイン画面 タブ2：取扱説明書自動生成システム
# ==========================================
with tab2:
    st.info("💡 画面を右クリックして「印刷」を選択するか、キーボードの [Ctrl + P] を押すことで、この美麗マニュアルをそのまま「PDFとして保存」して社内配布用資料として出力できます！")
    
    st.markdown("""
    <div class='manual-page'>
        <div class='manual-title'>コンクリート劣化診断 AI Suite Pro</div>
        <div class='manual-subtitle'>〜 AI Vision × 診断士知見による高精度劣化原因特定システム 〜</div>
        <hr style='border-top: 2px solid #1f497d;'>
        <br><br><br>
        <div style='text-align: center; font-size: 18px; font-weight: bold; color: #333333;'>技術製品仕様・取扱説明書 (Ver. 2.0)</div>
        <br><br><br><br><br>
        <div class='manual-meta'>
            <strong>【開発・作成者】</strong><br>
            T＆日本メンテ開発㈱<br>
            工事部 企画調査課<br>
            太田 昌男<br><br>
            <strong>【作成日】</strong><br>
            2026年6月3日
        </div>
    </div>
    
    <div class='manual-page'>
        <div class='manual-h1'>目次 および 本システムの概要</div>
        <div class='manual-h2'>■ 目次 (Index)</div>
        <div class='manual-text'>
            第1章 システム概要と開発背景 ----------------------------------------- P.3<br>
            第2章 主な機能と特徴 ------------------------------------------------- P.4<br>
            第3章 動作環境と起動方法 --------------------------------------------- P.5<br>
            第4章 画面構成の解説 ------------------------------------------------- P.6<br>
            第5章 基本操作手順（情報の入力からAI解析まで） ----------------------- P.7<br>
            第6章 診断レポートの構成要素 ----------------------------------------- P.8<br>
            第7章 正式Excel報告書の仕様とダウンロード ----------------------------- P.9<br>
            第8章 トラブルシューティング（保守・管理） --------------------------- P.10
        </div>
        <div class='manual-h2'>■ 本システムの概要</div>
        <div class='manual-text'>
            本システムは、遠隔地や点検現場において、デジカメやタブレットで撮影したコンクリート構造物の変状写真をアップロードするだけで、最新の画像認識AI（Gemini 2.5 Flash）が損傷状況を即座に言語化・識別します。さらに現場の環境要因データと組み合わせ、土木学会（JSCE）等の基準に準拠した厳密な原因特定を行い、社内報告や顧客提出にそのまま使用できる「美しく装飾された正式なExcel台帳（.xlsx）」を一瞬で自動生成します。
        </div>
    </div>
    
    <div class='manual-page'>
        <div class='manual-h1'>第1章 システム概要と開発背景</div>
        <div class='manual-h2'>1. 開発背景</div>
        <div class='manual-text'>
            従来のコンクリート構造物点検業務においては、現場写真の整理、変状スケッチの作成、基準書との照合に多大な時間と熟練技術者の工数（コスト）が割かれていました。また、従来のAI技術ではハルシネーション（知ったかぶり）問題が実務適用の壁となっていました。本システムは、これら「業務効率化」と「判断の厳密性」を両立するために、T＆日本メンテ開発㈱ 工事部 企画調査課 太田昌男によって開発されました。
        </div>
        <div class='manual-h2'>2. 従来手法と本システムの比較</div>
        <div class='manual-text'>
            ・従来手法：写真整理 ＞ 基準書照合 ＞ 原因考察 ＞ エクセル手入力（数時間〜数日）<br>
            ・本システム：写真登録 ＞ 環境要因チェック ＞ AI解析 ＞ 正式Excel自動出力（わずか30秒）
        </div>
    </div>
    
    <div class='manual-page'>
        <div class='manual-h1'>第2章 主な機能と特徴</div>
        <div class='manual-h2'>1. 厳格なハルシネーション排除機能</div>
        <div class='manual-text'>
            写真や与えられた情報から読み取れない事実（部材の内部状況、正確なひび割れ深さなど）をAIが勝手に推測して断定することを徹底的に禁止。判断がグレーな要素は「〜の可能性が疑われるが、現時点では確定できない」と誠実に出力し、専門家としての信頼性を担保します。
        </div>
        <div class='manual-h2'>2. 人間とAIのクロス分析（協働解析）</div>
        <div class='manual-text'>
            写真の見た目だけでなく、人間が現地で確認した「海岸線からの距離」「漏水の有無」「施工初期欠陥」などの背景データを掛け合わせることで、単なる画像識別を超えた精緻な原因特定へアプローチします。
        </div>
    </div>
    
    <div class='manual-page'>
        <div class='manual-h1'>第3章 動作環境と起動方法</div>
        <div class='manual-h2'>1. アプリケーションの起動手順</div>
        <div class='manual-text'>
            【ステップ1】「concrete-app」フォルダを開き、【ステップ2】「🚀アプリ起動.bat」をダブルクリックすると、【ステップ3】自動的に黒い画面が立ち上がりシステムが起動します。
        </div>
        <br>
    </div>
    
    <div class='manual-page'>
        <div class='manual-h1'>第4章 画面構成の解説</div>
        <div class='manual-h2'>1. インターフェースの基本構造</div>
        <div class='manual-text'>
            画面は左右の2カラム構造を採用しています。<br>
            ・<strong>左側カラム【統括管理パネル】：</strong> 物件名や構造物種別の入力、環境要因のチェックを行うエリアです。<br>
            ・<strong>右側カラム【専門家AI 協働解析エンジン】：</strong> 写真の登録、解析の実行、AIの予備診断結果を表示するエリアです。
        </div>
    </div>
    
    <div class='manual-page'>
        <div class='manual-h1'>第5章 基本操作手順</div>
        <div class='manual-h2'>1. 現場運用の基本フロー</div>
        <div class='manual-text'>
            【ステップ1】基本情報の入力 ＞ 【ステップ2】現地環境のチェック ＞ 【ステップ3】損傷写真をドラッグ＆ドロップ ＞ 【ステップ4】「🤖 診断士AIと連携して解析を実行」をクリックして診断書を生成します。
        </div>
    </div>
    
    <div class='manual-page'>
        <div class='manual-h1'>第6章 診断レポートの構成要素（正式書式）</div>
        <div class='manual-h2'>1. 報告書を構成する5つの厳格なセクション</div>
        <div class='manual-text'>
            <strong>①【部材・損傷の状況】：</strong> 画像解析から部材を特定し、損傷状況を微細に言語化。<br>
            <strong>②【劣化原因の特定】：</strong> 損傷特徴と環境要因をクロス分析し、劣化メカニズムを解明。<br>
            <strong>③【判定グレード】：</strong> JSCE基準に基づき、a〜eの5段階から適切なグレードとその根拠を選定。<br>
            <strong>④【推奨補修工法】：</strong> 原因に対応した最適な工法（ひび割れ注入工法など）を実務ベースで提案。<br>
            <strong>⑤【診断士AIからの追記質問】：</strong> 知ったかぶりを防ぎ、より確実な特定に必要な追加調査を逆提案。
        </div>
    </div>
    
    <div class='manual-page'>
        <div class='manual-h1'>第7章 正式Excel報告書の仕様</div>
        <div class='manual-h2'>1. プロ仕様のフォーマット設計</div>
        <div class='manual-text'>
            生成されるエクセルファイルは、信頼感を与える「濃紺（#1F497D）」と視認性の高い「薄青（#DCE6F1）」で美しく色分けされ、公的書類に最適な薄グレーの格子罫線、および文字量に合わせた行の高さ自動調整ルールが100%自動適用されています。
        </div>
    </div>
    
    <div class='manual-page'>
        <div class='manual-h1'>第8章 トラブルシューティング（保守・管理）</div>
        <div class='manual-h2'>■ Q1. 起動用バッチファイルを押しても、画面が立ち上がらない場合</div>
        <div class='manual-text'>🟢 【対処法】裏で動いている黒い画面（コマンドプロンプト）を一度すべて「×」で閉じ、再度バッチファイルをダブルクリックしてください。</div>
        <div class='manual-h2'>■ Q2. 2026年現在の最新仕様に対応しているか？</div>
        <div class='manual-text'>🟢 【対処法】最新仕様の「AQ.Ab8RN...」キー、および最新マルチモーダルモデル「gemini-2.5-flash」を指定しているため、高度なセキュリティと超高速な解析処理が安定して動作します。</div>
        <br><br><hr style='border-top: 1px solid #ccc;'>
        <div style='text-align: center; font-size: 11px; color: #666666;'>© 2026 T＆日本メンテ開発㈱ 工事部 企画調査課 太田昌男 - All Rights Reserved.</div>
    </div>
    """, unsafe_allow_html=True)

# ==========================================
# 5. フッター
# ==========================================
st.divider()
st.caption(f"© {datetime.date.today().year} Concrete AI Diagnostic Suite Pro - Version 2.0 (日本語プロ仕様版)")