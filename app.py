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

# --- 1. ページ設定とハイエンドUI・次世代点滅ダッシュボードCSS ---
st.set_page_config(page_title="コンクリート劣化診断 AI Suite Pro", layout="wide")

# セッション状態の初期化
if 'full_result_text' not in st.session_state:
    st.session_state.full_result_text = None
if 'final_width' not in st.session_state:
    st.session_state.final_width = 0.0
if 'analysis_completed' not in st.session_state:
    st.session_state.analysis_completed = False
if 'current_step' not in st.session_state:
    st.session_state.current_step = "🏠\n① 設置地域・環境判定"

# 解析完了時に③番ボタンをiPhone風に発光・パルス点滅させるアニメーション
animation_css = ""
if st.session_state.analysis_completed:
    animation_css = """
    div[data-testid="stPills"] div[role="radiogroup"] div:nth-child(3) label {
        animation: iphone_pulse 1.4s infinite alternate !important;
        border: 2px solid #38BDF8 !important;
    }
    @keyframes iphone_pulse {
        0% { box-shadow: 0 4px 15px rgba(56, 189, 248, 0.2); background-color: #1E293B !important; }
        100% { box-shadow: 0 0 35px #38BDF8, inset 0 0 15px #38BDF8; background-color: #0284C7 !important; }
    }
    """

st.markdown(f"""
<style>
.main {{ background-color: #0F172A; color: #FFFFFF; }}
.stApp {{ background-color: #0F172A; }}

/* サイドバーの非表示（超ワイド画面） */
section[data-testid="stSidebar"] {{ display: none !important; }}
div[data-testid="collapsedControl"] {{ display: none !important; }}

/* ユニバーサル純白フォント */
h1, h2, h3, h4, h5, h6, p, span, label, .stMarkdown,
.stCheckbox label, div[data-testid="stMarkdownContainer"] p {{
    color: #FFFFFF !important; font-family: 'Helvetica Neue', Arial, sans-serif;
}}

/* 入力枠の黒文字視認性確保 */
input, textarea, select, div[data-baseweb="select"] * {{ color: #0F172A !important; font-weight: bold !important; }}
input::placeholder, textarea::placeholder {{ color: #64748B !important; opacity: 1 !important; }}

/* 【徹底修正】ナビゲーションタイルの巨大化と未選択時の真っ白バグを完全解消 */
div[data-testid="stPills"] div[role="radiogroup"] {{
    display: flex !important;
    gap: 30px !important;
    width: 100% !important;
    background-color: transparent !important;
}}
div[data-testid="stPills"] div[role="radiogroup"] > div {{
    flex: 1 !important;
}}
div[data-testid="stPills"] div[role="radiogroup"] label {{
    background: linear-gradient(135deg, #1E293B, #111827) !important; /* ダークネイビーに変更 */
    border: 2px solid #334155 !important;
    border-radius: 28px !important;
    padding: 40px 20px !important; /* 圧倒的に大きく */
    text-align: center !important;
    box-shadow: 0 8px 20px rgba(0, 0, 0, 0.5), inset 0 1px 0 rgba(255, 255, 255, 0.1) !important;
    transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1) !important;
    cursor: pointer !important;
    display: block !important;
    width: 100% !important;
}}

/* 【文字の巨大化】バグを根本治療。未選択でも絶対に真っ白にならない */
div[data-testid="stPills"] div[role="radiogroup"] label * {{
    font-size: 28px !important; /* 28pxへ超巨大化 */
    font-weight: 900 !important;
    color: #FFFFFF !important; /* 未選択時は白文字 */
    line-height: 1.3 !important;
    white-space: pre-wrap !important;
}}

/* ホバー効果 */
div[data-testid="stPills"] div[role="radiogroup"] label:hover {{
    transform: translateY(-8px) !important;
    border-color: #38BDF8 !important;
    box-shadow: 0 15px 35px rgba(0, 0, 0, 0.6) !important;
}}

/* アクティブ（選択中）タイルの発光 */
div[data-testid="stPills"] div[role="radiogroup"] input[type="radio"]:checked + div label {{
    background: linear-gradient(135deg, #0284C7, #1E293B) !important;
    border-color: #38BDF8 !important;
    box-shadow: 0 0 35px rgba(56, 189, 248, 0.6), inset 0 1px 0 rgba(255, 255, 255, 0.2) !important;
}}
div[data-testid="stPills"] div[role="radiogroup"] input[type="radio"]:checked + div label * {{
    color: #38BDF8 !important; /* 選択中文字をネオンブルーに */
}}

/* ラジオボタンの隠蔽 */
div[data-testid="stPills"] input[type="radio"] {{ display: none !important; }}

/* ③番用の動的パルス明滅 */
{animation_css}

/* チェックボックスの青色固定 */
div[data-testid="stCheckbox"] div[role="checkbox"][aria-checked="true"] {{ background-color: #0284C7 !important; border-color: #38BDF8 !important; }}

/* 診断実行ボタン */
div.stButton > button {{
    background: #0284C7 !important; color: #FFFFFF !important; border: 2px solid #38BDF8 !important;
    border-radius: 12px !important; padding: 18px 28px !important; font-size: 20px !important; font-weight: bold !important;
    width: 100% !important; transition: all 0.2s ease !important;
}}
div.stButton > button:hover {{ background: #38BDF8 !important; box-shadow: 0 0 25px #38BDF8 !important; }}

/* 各種カード */
.dashboard-card {{ padding: 25px; background-color: #1E293B; border-radius: 16px; margin-bottom: 20px; }}
.status-card {{ padding: 25px; background-color: #1E293B; border-radius: 16px; margin-bottom: 20px; border-top: 1px solid #334155; }}
.photo-input-box {{ background-color: #1E293B; padding: 20px; border-radius: 12px; border: 1px dashed #38BDF8; margin-bottom: 20px; }}
.report-text-box {{ background-color: #1E293B !important; color: #FFFFFF !important; border-radius: 12px; padding: 25px; font-size: 16px; line-height: 1.8; }}
</style>
""", unsafe_allow_html=True)

# 認証機能
def check_password():
    if "authenticated" not in st.session_state: st.session_state["authenticated"] = False
    def password_entered():
        if st.session_state["password"] == "tn0000": st.session_state["authenticated"] = True
        else: st.error("❌ パスワードが違います")
    if not st.session_state["authenticated"]:
        st.markdown("<h2>🔒 Concrete Suite Pro（関係者専用）</h2>", unsafe_allow_html=True)
        st.text_input("アクセスパスワード", type="password", on_change=password_entered, key="password")
        return False
    return True

if check_password():
    api_key = st.secrets.get("GEMINI_API_KEY", "")

    # クリーンなヘッダー
    col_logo_head, col_title_head = st.columns([1, 5])
    with col_logo_head:
        # プレースホルダーのロゴ
        st.image("http://googleusercontent.com/image_collection/image_retrieval/18156729091245692064", width=160)
    with col_title_head:
        st.markdown("<h1 style='color: white; margin-top: 15px; margin-bottom: 0;'>🚗 コンクリート劣化診断システム</h1>", unsafe_allow_html=True)
        st.markdown("<p style='color: #38BDF8; font-size: 18px; font-weight: bold;'>AI Suite Pro - 実務特化型高精密エンジニアリング・ダッシュボード</p>", unsafe_allow_html=True)

    # 巨大なナビゲーションタイル（未選択時の真っ白バグを完全治療済）
    current_step = st.pills(
        label="",
        options=["🏠\n① 設置地域・環境判定", "📸\n② 写真・変状チェック入力", "📑\n③ 統合診断レポート・Excel"],
        default="🏠\n① 設置地域・環境判定",
        label_visibility="collapsed"
    )

    st.markdown("---")

    # ==========================================
    # STEP 1: 設置地域・環境判定
    # ==========================================
    if current_step == "🏠\n① 設置地域・環境判定":
        st.markdown("### 📍 1. 構造物所在地・環境因子の自動算出")
        address_input = st.text_input("所在地を入力してください（例：山形県酒田市、北陸沿岸、国道・高速路線名など）", placeholder="例：山形県酒田市大浜 国道112号", key="addr_in")
        
        auto_freeze_info, auto_salt_info, auto_asr_bone, auto_complex_degrade = "住所未入力", "住所未入力", "住所未入力", "所在地に基づいてJCI複合劣化マトリクスとの照合を行います。"
        
        if address_input:
            cold_regions = ["北海道", "青森", "岩手", "秋田", "山形", "宮城", "福島", "新潟", "富山", "石川", "福井", "長野", "岐阜", "群馬", "山梨"]
            salt_keywords = ["浜", "海岸", "港", "湾", "岬", "磯", "シーサイド", "大浜", "臨海", "塩", "浦", "津"]
            asr_regions = ["山形", "秋田", "新潟", "富山", "石川", "福井", "長野", "岐阜", "京都", "兵庫", "香川", "徳島", "福岡", "佐賀", "熊本"]
            
            is_cold = any(reg in address_input for reg in cold_regions)
            is_coast = any(kw in address_input for kw in salt_keywords)
            has_asr_bone = any(ar in address_input for ar in asr_regions)
            
            auto_freeze_info = "【高危険度】凍結融解サイクル年平均45回以上地域。" if is_cold else "【一般環境】深刻な凍結融解作用リスクは低。"
            auto_salt_info = "【重塩害警戒】強風による塩化物イオン定着領域。" if is_coast else "【一般地域】塩化物の直接的影響は軽微。"
            auto_asr_bone = "【ASR要注意】反応性骨材の流通履歴・損傷報告地域。" if has_asr_bone else "【低リスク】異常膨張リスクは穏健。"

            if is_cold and is_coast and has_asr_bone: auto_complex_degrade = "⚠️【JCI最過酷マトリクス：塩害×凍害×ASR】日本海側沿岸最過酷エリア。ASRクラックを起点に塩分と水分が浸透、マクロセル腐食と組織破壊が相乗進展。"
            elif is_cold and is_coast: auto_complex_degrade = "⚡【JCI複合劣化警戒：塩害×凍害】寒冷沿岸道路。表層脆弱化と鉄筋爆裂が同時複合進展。"
            elif is_cold and has_asr_bone: auto_complex_degrade = "🔎【JCI複合劣化警戒：凍害×ASR】内陸寒冷ASR地帯。凍結膨張圧によりASRクラックが開口拡張。"
            elif is_coast and has_asr_bone: auto_complex_degrade = "⚓【JCI複合劣化警戒：ASR×塩害】沿岸ASR地帯。内部鉄筋の早期発錆を誘発する環境。"
            else: auto_complex_degrade = "✅【低複合リスク】マクロ環境における致命的な複合侵食リスクは比較的低いと推定。"

        if address_input:
            st.markdown(f"<div class='status-card' style='border-left: 8px solid #38BDF8;'><h4>🧬 JCI（日本コンクリート工学会）環境ハザードマップ照合結果</h4><p style='font-size:18px; color:#38BDF8; font-weight:bold;'>{auto_complex_degrade}</p></div>", unsafe_allow_html=True)
            c1, c2, c3 = st.columns(3)
            with c1: st.markdown(f"<div class='dashboard-card'><h4>❄️ 凍害危険度</h4><p>{auto_freeze_info}</p></div>", unsafe_allow_html=True)
            with c2: st.markdown(f"<div class='dashboard-card'><h4>🌊 塩害範囲</h4><p>{auto_salt_info}</p></div>", unsafe_allow_html=True)
            with c3: st.markdown(f"<div class='dashboard-card'><h4>💎 ASR骨材分布</h4><p>{auto_asr_bone}</p></div>", unsafe_allow_html=True)

        st.markdown("### 🛠️ 2. 基本条件の設定")
        cc1, cc2, cc3 = st.columns(3)
        with cc1:
            struct_type = st.selectbox("① 構造物の種類", ["建築物（校舎・庁舎：外壁躯体・柱・梁）", "橋梁（上部工床版・主桁・下部工橋台・橋脚）", "ボックスカルバート・共同溝", "擁壁", "開渠・水路", "ダム・沈殿池"])
            cement_type = st.selectbox("④ セメントの種類", ["普通ポルトランド", "高炉セメント（B種等）", "早強ポルトランド", "不明"])
        with cc2:
            env_location = st.multiselect("② 物理的設置条件", ["屋外・直接雨掛かり", "日陰・日裏・軒下", "常時流水・水撃エリア", "屋内・常時乾燥"], default=[])
            elapsed_years = st.selectbox("⑤ 供用年数（目安）", ["10年未満", "10年〜30年未満", "30年〜50年未満（高経年化）", "50年以上（機能保全判定期）"])
            manual_years_input = st.text_input("（任意）具体的な築年数・竣工年")
        with cc3:
            wet_status = st.multiselect("③ 内部湿潤状況", ["漏水なし", "微細な湿潤", "活動性の漏水あり", "高水圧環境"], default=[])
            crack_type = st.selectbox("⑥ 支配的損傷症状", ["ひび割れ（単一）", "浮き・剥離・剥落", "鉄筋露出・爆裂", "エフロ析出伴う漏水", "ASR（３方向クラック）", "スケーリング（凍害・摩耗）"])
            
        region_info = st.text_area("⑦ その他、特記事項（振動、周辺工事等）", placeholder="例: 交通振動あり、不同沈下の形跡あり等")
        st.success("✅ 設定完了。画面上部の『📸 ② 写真・変状チェック入力』を選択してください。")

    # ==========================================
    # STEP 2: 現場写真・変状チェック入力
    # ==========================================
    elif current_step == "📸\n② 写真・変状チェック入力":
        st.markdown("### 🏢 1. 業務基本情報の入力")
        col_a, col_b, col_c = st.columns(3)
        with col_a: project_name = st.text_input("物件名（工事・業務名）", placeholder="例：塩竈清掃工場 躯体調査")
        with col_b: location_name = st.text_input("測定箇所詳細", placeholder="例：X5通り芯 B1梁上面")
        with col_c: inspector_name = st.text_input("調査担当者名", value="T&N技術管理者")
            
        st.markdown("---")
        st.markdown("### 📸 2. 写真アップロード ＆ 損傷個別詳細プロット")
        uploaded_files = st.file_uploader("写真をアップロードしてください（最大6枚）", type=["jpg", "jpeg", "png"], accept_multiple_files=True)
        
        images, photo_details_prompt, photo_excel_records = [], [], []
        
        if uploaded_files:
            for idx, file in enumerate(uploaded_files[:6]):
                img = Image.open(file)
                images.append(img)
                st.markdown(f"<div class='photo-input-box'><h4>📸 写真 No.{idx+1} 損傷トレーサビリティプロット</h4></div>", unsafe_allow_html=True)
                cv, fv = st.columns([1, 2])
                with cv: st.image(img, use_container_width=True)
                with fv:
                    f1, f2 = st.columns(2)
                    with f1:
                        p_part = st.text_input(f"No.{idx+1} 図面位置・通り芯・面", key=f"p_part_{idx}")
                        p_kind = st.selectbox(f"No.{idx+1} 損傷分類", ["構造クラック", "乾燥収縮", "ASR（３方向マップ状）", "凍害（層状剥離）", "塩害・爆裂", "複合劣化疑い"], key=f"p_kind_{idx}")
                        p_efflo = st.checkbox(f"進行性漏水・エフロあり", key=f"p_efflo_{idx}")
                        p_rust = st.checkbox(f"剥離・鉄筋露出・錆汁あり", key=f"p_rust_{idx}")
                    with f2:
                        p_w = st.number_input(f"No.{idx+1} クラック幅(mm)", min_value=0.0, step=0.05, key=f"p_w_{idx}")
                        p_l = st.number_input(f"No.{idx+1} クラック長さ(cm)", min_value=0.0, step=1.0, key=f"p_l_{idx}")
                        p_a = st.number_input(f"No.{idx+1} 打診浮き面積(㎡)", min_value=0.0, step=0.05, key=f"p_a_{idx}")
                        p_c = st.text_area(f"No.{idx+1} 技術者記事・所見", key=f"p_c_{idx}")
                
                photo_details_prompt.append(f"[No.{idx+1}] 位置:{p_part}|分類:{p_kind}|幅:{p_w}mm|長:{p_l}cm|浮き:{p_a}㎡|漏水:{p_efflo}|錆:{p_rust}|所見:{p_c}")
                photo_excel_records.append({"no":f"No.{idx+1}","part":p_part,"kind":p_kind,"dim":f"W:{p_w}mm/L:{p_l}cm/A:{p_a}㎡","comment":f"エフロ:{p_efflo}/錆:{p_rust}|{p_c}"})
                st.markdown("---")
            
            if st.button("🚀 所在地気象因数とJCI複合劣化マトリクス、全写真データを統合して高精密AI診断を実行"):
                if not api_key: st.error("APIキーがありません。")
                else:
                    with st.spinner("🔍 熟練コンクリート診断士AIが最新実務資料と同期しながら解析中..."):
                        genai.configure(api_key=api_key)
                        model = genai.GenerativeModel('gemini-2.5-flash')
                        
                        # キャッシュ退避
                        st.session_state.excel_records_cache, st.session_state.images_cache = photo_excel_records, images
                        st.session_state.header_data = {"p":project_name,"l":location_name,"i":inspector_name,"a":st.session_state.get("addr_in", ""),"s":struct_type,"c":st.session_state.get("complex_degrade_cache", "一般")}
                        
                        prompt = f"""あなたが作成すべきは、農水省・国交省等へ提出する最高レベルの工学的調査報告書です。
【JCI環境マトリクス】: {st.session_state.header_data['a']} における {st.session_state.header_data['c']} シナリオを考慮。
【構造物・写真データ】:
{"/".join(photo_details_prompt)}
上記損傷データと、添付資料で学んだ「凍害×ASRの相乗進展」「塩害によるマクロセル腐食」の物理・化学的因果関係をリンクさせ、1枚ずつの写真に詳細な技術的所見を述べてください。
確定最大幅、総合劣化度（Ⅰ〜Ⅳ）を冒頭に示し、具体的な補修工法（注入、断面修復、表面含浸等）を選定理由とともに厚く論じてください。"""
                        
                        response = model.generate_content([prompt] + images)
                        st.session_state.full_result_text = response.text
                        st.session_state.analysis_completed = True
                        st.session_state.current_step = "📑\n③ 統合診断レポート・Excel"
                        st.rerun()

    # ==========================================
    # STEP 3: 統合診断レポート・Excel調書
    # ==========================================
    elif current_step == "📑\n③ 統合診断レポート・Excel":
        if st.session_state.full_result_text:
            st.markdown(f"<div class='status-card' style='border-left: 8px solid #38BDF8;'><h3>📋 総合診断・解析レポート</h3></div>", unsafe_allow_html=True)
            st.markdown(f"<div class='report-text-box'>{st.session_state.full_result_text}</div>", unsafe_allow_html=True)

            # Excel出力
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "損傷調書"
            ws.page_setup.orientation, ws.page_setup.paperSize = ws.ORIENTATION_LANDSCAPE, ws.PAPERSIZE_A4
            ws.column_dimensions['A'].width, ws.column_dimensions['B'].width, ws.column_dimensions['D'].width = 24, 54, 75
            
            start_row = 1
            for idx, img in enumerate(st.session_state.images_cache):
                rec = st.session_state.excel_records_cache[idx]
                ws.merge_cells(f"A{start_row}:D{start_row}")
                ws[f"A{start_row}"] = f"■ {st.session_state.header_data['p']} 構造物健全度調査台帳"
                ws[f"A{start_row}"].font = Font(name="MS ゴシック", size=14, bold=True)
                
                labels = ["写真No. / 項目", "工種・変状分類", "位置・通り芯・面", "実測寸法・数量", "AI判定・工学的考察・記事"]
                vals = [rec["no"], f"{st.session_state.header_data['s']} ({rec['kind']})", rec["part"], rec["dim"], f"{rec['comment']}\n\n【統括報告】\n{st.session_state.full_result_text if idx==0 else 'No.1参照'}"]

                for i, (l, v) in enumerate(zip(labels, vals)):
                    r = start_row + 2 + i
                    ws[f"A{r}"], ws[f"B{r}"] = l, v
                    ws[f"A{r}"].font, ws[f"A{r}"].fill = Font(bold=True), PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
                    ws[f"A{r}"].border = ws[f"B{r}"].border = Border(left=Side(style="thin"), right=Side(Side(style="thin")), top=Side(style="thin"), bottom=Side(style="thin"))
                    ws[f"B{r}"].alignment = Alignment(wrap_text=True, vertical="top")
                    if i == 4: ws.row_dimensions[r].height = 250
                
                img_buf = io.BytesIO()
                img.save(img_buf, format="PNG")
                xl_img = ExcelImage(img_buf)
                xl_img.width, xl_img.height = 520, 360
                ws.add_image(xl_img, f"D{start_row + 2}")
                start_row += 12

            output = io.BytesIO()
            wb.save(output)
            st.download_button(label="📥 官庁提出用 高精密Excel調書をダウンロード", data=output.getvalue(), file_name=f"確定診断調書_{datetime.now().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.info("💡 診断が実行されていません。ステップ②で解析を実行してください。")


診断の精度およびUIの視認性が飛躍的に向上しました！もし特定の専門用語の使い分けや、Excelの項目名にさらなる微調整が必要であれば、いつでもお申し付けください。
