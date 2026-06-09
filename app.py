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

# --- 1. ページ設定と超ワイド・実務特化ダッシュボードCSS ---
st.set_page_config(page_title="T&N コンクリート劣化診断 AI Suite Pro", layout="wide")

# セッション状態の完全初期化
if 'full_result_text' not in st.session_state:
    st.session_state.full_result_text = None
if 'final_width' not in st.session_state:
    st.session_state.final_width = 0.0
if 'analysis_completed' not in st.session_state:
    st.session_state.analysis_completed = False
if 'current_step' not in st.session_state:
    st.session_state.current_step = "ステップ①"

# 解析完了時にステップ③の選択枠をネオンブルーで動的明滅させる誘導CSS
animation_css = ""
if st.session_state.analysis_completed:
    animation_css = """
    div[data-testid="stPills"] div[role="radiogroup"] div:nth-child(3) label {
        animation: pulse_glow_fixed 1.4s infinite alternate !important;
        border: 2px solid #38BDF8 !important;
    }
    @keyframes pulse_glow_fixed {
        0% { box-shadow: 0 4px 15px rgba(56, 189, 248, 0.2); background-color: #1E293B !important; }
        100% { box-shadow: 0 0 25px #38BDF8, inset 0 0 12px #38BDF8; background-color: #0284C7 !important; }
    }
    """

st.markdown(f"""
<style>
.main {{ background-color: #0F172A; color: #FFFFFF; }}
.stApp {{ background-color: #0F172A; }}

/* 画面を圧迫する左側サイドバーおよび不要なシステム文字の完全非表示化 */
section[data-testid="stSidebar"] {{ display: none !important; }}
div[data-testid="collapsedControl"] {{ display: none !important; }}

/* 選択ツール上部の不要なStreamlit既定文字を非表示化 */
div[data-testid="stPills"] label[data-testid="stWidgetLabel"] {{
    display: none !important;
    height: 0px !important;
    margin: 0 !important;
    padding: 0 !important;
}}

/* フォント表示の最適化 */
h1, h2, h3, h4, h5, h6, p, span, label, .stMarkdown,
.stCheckbox label, div[data-testid="stMarkdownContainer"] p {{
    color: #FFFFFF !important; font-family: 'Helvetica Neue', Arial, sans-serif; font-weight: bold !important;
}}

/* 入力・選択コンポーネントの黒文字視認性確保 */
input, textarea, select, div[data-baseweb="select"] * {{ color: #0F172A !important; font-weight: bold !important; }}
input::placeholder, textarea::placeholder {{ color: #64748B !important; opacity: 1 !important; }}

div[data-testid="stMultiSelect"] span[data-baseweb="tag"] * {{
    color: #0F172A !important;
    font-weight: bold !important;
}}

/* チェックボックスを信頼の青に統一（赤みの完全根絶） */
div[data-testid="stCheckbox"] div[role="checkbox"] {{
    border-color: #475569 !important;
}}
div[data-testid="stCheckbox"] div[role="checkbox"][aria-checked="true"] {{
    background-color: #0284C7 !important;
    border-color: #38BDF8 !important;
    box-shadow: 0 0 10px rgba(56, 189, 248, 0.5) !important;
}}
div[data-testid="stCheckbox"] div[role="checkbox"] svg {{
    stroke: #FFFFFF !important;
    fill: none !important;
}}

/* 【バグ完全修正】重なりを解決し、実務に適した巨大な3セクション切り替えボックスを新設 */
div[data-testid="stPills"] div[role="radiogroup"] {{
    display: flex !important;
    gap: 20px !important;
    width: 100% !important;
    background-color: transparent !important;
    padding-top: 10px !important;
}}
div[data-testid="stPills"] div[role="radiogroup"] > div {{
    flex: 1 !important;
}}
div[data-testid="stPills"] div[role="radiogroup"] label {{
    background: linear-gradient(135deg, #1E293B, #111827) !important;
    border: 2px solid #334155 !important;
    border-radius: 14px !important;
    padding: 24px 15px !important;
    text-align: center !important;
    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.3) !important;
    cursor: pointer !important;
    display: block !important;
    width: 100% !important;
}}
div[data-testid="stPills"] div[role="radiogroup"] label:hover {{
    transform: translateY(-3px) !important;
    border-color: #475569 !important;
}}

/* ボックス内のテキスト書式強制固定 */
div[data-testid="stPills"] div[role="radiogroup"] label * {{
    font-size: 20px !important;
    font-weight: 800 !important;
    color: #FFFFFF !important;
    line-height: 1.4 !important;
}}

/* 現在選択されているボックスの強調・発光 */
div[data-testid="stPills"] div[role="radiogroup"] input[type="radio"]:checked + div label {{
    background: linear-gradient(135deg, #0284C7, #1E293B) !important;
    border-color: #38BDF8 !important;
    box-shadow: 0 0 20px rgba(56, 189, 248, 0.4) !important;
}}
div[data-testid="stPills"] div[role="radiogroup"] input[type="radio"]:checked + div label * {{
    color: #38BDF8 !important;
}}

div[data-testid="stPills"] input[type="radio"] {{
    display: none !important;
}}

{animation_css}

/* 解析ボタン専用スタイル */
div.stButton > button {{
    background: #0284C7 !important;
    color: #FFFFFF !important;
    border: 2px solid #38BDF8 !important;
    border-radius: 10px !important;
    padding: 16px 24px !important;
    font-size: 18px !important;
    font-weight: bold !important;
    width: 100% !important;
}}
div.stButton > button:hover {{
    background: #38BDF8 !important;
    box-shadow: 0 0 20px #38BDF8 !important;
}}

/* 各種コンテナ */
.dashboard-card {{ padding: 25px; background-color: #1E293B; border-radius: 16px; border: 1px solid #334155; margin-bottom: 20px; }}
.status-card {{ padding: 25px; background-color: #1E293B; border-radius: 16px; margin-bottom: 20px; border-top: 1px solid #334155; border-right: 1px solid #334155; border-bottom: 1px solid #334155; }}
.photo-input-box {{ background-color: #1E293B; padding: 20px; border-radius: 12px; border: 1px dashed #38BDF8; margin-bottom: 20px; }}

.report-text-box {{
    background-color: #1E293B !important;
    color: #FFFFFF !important;
    border: 1px solid #475569;
    border-radius: 12px;
    padding: 25px;
    font-size: 16px;
    line-height: 1.8;
    white-space: pre-wrap;
}}
</style>
""", unsafe_allow_html=True)

# 閉域セキュリティパスワード認証
def check_password():
    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False
    def password_entered():
        if st.session_state["password"] == "tn0000":
            st.session_state["authenticated"] = True
            del st.session_state["password"]
        else:
            st.error("❌ アクセスパスワードが違います")
    if not st.session_state["authenticated"]:
        if os.path.exists("logo.png"): 
            st.image("logo.png", width=250)
        st.markdown("<h2>🔒 コンクリート劣化診断 AI Suite Pro（関係者専用）</h2>", unsafe_allow_html=True)
        st.text_input("パスワードを入力してください", type="password", on_change=password_entered, key="password")
        return False
    return True

if check_password():
    api_key = st.secrets.get("GEMINI_API_KEY", "")

    # 業務ヘッダー
    col_l, col_r = st.columns([1, 5])
    with col_l:
        if os.path.exists("logo.png"): st.image("logo.png", width=150)
    with col_r:
        st.markdown("<h1 style='margin-top:5px; margin-bottom:0;'>コンクリート構造物高精密総合診断システム</h1>", unsafe_allow_html=True)
        st.markdown("<p style='color: #38BDF8; font-size:15px; margin-top:0;'>農林水産省機能保全手引き・日本コンクリート工学会複合劣化判定マップ完全準拠</p>", unsafe_allow_html=True)

    # 上部3セクション切り替えボックス（テキストの重なりを完全修復）
    current_step = st.pills(
        label="",
        options=["ステップ①：設置地域・環境判定", "ステップ②：現場写真・変状チェック入力", "ステップ③：統合診断レポート・Excel調書"],
        default="ステップ①：設置地域・環境判定",
        label_visibility="collapsed"
    )

    st.markdown("---")

    # ==========================================
    # SCREEN 1: 設置地域・環境判定
    # ==========================================
    if current_step == "ステップ①：設置地域・環境判定":
        st.markdown("### 📍 1. 対象構造物の所在地入力（環境因子の自動算出）")
        address_input = st.text_input("構造物の設置住所・施設名を入力してください", placeholder="例：山形県酒田市大浜 国道112号", key="addr_in")
        
        auto_freeze_info = "住所未入力"
        auto_salt_info = "住所未入力"
        auto_asr_bone = "住所未入力"
        auto_complex_degrade = "住所が入力されると、JCI複合劣化ハザードマップ（塩害×凍害、凍害×ASR、ASR×塩害）との適合性を自動判定します。"
        
        if address_input:
            cold_regions = ["北海道", "青森", "岩手", "秋田", "山形", "宮城", "福島", "新潟", "富山", "石川", "福井", "長野", "岐阜", "群馬", "山梨"]
            salt_keywords = ["浜", "海岸", "港", "湾", "岬", "磯", "シーサイド", "大浜", "臨海", "塩", "浦", "津"]
            asr_regions = ["山形", "秋田", "新潟", "富山", "石川", "福井", "長野", "岐阜", "京都", "兵庫", "香川", "徳島", "福岡", "佐賀", "熊本"]
            
            is_cold = any(reg in address_input for reg in cold_regions)
            is_coast = any(kw in address_input for kw in salt_keywords)
            has_asr_bone = any(ar in address_input for ar in asr_regions)
            
            if is_cold: auto_freeze_info = "【高危険度】冬季凍結融解サイクル年平均45回以上地域。膨張圧リスク高。"
            else: auto_freeze_info = "【一般環境】深刻な凍結融解作用を受ける確率は比較的低いエリア。"
                
            if is_coast: auto_salt_info = "【重塩害警戒（S地域）】沿岸近傍。強風による高濃度飛来塩分の定着危険地帯。"
            elif is_cold and any(road in address_input for road in ["国道", "高速", "インター", "JCT", "バイパス", "道"]):
                auto_salt_info = "【塩害警戒（融雪剤散布路線）】寒冷地幹線道路。凍結防止剤散布による塩化物イオン供給環境。"
            else: auto_salt_info = "【一般地域】塩化物の直接的影響を受けにくい内陸エリア。"

            if has_asr_bone: auto_asr_bone = "【ASR要注意地域】アルカリシリカ反応性骨材の過去の流通履歴・損傷報告地域。"
            else: auto_asr_bone = "【低リスク】ASRに起因する異常膨張の潜在リスクは比較的穏健なエリア。"

            if is_cold and is_coast and has_asr_bone:
                auto_complex_degrade = "⚠️【JCIトリプル複合劣化：塩害×凍害×ASR】日本海側沿岸最過酷マトリクス。ASRクラックを起点に塩分と水分が浸透、マクロセル腐食と組織スケーリングが相乗進展する最過酷領域。"
            elif is_cold and is_coast:
                auto_complex_degrade = "⚡【JCI複合劣化：塩害×凍害】寒冷沿岸、または融雪剤散布を受ける道路橋。表層脆弱化（スケーリング）と鉄筋腐食（爆裂現象）が同時複合進展。"
            elif is_cold and has_asr_bone:
                auto_complex_degrade = "🔎【JCI複合劣化：凍害×ASR】内陸寒冷ASR地帯。ゲルの吸水膨張クラックへ冬季水分が浸入し、凍結膨張圧で開口が拡張する複合スパイラル領域。"
            elif is_coast and has_asr_bone:
                auto_complex_degrade = "⚓【JCI複合劣化：ASR×塩害】沿岸ASR地帯。ASRクラックから外部塩化物イオンが超急速に内部拡散し、早期鉄筋発錆を引き起こす環境。"
            else:
                auto_complex_degrade = "✅【単一/低複合リスク】現時点の立地条件から、JCI指針に定められる致命的な複合侵食リスクは比較的低いと推定されます。"

        if address_input:
            st.markdown(f"<div class='status-card' style='border-left: 8px solid #38BDF8;'><h4>🧬 日本コンクリート工学会（JCI）複合劣化ハザードマップ自動照合</h4><p style='font-size:15px; color:#38BDF8; font-weight:bold; margin-bottom:0;'>{auto_complex_degrade}</p></div>", unsafe_allow_html=True)
            c1, c2, c3 = st.columns(3)
            with c1: st.markdown(f"<div class='dashboard-card'><h4>❄️ 凍害危険度（分布）</h4><p style='font-size:14px; color:#CBD5E1;'>{auto_freeze_info}</p></div>", unsafe_allow_html=True)
            with c2: st.markdown(f"<div class='dashboard-card'><h4>🌊 塩害範囲（道路橋・沿岸）</h4><p style='font-size:14px; color:#CBD5E1;'>{auto_salt_info}</p></div>", unsafe_allow_html=True)
            with c3: st.markdown(f"<div class='dashboard-card'><h4>💎 ASR反応性骨材分布</h4><p style='font-size:14px; color:#CBD5E1;'>{auto_asr_bone}</p></div>", unsafe_allow_html=True)

        st.markdown("### 🛠️ 2. 構造物基本条件の設定（旧サイドバー項目統合）")
        cc1, cc2, cc3 = st.columns(3)
        with cc1:
            struct_type = st.selectbox("① 構造物の種類", [
                "（未選択・写真から自動判定）", "建築物（校舎・庁舎・屋内運動場：外壁躯体・柱・梁）", "橋梁（上部工床版・主桁・下部工橋台・橋脚・遊間伸縮部）", "ボックスカルバート・縦型土留共同溝・水路暗渠", "擁壁（重力式擁壁・鉄筋コンクリートL型擁壁）", "開渠・水路構造物・開閉所変電基礎", "ダム（重力式コンクリートダム・アーチダム・越流部躯体）"
            ])
            cement_type = st.selectbox("④ 使用セメントの種類", ["（未選択）", "普通ポルトランドセメント", "高炉セメント（B種等：複合劣化・塩害・ASR抑制）", "早強ポルトランドセメント", "不明（経年構造物）"])
        with cc2:
            env_location = st.multiselect("② 物理的設置条件（複数選択可）", ["屋外・直接雨掛かり（乾湿の繰り返し）", "日陰・日裏・軒下（湿気滞留環境）", "常時流水・水撃を受ける環境（流水摩耗・スケーリング）", "屋内・常時乾燥（中性化環境）"], default=[])
            elapsed_years = st.selectbox("⑤ 供用年数（大目安）", ["（未選択）", "10年未満", "10年以上〜30年未満", "30年以上〜50年未満（高経年化）", "50年以上（農水・国交省機能保全判定期）"])
            manual_years_input = st.text_input("（任意項目）具体的な築年数・竣工年（例：築42年、1984年竣工）")
        with cc3:
            wet_status = st.multiselect("③ 内部湿潤・漏水状況（複数選択可）", ["漏水なし（常時乾燥）", "微細な湿潤（変色・湿気あり）", "活動性の漏水あり（水みち形成・進行性エフロ）", "高水圧環境（ダム・沈殿池等の浸透圧環境）"], default=[])
            crack_type = st.selectbox("⑥ 目視・打診での支配的損傷", ["（未選択・写真から自動判定）", "ひび割れ（単一クラック・構造応力/乾燥収縮）", "浮き・剥離・コンクリート塊の剥落（打診中空音）", "鉄筋露出・爆裂現象（内部腐食錆水の滲出伴う）", "エフロレッセンス（遊離石灰）の析出を伴う活動性漏水", "亀甲状ひび割れ（ASR・アルカリ骨材反応の３方向マップ状クラック）", "スケーリング・微細組織のうろこ状脆弱化（凍害・物理的摩耗）"])
            
        region_info = st.text_area("⑦ その他、現場特記・周辺状況", placeholder="例: 近傍に大型車両の交通量が多く微振動あり、等")
        st.success("✅ 設定完了：画面最上部へスクロールし、中央の『ステップ②：現場写真・変状チェック入力』ボックスを選択してください。")

    # ==========================================
    # SCREEN 2: 現場写真・変状チェック入力
    # ==========================================
    elif current_step == "ステップ②：現場写真・変状チェック入力":
        st.markdown("### 🏢 1. 提出用 業務基本情報の入力")
        col_a, col_b, col_c = st.columns(3)
        with col_a: project_name = st.text_input("項目A：物件名（工事名・業務名）", placeholder="（例：塩竈清掃工場 躯体調査）")
        with col_b: location_name = st.text_input("項目B：調査位置・測定箇所詳細", placeholder="（例：沈殿池 南面壁）")
        with col_c: inspector_name = st.text_input("項目C：調査担当者（コンクリート診断士名）", value="T&N技術管理者")
            
        st.markdown("---")
        st.markdown("### 🔧 2. 技術者による構造・施工要因の補足チェック")
        ch1, ch2, ch3 = st.columns(3)
        with ch1:
            cb_shear = st.checkbox("構造的要因・不同沈下・土圧・耐震性能上のせん断ひび割れ疑い")
            cb_janka = st.checkbox("施工起因のジャンカ・コールドジョイント・初期欠陥あり")
        with ch2:
            cb_wet_h = st.checkbox("常時湿潤・漏水（ASR・遊離石灰溶出リスクの補強）")
            cb_cover = st.checkbox("設計かぶり厚の不足または中性化の鉄筋位置到達の疑い")
        with ch3:
            cb_joint = st.checkbox("施工目地・伸縮目地・伸縮装置周辺部からの漏水・損傷進展")
            
        selected_factors = []
        if cb_shear: selected_factors.append("不同沈下・土圧・構造応力または地動起因のせん断ひび割れ（耐震性能影響）")
        if cb_janka: selected_factors.append("施工起因のジャンカ・初期欠陥の目視確認あり")
        if cb_wet_h: selected_factors.append("常時湿潤・高水圧・漏水（ASR・遊離石灰溶出リスク大）")
        if cb_cover: selected_factors.append("設計かぶり厚の不足または中性化の鉄筋位置到達")
        if cb_joint: selected_factors.append("施工目地・打継目地部からの滞水・漏水進展")
        human_factors_text = "、".join(selected_factors) if selected_factors else "特になし"

        st.markdown("---")
        st.markdown("### 📸 3. 現場写真アップロード ＆ 個別損傷トレーサビリティプロット")
        uploaded_files = st.file_uploader("写真をアップロードしてください（最大6枚・一括Excel調書出力対応）", type=["jpg", "jpeg", "png"], accept_multiple_files=True)
        
        images = []
        photo_details_prompt = []
        photo_excel_records = []
        
        if uploaded_files:
            if len(uploaded_files) > 6:
                st.warning("⚠️ 最初の6枚のみを処理します。")
                uploaded_files = uploaded_files[:6]
                
            for idx, file in enumerate(uploaded_files):
                img = Image.open(file)
                images.append(img)
                
                st.markdown(f"<div class='photo-input-box'><h4>📸 写真 No.{idx+1} 損傷詳細プロット（通り芯・打診数量連動）</h4></div>", unsafe_allow_html=True)
                col_img_view, col_form_view = st.columns([1, 2])
                with col_img_view: st.image(img, use_container_width=True)
                
                with col_form_view:
                    f_col1, f_col2 = st.columns(2)
                    with f_col1:
                        p_part = st.text_input(f"No.{idx+1} 損傷図面位置・通り芯・部材番号（必須）", key=f"p_part_{idx}", placeholder="例: X5通り芯 B1梁上面、橋台ハンチ左面")
                        p_crack_kind = st.selectbox(f"No.{idx+1} ひび割れ・変状パターンの特定", ["構造クラック（応力・せん断方向ひび割れ）", "乾燥収縮クラック（経年・環境性ひび割れ）", "アルカリ骨材反応（ASR：３方向マップ状・ゲル析出を伴う）", "凍害（層状剥離・うろこ状脆弱化・ポップアウト）", "塩害・融雪剤侵食（鉄筋に沿った直線ひび割れ・爆裂現象）", "複合劣化疑い（2因子以上の変状が同時重畳している状態）"], key=f"p_kind_{idx}")
                        p_efflo = st.checkbox(f"進行性漏水・白華（エフロレッセンス）を随伴する", key=f"p_efflo_{idx}")
                        p_rust = st.checkbox(f"断面剥離・鉄筋露出・錆汁の滲出を随伴する", key=f"p_rust_{idx}")
                    with f_col2:
                        p_width = st.number_input(f"No.{idx+1} 実測クラック幅 (mm)", min_value=0.0, step=0.05, value=0.0, key=f"p_width_{idx}")
                        p_length = st.number_input(f"No.{idx+1} 実測クラック長さ (cm)", min_value=0.0, step=1.0, value=0.0, key=f"p_length_{idx}")
                        p_area = st.number_input(f"No.{idx+1} 打診による浮き・剥離推定数量 (㎡)", min_value=0.00, step=0.05, value=0.00, key=f"p_area_{idx}")
                        p_comment = st.text_area(f"No.{idx+1} 📝 記事・打診所見・補修設計への補足", key=f"p_comm_{idx}", placeholder="例: 伸縮継手部からの水みち形成を確認。打診にて浮きを確認。修復設計数量へ反映のこと。")
                
                efflo_str = "確認" if p_efflo else "なし"
                rust_str = "確認" if p_rust else "なし"
                detail_txt = f"[写真No.{idx+1}] 位置通り芯:{p_part} | JCI分類:{p_crack_kind} | 幅:{p_width}mm | 長さ:{p_length}cm | 打診浮き数量:{p_area}㎡ | 漏水エフロ:{efflo_str} | 鉄筋露出錆:{rust_str} | 技術者記事:{p_comment}"
                photo_details_prompt.append(detail_txt)
                
                photo_excel_records.append({
                    "no": f"No.{idx+1}",
                    "part": p_part if p_part else f"現場撮影箇所 {idx+1}",
                    "kind": p_crack_kind,
                    "dim": f"W={p_width}mm / L={p_length}cm / 浮き={p_area}㎡" if (p_width > 0 or p_length > 0 or p_area > 0) else "基準物なし（現地実測推奨）",
                    "comment": f"【随伴】エフロ:{efflo_str} / 鉄筋露出:{rust_str}。 所見: {p_comment}"
                })
                st.markdown("---")
            
            # 高精密解析ボタン
            if st.button("🚀 所在地気象因数とJCI複合劣化マトリクス、全写真データを統合して高精密AI診断を実行"):
                st.session_state.analysis_completed = False
                st.session_state.full_result_text = None
                
                if not api_key:
                    st.error("APIキーが設定されていません。")
                else:
                    with st.spinner("🔍 熟練コンクリート診断士AIがJCI複合劣化ハザードマップ・各種実績成果と完全同期しながら精密解析中..."):
                        max_retries = 3
                        for attempt in range(max_retries):
                            try:
                                genai.configure(api_key=api_key)
                                model = genai.GenerativeModel('gemini-2.5-flash')
                                
                                # 入力退避キャッシュ
                                st.session_state.excel_records_cache = photo_excel_records
                                st.session_state.struct_type_cache = struct_type
                                st.session_state.project_name_cache = project_name
                                st.session_state.location_name_cache = location_name
                                st.session_state.inspector_name_cache = inspector_name
                                st.session_state.images_cache = images
                                
                                addr_in_val = st.session_state.get("addr_in", "未指定")
                                st.session_state.address_input_cache = addr_in_val
                                
                                cold_regions = ["北海道", "青森", "岩手", "秋田", "山形", "宮城", "福島", "新潟", "富山", "石川", "福井", "長野", "岐阜", "群馬", "山梨"]
                                salt_keywords = ["浜", "海岸", "港", "湾", "岬", "磯", "シーサイド", "大浜", "臨海", "塩", "浦", "津"]
                                asr_regions = ["山形", "秋田", "新潟", "富山", "石川", "福井", "長野", "岐阜", "京都", "兵庫", "香川", "徳島", "福岡", "佐賀", "熊本"]
                                
                                is_cold = any(reg in addr_in_val for reg in cold_regions)
                                is_coast = any(kw in addr_in_val for kw in salt_keywords)
                                has_asr_bone = any(ar in addr_in_val for ar in asr_regions)
                                
                                calc_complex = "一般環境"
                                if is_cold and is_coast and has_asr_bone: calc_complex = "⚠️【JCIトリプル複合劣化警報：塩害×凍害×ASR】日本海側沿岸最過酷マトリクス。ASRクラックを起点に塩分と水分が浸透、マクロセル腐食と組織スケーリングが相乗進展。"
                                elif is_cold and is_coast: calc_complex = "⚡【JCI複合劣化要警戒：塩害×凍害】寒冷沿岸、または融雪剤散布を受ける道路橋。表層脆弱化と鉄筋腐食（爆裂）が複合進展。"
                                elif is_cold and has_asr_bone: calc_complex = "🔎【JCI複合劣化要警戒：凍害×ASR】内陸寒冷ASR地帯。ゲルの吸水膨張クラックへ冬季水分が浸入し凍結膨張圧で開口拡張。"
                                elif is_coast and has_asr_bone: calc_complex = "⚓【JCI複合劣化要警戒：ASR×塩害】沿岸ASR地帯。ASRクラックから外部塩化物イオンが超急速に内部拡散し早期鉄筋発錆を引き起こす環境。"
                                
                                st.session_state.complex_degrade_cache = calc_complex
                                
                                env_text = "、".join(env_location) if env_location else "指定なし"
                                wet_text = "、".join(wet_status) if wet_status else "指定なし"
                                photo_details_joined = "\n".join(photo_details_prompt)
                                
                                # 学習マニュアル・実績を網羅した最高精密プロンプト
                                prompt = f"""
あなたは日本最高峰の「コンクリート診断士」であり、日本コンクリート工学会（JCI）の「複合劣化コンクリート構造物の評価と維持管理計画研究委員会 報告書」のハザードマップ・相乗進展メカニズム、ならびに農水省・国交省の機能保全・点検マニュアルを完全にマスターしている専門家です。
テンプレート記述や汎用回答を完全に排除し、提示されたシステム自動算出のJCI環境データ、および【技術者が写真ごとに通り芯・図面位置・実測寸法・打診数量を精緻に入力したデータ】を極めて高度にビジョンリンク解析し、完全オーダーメイドの公式報告書をゼロから起稿してください。

【対象構造物の所在地・JCI複合劣化環境マトリクスデータ（システム自動算出）】
- 住所・施設名: {addr_in_val}
- 判定されたJCIマクロ劣化環境・リスク度シナリオ: {calc_complex}
- 地域凍害危険度分布: {"高危険度" if is_cold else "一般"}
- 地域塩害警戒範囲（道路橋・飛来塩分）: {"警戒地域" if is_coast else "一般"}
- 反応性骨材分布・ASR損傷報告エリア: {"分布確認" if has_asr_bone else "一般"}

【構造物全体の条件設定】
- 構造物種別: {struct_type}
- 物理的設置条件: {env_text}
- 内部湿潤・漏水状況: {wet_text}
- 使用セメント / 供用年数: {cement_type} / {elapsed_years} ({manual_years_input})
- 主たる劣化症状（マクロ）: {crack_type}
- 構造・施工要因の補足: {human_factors_text}

【📸 アップロードされた各写真の個別高精度パラメータ（通り芯・打診数量連動解析）】
{photo_details_joined}

【絶対厳守命令】
1. 写真データの中にクラックスケールが無く、かつ上記の写真個別パラメータでも幅・長さが「0.0」となっている場合は、絶対に寸法を数値としてハルシネーション（捏造）しないでください。その場合は必ず文章の冒頭で「【寸法判定保留】写真から正確な縮尺基準が確認できず実測値も無いため、数値推測を保留します。正確な劣化度評価のため縮尺基準の提供を求めます」と記載し、ユーザーへ逆質問してください。入力がある場合はその確定数値を論拠にしてください。
2. 判定基準として、JCI複合劣化指針および各点検マニュアルに則り、ひび割れ幅、漏水、エフロ、浮き剥離、鉄筋露出の有無、および2因子以上の複合重畳性を総合評価し、以下の4区分から該当する【劣化度】を必ず選定・明記してください。
   ・劣化度Ⅰ（軽微・経過観察）: ひび割れ幅0.2mm未満、漏水・錆汁なし。単一の初期欠陥・乾燥収縮等。表面含浸等の予防保全。
   ・劣化度Ⅱ（中期・要補修）: ひび割れ幅0.2mm以上1.0mm未満、またはエフロ析出。あるいは軽微な2因子の複合初期症状。低圧エポキシ樹脂注入工法。
   ・劣化度Ⅲ（重大・早期補修）: ひび割れ幅1.0mm以上、または活動性漏水、剥離、鉄筋露出（爆裂）。JCI複合劣化の相乗進展（塩害×凍害、凍害×ASR、ASR×塩害など）が明瞭に確認される状態。ポリマーセメントモルタル充填工法（断面修復工法）。
   ・劣化度Ⅳ（激甚・緊急対応）: 構造変形、耐震性能を揺るがす深刻なせん断クラック、激しい漏水・滞水、部材の機能喪失。
3. 考察には専門用語（鉄筋の不動態被膜破壊、マクロセル腐食、アルカリシリカゲル、遊離石灰の溶出、膨張圧、流水による物理的摩耗、複合スパイラル進展、剪断応力など）を適切に用い、単一侵食ではなく、JCI報告書が定義する『複合劣化の相乗メカニズム（例：ASRひび割れからの塩化物イオン拡散加速など）』の因果関係を詳細に展開してください。

出力は、確認できた場合のみ「確定最大ひび割れ幅: 〇.〇 mm / 総合劣化度: 〇」を必ず冒頭の1行目に示し、その後【JCI指針に基づく工学的複合劣化原因の深い推測】、【各写真（No.1〜）に対する個別の詳細技術的所見】、【推奨される具体的な補修工法とその選定理由】、【健全度・複合進行度特定のための内部詳細調査（シュミットハンマー、コア採取による圧縮強度・静弾性係数、全塩化物イオン量測定、ドリル削孔による中性化深さ試験等）の推奨】を、そのまま官庁に提出できる極厚な長文（各400文字以上）で出力してください。JSONは不要です。
"""
                                request_contents = [prompt] + images
                                response = model.generate_content(request_contents)
                                st.session_state.full_result_text = response.text
                                
                                final_w = p_width
                                if final_w == 0:
                                    try:
                                        match = re.search(r"最大ひび割れ幅:\s*([0-9.]+)", st.session_state.full_result_text)
                                        if match: final_w = float(match.group(1))
                                    except Exception: final_w = 0.0
                                st.session_state.final_width = final_w
                                st.session_state.analysis_completed = True
                                st.success("🎉 高精密AI解析が正常完了しました。画面上部の『ステップ③：統合診断レポート・Excel調書』ボックスを選択してください。")
                                break 
                            except Exception as e:
                                if "429" in str(e) or "quota" in str(e).lower():
                                    if attempt < max_retries - 1:
                                        time.sleep(2.5)
                                        continue
                                raise e
        else:
            st.info("ℹ️ 写真がアップロードされていません。現場写真をドロップしてください。")

    # ==========================================
    # SCREEN 3: 統合診断レポート・Excel調書
    # ==========================================
    elif current_step == "ステップ③：統合診断レポート・Excel調書":
        if st.session_state.full_result_text:
            fw = st.session_state.final_width
            if "劣化度Ⅲ" in st.session_state.full_result_text or "劣化度Ⅳ" in st.session_state.full_result_text or fw >= 1.0:
                color_code, status_title = "#EF4444", f"🔴 【劣化度Ⅲ〜Ⅳ：複合劣化相乗進展・早期補修対応】判定最大幅: {fw} mm"
                alert_desc = "⚠️ JCI・指針基準：2因子以上の侵食が相乗的に重畳し劣化期へ突入しています。早期の断面修復工法および構造安全性の確保、追跡詳細調査が必須です。"
            elif "劣化度Ⅱ" in st.session_state.full_result_text or (0.2 <= fw < 1.0):
                color_code, status_title = "#EAB308", f"🟡 【劣化度Ⅱ：中期劣化・機能保持補修】判定最大幅: {fw} mm"
                alert_desc = "💡 JCI・指針基準：環境因子による有害な損傷、または複合劣化の初期兆候を検知しました。「低圧エポキシ樹脂注入工法」等の選定を推奨します。"
            elif fw > 0:
                color_code, status_title = "#10B981", f"🟢 【劣化度Ⅰ：単一微細損傷・経過観察フェーズ】判定最大幅: {fw} mm"
                alert_desc = "✅ JCI・指針基準：安全性への直接的影響は軽微です。表面含浸工法による予防保全、または目視経過観察となります。"
            else:
                color_code, status_title = "#3B82F6", "🔵 【寸法・複合劣化度判定保留：現地実測要請】"
                alert_desc = "ℹ️ 写真および個別入力欄から正確な縮尺基準が確認できないため判定を保留しています。現地実測値または縮尺基準を確認してください。"
            
            st.markdown(f"<div class='status-card' style='border-left: 8px solid {color_code};'><h3 style='color: {color_code} !important; margin:0; font-size:22px;'>{status_title}</h3><p style='color: #F1F5F9 !important; font-size: 14px; margin: 8px 0 0 0; font-weight: bold;'>{alert_desc}</p></div>", unsafe_allow_html=True)
            st.markdown("<h4 style='color: white; margin-top:20px;'>📑 AI Suite Pro 高精密工学的統合解析レポート</h4>", unsafe_allow_html=True)
            st.markdown(f"<div class='report-text-box'>{st.session_state.full_result_text}</div>", unsafe_allow_html=True)

            # --- Excel調書出力モジュール ---
            try:
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "コンクリート構造物健全度調書"
                ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
                ws.page_setup.paperSize = ws.PAPERSIZE_A4
                ws.views.sheetView[0].showGridLines = False

                font_header = Font(name="MS ゴシック", size=14, bold=True)
                font_label = Font(name="MS ゴシック", size=11, bold=True)
                font_data = Font(name="MS ゴシック", size=11)
                
                thin_side = Side(border_style="thin", color="000000")
                border_cell = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
                fill_label = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

                ws.column_dimensions['A'].width = 24
                ws.column_dimensions['B'].width = 54  
                ws.column_dimensions['C'].width = 2
                ws.column_dimensions['D'].width = 72  
                
                p_name_xl = st.session_state.project_name_cache if st.session_state.project_name_cache else "コンクリート構造物健全度調査業務"
                l_name_xl = st.session_state.location_name_cache if st.session_state.location_name_cache else "現場調査対象箇所"
                if st.session_state.address_input_cache: p_name_xl = f"{st.session_state.address_input_cache} {p_name_xl}"

                start_row = 1
                for idx, img in enumerate(st.session_state.images_cache):
                    rec = st.session_state.excel_records_cache[idx]
                    
                    ws.merge_cells(f"A{start_row}:D{start_row}")
                    ws[f"A{start_row}"] = f"■ {p_name_xl} 構造物調査状況写真台帳"
                    ws[f"A{start_row}"].font = font_header
                    
                    ws.merge_cells(f"A{start_row+1}:D{start_row+1}")
                    ws[f"A{start_row+1}"] = f"施設位置・通り芯： {l_name_xl}  |  調査担当：{st.session_state.inspector_name_cache}"
                    ws[f"A{start_row+1}"].font = font_label
                    ws[f"A{start_row+1}"].alignment = Alignment(horizontal="right")

                    info_labels = ["写真No. / 撮影項目", "工種・部材・変状分類", "位置・通り芯プロット面", "手動入力・実測寸法数量", "JCI複合劣化マトリクス特性", "AI判定・工学的考察・記事"]
                    info_values = [rec["no"], f"{st.session_state.struct_type_cache} ({rec['kind']})", rec["part"], rec["dim"], st.session_state.complex_degrade_cache, f"{rec['comment']}\n\n【全体複合劣化統括報告】\n{st.session_state.full_result_text if idx == 0 else 'No.1写真の全体統括レポートを参照'}"]

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
                        
                        if "考察" in label or "記事" in label:
                            text_len = len(str(value))
                            ws.row_dimensions[r].height = max(160, min(450, int(text_len * 0.40)))
                        else:
                            ws.row_dimensions[r].height = 26

                    img_buffer = io.BytesIO()
                    img.save(img_buffer, format="PNG")
                    img_buffer.seek(0)
                    xl_img = ExcelImage(img_buffer)
                    xl_img.width, xl_img.height = 480, 330
                    ws.add_image(xl_img, f"D{start_row + 3}")
                    start_row += 13

                output = io.BytesIO()
                wb.save(output)
                
                st.markdown("---")
                st.download_button(
                    label="📥 官庁・役所提出用 高精密Excel調書をダウンロード",
                    data=output.getvalue(),
                    file_name=f"【確定複合劣化調書】{p_name_xl}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as excel_err:
                st.error(f"Excel写真台帳の生成中にエラーが発生しました: {excel_err}")
        else:
            st.info("💡 『ステップ②：現場写真・変状チェック入力』画面でデータを入力し、最下部の解析実行ボタンを押してください。診断が完了するとレポートとExcel出力モジュールがここへ自動展開されます。")
