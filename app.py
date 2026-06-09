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

# --- 1. ページ基本設定と超高度化UIアーキテクチャ ---
st.set_page_config(page_title="コンクリート劣化診断 AI Suite Pro", layout="wide", initial_sidebar_state="collapsed")

# セッション状態の安全な初期化
if 'current_step' not in st.session_state: st.session_state.current_step = "step1"
if 'full_result_text' not in st.session_state: st.session_state.full_result_text = None
if 'final_width' not in st.session_state: st.session_state.final_width = 0.0
if 'analysis_completed' not in st.session_state: st.session_state.analysis_completed = False

# ナビゲーション用のコールバック関数（エラーを出さずに瞬時に画面を切り替える）
def set_step(step):
    st.session_state.current_step = step

# --- プロ仕様の洗練されたCSS定義 ---
st.markdown("""
<style>
/* 全体テーマ設定（ダークモード強制固定） */
.main { background-color: #0F172A; color: #F8FAFC; font-family: 'Helvetica Neue', Arial, sans-serif; }
.stApp { background-color: #0F172A; }
[data-testid="stSidebar"], [data-testid="collapsedControl"] { display: none !important; }

/* 共通フォントと入力UIの視認性担保 */
h1, h2, h3, h4, h5, p, span, label, .stMarkdown p { color: #F8FAFC !important; font-weight: 600 !important; }
input, textarea, select, div[data-baseweb="select"] * { color: #0F172A !important; font-weight: bold !important; }

/* チェックボックスのプロフェッショナルな配色固定（赤み根絶） */
div[data-testid="stCheckbox"] div[role="checkbox"] { border-color: #475569 !important; }
div[data-testid="stCheckbox"] div[role="checkbox"][aria-checked="true"] {
    background-color: #0284C7 !important; border-color: #38BDF8 !important; box-shadow: 0 0 10px rgba(56, 189, 248, 0.4) !important;
}
div[data-testid="stCheckbox"] div[role="checkbox"] svg { stroke: #FFFFFF !important; fill: none !important; }

/* ========================================================================= */
/* 【最重要修正】ナビゲーション用巨大タイルの構築（絶対に白くならない完全防壁） */
/* ========================================================================= */
/* 2番目のカラムブロック（ナビゲーション）内のボタンを強制的に巨大タイル化 */
div[data-testid="stHorizontalBlock"]:nth-of-type(2) div[data-testid="column"] div.stButton > button {
    background: linear-gradient(135deg, #1E293B, #0F172A) !important; /* 強制的にダークネイビー */
    border: 2px solid #334155 !important;
    border-radius: 24px !important; /* 美しい角丸 */
    min-height: 160px !important; /* 縦幅を160pxに巨大化 */
    width: 100% !important;
    display: flex !important;
    align-items: center !important;
    justify-content: center !important;
    box-shadow: 0 6px 15px rgba(0,0,0,0.4) !important;
    transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1) !important;
}
/* ボタン内の文字を強制的に特大化 */
div[data-testid="stHorizontalBlock"]:nth-of-type(2) div[data-testid="column"] div.stButton > button p {
    font-size: 26px !important; /* 文字を26pxへ超巨大化 */
    font-weight: 900 !important;
    color: #94A3B8 !important; /* 未選択時は明るいグレー */
    line-height: 1.5 !important;
    white-space: pre-wrap !important; /* 改行を完全に許可 */
    margin: 0 !important;
}
/* ホバー時の美しい浮き上がり */
div[data-testid="stHorizontalBlock"]:nth-of-type(2) div[data-testid="column"] div.stButton > button:hover {
    transform: translateY(-6px) !important;
    border-color: #38BDF8 !important;
    box-shadow: 0 15px 30px rgba(0,0,0,0.6) !important;
}
div[data-testid="stHorizontalBlock"]:nth-of-type(2) div[data-testid="column"] div.stButton > button:hover p {
    color: #FFFFFF !important; /* マウスを乗せたら白く光る */
}

/* 診断実行ボタン専用デザイン（メインボタン） */
button[data-testid="baseButton-primary"] {
    background: #0284C7 !important;
    border: 2px solid #38BDF8 !important;
    border-radius: 12px !important;
    min-height: 70px !important;
    width: 100% !important;
    transition: all 0.3s ease !important;
}
button[data-testid="baseButton-primary"] p {
    font-size: 22px !important;
    font-weight: bold !important;
    color: #FFFFFF !important;
}
button[data-testid="baseButton-primary"]:hover {
    background: #38BDF8 !important;
    box-shadow: 0 0 25px #38BDF8 !important;
}

/* 構造化コンテナデザイン */
.dashboard-card { padding: 25px; background-color: #1E293B; border-radius: 12px; border: 1px solid #334155; margin-bottom: 20px; }
.status-card { padding: 25px; background-color: #1E293B; border-radius: 12px; border-top: 1px solid #334155; border-right: 1px solid #334155; border-bottom: 1px solid #334155; margin-bottom: 20px; }
.photo-input-box { background-color: #1E293B; padding: 25px; border-radius: 12px; border: 1px dashed #38BDF8; margin-bottom: 25px; }
.report-text-box { background-color: #1E293B !important; border: 1px solid #475569; border-radius: 12px; padding: 30px; font-size: 16px; line-height: 1.8; white-space: pre-wrap; }
</style>
""", unsafe_allow_html=True)

# 現在のステップに応じた「選択中タイルの青色発光」CSSの動的注入
active_index = {"step1": 1, "step2": 2, "step3": 3}[st.session_state.current_step]
st.markdown(f"""
<style>
div[data-testid="stHorizontalBlock"]:nth-of-type(2) div[data-testid="column"]:nth-child({active_index}) div.stButton > button {{
    background: linear-gradient(135deg, #0284C7, #1E293B) !important;
    border: 3px solid #38BDF8 !important;
    box-shadow: 0 0 35px rgba(56, 189, 248, 0.5) !important;
}}
div[data-testid="stHorizontalBlock"]:nth-of-type(2) div[data-testid="column"]:nth-child({active_index}) div.stButton > button p {{
    color: #FFFFFF !important;
    text-shadow: 0 0 10px rgba(255, 255, 255, 0.5) !important;
}}
</style>
""", unsafe_allow_html=True)

# 解析完了時に③番ボタンをパルス点滅させるアニメーション
if st.session_state.analysis_completed and active_index != 3:
    st.markdown("""
    <style>
    @keyframes pulse_glow {
        0% { box-shadow: 0 0 10px rgba(56,189,248,0.2); background: #1E293B !important; border-color: #334155 !important; }
        100% { box-shadow: 0 0 40px #38BDF8, inset 0 0 20px #38BDF8; background: #0284C7 !important; border-color: #38BDF8 !important; }
    }
    div[data-testid="stHorizontalBlock"]:nth-of-type(2) div[data-testid="column"]:nth-child(3) div.stButton > button {
        animation: pulse_glow 1.2s infinite alternate !important;
    }
    div[data-testid="stHorizontalBlock"]:nth-of-type(2) div[data-testid="column"]:nth-child(3) div.stButton > button p {
        color: #FFFFFF !important;
    }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 閉域セキュリティ認証 ---
def check_password():
    if "authenticated" not in st.session_state: st.session_state["authenticated"] = False
    def password_entered():
        if st.session_state["password"] == "tn0000": st.session_state["authenticated"] = True
        else: st.error("❌ パスワードが違います")
    if not st.session_state["authenticated"]:
        if os.path.exists("logo.png"): st.image("logo.png", width=200)
        st.markdown("<h2 style='text-align: center;'>🔒 コンクリート劣化診断システム</h2>", unsafe_allow_html=True)
        st.text_input("アクセスパスワードを入力", type="password", on_change=password_entered, key="password")
        return False
    return True

if check_password():
    api_key = st.secrets.get("GEMINI_API_KEY", "")

    # 業務ヘッダー（1番目のカラムブロック）
    col_logo, col_title = st.columns([1, 6])
    with col_logo:
        if os.path.exists("logo.png"): st.image("logo.png", width=140)
    with col_title:
        st.markdown("<h1 style='margin-bottom: 0;'>コンクリート構造物 高精密総合診断システム</h1>", unsafe_allow_html=True)
        st.markdown("<p style='color: #38BDF8; font-size: 16px; margin-top: 5px;'>農林水産省機能保全手引き・JCI複合劣化マトリクス完全準拠</p>", unsafe_allow_html=True)

    # 巨大ナビゲーションメニュー（2番目のカラムブロック）
    col_nav1, col_nav2, col_nav3 = st.columns(3)
    with col_nav1: st.button("🏠\n① 設置地域・環境判定", on_click=set_step, args=("step1",), use_container_width=True)
    with col_nav2: st.button("📸\n② 写真・変状チェック入力", on_click=set_step, args=("step2",), use_container_width=True)
    with col_nav3: st.button("📑\n③ 統合診断レポート・Excel", on_click=set_step, args=("step3",), use_container_width=True)

    st.markdown("---")

    # ==========================================
    # STEP 1: 設置地域・環境判定
    # ==========================================
    if st.session_state.current_step == "step1":
        st.markdown("### 📍 1. 構造物所在地およびJCI環境マッピング解析")
        address_input = st.text_input("所在地を入力（例：山形県酒田市、北陸沿岸、国道・高速路線等）", placeholder="例：山形県酒田市大浜 国道112号", key="addr_in")
        
        auto_freeze, auto_salt, auto_asr, auto_complex = "未判定", "未判定", "未判定", "所在地に基づくJCI複合劣化マトリクスとの照合待機中"
        
        if address_input:
            cold_regions = ["北海道", "青森", "岩手", "秋田", "山形", "宮城", "福島", "新潟", "富山", "石川", "福井", "長野", "岐阜", "群馬", "山梨"]
            salt_keywords = ["浜", "海岸", "港", "湾", "岬", "磯", "シーサイド", "大浜", "臨海", "塩", "浦", "津"]
            asr_regions = ["山形", "秋田", "新潟", "富山", "石川", "福井", "長野", "岐阜", "京都", "兵庫", "香川", "徳島", "福岡", "佐賀", "熊本"]
            
            is_cold = any(r in address_input for r in cold_regions)
            is_coast = any(k in address_input for k in salt_keywords)
            has_asr = any(a in address_input for a in asr_regions)
            
            auto_freeze = "【高危険度】凍結融解サイクル年平均45回以上。スケーリング・組織破壊リスク大。" if is_cold else "【一般】凍結融解の深刻な作用確率は低。"
            auto_salt = "【重塩害警戒】強風による高濃度飛来塩分の定着危険領域。" if is_coast else ("【塩害警戒】凍結防止剤散布による塩化物供給環境。" if is_cold and any(road in address_input for road in ["国道","高速","インター"]) else "【一般】塩化物の直接的影響は軽微。")
            auto_asr = "【ASR要注意】反応性骨材の過去の流通・損傷報告地域。" if has_asr else "【低リスク】異常膨張リスクは穏健。"

            if is_cold and is_coast and has_asr: auto_complex = "⚠️ 【最過酷マトリクス：塩害×凍害×ASR】日本海側沿岸。ASRクラックを起点に塩分・水分が浸透、マクロセル腐食と組織破壊が相乗進展する極めて過酷な環境。"
            elif is_cold and is_coast: auto_complex = "⚡ 【複合劣化警戒：塩害×凍害】表層脆弱化（スケーリング）と鉄筋腐食（爆裂現象）が同時複合進展する領域。"
            elif is_cold and has_asr: auto_complex = "🔎 【複合劣化警戒：凍害×ASR】ゲルの吸水膨張クラックへ冬季水分が浸入し、凍結膨張圧で開口が拡張する複合領域。"
            elif is_coast and has_asr: auto_complex = "⚓ 【複合劣化警戒：ASR×塩害】ASRクラックから塩化物イオンが内部へ急速拡散し、早期発錆を誘発する環境。"
            else: auto_complex = "✅ 【単一要因環境】現在の立地条件において、致命的な複合侵食リスクは比較的低いと推定されます。"

        if address_input:
            st.markdown(f"<div class='status-card' style='border-left: 8px solid #38BDF8;'><h4>🧬 JCI（日本コンクリート工学会）複合劣化判定</h4><p style='font-size:18px; color:#38BDF8;'>{auto_complex}</p></div>", unsafe_allow_html=True)
            c1, c2, c3 = st.columns(3)
            with c1: st.markdown(f"<div class='dashboard-card'><h4>❄️ 凍害危険度</h4><p>{auto_freeze}</p></div>", unsafe_allow_html=True)
            with c2: st.markdown(f"<div class='dashboard-card'><h4>🌊 塩害・融雪剤</h4><p>{auto_salt}</p></div>", unsafe_allow_html=True)
            with c3: st.markdown(f"<div class='dashboard-card'><h4>💎 ASR骨材分布</h4><p>{auto_asr}</p></div>", unsafe_allow_html=True)

        st.markdown("### 🛠️ 2. 構造物基本条件の設定")
        cc1, cc2, cc3 = st.columns(3)
        with cc1:
            struct_type = st.selectbox("① 構造物の種類", ["建築物（校舎・庁舎：外壁・柱・梁）", "橋梁（床版・主桁・橋台・橋脚）", "ボックスカルバート・共同溝", "擁壁", "開渠・水路", "ダム・沈殿池"])
            cement_type = st.selectbox("④ セメント種類", ["普通ポルトランド", "高炉セメント（B種等）", "早強ポルトランド", "不明"])
        with cc2:
            env_location = st.multiselect("② 物理的設置条件", ["屋外・直接雨掛かり", "日陰・日裏・軒下", "常時流水・水撃エリア", "屋内・常時乾燥"], default=[])
            elapsed_years = st.selectbox("⑤ 供用年数", ["10年未満", "10年〜30年未満", "30年〜50年未満", "50年以上"])
            manual_years_input = st.text_input("（任意）具体的な築年数・竣工年")
        with cc3:
            wet_status = st.multiselect("③ 内部湿潤状況", ["漏水なし", "微細な湿潤", "活動性の漏水あり", "高水圧環境"], default=[])
            crack_type = st.selectbox("⑥ 支配的損傷症状", ["ひび割れ（単一）", "浮き・剥離・剥落", "鉄筋露出・爆裂", "エフロ析出伴う漏水", "ASR（３方向クラック）", "スケーリング（凍害・摩耗）"])
        
        region_info = st.text_area("⑦ 現場特記事項", placeholder="例: 交通振動あり、不同沈下の形跡あり等")
        st.success("✅ 設定が完了しました。画面上部の巨大な『📸 ② 写真・変状チェック入力』ボタンを押してください。")

    # ==========================================
    # STEP 2: 現場写真・変状チェック入力
    # ==========================================
    elif st.session_state.current_step == "step2":
        st.markdown("### 🏢 1. 業務基本情報の入力")
        col_a, col_b, col_c = st.columns(3)
        with col_a: project_name = st.text_input("物件名（工事・業務名）", placeholder="例：塩竈清掃工場 躯体調査")
        with col_b: location_name = st.text_input("測定箇所詳細", placeholder="例：X5通り芯 B1梁上面")
        with col_c: inspector_name = st.text_input("調査担当者名", value="T&N技術管理者")
            
        st.markdown("---")
        st.markdown("### 🔧 2. 技術的要因の補足チェック")
        ch1, ch2, ch3 = st.columns(3)
        with ch1:
            cb_shear = st.checkbox("構造要因・不同沈下・土圧・耐震上のせん断ひび割れ疑い")
            cb_janka = st.checkbox("施工起因のジャンカ・コールドジョイント・初期欠陥")
        with ch2:
            cb_wet_h = st.checkbox("常時湿潤・漏水（ASR・遊離石灰溶出リスク大）")
            cb_cover = st.checkbox("設計かぶり厚不足または中性化の鉄筋位置到達の疑い")
        with ch3:
            cb_joint = st.checkbox("施工目地・伸縮目地・伸縮装置周辺部からの漏水・損傷")
            
        selected_factors = []
        if cb_shear: selected_factors.append("不同沈下・土圧等のせん断ひび割れ")
        if cb_janka: selected_factors.append("施工起因の初期欠陥")
        if cb_wet_h: selected_factors.append("常時湿潤・高水圧・漏水")
        if cb_cover: selected_factors.append("かぶり厚不足・中性化進行")
        if cb_joint: selected_factors.append("目地部からの漏水進展")
        human_factors_text = "、".join(selected_factors) if selected_factors else "特になし"

        st.markdown("---")
        st.markdown("### 📸 3. 現場写真アップロード ＆ 損傷詳細プロット")
        uploaded_files = st.file_uploader("写真をアップロード（最大6枚）", type=["jpg", "jpeg", "png"], accept_multiple_files=True)
        
        images, photo_details_prompt, photo_excel_records = [], [], []
        
        if uploaded_files:
            for idx, file in enumerate(uploaded_files[:6]):
                img = Image.open(file)
                images.append(img)
                st.markdown(f"<div class='photo-input-box'><h4>📸 写真 No.{idx+1} 損傷プロット（通り芯・打診数量連動）</h4></div>", unsafe_allow_html=True)
                cv, fv = st.columns([1, 2])
                with cv: st.image(img, use_container_width=True)
                with fv:
                    f1, f2 = st.columns(2)
                    with f1:
                        p_part = st.text_input(f"No.{idx+1} 損傷位置・通り芯・面", key=f"p_part_{idx}")
                        p_kind = st.selectbox(f"No.{idx+1} 損傷分類", ["構造クラック", "乾燥収縮", "ASR（３方向マップ状）", "凍害（層状剥離）", "塩害・爆裂", "複合劣化疑い"], key=f"p_kind_{idx}")
                        p_efflo = st.checkbox(f"進行性漏水・エフロあり", key=f"p_efflo_{idx}")
                        p_rust = st.checkbox(f"剥離・鉄筋露出・錆汁あり", key=f"p_rust_{idx}")
                    with f2:
                        p_w = st.number_input(f"No.{idx+1} 幅(mm)", min_value=0.0, step=0.05, key=f"p_w_{idx}")
                        p_l = st.number_input(f"No.{idx+1} 長さ(cm)", min_value=0.0, step=1.0, key=f"p_l_{idx}")
                        p_a = st.number_input(f"No.{idx+1} 浮き・剥離面積(㎡)", min_value=0.00, step=0.05, key=f"p_a_{idx}")
                        p_c = st.text_area(f"No.{idx+1} 技術者所見", key=f"p_c_{idx}")
                
                photo_details_prompt.append(f"[No.{idx+1}] 位置:{p_part}|分類:{p_kind}|幅:{p_w}mm|長:{p_l}cm|浮き:{p_a}㎡|漏水:{p_efflo}|錆:{p_rust}|所見:{p_c}")
                photo_excel_records.append({"no":f"No.{idx+1}","part":p_part,"kind":p_kind,"dim":f"W:{p_w}mm/L:{p_l}cm/A:{p_a}㎡","comment":f"エフロ:{p_efflo}/錆:{p_rust}|{p_c}"})
                st.markdown("---")
            
            # 安全・確実なAI解析の実行
            if st.button("🚀 環境マトリクスと全写真データを統合して高精密AI診断を実行", key="execute_analysis_btn", type="primary"):
                if not api_key: st.error("APIキーが設定されていません。")
                else:
                    with st.spinner("🔍 熟練コンクリート診断士AIが最新指針と同期しながら精密解析中..."):
                        try:
                            genai.configure(api_key=api_key)
                            model = genai.GenerativeModel('gemini-2.5-flash')
                            
                            # セッションキャッシュへ確実な保存
                            st.session_state.excel_records_cache, st.session_state.images_cache = photo_excel_records, images
                            st.session_state.header_data = {
                                "p": project_name, "l": location_name, "i": inspector_name,
                                "a": st.session_state.get("addr_in", ""), "s": struct_type,
                                "c": st.session_state.get("complex_degrade_cache", "一般環境")
                            }
                            
                            prompt = f"""あなたが作成すべきは、農水省・国交省等へ提出する最高レベルの工学的調査報告書です。
【JCI環境マトリクス】: {st.session_state.header_data['a']} における {st.session_state.header_data['c']} シナリオを考慮。
【構造物・写真データ】:
{"/ ".join(photo_details_prompt)}
上記損傷データと、学んだ「複合劣化の相乗進展（凍害×ASR等）」の因果関係をリンクさせ、各写真に詳細な技術的所見を述べてください。
確定最大幅、総合劣化度（Ⅰ〜Ⅳ）を冒頭に示し、具体的な補修工法（注入、断面修復等）を選定理由とともに厚く論じてください。"""
                            
                            response = model.generate_content([prompt] + images)
                            st.session_state.full_result_text = response.text
                            
                            # 正規表現による最大幅抽出の安定化
                            final_w = 0.0
                            match = re.search(r"最大ひび割れ幅:\s*([0-9.]+)", response.text)
                            if match: final_w = float(match.group(1))
                            st.session_state.final_width = final_w
                            st.session_state.analysis_completed = True
                            
                            # 診断完了後、確実にStep3へ遷移
                            st.session_state.current_step = "step3"
                            st.rerun()
                            
                        except Exception as e:
                            st.error(f"エラーが発生しました。アクセス集中（429）の場合は時間をおいて再度お試しください。詳細: {e}")

    # ==========================================
    # STEP 3: 統合診断レポート・Excel調書
    # ==========================================
    elif st.session_state.current_step == "step3":
        if st.session_state.full_result_text:
            fw = st.session_state.final_width
            if "劣化度Ⅲ" in st.session_state.full_result_text or "劣化度Ⅳ" in st.session_state.full_result_text or fw >= 1.0:
                color_code, status_title, alert_desc = "#EF4444", f"🔴 【劣化度Ⅲ〜Ⅳ：早期補修対応】判定最大幅: {fw} mm", "⚠️ 2因子以上の侵食が重畳し劣化期へ突入。早期の断面修復および詳細調査が必須です。"
            elif "劣化度Ⅱ" in st.session_state.full_result_text or (0.2 <= fw < 1.0):
                color_code, status_title, alert_desc = "#EAB308", f"🟡 【劣化度Ⅱ：中期劣化・機能保持】判定最大幅: {fw} mm", "💡 有害な損傷・複合劣化の初期兆候。「低圧エポキシ樹脂注入工法」等の選定を推奨します。"
            elif fw > 0:
                color_code, status_title, alert_desc = "#10B981", f"🟢 【劣化度Ⅰ：経過観察フェーズ】判定最大幅: {fw} mm", "✅ 影響は軽微です。表面含浸工法等の予防保全、または目視経過観察となります。"
            else:
                color_code, status_title, alert_desc = "#3B82F6", "🔵 【複合劣化度判定保留：現地実測要請】", "ℹ️ 正確な縮尺基準が確認できないため判定を保留しています。現地実測値を確認してください。"
            
            st.markdown(f"<div class='status-card' style='border-left: 8px solid {color_code};'><h3>{status_title}</h3><p style='font-size:15px; margin-top:8px;'>{alert_desc}</p></div>", unsafe_allow_html=True)
            st.markdown("### 📑 AI Suite Pro 統合解析レポート")
            st.markdown(f"<div class='report-text-box'>{st.session_state.full_result_text}</div>", unsafe_allow_html=True)

            # --- Excel調書出力モジュール ---
            try:
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "損傷調書"
                ws.page_setup.orientation, ws.page_setup.paperSize = ws.ORIENTATION_LANDSCAPE, ws.PAPERSIZE_A4
                ws.column_dimensions['A'].width, ws.column_dimensions['B'].width, ws.column_dimensions['D'].width = 24, 54, 75
                
                thin_border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
                
                start_row = 1
                for idx, img in enumerate(st.session_state.images_cache):
                    rec = st.session_state.excel_records_cache[idx]
                    ws.merge_cells(f"A{start_row}:D{start_row}")
                    ws[f"A{start_row}"] = f"■ {st.session_state.header_data['p']} 構造物健全度調査台帳"
                    ws[f"A{start_row}"].font = Font(name="MS ゴシック", size=14, bold=True)
                    
                    ws.merge_cells(f"A{start_row+1}:D{start_row+1}")
                    ws[f"A{start_row+1}"] = f"位置・通り芯： {st.session_state.header_data['l']} | 担当：{st.session_state.header_data['i']}"
                    ws[f"A{start_row+1}"].font, ws[f"A{start_row+1}"].alignment = Font(name="MS ゴシック", size=11, bold=True), Alignment(horizontal="right")
                    
                    labels = ["写真No. / 撮影項目", "工種・変状分類", "位置・通り芯・面", "実測寸法・打診数量", "JCI複合劣化マトリクス", "AI判定・工学的考察"]
                    vals = [rec["no"], f"{st.session_state.header_data['s']} ({rec['kind']})", rec["part"], rec["dim"], st.session_state.header_data['c'], f"{rec['comment']}\n\n【統括報告】\n{st.session_state.full_result_text if idx==0 else 'No.1写真レポートを参照'}"]

                    for i, (l, v) in enumerate(zip(labels, vals)):
                        r = start_row + 2 + i
                        ws[f"A{r}"], ws[f"B{r}"] = l, v
                        ws[f"A{r}"].font, ws[f"A{r}"].fill = Font(name="MS ゴシック", size=11, bold=True), PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
                        ws[f"B{r}"].font = Font(name="MS ゴシック", size=11)
                        ws[f"A{r}"].border = thin_border
                        ws[f"B{r}"].border = thin_border
                        ws[f"A{r}"].alignment = Alignment(horizontal="center", vertical="center")
                        ws[f"B{r}"].alignment = Alignment(wrap_text=True, vertical="top")
                        if i == 5: ws.row_dimensions[r].height = 250
                    
                    img_buf = io.BytesIO()
                    img.save(img_buf, format="PNG")
                    xl_img = ExcelImage(img_buf)
                    xl_img.width, xl_img.height = 520, 360
                    ws.add_image(xl_img, f"D{start_row + 2}")
                    start_row += 13

                output = io.BytesIO()
                wb.save(output)
                st.download_button(label="📥 官庁提出用 高精密Excel調書をダウンロード", data=output.getvalue(), file_name=f"確定診断調書_{datetime.now().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception as excel_err:
                st.error(f"Excel写真台帳の生成中にエラーが発生しました: {excel_err}")
        else:
            st.info("💡 診断が実行されていません。ステップ②で現場写真をアップロードし、解析を実行してください。")
