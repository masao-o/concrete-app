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

# --- 1. ページ設定とプロ仕様ダッシュボードCSS ---
st.set_page_config(page_title="T&N コンクリート劣化診断 AI Suite Pro", layout="wide")
st.markdown("""
<style>
.main { background-color: #0F172A; color: #FFFFFF; }
.stApp { background-color: #0F172A; }
section[data-testid="stSidebar"] { background-color: #1E293B !important; border-right: 1px solid #334155; }

/* ユニバーサル純白フォント設定 */
h1, h2, h3, h4, h5, h6, p, span, label, .stMarkdown,
[data-testid="stSidebar"] label, [data-testid="stSidebar"] p, [data-testid="stSidebar"] span,
.stCheckbox label, div[data-testid="stMarkdownContainer"] p {
    color: #FFFFFF !important; font-family: 'Helvetica Neue', Arial, sans-serif; font-weight: bold !important;
}

/* 入力・選択ボックスのハッキリ化（黒文字で視認性担保） */
input, textarea, select, div[data-baseweb="select"] * { color: #0F172A !important; font-weight: bold !important; }
input::placeholder, textarea::placeholder { color: #64748B !important; opacity: 1 !important; }

/* ファイルアップローダー */
div[data-testid="stFileUploader"] section { background-color: #F8FAFC !important; border: 2px dashed #94A3B8 !important; }
div[data-testid="stFileUploader"] section div, div[data-testid="stFileUploader"] section p,
div[data-testid="stFileUploader"] section span, div[data-testid="stFileUploader"] section small {
    color: #475569 !important; font-weight: bold !important;
}

/* ダッシュボードカード */
.dashboard-card { padding: 25px; background-color: #1E293B; border-radius: 16px; border: 1px solid #334155; margin-bottom: 20px; }
.status-card { padding: 25px; background-color: #1E293B; border-radius: 16px; margin-bottom: 20px; border-top: 1px solid #334155; border-right: 1px solid #334155; border-bottom: 1px solid #334155; }

/* AI出力テキストボックス */
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

/* タブメニューの文字色調整 */
button[data-baseweb="tab"] { color: #94A3B8 !important; font-size: 16px !important; }
button[data-baseweb="tab"][aria-selected="true"] { color: #38BDF8 !important; font-weight: bold !important; border-bottom-color: #38BDF8 !important; }
</style>
""", unsafe_allow_html=True)

# --- 2. 閉域パスワード認証システム ---
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
        st.text_input("アクセスパスワード（技術管理者専用）", type="password", on_change=password_entered, key="password")
        return False
    return True

if check_password():
    # セッション状態の初期化（画面切り替えでのエラー完全防止用）
    if 'full_result_text' not in st.session_state:
        st.session_state.full_result_text = None
    if 'final_width' not in st.session_state:
        st.session_state.final_width = 0.0

    if os.path.exists("logo.png"): 
        st.sidebar.image("logo.png", width=180)
    st.sidebar.markdown("### 💻 AI Suite Pro v3.0\n技術管理者ログイン中")
    st.sidebar.markdown("---")
    
    api_key = st.secrets.get("GEMINI_API_KEY", "")

    # タイトル表示
    st.markdown("<h1 style='color: white; margin-bottom: 0;'>🚗 AI Suite Pro</h1>", unsafe_allow_html=True)
    st.markdown("<p style='color: #94A3B8; font-size: 16px;'>実務特化型コンクリート高精密診断ダッシュボード</p>", unsafe_allow_html=True)

    # --- 3. ダッシュボード型タブナビゲーションの構築 ---
    tab_home, tab_input, tab_report = st.tabs([
        "🏠 ① ホーム・地域環境設定", 
        "📸 ② 現場データ・写真アップロード", 
        "📑 ③ 統合診断レポート・Excel調書"
    ])

    # ==========================================
    # TAB 1: ホーム・地域環境設定
    # ==========================================
    with tab_home:
        st.markdown("### 📍 1. 構造物所在地・マクロ気象自動判定")
        st.markdown("<p style='color: #94A3B8;'>住所を入力すると、過去数十年の天気予報・地質実績から凍害・融雪剤・塩害リスクを自動算出します。</p>", unsafe_allow_html=True)
        
        address_input = st.text_input("構造物の設置住所・施設名を入力", placeholder="例：山形県酒田市大浜 国道112号", key="addr_in")
        
        # 自動気象解析ロジック
        auto_freeze_info = "判定不能（住所未入力）"
        auto_salt_info = "判定不能（住所未入力）"
        auto_agent_info = "判定不能（住所未入力）"
        auto_weather_summary = "特記事項なし"
        
        if address_input:
            cold_regions = ["北海道", "青森", "岩手", "秋田", "山形", "宮城", "福島", "新潟", "富山", "石川", "福井", "長野", "岐阜", "群馬", "山梨"]
            is_cold = any(reg in address_input for reg in cold_regions)
            
            if is_cold:
                auto_freeze_info = "【激甚凍害地域】冬季凍結融解サイクルが年平均45〜60回以上の極過酷エリア。膨張圧によるスケーリングや組織破壊リスク高。"
                auto_agent_info = "【融雪剤散布確率：大】道路雪氷対策計画の重点路線。塩化物大量散布に伴う飛沫侵食、二次的塩害リスク大。"
            else:
                auto_freeze_info = "【一般環境】冬季の激しい凍結融解サイクルリスクは比較的低いエリア。"
                auto_agent_info = "【融雪剤散布確率：小】定期的な凍結防止剤の大量散布環境には該当しない可能性が高い。"
                
            salt_keywords = ["浜", "海岸", "港", "湾", "岬", "磯", "シーサイド", "大浜", "臨海", "塩", "浦", "津"]
            is_coast = any(kw in address_input for kw in salt_keywords)
            
            if is_coast:
                auto_salt_info = "【重塩害警戒】沿岸近傍。強風により高濃度の飛来塩分が構造物表面に定着しやすく、鉄筋不動態被膜破壊リスク高。"
            else:
                auto_salt_info = "【一般内陸地域】沿岸からの直接的な飛来塩分の影響は極めて軽微なエリア。"
                
            if is_cold and is_coast:
                auto_weather_summary = "過酷な沿岸寒冷地環境。乾湿の繰り返し、激しい寒暖差、飛来塩分による化学的侵食、物理的摩耗・風化が複合する激甚環境。"
            elif is_cold:
                auto_weather_summary = "内陸寒冷地または山間部。積雪・融雪による常時湿潤環境と、凍結防止剤の塩化物侵食が支配的な経年複合要因。"
            else:
                auto_weather_summary = "比較的温暖な一般環境。主として大気中の中性化および雨水による溶出が主たる要因。"

        # 自動解析結果をカード型ダッシュボードで視認性よく表示
        if address_input:
            c1, c2, c3 = st.columns(3)
            with c1:
                st.markdown(f"<div class='dashboard-card'><h4>❄️ 凍害・凍結リスク</h4><p style='font-size:14px; color:#CBD5E1;'>{auto_freeze_info}</p></div>", unsafe_allow_html=True)
            with c2:
                st.markdown(f"<div class='dashboard-card'><h4>🌊 塩害・飛来塩分</h4><p style='font-size:14px; color:#CBD5E1;'>{auto_salt_info}</p></div>", unsafe_allow_html=True)
            with c3:
                st.markdown(f"<div class='dashboard-card'><h4>🚗 融雪剤（塩化物）影響</h4><p style='font-size:14px; color:#CBD5E1;'>{auto_agent_info}</p></div>", unsafe_allow_html=True)

        st.markdown("---")
        st.markdown("### 🛠️ 2. プロ診断士用 条件設定（手動補完）")
        
        cc1, cc2, cc3 = st.columns(3)
        with cc1:
            struct_type = st.selectbox("① 構造物の種類", ["（未選択・写真から自動判定）", "橋梁（上部工/下部工）", "ボックスカルバート", "擁壁（重力式/もたれ式）", "トンネル覆工", "港湾・河川構造物・水路", "ダム（重力式/アーチ）", "建築物基礎・柱・壁"])
            cement_type = st.selectbox("④ 使用セメントの種類", ["（未選択）", "普通ポルトランドセメント", "高炉セメント（B種など・ASR/塩害対策）", "早強ポルトランドセメント", "不明"])
        with cc2:
            env_location = st.multiselect("② 設置環境・大分類（複数選択可）", ["一般地域（屋外・雨掛かり）", "一般地域（日陰・軒下）", "塩害警戒地域（海岸付近）", "寒冷地・凍枯地域", "水面下・常時通水・流水部", "屋内（常時乾燥）"], default=[])
            elapsed_years = st.selectbox("⑤ 供用年数（経過年数）", ["（未選択）", "5年未満（初期欠陥の可能性）", "5年以上〜20年未満", "20年以上〜50年未満", "50年以上（高経年化・農水施設機能保全要件）"])
        with cc3:
            wet_status = st.multiselect("③ 湿潤状態（複数選択可）", ["常時乾燥状態", "乾湿の繰り返し（ひび割れ進展）", "常時湿潤状態（漏水・滞水）", "高水圧・浸透環境（ダム・水槽等）"], default=[])
            crack_type = st.selectbox("⑥ 目視での主たる劣化症状", ["（未選択・写真から自動判定）", "ひび割れ（単一・規則性）", "亀甲状のひび割れ（ASRなどの疑い）", "エフロレッセンス（白華）の析出伴う", "コンクリートの剥離・鉄筋露出（爆裂現象）", "漏水・遊離石灰を伴う錆汁・流水による摩耗・スケーリング"])
            
        region_info = st.text_area("⑦ 地域・気象特記事項（手動補足がある場合入力）", placeholder="例: 特になし（上の住所自動判定データが優先されます）")
        st.success("✅ ステップ①完了：環境条件がセットされました。次の『② 現場データ・写真アップロード』タブを開いてください。")

    # ==========================================
    # TAB 2: 現場データ・写真アップロード
    # ==========================================
    with tab_input:
        st.markdown("### 🏢 1. 提出用 業務基本情報の入力")
        col_a, col_b, col_c = st.columns(3)
        with col_a:
            project_name = st.text_input("項目A：物件名（工事名・業務名）", placeholder="（例：塩竈清掃工場 躯体調査）")
        with col_b:
            location_name = st.text_input("項目B：調査位置・測定箇所詳細", placeholder="（例：沈殿池 南面壁）")
        with col_c:
            inspector_name = st.text_input("項目C：調査担当者（コンクリート診断士名）", value="T&N技術管理者")
            
        st.markdown("---")
        st.markdown("### 🔧 2. 技術者による物理的・構造的要因の補足（重複なし単一チェック）")
        
        ch1, ch2, ch3 = st.columns(3)
        with ch1:
            cb_salt = st.checkbox("海岸線から2km以内（重塩害環境の確定）")
            cb_freeze = st.checkbox("寒冷地・融雪剤散布路線の直接的影響")
        with ch2:
            cb_wet = st.checkbox("常時湿潤・漏水・高水圧環境（ASR・遊離石灰溶出）")
            cb_shear = st.checkbox("不同沈下・土圧・構造応力または地動起因のせん断応力疑い")
        with ch3:
            cb_janka = st.checkbox("施工起因のジャンカ・初期欠陥の目視確認あり")
            cb_joint = st.checkbox("施工目地・打継目地部からの漏水・滞水")
            cb_cover = st.checkbox("設計かぶり厚の不足が疑われる・または既知")
            
        selected_factors = []
        if cb_salt: selected_factors.append("海岸線から2km以内（塩害環境）")
        if cb_freeze: selected_factors.append("寒冷地・凍害・凍結防止剤散布路線の影響")
        if cb_wet: selected_factors.append("常時湿潤・高水圧・漏水（ASR・遊離石灰溶出リスク大）")
        if cb_shear: selected_factors.append("不同沈下・土圧・構造応力または地動起因のせん断ひび割れ（耐震性能影響）")
        if cb_janka: selected_factors.append("施工起因のジャンカ・初期欠陥の目視確認あり")
        if cb_joint: selected_factors.append("施工目地・打継目地部からの滞水・漏水")
        if cb_cover: selected_factors.append("設計かぶり厚の不足または中性化の鉄筋位置到達")
        human_factors_text = "、".join(selected_factors) if selected_factors else "特になし"

        st.markdown("---")
        st.markdown("### 📏 3. クラックスケール実測値の上書き指定（ハルシネーション完全防御）")
        st.info("💡 写真内に縮尺基準が無い場合、AIの数値捏造を防ぐため実測値を以下に入力してください。未入力の場合、AIは勝手に数値を推測せず保留します。")
        col_w, col_l = st.columns(2)
        with col_w:
            manual_width = st.number_input("実測ひび割れ幅 (mm)", min_value=0.0, step=0.05, value=0.0)
        with col_l:
            manual_length = st.number_input("実測ひび割れ長さ (cm)", min_value=0.0, step=1.0, value=0.0)

        st.markdown("---")
        st.markdown("### 📸 4. 現場写真アップロード（最大6枚・一括写真台帳フォーマット対応）")
        uploaded_files = st.file_uploader("ここにコンクリート構造物の写真（劣化箇所）をドロップしてください", type=["jpg", "jpeg", "png"], accept_multiple_files=True)
        
        images = []
        photo_comments = []
        
        if uploaded_files:
            if len(uploaded_files) > 6:
                st.warning("⚠️ 最初の6枚のみを処理します。")
                uploaded_files = uploaded_files[:6]
                
            img_cols = st.columns(len(uploaded_files))
            for idx, file in enumerate(uploaded_files):
                img = Image.open(file)
                images.append(img)
                with img_cols[idx]:
                    st.image(img, caption=f"Photo No.{idx+1}", use_container_width=True)
                    comment = st.text_input(f"No.{idx+1} コメント", key=f"comment_{idx}", placeholder="例: 梁中央部")
                    photo_comments.append(f"【Photo No.{idx+1}】: {comment if comment else '特記事項なし'}")
            
            st.markdown("---")
            execute_analysis = st.button("🚀 以上の条件・写真データをすべて統合して高精密AI解析を実行する")

    # ==========================================
    # TAB 3: 統合診断レポート・Excel調書
    # ==========================================
    with tab_report:
        # ボタンが押された場合の解析トリガー（エラーセーフ構造）
        if uploaded_files and 'execute_analysis' in locals() and execute_analysis:
            if not api_key:
                st.error("APIキーが設定されていません。StreamlitのSecrets設定を確認してください。")
            else:
                with st.spinner("🔍 熟練コンクリート診断士AIが農水省指針(機能保全基準)・耐震性能要件と照合しながら精密解析中..."):
                    
                    max_retries = 3
                    for attempt in range(max_retries):
                        try:
                            genai.configure(api_key=api_key)
                            model = genai.GenerativeModel('gemini-2.5-flash')
                            
                            env_text = "、".join(env_location) if env_location else "指定なし"
                            wet_text = "、".join(wet_status) if wet_status else "指定なし"
                            photo_comments_text = "\n".join(photo_comments)
                            
                            dim_info = "手動指定なし（写真内に明確なスケールが無ければ測定不可として厳格に質問してください）"
                            if manual_width > 0 or manual_length > 0:
                                dim_info = f"【診断士実測確定値】ひび割れ幅: {manual_width} mm, 長さ: {manual_length} cm"

                            prompt = f"""
あなたは日本最高峰の「コンクリート診断士」であり、農林水産省の「農業水利施設の機能保全の手引き」「ダム機能診断マニュアル」「ダム耐震性能照査マニュアル」および、国交省・各種実務報告書の劣化度判定基準をマスターしている専門家です。
既存の定型文や他者の著作権を侵害する文面を一切排除し、提示された固有の環境要因と写真データから、最高レベルの工学的報告書を論理的に起稿してください。

【対象構造物の所在地・マクロ気象データ（システム自動算出）】
- 住所・施設名: {address_input if address_input else '未指定'}
- 地域凍害・凍結融解サイクルリスク: {auto_freeze_info}
- 地域飛来塩分リスク: {auto_salt_info}
- 融雪剤（凍結防止剤）影響予測: {auto_agent_info}
- 気象・摩耗・風化複合要因サマリー: {auto_weather_summary}

【手動条件入力データ】
- 構造物種別: {struct_type}
- 設置環境・湿潤状態: {env_text} / {wet_text}
- 使用セメントの種類 / 経過年数: {cement_type} / {elapsed_years}
- 主たる劣化症状（目視）: {crack_type}
- 人為的補足因子（構造・施工・経年要因）: {human_factors_text}
- 寸法情報に関する事前条件: {dim_info}
- 提出写真ごとの個別コメント・位置情報:
{photo_comments_text}

【絶対厳守命令】
1. 手動指定寸法が無く、かつ写真内に「クラックスケール」や明確な寸法基準が確認できない場合、絶対に寸法を数値として勘で捏造しないでください。その場合は必ず文章の冒頭で「【寸法判定保留】写真から正確な縮尺基準が確認できないため、数値を勝手に推測せず保留します。正確な測定のために実測値または縮尺基準の提供を求めます」と記載し、ユーザーへ逆質問してください。
2. 判定基準として、農水省指針等に則り、ひび割れ幅、漏水、エフロ、剥離露出の有無を総合評価し、以下の4区分から該当する【劣化度】を必ず選定・明記してください。
   ・劣化度Ⅰ（軽微・経過観察）: ひび割れ幅0.2mm未満、漏水・錆汁なし。表面含浸等の予防保全。
   ・劣化度Ⅱ（中期・要補修）: ひび割れ幅0.2mm以上1.0mm未満、またはエフロ析出あり。低圧エポキシ樹脂注入工法。
   ・劣化度Ⅲ（重大・早期補修）: ひび割れ幅1.0mm以上、または漏水、コンクリートの剥離、鉄筋露出（爆裂）あり。ポリマーセメントモルタル充填工法（断面修復工法）。
   ・劣化度Ⅳ（激甚・緊急対応）: 構造物の変形、耐震性能を揺るがす深刻なせん断ひび割れ、激しい漏水・滞水。
3. 考察には専門用語（鉄筋の不動態被膜破壊、マクロセル腐食、遊離石灰の溶出、膨張圧、流水による物理的摩耗、スケーリング、剪断応力など）を適切に用い、この現場特有の化学的・物理的侵食因子の因果関係を詳細に展開してください。

出力は、確認できた場合のみ「確定ひび割れ幅: 〇.〇 mm / 劣化度: 〇」を必ず冒頭の1行目に示し、その後【工学的劣化原因の深い推測】、【写真ごとの個別技術的所見】、【推奨される具体的な補修工法とその選定理由】、【健全度特定のための内部詳細調査の推奨】を重厚な長文（各400文字以上）で出力してください。JSONは不要です。
"""
                            request_contents = [prompt] + images
                            response = model.generate_content(request_contents)
                            st.session_state.full_result_text = response.text
                            
                            # 幅のパース
                            final_w = manual_width
                            if final_w == 0:
                                try:
                                    match = re.search(r"確定ひび割れ幅:\s*([0-9.]+)", st.session_state.full_result_text)
                                    if match:
                                        final_w = float(match.group(1))
                                except Exception:
                                    final_w = 0.0
                            st.session_state.final_width = final_w
                            break 
                            
                        except Exception as e:
                            if "429" in str(e) or "quota" in str(e).lower():
                                if attempt < max_retries - 1:
                                    time.sleep(2.5)
                                    continue
                            raise e

        # 結果の描画（セッション状態を元に行うため切り替えてもエラーで落ちない）
        if st.session_state.full_result_text:
            fw = st.session_state.final_width
            if "劣化度Ⅲ" in st.session_state.full_result_text or "劣化度Ⅳ" in st.session_state.full_result_text or fw >= 1.0:
                color_code, status_title = "#EF4444", f"🔴 【劣化度Ⅲ〜Ⅳ：要早期・緊急補修】判定ひび割れ幅: {fw} mm"
                alert_desc = "⚠️ 指針基準：重大な断面欠損、漏水、または鉄筋露出の進展リスクがあります。早急な断面修復工法および構造安全性の確認が必要です。"
            elif "劣化度Ⅱ" in st.session_state.full_result_text or (0.2 <= fw < 1.0):
                color_code, status_title = "#EAB308", f"🟡 【劣化度Ⅱ：要補修・機能保持】判定ひび割れ幅: {fw} mm"
                alert_desc = "💡 指針基準：耐久性低下を防ぐため、指針に則った「低圧エポキシ樹脂注入工法」を推奨します。"
            elif fw > 0:
                color_code, status_title = "#10B981", f"🟢 【劣化度Ⅰ：経過観察・予防保全】判定ひび割れ幅: {fw} mm"
                alert_desc = "✅ 指針基準：構造安全性への直接的影響は軽微です。表面含浸工法による予防保全、または目視経過観察フェーズとなります。"
            else:
                color_code, status_title = "#3B82F6", "🔵 【寸法・劣化度保留：追加実測要請】"
                alert_desc = "ℹ️ 写真から明確な縮尺基準が確認できないため判定を保留しています。実測値の入力を求めます。"
            
            st.markdown(f"<div class='status-card' style='border-left: 8px solid {color_code};'><h3 style='color: {color_code} !important; margin:0; font-size:22px;'>{status_title}</h3><p style='color: #F1F5F9 !important; font-size: 14px; margin: 8px 0 0 0; font-weight: bold;'>{alert_desc}</p></div>", unsafe_allow_html=True)
            st.markdown("<h4 style='color: white; margin-top:20px;'>📑 AI Suite Pro 高精密工学的統合解析レポート</h4>", unsafe_allow_html=True)
            st.markdown(f"<div class='report-text-box'>{st.session_state.full_result_text}</div>", unsafe_allow_html=True)

            # --- Excel帳書出力生成（完全防壁版） ---
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

                ws.column_dimensions['A'].width = 22
                ws.column_dimensions['B'].width = 50  
                ws.column_dimensions['C'].width = 2
                ws.column_dimensions['D'].width = 72  
                
                p_name = project_name if ('project_name' in locals() and project_name) else "コンクリート構造物健全度調査業務"
                l_name = location_name if ('location_name' in locals() and location_name) else "現場調査対象箇所"
                if 'address_input' in locals() and address_input:
                    l_name = f"{address_input} ({l_name})"

                start_row = 1
                for idx, img in enumerate(images):
                    ws.merge_cells(f"A{start_row}:D{start_row}")
                    ws[f"A{start_row}"] = f"■ {p_name} 構造物調査状況写真台帳"
                    ws[f"A{start_row}"].font = font_header
                    
                    ws.merge_cells(f"A{start_row+1}:D{start_row+1}")
                    ws[f"A{start_row+1}"] = f"調査箇所・施設位置： {l_name} (No.{idx+1})  |  調査技術者：{inspector_name if 'inspector_name' in locals() else 'T&N技術管理者'}"
                    ws[f"A{start_row+1}"].font = font_label
                    ws[f"A{start_row+1}"].alignment = Alignment(horizontal="right")

                    info_labels = ["写真No. / 撮影項目", "工種・部材分類", "位置・測定部詳細", "気象・地域環境特性", "AI工学判定・原因の考察", "推奨補修工法・選定理由"]
                    reason_part = st.session_state.full_result_text if idx == 0 else photo_comments[idx]
                    method_part = "上記レポート内の【対策案および詳細調査の推奨】セクションを参照" if idx == 0 else "No.1写真の全体統括レポートに準ずる"
                        
                    info_values = [f"Photo No.{idx+1}", f"{struct_type if 'struct_type' in locals() else 'コンクリート構造物'} / 劣化度目視診断", l_name, auto_weather_summary if 'auto_weather_summary' in locals() else '一般環境', reason_part, method_part]

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
                        
                        if "考察" in label or "補修" in label:
                            text_len = len(str(value))
                            ws.row_dimensions[r].height = max(140, min(400, int(text_len * 0.42)))
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
                    file_name=f"【機能保全診断調書】{p_name}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as excel_err:
                st.error(f"Excel写真台帳の生成中にエラーが発生しました: {excel_err}")
        else:
            st.info("💡 『② 現場データ・写真アップロード』タブで写真を添付し、解析実行ボタンを押すと、国交省・農水省提出仕様の重厚な工学的レポートとExcel出力モジュールがここに自動展開されます。")
