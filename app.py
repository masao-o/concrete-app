import streamlit as st
import google.generativeai as genai
from PIL import Image

# パスワードセッションの初期化
if 'authenticated' not in st.session_state:
    st.session_state['authenticated'] = False

def check_password():
    def password_entered():
        if st.session_state["password"] == "masao0605":
            st.session_state["authenticated"] = True
            del st.session_state["password"]
        else:
            st.sidebar.error("パスワードが違います")
            
    if not st.session_state["authenticated"]:
        st.title("🔒 コンクリート劣化診断 AI Suite Pro")
        st.markdown("このアプリは関係者限定です。パスワードを入力してください。")
        st.text_input("パスワード", type="password", on_change=password_entered, key="password")
        return False
    return True

if check_password():
    st.title("🏗️ コンクリート劣化診断 AI Suite Pro")
    st.markdown("構造物の写真と環境条件を組み合わせ、AIがプロの視点で高精度な診断を行います。")
    st.markdown("---")

    # 📘 ユーザーから大好評だった「アプリの使い方・説明書」
    with st.expander("📘 本アプリの取扱説明書（マニュアル）を開く", expanded=False):
        st.markdown("""
        ### 【アプリの使い方】
        1. **環境条件の選択：** 左側のメニューから、対象構造物の「置かれている環境」や「乾湿状態」などを正確に選択してください。
        2. **写真のアップロード：** 診断したいコンクリート構造物のひび割れや劣化部分の写真をアップロードします。
        3. **AI診断の実行：** 「AI診断を開始」ボタンを押すと、選択された環境条件を加味して、AIが劣化原因の推測と対策案を詳しく解説します。
        
        ### 【診断のヒント】
        * 写真はできるだけ明るく、ピントが合ったものをアップロードすると精度が向上します。
        * クラックスケールなどが一緒に写っていると、AIがより具体的な規模を認識しやすくなります。
        """)

    # 直接本物のAPIキーを指定（エラーを絶対に起こさない対策）
    api_key = "AIzaSyD-O647K9Xg-mH4N0_Prc"

    # 🛠️ 大復活：「いろんな環境・条件を選択できる」サイドバーメニュー
    st.sidebar.header("🛠️ 構造物の環境・条件設定")
    
    env_location = st.sidebar.selectbox(
        "① 設置環境（場所）",
        ["屋外（雨が直接当たる）", "屋外（軒下など日陰）", "屋内（常時乾燥）", "地下・土中", "海岸付近（塩害の恐れあり）"]
    )
    
    wet_status = st.sidebar.selectbox(
        "② コンクリートの乾湿状態",
        ["常時乾燥している", "周期的に湿潤と乾燥を繰り返す", "常時湿潤・水分を含んでいる"]
    )
    
    crack_type = st.sidebar.selectbox(
        "③ 目視での主な症状（任意）",
        ["ひび割れ（クラック）のみ", "エフロレッセンス（白華現象）がある", "コンクリートの剥離・鉄筋露出あり", "全体的な変色・劣化"]
    )

    # メイン画面のファイルアップローダー
    uploaded_file = st.file_uploader("写真をアップロードしてください", type=["jpg", "jpeg", "png"])

    if uploaded_file is not None:
        image = Image.open(uploaded_file)
        st.image(image, caption="アップロードされた調査写真", use_container_width=True)
        
        st.markdown("---")
        if st.button("🔍 選択した条件でAI診断を開始"):
            with st.spinner("AIが環境条件と写真を総合的に解析中..."):
                try:
                    genai.configure(api_key=api_key)
                    model = genai.GenerativeModel('gemini-1.5-pro')
                    
                    # 選択された環境条件をAIへの命令（プロンプト）に自動で組み込む
                    prompt = f"""
                    あなたはコンクリート診断士の最高峰プロフェッショナルです。
                    以下の【調査対象の環境条件】と【アップロードされた写真】を総合的に分析し、プロの視点で非常に詳しく日本語で診断してください。
                    
                    【調査対象の環境条件】
                    ・設置環境: {env_location}
                    ・乾湿状態: {wet_status}
                    ・目視の主な症状: {crack_type}
                    
                    【出力すべき内容】
                    1. 総合的な劣化状況の診断（ひび割れ、エフロ、剥離などの状態）
                    2. この環境条件だからこそ考えられる「劣化原因の深い推測」
                    3. 実務に基づいた、今後の適切な「対策案・補修工法の提案」
                    """
                    
                    response = model.generate_content([prompt, image])
                    
                    st.markdown("### 📊 プロのAI診断結果")
                    st.write(response.text)
                except Exception as e:
                    st.error(f"エラーが発生しました: {e}")
