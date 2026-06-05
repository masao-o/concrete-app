import streamlit as st
import google.generativeai as genai
from PIL import Image
import os

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
    st.title("コンクリート劣化診断 AI Suite Pro")
    st.markdown("構造物の写真をアップロードすると、AIが劣化原因の推測と対策案を提示します。")

    # APIキーの取得
    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        api_key = st.secrets.get("GEMINI_API_KEY", "")

    uploaded_file = st.file_uploader("写真をアップロードしてください", type=["jpg", "jpeg", "png"])

    if uploaded_file is not None:
        image = Image.open(uploaded_file)
        st.image(image, caption="アップロードされた写真", use_container_width=True)
        
        if st.button("AI診断を開始"):
            if not api_key:
                st.error("APIキーが設定されていません。")
            else:
                with st.spinner("AIが写真を解析中..."):
                    try:
                        genai.configure(api_key=api_key)
                        model = genai.GenerativeModel('gemini-1.5-pro')
                        
                        prompt = "このコンクリート写真の劣化状況（ひび割れ、エフロレッセンス、剥離など）をプロの視点で診断し、原因の推測と適切な対策案を詳しく日本語で説明してください。"
                        response = model.generate_content([prompt, image])
                        
                        st.markdown("### 📊 診断結果")
                        st.write(response.text)
                    except Exception as e:
                        st.error(f"エラーが発生しました: {e}")
