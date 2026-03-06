# --- 側邊欄：設定 API Key 與 發件人 ---
st.sidebar.header("1. 系統設定")

# 嘗試從 Streamlit Secrets 讀取，讀不到再開放輸入框
default_api_key = st.secrets.get("SENDGRID_API_KEY", "")
default_sender = st.secrets.get("SENDER_EMAIL", "")

api_key = st.sidebar.text_input("輸入 SendGrid API Key", value=default_api_key, type="password")
sender_email = st.sidebar.text_input("輸入預設發件人 Email", value=default_sender)
subject = st.sidebar.text_input("輸入信件主旨", "這是一封測試信件")
