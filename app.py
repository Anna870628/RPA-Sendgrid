import streamlit as st
import pandas as pd
import mammoth
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import Mail

# 設定網頁標題與寬度
st.set_page_config(page_title="SendGrid 批次發信系統", layout="wide")

st.title("✉️ SendGrid 批次發信系統")

# --- 安全讀取 Secrets (避免本機測試時找不到檔案而報錯) ---
try:
    default_api_key = st.secrets.get("SENDGRID_API_KEY", "")
    default_sender = st.secrets.get("SENDER_EMAIL", "")
except FileNotFoundError:
    default_api_key = ""
    default_sender = ""

# --- 側邊欄：設定 API Key 與 發件人 ---
st.sidebar.header("1. 系統設定")
api_key = st.sidebar.text_input("輸入 SendGrid API Key", value=default_api_key, type="password")
sender_email = st.sidebar.text_input("輸入預設發件人 Email", value=default_sender)
subject = st.sidebar.text_input("輸入信件主旨", "這是一封測試信件")

# --- 區塊 1：名單匯入與瀏覽 ---
st.header("2. 匯入收件人名單")
excel_file = st.file_uploader("上傳 Excel 檔案 (.xlsx)", type=["xlsx"])

recipient_emails = []

if excel_file is not None:
    df = pd.read_excel(excel_file)
    st.write("📋 **名單預覽：**")
    st.dataframe(df, height=200)
    
    # 讓使用者選擇哪一個欄位是 Email
    email_col = st.selectbox("請選擇作為「收件信箱」的欄位", df.columns)
    
    if email_col:
        recipient_emails = df[email_col].dropna().tolist()
        st.success(f"已成功載入 {len(recipient_emails)} 筆 Email 名單！")

st.divider()

# --- 區塊 2：匯入 Word 內容與編輯 ---
st.header("3. 編輯信件內容")
word_file = st.file_uploader("上傳 Word 檔案 (.docx) 作為信件草稿", type=["docx"])

html_content = ""

if word_file is not None:
    # 使用 mammoth 將 docx 轉換為 HTML 保留格式
    result = mammoth.convert_to_html(word_file)
    html_content = result.value

st.markdown("📝 **手動編輯內容區 (HTML格式支援)：**")
edited_html = st.text_area("你可以在這裡直接修改文字或加入 HTML 標籤（如 <b>粗體</b>, <br>換行）", value=html_content, height=200)

st.divider()

# --- 區塊 3：預覽實際發出畫面 ---
st.header("4. 信件預覽")
st.markdown("👀 **收件人看到的信件畫面將如下：**")
with st.container(border=True):
    st.components.v1.html(edited_html, height=300, scrolling=True)

st.divider()

# --- 區塊 4：執行發送 ---
st.header("5. 確認發送")
if st.button("🚀 點擊發送給所有名單", type="primary"):
    if not api_key or not sender_email:
        st.error("請先在左側邊欄填寫 SendGrid API Key 與發件人 Email！")
    elif not recipient_emails:
        st.error("請先上傳名單並選擇正確的 Email 欄位！")
    elif not edited_html:
        st.error("信件內容不可為空！")
    else:
        success_count = 0
        error_count = 0
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        sg = SendGridAPIClient(api_key)
        
        for i, email in enumerate(recipient_emails):
            message = Mail(
                from_email=sender_email,
                to_emails=email,
                subject=subject,
                html_content=edited_html
            )
            try:
                response = sg.send(message)
                if response.status_code in [200, 201, 202]:
                    success_count += 1
            except Exception as e:
                error_count += 1
                st.error(f"發送至 {email} 失敗: {e}")
            
            progress_bar.progress((i + 1) / len(recipient_emails))
            status_text.text(f"正在發送... ({i+1}/{len(recipient_emails)})")
            
        st.success(f"🎉 發送完成！成功: {success_count} 筆，失敗: {error_count} 筆。")
