import streamlit as st

# 設定網頁標題與排版
st.set_page_config(page_title="學術格式校對 AI", layout="centered")

st.title("📄 學術格式與錯字校對器")
st.markdown("上傳你的 Word, PDF 或 PPT，AI 將自動抓出格式瑕疵與錯字。")

# 側邊欄：格式選擇
st.sidebar.header("設定")
format_choice = st.sidebar.selectbox(
    "請選擇目標格式規範：",
    ["APA 格式", "MLA 格式", "Chicago 格式"]
)

# 主畫面：檔案上傳區塊
uploaded_file = st.file_uploader(
    "上傳文件 (支援 .docx, .pdf, .pptx)", 
    type=["docx", "pdf", "pptx"]
)

if uploaded_file is not None:
    st.success(f"已成功上傳：{uploaded_file.name}")
    st.info(f"目前選擇的校對標準：{format_choice}")

    # 預留按鈕，之後用來觸發 AI 運算
    if st.button("開始 AI 校對"):
        st.warning("AI 解析引擎尚未串接，請等待下一步開發！")
