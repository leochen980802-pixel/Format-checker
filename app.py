import streamlit as st
import google.generativeai as genai
import docx
import fitz  # PyMuPDF
from pptx import Presentation
import json

# --- 1. 設定與初始化 ---
st.set_page_config(page_title="學術格式校對 AI", layout="centered")

# 從 Streamlit 雲端環境安全地取得 API Key
try:
    api_key = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=api_key)
except Exception as e:
    st.error("系統找不到 API 金鑰。請確認你已經在 Streamlit 的 Advanced settings 中設定了 GEMINI_API_KEY。")
    st.stop()

# --- 2. 核心功能：讀取不同格式的文件 ---
def extract_text(file, file_type):
    text = ""
    try:
        if file_type == "docx":
            doc = docx.Document(file)
            for para in doc.paragraphs:
                text += para.text + "\n"
        elif file_type == "pdf":
            doc = fitz.open(stream=file.read(), filetype="pdf")
            for page in doc:
                text += page.get_text()
        elif file_type == "pptx":
            ppt = Presentation(file)
            for slide in ppt.slides:
                for shape in slide.shapes:
                    if hasattr(shape, "text"):
                        text += shape.text + "\n"
    except Exception as e:
        st.error(f"檔案讀取失敗：{e}")
    return text

# --- 3. 核心功能：呼叫 AI 進行校對 ---
def proofread_text(text, format_style):
    # 【升級 1：改用 Pro 模型】處理複雜邏輯與細節的精準度遠高於 Flash
    target_model = "gemini-1.5-pro-latest" 
    
    try:
        # 動態尋找可用的 Pro 模型
        for m in genai.list_models():
            if 'generateContent' in m.supported_generation_methods:
                if 'pro' in m.name:
                    target_model = m.name
                    break
    except Exception as e:
        pass

    model = genai.GenerativeModel(target_model)
    
    # 【升級 2：強化提示詞】賦予專家身分、明確的檢查步驟與極度嚴格的語氣
    prompt = f"""
    你現在是一位擁有 20 年經驗的嚴苛學術期刊主編。你的任務是進行極度精準的文字與格式校對。
    請嚴格執行以下步驟：
    1. 逐字掃描，找出「所有」錯別字、漏字與不通順的文法。
    2. 嚴格檢查格式是否「完全」符合【{format_style}】規範（包含引註格式、標點符號全半形、大小寫等）。
    3. 寧可嚴格，不可錯漏。必須抓出所有問題。

    請務必以 JSON 陣列格式回傳，絕對不要包含任何其他文字、問候語或 Markdown 標記，格式嚴格如下：
    [
      {{"original": "錯誤的句子或單字", "issue": "具體的錯誤原因說明", "fix": "建議的精確修改內容"}}
    ]
    
    待校對文字：
    {text[:4000]} 
    """
    
    try:
        # 【升級 3：鎖死 Temperature】將隨機性降到 0.0，確保每次輸出一致且具最高邏輯性
        response = model.generate_content(
            prompt,
            generation_config={"temperature": 0.0}
        )
        result_text = response.text
        
        # 暴力萃取法 (保持不變，確保系統不崩潰)
        start_idx = result_text.find('[')
        end_idx = result_text.rfind(']')
        
        if start_idx != -1 and end_idx != -1:
            clean_json = result_text[start_idx:end_idx+1]
            return json.loads(clean_json)
        else:
            st.error("AI 未回傳標準的 JSON 格式，請再試一次。")
            return None
            
    except Exception as e:
        st.error(f"AI 解析資料失敗，錯誤訊息：{e}")
        return None


    # 備註：初期測試先截取前 4000 字，避免運算時間過長或超出 API 限制
    
    try:
        response = model.generate_content(prompt)
        # 清理 AI 可能產生的 Markdown 標籤，確保 JSON 格式正確
        result_text = response.text.strip().replace('```json', '').replace('```', '').strip()
        return json.loads(result_text)
    except Exception as e:
        st.error(f"AI 校對失敗，請檢查文字內容或稍後再試：{e}")
        return []

# --- 4. 介面與互動邏輯 ---
st.title("📄 學術格式與錯字校對工具")
st.markdown("上傳你的 Word, PDF 或 PPT，AI 將自動抓出格式瑕疵與錯字。")

st.sidebar.header("設定")
format_choice = st.sidebar.selectbox(
    "請選擇目標格式規範：",
    ["APA 格式", "MLA 格式", "Chicago 格式"]
)

uploaded_file = st.file_uploader(
    "上傳文件 (支援 .docx, .pdf, .pptx)", 
    type=["docx", "pdf", "pptx"]
)

if uploaded_file is not None:
    # 取得檔案副檔名
    file_ext = uploaded_file.name.split(".")[-1].lower()
    st.success(f"已成功上傳：{uploaded_file.name}")
    
    if st.button("開始 AI 校對"):
        with st.spinner("AI 正在努力閱讀與校對中，這可能需要幾十秒，請稍候..."):
            
            # 1. 將上傳的檔案轉換為純文字
            raw_text = extract_text(uploaded_file, file_ext)
            
            if not raw_text.strip():
                st.warning("無法從檔案中讀取到文字，請確認檔案內容並非純圖片。")
            else:
            # 2. 將文字與選擇的格式送給 API
                issues = proofread_text(raw_text, format_choice)
                
            # 3. 將陣列中的錯誤一條一條列印在畫面上
            if issues is None:
                # 如果是 None，代表上面已經印出紅字錯誤了，這裡什麼都不做 (pass)
                pass
            elif len(issues) > 0:
                    st.subheader("📝 發現以下格式或錯字問題：")
                    for idx, item in enumerate(issues):
                        with st.expander(f"問題 {idx + 1}：{item.get('issue', '格式問題')}"):
                            st.write("**原文：**", item.get("original", ""))
                            st.write("**建議修改：**", item.get("fix", ""))
            else:
                    # 只有真正回傳了空陣列 []，才代表完全沒錯字
                    st.success("太棒了！AI 沒有發現明顯的格式錯誤或錯字。")

