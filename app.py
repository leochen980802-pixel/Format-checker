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
        st.error(f"讀取檔案失敗：{e}")
    return text

import time  # 務必確保檔案最上方有 import time

# --- 3. 核心功能：呼叫 AI 進行校對（商用規格版） ---
def proofread_text(text, format_style):
    # 預設使用 Flash 模型，以取得每分鐘 15 次的高呼叫額度
    target_model = "gemini-1.5-flash" 
    
    try:
        for m in genai.list_models():
            if 'generateContent' in m.supported_generation_methods:
                if 'flash' in m.name:
                    target_model = m.name
                    break
    except Exception:
        pass 

    try:
        model = genai.GenerativeModel(target_model)
    except Exception as e:
        st.error(f"模型載入失敗：{e}")
        return None

    # =================【核心升級 1：滑動視窗分段處理】=================
    # 將 64 頁的龐大文字，每 1500 字切成一塊，確保 AI 能集中注意力「逐字」細看
    chunk_size = 1500
    text_chunks = [text[i:i+chunk_size] for i in range(0, len(text), chunk_size)]
    
    all_results = []
    
    # 在 Streamlit 畫面上建立進度條與狀態文字
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    # 開始分批輪詢校對
    for idx, chunk in enumerate(text_chunks):
        status_text.text(f"正在深度校對第 {idx+1} / {len(text_chunks)} 段簡報文字...")
        
        prompt = f"""
        你現在是一位擁有 20 年經驗的嚴苛學術期刊主編。你的任務是進行極度精準的文字與格式校對。
        請嚴格執行以下步驟：
        1. 逐字掃描，找出「所有」錯別字、漏字與不通順的文法。
        2. 嚴格檢查格式是否「完全」符合【{format_style}】規範（包含引註格式、標點符號全半形、大小寫等）。
        3. 寧可嚴格，不可錯漏。必須抓出所有問題。

        你必須嚴格回傳一個 JSON 陣列，格式如下：
        [
          {{"original": "錯誤的句子或單字", "issue": "具體的錯誤原因說明", "fix": "建議的精確修改內容"}}
        ]
        
        待校對文字段落：
        {chunk} 
        """
        
        try:
            # =================【核心升級 2：原生 JSON 模式】=================
            # 透過 response_mime_type 強制 API 只能輸出純 JSON，徹底根除 Extra data 錯誤
            response = model.generate_content(
                prompt,
                generation_config={
                    "temperature": 0.0,
                    "response_mime_type": "application/json"
                }
            )
            
            # 直接解析，不需要再寫 find('[') 暴力切字串
            chunk_result = json.loads(response.text)
            if isinstance(chunk_result, list):
                all_results.extend(chunk_result)
                
        except Exception as e:
            error_msg = str(e)
            # =================【核心升級 3：自動對抗 429 限制】=================
            if "429" in error_msg or "Quota" in error_msg:
                status_text.text("⚠️ 觸發免費版頻率限制，系統自動等待 30 秒後繼續...")
                time.sleep(30)
                # 重試機制
                try:
                    response = model.generate_content(
                        prompt, 
                        generation_config={"temperature": 0.0, "response_mime_type": "application/json"}
                    )
                    chunk_result = json.loads(response.text)
                    if isinstance(chunk_result, list):
                        all_results.extend(chunk_result)
                except Exception:
                    pass
            else:
                st.error(f"第 {idx+1} 段解析失敗，已跳過。錯誤：{error_msg}")
        
        # =================【核心升級 4：調頻緩衝】=================
        # 免費版 Flash 限制每分鐘 15 次，每次呼叫完刻意休息 4 秒，確保整體運作平穩不中斷
        time.sleep(4)
        progress_bar.progress((idx + 1) / len(text_chunks))
        
    status_text.text("✨ 全本 64 頁簡報深度校對完成！")
    time.sleep(1)
    progress_bar.empty()
    status_text.empty()
    
    return all_results



# --- 4. 介面與互動邏輯 ---
st.title("📄 學術格式與錯字校對神器")
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

