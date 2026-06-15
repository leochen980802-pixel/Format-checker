import streamlit as st
import json
import os
from datetime import datetime
import google.generativeai as genai
import time

# ==========================================
# 0. Gemini API 金鑰配置
# ==========================================
api_key = os.environ.get("GEMINI_API_KEY", "")
if not api_key:
    try:
        api_key = st.secrets["GEMINI_API_KEY"]
    except Exception:
        pass

if api_key:
    genai.configure(api_key=api_key)
else:
    st.sidebar.warning("🔑 未偵測到系統環境金鑰，請確保您已配置 GEMINI_API_KEY")

# ==========================================
# 1. 歷史紀錄系統檔案儲存邏輯 (嚴格縮排防呆)
# ==========================================
HISTORY_FILE = "proofread_history.json"

def load_history():
    if os.path.exists(HISTORY_FILE):
        with open(HISTORY_FILE, "r", encoding="utf-8") as f:
            try:
                return json.load(f)
            except json.JSONDecodeError:
                return []
    return []

def save_history(file_name, results):
    history = load_history()
    new_record = {
        "id": datetime.now().strftime("%Y%m%d%H%M%S"),
        "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "file_name": file_name,
        "results": results
    }
    history.insert(0, new_record) 
    
    with open(HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump(history, f, ensure_ascii=False, indent=4)

# ==========================================
# 2. 多格式檔案文字萃取器 (完美對應你的 PyMuPDF)
# ==========================================
def extract_text_from_file(file):
    filename = file.name.lower()
    
    # 處理純文字檔
    if filename.endswith('.txt'):
        return file.getvalue().decode("utf-8", errors="ignore")
        
    # 處理 PDF 檔 (對應你 requirements.txt 中的 PyMuPDF)
    elif filename.endswith('.pdf'):
        try:
            import fitz  # PyMuPDF 的官方內引導模組名稱
            # 讀取 Streamlit 的檔案緩衝區
            doc = fitz.open(stream=file.getvalue(), filetype="pdf")
            text_runs = []
            for page in doc:
                text_runs.append(page.get_text())
            return "".join(text_runs)
        except ImportError:
            st.error("🚨 系統未正確載入 PyMuPDF 套件，請確認雲端部署狀態。")
            return ""
        except Exception as e:
            st.error(f"🚨 PDF 解析發生未知錯誤：{e}")
            return ""
            
    # 處理 PPTX 簡報檔
    elif filename.endswith('.pptx'):
        try:
            from pptx import Presentation
            prs = Presentation(file)
            text_runs = []
            for slide in prs.slides:
                for shape in slide.shapes:
                    if hasattr(shape, "text") and shape.text.strip():
                        text_runs.append(shape.text)
            return "\n".join(text_runs)
        except ImportError:
            st.error("🚨 系統未安裝 python-pptx 套件。")
            return ""
            
    return ""

# ==========================================
# 3. AI 核心校對大腦 (分段滑動視窗 + 原生 JSON)
# ==========================================
def proofread_text(text, format_style):
    target_model = "gemini-1.5-flash" 
    try:
        model = genai.GenerativeModel(target_model)
    except Exception as e:
        st.error(f"模型載入失敗：{e}")
        return None

    chunk_size = 1500
    text_chunks = [text[i:i+chunk_size] for i in range(0, len(text), chunk_size)]
    
    all_results = []
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    for idx, chunk in enumerate(text_chunks):
        status_text.text(f"正在進行深度校對：第 {idx+1} / {len(text_chunks)} 段文字段落...")
        
        prompt = f"""
        你現在是一位擁有 20 年經驗的嚴苛學術期刊主編。你的任務是進行極度精準的文字與格式校對。
        請嚴格執行以下步驟：
        1. 逐字掃描，找出「所有」錯別字、漏字與不通順的文法。
        2. 嚴格檢查格式是否「完全」符合【{format_style}】規範（包含引註格式、標點符號全半形、大小寫等）。

        你必須嚴格回傳一個 JSON 陣列，格式如下：
        [
          {{"original": "錯誤的句子或單字", "issue": "具體的錯誤原因說明", "fix": "建議的精確修改內容"}}
        ]
        
        待校對文字段落：
        {chunk} 
        """
        
        try:
            response = model.generate_content(
                prompt,
                generation_config={
                    "temperature": 0.0,
                    "response_mime_type": "application/json"
                }
            )
            
            chunk_result = json.loads(response.text)
            if isinstance(chunk_result, list):
                all_results.extend(chunk_result)
                
        except Exception as e:
            error_msg = str(e)
            if "429" in error_msg or "Quota" in error_msg:
                status_text.text("⚠️ 觸發免費版頻率限制，系統自動等待 30 秒後繼續...")
                time.sleep(30)
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
                st.error(f"第 {idx+1} 段解析失敗，已自動跳過。錯誤：{error_msg}")
        
        time.sleep(4)
        progress_bar.progress((idx + 1) / len(text_chunks))
        
    status_text.text("✨ 全本簡報深度校對完成！")
    time.sleep(1)
    progress_bar.empty()
    status_text.empty()
    
    return all_results

# ==========================================
# 4. 側邊欄：歷史紀錄切換與顯示邏輯
# ==========================================
st.sidebar.title("🗂️ 校對歷史紀錄")
history_data = load_history()

if "selected_history" not in st.session_state:
    st.session_state.selected_history = None
if "current_viewing_file" not in st.session_state:
    st.session_state.current_viewing_file = ""

if not history_data:
    st.sidebar.info("目前還沒有任何校對紀錄。")
else:
    for record in history_data:
        button_label = f"📄 {record['file_name']} \n({record['time']})"
        if st.sidebar.button(button_label, key=record['id']):
            st.session_state.selected_history = record['results']
            st.session_state.current_viewing_file = record['file_name']

if st.session_state.selected_history:
    st.info(f"👀 歷史回顧模式：正在查看【{st.session_state.current_viewing_file}】的校對結果")
    if st.button("⬅️ 返回主畫面（新增校對）"):
        st.session_state.selected_history = None
        st.rerun()
        
    for item in st.session_state.selected_history:
        st.error(f"❌ 原文：{item.get('original', '')}")
        st.warning(f"💡 問題：{item.get('issue', '')}")
        st.success(f"✅ 建議：{item.get('fix', '')}")
        st.markdown("---")
        
    st.stop()

# ==========================================
# 5. 主要運作 UI 介面 (格式選擇永遠置頂可選)
# ==========================================
st.title("📝 高規格論文/簡報 AI 完美校對工具")
st.write("支援上傳 `.txt`, `.pdf`, `.pptx` 格式檔案。系統將採用滑動視窗進行無死角逐字審查。")

# 💡 修正：將格式選擇移到最外層，不管有沒有傳檔案都能直接選！
format_style = st.selectbox(
    "1️⃣ 請選擇您要遵循的學術/排版格式規範：", 
    ["APA 格式規範", "Chicago 格式規範", "MLA 格式規範", "通用商業精準簡報規範"]
)

uploaded_file = st.file_uploader("2️⃣ 請上傳您的簡報或文稿檔案", type=["txt", "pdf", "pptx"])

if uploaded_file is not None:
    with st.spinner("正在解析檔案內的所有文字內容..."):
        extracted_text = extract_text_from_file(uploaded_file)
    
    if not extracted_text.strip():
        st.warning("⚠️ 無法從檔案中擷取出有效文字，請確認該檔案非純圖片掃描檔。")
    else:
        st.success(f"成功載入檔案！總字數約為 {len(extracted_text)} 字。")
        
        if st.button("🚀 3️⃣ 開始全自動分段深度校對"):
            all_results = proofread_text(extracted_text, format_style)
            
            if all_results:
                save_history(uploaded_file.name, all_results)
                st.success("🎉 全本校對完成！結果已同步備份至左側歷史紀錄面板。")
                
                for item in all_results:
                    st.error(f"❌ 原文：{item.get('original', '')}")
                    st.warning(f"💡 問題：{item.get('issue', '')}")
                    st.success(f"✅ 建議：{item.get('fix', '')}")
                    st.markdown("---")
