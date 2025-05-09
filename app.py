import streamlit as st
from docx import Document
from collections import defaultdict
import matplotlib.pyplot as plt
import re
from datetime import datetime
import io
from matplotlib.font_manager import FontProperties
import pandas as pd
from bs4 import BeautifulSoup
import requests
import jieba
import jieba.analyse
from collections import Counter
from tqdm import tqdm  # Can be used with Streamlit, but careful with output
import os

# --- Font Setup (Local Workaround) ---
# This is NOT deployment-friendly. For deployment, ensure the font is available
# in the environment.
LOCAL_FONT_PATH = './NotoSansCJKjp-Regular.otf'  # Change if you have a different path
font_installed = False
if os.path.exists(LOCAL_FONT_PATH):
    try:
        font = FontProperties(fname=LOCAL_FONT_PATH)
        plt.rcParams['font.sans-serif'] = ['Noto Sans CJK TC', 'Noto Sans CJK SC', 'Noto Sans CJK JP']
        plt.rcParams['axes.unicode_minus'] = False
        font_installed = True
    except Exception as e:
        st.warning(f"Font loading issue: {e}. Plots may not display correctly.")
else:
    st.warning("Font file not found. Plots may not display correctly.")

# --- Helper Functions ---
# --- Helper Functions ---
def process_uploaded_files(uploaded_files, column_names):  # 移除 @st.cache_data 裝飾器
    date_pattern = re.compile(r"\b(202[0-5])(?:-|年|\/|\.)(0?[1-9]|1[0-2])(?:-|月|\/|\.)(0?[1-9]|[12]\d|3[01])(?:日)?\b")
    year_month_counts = defaultdict(lambda: defaultdict(int))
    for name, content in zip(column_names, uploaded_files):
        try:
            file_bytes = content.getvalue()  # 使用 getvalue() 來獲取檔案內容
            doc = Document(io.BytesIO(file_bytes))
            for para in doc.paragraphs:
                match = date_pattern.search(para.text)
                if match:
                    year_str, month_str, _ = match.groups()
                    try:
                        date_obj = datetime(int(year_str), int(month_str), 1)
                        start_date = datetime(2020, 1, 1)
                        end_date = datetime(2025, 4, 30)
                        if start_date <= date_obj <= end_date:
                            year_month_key = f"{year_str}-{int(month_str):02d}"
                            year_month_counts[name][year_month_key] += 1
                    except ValueError:
                        continue
        except Exception as e:
            st.error(f"Error processing file {name}: {e}")
    
    # 將 defaultdict(lambda: defaultdict(int)) 轉換為一般的 dict
    result = {}
    for name, counts in year_month_counts.items():
        result[name] = dict(counts)
    
    return result

def plot_data(year_month_counts, original_column_names, new_column_names):
    all_months = sorted(set(k for d in year_month_counts.values() for k in d.keys()))
    
    if not all_months:
        st.warning("沒有找到符合條件的日期數據，無法繪製圖表。")
        return
        
    fig, ax = plt.subplots(figsize=(12, 6))  # Create figure and axes
    
    for i, original_name in enumerate(original_column_names):
        if original_name in year_month_counts:
            counts = year_month_counts[original_name]
            y_vals = [counts.get(month, 0) for month in all_months]
            label_name = new_column_names[i] if i < len(new_column_names) else original_name
            
            if font_installed:
                ax.plot(all_months, y_vals, label=label_name, marker='o')
            else:
                ax.plot(all_months, y_vals, label=label_name, marker='o')
    
    if font_installed:
        ax.set_title("2020年1月至2025年4月 國台辦各欄目新聞稿數量變化", fontproperties=font)
        ax.set_xlabel("年份-月份", fontproperties=font)
        ax.set_ylabel("新聞稿數量", fontproperties=font)
    else:
        ax.set_title("2020年1月至2025年4月 國台辦各欄目新聞稿數量變化")
        ax.set_xlabel("年份-月份")
        ax.set_ylabel("新聞稿數量")
    
    if len(all_months) > 12:
        # 如果月份太多，只顯示部分月份標籤
        step = len(all_months) // 12
        plt.xticks(all_months[::step], rotation=45)
    else:
        plt.xticks(rotation=45)
        
    ax.legend()
    plt.tight_layout()
    plt.grid(True)
    st.pyplot(fig)  # Use st.pyplot to display the plot

def parse_list_docx(file_content):
    try:
        doc = Document(io.BytesIO(file_content))
        text_content = "\n".join(p.text for p in doc.paragraphs)
        items = []
        
        # 嘗試解析文檔中的項目
        # 注意：因為這是從Word檔解析而不是真正的HTML，此處可能需要調整
        lines = text_content.split('\n')
        for line in lines:
            date_match = re.search(r'\[(202\d-\d{1,2}-\d{1,2})\]', line)
            if date_match:
                date_str = date_match.group(1)
                # 從日期後取得標題和URL
                title_match = re.search(r'\]\s*(.*?)(?:\s*http|$)', line)
                url_match = re.search(r'(https?://[^\s]+)', line)
                
                title = title_match.group(1).strip() if title_match else ""
                url = url_match.group(1).strip() if url_match else ""
                
                try:
                    dt = datetime.strptime(date_str, "%Y-%m-%d")
                    if datetime(2020, 1, 1) <= dt <= datetime(2025, 4, 30):
                        items.append({"date": dt, "title": title, "url": url})
                except:
                    continue
        
        return items
    except Exception as e:
        st.error(f"解析文件時發生錯誤: {e}")
        return []

def fetch_content(url):
    try:
        resp = requests.get(url, timeout=5)
        resp.encoding = 'gb2312'
        soup = BeautifulSoup(resp.text, 'html.parser')
        editor = soup.find('div', class_='TRS_Editor')
        return editor.get_text(separator='\n').strip() if editor else ''
    except Exception as e:
        st.warning(f"⚠️ Could not read {url}: {e}")
        return ''

# --- Main Streamlit App ---
st.title("國台辦新聞稿分析")

# 第一部分：檔案上傳與新聞稿數量變化
st.header("新聞稿數量變化分析")
uploaded_files_count = st.file_uploader("上傳五個Word檔案", type="docx", accept_multiple_files=True)
new_column_names = ["台辦動態", "交流交往", "政務要聞", "部門涉台", "新聞發佈"]

if uploaded_files_count and len(uploaded_files_count) == 5:
    with st.spinner("正在處理檔案並分析數據..."):
        original_column_names = [file.name for file in uploaded_files_count]  # Get filenames
        year_month_counts = process_uploaded_files(uploaded_files_count, original_column_names)
        if year_month_counts:
            plot_data(year_month_counts, original_column_names, new_column_names)
        else:
            st.warning("沒有找到有效的日期數據來繪製圖表。")
elif uploaded_files_count:
    st.warning(f"請上傳五個檔案。目前已上傳 {len(uploaded_files_count)} 個。")
else:
    st.info("請上傳五個Word檔案以開始分析。")

# 分隔線
st.markdown("---")

# 第二部分：新聞稿關鍵詞分析
st.header("單篇新聞稿關鍵詞分析")
uploaded_file_keyword = st.file_uploader("上傳單個Word檔案以進行關鍵詞分析", type="docx")

if uploaded_file_keyword is not None:
    try:
        file_bytes = uploaded_file_keyword.getvalue()  # 使用 getvalue() 獲取檔案內容
        document = Document(io.BytesIO(file_bytes))
        text_for_analysis = "\n".join([para.text for para in document.paragraphs])
        pattern = r"\b(2020|2021|2022|2023|2024|2025)-(0[1-9]|1[0-2])-(0[1-9]|[12]\d|3[01])\b"
        matches = re.findall(pattern, text_for_analysis)
        year_months = [f"{y}-{m}" for y, m, d in matches]
        counts = Counter(year_months)
        
        years = list(range(2020, 2026))
        months = [f"{i:02d}" for i in range(1, 13)]
        data = []
        for y in years:
            row = {'year': y}
            for m in months:
                row[m] = counts.get(f"{y}-{m}", 0)
            data.append(row)
        
        df = pd.DataFrame(data)
        st.subheader("新聞稿每月數量統計")
        st.dataframe(df)
        
        df_long = df.melt(id_vars='year', value_vars=months, var_name='month', value_name='count')
        df_long['date'] = pd.to_datetime(df_long['year'].astype(str) + '-' + df_long['month'])
        df_long = df_long[df_long['date'] <= '2025-04-30']
        df_long = df_long.sort_values('date')
        
        fig, ax = plt.subplots(figsize=(12, 6))
        ax.plot(df_long['date'], df_long['count'], marker='o')
        ax.set_xlabel('日期')
        ax.set_ylabel('新聞稿數量')
        ax.set_title('新聞稿每月數量 (2020–2025.04)')
        plt.xticks(rotation=45)
        plt.grid(True)
        plt.tight_layout()
        st.pyplot(fig)
    except Exception as e:
        st.error(f"Error processing document: {e}")

# 分隔線
st.markdown("---")

# 第三部分：新聞列表解析與全文爬取
st.header("新聞列表解析與全文爬取")
uploaded_file_list = st.file_uploader("上傳包含新聞列表的Word檔案", type="docx", key="file_uploader_list")

if uploaded_file_list:
    try:
        file_bytes = uploaded_file_list.getvalue()  # 使用 getvalue() 獲取檔案內容
        all_items = parse_list_docx(file_bytes)
        
        if all_items:
            st.success(f"✅ 擷取到 {len(all_items)} 篇新聞列表")
            
            contents = []
            titles = []
            
            # Use st.spinner() to show a loading message during scraping
            with st.spinner(f"⏳ 正在爬取全部 {len(all_items)} 篇新聞稿全文..."):
                for i, item in enumerate(all_items):
                    titles.append(item["title"])
                    if item["url"].startswith("http"):
                        content = fetch_content(item["url"])
                        contents.append(content)
                    else:
                        contents.append("")
            
            successful_fetches = sum(1 for c in contents if c)
            st.success(f"✅ 完成，共成功擷取 {successful_fetches} 篇內容。")
        else:
            st.warning("未能從文件中擷取到新聞列表，請確認文件格式是否正確。")
    except Exception as e:
        st.error(f"處理文件時發生錯誤: {e}")
        
        all_text = "\n".join(contents)
        keywords_weighted = jieba.analyse.extract_tags(
            all_text, topK=30, withWeight=True, allowPOS=("n", "v", "vn", "nr", "ns")
        )
        
        st.subheader("前 30 關鍵詞（含權重）")
        for word, weight in keywords_weighted:
            st.write(f"{word}: {weight:.3f}")
        
        filtered_keywords = [w for w, weight in keywords_weighted]
        all_words = jieba.lcut(all_text)
        filtered_words = [w for w in all_words if w in filtered_keywords and len(w) > 1]
        freq_counter = Counter(filtered_words)
        
        st.subheader("關鍵詞出現頻率")
        st.bar_chart(pd.DataFrame(freq_counter.most_common(20), columns=["關鍵詞", "出現次數"]).set_index("關鍵詞"))
    except Exception as e:
        st.error(f"Error processing document list: {e}")
