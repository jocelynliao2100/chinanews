import streamlit as st
from docx import Document
from collections import defaultdict, Counter
import matplotlib.pyplot as plt
import matplotlib as mpl
import re, io, os
from datetime import datetime
import pandas as pd
import jieba
import jieba.analyse
import requests

# 嘗試載入 BeautifulSoup，不影響字體
try:
    from bs4 import BeautifulSoup
except ImportError:
    BeautifulSoup = None

# --- 1. Font Setup (Matplotlib) ---
LOCAL_FONT_PATH = "NotoSansCJKjp-Regular.otf"
if os.path.exists(LOCAL_FONT_PATH):
    # 1.1 加入到 Matplotlib 的 fontManager
    mpl.font_manager.fontManager.addfont(LOCAL_FONT_PATH)
    # 1.2 設定為全域字體
    mpl.rcParams["font.family"] = mpl.font_manager.FontProperties(fname=LOCAL_FONT_PATH).get_name()
    mpl.rcParams["axes.unicode_minus"] = False
else:
    st.warning(f"找不到字體檔 {LOCAL_FONT_PATH}，中文可能顯示不正常。")

# --- 2. 工具函式 ---
def process_uploaded_files(uploaded_files, column_names):
    date_re = re.compile(
        r"\b(202[0-5])[-/年\.](0?[1-9]|1[0-2])[-/月\.](0?[1-9]|[12]\d|3[01])日?\b"
    )
    counts = defaultdict(lambda: defaultdict(int))

    for name, uploaded in zip(column_names, uploaded_files):
        data = uploaded.getvalue()
        doc = Document(io.BytesIO(data))
        for p in doc.paragraphs:
            m = date_re.search(p.text)
            if not m:
                continue
            y, mth, _ = m.groups()
            ym = f"{y}-{int(mth):02d}"
            dt = datetime(int(y), int(mth), 1)
            if datetime(2020,1,1) <= dt <= datetime(2025,4,30):
                counts[name][ym] += 1

    # 轉成一般 dict
    return {k: dict(v) for k, v in counts.items()}

def plot_data(counts, original_names, display_names):
    # 全部月份
    all_m = sorted({ym for d in counts.values() for ym in d})
    # 轉 datetime
    dates = [datetime.strptime(ym, "%Y-%m") for ym in all_m]

    fig, ax = plt.subplots(figsize=(10,5))
    for orig, disp in zip(original_names, display_names):
        y = [counts.get(orig,{}).get(ym, 0) for ym in all_m]
        ax.plot(dates, y, marker="o", label=disp)

    ax.set_title("2020.01–2025.04 國台辦各欄目新聞稿數量變化")
    ax.set_xlabel("年月")
    ax.set_ylabel("新聞稿數量")
    ax.legend()
    ax.grid(True)
    plt.xticks(rotation=45)
    plt.tight_layout()
    st.pyplot(fig)

# --- 3. Streamlit App ---

st.title("🇨🇳 國台辦新聞稿分析")

st.header("一、新聞稿數量變化")
uploaded = st.file_uploader(
    "請一次上傳 5 個 .docx（台辦動態、交流交往、政務要聞、部門涉台、新聞發佈）",
    type="docx", accept_multiple_files=True, key="uploader1"
)
display_names = ["台辦動態", "交流交往", "政務要聞", "部門涉台", "新聞發佈"]

if uploaded:
    if len(uploaded) != 5:
        st.warning(f"請上傳 5 個檔，目前：{len(uploaded)}")
    else:
        originals = [f.name for f in uploaded]
        with st.spinner("🔄 資料處理中..."):
            cnts = process_uploaded_files(uploaded, originals)
        plot_data(cnts, originals, display_names)

        # 額外：找出最多月份 & 該月前20關鍵字
        total = Counter()
        for d in cnts.values():
            total.update(d)
        if total:
            month, num = total.most_common(1)[0]
            st.metric("🎯 新聞稿最多月份", month, f"{num} 篇")

            # 蒐集該月所有段落
            segs = []
            date_re = re.compile(
                r"\b(202[0-5])[-/年\.](0?[1-9]|1[0-2])[-/月\.](0?[1-9]|[12]\d|3[01])日?\b"
            )
            for up in uploaded:
                doc = Document(io.BytesIO(up.getvalue()))
                for p in doc.paragraphs:
                    m = date_re.search(p.text)
                    if m and f"{m.group(1)}-{int(m.group(2)):02d}"==month:
                        segs.append(p.text)

            words = [w for seg in segs for w in jieba.lcut(seg) if len(w)>1]
            top20 = Counter(words).most_common(20)
            st.subheader(f"📑 {month} 前20關鍵詞")
            cols = st.columns(2)
            for i, (w,c) in enumerate(top20):
                cols[i%2].write(f"{i+1}. {w}：{c}")

st.markdown("---")
st.header("二、單篇關鍵詞分析")
single = st.file_uploader("上傳單篇新聞稿 (.docx)", type="docx", key="uploader2")
if single:
    data = single.getvalue()
    doc = Document(io.BytesIO(data))
    txt = "\n".join(p.text for p in doc.paragraphs)
    tags = jieba.analyse.extract_tags(txt, topK=30, withWeight=True)
    st.subheader("前30關鍵詞(含權重)")
    for w,wt in tags:
        st.write(f"{w}：{wt:.3f}")

st.markdown("---")
st.header("三、新聞列表解析與全文爬取")
list_up = st.file_uploader("上傳含列表的 Word 檔", type="docx", key="uploader3")
def parse_list_docx(buf):
    items=[]
    doc=Document(io.BytesIO(buf))
    for line in "\n".join(p.text for p in doc.paragraphs).splitlines():
        m=re.search(r"\[(\d{4}-\d{1,2}-\d{1,2})\]\s*(.*?)\s*(https?://\S+)?", line)
        if m:
            d,t,u=m.groups()
            try:
                dt=datetime.strptime(d, "%Y-%m-%d")
                if datetime(2020,1,1)<=dt<=datetime(2025,4,30):
                    items.append({"date":dt, "title":t or "", "url":u or ""})
            except: pass
    return items

def fetch_text(url):
    if not BeautifulSoup: return ""
    try:
        r=requests.get(url,timeout=5); r.encoding=r.apparent_encoding
        s=BeautifulSoup(r.text,"html.parser").find("div",class_="TRS_Editor")
        return s.get_text("\n") if s else ""
    except: return ""

if list_up:
    buf=list_up.getvalue()
    lst=parse_list_docx(buf)
    if lst:
        st.success(f"解析到 {len(lst)} 篇新聞")
        texts=[]
        with st.spinner("爬取全文中..."):
            for it in lst:
                if it["url"].startswith("http"):
                    texts.append(fetch_text(it["url"]))
        st.success(f"成功爬取 {len([t for t in texts if t])} 篇")
    else:
        st.warning("無法解析任何新聞列表")
