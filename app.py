# streamlit_app.py

import streamlit as st
from docx import Document
import io, re
from collections import defaultdict
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib as mpl

# --- 加載專案中同目錄的中文字體檔 ---
# 請先下載並放置 NotoSansCJKtc-Regular.otf 到本檔同一資料夾
mpl.font_manager.fontManager.addfont("NotoSansCJKtc-Regular.otf")
mpl.rcParams['font.family'] = 'Noto Sans CJK TC'
mpl.rcParams['axes.unicode_minus'] = False

st.title("國台辦新聞稿各欄目月度統計 (2020–2025.04)")

st.markdown("""
請一次性上傳**五個**對應以下欄目的 `.docx` 檔，系統將自動解析並統計：
- 台辦動態  
- 交流交往  
- 政務要聞  
- 部門涉台  
- 要聞 (新聞發佈)  
""")

uploaded = st.file_uploader(
    "上傳 5 個 .docx 檔案", 
    type="docx", 
    accept_multiple_files=True
)

if uploaded:
    if len(uploaded) != 5:
        st.warning(f"目前上傳了 {len(uploaded)} 個檔案，請上傳 5 個。")
    else:
        # 檔名關鍵字 → 欄目對應
        category_map = {
            "動態": "台辦動態",
            "交流交往": "交流交往",
            "政務要聞": "政務要聞",
            "部門涉台": "部門涉台",
            "要聞": "新聞發佈",
            "新闻要闻": "新聞發佈",
        }
        # 日期擷取正則
        date_pattern = re.compile(
            r"\b(202[0-5])[-/年\.]"
            r"(0?[1-9]|1[0-2])[-/月\.]"
            r"(0?[1-9]|[12]\d|3[01])日?\b"
        )

        # 統計資料：欄目 → { "YYYY-MM": count }
        counts = defaultdict(lambda: defaultdict(int))
        # 同時保留 (datetime, 文字) 用於關鍵字分析
        text_date_list = []

        for file in uploaded:
            fname = file.name
            # 自動判別欄目
            category = next(
                (v for k, v in category_map.items() if k in fname),
                fname
            )

            doc = Document(io.BytesIO(file.getvalue()))
            for para in doc.paragraphs:
                m = date_pattern.search(para.text)
                if m:
                    y, m_, d = m.groups()
                    ym = f"{y}-{int(m_):02d}"
                    dt = pd.to_datetime(f"{y}-{int(m_):02d}-{int(d):02d}")
                    # 僅限 2020-01 ~ 2025-04
                    if "2020-01" <= ym <= "2025-04":
                        counts[category][ym] += 1
                        text_date_list.append((dt, para.text))

        # 完整月份索引
        all_months = pd.date_range(
            "2020-01-01", "2025-04-01", freq="MS"
        ).strftime("%Y-%m")
        # 建立 DataFrame
        df = pd.DataFrame(index=all_months, columns=counts.keys()).fillna(0)
        for cat, data in counts.items():
            for ym, c in data.items():
                if ym in df.index:
                    df.at[ym, cat] = c
        df.index = pd.to_datetime(df.index + "-01")

        # 折線圖
        st.subheader("折線圖")
        fig, ax = plt.subplots(figsize=(10, 5))
        for col in df.columns:
            ax.plot(df.index, df[col], marker='o', label=col)
        ax.set_xlabel("年月")
        ax.set_ylabel("新聞稿數量")
        ax.set_title("國台辦各欄目新聞稿數量變化 (2020.01–2025.04)")
        ax.legend()
        ax.grid(True)
        plt.xticks(rotation=45)
        st.pyplot(fig)

        # 2. 新聞稿最多的月份
        total_per_month = df.sum(axis=1)
        most_month = total_per_month.idxmax().strftime("%Y-%m")
        st.metric("🎯 新聞稿最多的月份", most_month)

        # 3. 該月份前20大中文關鍵字
        texts_in_most = [
            txt for dt, txt in text_date_list
            if dt.strftime("%Y-%m") == most_month
        ]
        import jieba
        from collections import Counter
        all_words = []
        for txt in texts_in_most:
            all_words += [w for w in jieba.lcut(txt) if len(w) > 1]
        top20 = Counter(all_words).most_common(20)

        st.write(f"#### 📑 {most_month} 前20大中文關鍵字")
        for word, freq in top20:
            st.write(f"- {word}：{freq} 次")

        # 原始統計表
        st.subheader("原始統計表")
        st.dataframe(df.astype(int), use_container_width=True)
