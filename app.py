# streamlit_app.py

import streamlit as st
from docx import Document
import io, re
from collections import defaultdict
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib

# --- 中文字體設定（視環境而定，可改成你系統內的字體） ---
matplotlib.rcParams['font.sans-serif'] = ['SimHei']
matplotlib.rcParams['axes.unicode_minus'] = False

st.title("國台辦新聞稿各欄目月度統計 (2020–2025.04)")

st.markdown("""
請一次性上傳**五個**對應以下欄目的 `.docx` 檔，系統將自動解析並統計：
- 台辦動態  
- 交流交往  
- 政務要聞  
- 部門涉台  
- 要聞 (新聞發佈)  
""")

uploaded = st.file_uploader("上傳 5 個 .docx 檔案", type="docx", accept_multiple_files=True)

if uploaded:
    if len(uploaded) != 5:
        st.warning(f"目前上傳了 {len(uploaded)} 個檔案，請上傳 5 個。")
    else:
        # 定義各檔案名關鍵字到欄目中文的對應
        category_map = {
            "動態": "台辦動態",
            "交流交往": "交流交往",
            "政務要聞": "政務要聞",
            "部門涉台": "部門涉台",
            "要聞": "新聞發佈",
            "新闻要闻": "新聞發佈",
        }
        # 用於日期擷取的正則
        date_pattern = re.compile(
            r"\b(202[0-5])[-/年\.](0?[1-9]|1[0-2])[-/月\.](0?[1-9]|[12]\d|3[01])日?\b"
        )

        # 統計資料結構：欄目 → { "YYYY-MM": count }
        counts = defaultdict(lambda: defaultdict(int))

        for file in uploaded:
            # 嘗試自動判別欄目
            fname = file.name
            category = None
            for key, val in category_map.items():
                if key in fname:
                    category = val
                    break
            if category is None:
                category = fname  # fallback：直接用檔名

            # 讀取 .docx
            doc = Document(io.BytesIO(file.getvalue()))
            for para in doc.paragraphs:
                m = date_pattern.search(para.text)
                if m:
                    y, m_, d = m.groups()
                    ym = f"{y}-{int(m_):02d}"
                    # 只統計 2020-01 ~ 2025-04
                    if "2020-01" <= ym <= "2025-04":
                        counts[category][ym] += 1

        # 建立完整月份索引
        all_months = pd.date_range("2020-01-01", "2025-04-01", freq="MS").strftime("%Y-%m")
        # 建立 DataFrame
        df = pd.DataFrame(index=all_months, columns=counts.keys()).fillna(0)
        for cat, data in counts.items():
            for ym, c in data.items():
                if ym in df.index:
                    df.at[ym, cat] = c
        df.index = pd.to_datetime(df.index + "-01")

        st.subheader("折線圖")
        # 用 matplotlib 繪圖並顯示
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

        st.subheader("原始統計表")
        st.dataframe(df.astype(int), use_container_width=True)
