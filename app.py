import streamlit as st
import pandas as pd
import plotly.graph_objects as go

# 頁面標題
st.title("中國新聞熱度地圖")

# 地點和出現次數資料
data = {
    "地点": ["贵州", "成都", "四川", "武汉", "重庆", "广东", "福建", "珠海",
             "广州", "青岛", "南京", "东莞", "山东"],
    "次数": [289, 260, 232, 226, 225, 202, 174, 160, 157, 155, 133, 132, 123]
}

df = pd.DataFrame(data)

# 地點的經緯度 (省級取省會)
location_data = {
    "贵州": {"lon": 106.8748, "lat": 26.8154},  # 貴陽
    "成都": {"lon": 104.0665, "lat": 30.5726},
    "四川": {"lon": 102.0098, "lat": 30.6171},
    "武汉": {"lon": 114.2986, "lat": 30.5844},
    "重庆": {"lon": 107.8740, "lat": 30.0572},
    "广东": {"lon": 113.2644, "lat": 23.1291},
    "福建": {"lon": 118.1130, "lat": 26.0789},
    "珠海": {"lon": 113.5668, "lat": 22.2707},
    "广州": {"lon": 113.2644, "lat": 23.1291},
    "青岛": {"lon": 120.3849, "lat": 36.0671},
    "南京": {"lon": 118.7969, "lat": 32.0617},
    "东莞": {"lon": 113.7500, "lat": 23.0408},
    "山东": {"lon": 118.1654, "lat": 36.6815}  # 濟南
}

# 加入經緯度
df["lon"] = df["地点"].map(lambda x: location_data[x]["lon"])
df["lat"] = df["地点"].map(lambda x: location_data[x]["lat"])

# 繪製地圖
fig = go.Figure(data=go.Scattergeo(
    lon=df['lon'],
    lat=df['lat'],
    text=df['地点'] + ": " + df['次数'].astype(str) + " 次",
    mode='markers',
    marker=dict(
        size=df['次数'] / 10,
        sizemode='diameter',
        color=df['次数'],
        colorscale='Reds',
        opacity=0.8,
        colorbar=dict(title="出现次数")
    ),
))

fig.update_layout(
    title_text='中國地點新聞熱度圖',
    geo=dict(
        scope='asia',
        showland=True,
        landcolor="LightGreen",
        showocean=True,
        oceancolor="LightBlue",
        projection_type='natural earth'
    ),
)

# 顯示地圖
st.plotly_chart(fig)
