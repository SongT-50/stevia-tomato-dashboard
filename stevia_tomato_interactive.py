import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# 1. 데이터 불러오기
file = '주문이력목록_20250524.xlsx'
df = pd.read_excel(file)

# 2. 확정 의미의 상태만 필터링 ('구매확정', '확정', '운송', '운송완료')
확정상태 = ['구매확정', '확정', '운송', '운송완료']
filtered = df[df['진행상태'].isin(확정상태)].copy()

# 3. 날짜 컬럼 전처리
filtered['주문일자'] = pd.to_datetime(filtered['주문일자'], format='%Y%m%d')
filtered['월'] = filtered['주문일자'].dt.strftime('%Y-%m')

# 4. 중도매인 번호 추출
filtered['중도매인번호'] = filtered['구매자'].str.extract(r'(\d+)').astype(int)

# 5. 월별 주문건수 (모든 월 포함, 0건도 표시)
all_months = pd.date_range('2024-12-01', '2025-05-01', freq='MS').strftime('%Y-%m')
monthly = filtered.groupby('월').size().reindex(all_months, fill_value=0).reset_index()
monthly.columns = ['월', '주문건수']

# 공통 폰트 설정
font_big = dict(family='Malgun Gothic', size=22)
title_font_size = 32
axis_title_font_size = 26
legend_font_size = 20

# 6. 월별 주문건수 그래프 (annotation: x축 첫번째 막대 위 여백, 겹치지 않게)
fig2 = px.bar(
    monthly,
    x='월',
    y='주문건수',
    title='월별 중도매인 주문건수 (확정상태+운송 포함)',
    labels={'월': '월', '주문건수': '주문건수'},
    color='주문건수',
    color_continuous_scale='Blues'
)
fig2.update_layout(
    title_x=0.5,
    font=font_big,
    title_font_size=title_font_size,
    xaxis_title_font_size=axis_title_font_size,
    yaxis_title_font_size=axis_title_font_size,
    legend_font_size=legend_font_size
)
annotation_text = (
    '2024년: 760kg / 6,650,000원<br>'
    '2025년: 12,756kg / 93,069,500원'
)
# x축 첫번째 값 위쪽에 annotation (겹치지 않게 y=1.13)
fig2.add_annotation(
    x=all_months[0], y=1.13, xref='x', yref='paper',
    text=annotation_text,
    showarrow=False,
    align='left',
    font=font_big,
    bgcolor='rgba(255,255,255,0.7)',
    bordercolor='black',
    borderwidth=1
)
fig2.write_html('monthly_trend_interactive_confirmed.html')

# 7. 일별 주문건수 추이
# (확정상태+운송 포함)
daily = filtered.groupby('주문일자').size().reset_index(name='주문건수')
fig1 = px.line(
    daily,
    x='주문일자',
    y='주문건수',
    title='일별 중도매인 주문건수 추이 (확정상태+운송 포함)',
    markers=True,
    labels={'주문일자': '날짜', '주문건수': '주문건수'}
)
fig1.update_traces(line=dict(color='royalblue', width=3), marker=dict(size=10, color='royalblue'))
fig1.update_layout(
    title_x=0.5,
    font=font_big,
    title_font_size=title_font_size,
    xaxis_title_font_size=axis_title_font_size,
    yaxis_title_font_size=axis_title_font_size,
    legend_font_size=legend_font_size
)
fig1.write_html('daily_trend_interactive_confirmed.html')

# 8. 중도매인별 월별 주문건수 히트맵
df_heatmap = filtered.groupby(['월', '중도매인번호']).size().reset_index(name='주문건수')
fig3 = px.density_heatmap(
    df_heatmap,
    x='월',
    y='중도매인번호',
    z='주문건수',
    title='중도매인별 월별 주문건수 히트맵 (확정상태+운송 포함)',
    labels={'월': '월', '중도매인번호': '중도매인 번호', '주문건수': '주문건수'},
    color_continuous_scale='YlOrRd'
)
fig3.update_layout(
    title_x=0.5,
    font=font_big,
    title_font_size=title_font_size,
    xaxis_title_font_size=axis_title_font_size,
    yaxis_title_font_size=axis_title_font_size,
    legend_font_size=legend_font_size
)
fig3.write_html('merchant_monthly_heatmap_confirmed.html')

# 9. 중도매인별 월별 주문건수 변화 (라인 그래프)
fig4 = px.line(
    df_heatmap,
    x='월',
    y='주문건수',
    color='중도매인번호',
    title='중도매인별 월별 주문건수 변화 (확정상태+운송 포함)',
    labels={'월': '월', '주문건수': '주문건수', '중도매인번호': '중도매인 번호'}
)
fig4.update_layout(
    title_x=0.5,
    font=font_big,
    title_font_size=title_font_size,
    xaxis_title_font_size=axis_title_font_size,
    yaxis_title_font_size=axis_title_font_size,
    legend_font_size=legend_font_size
)
fig4.write_html('merchant_monthly_line_confirmed.html')

# 10. 월별-중도매인별 주문건수 피벗테이블 (HTML로 저장)
pivot = filtered.pivot_table(index='월', columns='중도매인번호', values='주문번호', aggfunc='count', fill_value=0)
pivot.to_html('pivot_month_merchant.html', encoding='utf-8')

# 11. 2025년 5월 한 달만 중도매인별 주문건수 막대그래프
may = filtered[filtered['월'] == '2025-05']
may_merchant = may.groupby('중도매인번호').size().reset_index(name='주문건수')
fig_may = px.bar(
    may_merchant,
    x='중도매인번호',
    y='주문건수',
    title='2025년 5월 중도매인별 주문건수 (확정상태+운송 포함)',
    labels={'중도매인번호': '중도매인 번호', '주문건수': '주문건수'},
    color='주문건수',
    color_continuous_scale='Blues'
)
fig_may.update_layout(
    title_x=0.5,
    font=font_big,
    title_font_size=title_font_size,
    xaxis_title_font_size=axis_title_font_size,
    yaxis_title_font_size=axis_title_font_size,
    legend_font_size=legend_font_size
)
fig_may.write_html('may_merchant_bar_confirmed.html')

# 12. index.html 생성 (모든 결과물 링크)
with open('index.html', 'w', encoding='utf-8') as f:
    f.write('<!DOCTYPE html>\n<html lang="ko">\n<head>\n  <meta charset="UTF-8">\n  <title>스테비아 토마토 거래 분석 대시보드</title>\n  <meta name="viewport" content="width=device-width, initial-scale=1">\n  <style>\n    body { font-family: \'Malgun Gothic\', sans-serif; margin: 0; padding: 0; background: #f8f8f8; }\n    h2 { text-align: center; margin-top: 24px; font-size: 2em; }\n    ul { list-style: none; padding: 0; max-width: 500px; margin: 24px auto; }\n    li { margin: 16px 0; }\n    a {\n      display: block;\n      background: #1976d2;\n      color: #fff;\n      text-decoration: none;\n      padding: 22px 0;\n      border-radius: 8px;\n      font-size: 1.4em;\n      text-align: center;\n      box-shadow: 0 2px 8px rgba(0,0,0,0.07);\n      transition: background 0.2s;\n    }\n    a:hover { background: #1565c0; }\n    @media (max-width: 600px) {\n      h2 { font-size: 1.3em; }\n      a { font-size: 1.1em; padding: 18px 0; }\n    }\n  </style>\n</head>\n<body>\n  <h2>스테비아 토마토 거래 분석 대시보드</h2>\n  <ul>\n    <li><a href="monthly_trend_interactive_confirmed.html" target="_blank">월별 중도매인 주문건수 (확정상태+운송 포함)</a></li>\n    <li><a href="daily_trend_interactive_confirmed.html" target="_blank">일별 중도매인 주문건수 추이 (확정상태+운송 포함)</a></li>\n    <li><a href="merchant_monthly_heatmap_confirmed.html" target="_blank">중도매인별 월별 주문건수 히트맵 (확정상태+운송 포함)</a></li>\n    <li><a href="merchant_monthly_line_confirmed.html" target="_blank">중도매인별 월별 주문건수 변화 (확정상태+운송 포함)</a></li>\n    <li><a href="pivot_month_merchant.html" target="_blank">월별-중도매인별 주문건수 표</a></li>\n    <li><a href="may_merchant_bar_confirmed.html" target="_blank">2025년 5월 중도매인별 주문건수</a></li>\n  </ul>\n</body>\n</html>')

print('모든 그래프와 index.html을 "주문건수" 용어로 통일하고, annotation 위치도 조정 완료!') 