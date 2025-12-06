"""
스테비아 토마토 거래 분석 대시보드
=====================================

사용법:
1. 아래 [설정] 부분에서 데이터 파일 경로를 변경하세요
2. 터미널에서 실행: python stevia_tomato_interactive.py
3. 생성된 index.html을 브라우저에서 열어보세요

새 데이터 업로드 시:
- DATA_FILE 경로만 새 파일로 변경하면 됩니다
- 날짜 범위는 자동으로 계산됩니다
"""

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime

# ============================================================
# [설정] - 여기만 수정하세요!
# ============================================================

# 데이터 파일 경로 (새 파일 업로드 시 이 경로만 변경)
DATA_FILE = '주문이력목록_20250524.xlsx'

# 확정으로 인정할 상태들
CONFIRMED_STATUS = ['구매확정', '확정', '운송', '운송완료']

# 연도별 요약 정보 (수동 입력 - 필요시 업데이트)
YEARLY_SUMMARY = {
    '2024': {'kg': 760, 'won': 6650000},
    '2025': {'kg': 12756, 'won': 93069500},
}

# 출력 폴더 (현재 폴더에 저장)
OUTPUT_DIR = './'

# ============================================================
# [스타일 설정]
# ============================================================

FONT_SETTINGS = dict(family='Malgun Gothic', size=22)
TITLE_FONT_SIZE = 32
AXIS_TITLE_FONT_SIZE = 26
LEGEND_FONT_SIZE = 20

# ============================================================
# [메인 코드] - 수정 불필요
# ============================================================

def load_and_filter_data(file_path):
    """데이터 로드 및 필터링"""
    print(f"📂 데이터 로드 중: {file_path}")
    df = pd.read_excel(file_path)

    # 확정 상태만 필터링
    filtered = df[df['진행상태'].isin(CONFIRMED_STATUS)].copy()
    print(f"✅ 확정 상태 필터링 완료: {len(filtered)}건")

    # 날짜 전처리
    filtered['주문일자'] = pd.to_datetime(filtered['주문일자'], format='%Y%m%d')
    filtered['월'] = filtered['주문일자'].dt.strftime('%Y-%m')

    # 중도매인 번호 추출
    filtered['중도매인번호'] = filtered['구매자'].str.extract(r'(\d+)').astype(int)

    return filtered


def get_month_range(filtered_df):
    """데이터에서 월 범위 자동 계산"""
    min_date = filtered_df['주문일자'].min().replace(day=1)
    max_date = filtered_df['주문일자'].max().replace(day=1)
    return pd.date_range(min_date, max_date, freq='MS').strftime('%Y-%m')


def apply_common_layout(fig):
    """공통 레이아웃 적용"""
    fig.update_layout(
        title_x=0.5,
        font=FONT_SETTINGS,
        title_font_size=TITLE_FONT_SIZE,
        xaxis_title_font_size=AXIS_TITLE_FONT_SIZE,
        yaxis_title_font_size=AXIS_TITLE_FONT_SIZE,
        legend_font_size=LEGEND_FONT_SIZE
    )
    return fig


def save_html(fig, filename):
    """CDN 방식으로 HTML 저장 (파일 크기 대폭 감소)"""
    filepath = f"{OUTPUT_DIR}{filename}"
    fig.write_html(filepath, include_plotlyjs='cdn')
    print(f"💾 저장 완료: {filename}")


def create_monthly_trend(filtered, all_months):
    """월별 주문건수 막대 그래프"""
    monthly = filtered.groupby('월').size().reindex(all_months, fill_value=0).reset_index()
    monthly.columns = ['월', '주문건수']

    fig = px.bar(
        monthly, x='월', y='주문건수',
        title='월별 중도매인 주문건수 (확정상태+운송 포함)',
        labels={'월': '월', '주문건수': '주문건수'},
        color='주문건수',
        color_continuous_scale='Blues'
    )

    # 연도별 요약 annotation
    summary_lines = []
    for year, data in sorted(YEARLY_SUMMARY.items()):
        summary_lines.append(f"{year}년: {data['kg']:,}kg / {data['won']:,}원")
    annotation_text = '<br>'.join(summary_lines)

    fig.add_annotation(
        x=all_months[0], y=1.13, xref='x', yref='paper',
        text=annotation_text,
        showarrow=False, align='left',
        font=FONT_SETTINGS,
        bgcolor='rgba(255,255,255,0.7)',
        bordercolor='black', borderwidth=1
    )

    apply_common_layout(fig)
    save_html(fig, 'monthly_trend_interactive_confirmed.html')


def create_daily_trend(filtered):
    """일별 주문건수 라인 그래프"""
    daily = filtered.groupby('주문일자').size().reset_index(name='주문건수')

    fig = px.line(
        daily, x='주문일자', y='주문건수',
        title='일별 중도매인 주문건수 추이 (확정상태+운송 포함)',
        markers=True,
        labels={'주문일자': '날짜', '주문건수': '주문건수'}
    )
    fig.update_traces(
        line=dict(color='royalblue', width=3),
        marker=dict(size=10, color='royalblue')
    )

    apply_common_layout(fig)
    save_html(fig, 'daily_trend_interactive_confirmed.html')


def create_merchant_heatmap(filtered):
    """중도매인별 월별 히트맵"""
    df_heatmap = filtered.groupby(['월', '중도매인번호']).size().reset_index(name='주문건수')

    fig = px.density_heatmap(
        df_heatmap, x='월', y='중도매인번호', z='주문건수',
        title='중도매인별 월별 주문건수 히트맵 (확정상태+운송 포함)',
        labels={'월': '월', '중도매인번호': '중도매인 번호', '주문건수': '주문건수'},
        color_continuous_scale='YlOrRd'
    )

    apply_common_layout(fig)
    save_html(fig, 'merchant_monthly_heatmap_confirmed.html')

    return df_heatmap


def create_merchant_line(df_heatmap):
    """중도매인별 월별 라인 그래프"""
    fig = px.line(
        df_heatmap, x='월', y='주문건수', color='중도매인번호',
        title='중도매인별 월별 주문건수 변화 (확정상태+운송 포함)',
        labels={'월': '월', '주문건수': '주문건수', '중도매인번호': '중도매인 번호'}
    )

    apply_common_layout(fig)
    save_html(fig, 'merchant_monthly_line_confirmed.html')


def create_pivot_table(filtered):
    """월별-중도매인별 피벗테이블"""
    pivot = filtered.pivot_table(
        index='월', columns='중도매인번호',
        values='주문번호', aggfunc='count', fill_value=0
    )

    filepath = f"{OUTPUT_DIR}pivot_month_merchant.html"
    pivot.to_html(filepath, encoding='utf-8')
    print(f"💾 저장 완료: pivot_month_merchant.html")


def create_latest_month_bar(filtered):
    """가장 최근 월의 중도매인별 막대 그래프"""
    latest_month = filtered['월'].max()
    latest_data = filtered[filtered['월'] == latest_month]
    merchant_data = latest_data.groupby('중도매인번호').size().reset_index(name='주문건수')

    fig = px.bar(
        merchant_data, x='중도매인번호', y='주문건수',
        title=f'{latest_month} 중도매인별 주문건수 (확정상태+운송 포함)',
        labels={'중도매인번호': '중도매인 번호', '주문건수': '주문건수'},
        color='주문건수',
        color_continuous_scale='Blues'
    )

    apply_common_layout(fig)
    save_html(fig, 'latest_month_merchant_bar.html')

    return latest_month


def create_index_html(latest_month):
    """메인 대시보드 페이지 생성"""
    today = datetime.now().strftime('%Y-%m-%d')

    html_content = f'''<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8">
  <title>스테비아 토마토 거래 분석 대시보드</title>
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <style>
    body {{
      font-family: 'Malgun Gothic', sans-serif;
      margin: 0;
      padding: 20px;
      background: #f8f8f8;
    }}
    .container {{
      max-width: 600px;
      margin: 0 auto;
    }}
    h1 {{
      text-align: center;
      color: #333;
      font-size: 1.8em;
      margin-bottom: 10px;
    }}
    .update-info {{
      text-align: center;
      color: #666;
      font-size: 0.9em;
      margin-bottom: 30px;
    }}
    .card {{
      background: white;
      border-radius: 12px;
      padding: 20px;
      margin-bottom: 16px;
      box-shadow: 0 2px 8px rgba(0,0,0,0.08);
    }}
    .card h3 {{
      margin: 0 0 12px 0;
      color: #333;
      font-size: 1.1em;
    }}
    .card a {{
      display: block;
      background: #1976d2;
      color: white;
      text-decoration: none;
      padding: 14px 20px;
      border-radius: 8px;
      text-align: center;
      font-size: 1em;
      transition: background 0.2s;
    }}
    .card a:hover {{
      background: #1565c0;
    }}
    .section-title {{
      font-size: 1em;
      color: #666;
      margin: 24px 0 12px 0;
      padding-bottom: 8px;
      border-bottom: 2px solid #e0e0e0;
    }}
    @media (max-width: 600px) {{
      h1 {{ font-size: 1.4em; }}
      .card a {{ padding: 12px 16px; }}
    }}
  </style>
</head>
<body>
  <div class="container">
    <h1>스테비아 토마토 거래 분석</h1>
    <p class="update-info">마지막 업데이트: {today}</p>

    <div class="section-title">📊 전체 추이</div>

    <div class="card">
      <h3>월별 주문건수</h3>
      <a href="monthly_trend_interactive_confirmed.html">차트 보기</a>
    </div>

    <div class="card">
      <h3>일별 주문건수 추이</h3>
      <a href="daily_trend_interactive_confirmed.html">차트 보기</a>
    </div>

    <div class="section-title">👥 중도매인 분석</div>

    <div class="card">
      <h3>중도매인별 월별 히트맵</h3>
      <a href="merchant_monthly_heatmap_confirmed.html">차트 보기</a>
    </div>

    <div class="card">
      <h3>중도매인별 월별 추이</h3>
      <a href="merchant_monthly_line_confirmed.html">차트 보기</a>
    </div>

    <div class="card">
      <h3>월별-중도매인별 표</h3>
      <a href="pivot_month_merchant.html">표 보기</a>
    </div>

    <div class="section-title">📅 최근 월 ({latest_month})</div>

    <div class="card">
      <h3>{latest_month} 중도매인별 주문건수</h3>
      <a href="latest_month_merchant_bar.html">차트 보기</a>
    </div>
  </div>
</body>
</html>'''

    filepath = f"{OUTPUT_DIR}index.html"
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(html_content)
    print(f"💾 저장 완료: index.html")


def main():
    """메인 실행 함수"""
    print("=" * 50)
    print("🍅 스테비아 토마토 거래 분석 시작")
    print("=" * 50)

    # 1. 데이터 로드
    filtered = load_and_filter_data(DATA_FILE)

    # 2. 월 범위 자동 계산
    all_months = get_month_range(filtered)
    print(f"📅 분석 기간: {all_months[0]} ~ {all_months[-1]}")

    # 3. 시각화 생성
    print("\n📊 시각화 생성 중...")
    create_monthly_trend(filtered, all_months)
    create_daily_trend(filtered)
    df_heatmap = create_merchant_heatmap(filtered)
    create_merchant_line(df_heatmap)
    create_pivot_table(filtered)
    latest_month = create_latest_month_bar(filtered)

    # 4. 인덱스 페이지 생성
    create_index_html(latest_month)

    print("\n" + "=" * 50)
    print("✅ 모든 작업 완료!")
    print("📁 index.html을 브라우저에서 열어보세요")
    print("=" * 50)


if __name__ == '__main__':
    main()
