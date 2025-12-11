# Streamlit app for 시각화: 코오롱인더스트리 DX아카데미 교육 만족도
# 작성: 자동생성 (프롬프트 엔지니어링 기반)
# 필요 라이브러리: streamlit, pandas, plotly.express, collections, re
# 사용법: streamlit run app.py
# 주의: 워드클라우드 생성을 위해 'wordcloud' 및 'matplotlib' 라이브러리 설치 필요 (현재 워드 클라우드 코드 삭제됨)

import streamlit as st
import pandas as pd
import plotly.express as px
from collections import Counter
import re
from typing import List, Dict

# 워드클라우드 및 Matplotlib 관련 라이브러리 제거 (사용하지 않음)
# import matplotlib.pyplot as plt
# from wordcloud import WordCloud
# FONT_PATH 관련 로직 제거

# ---------------------------
# 설정 및 상수
# ---------------------------
DATA_PATH = "(DX 아카데미) 교육 만족도_Track1_5차_재무회계원가구매(마곡)_20251202화(누적)_평균제외streamlit.xlsx"

# 객관식(리커트) 컬럼 목록 (요청대로)
LIKERT_COLS = [
    '해당 교육과정의 내용과 흐름에 전반적으로 만족하십니까?',
    '교육 난이도는 본인의 수준에 적절했나요?',
    'ChatGPT를 활용하여 분석을 수행하는 것이 도움이 된다고 생각하십니까?',
    '참여한 교육 과정을 주변에 추천할 의향이 있으십니까?',
    '교육 일자 및 시간 배분이 적절했나요? (전체교육시간, 실습 비중)',
    '제공된 교육 과정 및 교안은 실무에서도 참고할 수 있을 만큼 실용적인가요?',
    '실습 환경 (RAS, Python) 은 데이터 분석 흐름을 체험하기에 적절하였나요?',
    '강사들은 교육 내용을 명확하고 이해하기 쉽게 전달했다',
    '질문이나 실습 요청사항에 도움을 주고, 수강생과 원활하게 소통했다'
]

# 히트맵 X축 가독성 개선을 위한 문항 단축/줄바꿈 목록
LIKERT_COLS_SHORT = {
    '해당 교육과정의 내용과 흐름에 전반적으로 만족하십니까?': '전반적 만족도',
    '교육 난이도는 본인의 수준에 적절했나요?': '교육\n난이도\n적절성',
    'ChatGPT를 활용하여 분석을 수행하는 것이 도움이 된다고 생각하십니까?': 'ChatGPT\n활용\n도움 여부',
    '참여한 교육 과정을 주변에 추천할 의향이 있으십니까?': '주변\n추천\n의향',
    '교육 일자 및 시간 배분이 적절했나요? (전체교육시간, 실습 비중)': '일자/시간\n배분\n적절성',
    '제공된 교육 과정 및 교안은 실무에서도 참고할 수 있을 만큼 실용적인가요?': '교안\n실무\n실용성',
    '실습 환경 (RAS, Python) 은 데이터 분석 흐름을 체험하기에 적절하였나요?': '실습 환경\n적절성',
    '강사들은 교육 내용을 명확하고 이해하기 쉽게 전달했다': '강사\n내용\n전달력',
    '질문이나 실습 요청사항에 도움을 주고, 수강생과 원활하게 소통했다': '강사\n소통/지원\n만족도' # '소통/피드백'에서 '소통/지원'으로 수정됨
}

# 주관식 컬럼 목록
OPEN_COLS = [
    '추가로 배우고 싶은 심화 주제나 필요 역량이 있다면 무엇인가요?',
    '실습이나 자료와 관련하여 좋았던 점이나 개선이 필요한 점이 있다면 자유롭게 적어주세요. (선택)',
    '강의나 강사와 관련하여 인상 깊었던 점이나 개선이 필요한 점이 있다면 자유롭게 적어주세요. (선택)',
    '종합적인 교육 소감 및 정규화 교육을 위해 보완하거나 개선해야할 사항을 조언해주세요'
]

TIMESTAMP_COL = '타임스탬프'
COURSE_COL = '수강 교육과정'
HQ_COL = '소속본부'
DEPT_COL = '소속부서'
PREV_COURSE_COL = '본 강의 이전에 수강한 선행 교육과정을 모두 선택해 주세요.'

# Korean stopwords (간단한 기본 목록 — 필요시 확장 가능)
KOR_STOPWORDS = set([
    '입니다', '있습니다', '습니다', '있어요', '같습니다', '같아요', '그리고',
    '하지만', '또한', '관련', '부분', '좋습니다', '좋아요', '필요', '필요합니다',
    '합니다', '하는', '했습니다', '했습니다.', '하는것', '것', '수', '해주셨으면',
    '했습니다', '많이', '좀', '더', '없이', '있음', '없음', '있다', '없다'
])

# 리커트 문자열 -> 숫자 매핑 (한국어 표현 및 영어 표현 포함)
LIKERT_MAP = {
    '매우 그렇다': 5, '매우 그렇습니다': 5, '매우 그렇습니다.':5,
    '그렇다': 4, '그렇습니다': 4,
    '보통': 3, '중간': 3, '보통입니다':3,
    '그렇지 않다': 2, '그렇지 않습니다':2,
    '전혀 그렇지 않다': 1, '전혀 그렇지 않습니다':1,
    '매우 만족':5, '만족':4, '보통':3, '불만족':2, '매우 불만족':1,
    # 영어/숫자 style
    '5':5, '4':4, '3':3, '2':2, '1':1, '0':0
}

# ---------------------------
# 유틸리티 함수
# ---------------------------

@st.cache_data
def load_data(path: str) -> pd.DataFrame:
    """
    데이터 로드 (캐시 적용). 파일이 없거나 컬럼이 다른 경우 예외 발생.
    타임스탬프는 날짜만 추출하여 yyyy-mm-dd 형식의 datetime으로 저장.
    """
    try:
        df = pd.read_excel(path)
    except FileNotFoundError:
        raise FileNotFoundError(f"파일을 찾을 수 없습니다: {path}")
    except Exception as e:
        raise RuntimeError(f"파일 로딩 중 오류 발생: {e}")

    # 컬럼 존재 확인 (원본 컬럼 일부가 누락되면 사용자에게 친절하게 안내)
    expected_cols = [TIMESTAMP_COL, COURSE_COL, HQ_COL, DEPT_COL] + LIKERT_COLS + OPEN_COLS + [PREV_COURSE_COL]
    missing = [c for c in expected_cols if c not in df.columns]
    if missing:
        # 일부 컬럼은 실제 데이터에 없을 수 있으므로 경고만 주고 진행
        st.warning(f"다음 예상 컬럼이 데이터에 없습니다(경고): {missing}")

    # 타임스탬프 처리: 날짜만 사용
    if TIMESTAMP_COL in df.columns:
        try:
            # pandas to_datetime, 형식 다양성 허용
            df[TIMESTAMP_COL] = pd.to_datetime(df[TIMESTAMP_COL], errors='coerce')
            # 날짜(YYYY-MM-DD)만 남기기
            df['date'] = df[TIMESTAMP_COL].dt.date
        except Exception:
            # 파싱 실패 시 생성 안함
            df['date'] = pd.NaT

    return df

def map_likert_value(val):
    """
    리커트 값을 숫자로 변환.
    - 이미 숫자면 그대로 사용
    - 문자열이면 매핑 테이블로 시도
    - 실패하면 NaN 반환
    """
    if pd.isna(val):
        return pd.NA
    # 숫자 계열
    try:
        # 문자열로 된 숫자나 실수형 처리
        if isinstance(val, (int, float)):
            return val
        s = str(val).strip()
        # 바로 정수로 변환 가능한 경우
        if s.isdigit():
            return int(s)
        # 매핑에서 찾기 (대소문자 구분 없이)
        if s in LIKERT_MAP:
            return LIKERT_MAP[s]
        # 소문자 처리
        s_lower = s.lower()
        if s_lower in LIKERT_MAP:
            return LIKERT_MAP[s_lower]
        # 한글의 경우 공백 제거 후 재시도
        s_norm = re.sub(r'\s+', '', s)
        if s_norm in LIKERT_MAP:
            return LIKERT_MAP[s_norm]
    except Exception:
        return pd.NA
    return pd.NA

def tokenize_korean(text: str) -> List[str]:
    """
    한글 단어 토큰화: 정규표현식으로 한글 블록만 추출.
    불용어 제거 후 반환.
    """
    if not isinstance(text, str):
        return []
    tokens = re.findall(r"[가-힣]+", text)
    tokens = [t for t in tokens if t and t not in KOR_STOPWORDS and len(t) > 1]  # 길이 2 이상 필터
    return tokens

def top_n_words(series: pd.Series, n: int = 10) -> List[tuple]:
    """
    주어진 텍스트 시리즈에서 상위 n개 단어와 빈도 수 반환.
    """
    all_tokens = []
    for item in series.dropna().astype(str):
        all_tokens.extend(tokenize_korean(item))
    counter = Counter(all_tokens)
    return counter.most_common(n)

def categorize_course(name: str) -> str:
    """
    원본 '수강 교육과정' 텍스트를 공통/Track1/Track2/Track3 등으로 분류하는 보조함수.
    내용에 'Track1', 'Track2' 등 또는 '공통'이 포함되면 해당 카테고리 반환, 아니면 '기타'.
    """
    if not isinstance(name, str):
        return '기타'
    s = name.lower()
    if 'track1' in s or 'track 1' in s or 'track1 교육' in s:
        return 'Track1 교육'
    if 'track2' in s or 'track 2' in s or 'track2 교육' in s:
        return 'Track2 교육'
    if 'track3' in s or 'track 3' in s or 'track3 교육' in s:
        return 'Track3 교육'
    if '공통' in name or 'common' in s:
        return '공통교육'
    # 기타는 원본값 그대로 반환(또는 '기타')
    return name
    
# 워드클라우드 함수는 삭제됨
# def generate_wordcloud(text_series: pd.Series, title: str):
#     ...

# ---------------------------
# 스트림릿 UI 시작
# ---------------------------

st.set_page_config(page_title="DX아카데미 교육 만족도 대시보드", layout="wide")
st.title("코오롱인더스트리 DX아카데미 — 교육 만족도 대시보드")

# 데이터 로드 (예외 처리 포함)
try:
    df = load_data(DATA_PATH)
except FileNotFoundError as fe:
    st.error(str(fe))
    st.stop()
except Exception as e:
    st.error(f"데이터 로드 실패: {e}")
    st.stop()

# 원본 데이터 보기(expander)
with st.expander("원본 데이터 보기 (Raw Data)"):
    st.dataframe(df.head(200))

# 기본 전처리: 리커트 컬럼 숫자화
for col in LIKERT_COLS:
    if col in df.columns:
        df[col + '_num'] = df[col].apply(map_likert_value)

# 교육과정 카테고리 컬럼 추가
if COURSE_COL in df.columns:
    df['course_cat'] = df[COURSE_COL].apply(categorize_course)
else:
    df['course_cat'] = '알수없음'

# 사이드바 필터 UI
st.sidebar.header("필터")
# 날짜(타임스탬프) 목록 준비 (date 컬럼 사용)
if 'date' in df.columns:
    unique_dates = sorted([d for d in df['date'].dropna().unique()], reverse=True)
    # 문자열로 변환하여 표시
    unique_dates_str = [str(d) for d in unique_dates]
else:
    unique_dates_str = []

# multiselect for dates (전체 선택 가능)
selected_dates = st.sidebar.multiselect("타임스탬프 (날짜 선택)", options=["전체"] + unique_dates_str, default=["전체"] if unique_dates_str else ["전체"])

# 수강 교육과정 selectbox (고정 옵션 + 실제 고유값 포함)
course_options = ["전체", "공통교육", "Track1 교육", "Track2 교육", "Track3 교육"]
# Append any other unique raw values to allow selection
if COURSE_COL in df.columns:
    raw_courses = df[COURSE_COL].dropna().astype(str).unique().tolist()
    for rc in raw_courses:
        if rc not in course_options:
            course_options.append(rc)
selected_course = st.sidebar.selectbox("수강 교육과정", course_options, index=0)

# 소속본부, 소속부서 selectbox (전체 포함)
hq_options = ["전체"]
dept_options = ["전체"]
if HQ_COL in df.columns:
    hq_options += sorted(df[HQ_COL].dropna().astype(str).unique().tolist())
if DEPT_COL in df.columns:
    dept_options += sorted(df[DEPT_COL].dropna().astype(str).unique().tolist())

selected_hq = st.sidebar.selectbox("소속본부", hq_options, index=0)
selected_dept = st.sidebar.selectbox("소속부서", dept_options, index=0)

# 필터 적용 함수
def apply_filters(df: pd.DataFrame):
    filt = pd.Series([True] * len(df))
    # 날짜 필터: "전체"이면 무시, 아니면 선택한 날짜들만
    if selected_dates and "전체" not in selected_dates:
        # convert to date objects ensure matching
        sel_dates = []
        for s in selected_dates:
            try:
                sel_dates.append(pd.to_datetime(s).date())
            except Exception:
                pass
        filt &= df.get('date').isin(sel_dates)
    # 교육과정 필터
    if selected_course and selected_course != "전체":
        # if user selected canonical Track names, match course_cat; otherwise match original text
        if selected_course in ['공통교육','Track1 교육','Track2 교육','Track3 교육']:
            filt &= df.get('course_cat') == selected_course
        else:
            filt &= df.get(COURSE_COL).astype(str) == selected_course
    # 소속본부
    if selected_hq and selected_hq != "전체":
        filt &= df.get(HQ_COL).astype(str) == selected_hq
    # 소속부서
    if selected_dept and selected_dept != "전체":
        filt &= df.get(DEPT_COL).astype(str) == selected_dept
    return df.loc[filt].copy()

filtered_df = apply_filters(df)

# 탭 구성
tab1, tab2, tab3, tab4 = st.tabs(["Today 교육", "교육과정 별 시각화", "선행교육 수강 비교", "주관식 의견"])

# ---------------------------
# Tab 1: Today 교육
# ---------------------------
with tab1:
    st.header("Today 교육 (선택 필터 기준)") # 헤더 변경: '최신 날짜 기준'이 아님
    
    # 필터링된 데이터프레임 (filtered_df)을 그대로 사용
    recent_df = filtered_df.copy()

    if recent_df.empty:
        st.info("선택된 필터로 결과가 없습니다.")
    else:
        # 객관식 평균 계산 (숫자화된 컬럼 사용)
        avg_vals = {}
        for col in LIKERT_COLS:
            num_col = col + '_num'
            if num_col in recent_df.columns:
                avg_vals[col] = pd.to_numeric(recent_df[num_col], errors='coerce').mean()
            else:
                avg_vals[col] = None
        # DataFrame으로 변환하여 라인 차트 생성
        avg_df = pd.DataFrame({
            '문항': list(avg_vals.keys()),
            '평균점수': [v if v is not None else pd.NA for v in avg_vals.values()]
        }).dropna()

        if not avg_df.empty:
            
            # ********** 새로운 정렬 토글 버튼 추가 **********
            st.markdown("##### 정렬 옵션")
            sort_by_score = st.toggle('평균 점수가 높은 순으로 정렬', value=False, help="끄면 (기본값) 문항 순서대로, 켜면 점수 순으로 정렬됩니다.")
            
            # 2. 정렬 로직 및 Plotly 설정 정의
            if sort_by_score:
                # 1. 점수 순으로 정렬 (Plotly의 bottom-up을 고려하여 오름차순 정렬)
                avg_df = avg_df.sort_values(by='평균점수', ascending=True) 
                
                # Plotly 설정: 값에 따라 정렬 (total ascending은 bar height에 따라 정렬)
                yaxis_config = {
                    'categoryorder':'total ascending', 
                    'tickfont': dict(size=16) 
                }
                current_title = "객관식 문항별 평균 (점수 높은 순)"
            else:
                # 2. 기본값: LIKERT_COLS 순서대로 정렬 (Plotly의 bottom-up을 고려하여 역순 정렬)
                
                # DataFrame을 LIKERT_COLS 목록과 동일한 순서로 재정렬 
                avg_df['문항'] = pd.Categorical(avg_df['문항'], categories=LIKERT_COLS, ordered=True)
                avg_df = avg_df.sort_values('문항', ascending=False)
                
                # Plotly 설정: LIKERT_COLS 역순을 기준으로 배열 순서 강제
                yaxis_config = {
                    'categoryorder':'array', 
                    'categoryarray': LIKERT_COLS[::-1], 
                    'tickfont': dict(size=16) 
                }
                current_title = "객관식 문항별 평균 (문항 순서대로)"
            
            # 꺾은선 그래프 대신 수평 막대 그래프 사용 (가시성 및 비교 용이성 높임)
            fig = px.bar(
                avg_df, 
                x='평균점수', 
                y='문항', 
                orientation='h', # 수평 막대 그래프 설정
                title=current_title, # 정렬에 따른 동적 제목
                text='평균점수' # 막대 끝에 점수 표시
            )
            
            # 막대 그래프 레이아웃 업데이트
            fig.update_traces(texttemplate='%{text:.2f}', textposition='outside')
            fig.update_layout(
                xaxis_title="평균 점수 (0-5)",
                yaxis_title="문항",
                yaxis=yaxis_config, # 동적으로 설정된 yaxis_config 사용
                # X축 범위를 0부터 5.1까지 고정하여 5가 보이게 함
                xaxis=dict(range=[0, 5.1]), 
                width=1000, # 차트의 물리적 너비를 1000px로 설정 
                font=dict(size=16) # 전체 글꼴 크기를 16으로 키움 
            )
            
            st.plotly_chart(fig) # use_container_width=True 제거
        else:
            st.info("객관식 컬럼의 숫자형 변환 결과가 없습니다.")

        # 주관식: 각 주관식 컬럼별 상위 5개 단어 표시
        st.subheader("주관식 - 상위 단어 (상위 5개)")
        cols = st.columns(len(OPEN_COLS))
        for i, col in enumerate(OPEN_COLS):
            with cols[i]:
                # 해당 주관식 문항의 제목을 먼저 표시
                st.markdown(f"**{col}**") 
                
                if col in recent_df.columns:
                    top5 = top_n_words(recent_df[col].dropna().astype(str), n=5)
                    if top5:
                        words_df = pd.DataFrame(top5, columns=['단어', '빈도'])
                        st.table(words_df)
                    else:
                        st.write("주관식 응답이 없거나 한글 단어가 적습니다.")
                else:
                    st.write("컬럼 없음")

# ---------------------------
# Tab 2: 교육과정 별 시각화
# ---------------------------
with tab2:
    st.header("교육과정 별 객관식 점수 비교")
    # 그룹화: course_cat 기준으로 평균 계산
    if filtered_df.empty:
        st.info("선택된 필터로 결과가 없습니다.")
    else:
        # 그룹별 평균 계산 (각 문항별)
        group_cols = ['course_cat']
        agg_df = filtered_df.groupby('course_cat')[[c + '_num' for c in LIKERT_COLS if (c + '_num') in filtered_df.columns]].mean().reset_index()
        if agg_df.empty:
            st.info("데이터가 충분하지 않습니다.")
        else:
            # Melt to long format for easy plotting
            melt_df = agg_df.melt(id_vars='course_cat', var_name='문항', value_name='평균점수')
            # remove suffix '_num' from 문항 이름
            melt_df['문항'] = melt_df['문항'].str.replace('_num', '', regex=False)
            
            # ********** LIKERT_COLS_SHORT 정의 및 적용 시작 **********
            # 히트맵 X축 가독성 개선을 위한 문항 단축/줄바꿈 목록
            LIKERT_COLS_SHORT = {
                '해당 교육과정의 내용과 흐름에 전반적으로 만족하십니까?': '전반적 만족도',
                '교육 난이도는 본인의 수준에 적절했나요?': '교육\n난이도\n적절성',
                'ChatGPT를 활용하여 분석을 수행하는 것이 도움이 된다고 생각하십니까?': 'ChatGPT\n활용\n도움 여부',
                '참여한 교육 과정을 주변에 추천할 의향이 있으십니까?': '주변\n추천\n의향',
                '교육 일자 및 시간 배분이 적절했나요? (전체교육시간, 실습 비중)': '일자/시간\n배분\n적절성',
                '제공된 교육 과정 및 교안은 실무에서도 참고할 수 있을 만큼 실용적인가요?': '교안\n실무\n실용성',
                '실습 환경 (RAS, Python) 은 데이터 분석 흐름을 체험하기에 적절하였나요?': '실습 환경\n적절성',
                '강사들은 교육 내용을 명확하고 이해하기 쉽게 전달했다': '강사\n내용\n전달력',
                '질문이나 실습 요청사항에 도움을 주고, 수강생과 원활하게 소통했다': '강사\n소통/지원\n만족도' # '소통/피드백'에서 '소통/지원'으로 수정됨
            }
            
            # 히트맵을 위해 문항 이름을 단축 버전으로 변환
            melt_df['문항_short'] = melt_df['문항'].map(LIKERT_COLS_SHORT)
            # ********** LIKERT_COLS_SHORT 정의 및 적용 끝 **********
            
            # --- 히트맵 코드 시작 ---
            
            # 피벗 테이블 생성 (히트맵에 적합한 형태로 변환), '문항_short' 사용
            heatmap_df = melt_df.pivot_table(index='course_cat', columns='문항_short', values='평균점수', aggfunc='mean')
            
            # ********** 히트맵 X축 순서를 LIKERT_COLS 순서대로 강제 (유지) **********
            # LIKERT_COLS_SHORT의 '값' 부분 순서(줄 바꿈된 문항)를 추출하여 배열 정의
            ordered_x_axis = [LIKERT_COLS_SHORT[col] for col in LIKERT_COLS if col in LIKERT_COLS_SHORT]
            heatmap_df = heatmap_df[ordered_x_axis]
            # *******************************************************************
            
            fig_heatmap = px.imshow(
                heatmap_df,
                text_auto=".2f", # 소수점 둘째 자리까지 텍스트 표시
                aspect="auto",
                color_continuous_scale=px.colors.sequential.Viridis, # 색상 스케일
                title="교육과정 별 객관식 문항 평균 히트맵"
            )
            
            # 색상바 범위 설정 (0점 ~ 5점)
            fig_heatmap.update_coloraxes(colorbar_title="평균 점수 (0-5)", cmin=0, cmax=5)
            
            # X축 레이블 가독성 개선: 회전 0도, 하단 마진 증가
            fig_heatmap.update_xaxes(
                title_text="문항", 
                side="top",
                tickangle=0, # 레이블 회전 0도로 고정 (줄 바꿈 적용)
                tickfont=dict(size=12)
            ) 
            fig_heatmap.update_yaxes(title_text="교육과정")
            
            fig_heatmap.update_layout(
                margin=dict(t=80, b=180) # 상단 마진을 80으로, 하단 마진을 180으로 증가 (제목과 레이블 간 간격 확보)
            )
            
            st.plotly_chart(fig_heatmap, use_container_width=True)

            # --- 히트맵 코드 끝 ---
            
            st.subheader("교육과정 별 막대 그래프") 
            # 막대 그래프는 원본 문항 이름('문항')을 그대로 사용
            fig2 = px.bar(
                melt_df, 
                x='문항', 
                y='평균점수', 
                color='course_cat', 
                barmode='group',
                title="수강 교육과정 별 객관식 문항 평균 비교",
                text='평균점수' # 막대 위에 평균 점수 값 표시를 위한 설정
            )
            
            # 막대 위에 평균 점수 값 표시 및 형식 지정
            fig2.update_traces(texttemplate='%{text:.2f}', textposition='outside')
            
            # Y축 범위 최댓값 확장 및 레이아웃 상단 마진 추가 (텍스트 잘림 방지)
            fig2.update_layout(
                xaxis_title="문항", 
                yaxis_title="평균 점수 (0-5)", 
                legend_title="교육과정",
                yaxis=dict(range=[0, 5.4]), # Y축 최대 범위 5.4로 확장
                margin=dict(t=50) # 상단 마진 추가
            )
            st.plotly_chart(fig2, use_container_width=True)

# ---------------------------
# Tab 3: 선행교육 수강 비교
# ---------------------------
with tab3:
    st.header("선행교육 수강 여부에 따른 비교")
    if PREV_COURSE_COL not in filtered_df.columns:
        st.info("데이터에 선행교육 관련 컬럼이 없습니다.")
    else:
        # 수강 여부 컬럼 생성: 비어있지 않음 -> '수강함', 비어있음 -> '수강안함'
        filtered_df['선행수강여부'] = filtered_df[PREV_COURSE_COL].apply(lambda x: '수강안함' if pd.isna(x) or str(x).strip()=='' else '수강함')
        compare_df = filtered_df.groupby('선행수강여부')[[c + '_num' for c in LIKERT_COLS if (c + '_num') in filtered_df.columns]].mean().reset_index()
        if compare_df.empty:
            st.info("데이터가 충분하지 않습니다.")
        else:
            cmp_melt = compare_df.melt(id_vars='선행수강여부', var_name='문항', value_name='평균점수')
            cmp_melt['문항'] = cmp_melt['문항'].str.replace('_num','', regex=False)
            
            fig3 = px.bar(
                cmp_melt, 
                x='문항', 
                y='평균점수', 
                color='선행수강여부', 
                barmode='group',
                title="선행교육 수강 여부에 따른 객관식 평균 비교",
                text='평균점수' # 막대 위에 평균 점수 값 표시를 위한 설정
            )
            
            # 막대 위에 평균 점수 값 표시 및 형식 지정
            fig3.update_traces(texttemplate='%{text:.2f}', textposition='outside')
            
            # Y축 범위 최댓값 확장 및 레이아웃 상단 마진 추가 (텍스트 잘림 방지)
            fig3.update_layout(
                xaxis_title="문항", 
                yaxis_title="평균 점수 (0-5)", 
                legend_title="선행수강여부",
                yaxis=dict(range=[0, 5.4]), # Y축 최대 범위 5.4로 확장 (Tab 2와 동일하게 적용)
                margin=dict(t=50) # 상단 마진 추가
            )
            st.plotly_chart(fig3, use_container_width=True)

# ---------------------------
# Tab 4: 주관식 의견
# ---------------------------
with tab4:
    st.header("주관식 응답 분석 (단어 빈도 및 테이블)") # 헤더를 기존의 "단어 빈도 및 테이블"로 복원
    
    # ------------------ 워드 클라우드 섹션 및 관련 코드 삭제 ------------------
    # st.subheader("키워드 빈도 워드 클라우드") <-- 삭제됨
    # generate_wordcloud 호출 로직 삭제됨
    # st.markdown("---") <-- 삭제됨
    
    st.subheader("주관식 문항별 단어 빈도표") # 헤더를 "주관식 문항별 단어 빈도표"로 유지
    # ----------------------------------------------------------------------

    if filtered_df.empty:
        st.info("선택된 필터로 결과가 없습니다.")
    else:
        for col in OPEN_COLS:
            st.subheader(col)
            if col in filtered_df.columns:
                # 상위 10개 단어
                top10 = top_n_words(filtered_df[col].dropna().astype(str), n=10)
                if top10:
                    words_df = pd.DataFrame(top10, columns=['단어', '빈도'])
                    # 막대그래프 (Plotly)
                    fig = px.bar(words_df, x='단어', y='빈도', title=f"'{col}' 빈도 상위 10개")
                    fig.update_layout(xaxis_title="단어", yaxis_title="빈도수")
                    st.plotly_chart(fig, use_container_width=True)
                    # 표로도 출력
                    st.dataframe(words_df)
                else:
                    st.write("주관식 응답이 없거나 한국어 단어 추출 결과가 없습니다.")
            else:
                st.write("컬럼 없음")

# ---------------------------
# 하단: 추가 정보 및 예외 처리 안내
# ---------------------------
st.markdown("---")
st.markdown("**유의사항:** 데이터 컬럼명이 다르거나 파일 경로가 다를 경우 오류가 발생할 수 있습니다. 필요하면 컬럼명을 확인 후 수정해 주세요.")
