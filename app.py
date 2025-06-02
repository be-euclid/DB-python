import streamlit as st
import pandas as pd
import re
import matplotlib.pyplot as plt
import numpy as np
import matplotlib
import platform

# 한글 폰트 설정 함수
def set_korean_font():
    if platform.system() == 'Windows':
        matplotlib.rc('font', family='Malgun Gothic')
    elif platform.system() == 'Darwin':  # macOS
        matplotlib.rc('font', family='AppleGothic')
    else:  # Linux/Streamlit Cloud
        matplotlib.rc('font', family='NanumGothic')
    matplotlib.rcParams['axes.unicode_minus'] = False

set_korean_font()

@st.cache_data
def load_all_data(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    sheet_names = xls.sheet_names
    df_list = []
    for sheet in sheet_names:
        try:
            df = pd.read_excel(xls, sheet_name=sheet)
            df['Year(sheet)'] = sheet
            df_list.append(df)
        except Exception:
            continue
    if df_list:
        return pd.concat(df_list, ignore_index=True)
    else:
        return pd.DataFrame()

def name_to_initials(full_name):
    parts = full_name.strip().split()
    if len(parts) == 3:
        last, first, patronymic = parts
        return f"{last} {first[0]}.{patronymic[0]}."
    elif len(parts) == 2:
        last, first = parts
        return f"{last} {first[0]}."
    else:
        return full_name

def normalize_name(name):
    return re.sub(r'\s+', ' ', name.strip().lower())

def match_names(df, name_input):
    name_input_norm = normalize_name(name_input)
    initials_input = name_to_initials(name_input)
    initials_input_norm = normalize_name(initials_input)
    name_cols = [col for col in df.columns if 'Name' in str(col) or 'name' in str(col)]
    if not name_cols:
        return pd.DataFrame()
    name_col = name_cols[0]
    def is_match(row_name):
        row_name_norm = normalize_name(str(row_name))
        row_initials_norm = normalize_name(name_to_initials(str(row_name)))
        return (
            name_input_norm == row_name_norm or
            initials_input_norm == row_name_norm or
            name_input_norm == row_initials_norm or
            initials_input_norm == row_initials_norm
        )
    mask = df[name_col].apply(is_match)
    return df[mask]

def reorder_columns(df):
    name_col = [col for col in df.columns if 'name' in str(col).lower()]
    year_col = [col for col in df.columns if 'year(sheet)' in str(col).lower()]
    if not name_col or not year_col:
        return df
    name_col = name_col[0]
    year_col = year_col[0]
    cols = list(df.columns)
    name_idx = cols.index(name_col)
    cols_wo_year = [c for c in cols if c != year_col]
    new_cols = cols_wo_year[:name_idx+1] + [year_col] + cols_wo_year[name_idx+1:]
    return df[new_cols]

def drop_all_none_columns(df):
    df = df.dropna(axis=1, how='all')
    return df.loc[:, ~(df == '').all()]

def get_position_counts_top7(df):
    pos_cols = [col for col in df.columns if 'position/title' in str(col).lower()]
    if not pos_cols:
        return None
    pos_col = pos_cols[0]
    counts = df[pos_col].value_counts(dropna=True)
    # 상위 7개, 나머지는 Others로 묶기
    if len(counts) > 7:
        top7 = counts[:7]
        others = counts[7:].sum()
        top7['Others'] = others
        return top7
    else:
        return counts

def get_party_counts(df):
    party_cols = [col for col in df.columns if 'party membership' in str(col).lower()]
    if not party_cols:
        return None
    party_col = party_cols[0]
    # None, NaN, 빈 문자열, 공백 등은 모두 Non-Party로 대체
    party_data = df[party_col].replace([None, np.nan, '', ' '], 'Non-Party')
    party_data = party_data.fillna('Non-Party')
    party_data = party_data.apply(lambda x: 'Non-Party' if str(x).strip() == '' else x)
    counts = party_data.value_counts()
    return counts

def normalize_party(x):
    if pd.isna(x) or str(x).strip() == '' or str(x).strip().lower() in ['non-party', 'non party', 'nonparty']:
        return 'Non-Party'
    return str(x).strip()

def get_party_counts_and_col(df):
    party_cols = [col for col in df.columns if 'party membership' in str(col).lower()]
    if not party_cols:
        return None, None
    party_col = party_cols[0]
    party_data = df[party_col].apply(normalize_party)
    counts = party_data.value_counts()
    return counts, party_col, party_data

def pie_chart_with_counts(counts, title, total):
    labels = [f"{idx} ({cnt})" for idx, cnt in zip(counts.index, counts.values)]
    def autopct_format(pct):
        count = int(round(pct * total / 100.0))
        return f"{pct:.1f}%\n({count})"
    fig, ax = plt.subplots(figsize=(6, 6))
    wedges, texts, autotexts = ax.pie(
        counts, labels=labels, autopct=autopct_format, startangle=90, counterclock=False,
        textprops={'fontsize': 10}
    )
    for autotext in autotexts:
        autotext.set_fontsize(9)
    ax.axis('equal')
    plt.figtext(0.98, 0.02, f"총 인원: {total}명", ha='right', va='bottom', fontsize=10, color='gray')
    plt.title(title, fontsize=13)
    return fig

uploaded_file = st.file_uploader("엑셀 파일(.xlsx) 업로드", type="xlsx")
if uploaded_file:
    df_all = load_all_data(uploaded_file)
    menu = st.sidebar.radio("메뉴 선택", ["이름 검색", "연도별 직위 분포", "연도별 Party Membership 분포"])
    if menu == "연도별 Party Membership 분포":
        years = sorted(df_all['Year(sheet)'].unique())
        selected_year = st.selectbox("연도를 선택하세요", years, key="party_year")
        year_df = df_all[df_all['Year(sheet)'] == selected_year]
        # None/공백(Non-Party) 숨기기 체크박스
        hide_nonparty = st.checkbox("'Non-Party' (None/공백 포함) 숨기기")
        counts, party_col, party_data = get_party_counts_and_col(year_df)
        if counts is not None and not counts.empty:
            st.subheader(f"{selected_year}년 Party Membership별 인원 분포")
            party_df = counts.rename('인원수').reset_index().rename(columns={'index': 'Party Membership'})
            if hide_nonparty:
                party_df = party_df[party_df['Party Membership'].str.lower() != 'non-party']
            if not party_df.empty:
                selected_party = st.radio("Party Membership을 선택하세요", party_df['Party Membership'])
                st.dataframe(party_df)
                filtered_df = year_df.copy()
                filtered_df[party_col] = party_data  # 정규화된 값으로 대체
                result = filtered_df[filtered_df[party_col].str.lower() == selected_party.strip().lower()]
                if not result.empty:
                    result = reorder_columns(result)
                    st.success(f"{selected_party} Party Membership을 가진 인물 목록 ({len(result)}명):")
                    st.dataframe(result)
                else:
                    st.warning("해당 Party Membership을 가진 인물이 없습니다.")
            else:
                st.warning("표시할 Party Membership이 없습니다.")
        else:
            st.warning("Party Membership 컬럼이 없거나 데이터가 없습니다.")
                
    elif menu == "이름 검색":
        name_input = st.text_input("이름을 입력하세요(예: Nemchinov Vasily Sergeevich)")
        hide_none = st.checkbox("None/빈값 컬럼 숨기기")
        if name_input:
            result = match_names(df_all, name_input)
            if not result.empty:
                result = reorder_columns(result)
                if hide_none:
                    result = drop_all_none_columns(result)
                st.success(f"{name_input} 검색 결과 :")
                st.dataframe(result)
            else:
                st.warning("해당 이름을 찾을 수 없습니다.")
                
    elif menu == "연도별 직위 분포":
        years = sorted(df_all['Year(sheet)'].unique())
        selected_year = st.selectbox("연도를 선택하세요", years)
        year_df = df_all[df_all['Year(sheet)'] == selected_year]
        counts = get_position_counts_top7(year_df)
        if counts is not None and not counts.empty:
            st.subheader(f"{selected_year}년 Position/Title별 인원 분포 (상위 7개 + Others)")
            total = int(counts.sum())
            fig = pie_chart_with_counts(counts, f"{selected_year}년 Position/Title 분포", total)
            st.pyplot(fig)
            st.dataframe(counts.rename('인원수'))
        else:
            st.warning("Position/Title 컬럼이 없거나 데이터가 없습니다.")
