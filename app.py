import streamlit as st
import pandas as pd
import re

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
    """
    'Nemchinov Vasily Sergeevich' → 'Nemchinov V.S.'
    """
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
    # 정규화
    name_input_norm = normalize_name(name_input)
    initials_input = name_to_initials(name_input)
    initials_input_norm = normalize_name(initials_input)

    name_cols = [col for col in df.columns if 'Name' in str(col) or 'name' in str(col)]
    if not name_cols:
        return pd.DataFrame()
    name_col = name_cols[0]

    # 모든 이름을 정규화 및 이니셜화
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

st.title("DB 이름 검색기")

uploaded_file = st.file_uploader("엑셀 파일(.xlsx) 업로드", type="xlsx")
if uploaded_file:
    df_all = load_all_data(uploaded_file)
    name_input = st.text_input("이름을 입력하세요(예: Nemchinov Vasily Sergeevich)")
    if name_input:
        result = match_names(df_all, name_input)
        if not result.empty:
            st.success(f"{name_input} 검색 결과 :")
            st.dataframe(result)
        else:
            st.warning("해당 이름을 찾을 수 없습니다.")
