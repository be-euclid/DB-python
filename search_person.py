import pandas as pd

def search_person_info(file_path, search_name):
    xls = pd.ExcelFile(file_path)
    sheet_names = xls.sheet_names
    results = []

    for sheet in sheet_names:
        try:
            df = pd.read_excel(file_path, sheet_name=sheet)
            name_cols = [col for col in df.columns if 'Name' in col or 'name' in col]
            if not name_cols:
                continue
            name_col = name_cols[0]
            mask = df[name_col].astype(str).str.strip().str.lower() == search_name.strip().lower()
            person_rows = df[mask]
            if not person_rows.empty:
                person_rows['Year(sheet)'] = sheet
                results.append(person_rows)
        except Exception as e:
            print(f"시트 {sheet} 처리 중 오류: {e}")

    if results:
        all_results = pd.concat(results, ignore_index=True)
        print(all_results.T)
        return all_results
    else:
        print("해당 이름을 찾을 수 없습니다.")
        return None

if __name__ == "__main__":
    file_path = "DB.xlsx"
    search_name = input("검색할 이름을 입력하세요: ")
    search_person_info(file_path, search_name)
