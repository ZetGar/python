#-------------------------------
#주간근태 마무리!!!
#-------------------------------
#▶▶▶무조건!! 97-2002 엑셀형식 워크시트 형식으로 다름이름 저장 갈겨!!!◀◀◀
#-------------------------------
#▶▶▶무조건!! 97-2002 엑셀형식 워크시트 형식으로 다름이름 저장 갈겨!!!◀◀◀
#-------------------------------
#▶▶▶무조건!! 97-2002 엑셀형식 워크시트 형식으로 다름이름 저장 갈겨!!!◀◀◀
#-------------------------------
#▶▶▶무조건!! 97-2002 엑셀형식 워크시트 형식으로 다름이름 저장 갈겨!!!◀◀◀
#-------------------------------
#▶▶▶무조건!! 97-2002 엑셀형식 워크시트 형식으로 다름이름 저장 갈겨!!!◀◀◀
#-------------------------------

import pandas as pd

from datetime import datetime

# 파일 불러오기
file_path = r"E:\w_test\w_2_3_s.xlsx"
df = pd.read_excel(file_path)

# 현재 날짜를 기준으로 연도와 월을 가져옴
today = datetime.now()
year_month = today.strftime('%Y-%m')

# 첫 번째 행의 데이터를 기준으로 날짜 형식 변경
first_row = df.iloc[0]  # 첫 번째 행 선택
for col in df.columns:
    # 열 이름이 '일'로 끝나는 경우에만 처리
    if col.endswith('일'):
        # '일' 문자를 제외하고 숫자 부분만 추출하여 날짜 생성
        day = int(col[:-1])
        date = f"{year_month}-{day:02d}"  # yyyy-mm-dd 형식으로 날짜 생성
        df.rename(columns={col: date}, inplace=True)  # 컬럼 이름 변경

# 변경된 DataFrame을 새 파일로 저장
output_file_path = r"E:\w_test\updated_dates.xlsx"
df.to_excel(output_file_path, index=False)
print(f"변환된 데이터가 {output_file_path} 경로에 저장되었습니다.")


# 파일1과 파일2를 읽어들임
file1 = pd.read_excel(r'E:\w_test\w_2_3.xlsx')
file2 = pd.read_excel(r'E:\w_test\updated_dates.xlsx')
file3 = pd.read_excel(r'E:\w_test\w_2_3_h.xlsx')

# 파일1의 '근무일자' 값을 문자열로 변환하여 yyyy-mm-dd 형식으로 고정
file1['근무일자'] = file1['근무일자'].astype(str)

# 파일2의 첫 행을 문자열로 변환하여 yyyy-mm-dd 형식으로 고정
file2.columns = file2.columns.astype(str)

# '근무일자' 값을 파일2의 열 이름과 비교하여 해당하는 값을 찾아 업데이트
for index, row in file1.iterrows():
    date_to_find = str(row['근무일자'])
    if date_to_find in file2.columns:
        column_to_search = file2[date_to_find]
        
        name_info = row['이름']
        if pd.notnull(name_info):
            matching_cells = column_to_search[column_to_search.astype(str).str.contains(name_info)]
            if not matching_cells.empty:
                value_to_update = matching_cells.iloc[0]
                file1.at[index, '출근판정'] = value_to_update
            else:
                # 값이 없을 때 파일1의 값을 그대로 유지
                file1.at[index, '출근판정'] = row['출근판정']
        else:
            # 이름 정보가 없을 때 파일1의 값을 그대로 유지
            file1.at[index, '출근판정'] = row['출근판정']
    else:
        # 해당하는 날짜가 파일2에 없을 때 파일1의 값을 그대로 유지
        file1.at[index, '출근판정'] = row['출근판정']

# 업데이트된 파일1을 새로운 엑셀 파일로 저장
file1.to_excel(r'E:\w_test\updated_w_2_3.xlsx', index=False)

file4 = pd.read_excel(r'E:\w_test\updated_w_2_3.xlsx')

# 파일4과 파일3 읽어오기
df_file4 = pd.read_excel(r'E:\w_test\updated_w_2_3.xlsx')
df_file3 = pd.read_excel(r'E:\w_test\w_2_3_h.xlsx')

# '근무일자'와 '이름' 데이터를 강제로 텍스트로 인식
df_file4['근무일자'] = df_file4['근무일자'].astype(str)
df_file4['이름'] = df_file4['이름'].astype(str)

# '날짜'와 '성명' 데이터를 강제로 텍스트로 인식
df_file3['날짜'] = df_file3['날짜'].astype(str)
df_file3['성명'] = df_file3['성명'].astype(str)

# '날짜' 열에서 앞에서 10글자만 추출하여 변경
df_file3['날짜'] = df_file3['날짜'].str[:10]

# 파일1과 파일2를 대조하여 조건에 맞는 행 찾기
for index, row in df_file4.iterrows():
    match = df_file3[(df_file3['날짜'] == row['근무일자']) & (df_file3['성명'] == row['이름'])]
    if not match.empty:
        # 조건에 맞는 행이 있을 경우
        if match.iloc[0]['근태구분'] != '오후반차':
            # '오후반차'가 아닌 경우 파일1의 해당 행의 '출근판정' 값으로 덮어쓰기
            df_file4.loc[index, '출근판정'] = match.iloc[0]['근태구분']

# 최종 결과를 새 파일로 저장
output_file_path = r"E:\w_test\final.xlsx"
df_file4.to_excel(output_file_path, index=False)
print(f"최종 결과가 {output_file_path} 경로에 저장되었습니다.")


# 'final' 파일을 file_path로 선언
file_path = r'E:\w_test\final.xlsx'

def update_attendance(file_path):
    print(file_path)
    
    try:
        # 엑셀 파일 읽기
        df = pd.read_excel(file_path)
        
        # 출근판정을 '파견'으로 변경할 이름 리스트
        names_to_update = ['윤길준', '임재원', '이영석', '박정상', '강철', '임무상', '전종훈', '이경록', '이예린', '이승훈', '박다인', '정관홍', '김홍근', '김종하', '김태용', '이철희', '장보경', '장영순', '강철', '한주석', '윤서희']
        
        # '이름' 열에서 출근판정을 '파견'으로 변경
        df.loc[df['이름'].isin(names_to_update), '출근판정'] = '파견'
        
        # '조직' 값이 '퇴사', '임원', '일용직', '유니닥스'인 행 삭제
        df = df[~df['조직'].isin(['퇴사', '임원', '일용직', '유니닥스'])]
        
        # 결과 파일의 경로 및 파일명 설정
        output_file_path = os.path.join(os.path.dirname(file_path), 'result.xlsx')
        
        # 변경된 내용을 새로운 파일로 저장
        df.to_excel(output_file_path, index=False)
        
        print(f"새로운 파일이 {output_file_path} 경로에 생성되었습니다.")
    except Exception as e:
        print(f"오류 발생: {e}") 

# update_attendance 함수를 호출하여 'final' 파일을 처리하고 result 파일로 저장합니다.
update_attendance(file_path)


