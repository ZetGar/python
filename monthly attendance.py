#==========================
# 1단계 : 파일 합치기
#==========================

import os
import pandas as pd

# 폴더 경로 설정
folder_path = r'E:\M_test'

# 모든 엑셀 파일 목록 가져오기
file_list = os.listdir(folder_path)

# 엑셀 파일에서 데이터 읽어오기 및 합치기
combined_data = pd.DataFrame()
for file_name in file_list:
    if file_name.endswith('.xlsx') or file_name.endswith('.xls'):
        file_path = os.path.join(folder_path, file_name)
        # 엑셀 파일에서 데이터 읽어오기 (필터 해제)
        excel_data = pd.read_excel(file_path, sheet_name='Sheet1')  # 'Sheet1'로 수정
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
            for sheet in writer.book.sheetnames:
                ws = writer.book[sheet]
                ws.auto_filter.ref = ws.dimensions
        # 읽어온 데이터를 합치기
        combined_data = pd.concat([combined_data, excel_data])

# 합쳐진 데이터를 새로운 엑셀 파일로 저장
output_file_path = os.path.join(folder_path, 'combined_data.xlsx')
combined_data.to_excel(output_file_path, index=False)

print("합쳐진 데이터가", output_file_path, "에 저장되었습니다.")


#=======================================
# 2단계 : 지각 횟수, 합계 이름별 구하기
#=======================================

import pandas as pd

# 원본 데이터 파일 경로
input_file_path = r'E:\M_test\combined_data.xlsx'

# 결과 파일 저장 경로
output_file_path = r'E:\M_test\late_summary.xlsx'

# 원본 데이터 파일 읽기
data = pd.read_excel(input_file_path)

# 출근판정 열을 텍스트 형식으로 고정
data['출근판정'] = data['출근판정'].astype(str)

# 이름별로 출근판정이 '지각'인 행을 찾기
late_entries = data[data['출근판정'] == '지각']

# 이름별로 출근판정이 '지각'인 행의 지각시간 합 구하기
late_times_sum = late_entries.groupby('이름')['지각시간'].apply(lambda x: pd.to_timedelta(x + ':00').sum())

# 이름별로 출근판정이 '지각'인 행의 횟수 구하기
late_counts = late_entries.groupby('이름').size()

# 새로운 데이터프레임 생성
summary_data = pd.DataFrame({
    '이름': late_times_sum.index,
    '조직': late_entries.groupby('이름')['조직'].first(),
    '지각시간': late_times_sum.dt.components['hours'].astype(str) + ':' +
               late_times_sum.dt.components['minutes'].astype(str),
    '지각횟수': late_counts
})

# 결과 파일로 저장
summary_data.to_excel(output_file_path, index=False)

print("이름별 지각 정보가", output_file_path, "에 저장되었습니다.")

#=======================================
# 3단계 : 지각 추후에 연차 올렸나 확인
#=======================================

import pandas as pd

# '출근판정'이 '지각'인 행들만 가져오기
combined_data = pd.read_excel(r'E:\M_test\combined_data.xlsx')
h_confirm = combined_data[combined_data['출근판정'] == '지각']

# 'h_confirm' 파일 생성
h_confirm.to_excel(r'E:\M_test\h_confirm.xlsx', index=False)

# 'M_H' 파일 불러오기
m_h_data = pd.read_excel(r'E:\M_test\M_H.xlsx')

# '날짜' 열의 왼쪽 10글자를 날짜로 인식
m_h_data['날짜'] = pd.to_datetime(m_h_data['날짜'].str[:10])

# '성명'과 '날짜'가 'h_confirm' 파일의 '이름'과 '근무일자'와 동시에 일치하는 경우에 대해 출근판정 열 업데이트
for index, row in m_h_data.iterrows():
    mask = (h_confirm['이름'] == row['성명']) & (h_confirm['근무일자'] == row['날짜'])
    if not h_confirm[mask].empty:
        # h_confirm의 출근판정 열을 M_H의 근태구분 데이터로 덮어씌움
        m_h_data.loc[index, '출근판정'] = h_confirm[mask]['근태구분'].values[0]

# 수정된 데이터를 새로운 엑셀 파일로 저장
m_h_data.to_excel(r'E:\M_test\M_H_updated.xlsx', index=False)


#=======================================
# 4단계 : 순수 연장시간 계산
#=======================================


import pandas as pd

# 2단계: '퇴근판정'이 '연장근무'인 행을 복사하여 붙여넣습니다.
combined_data_path = r'E:\M_test\combined_data.xlsx'
combined_data = pd.read_excel(combined_data_path)
result_df = combined_data[combined_data['퇴근판정'] == '연장근무'].copy()

# '연장근무시간'과 '지각시간'을 분(minute) 단위의 숫자로 변환합니다.
def time_to_minutes(t):
    if pd.isna(t):
        return 0
    hours, minutes = map(int, t.split(':'))
    return hours * 60 + minutes

result_df['연장근무시간'] = result_df['연장근무시간'].apply(time_to_minutes)
result_df['지각시간'] = result_df['지각시간'].apply(time_to_minutes)

# 이름별로 데이터를 그룹화하고, 연장근무시간과 지각시간의 합계, 연장근무 횟수를 계산합니다.
grouped = result_df.groupby(['이름', '조직']).agg({
    '연장근무시간': 'sum',
    '지각시간': 'sum',
    '퇴근판정': 'count'  # 연장근무 횟수
}).reset_index()

grouped.rename(columns={'퇴근판정': '연장근무 갯수'}, inplace=True)

# 결과값과 1.5배 결과값 계산
grouped['결과값'] = grouped['연장근무시간'] - grouped['지각시간']
grouped['1.5배 결과값'] = grouped['결과값'] * 1.5

# 결과값을 hh:mm 형식으로 변환하는 함수
def minutes_to_time(minutes):
    hours = minutes // 60
    minutes = minutes % 60
    return f"{int(hours)}:{int(minutes):02d}"

grouped['결과값'] = grouped['결과값'].apply(minutes_to_time)
grouped['1.5배 결과값'] = grouped['1.5배 결과값'].apply(lambda x: minutes_to_time(int(x)))

# 최종 결과에 필요한 열만 포함하여 저장
final_columns = ['이름', '조직', '연장근무 갯수', '결과값', '1.5배 결과값']
final_df = grouped[final_columns]

# 결과를 WOW 파일에 저장
wow_file_path = r'E:\M_test\wow.xlsx'
final_df.to_excel(wow_file_path, index=False, sheet_name='Sheet1')

#=======================================
# 5단계 : 연차계산
#=======================================

import pandas as pd

# 파일 경로
file_path = 'E:/M_test/M_H.xlsx'

# Excel 파일을 읽어옵니다.
df = pd.read_excel(file_path)

# 근태구분에 따른 치환 값 정의
attendance_mapping = {
    '경조휴가': 0,
    '교육': 0,
    '기타': 1,
    '년월차': 1,
    '리프레쉬': 0,
    '법정휴가': 0,
    '병가': 0,
    '보상연차': 1,
    '보상오전': 0.5,
    '보상오후': 0.5,
    '오전반차': 0.5,
    '오후반차': 0.5,
    '훈련': 0
}

# 치환하기 전에 '보상연차', '보상오전', '보상오후'에 해당하는 데이터를 선별하여 각각 합계를 계산
compensation_df = df[df['근태구분'].isin(['보상연차', '보상오전', '보상오후'])]
compensation_summary = compensation_df.groupby('성명').size().reset_index(name='보상합계')

# 이제 전체 근태구분을 치환합니다.
df['근태구분'] = df['근태구분'].map(attendance_mapping)

# '성명' 별로 총합 계산
total_summary = df.groupby('성명')['근태구분'].sum().reset_index()

# 최종 데이터프레임을 합칩니다: total_summary에 compensation_summary를 병합
final_summary = pd.merge(total_summary, compensation_summary, on='성명', how='left').fillna(0)

# '부서명' 매칭 추가
# 원본 데이터에서 '성명'과 '부서명'을 매칭
department_mapping = df[['성명', '부서명']].drop_duplicates().set_index('성명')['부서명'].to_dict()
final_summary['부서명'] = final_summary['성명'].map(department_mapping)

# 새로운 Excel 파일로 저장
final_summary.to_excel('E:/M_test/M_H_final_summary.xlsx', index=False)
