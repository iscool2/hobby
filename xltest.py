# ======================================================
# Excel file (.xlsx) read/write example with pandas 
# ======================================================
import pandas as pd
Fname = 'rawdata.xlsx'
SI = 'input'          #input sheet name
SO = 'output'         #output sheet name or sheet order number

df = pd.read_excel(Fname, sheet_name=[SI,SO], header = 0)
si = df[SI]
so = df[SO]
print(si.index)
print(so.index)

rdRow = 0
wrRow = 0
while rdRow<(len(si)-4):
    so.loc[wrRow, 'Word']       = si.loc[rdRow, 'DATA']
    so.loc[wrRow, 'Chinese']    = si.loc[rdRow+1, 'DATA']
    so.loc[wrRow, 'PingYin']    = si.loc[rdRow+2, 'DATA']
    so.loc[wrRow, 'English']    = si.loc[rdRow+3, 'DATA']   
    # increment output row index
    rdRow = rdRow + 4
    wrRow = wrRow + 1

with pd.ExcelWriter(Fname) as w:
    si.to_excel(w, sheet_name=SI, index=False)
    so.to_excel(w, sheet_name=SO, index=False)

print('bye bye~~')

# 3. header : 헤더(열) 지정
# - 열 이름(헤더)으로 사용할 행 지정 / 첫 행이 헤더가 아닌 경우 header = None 
# pd.read_excel('파일명.xlsx', header = 1)


# 4. names : 열 이름 변경
# - 불러오는 열의 개수와 일치해야한다.
# pd.read_excel('파일명.xlsx', names = ['col1', 'col2'])


# 5. usecols : 불러올 열 지정
# # 이름으로 지정
# pd.read_excel('파일명.xlsx', usecols = ['사용할열_1', '사용할열_2'])
# # 번호로 지정
# pd.read_excel('파일명.xlsx', usecols = [0, 1])


# 6. na_values : 결측값 인식하기
# - 결측값(NA / NaN)으로 인식 할 문자열 지정
# - '', '# N / A', '# N / AN / A', '#NA', '-1. # IND', '-1. # QNAN', '-NaN', '-nan', '1. # IND', '1. # QNAN', '<NA>', 'N / A', 'NA', 'NULL', 'NaN', 'n / a ','nan ','null '는 기본적으로 결측값으로 인식된다.
# pd.read_excel('파일명.xlsx', na_values = '결측값의_형태')


# 7. 불러올 행 제한
# nrows : 불러올 행 개수 제한 / 처음 ~ n번째 행만 불러오기 
# skiprows : 처음 ~ n번째 행 제외 / n+1번째 ~ 마지막까지
# skipfooter : 뒤에서 n개 제외
# pd.read_excel('파일명.xlsx', skiprows = n)   # 앞에서 n개 행 생략
# pd.read_excel('파일명.xlsx', nrows = n)   # 처음 ~ n번째
# pd.read_excel('example.xlsx', skipfooter = n)   # 뒤에서 n개 행 생략



