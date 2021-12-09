#엑셀 파일 합치기
#폴더안의 파일들을 효율적으로 읽는 코드

import pandas as pd
import os

all_data = pd.DataFrame()

for (path, dir, files) in os.walk("D:\#.Secure Work Folder\\2021년 업무\과제진척도\\211208\\"):
    for filename in files:
        ext = os.path.splitext(filename)[-1]
        if ext == '.xlsx':
            file_name = "D:\#.Secure Work Folder\\2021년 업무\과제진척도\\211208\\"+filename
            df = pd.read_excel(file_name)
            all_data = all_data.append(df, ignore_index = True)
            #print("%s/%s" % (path, filename))

#데이터갯수확인
print(all_data.shape)

#데이터 잘 들어오는지 확인
print (all_data.head())

#파일저장
print (all_data.to_excel("D:\#.Secure Work Folder\\2021년 업무\과제진척도\\211208\\total.xlsx", header=False, index=False))
