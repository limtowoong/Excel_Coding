```
from openpyxl import load_workbook
import re
import os

folder_path = r'C:\Users\Office\Desktop\PY_TEST'

for file in os.listdir(folder_path):

    file_name, ext = os.path.splitext(file)

    match = re.search(r'(\d{4}년)_.*_(\w+)_\d$', file_name)

    if match:
        print(f"매칭된 파일: {file}")
        year = match.group(1)
        meter = match.group(2)
        break
    else:
        print(f"매칭되지 않은 파일: {file}")

print("원하는 파일 생성 수 : ", end=' ')
end_num = int(input())

for i in range(1, end_num):

    source_path = f'{year}_CNU_AMIGO_{meter}_{i}.xlsx'
    wb = load_workbook(source_path)

    ws = wb.active

    d1008_value = ws['D1008'].value

    for j in range(9, 1009):
        ws[f'D{j}'] = d1008_value + (j - 9) + 1
    destination_path = f'{year}_CNU_AMIGO_{meter}_{i+1}.xlsx'
    wb.save(destination_path)
```
