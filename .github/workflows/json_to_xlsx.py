import json
import openpyxl

data = open('.github/workflows/report 1.json','r',encoding = 'utf-8')
json_data = json.load(data) # 開いたJSONファイルを、Pythonで使えるデータに変換する


wb = openpyxl.Workbook()
sheet = wb.worksheets[0]

header_name = 'テストケース名'
header_time = '実行日時'
header_person = '実行者名'
header_results = '結果'
header_remark ='備考'


test = json_data['test_results']

filename = json_data['test_name']
file_name = f'{filename}xlsx'



for value in test:
    # name はリストtest_resultsの中にある辞書を１つずつ取り出す。
    test_case_name = value['test_case_name']
    test_officer = value['test_officer']
    test_date = value['test_date']
    result = value['result']
    remark = value.get('remark')



    test_report = [
        [header_name,header_time,header_person,header_results,header_time,header_remark],
        [test_case_name,test_officer,test_date,result,remark],
    ]

for i in test_report:
    sheet.append(i)



wb.save(file_name)
wb.close