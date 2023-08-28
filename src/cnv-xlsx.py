import openpyxl as px
import json
import csv

# 出力データ
json_list = []
# 見出し行退避用
cap = []

#
# Excelファイルの置き場所を指定
#
inputdir = '/root/input'
outputdir = '/root/output'

#
# Excelファイルを開く
#
wb = px.load_workbook( inputdir + '/RESULT.xlsx')
print(type(wb))

#
# ワークシートを開く
#
ws = wb.worksheets[0]

#
# セルを読み込む
#

# 1～12行
for rows in range(1, 12):
    rowdata = ""
    # 1~6列目
#    for cols in range(1, 6):
        # print(ws.cell(rows, cols).value)
        
    if rows == 1:
        for cols in range(1, 7):
            cap.append( ws.cell(rows, cols).value )
    else:
            # print(ws.cell(rows, cols).number_format + ': ' + str(ws.cell(rows, cols).value) )
            # rowdata = rowdata + cap[cols - 1] + ': ' + str(ws.cell(rows, cols).value) + ','
        print      ("セルの書式は {0} です".format(ws.cell(rows, 1).number_format))

        rowdata = {
            cap[0]: str(ws.cell(rows, 1).value),
            cap[1]: str(ws.cell(rows, 2).value),
            cap[2]: str(ws.cell(rows, 3).value),
            cap[3]: str(ws.cell(rows, 4).value),
            cap[4]: str(ws.cell(rows, 5).value),
            cap[5]: str(ws.cell(rows, 6).value)
        }
            
    #rowdata = rowdata + '}'
    if rows > 1:
        json_list.append(rowdata)

p2j_data  = json.dumps(json_list, indent=2)
print(json.dumps(json_list, indent=2))

with open(outputdir + '/output.json', 'w') as f:
    json.dump(json_list, f)

wb.close
