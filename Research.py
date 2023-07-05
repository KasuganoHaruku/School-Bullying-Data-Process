import openpyxl
from openpyxl import load_workbook

path = "C:/Users/haruk/Downloads/226347839_0_青少年校园霸凌与心理健康调查_208_207.xlsx"
wb = load_workbook(path)
sheet = wb['Sheet1']
sheet2 = wb.create_sheet("Data")
print(wb.sheetnames)
m_row = sheet.max_row

s1_start = (sheet['J2'].column)
s1_end = (sheet['X2'].column)
i = s1_start
j = 2
sta = 0
while(j <= m_row):
    while(i <= s1_end):
        score = sheet.cell(row=j,column=i).value
        tmp = i-s1_start+1
        if tmp in (3,6,10,15):
            if(score == "极其相符"):
                score = 1
            elif(score == "非常相符"):
                score = 2
            elif(score == "中等相符"):
                score = 3
            elif(score == "部分相符"):
                score = 4
            elif(score == "完全不符"):
                score = 5
        else:
            if(score == "极其相符"):
                score = 5
            elif(score == "非常相符"):
                score = 4
            elif(score == "中等相符"):
                score = 3
            elif(score == "部分相符"):
                score = 2
            elif(score == "完全不符"):
                score = 1
        sta = sta + score
        sheet2.cell(row=j,column=i,value=score).value
        sheet2.cell(row=1,column=i,value=tmp).value
        i = i+1
    sheet2.cell(row=j,column=s1_end + 1,value=sta).value
    i = s1_start
    sta = 0
    j = j+1
print(i - s1_start)
print(j-1)
wb.save("data.xlsx")