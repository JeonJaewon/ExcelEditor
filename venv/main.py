import openpyxl
from openpyxl.styles import PatternFill
from openpyxl import Workbook
from tkinter import *
from tkinter import filedialog

root = Tk()
lbl = Label(root, text="편집을 원하는 엑셀 파일을 선택하세요.")
lbl.pack()

lbl2 = Label(root, text="고객 구분 기준으로 할 엑셀 열을 선택하세요 (A~Z).")
lbl2.pack()

btn = Button(root, text="OK")
btn.pack()

root.filename = filedialog.askopenfilename(initialdir = "./",title = "Select file",filetypes = (("excel files","*.xlsx"),("all files","*.*")))
fileName = root.filename


ADDRESS_COL_NUM = 2
fileName = "송장출력.xlsx"
book = openpyxl.load_workbook(fileName) #엑셀 파일
sheet = book.worksheets[0]
prevAddr = ""
curAddr = ""
color1 = "FFC7CE" #부농
color2 = "FFFFFF" #흰색
curColor = color1
for row in sheet.rows:
    curAddr = row[ADDRESS_COL_NUM].value
    if(prevAddr == curAddr): #내 윗칸과 주소가 같다면
        for cell in row:
            cell.fill = PatternFill(start_color=curColor, fill_type="solid")
    else:
        curColor = color2 if (curColor == color1) else color1
        for cell in row:
            cell.fill = PatternFill(start_color=curColor, fill_type="solid")
    prevAddr = curAddr

#  띄어쓰기도 인식해서 주소가 '완전히 동일'하게 입력되어야함.
#  ex) '광진구 ' 와 '광진구' 는 다른 주소로 인식함.
book.save("송장출력.xlsx")

