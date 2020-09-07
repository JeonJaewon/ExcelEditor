import openpyxl
from openpyxl.styles import PatternFill
from openpyxl import Workbook
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
FILE_NAME = "default_name.xlsx"
GRAY = "AEAAAA"  # 회색
WHITE = "FFFFFF"  # 흰색

root = Tk()
root.geometry("300x250")

file_frame = Frame(root)
lbl = Label(file_frame, text="편집을 원하는 엑셀 파일을 선택하세요.")
lbl.pack()

def openFile():
    root.filename = filedialog.askopenfilename(initialdir="./", title="Select file",
                                               filetypes=(("excel files", "*.xlsx"), ("all files", "*.*")))
    global FILE_NAME
    FILE_NAME = root.filename
    print(root.filename)
btn = Button(file_frame, text="파일 열기", command=openFile)
btn.pack()
file_frame.pack(pady=20)


frame_col = Frame(root)
lbl2 = Label(frame_col, text="고객 구분 기준으로 할 엑셀 열을 선택하세요 (A~Z)")
lbl2.pack()

addressCol = IntVar()
radios = []
for i in range (0,8):
    radios.append(Radiobutton(frame_col, text=chr(ord("A")+i), variable=addressCol, value=i))
    if i == 3:
        radios[i].invoke()
    radios[i].pack(side=LEFT)
frame_col.pack(pady=20)

def start():
    if FILE_NAME == "default_name.xlsx":
        messagebox.showwarning("오류", "파일을 선택하세요")
        return
    print("started col : "+ str(addressCol.get()))
    book = openpyxl.load_workbook(FILE_NAME)  # 엑셀 파일
    sheet = book.worksheets[0]
    # 주소를 기준으로 전 사람과 현재 사람을 비교!
    prevAddr = ""
    curAddr = ""
    curColor = GRAY
    for row in sheet.rows:
        curAddr = row[addressCol.get()].value
        if (prevAddr == curAddr):  # 내 윗칸과 주소가 같다면
            for cell in row:
                cell.fill = PatternFill(start_color=curColor, fill_type="solid")  # 같은색으로 칠하기
        else:  # 주소가 다르다면
            curColor = WHITE if (curColor == GRAY) else GRAY  # 다른새긍로 칠하기
            for cell in row:
                cell.fill = PatternFill(start_color=curColor, fill_type="solid")
        prevAddr = curAddr

    #  띄어쓰기도 인식해서 주소가 '완전히 동일'하게 입력되어야함.
    #  ex) '광진구 ' 와 '광진구' 는 다른 주소로 인식함.
    print("done")
    messagebox.showinfo("성공", "작업을 완료했습니다.")
    book.save("송장출력.xlsx")


btn = Button(root, text="Start", command=start, width=60)
btn.pack(pady=20)

root.mainloop()
