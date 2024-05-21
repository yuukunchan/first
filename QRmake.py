import tkinter as tk
import tkinter.ttk as ttk
import gspread
from google.oauth2 import service_account #google-auth
import customtkinter as ctk
import qrcode
import PIL.ImageTk #Pillow
import json
from tkinter import filedialog
import os
import openpyxl

class Hikisuu:
    
    def __init__(self) -> None:
        
        if os.path.isfile("settei.json") == False:
            with open("settei.json","w") as f:        
                json.dump(self.settei,f,ensure_ascii=False)
        else:
            with open("settei.json","r") as f:
                self.settei = json.load(f)
            self.keyfile = self.settei["jファイル"]
            self.sp_id = self.settei["sp_id"]
            self.seat_no = self.settei["シート番号"]
            self.cl_r_no = self.settei["右端列番号"] 
            self.cl_l_no = self.settei["左端列番号"]
            self.cl_name_no = self.settei["名前列番号"]
            self.cl_line_no = self.settei["行番号"]
                                
    qr_dic={}
    wb=0
    pulldown=["更新ボタンを押してください"]
    sp_id=''
    img = 0
    combobox=0
    keyfile = ""
    scope = []
    ws = ""
    
    settei = {'jファイル': '', 'sp_id': '', "シート番号":"","右端列番号":"","左端列番号":"C","名前列番号":"B","行番号":""}
    file_name =""
    path =0
    testfile=0
    seat_no=0
    cl_r_no=""
    cl_l_no="A"
    cl_name_no="B"
    cl_line_no=0
    
H=Hikisuu()

#メイン
root = tk.Tk()
root.title("QRコード生成")
root.geometry("600x620")

#タブ作成
tabframe = ttk.Notebook(root)
tab1 = tk.Frame(tabframe,width=600)
tabframe.add(tab1,text="コード生成")
tab2 = tk.Frame(tabframe,width=600)
tabframe.add(tab2,text="設定")

#スプレッドシートと連携する関数
def load(self):
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
    Keyfile = self.keyfile
    credentials = service_account.Credentials.from_service_account_file(Keyfile)
    scoped_credentials = credentials.with_scopes(scope)
    gc = gspread.Client(auth=scoped_credentials)
    sp_id = self.sp_id
    wb = gc.open_by_key(sp_id)
    self.ws = wb.get_worksheet(self.seat_no)
    
#プルダウンリストを更新する関数
def saiyomikomi(self):      
    load(self)
    ws=self.ws
    #print(self.cl_name_no)
    for a in range(int(self.cl_line_no),600):
        cell = self.cl_name_no +str(a)
        check = ws.acell(cell).value
        if check is None:
            cell_range = self.cl_name_no + str(self.cl_line_no) + ":" +self.cl_name_no + str(a) 
            break   
        else:
            cell_range = self.cl_name_no + str(self.cl_line_no) + ":" +self.cl_name_no + str(a-1)
    sp_range = ws.get(cell_range)
    self.pulldown = []
    for i in range(len(sp_range)):
       self.pulldown.append(sp_range[i][0])
       self.qr_dic[sp_range[i][0]] = i           
    combobox.configure(values=self.pulldown)
      
#QR生成関数
def make_qr(self):
    
    load(self)#スプレッドシートと連携する関数の呼び出し
    
    ws = self.ws
    item = ws.get(self.cl_l_no + str(self.cl_line_no -1) +":"+ self.cl_r_no + str(self.cl_line_no-1))[0]
    sn_key = combobox.get()
    
    sn = self.qr_dic[sn_key]
    date_no =self.cl_l_no + str(sn+self.cl_line_no) + ":" + self.cl_r_no + str(sn+self.cl_line_no)
    patient_date = ws.get(date_no)[0]
    qr_text = []
    item_range = openpyxl.utils.column_index_from_string(self.cl_r_no) - openpyxl.utils.column_index_from_string(self.cl_l_no)+1
    
    if len(patient_date) != item_range:
        
        for k in range(item_range - len(patient_date)):
            patient_date.append("")
        
    for j in range(item_range):
        if patient_date[j] != "":
            qr_text.append(item[j])
            qr_text.append("：")
            qr_text.append(patient_date[j]+ "\r\n")
                
    qr_text = "".join(qr_text)
    qr = qrcode.QRCode(box_size=2)
    qr.add_data(qr_text.encode('shift_jis'))
    qr.make()
    qr_img = qr.make_image()
    
    self.t = sn_key
    patient_name.configure(text=self.t+"さんのQRコード")
    self.img =PIL.ImageTk.PhotoImage(image = qr_img)
    patient_qr.configure(image=self.img)


def file_btn_click(self):
    self.path = filedialog.askopenfilename()
    self.file_name.set(self.path)
    file_name_la.configure(textvariable=self.file_name)    

def file_save(self):
    if file_name_la.get() != "":
        self.keyfile = str(os.path.basename(self.path))
        j_name_la.configure(text="キーファイル　：　"+self.keyfile)
        self.settei["jファイル"]= self.keyfile
        with open("settei.json","w") as f:        
            json.dump(self.settei,f,ensure_ascii=False)
           
def sp_id_btn_click(self):
    if sp_id_la.get() != "":
        self.sp_id = sp_id_la.get()
        sp_name_la.configure(text="sp_id　：　" + self.sp_id)
        self.settei["sp_id"] = self.sp_id
        with open("settei.json","w") as f:        
            json.dump(self.settei,f,ensure_ascii=False)

def seat_btn_click(self):
    if seat_no_la.get() != "":
        self.seat_no = int(seat_no_la.get())
        seat_name_la.configure(text="シート番号　：　" + str(self.seat_no))
        self.settei["シート番号"] = self.seat_no
        with open("settei.json","w") as f:        
            json.dump(self.settei,f,ensure_ascii=False)

def cl_r_click(self):
    if cl_r_la.get() != "":
        self.cl_r_no = cl_r_la.get()
        cl_r_name_la.configure(text="列番号　"+ self.cl_r_no + "　まで")
        self.settei["右端列番号"] = self.cl_r_no
        with open("settei.json","w") as f:        
            json.dump(self.settei,f,ensure_ascii=False)

def cl_l_click(self):
    if cl_l_la.get() != "":
        self.cl_l_no = cl_l_la.get()
        cl_l_name_la.configure(text="列番号　" + self.cl_l_no + "　から")
        self.settei["左端列番号"] = self.cl_l_no
        with open("settei.json","w") as f:        
            json.dump(self.settei,f,ensure_ascii=False)

def cl_name_click(self):
    if cl_name_la.get() != "":
        self.cl_name_no = cl_name_la.get()
        cl_name_name_la.configure(text="名前が載る列　" + self.cl_name_no + "　列")
        self.settei["名前列番号"] = self.cl_name_no
        with open("settei.json","w") as f:        
            json.dump(self.settei,f,ensure_ascii=False)
            
def cl_line_click(self):
    if cl_line_la.get() != "":
        self.cl_line_no = int(cl_line_la.get())
        cl_line_name_la.configure(text="行番号　"+ str(self.cl_line_no) + "　まで")
        self.settei["行番号"] = self.cl_line_no
        with open("settei.json","w") as f:        
            json.dump(self.settei,f,ensure_ascii=False)            

def hozon(self):
    hozon_pop = tk.messagebox.askyesno("確認","入力値を保存しますか？")
    if hozon_pop == True:
        file_save(self)
        sp_id_btn_click(self)
        seat_btn_click(self)
        cl_r_click(self)
        cl_l_click(self)
        cl_name_click(self)
        cl_line_click(self)

def qrchange(event):
    make_qr(H)

#tab1

#生成ボタン
#make_qr_btn = ctk.CTkButton(tab1,text="コード生成",command= lambda:make_qr(H))
#make_qr_btn.grid(row=1,column=0,columnspan=3,pady=20)

#読み込みボタン
quit_btn = ctk.CTkButton(tab1,text = '更新',command = lambda:saiyomikomi(H),width=100)
quit_btn.grid(row=0,column=2)

#プルダウン
combobox = ctk.CTkComboBox(master=tab1,values=H.pulldown,width=300,command = qrchange)
combobox.grid(row=0,column=0,columnspan=2)

#QRコードの名称
#patient_name = ctk.CTkLabel(tab1,text="生成ボタンを押すと↓にQRコードが表示されます")
patient_name = ctk.CTkLabel(tab1,text="名前を選択すると↓にQRコードが表示されます\n(読込に少し時間がかかります)")
patient_name.grid(row=2,column=0,columnspan=3,pady=5)


#QR表示ラベル
patient_qr = tk.Label(tab1,text="テスト")
na = ""
H.img = tk.PhotoImage(file = na)
patient_qr.configure(image=H.img)
patient_qr.grid(row=3,column=0,columnspan=3)

#tab2

#説明1
setumei_la1 = ctk.CTkLabel(tab2,width=400,text="プスレッドシートのjsonファイルを選んでください")
setumei_la1.grid(row=0,column=0,columnspan=3,pady=5)

#ファイルパス表示テキストボックス
H.file_name = tk.StringVar()
file_name_la = ctk.CTkEntry(tab2,textvariable = H.file_name,width=200)
file_name_la.grid(row=1,column=0,columnspan=2)

#ファイル選択ぼたん
file_btn = ctk.CTkButton(tab2,
                         text="ファイル名参照",
                         width=100,
                         command=lambda:file_btn_click(H))
file_btn.grid(row=1,column=2)

#設定ボタン
file_save_btn = ctk.CTkButton(tab2,
                          text="jsonファイルを設定",
                          width=100,
                          command=lambda:file_save(H))
#file_save_btn.grid(row=,column=)

#説明2
setumei_la2 = ctk.CTkLabel(tab2,text="sp_idを入力してください")
setumei_la2.grid(row=3,column=0,columnspan=3,pady=5)

#sp_id入力欄
sp_id_la = ctk.CTkEntry(tab2,textvariable = "",width=300)
sp_id_la.grid(row=4,column=0,columnspan=3)

#sp_idボタン
sp_id_btn = ctk.CTkButton(tab2,
                          text="sp_idを設定",
                          width=100,
                          command=lambda:sp_id_btn_click(H))
#sp_id_btn.grid(row=,column=,columnspan=)

#説明3
setumei_la3 = ctk.CTkLabel(tab2,text="シート番号を入力してください(一番左が0番です)")
setumei_la3.grid(row=6,column=0,columnspan=2,pady=20)

#シートナンバー記入欄
seat_no_la = ctk.CTkEntry(tab2,textvariable = "",width=30,)
seat_no_la.grid(row=6,column=2)

#シートナンバー設定ボタン
seat_btn = ctk.CTkButton(tab2,
                         text="シート番号を設定",
                         width=100,
                         command=lambda:seat_btn_click(H))
#seat_btn.grid(row=,column=)

#説明6
setumei_la6 = ctk.CTkLabel(tab2,text="QRコード化する列範囲(アルファベット)")
setumei_la6.grid(row=8,column=0,columnspan=3,pady=5)

#列番号入力欄
cl_l_la = ctk.CTkEntry(tab2,textvariable = "",width=30,)
cl_l_la.grid(row=9,column=1,sticky=tk.W)

#列番号設定ボタン
cl_l_btn = ctk.CTkButton(tab2,
                       text="開始列番号を設定",
                       width=100,
                       command=lambda:cl_l_click(H))
#cl_l_btn.grid(row=10,column=0)

setumei_la7 = ctk.CTkLabel(tab2,text="から")
setumei_la7.grid(row=9,column=1)

#列番号入力欄
cl_r_la = ctk.CTkEntry(tab2,textvariable = "",width=30,)
cl_r_la.grid(row=9,column=2)

setumei_la7 = ctk.CTkLabel(tab2,text="まで")
setumei_la7.grid(row=9,column=2,sticky=tk.E)


#列番号設定ボタン
cl_r_btn = ctk.CTkButton(tab2,
                       text="終了列番号を設定",
                       width=100,
                       command=lambda:cl_r_click(H))
#cl_r_btn.grid(row=,column=)

#説明8
setumei_la8 = ctk.CTkLabel(tab2,text="名前が載る列(プルダウンリストとなる列)")
setumei_la8.grid(row=11,column=0,columnspan=2,pady=10)

#プルダウンに入る列の記入欄
cl_name_la = ctk.CTkEntry(tab2,textvariable = "",width=30,)
cl_name_la.grid(row=11,column=2)

#プルダウン列の設定ボタン
cl_name_btn = ctk.CTkButton(tab2,
                         text="名前の列を設定",
                         width=100,
                         command=lambda:cl_name_click(H))
#cl_name_btn.grid(row=,column=)

#説明9
setumei_la9 = ctk.CTkLabel(tab2,text="1人目の患者データが載る行")
setumei_la9.grid(row=13,column=0,columnspan=2,pady=10)



#行番号入力欄
cl_line_la = ctk.CTkEntry(tab2,textvariable = "",width=30,)
cl_line_la.grid(row=13,column=2)

#行番号設定ボタン
cl_line_btn = ctk.CTkButton(tab2,
                            text="開始行番号を設定",
                            width=100,
                            command=lambda:cl_line_click(H))
#cl_line_btn.grid(row=10,column=0)


#一括保存ボタン
hozon_btn = ctk.CTkButton(tab2,
                          text="保存",
                          width=100,
                          command=lambda:hozon(H))
hozon_btn.grid(row=15,column=1)


#説明4
setumei_la4 = ctk.CTkLabel(tab2,text="現在の設定値")
setumei_la4.grid(row=19,column=0,columnspan=3,pady=5)

#設定されたファイル名ラベル
j_name_la = ctk.CTkLabel(tab2,text = "ファイル名　：　"+ H.settei["jファイル"])
j_name_la.grid(row=20,column=0,columnspan=3)

#設定されたsp_id
sp_name_la = ctk.CTkLabel(tab2,text = "sp_id　：　"+ H.settei["sp_id"])
sp_name_la.grid(row=21,column=0,columnspan=3)

#設定されたシート番号
seat_name_la = ctk.CTkLabel(tab2,text = "シート番号　：　"+ str(H.settei["シート番号"]))
seat_name_la.grid(row=22,column=0,columnspan=3)

#設定された列番号
setumei_la7 = ctk.CTkLabel(tab2,text="QRコード化する範囲　：　")
setumei_la7.grid(row=23,column=0)

cl_l_name_la = ctk.CTkLabel(tab2,text = "列番号　"+ H.settei["左端列番号"]+"　から")
cl_l_name_la.grid(row=23,column=1,)

cl_r_name_la = ctk.CTkLabel(tab2,text = "列番号　"+ H.settei["右端列番号"]+"　まで")
cl_r_name_la.grid(row=23,column=2,)

#設定された列番号
cl_name_name_la = ctk.CTkLabel(tab2,text = "名前が載る列　"+ H.settei["名前列番号"]+"　列(基本的にB列のままでOK)")
cl_name_name_la.grid(row=24,column=0,columnspan=3)

#設定された列番号
cl_line_name_la = ctk.CTkLabel(tab2,text = "1人目のデータが載る行　"+ str(H.settei["行番号"])+"　行目から(基本的に2行目のままでOK)")
cl_line_name_la.grid(row=25,column=0,columnspan=3)


tabframe.pack()

root.mainloop()

