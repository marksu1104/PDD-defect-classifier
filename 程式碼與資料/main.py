import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import ctypes
import loadmodel
import subprocess

class PDD(tk.Tk):
    
    def __init__(self,root):
        
        self.root=root
        self.root.title('PDD印刷不良分類')
        #調用api取得目前的縮放因子
        self.ScaleFactor = ctypes.windll.shcore.GetScaleFactorForDevice(0)
        #設置縮放因子
        self.root.call('tk', 'scaling', self.ScaleFactor/75)
        
        self.createWidget()

    def createWidget(self):
        
        #建立frame(for 按鈕並列)
        self.frame1 = tk.Frame(self.root)
        
        #建立變數
        self.success_var = tk.StringVar()  #文字變數:成功提示文字
        self.success_var.set('')
        self.little_var = tk.StringVar()   #文字變數:小提示
        self.little_var.set('')
        self.bar_int = tk.IntVar()         #整數變數:進度條
        self.bar_int.set(0)
 
        #建立物件
        self.mytittle = tk.Label(self.root, text='PDD印刷不良分類',         #開頭
                                 font=('Microsoft JhengHei UI',20,'bold'))
        self.mytittle.pack(pady=20)
        
        self.label = tk.Label(self.root,        #說明文字
                              text='請先點擊開啟檔案選擇要分類的PDD報告，再點擊執行。\n您選擇的檔案是(可一次選取多個檔案)：', 
                              font=('Microsoft JhengHei UI',10), justify='left')
        self.label.pack(anchor='w', padx=20)
        
        self.text = tk.Text(self.root, height=10,       #text  
                            width= 70, font=('Microsoft JhengHei UI',10))
        self.text.pack(padx=20)  
        
        self.success = tk.Label(self.root, textvariable=self.success_var,       #成功提示文字
                                font=('Arial',15), fg='#00DB00')
        self.success.pack()
        
        self.bar = ttk.Progressbar(root,length=600,variable=self.bar_int)       #進度條
        self.bar.pack()
        
        self.frame1.pack()      #要用frame前必須呼叫
        
        self.btn_open = tk.Button(self.frame1,      #開啟檔案按鈕
                text='開啟檔案',
                font=('Microsoft JhengHei UI',15),
                command=self.show
                )
        self.btn_open.grid(column=0,row=2,padx=10,pady=10)

        self.btn_do = tk.Button(self.frame1,        #執行按鈕
                        text='執行',
                        font=('Microsoft JhengHei UI',15),
                        command=self.do_classify
                        ) 
        self.btn_do.grid(column=1,row=2,padx=10,pady=10)
        
        self.little_label = tk.Label(self.root, textvariable=self.little_var,       #小提示
                                     font=('Microsoft JhengHei UI',10),justify='right')
        self.little_label.pack(padx=3,pady=3,anchor='e')
        
    def show(self):         #顯示檔案位置
        self.success_var.set('')
        self.bar_int.set(0)
        files = filedialog.askopenfilenames(initialdir='../',title="Select PDF file",
                                            filetypes=(("pdf", "*.pdf"),))
        self.text.delete(1.0,'end')
        for i, e in enumerate(files):          # 使用 enumerate 將串列變成帶有索引值的字典
            self.text.insert('end', f'{e}')         # Text 從後方加入內容
            if i != len(files):
                self.text.insert('end', '\n')       # 如果是不是最後一行，都加入換行符號
        
    def do_classify(self):           #執行按鈕動作
        self.bar_int.set(0)
        content = self.text.get(1.0,'end-1c')    # 取得文字輸入框內容，使用 end-1c 不要取得最後換行
        new_files = content.split('\n')          # 從換行符進行換行
        del new_files[-1]                        # 刪除最後一筆空資料
        
        if len(new_files) == 0:                  # 判斷是否有選擇檔案
            self.success_var.set('未讀取到檔案！') 
        else:   
            try:
                count = 1
                self.success_var.set('請稍後...')
                root.update()   #記得更新介面
                for i in new_files:
                    loadmodel.model(i)
                    self.bar_int.set(count/(len(new_files))*100)
                    count = count+1
                    root.update()  
                self.success_var.set('執行成功！')
                self.little_var.set('若要分類新的數據，請點擊開啟檔案')
                
                #開啟結果資料夾
                os.chdir('./結果')
                subprocess.Popen(r'explorer.exe .')
                os.chdir('../')                         #記得回上一頁
                    
                    
            except Exception as e:
                self.success_var.set('')
                messagebox.showwarning('錯誤', f'檔案、資料有誤！\n錯誤訊息：{e}')    
        

if __name__ == '__main__':
    root = tk.Tk()
    
    #調用api設置成由應用程序縮放
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
    
    pdd = PDD(root)
    root.mainloop()
