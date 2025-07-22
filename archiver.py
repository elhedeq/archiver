import tkinter as tk 
import tkinter.font as tkFont
from tkinter import ttk, messagebox, filedialog
import sqlite3
from PDFViewer import PDFViewerWidget
import shutil
import os
from datetime import datetime, timedelta
from PIL import Image, ImageTk
import sys
import traceback
import pandas as pd

# Global Error Handling Function
def handle_exception(exc_type, exc_value, exc_traceback):
    error_message = "".join(traceback.format_exception(exc_type, exc_value, exc_traceback))
    print(error_message)
    messagebox.showerror("Unexpected Error", f"حدث خطأ غير متوقع:\n{error_message}")

# Apply global error handling
sys.excepthook = handle_exception

LARGEFONT = ("Arial", 35)

class tkinterApp(tk.Tk):
    
    # __init__ function for class tkinterApp 
    def __init__(self, *args, **kwargs): 
        
        # __init__ function for class Tk
        tk.Tk.__init__(self, *args, **kwargs)
        self.shared_data = {}  # Dictionary to store values
        self.title("Archiver")
        self.geometry("1440x900")
        self.resizable(False, False)
        # Modify existing default font
        default_font = tkFont.nametofont("TkDefaultFont")
        default_font.configure(family="Arial", size=14)  # Set font globally
        screen_width = self.winfo_screenwidth()
        if screen_width < 1200:
            default_font.configure(family="Arial", size=10)  # Set font globally
            LARGEFONT = ("Arial", 25)
        else:
            default_font.configure(family="Arial", size=14)  # Set font globally
            LARGEFONT = ("Arial", 35)
        # creating a container
        container = tk.Frame(self)
        container.pack(side = "top", fill = "both", expand = True) 
 
        container.grid_rowconfigure(0, weight = 1)
        container.grid_columnconfigure(0, weight = 1)
 
        # initializing frames to an empty array
        self.frames = {}  
 
        # iterating through a tuple consisting
        # of the different page layouts
        for F in (StartPage, addLetter, search, editLetter):
 
            frame = F(container, self)
 
            # initializing frame of that object from
            # startpage, addLetter, search respectively with 
            # for loop
            self.frames[F] = frame 
 
            frame.grid(row = 0, column = 0, sticky ="nsew")
 
        self.show_frame(StartPage)

    def set_data(self, key, value):
        self.shared_data[key] = value

    def get_data(self, key):
        return self.shared_data.get(key, None)
 
    # to display the current frame passed as
    # parameter
    def show_frame(self, cont, **kwargs):
        frame = self.frames[cont]
        if isinstance(frame, search):
            frame.refreshKeywords()
            frame.refreshAdressee()
        if isinstance(frame, addLetter):
            frame.getNumber()
        if isinstance(frame, editLetter):
            frame.getOldValues(**kwargs)
        frame.tkraise()
 
# first window frame startpage
 
class StartPage(tk.Frame):
    def __init__(self, parent, controller): 
        self.conn = sqlite3.connect('archive.db')

        tk.Frame.__init__(self, parent)
         
        leftSpacer = tk.Label(self,width=30)
        leftSpacer.grid(row=0,column=0, sticky="nsew")
        
        self.logo = Image.open('logo.png')
        self.logo = self.logo.resize((int(915/6*5), int(667/6*5)))
        self.logo = ImageTk.PhotoImage(self.logo)
        img_width, img_height = self.logo.width(), self.logo.height()
        self.img = tk.Label(self, image=self.logo)
        self.img.grid(row=0,column=1,columnspan=4,rowspan=5, sticky="nsew")

        signature = tk.Label(self, text="", font=('Arial',11,'bold'), fg="blue",width=30)
        signature.grid(row=9,column=2, columnspan=2, sticky="nsew")

        style = ttk.Style()
        style.configure("Custom.TButton", font=('Arial',16))

        addLetterPageBtn = ttk.Button(self, text ="إضافة خطاب 📩", style='Custom.TButton', command = lambda : controller.show_frame(addLetter))
        addLetterPageBtn.grid(row = 7, column = 3,columnspan=2, padx = 10, pady = 10, ipadx=40, ipady=30, sticky="nsew")
 
        searchPageBtn = ttk.Button(self, text ="بحث 🔍", style='Custom.TButton', command = lambda : controller.show_frame(search))
        searchPageBtn.grid(row = 7, column = 1, columnspan=2, padx = 10, pady = 10, ipadx=40, ipady=30, sticky="nsew")

        exportBtn = ttk.Button(self, text ="excel تصدير البيانات لملف 📂", style='Custom.TButton', command = self.exportData)
        exportBtn.grid(row = 8, column = 1, columnspan=2, padx = 10, pady = 10, sticky="nsew")

        aboutBtn = ttk.Button(self, text ="مميزات البرنامج ✅", style='Custom.TButton', command = self.showInfo)
        aboutBtn.grid(row = 8, column = 3, columnspan=2, padx = 10, pady = 10, sticky="nsew")
        
    def exportData(self):
        filetypes = (
            ('pdf files', '*.pdf'),
            ('All files', '*.*')
        )

        folder = filedialog.askdirectory(title='Select Folder')
        if not folder:
            messagebox.showerror("Error", 'لم يتم تحديد المجلد')
            exit()
        outgoing_letters_query = f"""
                    SELECT number, strftime('%Y/%m/%d', date, 'unixepoch') AS date, adressees.name AS adressee, 
                    GROUP_CONCAT(IFNULL(letter_keywords.keyword, 'لا توجد كلمات بحث')) AS keywords ,
                    CASE 
                        WHEN received = 0 THEN 'email'
                        WHEN received = 1 THEN 'يدوي'
                    END AS received
                    FROM outgoing_letters LEFT JOIN adressees ON outgoing_letters.adressee = adressees.id 
                    LEFT JOIN outgoing_letter_keywords ON outgoing_letters.id = outgoing_letter_keywords.letterid 
                    LEFT JOIN letter_keywords ON outgoing_letter_keywords.keywordid = letter_keywords.id 
                    GROUP BY outgoing_letters.number"""
        
        incoming_letters_query = f"""
                    SELECT number, strftime('%Y/%m/%d', date, 'unixepoch') AS date, adressees.name AS adressee, GROUP_CONCAT(IFNULL(letter_keywords.keyword, 'لا توجد كلمات بحث')) AS keywords , incoming_letters."order"
                    FROM incoming_letters LEFT JOIN adressees ON incoming_letters.adressee = adressees.id 
                    LEFT JOIN incoming_letter_keywords ON incoming_letters.id = incoming_letter_keywords.letterid LEFT JOIN letter_keywords ON incoming_letter_keywords.keywordid = letter_keywords.id 
                    GROUP BY incoming_letters.number"""
        
        df_outgoing = pd.read_sql_query(outgoing_letters_query, self.conn)
        df_incoming = pd.read_sql_query(incoming_letters_query, self.conn)

        with pd.ExcelWriter(folder+'/سجل الخطابات.xlsx', engine="openpyxl") as writer:
            df_outgoing.to_excel(writer, sheet_name="صادر", index=False)
            df_incoming.to_excel(writer, sheet_name="وارد", index=False)

        messagebox.showinfo('تم الحفظ','تم حفظ البيانات بنجاح\n'+folder+'/سجل الخطابات.xlsx')
    
    def showInfo(self):
        features = """                                                                                                                                                                                      
                    ✅ تخزين واسترجاع الخطابات بسهولة
                    ✅ ترقيم الخطابات تلقائياً
                    ✅ البحث بالرقم، التاريخ، الجهة ، أو الكلمات المفتاحية
                    ✅ فتح وعرض ملفات الخطابات داخل البرنامج
                    ✅ تعديل البيانات، وحذف السجلات غير الضرورية
                    ✅ ترتيب الملفات داخل مجلدات حسب الجهة والسنة
                    ✅ إمكانية تصدير البيانات لملف excel 
                    
                    """
        
        messagebox.showinfo('مميزات البرنامج', features)

 
# second window frame addLetter 
class addLetter(tk.Frame):
    keyWordsOptions = []
    adresseeOptions = []
    letterTypeSelect = None
    letterNumber = None
    letterYear = None
    letterMonth = None
    letterDay = None
    adresseeSelect = None
    adresseeEntry = None
    keyWordsSelect = None
    letterTable = 'outgoing_letters'
    filename = None
    dayValues = {
        1: [i for i in range(1,32)],
        2: [i for i in range(1,30)],
        3: [i for i in range(1,32)],
        4: [i for i in range(1,31)],
        5: [i for i in range(1,32)],
        6: [i for i in range(1,31)],
        7: [i for i in range(1,32)],
        8: [i for i in range(1,32)],
        9: [i for i in range(1,31)],
        10: [i for i in range(1,32)],
        11: [i for i in range(1,31)],
        12: [i for i in range(1,32)]
    }
    
    def __init__(self, parent, controller):

        self.controller = controller
        self.conn = sqlite3.connect('archive.db')
        self.cursor = self.conn.cursor()

        self.cursor.execute('SELECT name FROM adressees')
        for i in self.cursor.fetchall():
            self.adresseeOptions.append(i[0])

        self.cursor.execute('SELECT keyword FROM letter_keywords')
        for i in self.cursor.fetchall():
            self.keyWordsOptions.append(i[0])

        tk.Frame.__init__(self, parent)

        self.logo = Image.open('logo.png')
        self.logo = self.logo.resize((int(915/8), int(667/8)))
        self.logo = ImageTk.PhotoImage(self.logo)
        img = tk.Label(self, image=self.logo)
        img.grid(row=0,column=5, sticky="nsew")

        title = ttk.Label(self, text ="إضافة خطاب", font = LARGEFONT)
        title.grid(row = 0, column = 3, padx = 10, pady = 10, sticky="nsew")
 
        searchPageBtn = ttk.Button(self, text ="بحث 🔍",command = lambda : self.controller.show_frame(search))
        searchPageBtn.grid(row = 9, column = 3, padx = 10, pady = 10, sticky="nsew")
        
        homeBtn = ttk.Button(self, text ="الصفحة الرئيسية 🏠", command = lambda : self.controller.show_frame(StartPage))
        homeBtn.grid(row = 9, column = 2, padx = 10, pady = 10, sticky="nsew")

        self.previewPlaceholder = tk.Canvas(self, width=500, height=700)
        self.previewPlaceholder.grid(row = 0, column = 0, rowspan=8, sticky="nsew")
        self.previewPlaceholder.config(state="disabled")

        openFileBtn = ttk.Button(self, text ="اختر ملف 📄", command = self.selectFile)
        openFileBtn.grid(row = 1, column = 5, padx = 10, pady = 10, sticky="nsew")

        letterTypeLabel = ttk.Label(self, text ="صادر / وارد")
        letterTypeLabel.grid(row = 2, column = 5, padx = 10, pady = 10, sticky="nsew")

        self.letterTypeSelect = ttk.Combobox(self, values=['صادر', 'وارد'])
        self.letterTypeSelect.current(0)
        self.letterTypeSelect.grid(row = 2, column = 3, padx = 10, pady = 10, sticky="nsew")
        self.letterTypeSelect.bind('<<ComboboxSelected>>', self.selectTable)
        
        letterNumberLabel = ttk.Label(self, text ="رقم الخطاب")
        letterNumberLabel.grid(row = 1, column = 3, padx = 10, pady = 10, sticky="nsew")

        self.letterNumberVar = tk.StringVar()
        self.letterNumber = tk.Entry(self,  textvariable=self.letterNumberVar)
        self.letterNumber.grid(row = 1, column = 2, padx = 10, pady = 10, sticky="nsew")

        dateLabel = ttk.Label(self, text = 'التاريخ')
        dateLabel.grid(row = 3, column = 5, padx = 10, pady = 10)

        self.combo_var = tk.StringVar()
        current_datetime = datetime.now()
        today = current_datetime.strftime("%Y-%m-%d")
        self.combo_var.set(today[:4])
        self.letterYear = ttk.Combobox(self, values=[i for i in range(1961,2100)],textvariable=self.combo_var)
        self.letterYear.grid(row = 3, column = 3, padx = 10, pady = 10, sticky="nsew")
        self.letterYear.bind('<<ComboboxSelected>>', lambda event:self.getNumber())
        self.combo_var.trace_add("write", self.getNumber)

        self.letterMonth = ttk.Combobox(self, values=[i for i in range(1,13)])
        self.letterMonth.grid(row = 3, column = 2, padx = 10, pady = 10, sticky="nsew")
        self.letterMonth.set(int(today[5:7]))
        self.letterMonth.bind('<<ComboboxSelected>>', self.setDayValues)

        self.letterDay = ttk.Combobox(self, values=[i for i in range(1,32)])
        self.letterDay.grid(row = 3, column = 1, padx = 10, pady = 10, sticky="nsew")
        self.letterDay.set(int(today[-2:]))

        adresseeLabel = ttk.Label(self, text ="الجهة")
        adresseeLabel.grid(row = 4, column = 5, padx = 10, pady = 10, sticky="nsew")

        self.adresseeEntry = ttk.Entry(self)
        self.adresseeEntry.grid(row = 4, column = 1, columnspan=4, padx = 10, pady = 10, sticky="nsew")
        self.adresseeEntry.bind('<KeyRelease>',self.updateadresseeOptions)

        self.adresseeSelect = tk.Listbox(self, height=5)
        self.adresseeSelect.grid(row=5,column=1, columnspan=4,padx=10,pady=10, sticky="nsew")
        for item in self.adresseeOptions:
            self.adresseeSelect.insert('end', item)
        self.adresseeSelect.bind('<<ListboxSelect>>',self.autocompleteAdressee)
        adressee_scrollbar = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.adresseeSelect.yview)
        adressee_scrollbar.grid(row=5, column=3, sticky="nse")
        self.adresseeSelect.configure(yscrollcommand=adressee_scrollbar.set)

        keywordsLabel = ttk.Label(self, text ="كلمات البحث \n (يفصل بينها بفاصلة)")
        keywordsLabel.grid(row = 6, column = 5, padx = 10, pady = 10, sticky="nsew")

        self.keywordsEntry = ttk.Entry(self)
        self.keywordsEntry.grid(row=6, column=1, columnspan=4, padx = 10, pady = 10, sticky="nsew")
        self.keywordsEntry.bind('<KeyRelease>',self.updatekeywordOptions)

        self.keyWordsSelect = tk.Listbox(self, height=5)
        self.keyWordsSelect.grid(row=7,column=1, columnspan=4,padx=10,pady=10, sticky="nsew")
        for item in self.keyWordsOptions:
            self.keyWordsSelect.insert('end', item)
        self.keyWordsSelect.bind('<<ListboxSelect>>',self.autocompleteKeywords)
        keywords_scrollbar = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.keyWordsSelect.yview)
        keywords_scrollbar.grid(row=7, column=3, sticky="nse")
        self.keyWordsSelect.configure(yscrollcommand=keywords_scrollbar.set)

        order_placeholder = ttk.Label(self)
        order_placeholder.grid(row=8,column=1,pady=50)

        submit = ttk.Button(self,text="حفظ 💾",command=self.saveData)
        submit.grid(row=9,column=5,padx=10,pady=10, sticky="nsew")
        self.getNumber()
 
    def setDayValues(self,event):
        self.letterDay.configure(values=self.dayValues[int(event.widget.get())])
        self.letterDay.set("")

    def getNumber(self,*args):
        year = self.combo_var.get().strip()
        if not year.isdigit() or len(year) != 4:
            return
        try:
            year = int(year)
        except ValueError:
            messagebox.showerror("Error", f"البيانات غير صحيحة \n السنة يجب أن تكون رقم")
        start_date = int((datetime.strptime(f'{year}-1-1', "%Y-%m-%d")-datetime(1970,1,1)).total_seconds())
        end_date = int((datetime.strptime(f'{year}-12-31', "%Y-%m-%d")-datetime(1970,1,1)).total_seconds())
        self.cursor.execute(f'SELECT number FROM {self.letterTable} WHERE date >= {start_date} AND date <= {end_date}  ORDER BY number DESC LIMIT 1')
        number = self.cursor.fetchone()
        if(number):
            number = int(number[0])
            number += 1
        else:
            number = 1
        self.letterNumberVar.set(str(number))

    def updatekeywordOptions(self, event):
        value = event.widget.get().split(',')[-1].strip()
        
        # get data from l
        if value == '':
            data = self.keyWordsOptions
        else:
            data = []
            for item in self.keyWordsOptions:
                if value.lower() in item.lower():
                    data.append(item)                
    
        # update data in listbox
        self.updateKeywordsData(data)
    
    def updateKeywordsData(self,data):
        
        # clear previous data
        self.keyWordsSelect.delete(0, 'end')
    
        # put new data
        for item in data:
            self.keyWordsSelect.insert('end', item)
    
    def autocompleteKeywords(self,event):
        listbox = event.widget
        selected_index = listbox.curselection()
        selected_item = [listbox.get(i) for i in selected_index]
        if selected_item:
            keywords = self.keywordsEntry.get().split(',')
            keywords[-1] = selected_item[0]
            self.keywordsEntry.delete(0,'end')
            self.keywordsEntry.insert('end', ','.join(keywords)+',')
    
    def selectTable(self,event):
        self.letterTable = 'outgoing_letters' if self.letterTypeSelect.get()=='صادر' else 'incoming_letters'
        self.getNumber()
        if self.letterTable == 'incoming_letters':
            self.orderLabel = ttk.Label(self, text ="التأشيرة")
            self.orderLabel.grid(row = 8, column = 5, columnspan=2, padx = 10, pady = 10, sticky="nsew")

            self.order = tk.Text(self,height=5, width=10)
            self.order.grid(row=8,column=1, columnspan=4,padx=10,pady=10,sticky="nsew")
            self.text_scrollbar = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.order.yview)
            self.text_scrollbar.grid(row=8, column=3, sticky="nse")
            self.order.configure(yscrollcommand=self.text_scrollbar.set)
            self.order.bind("<Control-c>", self.copy_text)
            self.order.bind("<Control-v>", self.paste_text)
            self.order.bind("<Control-x>", self.cut_text)
        else:
            try:
                self.orderLabel.destroy()
                self.order.destroy()
                self.text_scrollbar.destroy()
            except:
                pass

    def updateadresseeOptions(self, event):
        value = event.widget.get()
        
        # get data from l
        if value == '':
            data = self.adresseeOptions
        else:
            data = []
            for item in self.adresseeOptions:
                if value.lower() in item.lower():
                    data.append(item) 
        # update data in listbox
        self.updateAdresseeData(data)
    
    def updateAdresseeData(self,data):
        
        # clear previous data
        self.adresseeSelect.delete(0, 'end')
    
        # put new data
        for item in data:
            self.adresseeSelect.insert('end', item)
    
    def autocompleteAdressee(self, event):
        listbox = event.widget
        selected_index = listbox.curselection()
        selected_item = [listbox.get(i) for i in selected_index]
        if selected_item:
            self.adresseeEntry.delete(0,'end')
            self.adresseeEntry.insert('end',selected_item[0])
    
    def selectFile(self):
        filetypes = (
            ('pdf files', '*.pdf'),
            ('All files', '*.*')
        )

        self.filename = filedialog.askopenfilename(
            title='Open a file',
            initialdir='/',
            filetypes=filetypes)

        try:
            if self.filename[-3:] != 'pdf':
                messagebox.showerror("Error", f"اختر ملف من نوع PDF")
                return
            self.pdfView = PDFViewerWidget.create_pdf_viewer(self,self.filename)
            self.pdfView.grid(row = 0, column = 0, rowspan=8, sticky="nsew")
        except ValueError:
            messagebox.showerror("Error", f"اختر ملف من نوع PDF")
            return
    
    def saveData(self):
        order = ''
        if self.letterTable == 'incoming_letters':
            order = self.order.get("1.0", "end-1c").strip()
        letter = {
            'type' : self.letterTable,
            'number' : self.letterNumber.get().strip(),
            'year' : self.combo_var.get().strip(),
            'month' : self.letterMonth.get().strip(),
            'day' : self.letterDay.get().strip(),
            'adressee' : self.adresseeEntry.get().strip(),
            'keywords' : self.keywordsEntry.get().strip(),
            'order' : order
        }
        for (key,value) in letter.items():
            if not value or not self.filename:
                if key == 'order':
                    continue
                data = f"""
                        نوع الخطاب: {'صادر' if letter['type'] == 'outgoing_letters' else 'وارد' if letter['type'] == 'incoming_letters' else letter['type']} 
                        رقم الخطاب: {letter['number'] if letter['number'] else 'لا بيانات'}
                        تاريخ الخطاب: {letter['year'] +'/'+ letter['month'] +'/'+ letter['day'] if letter['year'] and letter['month'] and letter['day'] else 'غير مكتمل'}
                        الجهة: {letter['adressee'] if letter['adressee'] else 'لا بيانات'}
                        كلمات البحث: {letter['keywords'] if letter['keywords'] else 'لا بيانات'}
                        ملف الخطاب: {self.filename if self.filename else 'لا بيانات'}
                        """
                messagebox.showwarning("warning", f"البيانات غير مكتملة \n {data}")
                return
        # try:
        self.cursor.execute('SELECT id FROM adressees WHERE name = ?', (letter["adressee"],))
        adresseeId = self.cursor.fetchone()
        if not adresseeId:
            self.cursor.execute('INSERT INTO adressees (name) VALUES (?)', (letter["adressee"],))
            self.conn.commit()
            self.adresseeOptions.append(letter["adressee"])
            self.cursor.execute('SELECT id FROM adressees WHERE name = ?', (letter["adressee"],))
            adresseeId = self.cursor.fetchone()
        try:
            if self.letterTypeSelect.get() not in ('صادر','وارد'):
                raise Exception('نوع الخطاب يجب أن يكون "صادر" أو "وارد"')
            letter['number'] = int(letter['number'])
            try:
                letter['year'] = int(letter['year'])
            except ValueError:
                raise Exception('السنة يجب أن تكون رقم')
            try:
                letter['month'] = int(letter['month'])
            except ValueError:
                raise Exception('الشهر يجب أن يكون رقم')
            try:
                letter['day'] = int(letter['day'])
            except ValueError:
                raise Exception('اليوم يجب أن يكون رقم')
            if letter['year']<1961:
                raise Exception('السنة يجب أن تكون أكبر من 1960 \n أنشأت وزارة التعليم العالى عام 1961 وفقا للقرار الجمهورى رقم 1665 لعام 1961')
            if letter['month']<1 or letter['month'] > 12:
                raise Exception('الشهر يجب أن يكون بين 1 و 12')
            if letter['day']<1 or letter['day'] > 31:
                raise Exception('اليوم يجب أن يكون بين 1 و 31')
        except Exception as e:
            messagebox.showerror("Error", f"البيانات غير صحيحة \n {e}")
            return
        date = int((datetime.strptime(f'{letter["year"]}-{letter["month"]}-{letter["day"]}', "%Y-%m-%d")-datetime(1970,1,1)).total_seconds())
        query = {
            'query' : f'INSERT INTO {letter["type"]} (number, date, adressee) VALUES (?, ?, ?)',
            'values' : (letter["number"], date, adresseeId[0],)
        }
        if letter["order"]:
            query = {
                'query' : f'INSERT INTO {letter["type"]} (number, date, "order", adressee) VALUES (?, ?, ?, ?)',
                'values' : (letter["number"], date, letter["order"], adresseeId[0],)
            }
        print("registered letter data")
        # insert letter data into letter table
        self.cursor.execute(query['query'], query['values'])
        # get id of the recently registered latter
        self.cursor.execute(f'SELECT id FROM {letter["type"]} WHERE number = ? AND date = ?', (letter["number"],date))
        letterId = self.cursor.fetchone()[0]
        keywordsIds = []
        # get keywords from input
        keywords = letter["keywords"].split(',')
        for word in keywords:
            word = word.strip()
            # foreach word check if it exists
            if len(word) > 0:
                self.cursor.execute(f'SELECT id FROM letter_keywords WHERE keyword = ?', (word,))
                wordid = self.cursor.fetchone()
                if not wordid:
                    # if keyword doesn't exist insert it into letter_keywords table
                    self.cursor.execute('INSERT INTO letter_keywords (keyword) VALUES (?)', (word,))
                    self.keyWordsOptions.append(word)
                    self.conn.commit()
                self.cursor.execute(f'SELECT id FROM letter_keywords WHERE keyword = ?', (word,))
                wordid = self.cursor.fetchone()[0]
                # insert relationship between keyword and letter
                self.cursor.execute(f'SELECT keywordid FROM {letter["type"][:-1]}_keywords WHERE keywordid = ? AND letterid = ?', (wordid,letterId))
                exists = self.cursor.fetchone()
                if not exists:
                    print("registering letter keywords")
                    self.cursor.execute(f'INSERT INTO {letter["type"][:-1]}_keywords (letterid, keywordid) VALUES(?, ?)', (letterId,wordid))
        self.conn.commit()
        
        nestedDir = f"./{self.letterTypeSelect.get()}/{letter['adressee']}/{str(letter['year'])}"
        os.makedirs(nestedDir, exist_ok=True)
        newFilename = f"{str(letter["number"])}{self.letterTypeSelect.get()} {str(letter["day"])}-{str(letter["month"])}-{str(letter["year"])} -- {letter["adressee"]}.pdf"
        shutil.copy(self.filename, f"{nestedDir}/{newFilename}")
        messagebox.showinfo("Success", f"تم إضافة الخطاب بنجاح \n {newFilename}")
        self.clearData()
        # except Exception as e:
        #     messagebox.showerror("Error", f"An error occurred: {e}")
    
    def clearData(self):
        self.getNumber()
        self.adresseeEntry.delete(0, 'end')
        self.keywordsEntry.delete(0, 'end')
        if (self.pdfView):
            self.pdfView.destroy()
        self.adresseeSelect.delete(0,'end')
        for item in self.adresseeOptions:
            self.adresseeSelect.insert('end', item)
        self.keyWordsSelect.delete(0,'end')
        for item in self.keyWordsOptions:
            self.keyWordsSelect.insert('end', item)
        if self.order:
            self.order.delete("1.0", tk.END)

    def copy_text(self,event=None):
        try:
            text_widget.event_generate("<<Copy>>")
        except tk.TclError:
            pass  # Handle empty selection gracefully

    def paste_text(self,event=None):
        try:
            text_widget.event_generate("<<Paste>>")
        except tk.TclError:
            pass  # Handle empty clipboard gracefully

    def cut_text(self,event=None):
        try:
            text_widget.event_generate("<<Cut>>")
        except tk.TclError:
            pass  # Handle empty selection gracefully

# third window frame search
class search(tk.Frame): 
    keyWordsOptions = []
    adresseeOptions = []
    letterTypeSelect = None
    letterNumberEntry = None
    letterYear = None
    letterMonth = None
    letterDay = None
    adresseeSelect = None
    adresseeEntry = None
    keyWordsSelect = None
    letterTable = 'outgoing_letters'
    filename = None
    receivedFile = None
    saveReceivingFileBtn = None
    receivingMethod = None
    editLetterBtn = None
    deleteLetterBtn = None
    dayValues = {
        1: [i for i in range(1,32)],
        2: [i for i in range(1,30)],
        3: [i for i in range(1,32)],
        4: [i for i in range(1,31)],
        5: [i for i in range(1,32)],
        6: [i for i in range(1,31)],
        7: [i for i in range(1,32)],
        8: [i for i in range(1,32)],
        9: [i for i in range(1,31)],
        10: [i for i in range(1,32)],
        11: [i for i in range(1,31)],
        12: [i for i in range(1,32)]
    }

    def __init__(self, parent, controller):

        self.controller = controller

        self.conn = sqlite3.connect('archive.db')
        self.cursor = self.conn.cursor()

        self.cursor.execute('SELECT name FROM adressees')
        for i in self.cursor.fetchall():
            self.adresseeOptions.append(i[0])

        self.cursor.execute('SELECT keyword FROM letter_keywords')
        for i in self.cursor.fetchall():
            self.keyWordsOptions.append(i[0])
        
        tk.Frame.__init__(self, parent)
        
        self.logo = Image.open('logo.png')
        self.logo = self.logo.resize((int(915/8), int(667/8)))
        self.logo = ImageTk.PhotoImage(self.logo)
        img = tk.Label(self, image=self.logo)
        img.grid(row=0,column=5, sticky="nsew")

        title = ttk.Label(self, text ="بحث", font = LARGEFONT)
        title.grid(row = 0, column = 2, padx = 10, pady = 10, sticky="nsew")

        addLetterPageBtn = ttk.Button(self, text ="إضافة خطاب 📩", command = lambda : controller.show_frame(addLetter))
        addLetterPageBtn.grid(row = 8, column = 2, padx = 10, pady = 10, sticky="nsew")

        self.previewPlaceholder = tk.Canvas(self, width=500, height=700)
        self.previewPlaceholder.grid(row = 0, column = 0, rowspan=8, sticky='nsew')
        self.previewPlaceholder.config(state="disabled")

        homeBtn = ttk.Button(self, text ="الصفحة الرئيسية 🏠", command = lambda : controller.show_frame(StartPage))
        homeBtn.grid(row = 8, column = 1, padx = 10, pady = 10, sticky="nsew")

        # ----------------
        letterTypeLabel = ttk.Label(self, text ="صادر / وارد")
        letterTypeLabel.grid(row = 2, column = 5, padx = 10, pady = 10, sticky="nsew")

        self.letterTypeSelect = ttk.Combobox(self, values=['صادر', 'وارد'])
        self.letterTypeSelect.current(0)
        self.letterTypeSelect.grid(row = 2, column = 3, padx = 10, pady = 10, sticky="nsew")
        self.letterTypeSelect.bind('<<ComboboxSelected>>', self.selectTable)
        
        letterNumberLabel = ttk.Label(self, text ="رقم الخطاب")
        letterNumberLabel.grid(row = 2, column = 2, padx = 10, pady = 10, sticky="nsew")

        self.letterNumberEntry = ttk.Entry(self)
        self.letterNumberEntry.grid(row = 2, column = 1, padx = 10, pady = 10, sticky="nsew")

        dateLabel = ttk.Label(self, text = 'التاريخ')
        dateLabel.grid(row = 3, column = 5, padx = 10, pady = 10, sticky="nsew")

        self.letterYear = ttk.Combobox(self, values=[i for i in range(1961,2100)])
        self.letterYear.grid(row = 3, column = 3, padx = 10, pady = 10, sticky="nsew")

        self.letterMonth = ttk.Combobox(self, values=[i for i in range(1,13)])
        self.letterMonth.grid(row = 3, column = 2, padx = 10, pady = 10, sticky="nsew")
        self.letterMonth.bind('<<ComboboxSelected>>', self.setDayValues)

        self.letterDay = ttk.Combobox(self, values=[i for i in range(1,32)])
        self.letterDay.grid(row = 3, column = 1, padx = 10, pady = 10, sticky="nsew")

        adresseeLabel = ttk.Label(self, text ="الجهة")
        adresseeLabel.grid(row = 4, column = 5, padx = 10, pady = 10, sticky="nsew")

        self.adresseeEntry = ttk.Entry(self)
        self.adresseeEntry.grid(row = 4, column = 2, columnspan=3, padx = 10, pady = 10, sticky="nsew")
        self.adresseeEntry.bind('<KeyRelease>',self.updateadresseeOptions)

        self.adresseeSelect = tk.Listbox(self, height=5)
        self.adresseeSelect.grid(row=5,column=2, columnspan=3,padx=10,pady=10, sticky="nsew")
        for item in self.adresseeOptions:
            self.adresseeSelect.insert('end', item)
        self.adresseeSelect.bind('<<ListboxSelect>>',self.autocompleteAdressee)
        adressee_scrollbar = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.adresseeSelect.yview)
        adressee_scrollbar.grid(row=5, column=4, sticky="nse")
        self.adresseeSelect.configure(yscrollcommand=adressee_scrollbar.set)

        keywordsLabel = ttk.Label(self, text ="كلمات البحث \n (يفصل بينها بفاصلة)")
        keywordsLabel.grid(row = 6, column = 5, padx = 10, pady = 10, sticky="nsew")

        self.keywordsEntry = ttk.Entry(self)
        self.keywordsEntry.grid(row=6, column=2, columnspan=3, padx = 10, pady = 10, sticky="nsew")
        self.keywordsEntry.bind('<KeyRelease>',self.updatekeywordOptions)

        self.keyWordsSelect = tk.Listbox(self, height=5)
        self.keyWordsSelect.grid(row=7,column=2, columnspan=3,padx=10,pady=10, sticky="nsew")
        for item in self.keyWordsOptions:
            self.keyWordsSelect.insert('end', item)
        self.keyWordsSelect.bind('<<ListboxSelect>>',self.autocompleteKeywords)
        keywords_scrollbar = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.keyWordsSelect.yview)
        keywords_scrollbar.grid(row=7, column=4, sticky="nse")
        self.keyWordsSelect.configure(yscrollcommand=keywords_scrollbar.set)
        
        submit = ttk.Button(self,text="بحث 🔍",command=self.searchData)
        submit.grid(row=8,column=5,padx=10,pady=10, sticky="nsew")

        self.trv = ttk.Treeview(self, selectmode='browse', columns=("letter_number", "letter_date", "adressee","keywords", "order"), show='tree headings', height=4)
        self.trv.grid(row=9, column=0, columnspan=5, padx=20, pady=20, sticky="nsew")
        trv_scrollbar = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.trv.yview)
        trv_scrollbar.grid(row=9, column=4,padx=20, pady=20, sticky="nse")
        self.trv.configure(yscrollcommand=trv_scrollbar.set)
        self.trv.tag_configure('gray', background='lightgray')
        self.trv.tag_configure('normal', background='white')
        self.selectTable(None)
        # ----------------

    def setDayValues(self,event):
        self.letterDay.configure(values=self.dayValues[int(event.widget.get())])
        self.letterDay.set("")

    def updatekeywordOptions(self, event):
        value = event.widget.get()
        
        # get data from l
        if value == '':
            data = self.keyWordsOptions
        else:
            data = []
            for item in self.keyWordsOptions:
                if value.lower() in item.lower():
                    data.append(item)                
    
        # update data in listbox
        self.updateKeywordsData(data)
    
    def updateKeywordsData(self,data):
        
        # clear previous data
        self.keyWordsSelect.delete(0, 'end')
    
        # put new data
        for item in data:
            self.keyWordsSelect.insert('end', item)
    
    def autocompleteKeywords(self,event):
        listbox = event.widget
        selected_index = listbox.curselection()
        selected_item = [listbox.get(i) for i in selected_index]
        if selected_item:
            self.keywordsEntry.delete(0,'end')
            self.keywordsEntry.insert('end',selected_item[0])
    
    def setLetterType(self, event):
        pass
    
    def updateadresseeOptions(self, event):
        value = event.widget.get()
        
        # get data from l
        if value == '':
            data = self.adresseeOptions
        else:
            data = []
            for item in self.adresseeOptions:
                if value.lower() in item.lower():
                    data.append(item) 
    
        # update data in listbox
        self.updateAdresseeData(data)
    
    def updateAdresseeData(self,data):
        
        # clear previous data
        self.adresseeSelect.delete(0, 'end')
    
        # put new data
        for item in data:
            self.adresseeSelect.insert('end', item)
    
    def autocompleteAdressee(self, event):
        listbox = event.widget
        selected_index = listbox.curselection()
        selected_item = [listbox.get(i) for i in selected_index]
        if selected_item:
            self.adresseeEntry.delete(0,'end')
            self.adresseeEntry.insert('end',selected_item[0])
    
    def selectTable(self, event):
        if self.editLetterBtn:
            self.editLetterBtn.destroy()
        if self.deleteLetterBtn:
            self.deleteLetterBtn.destroy()
        self.letterTable = 'outgoing_letters' if self.letterTypeSelect.get()=='صادر' else 'incoming_letters'
        if self.letterTypeSelect.get().strip() == 'صادر':
            self.receivedBtn = ttk.Button(self, text ="التسليم 📬", command = self.addReceivingFile)
            self.receivedBtn.grid(row = 4, column = 1, padx = 10, pady = 10, sticky="nsew")
        else:
            self.receivedBtn.destroy()
            if self.receivingMethod:
                self.receivingMethod.destroy()
            if self.saveReceivingFileBtn:
                self.saveReceivingFileBtn.destroy()
            for row in self.trv.get_children():
                self.trv.delete(row)
    
    def addReceivingFile(self):
        if not self.trv.selection():
            return
        selected_item = self.trv.selection()
        if not selected_item:
            return
        selected_item = self.trv.item(selected_item[0], "values")
        nestedDir = f"./{self.letterTypeSelect.get()}/{selected_item[2]}/{str(selected_item[1][:4])}/تسليمات"
        self.receivedFile = f"تسليم {selected_item[0]}{self.letterTypeSelect.get()} {int(selected_item[1][-2:])}-{int(selected_item[1][5:7])}-{int(selected_item[1][:4])} -- {selected_item[2]}.pdf"
        if self.receivedFile and os.path.exists(f'{nestedDir}/{self.receivedFile}'):
            self.pdfView.destroy()
            self.pdfView = PDFViewerWidget.create_pdf_viewer(self,f'{nestedDir}/{self.receivedFile}', False)
            self.pdfView.grid(row = 0, column = 0, rowspan=8, sticky="nsew")
            return
        self.receivingMethod = ttk.Combobox(self, values=['email', 'يدوي'])
        self.receivingMethod.current(0)
        self.receivingMethod.grid(row = 5, column = 1, padx = 10, pady = 10, sticky="new")
        self.receivingMethod.bind('<<ComboboxSelected>>', self.selectFile)
    
    def selectFile(self, event):
        if self.receivingMethod.get().strip() != 'يدوي':
            if self.saveReceivingFileBtn:
                self.saveReceivingFileBtn.destroy()
            return
        self.saveReceivingFileBtn = ttk.Button(self, text ="حفظ صورة التسليم 📬", command = self.saveReceivingFile)
        self.saveReceivingFileBtn.grid(row = 6, column = 1, padx = 10, pady = 10, sticky="new")
        filetypes = (
            ('pdf files', '*.pdf'),
            ('All files', '*.*')
        )

        self.filename = filedialog.askopenfilename(
            title='Open a file',
            initialdir='/',
            filetypes=filetypes)

        try:
            if self.filename[-3:] != 'pdf':
                messagebox.showerror("Error", f"اختر ملف من نوع PDF")
                return
            self.pdfView = PDFViewerWidget.create_pdf_viewer(self,self.filename)
            self.pdfView.grid(row = 0, column = 0, rowspan=8, sticky="nsew")
        except ValueError:
            messagebox.showerror("Error", f"اختر ملف من نوع PDF")
            return
    
    def saveReceivingFile(self):
        if not(self.filename and os.path.exists(self.filename)):
            messagebox.showerror("Error", f"اختر ملف من نوع PDF")
            return
        selected_item = self.trv.selection()[0]
        selected_item = self.trv.item(selected_item, "values")
        nestedDir = f"./{self.letterTypeSelect.get()}/{selected_item[2]}/{str(selected_item[1][:4])}/تسليمات"
        os.makedirs(nestedDir, exist_ok=True)
        shutil.copy(self.filename, f"{nestedDir}/{self.receivedFile}")
        date = int((datetime.strptime(f'{selected_item[1]}', "%Y-%m-%d")-datetime(1970,1,1)).total_seconds())
        self.cursor.execute(f'UPDATE outgoing_letters SET received=1 WHERE number = {selected_item[0]} AND date = {date}')
        self.conn.commit()
        messagebox.showinfo("Success", f"تم إضافة صورة التسليم بنجاح \n {self.receivedFile}")

    def searchData(self):
        if self.saveReceivingFileBtn:
            self.saveReceivingFileBtn.destroy()
        if self.receivingMethod:
            self.receivingMethod.destroy()
        letter = {
            'number' : self.letterNumberEntry.get().strip(),
            'year' : self.letterYear.get().strip(),
            'month' : self.letterMonth.get().strip(),
            'day' : self.letterDay.get().strip(),
            'adressee' : self.adresseeEntry.get().strip(),
            'keywords' : self.keywordsEntry.get().strip()
        }
        field = ''
        if self.letterTable == 'incoming_letters':
            field += ', "order"'
        query = f"""SELECT number, date, adressees.name AS adressee, GROUP_CONCAT(IFNULL(letter_keywords.keyword, 'لا توجد كلمات بحث')) AS keywords {field}
                    FROM {self.letterTable} LEFT JOIN adressees ON {self.letterTable}.adressee = adressees.id 
                    LEFT JOIN {self.letterTable[:-1]}_keywords ON {self.letterTable}.id = {self.letterTable[:-1]}_keywords.letterid LEFT JOIN letter_keywords ON {self.letterTable[:-1]}_keywords.keywordid = letter_keywords.id """
        conditions = ""

        if letter['adressee']:
            conditions += f"adressees.name LIKE '%{letter['adressee']}%' AND "
        if letter['keywords']:
            conditions += f"""{self.letterTable}.id IN (
                                SELECT {self.letterTable[:-1]}_keywords.letterid
                                FROM {self.letterTable[:-1]}_keywords
                                JOIN letter_keywords ON {self.letterTable[:-1]}_keywords.keywordid = letter_keywords.id
                                WHERE letter_keywords.keyword LIKE '%{letter['keywords']}%'
                            ) AND """
        if letter['number']:
            conditions += f"number = '{str(letter['number'])}' AND "
        if letter['year']:
            if not (letter['month'] or letter['day']):
                start_date = int((datetime.strptime(f'{int(letter["year"])}-1-1', "%Y-%m-%d")-datetime(1970,1,1)).total_seconds())
                end_date = int((datetime.strptime(f'{int(letter["year"])}-12-31', "%Y-%m-%d")-datetime(1970,1,1)).total_seconds())
                conditions += f"date >= {start_date} AND date <= {end_date} AND "
            elif not letter['day']:
                start_date = int((datetime.strptime(f'{int(letter["year"])}-{int(letter["month"])}-1', "%Y-%m-%d")-datetime(1970,1,1)).total_seconds())
                if letter["month"] == 12:  # Handle December separately
                    end_date = datetime(int(letter["year"]) + 1, 1, 1) - timedelta(seconds=1)
                else:
                    end_date = datetime(int(letter["year"]), int(letter["month"]) + 1, 1) - timedelta(seconds=1)
                end_date = int((end_date-datetime(1970,1,1)).total_seconds())
                conditions += f"date >= {start_date} AND date <= {end_date} AND "
            else:
                date =  int((datetime.strptime(f'{int(letter["year"])}-{int(letter["month"])}-{int(letter["day"])}', "%Y-%m-%d")-datetime(1970,1,1)).total_seconds())
                conditions += f"date = {date} AND "

        if conditions :
            query += f" WHERE {conditions[:-4]} "
        query += f" GROUP BY {self.letterTable}.number "
        self.cursor.execute(query)
        print(query)
        result = self.cursor.fetchall()
        for row in self.trv.get_children():
            self.trv.delete(row)
        if not result:
            self.trv.insert("", "end", values=("","","لم يتم العثور على أي بيانات",))
            self.trv.tag_configure("nodata", foreground="red")
            self.trv.item(self.trv.get_children()[0], tags=("nodata",))
            return
        rows_numbers = [i for i in range(0,len(result)+1)]
        
        # width of columns and alignment
        self.trv.column("#0", width=20)
        self.trv.column("letter_number", width=30, anchor='c')
        self.trv.column("letter_date", width=80, anchor='c')
        self.trv.column("order", width=150, anchor='c')
        self.trv.column("adressee", width=250, anchor='c')
        self.trv.column("keywords", width=250, anchor='c')

        # Headings
        # respective columns
        self.trv.heading("#0", text="م")
        self.trv.heading("letter_number", text="رقم الخطاب")
        self.trv.heading("letter_date", text="التاريخ")
        self.trv.heading("order", text="التأشيرة")
        self.trv.heading("adressee", text="الجهة")
        self.trv.heading("keywords", text="كلمات البحث")
        my_tag = 'normal'
        i = 1
        for record in result: 
            my_tag='gray' if my_tag=='normal' else 'normal'
            if(self.letterTable == 'incoming_letters'):
                self.trv.insert("", 'end', iid=rows_numbers[i], text=rows_numbers[i],
                    values=(record[0], (datetime(1970, 1, 1) + timedelta(seconds=record[1])).strftime("%Y-%m-%d"), record[2], record[3], (record[4] if record[4] else '')), tags=(my_tag))
            else:
                self.trv.insert("", 'end', iid=rows_numbers[i], text=rows_numbers[i],
                    values=(record[0], (datetime(1970, 1, 1) + timedelta(seconds=record[1])).strftime("%Y-%m-%d"), record[2], record[3]), tags=(my_tag))
            i+=1
        self.trv.bind("<<TreeviewSelect>>", self.showSelection)
        # self.cursor.execute()

    def showSelection(self,event):
        if self.saveReceivingFileBtn:
            self.saveReceivingFileBtn.destroy()
        if self.receivingMethod:
            self.receivingMethod.destroy()
        selected_item = event.widget.selection()
        if not selected_item:
            return
        selected_item = event.widget.item(selected_item[0], "values")
        nestedDir = f"./{self.letterTypeSelect.get()}/{selected_item[2]}/{str(selected_item[1][:4])}"
        filename = f"{selected_item[0]}{self.letterTypeSelect.get()} {int(selected_item[1][-2:])}-{int(selected_item[1][5:7])}-{int(selected_item[1][:4])} -- {selected_item[2]}.pdf"
        filepath = f"{nestedDir}/{filename}"

        self.pdfView = PDFViewerWidget.create_pdf_viewer(self,filepath, False)
        self.pdfView.grid(row = 0, column = 0, rowspan=8, sticky="nsew")
        
        self.deleteLetterBtn = ttk.Button(self, text ="حذف الخطاب ☒", command = self.deleteLetter)
        self.deleteLetterBtn.grid(row = 8, column = 3, padx = 10, pady = 10, sticky="nsew")
        date = int((datetime.strptime(selected_item[1], "%Y-%m-%d")-datetime(1970,1,1)).total_seconds())
        query = f'SELECT id FROM {self.letterTable} WHERE number = {selected_item[0]} AND date = {date}'
        print(query)
        self.cursor.execute(query)
        self.letterid = self.cursor.fetchone()
        print(self.letterid)
        self.letterid = self.letterid[0]

        self.editLetterBtn = ttk.Button(self, text ="تعديل البيانات 📝", command = self.editLetter)
        self.editLetterBtn.grid(row = 8, column = 4, padx = 10, pady = 10, sticky="nsew")
    
    def deleteLetter(self):
        selected_item = self.trv.selection()
        if not selected_item:
            return
        selected_item = self.trv.item(selected_item[0], "values")
        date = int((datetime.strptime(f'{selected_item[1]}', "%Y-%m-%d")-datetime(1970,1,1)).total_seconds())
        delete_confirm = messagebox.askyesno("تأكيد الحذف","هل تريد فعلاً حذف الخطاب المحدد؟ سيتم حذف بياناته من قاعدة البيانات وكذلك حذف الملف من الجهاز ")
        if delete_confirm:
            self.cursor.execute(f'DELETE FROM {self.letterTable} WHERE number = {selected_item[0]} and date = {date}')
            self.conn.commit()
            nestedDir = f"./{self.letterTypeSelect.get()}/{selected_item[2]}/{str(selected_item[1][:4])}"
            filename = f"{selected_item[0]}{self.letterTypeSelect.get()} {int(selected_item[1][-2:])}-{int(selected_item[1][5:7])}-{int(selected_item[1][:4])} -- {selected_item[2]}.pdf"
            filepath = f"{nestedDir}/{filename}"
            self.pdfView.release_pdf()
            self.pdfView.destroy()
            os.remove(filepath)
            messagebox.showinfo("تم الحذف", "تم حذف الخطاب ")
            self.searchData()
            self.deleteLetterBtn.destroy()
            if self.editLetterBtn:
                self.editLetterBtn.destroy()

    def editLetter(self):
        self.controller.set_data('table', self.letterTable)
        self.controller.set_data('letterid', self.letterid)
        self.controller.show_frame(editLetter,table=self.letterTable,letterid=self.letterid)

    def refreshKeywords(self):
        self.keyWordsOptions.clear()
        self.cursor.execute('SELECT keyword FROM letter_keywords')
        for row in self.cursor.fetchall():
            self.keyWordsOptions.append(row[0])
        
        # Update Listbox Data
        self.keyWordsSelect.delete(0, 'end')
        for item in self.keyWordsOptions:
            self.keyWordsSelect.insert('end', item)
    
    def refreshAdressee(self):
        self.adresseeOptions.clear()
        self.cursor.execute('SELECT name FROM adressees')
        for row in self.cursor.fetchall():
            self.adresseeOptions.append(row[0])
        
        # Update Listbox Data
        self.adresseeSelect.delete(0, 'end')
        for item in self.adresseeOptions:
            self.adresseeSelect.insert('end', item)

    def wrap_text(text, width=20):
        return "\n".join(textwrap.wrap(text, width))

# fourth window frame editLetter 
class editLetter(tk.Frame):
    keyWordsOptions = []
    adresseeOptions = []
    letterTypeSelect = None
    letterNumber = None
    letterYear = None
    letterMonth = None
    letterDay = None
    adresseeSelect = None
    adresseeEntry = None
    keyWordsSelect = None
    letterTable = 'outgoing_letters'
    filename = None
    dayValues = {
        1: [i for i in range(1,32)],
        2: [i for i in range(1,30)],
        3: [i for i in range(1,32)],
        4: [i for i in range(1,31)],
        5: [i for i in range(1,32)],
        6: [i for i in range(1,31)],
        7: [i for i in range(1,32)],
        8: [i for i in range(1,32)],
        9: [i for i in range(1,31)],
        10: [i for i in range(1,32)],
        11: [i for i in range(1,31)],
        12: [i for i in range(1,32)]
    }
    
    def __init__(self, parent, controller):

        self.controller = controller
        self.conn = sqlite3.connect('archive.db')
        self.cursor = self.conn.cursor()

        self.cursor.execute('SELECT name FROM adressees')
        for i in self.cursor.fetchall():
            self.adresseeOptions.append(i[0])

        self.cursor.execute('SELECT keyword FROM letter_keywords')
        for i in self.cursor.fetchall():
            self.keyWordsOptions.append(i[0])

        tk.Frame.__init__(self, parent)

        self.logo = Image.open('logo.png')
        self.logo = self.logo.resize((int(915/8), int(667/8)))
        self.logo = ImageTk.PhotoImage(self.logo)
        img = tk.Label(self, image=self.logo)
        img.grid(row=0,column=5, sticky="nsew")

        title = ttk.Label(self, text ="تعديل خطاب", font = LARGEFONT)
        title.grid(row = 0, column = 3, padx = 10, pady = 10, sticky="nsew")
 
        searchPageBtn = ttk.Button(self, text ="بحث 🔍",command = lambda : self.controller.show_frame(search))
        searchPageBtn.grid(row = 9, column = 3, padx = 10, pady = 10, sticky="nsew")
        
        homeBtn = ttk.Button(self, text ="الصفحة الرئيسية 🏠", command = lambda : self.controller.show_frame(StartPage))
        homeBtn.grid(row = 9, column = 2, padx = 10, pady = 10, sticky="nsew")

        self.previewPlaceholder = tk.Canvas(self, width=500, height=700)
        self.previewPlaceholder.grid(row = 0, column = 0, rowspan=8, sticky="nsew")
        self.previewPlaceholder.config(state="disabled")

        openFileBtn = ttk.Button(self, text ="اختر ملف 📄", command = self.selectFile)
        openFileBtn.grid(row = 1, column = 5, padx = 10, pady = 10, sticky="nsew")

        letterTypeLabel = ttk.Label(self, text ="صادر / وارد")
        letterTypeLabel.grid(row = 2, column = 5, padx = 10, pady = 10, sticky="nsew")

        self.letterTypeSelect = ttk.Combobox(self, values=['صادر', 'وارد'])
        self.letterTypeSelect.current(0)
        self.letterTypeSelect.grid(row = 2, column = 3, padx = 10, pady = 10, sticky="nsew")
        self.letterTypeSelect.state(["disabled"])
        
        letterNumberLabel = ttk.Label(self, text ="رقم الخطاب")
        letterNumberLabel.grid(row = 1, column = 3, padx = 10, pady = 10, sticky="nsew")

        self.letterNumberVar = tk.StringVar()
        self.letterNumber = tk.Entry(self,  textvariable=self.letterNumberVar)
        self.letterNumber.grid(row = 1, column = 2, padx = 10, pady = 10, sticky="nsew")

        dateLabel = ttk.Label(self, text = 'التاريخ')
        dateLabel.grid(row = 3, column = 5, padx = 10, pady = 10)

        self.combo_var = tk.StringVar()
        current_datetime = datetime.now()
        today = current_datetime.strftime("%Y-%m-%d")
        self.combo_var.set(today[:4])
        self.letterYear = ttk.Combobox(self, values=[i for i in range(1961,2100)],textvariable=self.combo_var)
        self.letterYear.grid(row = 3, column = 3, padx = 10, pady = 10, sticky="nsew")

        self.letterMonth = ttk.Combobox(self, values=[i for i in range(1,13)])
        self.letterMonth.grid(row = 3, column = 2, padx = 10, pady = 10, sticky="nsew")
        self.letterMonth.set(int(today[5:7]))
        self.letterMonth.bind('<<ComboboxSelected>>', self.setDayValues)

        self.letterDay = ttk.Combobox(self, values=[i for i in range(1,32)])
        self.letterDay.grid(row = 3, column = 1, padx = 10, pady = 10, sticky="nsew")
        self.letterDay.set(int(today[-2:]))

        adresseeLabel = ttk.Label(self, text ="الجهة")
        adresseeLabel.grid(row = 4, column = 5, padx = 10, pady = 10, sticky="nsew")

        self.adresseeEntry = ttk.Entry(self)
        self.adresseeEntry.grid(row = 4, column = 1, columnspan=4, padx = 10, pady = 10, sticky="nsew")
        self.adresseeEntry.bind('<KeyRelease>',self.updateadresseeOptions)

        self.adresseeSelect = tk.Listbox(self, height=5)
        self.adresseeSelect.grid(row=5,column=1, columnspan=4,padx=10,pady=10, sticky="nsew")
        for item in self.adresseeOptions:
            self.adresseeSelect.insert('end', item)
        self.adresseeSelect.bind('<<ListboxSelect>>',self.autocompleteAdressee)
        adressee_scrollbar = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.adresseeSelect.yview)
        adressee_scrollbar.grid(row=5, column=4, sticky="nse")
        self.adresseeSelect.configure(yscrollcommand=adressee_scrollbar.set)

        keywordsLabel = ttk.Label(self, text ="كلمات البحث \n (يفصل بينها بفاصلة)")
        keywordsLabel.grid(row = 6, column = 5, padx = 10, pady = 10, sticky="nsew")

        self.keywordsEntry = ttk.Entry(self)
        self.keywordsEntry.grid(row=6, column=1, columnspan=4, padx = 10, pady = 10, sticky="nsew")
        self.keywordsEntry.bind('<KeyRelease>',self.updatekeywordOptions)

        self.keyWordsSelect = tk.Listbox(self, height=5)
        self.keyWordsSelect.grid(row=7,column=1, columnspan=4,padx=10,pady=10, sticky="nsew")
        for item in self.keyWordsOptions:
            self.keyWordsSelect.insert('end', item)
        self.keyWordsSelect.bind('<<ListboxSelect>>',self.autocompleteKeywords)
        keywords_scrollbar = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.keyWordsSelect.yview)
        keywords_scrollbar.grid(row=7, column=4, sticky="nse")
        self.keyWordsSelect.configure(yscrollcommand=keywords_scrollbar.set)

        order_placeholder = ttk.Label(self)
        order_placeholder.grid(row=8,column=1,pady=50)

        submit = ttk.Button(self,text="حفظ 💾",command=self.saveData)
        submit.grid(row=9,column=5,padx=10,pady=10, sticky="nsew")

    def getOldValues(self,table,letterid):
        self.letterid = letterid
        self.letterTable = table
        if not table:
            print('letterTable is None')
            return
        self.letterTable = table
        field = ', incoming_letters."order"' if table == 'incoming_letters' else ''
        query = f"""SELECT 
                    {table}.id,
                    {table}.number,
                    {table}.date,
                    adressees.name AS adressee,
                    GROUP_CONCAT(IFNULL(letter_keywords.keyword, 'لا توجد كلمات بحث')) AS keywords {field}
                FROM {table}
                LEFT JOIN adressees ON {table}.adressee = adressees.id
                LEFT JOIN {table[:-1]}_keywords ON {table}.id = {table[:-1]}_keywords.letterid
                LEFT JOIN letter_keywords ON {table[:-1]}_keywords.keywordid = letter_keywords.id
                WHERE {table}.id = {letterid}
                GROUP BY {table}.id, {table}.number, {table}.date, adressees.name
                """
        print(query)
        self.cursor.execute(query)
        data = self.cursor.fetchone()
        print(data)
        lt = 0 if table == 'outgoing_letters' else 1
        self.letterTypeSelect.current(lt)
        self.letterNumberVar.set(data[1])
        date = (datetime(1970, 1, 1) + timedelta(seconds=data[2])).strftime("%Y-%m-%d")
        self.letterYear.set(int(date[:4]))
        self.letterMonth.set(int(date[5:7]))
        self.letterDay.set(int(date[-2:]))
        self.adresseeEntry.delete(0,'end')
        self.adresseeEntry.insert(0, data[3])
        self.keywordsEntry.delete(0,'end')
        if data[4] != 'لا توجد كلمات بحث':
            self.keywordsEntry.insert(0, data[4]+', ')
        nestedDir = f"./{self.letterTypeSelect.get()}/{data[3]}/{str(date[:4])}"
        filename = f"{data[1]}{self.letterTypeSelect.get()} {int(date[-2:])}-{int(date[5:7])}-{int(date[:4])} -- {data[3]}.pdf"
        self.oldfilepath = f"{nestedDir}/{filename}"
        self.pdfView = PDFViewerWidget.create_pdf_viewer(self,self.oldfilepath,False)
        self.pdfView.grid(row = 0, column = 0, rowspan=8, sticky="nsew")
        if lt :
            self.order = tk.Text(self,height=5, width=10)
            self.order.grid(row=8,column=1, columnspan=4,padx=10,pady=10,sticky="nsew")
            self.text_scrollbar = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.order.yview)
            self.text_scrollbar.grid(row=8, column=4, sticky="nse")
            self.order.configure(yscrollcommand=self.text_scrollbar.set)
            self.order.bind("<Control-c>", self.copy_text)
            self.order.bind("<Control-v>", self.paste_text)
            self.order.bind("<Control-x>", self.cut_text)
            self.order.insert("1.0", data[4])

    def setDayValues(self,event):
        self.letterDay.configure(values=self.dayValues[int(event.widget.get())])
        self.letterDay.set("")

    def updatekeywordOptions(self, event):
        value = event.widget.get().split(',')[-1].strip()
        
        # get data from l
        if value == '':
            data = self.keyWordsOptions
        else:
            data = []
            for item in self.keyWordsOptions:
                if value.lower() in item.lower():
                    data.append(item)                
    
        # update data in listbox
        self.updateKeywordsData(data)
    
    def updateKeywordsData(self,data):
        
        # clear previous data
        self.keyWordsSelect.delete(0, 'end')
    
        # put new data
        for item in data:
            self.keyWordsSelect.insert('end', item)
    
    def autocompleteKeywords(self,event):
        listbox = event.widget
        selected_index = listbox.curselection()
        selected_item = [listbox.get(i) for i in selected_index]
        if selected_item:
            keywords = self.keywordsEntry.get().split(',')
            keywords[-1] = selected_item[0]
            self.keywordsEntry.delete(0,'end')
            self.keywordsEntry.insert('end', ','.join(keywords)+',')
    
    def selectTable(self,event):
        if self.letterTypeSelect.get()=='صادر':
            self.letterTable = 'outgoing_letters' 
        elif self.letterTypeSelect.get()=='وارد':
            self.letterTable = 'incoming_letters'
        if self.letterTable == 'incoming_letters':
            self.orderLabel = ttk.Label(self, text ="التأشيرة")
            self.orderLabel.grid(row = 8, column = 5, columnspan=2, padx = 10, pady = 10, sticky="nsew")

            self.order = tk.Text(self,height=5, width=10)
            self.order.grid(row=8,column=1, columnspan=4,padx=10,pady=10,sticky="nsew")
            self.text_scrollbar = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.order.yview)
            self.text_scrollbar.grid(row=8, column=4, sticky="nse")
            self.order.configure(yscrollcommand=self.text_scrollbar.set)
            self.order.bind("<Control-c>", self.copy_text)
            self.order.bind("<Control-v>", self.paste_text)
            self.order.bind("<Control-x>", self.cut_text)
        else:
            try:
                self.orderLabel.destroy()
                self.order.destroy()
                self.text_scrollbar.destroy()
            except:
                pass

    def updateadresseeOptions(self, event):
        value = event.widget.get()
        
        # get data from l
        if value == '':
            data = self.adresseeOptions
        else:
            data = []
            for item in self.adresseeOptions:
                if value.lower() in item.lower():
                    data.append(item) 
        # update data in listbox
        self.updateAdresseeData(data)
    
    def updateAdresseeData(self,data):
        
        # clear previous data
        self.adresseeSelect.delete(0, 'end')
    
        # put new data
        for item in data:
            self.adresseeSelect.insert('end', item)
    
    def autocompleteAdressee(self, event):
        listbox = event.widget
        selected_index = listbox.curselection()
        selected_item = [listbox.get(i) for i in selected_index]
        if selected_item:
            self.adresseeEntry.delete(0,'end')
            self.adresseeEntry.insert('end',selected_item[0])
    
    def selectFile(self):
        filetypes = (
            ('pdf files', '*.pdf'),
            ('All files', '*.*')
        )

        self.filename = filedialog.askopenfilename(
            title='Open a file',
            initialdir='/',
            filetypes=filetypes)

        try:
            if self.filename[-3:] != 'pdf':
                messagebox.showerror("Error", f"اختر ملف من نوع PDF")
                return
            self.pdfView = PDFViewerWidget.create_pdf_viewer(self,self.filename)
            self.pdfView.grid(row = 0, column = 0, rowspan=8, sticky="nsew")
        except ValueError:
            messagebox.showerror("Error", f"اختر ملف من نوع PDF")
            return
    
    def saveData(self):
        order = ''
        if self.letterTable == 'incoming_letters':
            order = self.order.get("1.0", "end-1c").strip()
        letter = {
            'type' : self.letterTable,
            'number' : self.letterNumber.get().strip(),
            'year' : self.combo_var.get().strip(),
            'month' : self.letterMonth.get().strip(),
            'day' : self.letterDay.get().strip(),
            'adressee' : self.adresseeEntry.get().strip(),
            'keywords' : self.keywordsEntry.get().strip(),
            'order' : order
        }
        for (key,value) in letter.items():
            if not value or not self.oldfilepath:
                if key == 'order':
                    continue
                data = f"""
                        نوع الخطاب: {'صادر' if letter['type'] == 'outgoing_letters' else 'وارد' if letter['type'] == 'incoming_letters' else letter['type']} 
                        رقم الخطاب: {letter['number'] if letter['number'] else 'لا بيانات'}
                        تاريخ الخطاب: {letter['year'] +'/'+ letter['month'] +'/'+ letter['day'] if letter['year'] and letter['month'] and letter['day'] else 'غير مكتمل'}
                        الجهة: {letter['adressee'] if letter['adressee'] else 'لا بيانات'}
                        كلمات البحث: {letter['keywords'] if letter['keywords'] else 'لا بيانات'}
                        ملف الخطاب: {self.oldfilepath if self.oldfilepath else 'لا بيانات'}
                        """
                messagebox.showwarning("warning", f"البيانات غير مكتملة \n {data}")
                return
        # try:
        self.cursor.execute('SELECT id FROM adressees WHERE name = ?', (letter["adressee"],))
        adresseeId = self.cursor.fetchone()
        if not adresseeId:
            self.cursor.execute('INSERT INTO adressees (name) VALUES (?)', (letter["adressee"],))
            self.conn.commit()
            self.adresseeOptions.append(letter["adressee"])
            self.cursor.execute('SELECT id FROM adressees WHERE name = ?', (letter["adressee"],))
            adresseeId = self.cursor.fetchone()
        try:
            if self.letterTypeSelect.get() not in ('صادر','وارد'):
                raise Exception('نوع الخطاب يجب أن يكون "صادر" أو "وارد"')
            letter['number'] = int(letter['number'])
            try:
                letter['year'] = int(letter['year'])
            except ValueError:
                raise Exception('السنة يجب أن تكون رقم')
            try:
                letter['month'] = int(letter['month'])
            except ValueError:
                raise Exception('الشهر يجب أن يكون رقم')
            try:
                letter['day'] = int(letter['day'])
            except ValueError:
                raise Exception('اليوم يجب أن يكون رقم')
            if letter['year']<1961:
                raise Exception('السنة يجب أن تكون أكبر من 1960 \n أنشأت وزارة التعليم العالى عام 1961 وفقا للقرار الجمهورى رقم 1665 لعام 1961')
            if letter['month']<1 or letter['month'] > 12:
                raise Exception('الشهر يجب أن يكون بين 1 و 12')
            if letter['day']<1 or letter['day'] > 31:
                raise Exception('اليوم يجب أن يكون بين 1 و 31')
        except Exception as e:
            messagebox.showerror("Error", f"البيانات غير صحيحة \n {e}")
            return
        date = int((datetime.strptime(f'{letter["year"]}-{letter["month"]}-{letter["day"]}', "%Y-%m-%d")-datetime(1970,1,1)).total_seconds())
        query = {
            'query' : f'UPDATE {letter["type"]} SET number = ?, date = ?, adressee=? WHERE id = {self.letterid}',
            'values' : (letter["number"], date, adresseeId[0],)
        }
        if letter["order"]:
            query = {
                'query' : f'UPDATE {letter["type"]} SET number = ?, date = ?, "order"=?, adressee=? WHERE id = {self.letterid}',
                'values' : (letter["number"], date, letter["order"], adresseeId[0],)
            }
        print("registered letter data")
        # insert letter data into letter table
        self.cursor.execute(query['query'], query['values'])
        self.cursor.execute(f'DELETE FROM {letter["type"][:-1]}_keywords WHERE letterid = ?', (self.letterid,))
        self.conn.commit()
        keywordsIds = []
        # get keywords from input
        keywords = letter["keywords"].split(',')
        for word in keywords:
            word = word.strip()
            # foreach word check if it exists
            if len(word) > 0:
                self.cursor.execute(f'SELECT id FROM letter_keywords WHERE keyword = ?', (word,))
                wordid = self.cursor.fetchone()
                if not wordid:
                    # if keyword doesn't exist insert it into letter_keywords table
                    self.cursor.execute('INSERT INTO letter_keywords (keyword) VALUES (?)', (word,))
                    self.keyWordsOptions.append(word)
                    self.conn.commit()
                self.cursor.execute(f'SELECT id FROM letter_keywords WHERE keyword = ?', (word,))
                wordid = self.cursor.fetchone()[0]
                # insert relationship between keyword and letter
                self.cursor.execute(f'SELECT keywordid FROM {letter["type"][:-1]}_keywords WHERE keywordid = ? AND letterid = ?', (wordid,self.letterid))
                exists = self.cursor.fetchone()
                if not exists:
                    print("registering letter keywords")
                    self.cursor.execute(f'INSERT INTO {letter["type"][:-1]}_keywords (letterid, keywordid) VALUES(?, ?)', (self.letterid,wordid))
        self.conn.commit()
        
        try:
            nestedDir = f"./{self.letterTypeSelect.get()}/{letter['adressee']}/{str(letter['year'])}"
            self.newfilepath = f'{nestedDir}/{letter['number']}{self.letterTypeSelect.get()} {int(letter['day'])}-{int(letter['month'])}-{int(letter['year'])} -- {letter['adressee']}.pdf'
            if not self.filename:
                self.filename = self.oldfilepath
            os.makedirs(nestedDir, exist_ok=True)
            if self.oldfilepath == self.newfilepath and self.filename:
                shutil.copy(self.filename, self.oldfilepath)
            else:
                shutil.copy(self.filename, self.newfilepath)
        except Exception as e:
            messagebox.showerror('Error',f'لا يمكن نسخ الملف \n {e}')
            print(e)
        messagebox.showinfo("Success", f"تم تعديل البيانات بنجاح \n {self.oldfilepath}")
        self.showUpdatedData()
        # except Exception as e:
        #     messagebox.showerror("Error", f"An error occurred: {e}")
    
    def showUpdatedData(self):
        self.controller.show_frame(search)

    def copy_text(self,event=None):
        try:
            text_widget.event_generate("<<Copy>>")
        except tk.TclError:
            pass  # Handle empty selection gracefully

    def paste_text(self,event=None):
        try:
            text_widget.event_generate("<<Paste>>")
        except tk.TclError:
            pass  # Handle empty clipboard gracefully

    def cut_text(self,event=None):
        try:
            text_widget.event_generate("<<Cut>>")
        except tk.TclError:
            pass  # Handle empty selection gracefully


if __name__ == "__main__":
    # Define database file name
    db_file = "archive.db"

    # Check if the database file exists, create if not
    if not os.path.exists(db_file):
        print("Database does not exist. Creating...")
        # Connect to the database (creates file if missing)
        conn = sqlite3.connect(db_file)
        cursor = conn.cursor()
        # Read and execute SQL script
        try:
            with open("db.sql", "r", encoding="utf-16") as file:
                sql_script = file.read()
            cursor.executescript(sql_script)  # Execute multiple SQL commands
            print("Database setup completed successfully!")
        except sqlite3.Error as e:
            print(f"Error executing SQL script: {e}")
        # Commit changes and close connection
        conn.commit()
        conn.close()
    # Driver Code
    app = tkinterApp()
    app.mainloop()