import openpyxl
import csv
import datetime
import os
from tkinter import *
from tkinter import ttk

wb = openpyxl.Workbook()
sheet = wb.active

class Short_Long:

    def __init__(self, master):
        #Styling UI
        master.title('Short & Long Description Cheat Sheet')
        # master.resizable(False, False)
        master.configure(background='#DAF7A6')

        self.style = ttk.Style()
        self.style.configure('TFrame', background='#DAF7A6')
        self.style.configure('TLabel', background='#DAF7A6')
        self.style.configure('Header.TLabel', foreground='white', background='#90c74c', font=('Aerial', 18,'bold'))

        #Frame Header Setup
        self.frame_header = ttk.Frame(master)
        self.frame_header.pack()
        # self.frame_header.columnconfigure(0, weight=1)
        
        self.logo = PhotoImage(file = 'Masonite PP Banner.png')
        ttk.Label(self.frame_header, image = self.logo).grid(row=0, column=0, columnspan=4)
        ttk.Label(self.frame_header, text = 'Short & Long Description', style='Header.TLabel').grid(row=0, column=0, rowspan=2, columnspan=4)
        ttk.Label(self.frame_header, text = 'Long', foreground='#90c74c').grid(row=2, column=1, sticky='e')
        
        self.longer = BooleanVar()
        self.entry_long = ttk.Checkbutton(self.frame_header, variable=self.longer, onvalue=True, offvalue=False)
        self.entry_long.grid(row=2, column=2, sticky='w')

        #Frame Content Setup
        self.frame_content = ttk.Frame(master)
        self.frame_content.pack(expand=True)
        
        #Labels for Questions
        ttk.Label(self.frame_content, text = 'Backend Question Name:').grid(row=0, column=0, pady=3, sticky='e')
        ttk.Label(self.frame_content, text = 'Backend Group Question Name:').grid(row=1, column=0, pady=3, sticky='e')
        ttk.Label(self.frame_content, text = 'Order of Description:').grid(row=2, column=0, pady=3, sticky='e')
        ttk.Label(self.frame_content, text = 'Display Format:').grid(row=3, column=0, pady=3, sticky='e')
        ttk.Label(self.frame_content, text = 'Ignore Blank Answer:').grid(row=4, column=0, pady=3, sticky='e')
        ttk.Label(self.frame_content, text = 'Backend Answer to Hide:').grid(row=5, column=0, pady=3, sticky='e')
        ttk.Label(self.frame_content, text = 'Type in a Prefix:').grid(row=6, column=0, pady=3, sticky='e')
        ttk.Label(self.frame_content, text = 'Type in a Suffix:').grid(row=7, column=0, pady=3, sticky='e')
        ttk.Label(self.frame_content, text = 'Type text of an answer to change:').grid(row=8, column=0, pady=3, sticky='e')
        ttk.Label(self.frame_content, text = 'Type Backend answer that was changed:').grid(row=9, column=0, pady=3, sticky='e')
        ttk.Label(self.frame_content, text = 'Fraction or Decimal:').grid(row=10, column=0, pady=3, sticky='e')
        
        #Entry Values for Questions
        self.DFormat = StringVar()
        self.IgnoreBA = BooleanVar()
        self.Fraction = StringVar()
        self.row = 0
        self.order=StringVar()
    
        self.entry_question = ttk.Entry(self.frame_content, width = 24)
        self.entry_group_question = ttk.Entry(self.frame_content, width = 24)
        self.entry_order = ttk.Entry(self.frame_content, width = 24, textvariable=self.order)
        self.entry_display = ttk.Combobox(self.frame_content, textvariable=self.DFormat)
        self.entry_display.config(values=('Answer Only', 'Name Only', 'Name and Answer'))
        self.entry_ignore = ttk.Checkbutton(self.frame_content)
        self.entry_ignore.config(variable=self.IgnoreBA, onvalue=True, offvalue=False)
        self.entry_hide = ttk.Entry(self.frame_content, width = 24)
        self.entry_prefix = ttk.Entry(self.frame_content, width = 24)
        self.entry_suffix = ttk.Entry(self.frame_content, width = 24)
        self.entry_change_answer = ttk.Entry(self.frame_content, width = 24)
        self.entry_change_backend = ttk.Entry(self.frame_content, width = 24)
        self.entry_decimal = ttk.Combobox(self.frame_content, textvariable=self.Fraction)
        self.entry_decimal.config(values=('Decimal', 'Fraction'))

        self.entry_question.grid(row=0, column=1, pady=5)
        self.entry_group_question.grid(row=1, column=1, pady=5)
        self.entry_order.grid(row=2, column=1, pady=5)
        self.entry_display.grid(row=3, column=1, pady=5)
        self.entry_ignore.grid(row=4, column=1, sticky='w', pady=5)
        self.entry_hide.grid(row=5, column=1, pady=5)
        self.entry_prefix.grid(row=6, column=1, pady=5)
        self.entry_suffix.grid(row=7, column=1, pady=5)
        self.entry_change_answer.grid(row=8, column=1, pady=5)
        self.entry_change_backend.grid(row=9, column=1, pady=5)
        self.entry_decimal.grid(row=10, column=1, pady=5)

        self.list_view = ttk.Treeview(self.frame_content)
        self.list_view.grid(row=0, column=2, rowspan=6, columnspan=2, padx=10, pady=5)

        self.adding = ttk.Button(self.frame_content, text = 'Add', command=self.add)
        self.adding.grid(row=12, column=1, pady=3)
        ttk.Button(self.frame_content, text = 'Spit Out', command=self.SpitOut).grid(row=12, column=2, pady=3)
        self.adding.state(['disabled'])

        def buttonFunction(*args):
            if self.entry_question.get() and self.entry_group_question.get() and self.entry_order.get():
                self.adding.state(['!disabled'])
        
        self.order.trace('w', buttonFunction)

    #Clears the entry fields and 
    def add(self):
        def DescriptionOrder(att2, att3, attValue1):
            DO = 'Default,'+ att2 + ',' + att3 + ',' + attValue1 
            return DO

        def DisplayFormat(att3, displayValue):
            DF = 'Default,DisplayFormat,' + att3 + ',' + displayValue
            return DF

        def IgnoringAnswer(att1, att3, boolValue):
            IA = att1 + ',IgnoreBlankAnswers,' + att3 + ',' + boolValue
            return IA

        def AnswerVisibility(att1, att3, att3b):
            AV = att1 + ',ItemVisible,' + att3 + ':' + att3b + ',False'
            return AV

        def Prefix(att1, att3, text):
            PRE = att1 + ',QuestionAnswerPrefix,' + att3 + ',' + text
            return PRE

        def Suffix(att1, att3, text):
            SUF = att1 + ',QuestionAnswerSuffix,' + att3 + ',' + text
            return SUF

        def ChangeDisplayName(att1, att3, att3b, text):
            CDN = att1 + ',QuestionAnswerName,' + att3 + ':' + att3b + ',' + text
            return CDN

        def DisplayFractionDecimal(att1, att3, boolValue):
            DFD = att1 + ',ReportAsDecimal,' + att3 + ',' + boolValue
            return DFD

        def NewEndLine(att1, att3, boolValue):
            NEL = att1 + ',EndsLine,' + att3 + ',' + boolValue
            return NEL

        A3 = self.entry_question.get()
        A1 = self.entry_group_question.get()
        AV1 = self.entry_order.get()
        Format = self.DFormat.get()
        Ignore = self.IgnoreBA.get()
        NewLine = self.longer.get()
        if Ignore:
            Ignore = 'True'
        else:
            Ignore = ''
        Visibility = self.entry_hide.get()
        PrefixText = self.entry_prefix.get()
        SuffixText = self.entry_suffix.get()
        ChangeName = self.entry_change_answer.get()
        if ChangeName:
            AnswerName = self.entry_change_backend.get()
        FractionDecimal = str(self.Fraction)
        if FractionDecimal == 'Decimal':
            FractionDecimal = 'True'
        elif FractionDecimal == 'Fraction':
            FractionDecimal = 'False'
        else:
            FractionDecimal = ''
        if NewLine:
            NewLine = 'True'
        else:
            NewLine = ''

        initRow = self.row+1

        if AV1:
            # if self.row == 0:
            #     sheet = wb.active
            self.row+=1
            DO1 = DescriptionOrder(A1, A3, AV1)
            text = DO1.split(',')
            for i, cell in enumerate(text):
                print(i)
                print(cell)
                print(self.row)
                sheet.cell(row=self.row, column=i+1).value = cell
                # sheet[col[i]+str(self.row)] = cell

        if Format:
            self.row+=1
            DF1 = DisplayFormat(A3, Format)
            text = DF1.split(',')
            for i, cell in enumerate(text):
                sheet.cell(row=self.row, column=i+1).value = cell

        if Ignore:
            self.row+=1
            IgnoreText = IgnoringAnswer(A1, A3, Ignore)
            text = IgnoreText.split(',')
            for i, cell in enumerate(text):
                sheet.cell(row=self.row, column=i+1).value = cell
        
        if Visibility:
            self.row+=1
            VisibleText = AnswerVisibility(A1, A3, Visibility)
            text = VisibleText.split(',')
            for i, cell in enumerate(text):
                sheet.cell(row=self.row, column=i+1).value = cell

        if PrefixText:
            self.row+=1
            PFT = Prefix(A1, A3, PrefixText)
            text = PFT.split(',')
            for i, cell in enumerate(text):
                sheet.cell(row=self.row, column=i+1).value = cell
        
        if SuffixText:
            self.row+=1
            SFT = Suffix(A1, A3, SuffixText)
            text = SFT.split(',')
            for i, cell in enumerate(text):
                sheet.cell(row=self.row, column=i+1).value = cell

        if ChangeName:
            self.row+=1
            CN = ChangeDisplayName(A1, A3, AnswerName, ChangeName)
            text = CN.split(',')
            for i, cell in enumerate(text):
                sheet.cell(row=self.row, column=i+1).value = cell
            
        if FractionDecimal:
            self.row+=1
            FD = DisplayFractionDecimal(A1, A3, FractionDecimal)
            text = FD.split(',')
            for i, cell in enumerate(text):
                sheet.cell(row=self.row, column=i+1).value = cell
        
        if NewLine:
            self.row+=1
            NL = NewEndLine(A1, A3, NewLine)
            text = NL.split(',')
            for i, cell in enumerate(text):
                sheet.cell(row=self.row, column=i+1).value = cell
        
        for x in range(initRow, self.row+1):
            print('***')
            Desc = str(AV1)
            while len(Desc) != 5:
                Desc = '0' + Desc
            sheet['E'+ str(x)] = Desc
        
        self.list_view.insert('', AV1, A3, text=AV1+' - '+A3)

        self.entry_question.delete(0, 'end')
        self.entry_group_question.delete(0, 'end')
        self.entry_order.delete(0, 'end')
        self.entry_display.delete(0, 'end')
        if Ignore == 'True':
            self.IgnoreBA.set(False)
        self.entry_hide.delete(0, 'end')
        self.entry_prefix.delete(0, 'end')
        self.entry_suffix.delete(0, 'end')
        self.entry_change_answer.delete(0, 'end')
        self.entry_change_backend.delete(0, 'end')
        self.entry_decimal.delete(0, 'end')
        self.adding.state(['disabled'])

    #Excel file is generated and clears the treeview pane
    def SpitOut(self):
        self.row = 0
        for x in self.list_view.get_children():
            self.list_view.delete(x)
        wb.save(os.getcwd() + '\\' + '{:%m-%d-%Y-%H_%M_%S} '.format(datetime.datetime.now()) + '_Short_Description.xlsx')
        for i in range(1,sheet.max_row+1):
            for j in range(1,sheet.max_column+1):
                sheet.cell(row=i, column=j).value = ''

def main():
    root = Tk()
    ShortAndLong = Short_Long(root)
    root.mainloop()
    
if __name__ == "__main__": main()



