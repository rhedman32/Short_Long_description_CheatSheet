import openpyxl
import csv
import datetime
import os
from tkinter import *
from tkinter import ttk

wb = openpyxl.Workbook()
sheet = wb.active
row = 1
column = ['A', 'B', 'C', 'D', 'E']
control = True

class Short_Long:

    def __init__(self, master):
        
        self.frame_header = ttk.Frame(master)
        self.frame_header.pack()
        self.frame_header.columnconfigure(0, weigh=1)
        
        self.logo = PhotoImage(file = 'Masonite PP Banner for Aaron.png')
        ttk.Label(self.frame_header, image = self.logo).grid(row=0, column=0, columnspan=4)
        ttk.Label(self.frame_header, text = 'Short & Long Description').grid(row=0, column=0, rowspan=2, columnspan=4)

        self.frame_content = ttk.Frame(master)
        self.frame_content.pack(expand=True)
        
        ttk.Label(self.frame_content, text = 'Backend Question Name:').grid(row=0, column=0, pady=3, sticky='e')
        ttk.Label(self.frame_content, text = 'Backend Group Question Name:').grid(row=1, column=0, pady=3, sticky='e')
        ttk.Label(self.frame_content, text = 'Order of Description:').grid(row=2, column=0, pady=3, sticky='e')

        self.entry_question = ttk.Entry(self.frame_content, width = 24)
        self.entry_group_question = ttk.Entry(self.frame_content, width = 24)
        self.entry_order = ttk.Entry(self.frame_content, width = 24)

        self.entry_question.grid(row=0, column=1)
        self.entry_group_question.grid(row=1, column=1)
        self.entry_order.grid(row=2, column=1)

        self.list_view = ttk.Treeview(self.frame_content)
        self.list_view.grid(row=0, column=2, rowspan=6, columnspan=2, padx=3, pady=3)

        ttk.Button(self.frame_content, text = 'Add', command=self.add).grid(row=6, column=1, pady=3)
        ttk.Button(self.frame_content, text = 'Spit Out').grid(row=6, column=2, pady=3)

    def add(self):
        self.entry_question.delete(0, 'end')
        self.entry_group_question.delete(0, 'end')
        self.entry_order.delete(0, 'end')

    # def SpitOut(self):

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

    # while control:
    #     A3 = input('Enter Question Backend Name:')
    #     A1 = input('Enter Question Group Backend Name:')
    #     AV1 = input('Enter description order:')
    #     Format = int(input('Answer Only:1/Name Only:2/Name and Answer:3 '))
    #     if Format == 1:
    #         Format = 'Answer Only'
    #     elif Format == 2:
    #         Format = 'Name Only'
    #     elif Format == 3:
    #         Format = 'Name and Answer'
    #     else:
    #         Format = ''
    #     Ignore = input('Ignore Blank asnwer? t/f: ')
    #     if Ignore == 't':
    #         Ignore = 'True'
    #     elif Ignore == 'f':
    #         Ignore = 'False'
    #     Visibility = input('Type Backend answer to Hide: ')
    #     PrefixText = input('Do you need something before your answer? ')
    #     SuffixText = input('Do you need something after? ')
    #     ChangeName = input('Do you want to change the name of the answer? ')
    #     if ChangeName:
    #         AnswerName = input('Type backend answer: ')
    #     FractionDecimal = input('Decimal:1 or Fraction:2 ')
    #     if FractionDecimal == 1:
    #         FractionDecimal = 'True'
    #     elif FractionDecimal == 2:
    #         FractionDecimal = 'False'
        
    #     initRow = row

    #     if AV1:
    #         DO1 = DescriptionOrder(A1, A3, AV1)
    #         text = DO1.split(',')
    #         for i, cell in enumerate(text):
    #             #print(cell)
    #             sheet[column[i]+str(row)] = cell

    #     if Format:
    #         row+=1
    #         DF1 = DisplayFormat(A3, Format)
    #         text = DF1.split(',')
    #         for i, cell in enumerate(text):
    #             sheet[column[i]+str(row)] = cell

    #     if Ignore:
    #         row+=1
    #         IgnoreText = IgnoringAnswer(A1, A3, Ignore)
    #         text = IgnoreText.split(',')
    #         for i, cell in enumerate(text):
    #             sheet[column[i]+str(row)] = cell
        
    #     if Visibility:
    #         row+=1
    #         VisibleText = AnswerVisibility(A1, A3, Visibility)
    #         text = VisibleText.split(',')
    #         for i, cell in enumerate(text):
    #             sheet[column[i]+str(row)] = cell

    #     if PrefixText:
    #         row+=1
    #         PFT = Prefix(A1, A3, PrefixText)
    #         text = PFT.split(',')
    #         for i, cell in enumerate(text):
    #             sheet[column[i]+str(row)] = cell
        
    #     if SuffixText:
    #         row+=1
    #         SFT = Suffix(A1, A3, SuffixText)
    #         text = SFT.split(',')
    #         for i, cell in enumerate(text):
    #             sheet[column[i]+str(row)] = cell

    #     if ChangeName:
    #         row+=1
    #         CN = ChangeDisplayName(A1, A3, AnswerName, ChangeName)
    #         text = CN.split(',')
    #         for i, cell in enumerate(text):
    #             sheet[column[i]+str(row)] = cell
            
    #     if FractionDecimal:
    #         row+=1
    #         FD = DisplayFractionDecimal(A1, A3, FractionDecimal)
    #         text = FD.split(',')
    #         for i, cell in enumerate(text):
    #             sheet[column[i]+str(row)] = cell

    #     for x in range(initRow, row+1):
    #         print('***')
    #         Desc = str(AV1)
    #         while len(Desc) != 5:
    #             Desc = '0' + Desc
    #         sheet['E'+ str(x)] = Desc 
    #     control = input('Continue? y/n: ')
    #     if control == 'n':
    #         break
    #     else:
    #         control = True
    #     row += 1

    # wb.save(os.getcwd() + '\\' + '{:%m-%d-%Y-%H_%M_%S} '.format(datetime.datetime.now()) + '_Short_Description.xlsx')
    # print('File Spit Out Completed')


def main():
    root = Tk()
    ShortAndLong = Short_Long(root)
    root.mainloop()
    
if __name__ == "__main__": main()



