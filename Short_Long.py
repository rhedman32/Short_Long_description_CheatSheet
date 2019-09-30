import openpyxl
import csv
import datetime
import os

wb = openpyxl.Workbook()
sheet = wb.active
row = 1
column = ['A', 'B', 'C', 'D', 'E']
control = True


def DescriptionOrder(att2, att3, attValue1):
    DO = 'Default,'+ att2 + ',' + att3 + ',' + attValue1 
    return DO

def DisplayFormat(att3, displayValue):
    DF = 'Default,DisplayFormat,' + att3 + ',' + displayValue
    return DF

def IgnoringAnswer(att1, att3, boolValue):
    IA = att1 + ',IgnoreBlankAnswer,' + att3 + ',' + boolValue
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

while control:
    A3 = input('Enter Question Backend Name:')
    A1 = input('Enter Question Group Backend Name:')
    A2 = A1
    AV1 = input('Enter description order:')
    Format = input('Answer Only/Name Only/Name and Answer')
    Ignore = input('Ignore Blank asnwer?T/F:')
    Visibility = input('Type Backend answer to Hide')
    PrefixText = input('Do you need something before your answer?')
    SuffixText = input('Do you need something after?')
    ChangeName = input('Do you want to change the name of the answer?')
    if ChangeName:
        AnswerName = input('Type backend answer')
    FractionDeciaml = input('Fraction or Decimal')
    initRow = row

    if AV1:
        DO1 = DescriptionOrder(A2, A3, AV1)
        text = DO1.split(',')
        for i, cell in enumerate(text):
            #print(cell)
            sheet[column[i]+str(row)] = cell

    if Format:
        row+=1
        DF1 = DisplayFormat(A3, Format)
        text = DF1.split(',')
        for i, cell in enumerate(text):
            sheet[column[i]+str(row)] = cell

    if Ignore:
        row+=1
        IgnoreText = IgnoringAnswer(A1, A3, Ignore)
        text = IgnoreText.split(',')
        for i, cell in enumerate(text):
            sheet[column[i]+str(row)] = cell
    
    if Visibility:
        row+=1
        VisibleText = AnswerVisibility(A1, A3, Visibility)
        text = VisibleText.split(',')
        for i, cell in enumerate(text):
            sheet[column[i]+str(row)] = cell

    if PrefixText:
        row+=1
        PFT = Prefix(A1, A3, PrefixText)
        text = PFT.split(',')
        for i, cell in enumerate(text):
            sheet[column[i]+str(row)] = cell
    
    if SuffixText:
        row+=1
        SFT = Suffix(A1, A3, SuffixText)
        text = PFT.split(',')
        for i, cell in enumerate(text):
            sheet[column[i]+str(row)] = cell

    if ChangeName:
        row+=1
        CN = ChangeDisplayName(A1, A3, AnswerName, ChangeName)
        text = CN.split(',')
        for i, cell in enumerate(text):
            sheet[column[i]+str(row)] = cell

    for x in range(initRow, row):
        print(row)
        sheet['E'+ str(x)] = '000'+ str(AV1)
    control = input('T/F')
    if control == 'F':
        break
    else:
        control = True      
    row += 1

wb.save(os.getcwd() + '\\' + '{:%m-%d-%Y-%H_%M_%S} '.format(datetime.datetime.now()) + '_Short_Description.xlsx')
print('File Spit Out Completed')






