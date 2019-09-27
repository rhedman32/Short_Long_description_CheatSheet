import openpyxl

wb = openpyxl.Workbook()
sheet = wb.active
row = '1'
column = 'A'



def DescriptionOrder(att2, att3, attValue):
    DO = 'Default,'+ att2 + ',' + att3 + ',' + attValue 
    return DO

def DisplayFormat(att3, displayValue):
    DF = 'Default,DisplayFormat,' + att3 + ',' + displayValue
    return DF

def IgnoringAnswer(att1, att3, boolValue):
    IA = att1 + ',IgnoreBlankAnswer,' + att3 + ',' + boolValue
    return IA

def AnswerVisibility(att1, att3, att3b, boolValue):
    AV = att1 + ',ItemVisible,' + att3 + ':' + att3b + ',' + boolValue
    return AV

def Prefix(att1, att3, text):
    PRE = att1 + ',QuestionAnswerPrefix,' + att3 + ',' + text
    return PRE

def Suffix(att1, att2, text):
    SUF = att1 + ',QuestionAnswerPrefix,' + att3 + ',' + text
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



