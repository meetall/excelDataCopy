
import names
import helper
from openpyxl import load_workbook
import copy

import dataRead

dataDir = 'D:\\Chenglin_HR_Work\\01_Payroll\\01_Commission\\2018_Commission\\MNC'

def writeFile(monthIndex, name):
    #get month name from index
    month = helper.monthNames[monthIndex-1]
    print("Processing data for " + name+ " for "+month)
    templateFileName = helper.getMonthFolder(dataDir,month)+'\\2018commission_'+name+'.xlsx'
    new_file_name = helper.getMonthFolder(dataDir,month)+'\\2018commission_'+name+'_'+month+'.xlsx'
    helper.copy_rename(dataDir, dataDir, templateFileName, new_file_name)

    wb = load_workbook(filename = new_file_name)

    copyNsData(wb, monthIndex, name)
    fillNsNumOnSummarySheet(wb, monthIndex)

    copyBRData(wb, monthIndex, name)
    fillBROnSummarySheet(wb, monthIndex)
    
    wb.save(new_file_name)
    
def copyNsData(wb, monthIndex, name):
    ns_sheet = wb.create_sheet('NS')
    ns_data = dataRead.readNSDataForEmployee(monthIndex, name)
    for row in range(0,len(ns_data)):
        for col in range(0,len(ns_data[row])):
            ns_sheet.cell(row=row+1,column=col+1).value = ns_data[row][col].value

def fillNsNumOnSummarySheet(wb, monthIndex):
    ns_sheet = wb['NS']
    if ns_sheet.max_row == 1:
        return

    for m in range(1,monthIndex+1):
        sum=0
        for row in range(2, ns_sheet.max_row+1):
            sum = sum + ns_sheet.cell(row=row, column=27+m).value
        fillSummary(wb, sum, 12, m+1)        

def copyBRData(wb, monthIndex, name):
    br_sheet = wb.create_sheet('BR')
    br_data = dataRead.readBRDataForEmployee(monthIndex, name)
    for row in range(0,len(br_data)):
        for col in range(0,len(br_data[row])):
            br_sheet.cell(row=row+1,column=col+1).value = br_data[row][col].value

def fillBROnSummarySheet(wb, monthIndex):
    br_sheet = wb['BR']
    if br_sheet.max_row == 1:
        return

    for m in range(1,monthIndex+1):
        fillSummary(wb, 
            br_sheet.cell(row=br_sheet.max_row,column=10+m).value,17,1+m)

def fillSummary(wb, value, row, col):
    summary = wb['Sales Achievements Details']
    summary.cell(row=row, column=col).value=value
