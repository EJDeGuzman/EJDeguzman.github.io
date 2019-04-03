# -*- coding: utf-8 -*-
"""
Created on Wed Mar 20 13:37:21 2019
@author: Edgar DeGuzman, Chris Sinclair

Instructons: 
    1)Open applicaiton 
    2)Applicaton should be used for a weekly basis
    3)Rename file (Ex: MonthYear_BookKeeping.xlsx) on line 17
    4)Applicaiton will take one instance of it being executed
    5)After that, any code that is executed will overwrite the previous data already present
"""
import xlsxwriter 
from pathlib import Path 
from xlrd import open_workbook
#from pandas import Series, Dataframe as df
import pandas as pd
budget = 1200

workbook = xlsxwriter.Workbook("Book Keeping.xlsx")                 #creates workbook named Book Keeping.xlsx
bk_worksheet = workbook.add_worksheet('Book Keeping Spreadheet')       #creates worksheet   

#path to the Book Keeping.xlsx file 
book_keeping_file = Path('C:\Users\ebdegu01\Documents\Programming Notes\Python Scripts\Book Keeping.xlsx')

bold = workbook.add_format({'bold': True})                      #bold format for text 
bold.set_font_size(14)                                          #sets the font size to 15   
money = workbook.add_format({'num_format': '$#,##0.00'})        #money format for cells

bk_worksheet.write('A1', 'Book Keeping Spreadsheet', bold)      #writes the heading for Book Keeping Spreadsheet
bk_worksheet.set_column(0,0,30)                                 #sets the column width
bk_worksheet.write('B1', 'Date: 03/20/2019', bold)              
bk_worksheet.set_column(1,1,23)
bk_worksheet.write('C1', 'Edgar DeGuzman', bold)
bk_worksheet.set_column(2,4,23)
bk_worksheet.write('$G$1', 'Budget', bold)
bk_worksheet.write('$H$1', budget, money)
bk_worksheet.write('$G$3','Net', bold)
bk_worksheet.write('$H$3', '=$H$1-$H$2', money)

#array for the headers in the excel file
book_array_content = ['Amount Paid', 'Transaction Purpose', 'Transaction Date', 'Tranaction Location', 'Payee']
row = 2                                                         #sets the row on the Excel sheet for row 3
column = 0                                                      #sets the column on the Excel sheet for column 1
for item in book_array_content:
    bk_worksheet.write(row, column, item, bold)
    column += 1
if book_keeping_file.is_file():
    row = 3                                                     #sets the row on Excel shet for row 4
    
    #function that will take the input information and put into an array
    def transaction_function():
        print('Enter the following information.')
        
        #variable to hold amount input that the user will enter
        amount_input = input('Amount paid for Transaction')
        if amount_input == "":
            print('No amount was entered!')
            #transaction_purpose_array = ['Rent', 'Utilities', 'Groceries', 'Restaurant', 'gas', 'Recreational']
        
        #User will enter the purpose of the transaction 
        transaction_purpose = raw_input('What was the transaction for? Rent,Utilities,Groceries,Restaurant,Gas,Social?')
        if transaction_purpose == "":
            print('The purpose of the payment was not specified!')
        #else:
            #print(transaction_purpose)
            
        #user will enter the date of the transaction
        transaction_date = raw_input('When did the transaction occur?')
        if transaction_date == "":
            print('The date was not entered!')
        
        #user will enter the location of the transaction
        transaction_location = raw_input('Where was the transaction?')
        if transaction_location == "":
            print('The location was not specified!')
        
        #user will enter who the transaction was paid to
        transaction_payee = raw_input('Who was the tranaction paid to?')
        if transaction_payee == "":
            print('The payee was not specified')
        
        #array for the input variables 
        book_array = [amount_input, transaction_purpose, transaction_date, transaction_location, transaction_payee]
        print(book_array)
        column = 0
        for x in book_array:
            bk_worksheet.write(row, column, x, money)
            column += 1
    transaction_function()                          #calls the transaction funtion     
    
                                     
    #User enters yes or no for transaction counter
    transaction_counter = raw_input('Do you need to make another transaction?')
    if transaction_counter == 'yes':
        while transaction_counter == 'yes':
            row += 1
            transaction_function()
            transaction_counter = raw_input('Do you need to make another transaction?') 
    elif transaction_counter == 'no':
        print('Thank you for entering your transaction information!')
    else:
        print('Error!')
    
    bk_worksheet.write('$G$2','Total',bold)
    bk_worksheet.write('$H$2','=SUM(A4:A20)',money)   
    transaction_purpose_col = 1
    transaction_purpose_array = ['Rent', 'Utilities', 'Groceries', 'Restaurant', 'Gas', 'Social']
    
    bk_worksheet.write('G6','Rent',bold)
    bk_worksheet.write('G7','Utilities',bold)
    bk_worksheet.write('G8','Groceries',bold)
    bk_worksheet.write('G9','Restaurant',bold)
    bk_worksheet.write('G10','Gas',bold)
    bk_worksheet.write('G11','Social',bold)
    bk_worksheet.write('G12','Total',bold)
    
    bk_worksheet.write('$H$6','=SUMIF(B4:B9,"=Rent",A4:A9)',money)
    bk_worksheet.write_formula('$H$7','=SUMIF(B4:B9,"=Utilities",A4:A9)',money)
    bk_worksheet.write_formula('$H$8','=SUMIF(B4:B9,"=Groceries",A4:A9)',money)
    bk_worksheet.write_formula('$H$9','=SUMIF(B4:B9,"=Restaurant",A4:A9)',money)
    bk_worksheet.write_formula('$H$10','=SUMIF(B4:B9,"=Gas",A4:A9)',money)
    bk_worksheet.write_formula('$H$11','=SUMIF(B4:B9,"=Social",A4:A9)',money)
    bk_worksheet.write_formula('$H$12','=SUM(H6:H11)',money)
    
    bkchart_worksheet = workbook.add_worksheet('Data Visuals')
    bkchart_worksheet.write('A1','Data Visualiztions',bold)
    bkchart_worksheet.set_column(0,0,22)
    bkchart_worksheet.write('A3','Rent',bold)
    bkchart_worksheet.write('A4','Utilities',bold)
    bkchart_worksheet.write('A5','Groceries',bold)
    bkchart_worksheet.write('A6','Restaurant',bold)
    bkchart_worksheet.write('A7','Gas',bold)
    bkchart_worksheet.write('A8','Social',bold)
    bkchart_worksheet.set_column(0,2,15)
    bkchart_worksheet.write('A9','Total',bold)
    
    bkchart_worksheet.write_formula('B3',"='Book Keeping Spreadheet'!H6",money)
    bkchart_worksheet.write_formula('B4',"='Book Keeping Spreadheet'!H7",money)
    bkchart_worksheet.write_formula('B5',"='Book Keeping Spreadheet'!H8",money)
    bkchart_worksheet.write_formula('B6',"='Book Keeping Spreadheet'!H9",money)
    bkchart_worksheet.write_formula('B7',"='Book Keeping Spreadheet'!H10",money)
    bkchart_worksheet.write_formula('B8',"='Book Keeping Spreadheet'!H11",money)
    bkchart_worksheet.write_formula('B9',"=SUM(B3:B8)",money)
    
    bkpiechart = workbook.add_chart({'type': 'pie'})
    
    bkpiechart.add_series({
        'categories' : "='Data Visuals'!A3:A8",
        'values' : "='Data Visuals'!B3:B8"
        })
    bkpiechart.set_title({'name': 'Budget Pie Chart'})
    
    bkpiechart.set_rotation(90)
    bkchart_worksheet.insert_chart('D3', bkpiechart)
    
    bklinechart = workbook.add_chart({'type': 'line'})
    bklinechart.add_series({
            'categories' : "='Book Keeping Spreadsheet'!C4:C20",
            'values' : "='Book Keeping Spreadsheet'!A4:A20"
            })
    bklinechart.set_title({'name': 'Budget Line Chart'})
    bkchart_worksheet.insert_chart('D20', bklinechart)
    
else: 
    print('The file does not exist!')
workbook.close()

    