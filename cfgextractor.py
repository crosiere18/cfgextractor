#!/usr/bin/python3

from openpyxl import Workbook

def output(xaxis, yaxis, data):
    wb = Workbook()

    ws = wb.active
    ws.title = 'Main'

    xlen = len(xaxis) + 2
    ylen = len(yaxis) + 2

    for col in range(2, xlen):
        ws.cell(row = 1, column = col).value = xaxis[col-2]
        
    for row in range(2, ylen):
        ws.cell(row = row, column = 1).value = yaxis[row-2]
    
    for cidx, col in enumerate(data):
        for ridx, cdata in enumerate(col):
            ws.cell(row=ridx+2,column=cidx+2).value = cdata


    wb.save('output.xlsx')

output(['a','b','c'],['1','2','3'],[['a1','a2','a3'],['b1','b2','b3'],['c1','c2','c3']])    
