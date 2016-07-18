#!/usr/bin/python3

from openpyxl import Workbook
from string import ascii_uppercase
import math

def ArgumentTooLarge(Exception):
    def __init__(self, value):
        self.value = value
    def __str__(self):
        return repr(self.value)

def xAxisLetter(rowlen):
    rowletter = ''
    while rowlen > 26:
        if rowlen >= 17576:
            raise ArgumentTooLarge(rowlen)
        elif rowlen > 676:
            val = math.floor(rowlen/676 - 1)
            rowletter += ascii_uppercase[val]
            rowlen -= (val + 1) * 676
        else:
            val = math.floor(rowlen/26 - 1)
            rowletter += ascii_uppercase[val]
            rowlen -= (val + 1) * 26

    rowletter += ascii_uppercase[rowlen % 26]
    return rowletter

def output(filename, sheetname, xaxis, yaxis, data):
    wb = Workbook()

    ws = wb.active

    xlen = len(xaxis) + 1
    ylen = len(yaxis) + 1

print(xAxisLetter(25))
print(xAxisLetter(26))
print(xAxisLetter(27))
print(xAxisLetter(52))
print(xAxisLetter(53))
