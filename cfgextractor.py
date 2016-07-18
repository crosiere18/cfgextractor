#!/usr/bin/python3

from openpyxl import Workbook
import sys

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

def readin(files):
    ret = {}
    chunk = {}
    chunknum = 0 
    for fname in files:
        cfgfile = open(fname, 'r')
        
        for line in cfgfile:
            line = line.strip()
            
            if not line or line.startswith('#') or line.startswith('define'): continue
            
            if line.startswith('}'):
                ret[chunknum] = chunk
                chunknum += 1
                chunk = {}
                continue 

            bits = line.split(None, 1)

            if len(bits) > 1:
                k, v = bits
                chunk[k] = v

        cfgfile.close()
    return ret

def dechunkify(delim, axis, chunkmap, field):
    for chunk in chunkmap:
        if field in chunkmap[chunk]:
            values = chunkmap[chunk][field].split(delim)
            for item in values:
                if item and item not in axis:
                    axis.append(item)

def start(filescfg,xfield,yfield,pivot,data):
    xaxis = []
    yaxis = []

    delim = ','
    
    infiles = open(filescfg, 'r')
    xfiles = infiles.readline().strip().split(delim)
    yfiles = infiles.readline().strip().split(delim)

    xchunkmap = readin(xfiles)
    ychunkmap = readin(yfiles) 

    dechunkify(delim, xaxis, xchunkmap, xfield)
    dechunkify(delim, yaxis, ychunkmap, yfield)

    print(xaxis)
    print(yaxis)

    return "blah"

if __name__ == "__main__":
    failstring = ("requires 5 arguments: xaxis field, list of files containing xaxis field,"
                "yaxis field, list of files containing yaxis field, pivot field, and desir"
                "ed data fields.")

    if len(sys.argv) != 6:
        print(sys.argv[0] + failstring)
    else:
        start(sys.argv[1],sys.argv[2],sys.argv[3],sys.argv[4],sys.argv[5])

output(['a','b','c'],['1','2','3'],[['a1','a2','a3'],['b1','b2','b3'],['c1','c2','c3']])    
