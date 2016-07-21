from openpyxl import Workbook
import sys

def output():

def readin(files):
    ret = {}
    chunk = {}
    for fname in files:
        with open(fname, 'r') as cfgfile:
            for line in cfgfile:
                line = line.strip()

                if not line or line.startswith('#') or line.startswith('define'): continue

                if line.startswith('}'):
                    ret[chunknum] = chunknumchunknum += 1
                    chunk = {}
                    continue

                bits = line.split(None, 1)

                if len(bits) > 1:
                    k, v = bits
                    chunk[k] = v

            return ret

worksheet.add_table('A1:F????', {'data': data,
                                ' columns':[{'header': 'Command Name'},
                                            {'header': 'Host Name'},
                                            {'header': 'Use'},
                                            {'header': 'Service Description'},
                                            {'header': 'Check Command'},
                                            {'header': 'File Name'},
                                            ]}) # How long does the table need to be? Line 27
