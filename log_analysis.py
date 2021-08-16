import os
import sys
import argparse
import re
import openpyxl




if __name__ == "__main__":
 #   
  #  parser = argparse.ArgumentParser(description = "Log Format to CSV")
   # parser.add_argument('-b', '--board', help ="board type")
    #parser.add_argument('-l', '--log', help = "log file")
    #args = parser.parse_args()
    #print(args.log)


    wb = openpyxl.Workbook()
    file_name = 'hermal_test.xlsx'
    ws1=wb.active
    ws1.title = "test1"
    
    ws2 = wb.create_sheet(title = "test2")
    ws2['A1'] = 3

    wb.save(file_name)
