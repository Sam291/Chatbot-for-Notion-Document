import pandas as pd
import re
from openpyxl import load_workbook

def isStringAvailable(ts, findString):
    ts = [str(i) for i in list(ts)]
    find = re.findall(findString, "|".join(ts).lower())
    if find:
        return True
    return False

def findData(findString):
    idx = df.apply(lambda ts: isStringAvailable(ts, findString), axis=1)
    return idx

def create_search_string(inputString):
    words = re.findall('[a-z]+', inputString, re.IGNORECASE)
    return '(?=.*' + ')(?=.*'.join(words).lower() + ').*'


inputString = input("Enter string/phrase to search (I am case agnostic)\n")
exactMatch = input('Exact match? (y)\n')
if inputString:
    inputFile = 'MASTER_BRD_new.xlsx'
    outputFile = 'output.xlsx'

    if exactMatch:
        print('Exact match')
        findString = inputString.lower()
    else:
        print('Random match')
        findString = create_search_string(inputString)
        
    writer = pd.ExcelWriter(outputFile, engine='openpyxl')
    found = 0
    workbook = load_workbook(inputFile)
    for worksheet in workbook:
        group = worksheet.title
        df = pd.read_excel(inputFile, group, engine='openpyxl')
        filtered_df = df[findData(findString)]
        if not filtered_df.empty:
            print(filtered_df)
            filtered_df.to_excel(writer, sheet_name=group, index=False, header=df.columns)
            found = 1
    if found:
        writer.close()
    else:
        print('Not found!')
