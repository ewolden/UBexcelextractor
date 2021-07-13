#!/usr/bin/env python
# -*- coding: utf-8 -*-
#pip install xlrd
#pip install openpyxl

import sys, getopt
from datetime import date
import pandas as pd
import numpy as np
import string

def col2num(col):
    num = 0
    for c in col:
        if c in string.ascii_letters:
            num = num * 26 + (ord(c.upper()) - ord('A')) + 1
    return num - 1

def print_usage():
    print('\nBruk: excelextractor.py -i <inputfil> -s <søkefil> -o <outputfil> -sc <søkekollone> -ic <innsettingskollone>')
    print('-i --input       -  full path eller relativ path')
    print('-s --search      -  full path eller relativ path')
    print('-o --output      -  full path eller relativ path')
    print('-k --searchCol  -  kollonenummer/bokstaver i inputfilen som det skal søkes i')
    print('-l --insertCol  -  kollonenummer/bokstaver i søkefilen som skal settes inn i outpufilen')
    print('\n')

def main(argv):
    input_filename = ''
    search_filename = ''
    output_filename = date.today().strftime('%Y%m%d') + 'output.xlsx'
    search_col = -1
    insert_col = -1

    try:
        opts, args = getopt.getopt(argv,'hi:s:o:k:l:',['help', 'input=', 'search=', 'output=', 'searchCol=', 'insertCol='])
    except getopt.GetoptError:
        print_usage()
        sys.exit(2)
    for opt, arg in opts:
        if opt == '-h':
            print_usage()
            sys.exit()
        elif opt in ("-i", "--input"):
            input_filename = arg
        elif opt in ("-s", "--search"):
            search_filename = arg
        elif opt in ("-o", "--output"):
            output_filename = arg
        elif opt in ('-k', '--searchCol'):
            search_col = arg
        elif opt in ('-l', '--insertCol'):
            insert_col = arg

    if input_filename == '':
        print(' -  Inputfil ikke spesifisert, avslutter. Skriv "excelextractor.py -h" for hjelp')
        sys.exit()  
    if search_filename == '':
        print(' -  Søkefil ikke spesifisert, avslutter. Skriv "excelextractor.py -h" for hjelp')
        sys.exit()  
    if output_filename == date.today().strftime('%Y%m%d') + 'output.xlsx':
        print(' -  Outputfil ikke spesifisert, skriver til ', output_filename)
    if search_col == -1:
        print(' -  Ingen kollone valgt for søk, søker i alle kolonner i inputfilen')
    elif isinstance(search_col, str):
        search_col = col2num(search_col)
    if insert_col == -1:
        print(' -  Ingen kollone valgt for innsetting, setter inn alle kolonner fra søkefilen')
    elif isinstance(insert_col, str):
        insert_col = col2num(insert_col)
    
    input_dfs = None
    search_dfs = None
    try:
        input_dfs = pd.read_excel(input_filename)
        search_dfs = pd.read_excel(search_filename)
    except Exception as e:
        print(' -  Fant ikke inputfil og/eller søkefil')
        sys.exit()

    print("Prosseserer...")

    found_rows = pd.Series(dtype = 'str')

    if insert_col == -1:
        input_dfs = input_dfs.reindex(input_dfs.columns.tolist() + search_dfs.columns.tolist(), axis=1)
    else:
        input_dfs = input_dfs.reindex(input_dfs.columns.tolist() + [search_dfs.columns.tolist()[insert_col]], axis = 1)

    if search_col == -1:
        for index, keyword in enumerate(search_dfs.iloc[:,0]):
            current_found_rows = pd.Series(dtype = 'str')
            for column in input_dfs:
                if(input_dfs[column].dropna().empty):
                    continue
                current_found_rows = current_found_rows | input_dfs[column].str.contains(str(keyword), na=False, regex=False)
            if insert_col == -1:
                for cell_index, cell_value in enumerate(search_dfs.iloc[:index,]):
                    for true_index in input_dfs.index[current_found_rows]:
                        input_dfs.iloc[true_index,len(input_dfs.columns)- len(search_dfs.columns) + cell_index] = str(search_dfs.iloc[index, cell_index])
            else:
                for true_index in input_dfs.index[current_found_rows]:
                        input_dfs.iloc[true_index,len(input_dfs.columns)- 1] = str(search_dfs.iloc[index, insert_col])
            found_rows = found_rows | current_found_rows
        
    else:
        for index, keyword in enumerate(search_dfs.iloc[:,0]):
            current_found_rows = pd.Series(dtype = 'str')
            current_found_rows = input_dfs.iloc[:,search_col].str.contains(str(keyword), na=False, regex=False)
            if insert_col == -1:
                for cell_index, cell_value in enumerate(search_dfs.iloc[:index,]):
                    for true_index in input_dfs.index[current_found_rows]:
                        input_dfs.iloc[true_index,len(input_dfs.columns)- len(search_dfs.columns) + cell_index] = str(search_dfs.iloc[index, cell_index])
            else:
                for true_index in input_dfs.index[current_found_rows]:
                        input_dfs.iloc[true_index,len(input_dfs.columns)- 1] = str(search_dfs.iloc[index, insert_col])
            found_rows = found_rows | current_found_rows

    
    output_dfs = input_dfs[found_rows]
    #print(output_dfs)
    output_dfs.to_excel(output_filename)
    print(" = Ferdig prossesert, output skrevet til: ", output_filename)
if __name__ == "__main__":
    main(sys.argv[1:])
