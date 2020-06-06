##################################################################################
# 
# MSExcelFindPassword 
# ---------------------------------------------------------------------------------
# Find Password of MS Excel File with custom wordlist "wordlist.lst" 
# ---------------------------------------------------------------------------------
# Yacine REZGUI <yacine.rezgui@gmail.com>
# Version 1.0
# ---------------------------------------------------------------------------------
# Prerequisites:
#  - win32com.client package need to install.
#  - Give proper path for excel file.
# 
# Usage :
# python MSExcelFindPassword.py <filename>
#							
###################################################################################

########## Builtin Package
import sys as sys
import os as os
import win32com.client as win32
from tqdm import tqdm

openedDoc = win32.gencache.EnsureDispatch('Excel.Application')

filename= sys.argv[1]
filepath = os.path.abspath(filename)

password_file = open ( 'wordlist.lst', 'r' )
passwords = password_file.readlines()
password_file.close()

passwords = [item.rstrip('\n') for item in passwords]

# Result store Path
results = open('results.txt', 'w')
print(filepath)
pwfind = 0


for password in tqdm(passwords):
    try:
        wb = openedDoc.Workbooks.Open(filepath, False, True, None, password)
        pwfind = 1
        results.write(password)
        results.close()

    except:
        continue

if pwfind == 1:
    results = open('results.txt', 'r')
    text = results.read()
    print(' ')
    print('Password Find : ')
    print('----------------------------')
    print(text)
    results.close()
else:
    print('Rien .....')
    
    
