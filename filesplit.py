import pandas as pd
import os
from openpyxl import load_workbook
from shutil import copyfile

file = input('File Path: ')
extension = os.path.splitext(file)[1]
filename = os.path.splitext(file)[0]
pth = os.path.dirname(file)
newfile = os.path.join(filename + '_2' + extension)
df = pd.read_excel(file)
colpick = input('Select Column: ')
cols = list(set(df[colpick].values))


def sendtofile(cols):
    for i in cols:
        df[df[colpick] == i].to_excel("{}/{}.xlsx".format(pth, i), sheet_name=i, index=False, encoding='utf-8')
    print('\nCompleted')
    print('Thanks for using this program.')
    return


def sendtosheet(cols):
    copyfile(file, newfile)
    for j in cols:
        writer = pd.ExcelWriter(newfile, engine='openpyxl')
        for myname in cols:
            mydf = df.loc[df[colpick] == myname]
            mydf.to_excel(writer, sheet_name=myname, index=False)
        writer.save()

    print('\nCompleted')
    print('Thanks for using this program.')
    return


print('Your data will split based on these values {} and create {} files or sheets based on next selection. If you are ready to proceed please type "Y" and hit enter. Hit "N" to exit.'.format(', '.join(cols), len(cols)))

while True:
    x = input('Ready to Proceed (Y/N): ').lower()
    if x == 'y':
        while True:
            s = input('Split into different Sheets or Files (S/F): ').lower()
            if s == 'f':
                sendtofile(cols)
                break
            elif s == 's':
                sendtosheet(cols)
                break
            else:
                continue
        break
    elif x == 'n':
        print('\nThanks for using this program.')
        break

    else:
        continue
