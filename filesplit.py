import pandas as pd
import os
from openpyxl import load_workbook
from shutil import copyfile

myfile = input('Please type The Excel file path with its extension [.xlsx]:   ')

if os.path.isfile(myfile):
    myfile_root, myfile_extension = os.path.splitext(myfile)

    if myfile_extension == '.xlsx':
        file = myfile
        filename, extension = os.path.splitext(file)
        pth = os.path.dirname(file)
        newfile = os.path.join(filename + '_2' + extension)
        df = pd.read_excel(file)
        column_name = input('Type the name of Column to split by: ')
        cols_name = list(set(df))

        if column_name in cols_name:
            cols = list(set(df[column_name].values))
            print(
                'Your data will split based on these values {} and create {} files or sheets based on next selection.'
                ' If you are ready to proceed please type "Y" and hit enter. Hit "N" to exit.'\
                    .format(', '.join(cols), len(cols)))

        else:
            print(f"The Column name was not correct. Please try again: {column_name}")
            exit()


    else:
        print(f"The file you chose was not a [.xlsx] Excel file. Please try again: {myfile}")
        exit()

else:
    print(f"The file does't exists. Please try again with a correct file path: {myfile}")
    exit()


def sendtofile(cols):
    for i in cols:
        df[df[column_name] == i].to_excel("{}/{}.xlsx".format(pth, i), sheet_name=i, index=False, encoding='utf-8')
    print('\nCompleted')
    print('Thanks for using this program.')
    return


def sendtosheet(cols):
    copyfile(file, newfile)
    for j in cols:
        writer = pd.ExcelWriter(newfile, engine='openpyxl')
        for myname in cols:
            mydf = df.loc[df[column_name] == myname]
            mydf.to_excel(writer, sheet_name=myname, index=False)
        writer.save()

    print('\nCompleted')
    print('Thanks for using this program.')
    return


while True:
    x = input('Ready to Proceed? (Y/N): ').lower()
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
