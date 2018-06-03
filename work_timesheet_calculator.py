#pylint: skip-file


def nothing(signum, frame):
    return


import signal
signal.signal(signal.SIGINT, nothing)

import win32com.client
from os import listdir
from os.path import isfile, join

xlApp = win32com.client.Dispatch("Excel.Application")
path_to_speadsheets = 'C:\\WorkTimesheets\\'
onlyfiles = []
try:
    onlyfiles = [f for f in listdir(path_to_speadsheets) if isfile(
        join(path_to_speadsheets, f)) and 'xls' in f.lower()]
except:
    path_to_speadsheets = input('\n------------------------------------------------------------------------------------------------------------------------\n' +
                                '\nThe files are not in the default path (C:\\WorkTimesheets\\). Type in the path to the timesheets.\n' +
                                '\nFormat: C:\\\\path\\\\to\\\\folder\\\\\n' +
                                '\n------------------------------------------------------------------------------------------------------------------------\n\n' +
                                '> ')
    try:
        onlyfiles = [f for f in listdir(path_to_speadsheets) if isfile(
            join(path_to_speadsheets, f)) and 'xls' in f.lower()]
        print('\nValid path.')
    except:
        print('\nInvalid path. Exiting...\n')
        input('Press any key to exit.\n\n')
        exit(1)
balance_minutes = 0

print('\n------------------------------------------------------------------------------------------------------------------------')
print('\nOpening timesheets...')
print('\n------------------------------------------------------------------------------------------------------------------------\n')
if len(onlyfiles) == 0:
    print('No excel file was found.')
for xlsfile in onlyfiles:
    workBook = xlApp.Workbooks.Open(path_to_speadsheets + xlsfile)
    for j in range(20):
        for i in range(60):
            try:
                cell = str(workBook.ActiveSheet.Cells(i+1, j+1))
                if cell == 'None':
                    continue
                if 'Débito' not in cell and 'Crédito' not in cell:
                    continue
                delta = cell.split(' - ')[0]
                hours = delta.split(':')[0]
                minutes = delta.split(':')[1]
                delta = 60 * int(hours) + int(minutes)
                text = cell.split(' - ')[1]
                if 'Débito' in text:
                    delta = -delta
                elif 'Crédito' not in text:
                    delta = 0
                print('Current Balance: ' + str(balance_minutes).ljust(4) + '   \t|\tDelta: ' +
                      str(delta) + '   \t|\tCell Info: ' + cell)
                balance_minutes += delta
            except:
                continue
    workBook.Close(SaveChanges=0)
xlApp.Quit()
print('\n------------------------------------------------------------------------------------------------------------------------\n')
print('Final Balance: ' + str(int(balance_minutes / 60)).zfill(2) +
      ':' + str(int(abs(balance_minutes) % 60)).zfill(2))
print('\n------------------------------------------------------------------------------------------------------------------------')
print('\nCreator: gabriel.dias.rezende.martins@gmail.com')
print('\n------------------------------------------------------------------------------------------------------------------------\n')
input('Press any key to exit.\n\n')
