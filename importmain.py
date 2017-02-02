from modules.importWIP import importWIPdata

# importWIPdata takes information in the form
# importWIPdata(DATABASE_NAME, importdirectory)

# import Midlands data
print('Importing Midlands data')
importWIPdata(
    './midlands/midlandswipdata.sqlite',
    'H:/Current Month Wips/2. Midlands')

# import London data
print('Importing London data')
importWIPdata(
    './london/londonwipdata.sqlite',
    'H:/Current Month Wips/1. London & South East')
