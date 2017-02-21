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

# import North East data
print('Importing North East data')
importWIPdata(
    './northeast/northeastwipdata.sqlite',
    'H:/Current Month Wips/3. North East'
)

# import North West data
print('Importing North West data')
importWIPdata(
    './northwest/northwestwipdata.sqlite',
    'H:/Current Month Wips/3b. North West'
)

# import South West data
print('Importing South West data')
importWIPdata(
    './southwest/southwestwipdata.sqlite',
    'H:/Current Month Wips/4. South West & Wales'
)
