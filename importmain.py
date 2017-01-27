from importWIP import importWIPdata

# importWIPdata takes information in the form
# importWIPdata(DATABASE_NAME, importdirectory)

# import Midlands data
importWIPdata(r'.\midlands\midlandswipdata.sqlite', r'H:\Current Month Wips\2. Midlands')

# import London data
importWIPdata(r'.\london\londonwipdata.sqlite', r'H:\Current Month Wips\1. London & South East')
