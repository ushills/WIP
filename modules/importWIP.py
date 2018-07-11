# import WIP data from WIP file

# import libraries
from openpyxl import load_workbook
import time
import datetime
import sqlite3
import os
import re

# supress warnings
import warnings
warnings.filterwarnings("ignore")


# importWIPdata takes information in the form
# importWIPdata(DATABASE_NAME, importdirectory)
def importWIPdata(DBNAME, directoryname):
    global DATABASE_NAME
    global directory
    DATABASE_NAME = DBNAME
    directory = directoryname
    print('importing from', directoryname, 'to', DBNAME)

    # first extract the files from the directory
    filelist = listFiles(os.path.normpath(directory))

    # check the files are wip files
    wipfiles = checkWipfile(filelist)

    # extract the data from the excel worksheet cells
    # and export it to the sqldatabase
    for wipfilename in wipfiles:
        wipData = importData(wipfilename)
        exportDataSql(wipData)


# search through directory to only return excel files
def listFiles(directory):
    print('searching for files.....')
    filelist = []
    for root, directories, filenames in os.walk(os.path.normpath(directory)):
        for filename in filenames:
            rawfilename = str(os.path.join(root, filename))
            # skip temp files beginning with ~
            if re.search(r"(~.*).*", rawfilename):
                continue
            # print rawfilename
            # check if the file is excel, i.e. ends .xlsx
            if re.search(r"(.*).xls?", rawfilename):
                # print 'Appending', rawfilename
                filelist.append(rawfilename)
    print('found', len(filelist), 'files')
    return filelist


# now check that the excel file is a wip file
def checkWipfile(filelist):
    filteredfilelist = []
    print('filtering files.....')
    worksheetName = 'Executive Summary'
    projectNameCell = ('A5')

    # go through the filelist and check each file
    for excelfile in filelist:
        # Try to open the WIP excel file
        try:
            wipWorkbook = load_workbook(
                filename=excelfile, read_only=True,
                data_only=True)
            # and try to open the worksheet
            wipWorksheet = wipWorkbook[worksheetName]
            # print 'checking file', excelfile
        except:
            # print excelfile, 'is not a wip file'
            continue
        if (wipWorksheet[projectNameCell].value) != 'Project Name:':
            # print excelfile, 'is not a wip file'
            continue
        else:
            filteredfilelist.append(excelfile)
            # print 'adding', excelfile, 'to wipfile list'
    print('filtered', len(filelist), 'files to', len(filteredfilelist), 'files')
    return filteredfilelist


# function to import wip data from excel file
def importData(wipfilename):

    # set variables for location of various fields
    # e.g. forecast sale, forecast cost, contribution, cost to date etc
    # excel cell references in the following format
    # variable = (excel row, excel column)
    # we will return the data as a dictionary
    # with the variable wipData
    wipData = dict()

    # header variables
    worksheetName = "Executive Summary"
    projectNameRef = ('B5')
    projectNumberRef = ('B6')
    wipDateRef = ('B7')

    # columns for this month and Now
    nowColumn = 'C'
    thisMonthColumn = 'F'

    # sales account variables
    agreedVariationsNoRef = (nowColumn + '20')
    budgetVariationsNoRef = (nowColumn + '21')
    submittedVariationsNoRef = (nowColumn + '22')
    orderValueRef = (thisMonthColumn + '18')
    agreedVariationsValueRef = (thisMonthColumn + '20')
    budgetVariationsValueRef = (thisMonthColumn + '21')
    submittedVariationsValueRef = (thisMonthColumn + '22')

    # Reserve variable
    reserveAgreedRef = (thisMonthColumn + '30')
    reserveBudgetRef = (thisMonthColumn + '31')
    reserveSubmittedRef = (thisMonthColumn + '32')
    contrachargesRef = (thisMonthColumn + '36')

    # cost variables
    currentCostRef = (thisMonthColumn + '43')
    costToCompleteRef = (thisMonthColumn + '44')
    defectProvisionRef = (thisMonthColumn + '48')

    # margin variables
    contractContributionRef = (thisMonthColumn + '55')
    bettermentsRisksRef = (thisMonthColumn + '56')
    managersViewRef = (thisMonthColumn + '57')

    # total variables
    saleSubtotalRef = (thisMonthColumn + '26')
    reserveSubtotalRef = (thisMonthColumn + '34')
    forecastSaleTotalRef = (thisMonthColumn + '38')
    variationsNoTotalRef = (nowColumn + '24')
    forecastCostTotalRef = (thisMonthColumn + '50')
    forecastMarginTotalRef = (thisMonthColumn + '59')

    # cash recovery variables
    latestApplicationRef = (thisMonthColumn + '67')
    totalCertifiedRef = (thisMonthColumn + '72')

    # Extract the data from the WIP excel spreadsheet and set variables

    # Try to open the WIP excel file
    try:
        wipWorkbook = load_workbook(
            filename=wipfilename, read_only=True,
            data_only=True)
        # and try to open the worksheet
        wipWorksheet = wipWorkbook[worksheetName]
        print('Opening workbook', worksheetName, 'in', wipfilename)
    except:
        print(wipfilename, ' does not exist, exiting')
        quit()

    # extract the cell information
    # using the format wipData['field'] = wipWorksheet([fieldRef].value)
    wipData['projectName'] = (wipWorksheet[projectNameRef].value)
    wipData['projectNumber'] = (wipWorksheet[projectNumberRef].value)
    wipData['wipDate'] = str(wipWorksheet[wipDateRef].value)
    wipData['wipDate'] = (wipData['wipDate'].rsplit(' '))[0]
    wipDateFormatted = wipData['wipDate']
    print('Extracting', wipData['projectNumber'], '-',
          wipData['projectName'], 'data for', wipDateFormatted, '\n')

    # extract variation information
    wipData['agreedVariationsNo'] = \
        int(wipWorksheet[agreedVariationsNoRef].value)
    wipData['budgetVariationsNo'] = \
        int(wipWorksheet[budgetVariationsNoRef].value)
    wipData['submittedVariationsNo'] = \
        int(wipWorksheet[submittedVariationsNoRef].value)
    wipData['variationsNoTotal'] = \
        int(wipWorksheet[variationsNoTotalRef].value)
    variationsNoTotalCheck = wipData['agreedVariationsNo'] \
        + wipData['budgetVariationsNo'] \
        + wipData['submittedVariationsNo']
    # check variations no balance
    if wipData['variationsNoTotal'] != variationsNoTotalCheck:
        print('Number of variations incorrect')
        quit()
    else:
        print('Variations.....okay')

    # extract sales information
    wipData['orderValue'] = int(wipWorksheet[orderValueRef].value)
    wipData['agreedVariationsValue'] = \
        int(wipWorksheet[agreedVariationsValueRef].value)
    wipData['budgetVariationsValue'] = \
        int(wipWorksheet[budgetVariationsValueRef].value)
    wipData['submittedVariationsValue'] = \
        int(wipWorksheet[submittedVariationsValueRef].value)
    wipData['saleSubtotal'] = int(wipWorksheet[saleSubtotalRef].value)
    saleSubtotalCheck = wipData['orderValue'] \
        + wipData['agreedVariationsValue'] \
        + wipData['budgetVariationsValue'] \
        + wipData['submittedVariationsValue']
    # check sales values balance
    if abs(saleSubtotalCheck - wipData['saleSubtotal']) > 5:
        print('Sales value incorrect')
        quit()
    else:
        print('Sales.....okay')

    # extract reserve information
    wipData['reserveAgreed'] = int(wipWorksheet[reserveAgreedRef].value)
    wipData['reserveBudget'] = int(wipWorksheet[reserveBudgetRef].value)
    wipData['reserveSubmitted'] = int(wipWorksheet[reserveSubmittedRef].value)
    wipData['contracharges'] = int(wipWorksheet[contrachargesRef].value)
    wipData['forecastSaleTotal'] = int(wipWorksheet[forecastSaleTotalRef].value)
    sumChecksale = wipData['orderValue'] \
        + wipData['agreedVariationsValue'] \
        + wipData['budgetVariationsValue'] \
        + wipData['submittedVariationsValue'] \
        + wipData['reserveAgreed'] \
        + wipData['reserveBudget'] \
        + wipData['reserveSubmitted'] \
        + wipData['contracharges']
    if abs(sumChecksale - wipData['forecastSaleTotal']) > 5:
        print('Forecast sale subtotal incorrect')
        quit()
    else:
        print('Forecast sale.....okay')

    # extract cost information
    wipData['currentCost'] = int(wipWorksheet[currentCostRef].value)
    wipData['costToComplete'] = int(wipWorksheet[costToCompleteRef].value)
    try:
        wipWorksheet[defectProvisionRef].value
        wipData['defectProvision'] = int(wipWorksheet[defectProvisionRef].value)
    except TypeError:
        wipData['defectProvision'] = 0
    wipData['forecastCostTotal'] = int(wipWorksheet[forecastCostTotalRef].value)
    sumCheckcost = wipData['currentCost'] \
        + wipData['costToComplete'] \
        + wipData['defectProvision']
    if abs(sumCheckcost - wipData['forecastCostTotal']) > 5:
        print('Forecast cost subtotal incorrect')
        quit()
    else:
        print('Forecast cost.....okay')

    # extract margin information
    wipData['contractContribution'] = \
        int(wipWorksheet[contractContributionRef].value)
    wipData['bettermentsRisks'] = int(wipWorksheet[bettermentsRisksRef].value)

    # check if the managersView is blank and would cause an error
    # if true add a 0
    try:
        wipData['managersView'] = int(wipWorksheet[managersViewRef].value)
    except:
        wipData['managersView'] = 0

    wipData['forecastMarginTotal'] = \
        int(wipWorksheet[forecastMarginTotalRef].value)
    sumCheckmargin = wipData['contractContribution'] \
        + wipData['bettermentsRisks'] \
        + wipData['managersView']
    if abs(sumCheckmargin - wipData['forecastMarginTotal']) > 5:
        print('Forecast margin subtotal incorrect')
        quit()
    else:
        print('Forecast margin.....okay')

    # extract cash information
    # if the field is blank insert a 0
    try:
        wipData['latestApplication'] = \
            int(wipWorksheet[latestApplicationRef].value)
    except:
        wipData['latestApplication'] = 0
    try:
        wipData['totalCertified'] = int(wipWorksheet[totalCertifiedRef].value)
    except:
        wipData['totalCertified'] = 0

    print('All data imported sucessfully\n')
    return wipData


# function to import data into an SQL database
def exportDataSql(dataList):
    # print dataList
    projectName = dataList['projectName']
    wipDate = dataList['wipDate']
    conn = sqlite3.connect(os.path.normpath(DATABASE_NAME))
    cur = conn.cursor()

    # delete the database table if it exists...for testing only
    # cur.execute('''DROP TABLE IF EXISTS wipdata''')

    # create job name table
    cur.execute('''
        CREATE TABLE IF NOT EXISTS projectname (
        id INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT UNIQUE,
        name TEXT UNIQUE
        )''')

    # create wipdata table if doesn't exist
    cur.execute('''
    CREATE TABLE IF NOT EXISTS wipdata (
        projectNumber TEXT, projectName INTEGER, wipDate TEXT,
        agreedVariationsNo INTEGER, budgetVariationsNo INTEGER,
        submittedVariationsNo INTEGER, variationsNoTotal INTEGER,
        orderValue INTEGER, agreedVariationsValue INTEGER,
        budgetVariationsValue INTEGER, submittedVariationsValue INTEGER,
        saleSubtotal INTEGER, reserveAgreed INTEGER, reserveBudget INTEGER,
        reserveSubmitted INTEGER, contracharges INTEGER,
        forecastSaleTotal INTEGER, currentCost INTEGER, costToComplete INTEGER,
        defectProvision INTEGER, forecastCostTotal INTEGER,
        contractContribution INTEGER, bettermentsRisks INTEGER,
        managersView INTEGER, forecastMarginTotal INTEGER,
        latestApplication INTEGER, totalCertified INTEGER,
        PRIMARY KEY (projectNumber, wipDate));''')

    # check if the project name exists in the table project name
    cur.execute('''INSERT OR IGNORE INTO projectname (name)
        VALUES ( ? )''', (projectName, ))
    cur.execute('SELECT id FROM projectname WHERE name = ?', (projectName, ))
    projectName_id = cur.fetchone()[0]
    dataList['projectName'] = projectName_id

    # run through the data and allocate to each field
    cur.executemany('''INSERT OR REPLACE INTO wipdata
        (projectNumber, projectName, wipDate, agreedVariationsNo,
        budgetVariationsNo, submittedVariationsNo, variationsNoTotal,
        orderValue, agreedVariationsValue, budgetVariationsValue,
        submittedVariationsValue, saleSubtotal, reserveAgreed,
        reserveBudget, reserveSubmitted, contracharges,
        forecastSaleTotal, currentCost, costToComplete,
        defectProvision, forecastCostTotal, contractContribution,
        bettermentsRisks, managersView, forecastMarginTotal,
        latestApplication, totalCertified)
    VALUES
        (:projectNumber, :projectName, :wipDate, :agreedVariationsNo,
        :budgetVariationsNo, :submittedVariationsNo, :variationsNoTotal,
        :orderValue, :agreedVariationsValue, :budgetVariationsValue,
        :submittedVariationsValue, :saleSubtotal, :reserveAgreed,
        :reserveBudget, :reserveSubmitted, :contracharges,
        :forecastSaleTotal, :currentCost, :costToComplete,
        :defectProvision, :forecastCostTotal, :contractContribution,
        :bettermentsRisks, :managersView, :forecastMarginTotal,
        :latestApplication, :totalCertified)''', [dataList])

    # Commit the changes to the database
    conn.commit()
    cur.close()
    conn.close()


# Call main routine
# main()
