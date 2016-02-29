# import WIP data from WIP file

# import libraries 
from openpyxl import load_workbook
import time
import sqlite3

# supress warnings
import warnings 
warnings.filterwarnings("ignore")

# declare WIP excel file name here, although we will create a list later
wipfilename = "WIP Data 50627.xlsx"

# main routine
def main():
    # extract the data from the excel worksheet cells
    wipData = importData(wipfilename)
    print wipData['projectNumber']
    exportDataSql(wipData)

# function to import wip data from excel file
def importData(filename):

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
        wipWorkbook = load_workbook(filename=wipfilename, read_only=True, data_only=True)
        # and try to open the worksheet
        wipWorksheet = wipWorkbook[worksheetName]
        print 'Opening workbook', worksheetName, 'in', wipfilename
    except:
        print wipfilename, ' does not exist, exiting'
        quit()

    # extract the cell information
    # using the format wipData['field'] = wipWorksheet([fieldRef].value)
    wipData['projectName'] = (wipWorksheet[projectNameRef].value)
    wipData['projectNumber'] = (wipWorksheet[projectNumberRef].value) 
    wipData['wipDate'] = (wipWorksheet[wipDateRef].value)
    wipDateFormatted = wipData['wipDate']
    print 'Extracting', wipData['projectNumber'], '-', wipData['projectName'], 'data for', wipDateFormatted,'\n'

    # extract variation information
    wipData['agreedVariationsNo'] = int(wipWorksheet[agreedVariationsNoRef].value)
    wipData['budgetVariationsNo'] = int(wipWorksheet[budgetVariationsNoRef].value)
    wipData['submittedVariationsNo'] = int(wipWorksheet[submittedVariationsNoRef].value)
    wipData['variationsNoTotal'] = int(wipWorksheet[variationsNoTotalRef].value)
    variationsNoTotalCheck = wipData['agreedVariationsNo'] + wipData['budgetVariationsNo'] + wipData['submittedVariationsNo']
    # check variations no balance
    if wipData['variationsNoTotal'] != variationsNoTotalCheck:
        print 'Number of variations incorrect'
        quit ()
    else:
        print 'Variations.....okay'

    # extract sales information
    wipData['orderValue'] = int(wipWorksheet[orderValueRef].value)
    wipData['agreedVariationsValue'] = int(wipWorksheet[agreedVariationsValueRef].value)
    wipData['budgetVariationsValue'] = int(wipWorksheet[budgetVariationsValueRef].value)
    wipData['submittedVariationsValue'] = int(wipWorksheet[submittedVariationsValueRef].value)
    wipData['saleSubtotal'] = int(wipWorksheet[saleSubtotalRef].value)
    saleSubtotalCheck = wipData['orderValue'] + wipData['agreedVariationsValue'] + wipData['budgetVariationsValue'] + wipData['submittedVariationsValue']
    # check sales values balance
    if abs(saleSubtotalCheck - wipData['saleSubtotal']) > 5:
        print 'Sales value incorrect'
        quit()
    else:
        print 'Sales.....okay'

    # extract reserve information
    wipData['reserveAgreed'] = int(wipWorksheet[reserveAgreedRef].value)
    wipData['reserveBudget'] = int(wipWorksheet[reserveBudgetRef].value)
    wipData['reserveSubmitted'] = int(wipWorksheet[reserveSubmittedRef].value)
    wipData['contracharges'] = int(wipWorksheet[contrachargesRef].value)
    wipData['forecastSaleTotal'] = int(wipWorksheet[forecastSaleTotalRef].value)
    sumChecksale = wipData['orderValue'] + wipData['agreedVariationsValue'] + wipData['budgetVariationsValue'] + wipData['submittedVariationsValue'] + wipData['reserveAgreed'] + wipData['reserveBudget'] + wipData['reserveSubmitted'] + wipData['contracharges']
    if abs(sumChecksale - wipData['forecastSaleTotal']) > 5:
        print 'Forecast sale subtotal incorrect'
        quit()
    else:
        print 'Forecast sale.....okay'

    # extract cost information
    wipData['currentCost'] = int(wipWorksheet[currentCostRef].value)
    wipData['costToComplete'] = int(wipWorksheet[costToCompleteRef].value)
    wipData['defectProvision'] = int(wipWorksheet[defectProvisionRef].value)
    wipData['forecastCostTotal'] = int(wipWorksheet[forecastCostTotalRef].value)
    sumCheckcost = wipData['currentCost'] + wipData['costToComplete'] + wipData['defectProvision']
    if abs(sumCheckcost - wipData['forecastCostTotal']) > 5:
        print 'Forecast cost subtotal incorrect'
        quit()
    else:
        print 'Forecast cost.....okay'

    # extract margin information
    wipData['contractContribution'] = int(wipWorksheet[contractContributionRef].value)
    wipData['bettermentsRisks'] = int(wipWorksheet[bettermentsRisksRef].value)
    try:
        wipData['managersView'] = int(wipWorksheet[managersViewRef].value)
    except:
        wipData['managersView'] = 0
    wipData['forecastMarginTotal'] = int(wipWorksheet[forecastMarginTotalRef].value)
    sumCheckmargin = wipData['contractContribution'] + wipData['bettermentsRisks'] + wipData['managersView']
    if abs(sumCheckmargin - wipData['forecastMarginTotal']) > 5:
        print 'Forecast margin subtotal incorrect'
        quit()
    else:
        print 'Forecast margin.....okay'


    # extract cash information
    wipData['latestApplication'] = int(wipWorksheet[latestApplicationRef].value)
    wipData['totalCertified'] = int(wipWorksheet[totalCertifiedRef].value)

    print 'All data imported sucessfully\n'
    return wipData 

#function to import data into an SQL database
def exportDataSql(dataList):
    print dataList
    conn = sqlite3.connect('wipdatadb.sqlite')
    cur = conn.cursor()

    # delete the database table if it exists...for testing only
    cur.execute('''DROP TABLE IF EXISTS wipdata''')

    # create the database
    cur.execute('''
    CREATE TABLE wipdata (
    projectNumber TEXT, projectName INTEGER, wipDate TEXT, agreedVariationsNo INTEGER,
    budgetVariationsNo INTEGER, submittedVariationsNo INTEGER, variationsNoTotal INTEGER,
    orderValue INTEGER, agreedVariationsValue INTEGER, budgetVariationsValue INTEGER,
    submittedVariationsValue INTEGER, saleSubtotal INTEGER, reserveAgreed INTEGER,
    reserveBudget INTEGER, reserveSubmitted INTEGER, contracharges INTEGER,
    forecastSaleTotal INTEGER, currentCost INTEGER, costToComplete INTEGER,
    defectProvision INTEGER, forecastCostTotal INTEGER, contractContribution INTEGER,
    bettermentsRisks INTEGER, managersView INTEGER, forecastMarginTotal INTEGER, 
    latestApplication INTEGER, totalCertified INTEGER)''')

    # run through the data and allocate to each field
    cur.executemany('''INSERT INTO wipdata
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
main()

