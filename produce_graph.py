# produce graph from database


# import libraries
import sqlite3
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import datetime
#import numpy as np
import re

# main routine
def main():
    searchData = inputJobNo()
    graphData = importDataSql(searchData)
    print graphData
    plotGraph(graphData)

# function to read data from SQL database based on job number & months
def importDataSql(searchData):
    # split out searchData and connect to database
    jobNo = searchData[0]
    months = searchData[1]
    conn = sqlite3.connect('wipdatadb.sqlite')
    cur = conn.cursor()

    # extract seachdata
    cur.execute('''
    SELECT wipDate, forecastMarginTotal FROM wipdata 
    JOIN projectName ON wipdata.projectName = projectname.id
    WHERE projectNumber = ?
    ORDER BY wipdate
    DESC LIMIT ?  
    ''', (jobNo, months))

    all_rows = cur.fetchall()
    print all_rows
    print len(all_rows) 
    
    return all_rows

# function to request job number to graph and months history
# i.e. LIMIT default graph last 12 months
def inputJobNo():
    while True:
        inp = raw_input('Enter the project to graph: ')
        if not re.search(r"\d{5}", inp):
            print 'Enter a valid job number:'
            continue
        else:
            jobNo = str(inp)
            break

    while True:
        inp = raw_input('Enter the range: ')
        if not re.search(r"\d{1,2}", inp):
            print 'range between 1 and 24 months'
            continue
        else:
            months = inp
            break

    return (jobNo, months)



# function to produce and output graph
def plotGraph(graphData):

    dates = []
    y = []

# loop through the list graphData to extract the x and y axis
    for data in graphData:
        dates.append(data[0])
        y.append(data[1])

    print dates
    print y

    x = [datetime.datetime.strptime(d, '%Y-%m-%d').date() for d in dates]
    print x
    #y = range(len(x))
    #plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%d/%m/%Y'))
    #plt.gca().xaxis.set_major_formatter(mdates.DayLocator())
# data is produced in the folllowing format
# [(u'2016-01-31', -1355036), (u'2015-12-31', -1354858), 
# (u'2015-11-30', -1233022), (u'2015-10-31', -983396), 
# (u'2015-09-30', -896752), (u'2015-08-31', -863247), 
# (u'2015-07-31', -678451), (u'2015-06-30', -563631), 
# (u'2015-05-31', -644057), (u'2015-04-30', -480037), 
# (u'2015-03-31', -331675), (u'2015-02-28', -371312)]
    plt.plot(x, y)
    plt.gcf().autofmt_xdate()
    plt.show
    plt.savefig('contribution.png', bbox_inches='tight')

# call main routine
main()
