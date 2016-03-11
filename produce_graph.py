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
    searchData = inputProjectNumber()
    graphData = importDataSql(searchData)
    plotGraph(graphData)

# function to read data from SQL database based on job number & months
def importDataSql(searchData):
    # temp variable to import the data we want
    # we will get this variable from the routine call in future
    request = "wipdate, projectNumber, projectname.name, forecastCostTotal, forecastSaleTotal, forecastMarginTotal"
    # print request
    
    # split out searchData and connect to database
    projectNumber = searchData[0]
    months = searchData[1]
    conn = sqlite3.connect('wipdatadb.sqlite')
    cur = conn.cursor()

    # extract seachdata
    cur.execute('''
    SELECT ''' + request + ''' 
    FROM wipdata 
    JOIN projectName
    ON wipdata.projectName = projectname.id
    WHERE projectNumber = :projectNumber 
    ORDER BY wipdate
    DESC LIMIT :months  
    ''', {'projectNumber': projectNumber, 'months': months})

    all_rows = cur.fetchall()
    # print all_rows
    # print len(all_rows) 
    
    return all_rows

# function to request job number to graph and months history
# i.e. LIMIT default graph last 12 months
def inputProjectNumber():
    while True:
        inp = raw_input('Enter the project to graph: ')
        if not re.search(r"\d{5}", inp):
            print 'Enter a valid job number:'
            continue
        else:
            projectNumber = str(inp)
            break

    while True:
        inp = raw_input('Enter the range: ')
        if not re.search(r"\d{1,2}", inp):
            print 'range between 1 and 24 months'
            continue
        else:
            months = inp
            break

    return (projectNumber, months)



# function to produce and output graph
def plotGraph(graphData):

    # create the lists
    dates = []
    projectNumber = []
    projectName = []
    forecastCost = []
    forecastSale = []
    forecastContribution = []

    # loop through the list graphData to extract the x and y axis
    # data is produced in the folllowing format
    # [(u'2016-01-31', -1355036), (u'2015-12-31', -1354858), 
    for data in graphData:
        dates.append(data[0])
        projectNumber.append(data[1])
        projectName.append(data[2])
        forecastCost.append(data[3]/1000)
        forecastSale.append(data[4]/1000)
        forecastContribution.append(data[5]/1000)

    # format the dates in the correct format to show
    dates = [datetime.datetime.strptime(d, '%Y-%m-%d').date() for d in dates]
 
    # These are the "Tableau 20" colors as RGB.  
    tableau20 = [(31, 119, 180), (174, 199, 232), (255, 127, 14), (255, 187, 120),  
                 (44, 160, 44), (152, 223, 138), (214, 39, 40), (255, 152, 150),  
                 (148, 103, 189), (197, 176, 213), (140, 86, 75), (196, 156, 148),  
                 (227, 119, 194), (247, 182, 210), (127, 127, 127), (199, 199, 199),  
                 (188, 189, 34), (219, 219, 141), (23, 190, 207), (158, 218, 229)]
    
    # Scale the RGB values to the [0, 1] range, which is the format matplotlib accepts.  
    for i in range(len(tableau20)):  
        r, g, b = tableau20[i]  
        tableau20[i] = (r / 255., g / 255., b / 255.)  

    # You typically want your plot to be ~1.33x wider than tall. This plot is a rare  
    # exception because of the number of lines being plotted on it.  
    # Common sizes: (10, 7.5) and (12, 9)  
    plt.figure(figsize=(12, 9))
    #fig = plt.figure() 

    # Remove the plot frame lines
    ax1 = plt.subplot(111)  
    ax1.spines["top"].set_visible(False)  
    ax1.spines["bottom"].set_visible(False)  
    ax1.spines["right"].set_visible(False)  
    ax1.spines["left"].set_visible(False)  

    # set xlim to show all dates
    plt.xlim(dates[-1], dates[0]+ (datetime.timedelta(days=32)))

    # Ensure that the axis ticks only show up on the bottom and left of the plot.  
    # Ticks on the right and top of the plot are generally unnecessary chartjunk.  
    ax1.get_xaxis().tick_bottom()  
    ax1.get_yaxis().tick_left() 
   
    # set the y axis to format values with commas, i.e. thousands

    # plot the data
    forecastCostLine = plt.plot(dates, forecastCost, lw=2.5, color=tableau20[0], label='Cost')
    forecastSaleLine = plt.plot(dates, forecastSale, lw=2.5, color=tableau20[1], label='Sale')
    forecastContributionLine = plt.plot(dates, forecastContribution, lw=2.5, color=tableau20[2], label='Contribution')
    
    # add a text label to the right end of every drawn line
    # first get the last value of each line
    yPosForecastCost = forecastCost[0]
    yPosForecastSale = forecastSale[0]
    yPosForecastContribution = forecastContribution[0]
    print dates[0]
    xPosLabel = (dates[0] + (datetime.timedelta(days=7)))

    plt.text(xPosLabel, yPosForecastCost, 'Cost', fontsize=14, color=tableau20[0], verticalalignment='center')
    plt.text(xPosLabel, yPosForecastSale, 'Sale', fontsize=14, color=tableau20[1], verticalalignment='center')
    plt.text(xPosLabel, yPosForecastContribution, 'Contribution', fontsize=14, color=tableau20[2], verticalalignment='center')

    # set the y axis label
    ylabel = u'\xA3/k'
    ax1.set_ylabel(ylabel, fontsize=14, rotation='vertical')
    
    # set the title of the graph
    title = projectName[0] + ' (' + projectNumber[0] + ')' + '\nForecast Sale, Cost and Contribution'
    # print 'title set to', title
    ax1.set_title(title, fontsize=17, ha='center')

    # add the legend
    # ax1.legend(frameon=False)
    plt.gcf().autofmt_xdate()

    # data source and notice
    # yPosNotice = min(min(forecastCost), min(forecastSale), min(forecastContribution))
    # plt.text(min(dates), yPosNotice, 'Data source: H:\Previous Months WIPS' 
    #        '\nAuthor: Ian Hill (ian.hill@interserve.com)', fontsize=10)
    
    # plot the graph and save
    # plt.show()
    plt.savefig((projectNumber[0]+' contribution graph.png'), bbox_inches='tight')

# call main routine
main()
