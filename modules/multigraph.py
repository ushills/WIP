# produce graph from database


# import libraries
import sqlite3
import matplotlib.pyplot as plt
from matplotlib.dates import MonthLocator, DateFormatter
import datetime
import os


# main routine
def printgraphs(DBNAME, directory):
    global DATABASE_NAME
    global outputdirectory
    DATABASE_NAME = (DBNAME)
    outputdirectory = directory
    plotGraphs()


# function to plot graphs - primary function
def plotGraphs():

    monthsToPlot = '12'

    # extract the most recent wip date from the database
    recentWipDate = mostRecentWip(DATABASE_NAME)

    # extract the projectList based on the most recent wip date
    projectList = recentProjectList(recentWipDate, DATABASE_NAME)

    # iterate through the projectList and plot graphs to each project
    sucessful = 0
    failed = 0
    for project in projectList:
        try:
            searchData = (project, monthsToPlot)
            # print 'plotting graphs for', searchData[0]

            # plot the Forecast Data graph
            request = "wipdate, projectNumber, projectname.name, \
                       forecastCostTotal, forecastSaleTotal, \
                       forecastMarginTotal, currentCost, totalCertified"
            graphData = importDataSql(searchData, request, DATABASE_NAME)
            plotForecastGraph(graphData)

            # plot the variation histogram
            request = "wipdate, projectNumber, projectname.name, \
                       agreedVariationsNo, budgetVariationsNo, \
                       submittedVariationsNo, agreedVariationsValue, \
                       budgetVariationsValue, submittedVariationsValue"
            graphData = importDataSql(searchData, request, DATABASE_NAME)
            plotVariationGraph(graphData)
            sucessful += 1

        except:
            print('skipping', searchData[0], 'data incorrect')
            print('-' * 25)
            failed += 1
            continue

    print('printed', sucessful, 'project graphs')
    if failed > 0:
        print(failed, 'projects failed to print')


# function to extract date of most recent wip in database
def mostRecentWip(database):

    # connect to the datebase
    try:
        conn = sqlite3.connect(os.path.normpath(database))
    except NameError:
        print("Database file", database, "does not exist")
        print("Failed in mostRecentWip")
    else:
        cur = conn.cursor()
        # extract the most recent wipdate
        cur.execute('''
        SELECT max(wipDate) as latestdate FROM wipdata
        ''')

        latestDate = cur.fetchone()
        # extract the latestDate string from the list returned
        latestDate = str(latestDate[0])

        return latestDate


# function to extract list of most recent wips
def recentProjectList(searchDate, database):
    projects = []

    # connect to the database
    try:
        conn = sqlite3.connect(os.path.normpath(database))
    except NameError:
        print("Database file", database, "does not exist")
        print("Failed in recentProjectList")
    else:
        cur = conn.cursor()

        # extract the list of most recent wips using the searchDate
        cur.execute('''
        SELECT projectNumber AS wipsThisMonth
        FROM wipdata
        WHERE (JulianDay(:searchDate) - JulianDay(wipDate)) < 35
        GROUP BY projectNumber''', {'searchDate': searchDate})

        projectList = cur.fetchall()

        for i in range(len(projectList)):
            projects.append(projectList[i][0])

        return projects


# function to read data from SQL database based on job number & months
def importDataSql(searchData, request, database):
    projectNumber = searchData[0]
    months = searchData[1]
    conn = sqlite3.connect(os.path.normpath(database))
    cur = conn.cursor()

    # extract seachdata
    # reverse sort order newest to oldest
    cur.execute('''
    SELECT * FROM (
    SELECT ''' + request + '''
    FROM wipdata
    JOIN projectName
    ON wipdata.projectName = projectname.id
    WHERE projectNumber = :projectNumber
    AND forecastSaleTotal > 0
    ORDER BY wipdate DESC
    LIMIT :months)
    ORDER BY wipdate ASC
    ''', {'projectNumber': projectNumber, 'months': months})
    return cur.fetchall()


# function to produce and output forecast graph
def plotForecastGraph(graphData):

    # Set the common variables

    # These are the "Tableau 20" colors as RGB.
    tableau20 = [(31, 119, 180), (174, 199, 232), (255, 127, 14),
                 (255, 187, 120), (44, 160, 44), (152, 223, 138),
                 (214, 39, 40), (255, 152, 150), (148, 103, 189),
                 (197, 176, 213), (140, 86, 75), (196, 156, 148),
                 (227, 119, 194), (247, 182, 210), (127, 127, 127),
                 (199, 199, 199), (188, 189, 34), (219, 219, 141),
                 (23, 190, 207), (158, 218, 229)]

    # Scale the RGB values to the [0, 1] range,
    # which is the format matplotlib accepts.
    for i in range(len(tableau20)):
        r, g, b = tableau20[i]
        tableau20[i] = (r / 255., g / 255., b / 255.)

    # You typically want your plot to be ~1.33x wider than tall.
    # This plot is a rare exception because of the number of lines
    # being plotted on it.
    # Common sizes: (10, 7.5) and (12, 9)
    plt.figure(figsize=(12, 9))
    # fig = plt.figure()

    # change to classic matplotlib styles
    plt.style.use('classic')

    fig, (ax, ax3) = plt.subplots(nrows=2)
    ax2 = ax.twinx()

    # Create the forcast totals graph
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

    # Ensure that the axis ticks only show up on the bottom and
    # left of the plot.
    # Ticks on the right and top of the plot are generally
    # unnecessary chartjunk
    ax.get_xaxis().tick_bottom()
    ax.get_yaxis().tick_left()
    ax2.get_xaxis().tick_bottom()
    ax2.get_yaxis().tick_right()

    # plot the data
    forecastSaleLine = ax.plot(
        dates, forecastSale, lw=2.5,
        color=tableau20[5], label='Forecast Sale')

    forecastCostLine = ax.plot(
        dates, forecastCost, lw=2.5,
        color=tableau20[0], label='Forecast Cost')

    forecastContributionLine = ax2.plot(
        dates, forecastContribution,
        lw=2.5, color=tableau20[2],
        label='Contribution')

    # set the y axis label
    ylabel = '\xA3k'
    ax.set_ylabel(ylabel, fontsize=14, rotation='vertical')
    ax2.set_ylabel(
        ylabel, fontsize=14, rotation='vertical',
        color=tableau20[2])

    for tl in ax2.get_yticklabels():
        tl.set_color(tableau20[2])

    # set the title of the graph
    title = (projectName[0] + ' (' + projectNumber[0] + ')\n')
    ax.set_title(title, fontsize=16, ha='center')
    plt.gcf().autofmt_xdate()

    # create the to-date graph
    # create the lists
    dates = []
    projectNumber = []
    projectName = []
    currentCost = []
    totalCertified = []

    # loop through the list graphData to extract the x and y axis
    # data is produced in the folllowing format
    # [(u'2016-01-31', -1355036), (u'2015-12-31', -1354858),
    for data in graphData:
        dates.append(data[0])
        projectNumber.append(data[1])
        projectName.append(data[2])
        currentCost.append(data[6]/1000)
        totalCertified.append(data[7]/1000)

    # format the dates in the correct format to show
    dates = [datetime.datetime.strptime(d, '%Y-%m-%d').date() for d in dates]

    # format the x axis date format
    months = MonthLocator()
    monthFormat = DateFormatter('%b-%Y')
    ax3.fmt_xdata = DateFormatter('%b-%Y')
    ax3.xaxis.set_major_locator(months)
    ax3.xaxis.set_major_formatter(monthFormat)
    ax.xaxis.set_major_locator(months)
    ax.xaxis.set_major_formatter(monthFormat)

    # create the axis
    # ax3 = plt.subplot(212)

    # Ensure that the axis ticks only show up on the bottom and
    # left of the plot.
    # Ticks on the right and top of the plot are generally
    # unnecessary chartjunk.
    ax3.get_xaxis().tick_bottom()
    ax3.get_yaxis().tick_left()

    # plot the data
    currentCostline = ax3.plot(
        dates, currentCost, lw=2.5,
        color=tableau20[11], label='Cost')
    totalCertifiedline = ax3.plot(
        dates, totalCertified, lw=2.5,
        color=tableau20[6], label='Certified')

    # set the y axis label
    ylabel = '\xA3k'
    ax3.set_ylabel(ylabel, fontsize=14, rotation='vertical')

    # add the legend
    # shrink the axis height by 10% at the bottom
    box = ax3.get_position()
    ax3.set_position([
                      box.x0, box.y0 + box.height * 0.1,
                      box.width, box.height * 0.9])

    # combine ax1 and ax2 labels
    lines, labels = ax.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    lines3, labels3 = ax3.get_legend_handles_labels()

    # place the legend below the axis
    ax3.legend(
        lines + lines2 + lines3, labels + labels2 + labels3,
        loc='upper center', fontsize=10, bbox_to_anchor=(0.5, -0.35),
        fancybox=True, shadow=True, ncol=5)

    # plot the graph and save
    # plt.show()
    plt.savefig(
        os.path.normpath(
            outputdirectory +
            projectNumber[0] + ' forecast totals graph.png'),
        bbox_inches='tight')
    plt.close('all')


# function to produce and output variation graph
def plotVariationGraph(graphData):

    # Set the common variables

    # These are the "Tableau 20" colors as RGB.
    tableau20 = [(31, 119, 180), (174, 199, 232), (255, 127, 14),
                 (255, 187, 120), (44, 160, 44), (152, 223, 138),
                 (214, 39, 40), (255, 152, 150), (148, 103, 189),
                 (197, 176, 213), (140, 86, 75), (196, 156, 148),
                 (227, 119, 194), (247, 182, 210), (127, 127, 127),
                 (199, 199, 199), (188, 189, 34), (219, 219, 141),
                 (23, 190, 207), (158, 218, 229)]

    # Scale the RGB values to the [0, 1] range,
    # which is the format matplotlib accepts.
    for i in range(len(tableau20)):
        r, g, b = tableau20[i]
        tableau20[i] = (r / 255., g / 255., b / 255.)

    # You typically want your plot to be ~1.33x wider than tall.
    # This plot is a rare exception because of the number of lines
    # being plotted on it.
    # Common sizes: (10, 7.5) and (12, 9)

    plt.figure(figsize=(12, 9))
    fig, (ax, ax2) = plt.subplots(nrows=2)

    # change to classic matplotlib styles
    plt.style.use('classic')

    # Create the forcast totals graph
    # create the lists
    dates = []
    projectNumber = []
    projectName = []
    agreedVariationNo = []
    budgetVariationNo = []
    submittedVariationNo = []
    agreedVariationValue = []
    budgetVariationValue = []
    submittedVariationValue = []

    # loop through the list graphData to extract the x and y axis
    # data is produced in the folllowing format
    # [(u'2016-01-31', -1355036), (u'2015-12-31', -1354858),
    for data in graphData:
        dates.append(data[0])
        projectNumber.append(data[1])
        projectName.append(data[2])
        agreedVariationNo.append(data[3])
        budgetVariationNo.append(data[4])
        submittedVariationNo.append(data[5])
        agreedVariationValue.append(data[6]/1000)
        budgetVariationValue.append(data[7]/1000)
        submittedVariationValue.append(data[8]/1000)

    # create the bins for the bar chart and set width of bar
    bins = list(range(len(dates)))
    widthDate = 0.9

    # format the dates in the correct format to show
    dates = [datetime.datetime.strptime(d, '%Y-%m-%d').date() for d in dates]

    # convert the dates to MMM-YYY
    xdate = []
    for d in dates:
        d = d.strftime('%b-%Y')
        xdate.append(d)
    dates = xdate

    # Ensure that the axis ticks only show up on the bottom and
    # left of the plot.
    # Ticks on the right and top of the plot are generally
    # unnecessary chartjunk.
    ax.get_xaxis().tick_bottom()
    ax.get_yaxis().tick_left()

    # plot the data
    # as this is a stacked graph first we need to zip the arrays for the
    # bottom statement
    cumulativeNoHist = [a+b for a, b in zip(agreedVariationNo,
                                            submittedVariationNo)]

    # plot the bars
    agreedVariationNoHist = ax.bar(
        bins, agreedVariationNo, width=widthDate,
        color=tableau20[5], label='Agreed')

    submittedVariationNoHist = ax.bar(
        bins, submittedVariationNo, bottom=agreedVariationNo,
        width=widthDate, color=tableau20[15], label='Submitted')

    budgetVariationNoHist = ax.bar(
        bins, budgetVariationNo, bottom=cumulativeNoHist,
        width=widthDate, color=tableau20[11], label='Budget')

    # set the y axis label
    ylabel = 'No of Variations'
    ax.set_ylabel(ylabel, fontsize=14, rotation='vertical')

    # set the title of the graph
    title = (projectName[0] + ' (' + projectNumber[0] + ')\n')
    ax.set_title(title, fontsize=16, ha='center')

    # Ensure that the axis ticks only show up on the bottom and
    # left of the plot.
    # Ticks on the right and top of the plot are generally
    # unnecessary chartjunk.
    ax2.get_xaxis().tick_bottom()
    ax2.get_yaxis().tick_left()

    # plot the data
    # first set the width of the bins for side by side bar chart
    widthDate = 0.3
    cumulativebin = [x + widthDate for x in bins]
    cumulativebin2 = [x + (2 * widthDate) for x in bins]

    # plot the bars
    agreedVariationValueHist = ax2.bar(
        bins, agreedVariationValue, width=widthDate,
        color=tableau20[5], label='Agreed')

    submittedVariationValueHist = ax2.bar(
        cumulativebin, submittedVariationValue, width=widthDate,
        color=tableau20[15], label='Submitted')

    budgetVariationValueHist = ax2.bar(
        cumulativebin2, budgetVariationValue, width=widthDate,
        color=tableau20[11], label='Budget')

    # set the y axis label
    ylabel = 'Value of Variations\n(\xA3k)'
    ax2.set_ylabel(ylabel, fontsize=14, rotation='vertical')

    # set the x axis label
    ax2.set_xticks(cumulativebin)
    ax2.set_xticklabels(dates)
    plt.gcf().autofmt_xdate()

    # add the legend
    # shrink the axis height by 10% at the bottom
    box = ax2.get_position()
    ax2.set_position([
        box.x0, box.y0 + box.height * 0.1,
        box.width, box.height * 0.9])

    # combine ax1 and ax2 labels
    lines, labels = ax.get_legend_handles_labels()

    # place the legend below the axis
    ax2.legend(
        lines, labels, loc='upper center', fontsize=10,
        bbox_to_anchor=(0.5, -0.35), fancybox=True,
        shadow=True, ncol=5)

    # plot the graph and save
    # plt.show()
    plt.savefig(
        os.path.normpath(outputdirectory +
                         projectNumber[0] +
                         ' variations graph.png'),
        bbox_inches='tight')
    plt.close('all')
