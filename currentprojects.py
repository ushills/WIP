# produce list of current projects from database

# import libraries
import sqlite3

# main routine
def main():
    recentProjects = recentProjectList(mostRecentWip())
    # print recentProjects
    # print len(recentProjects)
    for project in recentProjects:
        print project

def mostRecentWip():

    # connect to the datebase
    conn = sqlite3.connect('wipdatadb.sqlite')
    cur = conn.cursor()

    # extract the most recent wipdate
    cur.execute('''
    SELECT max(wipDate) as latestdate FROM wipdata
    ''')

    latestDate = cur.fetchall()
    #print (latestdate)

    return latestDate

def recentProjectList(searchDate):
    projects = []
    projectList = []

    # connect to the database
    conn = sqlite3.connect('wipdatadb.sqlite')
    cur = conn.cursor()

    # extract the list of most recent wips using the searchDate
    cur.execute('''
    SELECT projectNumber AS wipsThisMonth
    FROM wipdata
    WHERE (JulianDay(date('2016-12-05')) - JulianDay(wipDate)) < 35
    GROUP BY projectNumber''')

    projectList = cur.fetchall()

    for i in range(len(projectList)):
        projects.append(projectList[i][0])

    return projects



# call main routine
main()
