# produce graph from database


# import libraries
import sqlite3
import re

# main routine
def main():
    searchdata = inputJobNo()
    importDataSql(searchdata)

# function to read data from SQL database based on job number & months
def importDataSql(searchdata):
    # split out searchdata and connect to database
    jobNo = searchdata[0]
    months = searchdata[1]
    conn = sqlite3.connect('wipdatadb.sqlite')
    cur = conn.cursor()

    # extract seachdata
    cur.execute('''
    SELECT * FROM wipdata 
    JOIN projectName ON wipdata.projectName = projectname.id
    WHERE projectNumber = ?
    ORDER BY wipdate
    DESC LIMIT ?  
    ''', (jobNo, months))

    all_rows = cur.fetchall()
    print all_rows
    print len(all_rows) 

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
# call main routine
main()
