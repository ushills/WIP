from modules.multigraph import printgraphs

# printgraphs taken arguements in the form
# printgraphs(DATABASE_NAME, outputdirectory)

# print graphs for Midlands
print 'Printing Midlands graphs'
printgraphs('./midlands/midlandswipdata.sqlite', './midlands/graphs/')

# print graphs for London
print 'Printing London graphs'
printgraphs('./london/londonwipdata.sqlite', './london/graphs/')
