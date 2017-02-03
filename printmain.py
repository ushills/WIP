from modules.multigraph import printgraphs

# printgraphs taken arguements in the form
# printgraphs(DATABASE_NAME, outputdirectory)

# print graphs for Midlands
print('Printing Midlands graphs')
printgraphs('./midlands/midlandswipdata.sqlite', './midlands/graphs/')

# print graphs for London
print('Printing London graphs')
printgraphs('./london/londonwipdata.sqlite', './london/graphs/')

# print graphs for North East
print('Printing North East graphs')
printgraphs('./northeast/northeastwipdata.sqlite', './northeast/graphs/')

# print graphs for North West
print('Printing North West graphs')
printgraphs('./northwest/northwestwipdata.sqlite', './northwest/graphs/')

# print graphs for South West
print('Printing South West graphs')
printgraphs('./southwest/southwestwipdata.sqlite', './southwest/graphs/')
