from modules.multigraph import print_graphs

# print_graphs taken arguements in the form
# print_graphs(_database_name, output_directory)

# print graphs for Midlands
print("Printing Midlands graphs")
print_graphs("./midlands/midlandswipdata.sqlite", "./midlands/graphs/")

# print graphs for London
print("Printing London graphs")
print_graphs("./london/londonwipdata.sqlite", "./london/graphs/")

# print graphs for North East
print("Printing North East graphs")
print_graphs("./northeast/northeastwipdata.sqlite", "./northeast/graphs/")

# print graphs for North West
print("Printing North West graphs")
print_graphs("./northwest/northwestwipdata.sqlite", "./northwest/graphs/")

# print graphs for South West
print("Printing South West graphs")
print_graphs("./southwest/southwestwipdata.sqlite", "./southwest/graphs/")
