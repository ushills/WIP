from multigraph import printgraphs

# printgraphs taken arguements in the form
# printgraphs(DATABASE_NAME, outputdirectory)

# print graphs for Midlands
printgraphs (r'.\midlands\midlandswipdata.sqlite', r'.\midlands\graphs\\')

# print graphs for London
printgraphs (r'.\london\londonwipdata.sqlite', r'.\london\graphs\\')
