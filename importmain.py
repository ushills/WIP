from modules.importWIP import import_wip_data

# import_wip_data takes information in the form
# import_wip_data(_database_name, importdirectory)

# import Midlands data
print("Importing Midlands data")
import_wip_data(
    "./midlands/midlandswipdata.sqlite", "H:/Current Month Wips/2. Midlands"
)

# import London data
print("Importing London data")
import_wip_data(
    "./london/londonwipdata.sqlite", "H:/Current Month Wips/1. London & South East"
)

# import North East data
print("Importing North East data")
import_wip_data(
    "./northeast/northeastwipdata.sqlite", "H:/Current Month Wips/3. North East"
)

# import North West data
print("Importing North West data")
import_wip_data(
    "./northwest/northwestwipdata.sqlite", "H:/Current Month Wips/3b. North West"
)

# import South West data
print("Importing South West data")
import_wip_data(
    "./southwest/southwestwipdata.sqlite", "H:/Current Month Wips/4. South West & Wales"
)
