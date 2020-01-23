import shutil

print("Copying files....")
shutil.copytree("./london/graphs/", "H:/Performance Graphs/london", dirs_exist_ok=True )
shutil.copytree("./midlands/graphs/", "H:/Performance Graphs/midlands", dirs_exist_ok=True )
shutil.copytree("./northeast/graphs/", "H:/Performance Graphs/northeast", dirs_exist_ok=True )
shutil.copytree("./northwest/graphs/", "H:/Performance Graphs/northwest", dirs_exist_ok=True )
shutil.copytree("./southwest/graphs/", "H:/Performance Graphs/southwest", dirs_exist_ok=True )
print("Copying files complete!")