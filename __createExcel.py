import os
import pandas as pd
from datetime import datetime
fileSizes = {"KB": 1024, "MB":1048576, "GB" : 1073741824, "TB" : 1099511627776}

#input params
path = os.getcwd() #add any path (by default it is the current directory)
fileName = os.path.basename(path)+".xlsx" if (os.path.basename(path) != "") else "files.xlsx" #set the name of the excel file by default the parent folder name
fileSize = 1 #files listed will be greater than this size in MB
sizeUnit = "MB"
header = ['Index',os.path.basename(path)+' Name', 'size({})'.format(sizeUnit), 'lastModified'] # column names in the excel

fileSize *=fileSizes.get(sizeUnit) #convert size to bytes

#recursive listing
def getFiles(dirpath):
    subfiles = os.listdir(dirpath)
    files = []
    for i in subfiles:

        if not (i.endswith('$RECYCLE.BIN') or i.endswith('System Volume Information')): #avoid permission issues
            if os.path.isdir(dirpath+'/'+i):
                files.extend(getFiles(dirpath+'/'+i))
            elif os.path.getsize(dirpath+'/'+i) > fileSize:
                fileName = i
                hyperLink = '=HYPERLINK("'+ dirpath+'\\'+i+'", "'+i+'")'
                Size = round(os.path.getsize(dirpath+'/'+i)*(1/fileSizes.get(sizeUnit)),2)
                Date = datetime.utcfromtimestamp(os.path.getmtime(dirpath+'/'+i)).strftime('%d-%m-%Y %h:%M:%S')
                files.append([fileName, hyperLink, Size, Date])
    return files
            

files = getFiles(path) #get all file data

files = sorted(files, key = lambda x:x[2], reverse=True) #the index decides the sort (by default size descending)

fileLinks = [x[:] for x in files] #deepcopy the files list

for i in range(len(fileLinks)):
    fileLinks[i][0] = i #index

df = pd.DataFrame(fileLinks, columns=header) #create a dataFrame

for i in range(len(fileLinks)):
    fileLinks[i][1] = files[i][0] #get file names for dynamic column width set

dfNames = pd.DataFrame(fileLinks, columns = header) #create a dataFrame for dynamic column width set

writer = pd.ExcelWriter(fileName, engine='xlsxwriter') #for styling and column width set

df.to_excel(writer, sheet_name="Sheet1",index=False)

#add format to center the columns
format = writer.book.add_format()
format.set_align('center')

#dynamically calculate the column width and set it
for col in dfNames:
    maxLen = 0
    for mov in dfNames[col]:
        maxLen = max(len(str(mov)), maxLen, len(col))
    loc = dfNames.columns.get_loc(col)
    writer.sheets["Sheet1"].set_column(loc,loc,maxLen, format)

writer.save() #save the excel
