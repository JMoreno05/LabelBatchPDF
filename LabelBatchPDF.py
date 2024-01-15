'''
Created on May 27, 2022

@author: lady_
'''
#    future, varg program type
#    program location file different location for each type

import os, csv
from PyPDF2 import PdfFileReader, PdfFileWriter, PdfFileMerger

barcodePath = ""
expPath = ""
bagLabelPath = ""
errors = []
programPath = ""

def readAppLocation():
    global programPath
    print("Checking App Location ...")
    docPath = os.path.expanduser('~\Documents')
    # print("docpath",docPath)
    try:
        with open(docPath +"\LabelBatchPDF ProgramLocation.csv", 'r') as file:
            reader = csv.reader(file)
            for row in reader:
                programPath = row[0]
    except FileNotFoundError as e:
        print("The file 'LabelBatchPDF ProgramLocation.csv' was not found in your Documents folder. Please copy and paste from the 'LabelBatchPDF' folder")
        os.system("PAUSE")
        return False
    #print(programPath)
    return True

def checkDependency():
    global programPath
    print("Checking on Dependencies ...")
    isOk = True
    msg = ""
    
    folders = os.listdir(programPath)    
    if "new.pdf"  in folders:
            os.remove(programPath + "\\new.pdf")
            createNewPDF()
    else:
        createNewPDF()
            #print("New new.pdf created.")
    if "errors.csv"  in folders:
            os.remove(programPath + "\\errors.csv")

    if "LabelBatchPDF Settings.csv" not in folders:
        msg = "Settings File is missing."
        isOk = False
        #prompt to create new from scratch
    elif "items.csv"  not in folders:
        isOk = False
        msg = "CSV file does not exist (items.csv)"
    
    return isOk, msg

def createNewPDF(): #create unique file name and
    global programPath
    # print("create pdf")
    pdf = PdfFileWriter()
    with  open(programPath + "\\new.pdf", 'wb') as file:
        pdf.addBlankPage(162,108) #A4
        pdf.write(file)

def readAppSettings(): # set global variables 
    # for Barcode folder, EXP folder and Bag Folder and filename
    global barcodePath, expPath, bagLabelPath, programPath
    print("Reading App Settings ...")
    locSettings = programPath + "\LabelBatchPDF Settings.csv"

    try:
        with open(locSettings, "r", newline="") as file:
            reader = csv.reader(file)
            for row in reader:
                # print(row[0],row[1])
                if row[0] == "BarcodeFolder" :
                    barcodePath = row[1]
                    
                    # print("barcodePath", row[1])
                elif row[0] == "ExpirationLabelFolder":
                   
                    expPath = row[1]
                    # print(expPath)
                elif row[0] == "BagWarning.pdf Path":
              
                    bagLabelPath = row[1]
                    # print(bagLabelPath)
                else:
                    length = len(row[0])
                    if len(row[0]) > 13:
                        x = row[0][3:]
                        y = row[0][length - 10:]
                        if y == "codeFolder":
                            # print("match")
                            barcodePath = row[1]
                        # print(x, "edited row name", y)
                    # print("other", row[0], row[1])
        # os.system("PAUSE") #prevents cmd line from closing at end of program  
            
    except FileNotFoundError as e:
        print("'LabelBatchPDF Settings' not found", e)
        os.system("PAUSE") #prevents cmd line from closing at end of program
        return False
    except Exception as e:
        print(e)
        os.system("PAUSE") #prevents cmd line from closing at end of program
        return False
    return True

def readItemsFromCSV(): # returns list
    global programPath
    print("Reading Items ...")
    items = []
    with open(programPath + "\\items.csv", 'r', newline="") as file:
        reader = csv.reader(file)
        for row in reader:
            items.append(row) 
    # print(items)
    # os.system("PAUSE")
    
    return items

def checkPathNotFound():
    print("Checking paths from Settings ...")
    global barcodePath, expPath, bagLabelPath
    folders = []
    try:
        path = barcodePath
        folders = os.listdir(barcodePath)
        #print(folders, "bc")
        path = expPath
        folders = os.listdir(expPath)
        #print(folders, 'exp')
        path = bagLabelPath
        folders = os.listdir(bagLabelPath)
        #print(folders,"bag")
    except FileNotFoundError as e:
        print("Invalid Path", path,e )
        os.system("PAUSE") #prevents cmd line from closing at end of program
        return False
    # os.system("PAUSE")
    return True

def checkFileNotFound(task):
    print("Checking If item is missing labels...")
    msg =[task[0]]
    global barcodePath, expPath, bagLabelPath
    error = False
    folders = os.listdir(barcodePath)   
    barcode = task[0] + ".pdf"
    if barcode not in folders:
        # print("no barcode", barcode)
        msg.append("No Barcode")
        error = True

    folders = os.listdir(expPath)   
    exp = task[1] + ".pdf"
    if exp not in folders:
        msg.append("No Exp")
        error = True
        # print("no exp", exp)
    if task[3] == 'Y' and "BagWarning.pdf" not in folders:
        msg.append("No BagWarning.pdf")
        # print("no bag found","BagWarning.pdf")
        error = True
    
    if error:
        global errors
        errors.append(msg)
        return True
    else:
        return False

def exportErrorCSV():
    global errors
    print("Exporting errors...")
    # check if file exists and return list 
    #if not null, append to error list
    if errors == []:
        errors.append(["No Errors Found"])
    errors.insert(0,["SKU","Error"])
    with open(programPath + "\\errors.csv", 'w', newline="") as file:
        writer = csv.writer(file)
        writer.writerows(errors)
    # print(errors)
    # os.system("PAUSE")

def addLabelsToBatchPDF(dest, file, qty):
    # print("add label")
    with open(dest, 'rb') as destFile:
        destFileReader = PdfFileReader(destFile)

        with open(file, 'rb') as lblFile:
            pgReader = PdfFileReader(lblFile)
    
            merger =  PdfFileMerger(strict=True) #should user be warned of all problems
        
            merger.append(destFileReader)
            for  i in range(qty):
                merger.append(pgReader)
            
            merger.write(dest)
    # os.system("PAUSE")
    
def endOfProgram():
    # print("eof")
    os.startfile(programPath + "\\new.pdf")
    exportErrorCSV()
    os.startfile(programPath + "\\errors.csv")
    for error in errors:
        print(error)
    os.system("PAUSE") #prevents cmd line from closing at end of program    
            
def main():
    global barcodePath, expPath, bagLabelPath, errors, programPath
    if readAppLocation():
        #
        # #print("Current Dir", os.getcwd())
        #
        # os.system("PAUSE") #prevents cmd line from closing at end of program
        # #to change current dir
        # #os.getchdir(path)
        sucess, msg = checkDependency()
        if not sucess:
            print(msg)
            os.system("PAUSE") #prevents cmd line from closing at end of program
        else:
            readAppSettings()
            if checkPathNotFound():
                tasks = readItemsFromCSV()
                #print(barcodePath, expPath, bagLabelPath)
        
                for task in tasks:
                    if checkFileNotFound(task) == False:
                        print("Adding Labels to file ...")
                    # Barcode
                    #try:    #return error msg, append SKU and append to errors
                        addLabelsToBatchPDF(programPath + "\\new.pdf", barcodePath + "/" + task[0] + ".pdf", int(task[2]))
                        # Exp
                        addLabelsToBatchPDF(programPath + "\\new.pdf", expPath + "/" + task[1] + ".pdf", int(task[2]))
                        # bag label
                        if task[3] == 'Y': ## print bag label
                            addLabelsToBatchPDF(programPath + "\\new.pdf", bagLabelPath + "\\BagWarning.pdf", int(task[2]))
    
                    
                endOfProgram()
    os.system("PAUSE") #prevents cmd line from closing at end of program   
    
if __name__ == '__main__':
    main()
