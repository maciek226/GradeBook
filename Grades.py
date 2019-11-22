# Version 0.5
# Maciej Bartosz Lacki

######## TO-DO:
# + **Multispreadsheet** - search through more then one sheet at a time to find the your section
# + add a data logger
# + Live Search function
# + Preivew changes,
# + Simplify the code
# + Move to object orineted programming 


######## Change-Log
# V0.1 - basic functionality
# V0.3 - Add multiple grades at the same time 
#      - added file paths
#      - UX fixes
#      - See if the grade was already enterd
# V0.5 - Added Config File
#      - Added ability to save the data to the original file (use at your own risk)
#      - Input robustness improvments 

# Check it the input is an integer, or can be converted to an integer 

import os
import csv

# Dev Settings
overWrite = 1
enterGrades = 1

# Config File
workPath = os.getcwd()
configFile = os.path.join(workPath,'GradesConf.txt')


file= str('') #this will be a lists once we make multispreadsheet work 
outFile= str('') #this will be a lists once we make multispreadsheet work 
#check if config file exists
#### Error Handiling - check if the files exist
if os.path.isfile(configFile):
    # If config file exists read the data 
    print('Loading user settings...\n')
    with open(configFile, newline = '') as book:
        bookType = csv.Sniffer().sniff(book.read(1024))
        # Previous command moves the cursor forward, we have to move it back...
        book.seek(0)
        reader = csv.reader(book,bookType)
        for row in reader:
            if row[0] == 'in':
                if os.path.isfile(row[1]):
                    file = row[1]
                    print('Loading input sheet:\t' + str(row[1]) + '\n')
                else:
                    print('Specified input file does not exist: ' + str(row[1]) + '\n')
            elif row[0] == 'out':
                if os.path.isfile(row[1]):
                    outFile = row[1]
                    print('Loading output sheet:\t' + str(row[1]) + '\n')
                else:
                    print('Specified input file does not exist: ' + str(row[1]) + '\n')
            elif row[0] == 'overWrite':
                overWrite = row[1]
if (len(file) > 0 and len(outFile) > 0) or (len(file) > 0 and overWrite == 1):
    status = 1
else:
    status = 0

if status == 0:
    # If the file does not exist, ask for the paths 
    print('no user settings found, starting from scratch')
    while 1:
        file = input('Enter path to the sheet\n')
        if os.path.isfile(file) == 1:
            status = 1
            break

    while 1:
        outFile = input('Enter the output file \n')
        if os.path.isfile(outFile) == 1:
            status = 1
            break
#TODO - make a config file based on the user settings 
#    with open(configFile,'w+'newline = '') as book:
#        writer = csv.writer(book,delimiter='\t')
        
if status > 0:
    print('Welcome to the Grade Editor\n')
    print('To leave the applicaiton and save the results \ntype EXT into the command line, when asked for the user\n')
    print('If you accidentally selected wrong individual, \nyou can cancel the operation by typing CNC\n')
##################### Spreadsheet Setup




#Spreadsheet columns with the desired data
sNumCol = 2 
fNameCol = 1
lNameCol = 0
gradeCol = list()

#Path to the spreadsheet

sNum = list()
Name = list()
grade = list()

originalData = list()


with open(file,encoding='utf-16', newline = '') as book:
    # Get the formatting of the spreadsheet - later used to 
    bookType = csv.Sniffer().sniff(book.read(1024))
    # Previous command moves the cursor forward, we have to move it back...
    book.seek(0)
    reader = csv.reader(book,bookType)
    
    flag = 0;
    # Read each row, and record the data in the predefined lists
    # One list stores student numbers and the other stores names
    # Both contain the line index
    for row in reader:
        # First row contains the headers, choose one to fill in the grades
        if flag == 0:
            cc = 0
            for ent in row:
                    print(str(cc) + "\t" + str(row[cc]))
                    cc = cc + 1
            addmore = 'y'
            while addmore == 'y':
                gcIn = input('Select column of grades to enter\n\t')
                try:
                    addcol = int(gcIn)
                except:
                    addcol = -1
                if addcol != -1 and addcol< len(row):
                    gradeCol.append(addcol)
                    flag = 1
                    while 1:
                        addmore = input('Would you like to edit more grade columns? y/n\n\t')
                        if str(addmore) != 'y' and str(addmore) != 'n':
                            print('Invalid Input')
                        else:
                            break
                else:
                    print('Invalid Selection')

        # Save the data we want in the sheet, create a list of grades(copy the data)
        originalData.append(row)

        sNum.append([row[sNumCol],reader.line_num])
        Name.append([row[fNameCol],row[lNameCol],reader.line_num])
        glist = list()
        cc = 0
        for col in gradeCol:
            glist.append(row[gradeCol[cc]])
            cc = cc + 1
        grade.append(glist)


# Searching and enetering grades - works until canceled
while enterGrades:
    ID = -1 # carries the line number of the student
    
    userIn = input('Enter Name or Student Number\n\t')
    
    # Exit the loop (right now it goes directly to saving
    if userIn == 'EXT': 
        break

    # Check the input type provided by the user 
    mode = 0
    try:
        int(userIn)
    except ValueError:
        mode = 1
    

    if mode == 0:
        # Search by student number
        
        matching = list()

        # Scan all student numbers and find ones that match the search
        for ent in sNum:
                if str(userIn) in ent[0]: #ent = [student number, index]
                    matching.append(ent)

        # Get the length of the matching results and...
        if len(matching) == 1:
            # If there is only one match, save the index to ID and display the student info
            ID = int(matching[0][1])-1
            print("\t" + str(Name[ID][0]) + " " + str(Name[ID][1]) +" "+ str(sNum[ID][0]))
        
        elif len(matching) > 1:
            # If there are multiple possible fits, display them all... 
            print("Found these possibilities: ")
            cc = 0
            for ent in matching:
                print("\t" + str(cc) + "\t" + ent[0] + " " + str(Name[ent[1]-1][0])+ " " + str(Name[ent[1]-1][1])+"\n" )
                cc = cc+1

            # Let the user choose one of the options 
            while 1:
                refine = input("Select one of the option\n\t")

                # If non of the options work, cancel
                if refine == "CNC":
                    break

                try:
                    refineN = int(refine)
                except:
                    print('please enter a number')
                    refineN = -1

                # If the choice is valid, leave 
                if refineN != -1 and refineN < cc:
                    ID = matching[refineN][1]-1
                    print("Chosen: \n \t" + str(Name[ID][0]) + " " + str(Name[ID][1]) +" "+ str(sNum[ID][0]))
                    break
                else:
                    print('choose a valid number')

        # If no hits, say it, and try again
        elif len(matching) == 0:
            print("Found nothin...")

     
    elif mode == 1:
        # Search by name and last name 
        matching = list()

        # Scan all first and last names and find ones that match the search
        # For convineance the case of both the input and the values is ignored 
        for ent in Name:
                if str.lower(userIn) in str.lower(ent[0]) or (userIn) in str.lower(ent[1]):
                    matching.append(ent)

        # If there is only one match, save the index to ID and display the student info
        if len(matching) == 1: 
            ID = matching[0][2]-1
            print("Found: \n \t" + str(Name[ID][0]) + " " + str(Name[ID][1]) +" "+ str(sNum[ID][0]))
        
        elif len(matching) > 1:
            # If there are multiple possible fits, display them all... 
            print("Found these possibilities: ")
            cc = 0
            for ent in matching:
                print("\t" + str(cc) + "\t" + str(sNum[ent[2]-1][0]) + " " + str(Name[ent[2]-1][0])+ " " + str(Name[ent[2]-1][1])+"\n" )
                cc = cc+1
                
            # Let the user choose one of the options 
            while 1:
                refine = input("\t\tSelect one of the option\n\t")

                # If non of the options work, cancel
                if refine == 'CNC':
                    break
                
                try:
                    refineN = int(refine)
                except:
                    print('please enter a number')
                    refineN = -1

                # If the choice is valid, leave 
                if refineN != -1 and refineN < cc:
                    ID = matching[refineN][2]-1
                    print("Chosen: \n \t" + str(Name[ID][0]) + " " + str(Name[ID][1]) +" "+ str(sNum[ID][0]))
                    break
                else:
                    print('choose a valid number')

        # If no hits, say it, and try again
        elif len(matching) == 0:
            print("Found nothin...")

    # If Any of the loops succeeded at finding the student, enter the grade 
    if ID != -1:
        cc = 0
        for gr in grade[ID]:
            print("\t" + "Current Grade: " + gr)
            while 1:
                gradeIn = input('\tEnter Grade #' + str(cc+1) + "\n\t\t" )

                # If the choise was accidental, break 
                if gradeIn == 'CNC':
                    break

                # Validate the input
                try:
                    gradeInN = float(gradeIn)
                except:
                    gradeInN = -1

                # If the input is valid, save the input in the correct location, and leave
                if gradeInN != -1:
                    grade[ID][cc] = gradeIn
                    print('\n')
                    #print(str(grade[ID]))
                    cc = cc+1
                    break
                else:
                    print('Invalid Entery')


#Write the grades to a new file or overwrites the curret file 

if overWrite == 1:
    with open(file,'w',encoding='utf-16', newline = '') as book:
        writer = csv.writer(book,bookType)
        flag = 0
        dd = 0
        for row in originalData:
            # Do not change the first row
            if flag != 1:
                newRow = row
                cc = 0
                for col in gradeCol:
                    newRow[col] = grade[dd][cc]
                    cc = cc+1
                print(str(newRow))
                writer.writerow( newRow )
            else:
                writer.writerow( row )
                flag = 1
            dd = dd + 1
else:
    with open(outFile,'w+',encoding='utf-16', newline = '') as nBook:
        # Open the old file 
        with open(file,encoding='utf-16', newline = '') as book:
            
            reader = csv.reader(book,bookType)
            writer = csv.writer(nBook,bookType)
            book.seek(0)
            flag = 0

            # Copy the entier file, and substitute the grade 
            for row in reader:
                # Do not change the first row 
                if flag != 1:
                    newRow = row
                    cc = 0
                    for col in gradeCol:
                        newRow[col] = grade[reader.line_num-1][cc]
                        cc = cc+1
                    writer.writerow( newRow )
                else:
                    writer.writerow( row )
                    flag = 1
