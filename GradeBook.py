# GradeBook
# Maciej Lacki 2019

# limitations: error detection
#   It is possible for user to not select any grade column


### Use at your
import os
import csv
import datetime


class GradeBook:
    # Current root directory
    workDir = str
    # Assumed config file name/location
    configFile = str

    # Input Files
    inFile = list()
    outFile = list()

    fileFormat = csv.Sniffer()

    # Name of the config file
    configName = 'GradesConf.txt'

    # Student Data
    sNum = 2
    fName = 1
    lName = 0
    numSheets = 0
    gradeCols = list()

    sheets = list()

    # Commands
    exitCom = 'EXT'
    cancelCom = 'CNC'

    # Logger
    logPath = str
    loggerSetup = 0

    #### Settings
    # Overwrite grades
    overWrite = 1


###############################################################################
#                                   Setup                                     #
###############################################################################
def setup():
    gradebook.workDir = os.getcwd()
    gradebook.configFile = os.path.join(gradebook.workDir, gradebook.configName)
    if os.path.isfile(gradebook.configFile):
        # Config Exists, read it
        print('Config file found\n\t Loading Setting...\n')
        settingsImport()
    else:
        # Config does not exist...
        print('Config file not found')
        settingGet()
        settingsExport()

    if len(gradebook.inFile) > 0:
        readSheets()
    else:
        print('no sheets added')
        settingGet()
        settingsExport()
        readSheets()


def settingsImport():
    with open(gradebook.configFile, newline='') as book:
        bookType = csv.Sniffer().sniff(book.read(1024))

        book.seek(0)
        reader = csv.reader(book, bookType)

        print('Opening file(s):')
        for row in reader:  # Check each line
            if row[0] == 'in':  # If the first item in the line is in
                if os.path.isfile(row[1]):  # If the file exists
                    # Add the input file to the list
                    gradebook.inFile.append(row[1])
                    print('\t%s' % (str(row[1])))
                else:
                    # The file does not exist
                    print('File ', row[1], ' does not exist')
            elif row[0] == 'out':  # If the first item in the line is in
                if os.path.isfile(row[1]):  # If the file exists
                    # Add the input file to the list
                    gradebook.outFile.append(row[1])
                    print('Opening file \n \t ', row[1])
                else:
                    # The file does not exist
                    print('File ', row[1], ' does not exist')
            elif row[0] == 'overWrite':
                GradeBook.overWrite = row[1]
        print('\n')


def settingsExport():
    print('Exporting...')
    with open(gradebook.configFile, 'w+', newline='') as book:
        writer = csv.writer(book, delimiter='\t')
        for file in gradebook.inFile:
            writer.writerow(['in', file])


def settingGet():
    # If the config file does not exist, get the requiered paths
    while 1:
        while 1:
            file = str(input('Enter path to the sheet\n'))
            if os.path.isfile(file):
                gradebook.inFile.append(file)
                break
            elif 'cancel' in file:
                break

        if len(gradebook.inFile) > 0 and str(input('Add more sheets? (y/n) \n \t')) == 'n':
            gradebook.overWrite = 1
            break


# Collects all the data from the sheets found in the config file and puts them in a list of lists


def readSheets():
    # Counter of the sheets
    cc = 0
    # Open each sheet
    for sheet in gradebook.inFile:
        with open(gradebook.inFile[cc], 'r', encoding='utf-16', newline='') as book:
            # Check the type of the sheet, and bring the cursor back to 0
            book.seek(2048)
            bookType = csv.Sniffer().sniff(book.readline())
            gradebook.fileFormat = bookType
            book.seek(0)
            reader = csv.reader(book, bookType)

            flag = 0;
            dd = 0

            # Add a new list to for each sheet
            gradebook.sheets.append(list())
            for row in reader:
                # Add to the newly created list a new list to contain the user data
                gradebook.sheets[cc].append(list())

                # Add the data from each row
                for item in row:
                    gradebook.sheets[cc][dd].append(item)
                dd = dd + 1
        cc = cc + 1

    gradebook.numSheets = cc
    gradeColSelect()


###############################################################################
#                          User Interatction                                  #
###############################################################################

# Select the spreadsheet columns which will be overwriten
def gradeColSelect():
    cc = 0
    print('The Columns in the sheet:\n')
    # print the first column of the sheet
    for sheet in gradebook.sheets:

        gradebook.gradeCols.append(list())
        dd = 0
        for entery in sheet[0]:
            print("\t%d \t%s" % (dd, entery))
            dd = dd + 1

        while 1:
            selection = selectCol(dd)
            if selection != -1:
                gradebook.gradeCols[cc].append(selection)

            more = str(input('\nWould you like to add more grade columns? (y/n)\n\t'))

            if more == 'n':
                break
        cc = cc + 1


# Make a selection of the
def selectCol(number):
    while 1:
        selection = str(input('\nSelect the grading column \n\t'))
        try:
            if int(selection) >= 0 and int(selection) <= number:

                return (int(selection))
            else:
                print('\nSelection out of range')
        except:
            if selection == str(gradebook.cancelCom):
                return (-1)
            else:
                print('\nInput is invalid')


# Main interaction loop
def enterStud():
    # On until broken
    while 1:
        # Get the user input
        userIn = str(input('Enter Name or Student Number\n\t'))

        # If exit command - leave, else pass the input into search
        if str(userIn) == gradebook.exitCom:
            break
        elif str(userIn) == 'Settings':
            print('\tFeature coming in future version... maybe')
        else:
            student = search(userIn)

            if student != -1:
                enterGrades(student)
    if gradebook.overWrite == 0:
        writeSheet()
    elif gradebook.overWrite == 1:
        overwriteSheet()


###############################################################################
#                          Searching                                          #
###############################################################################
# Searches the list of students
def search(phrase):
    # Deafult serach mode - by number
    sMode = 0

    # If the entery is not an intiger change to name search mode
    try:
        int(phrase)
    except ValueError:
        sMode = 1

    # Search using one of the modes
    if sMode == 0:
        student = searchNumber(phrase)
    else:
        student = searchName(phrase)

    if student != -1:
        return (student)
    else:
        return (-1)


def searchName(name):
    matches = list()

    for sheet in range(0, gradebook.numSheets):
        index = 0
        flag = 0
        for entery in gradebook.sheets[sheet]:
            # Skip the first row
            if flag == 0:
                flag = 1
            else:
                if str.lower(name) in str.lower(entery[gradebook.fName]):
                    matches.append(
                        [entery[gradebook.fName], entery[gradebook.lName], entery[gradebook.sNum], sheet, index])

                if str.lower(name) in str.lower(entery[gradebook.lName]):
                    matches.append(
                        [entery[gradebook.fName], entery[gradebook.lName], entery[gradebook.sNum], sheet, index])
            index = index + 1

    if len(matches) == 1:
        return (matches[0])
    elif len(matches) > 1:
        selection = refineSearch(matches)
        if selection == -1:
            return (-1)
        else:
            return (matches[int(selection)])
    elif len(matches) == 0:
        print('\n\tNo Matches Found')
        return (-1)


def searchNumber(number):
    matches = list()
    for sheet in range(0, gradebook.numSheets):
        index = 0
        flag = 0
        for entery in gradebook.sheets[sheet]:
            if flag == 0:
                flag = 1
            else:
                if str(number) in entery[gradebook.sNum]:
                    matches.append(
                        [entery[gradebook.fName], entery[gradebook.lName], entery[gradebook.sNum], sheet, index])
            index = index + 1

    if len(matches) == 1:
        return (matches[0])
    elif len(matches) > 1:
        selection = refineSearch(matches)
        print(selection)
        if selection == -1:
            return (-1)
        else:
            return (matches[int(selection)])
    elif len(matches) == 0:
        print('\n\tNo Matches Found')
        return (-1)


def refineSearch(matches):
    print("\nFound following matches:")
    cc = 0
    for student in matches:
        print("\t%d - %s" % (cc, str(student[gradebook.sNum])), student[gradebook.lName], student[gradebook.fName],
              '\n')
        cc = cc + 1
    while 1:

        selection = str(input("Select the student:\n\t"))

        if str(selection) == str(gradebook.cancelCom):
            return (-1)

        # attempt to use the input to select the user, if invalid try again
        try:
            if int(selection) >= 0 and int(selection) < len(matches):
                return (selection)
            else:
                print("Please enter a valid number")
        except:
            print("Please enter a valid number")


def enterGrades(student):
    print('Selected')
    print("\t%s %s %s" % (student[gradebook.lName], student[gradebook.fName], student[gradebook.sNum]))
    cc = 1
    print('\tCurrent Grades:')

    for entery in gradebook.gradeCols[student[-2]]:
        print("\t\tG#%d:" % (cc), gradebook.sheets[student[-2]][student[-1]][entery])
        cc = cc + 1

    cc = 1
    for entery in gradebook.gradeCols[student[-2]]:
        while 1:
            grade = str(input('\n\tEnter Grade #%d\n\t\t' % (cc)))
            try:
                gradebook.sheets[student[-2]][student[-1]][entery] = float(grade)
                break
            except:
                print('Invalid Input')
        cc = cc + 1
    print("\n")


###############################################################################
#                          Write Sheets                                       #
###############################################################################
def writeSheet():
    cc = 0
    for sheet in gradebook.sheets:
        name = 'file' + str(cc + 1) + '.xls'

        with open(os.path.join(gradebook.workDir, name), 'w+', encoding='utf-16', newline='') as book:
            print(gradebook.fileFormat)
            writer = csv.writer(book, gradebook.fileFormat)
            book.seek(0)

            for row in gradebook.sheets[cc]:
                # print(str(row))
                writer.writerow(row)
        cc = cc + 1
        # print('\n\n\n\n')


def overwriteSheet():
    cc = 0
    for sheet in gradebook.inFile:
        with open(gradebook.inFile[cc], 'w', encoding='utf-16', newline='') as book:
            writer = csv.writer(book, gradebook.fileFormat)
            book.seek(0)

            for row in gradebook.sheets[cc]:
                writer.writerow(row)
        cc = cc + 1


def logger(Log):
    if GradBook.loggerSetup == 1:
        with open(GradeBook.logPath, 'a') as logger:
            time = datetime.now()
            logger.write(time.strftime("%y-%m-%d %H:%M:%S - "), Log)


def loggerSetup():
    GradeBook.logPath = os.path.join(GradeBook.logPath, "Log.txt")

    if os.path.exists(GradeBook.logPath):
        GradBook.loggerSetup = 1

        return True
    else:
        flag = 0


def inputCheck(string):
    # Check the type of the
    flag = 0


# Not integrated into code yet
def isInt(string):
    try:
        int(string)
        return (int(string))
    except ValueError:
        return False


# The class holding all the information
gradebook = GradeBook()

setup()
print('\n\n\n\n\n\n\n\n\n')
enterStud()
