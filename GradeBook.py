# GradeBook
# Maciej Lacki
# 2019

# Use at your own risk

# Version 0.6
#   Fixed inability to cancle selection in grade input
#   Improved the config file creation wizard
#   Config files now have save overwrite preferences
#   Moved all methods to the GradeBook Class
#   Fixed the naming scheme - now using python naming convention
#   Commented the code
#   Added the ability to add a new column to a sheet with command ADD

# Version 0.55
#   Fixed -1 after using CNC
#   Added the path to the selected sheet in the grade selection path

# limitations: error detection
#   It is possible for user to not select any grade column even multiple times

# TODO:
#   Loger - gather all the data in case something crashes
#   Restore - from the logs
#   Menu - Make it easier to select the columns, adjust their order
#   Add select save location
#   Ability to assign the same grades to more then one person
#   Add Doc Strings


import os
import csv
import datetime


class GradeBook:
    """Opens Spreadsheets from downloaded from blackboard and aids in
    inputing the grades

    The class contains all the data and methods used to interacts with
    the spreadsheets
    """
    version = "0.6"
    dev = 0

    # Current root directory
    working_dir = str

    # Assumed config file name/location
    config_dir = str

    # Input Files
    input_files = list()
    output_files = list()

    file_format = csv.Sniffer()

    # Name of the config file
    config_name = 'GradesConf.txt'

    # Student Data
    s_num_col = 2
    first_name_col = 1
    last_name_col = 0
    number_sheets = 0
    grading_columns = list()

    sheets = list()

    # Commands
    exit_command = 'EXT'
    cancel_command = 'CNC'
    add_sheet_command = 'ADD'

    # Logger
    log_path = str
    logger_setup = 0

    # Settings
    # Overwrite grades
    overwrite = 1

    # Options used in main menu
    initialized = 0
    menu_options = ["Start",
                    "Settings",
                    "Spreadsheet Setup",
                    "Exit and save",
                    "Cancel"]
    ready = 0

    def __init__(self):
        print('Welcome to GradeBook V%s \n' % self.version)

###############################################################################
#                                   Setup                                     #
###############################################################################
    # Determines current working path, attempts to read the config file
    # Creates a config file if one is not present, or not working
    def setup(self):
        self.working_dir = os.getcwd()
        self.config_dir = os.path.join(self.working_dir, self.config_name)

        if os.path.isfile(self.config_dir):
            # Config Exists, read it
            print('Config file found\n\t Loading Setting...\n')
            self.import_settings()
        else:
            # Config does not exist...
            print('Config file not found')
            self.get_settings()
            self.export_settings()

        if len(self.input_files) > 0:
            # There is more then 0 input files defined - read them
            self.read_sheets()
        else:
            # No sheets were loaded, assume the config is broken, make new one
            print('no sheets added')
            self.get_settings()
            self.export_settings()
            self.read_sheets()

    # Read the settings file
    def import_settings(self):
        with open(self.config_dir, newline='') as book:
            bookType = csv.Sniffer().sniff(book.read(1024))

            book.seek(0)
            reader = csv.reader(book, bookType)

            print('Opening file(s):')
            for row in reader:  # Check each line
                if row[0] == 'in':  # If the first item in the line is in
                    if os.path.isfile(row[1]):  # If the file exists
                        # Add the input file to the list
                        self.input_files.append(row[1])
                        print('\t%s' % (str(row[1])))
                    else:
                        # The file does not exist
                        print('File ', row[1], ' does not exist')
                elif row[0] == 'out':  # If the first item in the line is in
                    if os.path.isfile(row[1]):  # If the file exists
                        # Add the input file to the list
                        self.output_files.append(row[1])
                        print('Opening file \n \t ', row[1])
                    else:
                        # The file does not exist
                        print('File ', row[1], ' does not exist')
                elif row[0] == 'overwrite':
                    self.overwrite = row[1]
                elif row[0] == 'dev':
                    self.dev = row[1]
            print('\n')

    # Save or overwrite the settings
    def export_settings(self):
        print('\n\t Saving settings...\n')
        with open(self.config_dir, 'w+', newline='') as book:
            writer = csv.writer(book, delimiter='\t')

            for file in self.input_files:
                writer.writerow(['in', file])

            writer.writerow(["overwrite", self.overwrite])
            writer.writerow(["dev", self.dev])

    # Get the neccesserry user preferences
    def get_settings(self):
        # If the config file does not exist, get the requiered paths
        while 1:
            while 1:
                file = str(input('Enter path to the sheet\n'))
                if os.path.isfile(file):
                    self.input_files.append(file)
                    break
                elif 'cancel' in file:
                    break

            if len(self.input_files) > 0 and str(input('Add more sheets? (y/n)\n\t')) == 'n':
                self.overwrite = 1
                break
        while 1:
            overwrite = str(input("Would you like to overwrite the sheets?(y/n)\n\t"))
            if overwrite == 'y':
                self.overwrite = 1
                break
            elif overwrite == 'n':
                self.overwrite = 0
                break
            else:
                print("Invalid Input")

    # Collects all the data from the sheets found
    # in the config file and puts them in a list of lists
    def read_sheets(self):
        # Counter of the sheets
        cc = 0

        # Open each sheet
        for sheet in self.input_files:

            with open(self.input_files[cc], 'r',
                      encoding='utf-16', newline='') as book:

                # Check the type of the sheet, and bring the cursor back to 0
                book.seek(2048)
                bookType = csv.Sniffer().sniff(book.readline())
                self.file_format = bookType
                book.seek(0)
                reader = csv.reader(book, bookType)

                dd = 0

                # Add a new list to for each sheet
                self.sheets.append(list())
                for row in reader:
                    # Add to the newly created list a
                    # new list to contain the user data
                    self.sheets[cc].append(list())

                    # Add the data from each row
                    for item in row:
                        self.sheets[cc][dd].append(item)
                    dd = dd + 1
            cc = cc + 1

        self.number_sheets = cc

        # Select the grading columns
        self.column_selection()


###############################################################################
#                          User Interatction                                  #
###############################################################################

    # Select the spreadsheet columns which will be overwriten
    def column_selection(self):

        cc = 0
        print('The Columns in the sheet to add a new column type %s:\n' % self.add_sheet_command)
        # print the first column of the sheet
        for sheet in self.sheets:

            self.grading_columns.append(list())
            dd = 0
            print(("\t%s\n") % str(self.input_files[cc]))

            for entery in sheet[0]:

                print("\t%d \t%s" % (dd, entery))
                dd = dd + 1

            while 1:

                selection = self.select_column(dd, cc)

                if selection != -1:
                    self.grading_columns[cc].append(selection)

                more = str(input('\nWould you like to add more grade columns? (y/n/ADD)\n\t'))

                if more == 'n':
                    break
            cc = cc + 1

    # Make a selection of the sheet
    def select_column(self, number, sheet):
        while 1:
            selection = str(input('Select the grading column \n\t'))
            try:
                if int(selection) >= 0 and int(selection) <= number:

                    return (int(selection))
                else:
                    print('\nSelection out of range')
            except(ValueError):

                if selection == self.add_sheet_command:
                    new_column_name = str(input('Enter the name of the new collumn: '))
                    self.add_column(sheet, new_column_name)
                    return(number)
                elif selection == str(self.cancel_command):
                    return (-1)
                else:
                    print('\nInput is invalid')

    # Add a column to the spreadsheet
    def add_column(self, sheet, column_name):

        cc = 0
        for student in self.sheets[sheet]:
            self.sheets[sheet][cc].append('')
            cc = cc + 1

        self.sheets[sheet][0][-1] = column_name

    # Main interaction loop
    def enter_student(self):

        print("\nYou can now search for the studentsusing their student",
              "number or name \nYou do not need to provide full student",
              " number or full name \nIf you select wrong person type CNC",
              " to cancel the selectrion \nTo save all the data and",
              " exit the application type EXT\n")

        # On until broken
        while 1:
            # Get the user input
            userIn = str(input('Enter Name or Student Number\n\t'))

            # If exit command - leave, else pass the input into search
            if str(userIn) == self.exit_command:
                break
            elif str(userIn) == 'Settings':
                print('\tFeature coming in future version... maybe')
            else:
                student = self.search(userIn)

                if student != -1:
                    self.enter_grades(student)
        # Once the search loop is initiated, save the sheets
        while 1:
            # Check if the sheets are writable
            status = self.check_sheet()
            if all(status):
                if self.overwrite == 0:
                    self.write_sheet()
                elif self.overwrite == 1:
                    self.overwrite_sheet()
                break
            else:
                # If sheets are not writable let the user know
                saveas = str(input("One of the sheets is open"
                                   " by another program, please close it,"
                                   " and press enter or type saveas\n"))
                if saveas == "saveas":
                    self.write_sheet()
                    break


###############################################################################
#                          Searching                                          #
###############################################################################

    # Searches the list of students
    def search(self, phrase):

        # Deafult serach mode - by number
        sMode = 0

        # If the entery is not an intiger change to name search mode
        try:
            int(phrase)
        except ValueError:
            sMode = 1

        # Search using one of the modes
        if sMode == 0:
            student = self.search_number(phrase)
        else:
            student = self.search_name(phrase)

        if student != -1:
            return (student)
        else:
            return (-1)

    def search_name(self, name):
        matches = list()

        # Check each sheet
        for sheet in range(0, self.number_sheets):
            index = 0
            flag = 0

            # In each sheet find matches
            for entery in self.sheets[sheet]:
                # Skip the first row
                if flag == 0:
                    flag = 1
                else:
                    # If a match is found, append the info, and the index
                    if str.lower(name) in str.lower(entery[self.first_name_col]):
                        matches.append([entery[self.first_name_col],
                                        entery[self.last_name_col],
                                        entery[self.s_num_col], sheet, index])

                    if str.lower(name) in str.lower(entery[self.last_name_col]):
                        matches.append([entery[self.first_name_col],
                                        entery[self.last_name_col],
                                        entery[self.s_num_col], sheet, index])
                index = index + 1

        # Check if any matches were found
        if len(matches) == 1:
            # If one match is found return
            return (matches[0])

        elif len(matches) > 1:
            # If multiple are found - refine the selection
            selection = self.refine_search(matches)
            if selection == -1:
                return (-1)
            else:
                return (matches[int(selection)])

        elif len(matches) == 0:
            # If non are found - let the user know
            print('\n\tNo Matches Found')
            return (-1)

    def search_number(self, number):
        matches = list()

        # In each sheet find matches
        for sheet in range(0, self.number_sheets):
            index = 0
            flag = 0
            for entery in self.sheets[sheet]:
                # Skip the first row
                if flag == 0:
                    flag = 1
                else:
                    # If a match is found, append the info, and the index
                    if str(number) in entery[self.s_num_col]:
                        matches.append([entery[self.first_name_col],
                                        entery[self.last_name_col],
                                        entery[self.s_num_col],
                                        sheet,
                                        index])
                index = index + 1

        # Check if any matches were found
        if len(matches) == 1:
            # If one match is found return
            return (matches[0])

        elif len(matches) > 1:
            # If multiple are found - refine the selection
            selection = self.refine_search(matches)
            # print(selection)
            if selection == -1:
                return (-1)
            else:
                return (matches[int(selection)])

        elif len(matches) == 0:
            # If non are found - let the user know
            print('\n\tNo Matches Found')
            return (-1)

    # Gets a list of potnetial matches, and gets the user to select one
    def refine_search(self, matches):

        # print all the matches
        print("\nFound following matches:")
        cc = 0
        for student in matches:
            print("\t%d - %s" % (cc, str(student[self.s_num_col])),
                  student[self.last_name_col],
                  student[self.first_name_col], '\n')

            cc = cc + 1

        # Get user to select one
        while 1:
            selection = str(input("Select the student:\n\t"))

            # Check if user wants to cancel
            if str(selection) == str(self.cancel_command):
                return (-1)

            # attempt to use the input to select the user, if invalid try again
            try:
                if int(selection) >= 0 and int(selection) < len(matches):
                    return (selection)
                else:
                    print("Please enter a valid number")
            except(ValueError):
                print("Please enter a valid number")

    # Print the selected student, and gets the grade input
    # for each selected grading column
    def enter_grades(self, student):
        # Print the student info
        print('Selected')
        print("\t%s %s %s"
              % (student[self.last_name_col],
                 student[self.first_name_col],
                 student[self.s_num_col]))

        cc = 1
        print('\tCurrent Grades:')

        for entery in self.grading_columns[student[-2]]:
            print("\t\tG#%d:" % (cc),
                  self.sheets[student[-2]][student[-1]][entery])

            cc = cc + 1

        # Get the grade for each selected column
        cc = 1
        for entery in self.grading_columns[student[-2]]:
            # Runs until a correct input is recived - float or CNC
            while 1:
                grade = str(input('\n\tEnter Grade #%d\n\t\t' % (cc)))
                try:
                    self.sheets[student[-2]][student[-1]][entery] = float(grade)
                    break
                except(ValueError):
                    if grade == str(self.cancel_command):
                        break
                    else:
                        print('Invalid Input')
            cc = cc + 1

###############################################################################
#                          Write Sheets                                       #
###############################################################################

    # Writes a new sheet in the working directory
    def write_sheet(self):
        cc = 0
        for sheet in self.sheets:
            name = 'file' + str(cc + 1) + '.xls'

            with open(os.path.join(self.working_dir, name), 'w+',
                      encoding='utf-16', newline='') as book:

                writer = csv.writer(book, self.file_format)
                book.seek(0)

                for row in self.sheets[cc]:
                    writer.writerow(row)
            cc = cc + 1

    # Opens the sheet and prints the updated data
    def overwrite_sheet(self):
        cc = 0
        for sheet in self.input_files:

            with open(self.input_files[cc], 'w',
                      encoding='utf-16', newline='') as book:

                writer = csv.writer(book, self.file_format)
                book.seek(0)

                for row in self.sheets[cc]:
                    writer.writerow(row)
            cc = cc + 1

    # Test if the spreadsheet is writable
    def check_sheet(self):
        status = list()
        cc = 0
        for sheet in self.input_files:
            try:
                with open(self.input_files[cc], 'w', encoding='utf-16', newline=''):
                    status.append(True)
            except(ValueError):
                status.append(False)
            cc = cc + 1
        return(status)

    # Working in progress - logs all the user input and actions
    def logger(self, Log):
        if self.logger_setup == 1:
            with open(self.log_path, 'a') as logger:
                time = datetime.now()
                logger.write(time.strftime("%y-%m-%d %H:%M:%S - "), Log)

    def logger_setup(self):
        self.log_path = os.path.join(GradeBook.log_path, "Log.txt")

        if os.path.exists(GradeBook.log_path):
            self.logger_setup = 1

            return True
        else:
            flag = 0


# The class holding all the information
gb = GradeBook()

if gb.dev == 1:
    flag = 0
    #gb.settings()
    # gb.menu()
else:
    gb.setup()
    gb.enter_student()
