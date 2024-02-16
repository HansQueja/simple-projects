"""
    BSCS 2-3 | Group 2's Fishbone Diagram Generator using Python Script

    Members:
    Blanquer, Jhoanna Joy
    Bonjibod, Jesus Evan
    Flores, Honey
    Manio, Jessica
    Maravilla, Ernest Matthew
    Prado, Mariella Alleli
    Queja, Hans Christian

    -----------------------------------------------------------------------------

    ABOUT THIS PYTHON SCRIPT

    This fishbone python script takes in an excel file for reading, and processes
    the given file and separates it into three main division:

    1. Main Problem
    2. Root Causes of the Problem
    3. Sub Root Causes based on the Root Causes
    4. Micro Root Causes based on the Sub Root Causes

    The data collected will then be plotted into the screen, following the layout
    of a fishbone diagram.

    -----------------------------------------------------------------------------
    
    PREREQUISITES OF THE SCRIPT
    
    Install the following modules:
        
        pip install pandas
	    pip install openpyxl
	    pip install matplotlib

    The following are the intricacies of the said modules:
    1. Pandas: Used to read the excel file and store the data needed for the fishbone
                diagram.
    2. Matplotlib: Used to plot lines and shapes into the screen by utilizing
                    its mathematical methods.
    3. Openpyxl: Co-dependency of pandas.

    -----------------------------------------------------------------------------

    Reminders before using the script:
    1. The program can only take four columns of data (index/row, Root Cause, Sub
        Root Cause, Micro Cause). Any more than that is not supported.
    2. The program can only take a minimum of one (1) root cause, and a maximum
        of ten (10) root causes.
    3. The maximum number of micro root causes is six (6).
    4. The maximum amount of sub root causes is six (6), and the minimum amount is 
        three (3).
    5. The excel file should start on the A1 cell.

    -----------------------------------------------------------------------------

    How to use the python scrpit?

    1. Execute the program on the command line by typing "python fishboneChallenge.py".
    2. Enter the file directory of where the excel file is.
    3. Enter the file name of the excel file.
    4. Enter your desired fishbone's output file name.

    After executing the said steps, you will find your PNG and PDF copy of the fishbone
    diagram on the same directory as the python file.

    Thank you, and have fun!

"""
# this is used to open the created output file
import os

# the remaining imports are the ones stated before
import pandas as pd 
import matplotlib.pyplot as plt
import numpy as np

def main():

    print("\n\nWelcome to the Fishbone Diagram Generator!\n")

    # prepares the data structures for the main root causes, the sub root causes, and the sub root causes of the subroots
    rootCauses = []
    semiRootCauses = []
    microRootCauses = {}

    # we call the function to read an excel file, and then stores the data on the given data structures
    rootCauses, semiRootCauses, microRootCauses, mainProblem = readExcelFile()

    # using the data we got from the excel file, we will plot them into the screen
    plotCauses(rootCauses, semiRootCauses, microRootCauses, mainProblem)
    

def plotCauses(rootCauses, semiRootCauses, microRootCauses, mainProblem):

    # takes the total amount of root causes from the file
    # having a count on the root causes will help us divide the different roots on the top and bottom side of the fishbone
    rootCausesCount = len(rootCauses)

    # the program only takes in
    if rootCausesCount < 1 or rootCausesCount > 10:
       print("Sorry, the current program can only take 1-10 root causes.")
       exit(1)

    # prepares the output file name
    outputFileName = ""

    # asks the user for desired output file name
    # if user enters nothing, program will reprompt the user for valid file name
    while True:

        outputFileName = input("\nEnter your desired fishbone output name: ")
    
        if outputFileName == "":
            print("Invalid input. Enter a string to the command line.")
            continue

        break

    # remove the toolbar, we don't really need to use it
    plt.rcParams['toolbar'] = 'None'
    fig = plt.figure()
    
    # this allows us to have a screen or frame to place our data
    ax = fig.add_subplot(111)

    # Setting the size of our screen
    X_EXTENDER = np.round(rootCausesCount / 4.0)

    # the following for loop counts to get the max count of semi root causes
    MAX_HEIGHT_COUNT = 0
    
    for rows in semiRootCauses:
        currentCount = 0
        for items in rows:
            currentCount += 1
        if currentCount >= MAX_HEIGHT_COUNT:
            MAX_HEIGHT_COUNT = currentCount

    # if the count of root causes is >= 9, set the following max parameters
    # if it is less than, set the following other max parameters
    if rootCausesCount >= 9:
        MAX_X_AXIS = 1500 + (500 * X_EXTENDER)
        MAX_Y_AXIS = 100 * MAX_HEIGHT_COUNT * 2
    else:
        MAX_X_AXIS = 1000 + (500 * X_EXTENDER)
        MAX_Y_AXIS = 100 * MAX_HEIGHT_COUNT * 2

    # sets the size of the window and the working space
    fig.set_size_inches(14, 8.5)
    ax.axis([0, MAX_X_AXIS, 0,  MAX_Y_AXIS])

    # Creates the main arrow in the middle
    mainLineX = [400, MAX_X_AXIS - 200]
    mainLineY = [MAX_Y_AXIS / 2, MAX_Y_AXIS / 2]

    # creates the arrow heads of the main arrow
    arrowX = [MAX_X_AXIS - 200 - 20, MAX_X_AXIS - 200 + 5, MAX_X_AXIS - 200 - 20]
    arrowY = [MAX_Y_AXIS / 2 + 20, MAX_Y_AXIS / 2, MAX_Y_AXIS / 2 - 20]

    # add the main line, and the main problem
    ax.plot(mainLineX, mainLineY, color='black', linewidth = 3, linestyle='-')
    ax.plot(arrowX, arrowY, color='black', linewidth = 3, linestyle='-')

    # adds the main problem at the end of the middle arrow
    ax.text(MAX_X_AXIS - 140, MAX_Y_AXIS / 2, mainProblem, style='oblique', bbox={'facecolor': 'green', 'alpha': 0.5, 'pad': 10})

    # places texts on the screen (the title and the subtitle)
    ax.set_title('Fishbone Diagram Generator', fontsize=15)

    defaultLineX = 200
    defaultRootX = 165
    defaultX = 200
    defaultTopY = MAX_Y_AXIS - 60

    # pang usog per root problem
    defaultAdder = 0
    defaultLineAdder = 0

    # ito pangbaba ng mga values ng subroot
    defaultSubtractor = 80
    indexer = 0
    subrootXAdder = 0
    xNudger = 0

    # divides the rootCause list into two other list, one for the top of the main line and one for the bottom of the line
    midPointRootCauses = len(rootCauses) // 2
    topSideRoot = rootCauses[:midPointRootCauses + len(rootCauses) % 2]
    topSideRootCount = len(topSideRoot)
    bottomSideRoot = rootCauses[midPointRootCauses + len(rootCauses) % 2:]

    for roots in topSideRoot:

        # draws the diagonal line
        ax.plot([defaultLineX + defaultLineAdder, defaultLineX + defaultLineX + defaultLineAdder + 50], [defaultTopY, MAX_Y_AXIS / 2], 
                color='black', linestyle='-')

        # places the root problem
        ax.text(defaultLineX - 50 + defaultLineAdder, defaultTopY + 10, roots, style='oblique', 
                bbox={'facecolor': 'green', 'alpha': 0.5, 'pad': 8})

        # checks if the subproblem has another subproblem.
        # if it has one, then we will configure that first to put other arrows
        # if not, we will place the subproblem in the frame
        for subroot in semiRootCauses[indexer]:

            # is a way to check if the root number aligns with the subproblem number
            subrootIndex = subroot + str(indexer + 1)

            # if the subproblem has subproblems (nested), do the following
            if subrootIndex in microRootCauses.values():

                # collects all the sub-subproblems linked to that subproblem
                matching_keys = [key for key, value in microRootCauses.items() if value == subrootIndex]

                # puts into the frame the main subproblem first
                ax.annotate(subroot, xy=(defaultRootX + defaultAdder + 70 + xNudger, defaultTopY - defaultSubtractor), 
                            xytext=(defaultRootX + defaultAdder - 250 + xNudger, defaultTopY - defaultSubtractor), 
                            fontsize=8,arrowprops=dict(arrowstyle = "->", connectionstyle = "angle"))

                microRootNudgerTop = 0
                microRootNudgerBottom = 0
                microRootCount = 1

                matchingKeysLength = len(matching_keys)

                # plots the micro causes in relation to the position of the sub root cause that they are under
                for values in matching_keys:
                    
                    values = values.rstrip(values[-1])

                    if microRootCount > matchingKeysLength // 2:
                        ax.annotate(values, xy=(defaultRootX + defaultAdder + microRootNudgerTop + xNudger - 100, defaultTopY - defaultSubtractor), 
                            xytext=(defaultRootX + defaultAdder + microRootNudgerTop + xNudger - 140, defaultTopY - defaultSubtractor + 30), 
                            fontsize=5,arrowprops=dict(arrowstyle = "->"))
                        microRootNudgerTop += 80
                    else:
                        ax.annotate(values, xy=(defaultRootX + defaultAdder + microRootNudgerBottom + xNudger - 100, defaultTopY - defaultSubtractor), 
                            xytext=(defaultRootX + defaultAdder + microRootNudgerBottom + xNudger - 140, defaultTopY - defaultSubtractor - 30), 
                            fontsize=5,arrowprops=dict(arrowstyle = "->"))
                        microRootNudgerBottom += 80

                    microRootCount += 1

                defaultSubtractor += 80
                xNudger += 45

            # if the sub root cause has no micro causes, plot them already to the plotting window.
            else:
                ax.annotate(subroot, xy=(defaultRootX + defaultAdder + 70 + xNudger, defaultTopY - defaultSubtractor), 
                            xytext=(defaultRootX + defaultAdder - 150 + xNudger, defaultTopY - defaultSubtractor), 
                            fontsize=8,arrowprops=dict(arrowstyle = "->", connectionstyle = "angle"))
                defaultSubtractor += 80
                xNudger += 45
        
        # increment the positional values to move the coordinates to their next proper place
        defaultLineAdder += 400
        xNudger = 0
        subrootXAdder += 310
        defaultSubtractor = 80
        defaultAdder += 400
        indexer += 1
        if indexer > topSideRootCount - 1:
            break

    # resets the variables for the bottom part of the fishbone
    defaultTopY = 200
    defaultRootY = 120
    defaultAdder = 0
    defaultLineAdder = 0
    defaultSubtractor = 0
    subrootXAdder = 0
    xNudger = 0

    # same structure as the top part of the fishbone, but in respect to 0,0 of the window
    for roots in bottomSideRoot:
        ax.plot([defaultLineX + defaultLineAdder + 10, defaultLineX + defaultLineX + defaultLineAdder + 50], 
                [90, MAX_Y_AXIS / 2], color='black', linestyle='-')
        

        ax.text(defaultX - 20 + defaultLineAdder, 60, roots, style='oblique', 
                bbox={'facecolor': 'green', 'alpha': 0.5, 'pad': 8})

        for subroot in semiRootCauses[indexer]:

            subrootIndex = subroot + str(indexer + 1)

            if subrootIndex in microRootCauses.values():

                matching_keys = [key for key, value in microRootCauses.items() if value == subrootIndex]

                ax.annotate(subroot, xy=(defaultRootX + defaultAdder + 60 + xNudger, defaultRootY + defaultSubtractor), 
                            xytext=(defaultRootX + defaultAdder - 270 + xNudger, defaultRootY + defaultSubtractor), 
                            fontsize=8,arrowprops=dict(arrowstyle = "->", connectionstyle = "angle"))

                microRootNudgerTop = 0
                microRootNudgerBottom = 0
                microRootCount = 1

                matchingKeysLength = len(matching_keys)

                for values in matching_keys:
                    
                    values = values.rstrip(values[-1])

                    if microRootCount > matchingKeysLength // 2:
                        ax.annotate(values, xy=(defaultRootX + defaultAdder + microRootNudgerTop + 10, defaultTopY + defaultSubtractor - 78), 
                        xytext=(defaultRootX + defaultAdder + microRootNudgerTop - 50, defaultTopY + defaultSubtractor - 55), 
                        fontsize=5,arrowprops=dict(arrowstyle = "->"))
                        microRootNudgerTop += 90
                    else:
                        ax.annotate(values, xy=(defaultRootX + defaultAdder + microRootNudgerBottom + 10, defaultTopY + defaultSubtractor - 78), 
                        xytext=(defaultRootX + defaultAdder + microRootNudgerBottom - 50, defaultTopY + defaultSubtractor - 110), 
                        fontsize=5,arrowprops=dict(arrowstyle = "->"))
                        microRootNudgerBottom += 80

                    microRootCount +=  1

                defaultSubtractor += 80
                xNudger += 45

            else:
                ax.annotate(subroot, xy=(defaultRootX + defaultAdder + 70 + xNudger, defaultRootY + defaultSubtractor), 
                            xytext=(defaultRootX + defaultAdder - 150 + xNudger, defaultRootY + defaultSubtractor), 
                            fontsize=8,arrowprops=dict(arrowstyle = "->", connectionstyle = "angle"))
                defaultSubtractor += 80
                xNudger += 45
        
        xNudger = 0
        subrootXAdder += 300
        defaultSubtractor = 0
        defaultAdder += 390
        defaultLineAdder += 400
        indexer += 1
        if indexer > rootCausesCount - 1:
            break

    # prepares the file names and type for the output
    outputFileNamePNG = "%s.png" %(outputFileName) 
    outputFileNamePDF = "%s.pdf" %(outputFileName) 

    # saves the plotted values into a PNG and a PDF.
    plt.axis('off')
    plt.savefig(outputFileNamePNG, dpi = 100)
    plt.savefig(outputFileNamePDF, dpi = 100)
    plt.show()
    
    os.startfile(outputFileNamePDF)
    os.startfile(outputFileNamePNG)

    print("\n\nYour PNG and PDF copy of the Fishbone Diagram has now been saved to the scripts directory.")
    print("Thank you for using the Python Script!\n")
    

def readExcelFile():

    # this reads the file that we want to process, and puts the contents inside the object called file
    fileDirectory = ""
    fileName = ""

    while True:

        fileDirectory = input("\n\nEnter the file directory of the excel file you want to use: ")
    
        if fileDirectory == "":
            print("Invalid input. Enter a string to the command line.")
            continue

        break

    while True:

        fileName = input("\nEnter the file name: ")
    
        if fileName == "":
            print("Invalid input. Enter a string to the command line.")
            continue

        break

    finalFileName = "%s\%s.xlsx" % (fileDirectory, fileName)
    print("Your file directory is: ", finalFileName)
    # if the file directory/file name is invalid, exit the program
    try:

        file = pd.read_excel(finalFileName)
    except FileNotFoundError:
        print("File is not found/invalid. Try again.")
        exit(0)

    # the line gets the total number of rows of the excel file
    total_rows = file.shape[0]
    total_columns = file.shape[1]

    if total_columns == 2:
        print("Invalid, provide sub causes to the excel sheet.")
        print()
        exit(0)
    elif total_rows < 6:
        print("Invalid. Provide more rows and causes to the excel sheet.")
        print()
        exit(0)
    
    newColumnNames = {}

    if total_columns == 4:
        # these are new column names for the columns that we have for the object file, replacing the null column names as well
        newColumnNames = {'Main Effect': 'MainEffect', 'Unnamed: 1': 'RootCause', 'Unnamed: 2': 'MiniCause', 'Unnamed: 3': 'MicroCause'}
    elif total_columns == 3:
        print("no microcause")
        newColumnNames = {'Main Effect': 'MainEffect', 'Unnamed: 1': 'RootCause', 'Unnamed: 2': 'MiniCause'}

    file.rename(columns=newColumnNames, inplace=True)

    # this puts the first cell from the file (the main problem always), on a string called mainProblem
    mainProblem = file.iloc[0,0]

    # these are lists and a dictionary where we're going to store the variables
    rootCauseList = []
    miniCauseList = []
    newMiniCauseRow = []
    
    microCausedict = {}

    # this is a for loop that checks where to place the values available
    for index, row in file.iterrows():


        if total_columns == 4:
            # if the value is a root cause, append it to the rootCause list
            if pd.isna(row['MiniCause']) and pd.isna(row['MicroCause']) and not pd.isna(row['RootCause']):
                rootCauseList.append(row['RootCause'])

                if not newMiniCauseRow:
                    continue
                miniCauseList.append(newMiniCauseRow)
                newMiniCauseRow = []

            if row['MainEffect'] == 1 and row['RootCause'] == 2 and pd.isna(row['MicroCause']):
                newMiniCauseRow.append(row['MiniCause'])

            if row['MainEffect'] == 1 and row['RootCause'] == 2 and row['MiniCause'] == 3:
                microCausedict[row['MicroCause'] + str(len(rootCauseList))] = newMiniCauseRow[-1] + str(len(rootCauseList))

            if index == total_rows - 1:
                miniCauseList.append(newMiniCauseRow)
        
        elif total_columns == 3:
            # if the value is a root cause, append it to the rootCause list
            if pd.isna(row['MiniCause']) and not pd.isna(row['RootCause']):
                rootCauseList.append(row['RootCause'])

                if not newMiniCauseRow:
                    continue
                miniCauseList.append(newMiniCauseRow)
                newMiniCauseRow = []

            if row['MainEffect'] == 1 and row['RootCause'] == 2:
                newMiniCauseRow.append(row['MiniCause'])

            if index == total_rows - 1:
                miniCauseList.append(newMiniCauseRow)
    
    return rootCauseList, miniCauseList, microCausedict, mainProblem


if __name__ == "__main__":
    main()