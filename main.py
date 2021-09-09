# main.py
'''
    Import an excel sheet
    Filter by Not Complete
    Filter by Supplier
    Filter by TOC
    Save to excel sheet
'''
import pandas as pd
from os import system, path, mkdir
from time import strftime

from variables import TOCS, SUPPLIERS, my_path

def header():
    # Just a console cleaner
    system('cls')

def main_function():
    # Base Variable
    the_day = strftime('%Y_%m_%d')  # Give a date stamp
    today = strftime("%Y_%m_%d-%H_%M_%S")  # Gives an initial date time stamp

    # Create datestamped folder
    dir = path.join(my_path, the_day)
    if not path.exists(dir):
        mkdir(dir)

    header()
    # Separate out the TOCs
    # Import the excel sheet
    df = pd.read_excel('raw.xlsx')
    df.columns = df.columns.str.replace(' ','_') # Sort column headers with spaces

    # Iterate through each TOC
    for item in TOCS:
        # Filters - There is probably a way to do multiple filters in one row
        newdf = df[df.TOC == item]
        newdf = newdf[newdf.Class_3_Status == 'Not Completed']
        newdf = newdf[newdf.TOC == item]
        # Iterate over Suppliers
        for sup in SUPPLIERS:
            # Create TOC datestamped folder
            dir = path.join(my_path, the_day, item)
            if not path.exists(dir):
                mkdir(dir)
            newdf = newdf[newdf.Supplier == sup]
            filename = str(dir + '/' + item + '_' + sup + '_' + today +'.xlsx')
            newdf.to_excel(filename) # Create file

    # Just to show the code finished clean
    print('Exit')

if __name__ == '__main__':
    main_function()



