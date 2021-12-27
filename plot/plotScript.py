import pandas as pd
from pandas import DataFrame
from libplot import libplot as lp

"""
A script to plot discharge capacity vs cycle number and
first cycle charge/discharge for coin cells listed in 
a spreadsheet. This spreadsheet was made manually and
contains the 'Label', which is the name given to the
coin cell cathode sample by the author, the 'Active Mass (g)'
which is the active mass of the cathode sample in the cell,
the 'File Name', which is the name of the NDA file of the cell,
among other information.

This script uses the libplot library.
"""
# Change filepaths based on user preferences
user_filepath = r"/mnt/c/Users/xiang/" 
folder_filepath = "Documents/libedp/test_data/E-chem data (auto backup) - Copy2/"
save_file_path = user_filepath + folder_filepath
excel_file = "LCO_Echem_data_summary.xlsx"
sheet_name = "Performance Summary"

def main(user_filepath, folder_filepath, 
        save_file_path, excel_file, sheet_name):
    df = pd.ExcelFile(user_filepath + folder_filepath + 
                    excel_file).parse(sheet_name)

    labels = df['Label']
    file_names = df['File Name']
    active_mass = df['Active mass (g)']

    for item in range(len(file_names)):
        if item > 0:
            filename = file_names[item]
            label = labels[item]
            #numberOfCycles = len(df['Capacity of discharge(mAh)'])
            
            file = save_file_path + str(filename) + '.xls'
            mass = active_mass[item]
            
            if filename != 0:
                lp.plotDischargeCapacity(file, filename, mass, label, 
                                        showPlot=False,
                                        savePlot=True, 
                                        saveFilePath=save_file_path)
                lp.plotFirstCycle(file, filename, mass, label,
                                        showPlot=False,
                                        savePlot=True, 
                                        saveFilePath=save_file_path)

if __name__ = '__main__':
    main(user_filepath, folder_filepath, save_file_path, excel_file, sheet_name)