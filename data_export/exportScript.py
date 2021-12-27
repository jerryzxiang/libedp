import pandas as pd
from pandas import DataFrame
from libdataexport import libdataexport as lde

user_filepath = r"/mnt/c/Users/xiang/" #change username based on computer
folder_filepath = "Documents/libedp/test_data/E-chem data (auto backup) - Copy2/"
excel_file = "LCO_Echem_data_summary.xlsx"
sheet_name = "Performance Summary"

def main(user_filepath, folder_filepath, excel_file, sheet_name):
    save_file_path = user_filepath + folder_filepath
    df = pd.ExcelFile(user_filepath + folder_filepath + 
                    excel_file).parse(sheet_name)

    #labels=[]
    #labels.append(df['Label'])
    labels = df['Label']
    #file_names = [] # create empty list of file names
    #file_names.append(df['File Name'])
    file_names = df['File Name']
    #active_mass=[]
    #active_mass.append(df['Active mass (g)'])
    active_mass = df['Active mass (g)']

    for item in range(len(file_names)):
        if item > 0:
            filename = file_names[item]
            label = labels[item]
            #numberOfCycles = len(df['Capacity of discharge(mAh)'])
            
            file = save_file_path + str(filename) + '.xls'
            mass = active_mass[item]
            
            if filename != 0:
                print(lde.getSpecDisCapacity(file, filename, mass, label))

if __name__ == '__main__':
    main(user_filepath, folder_filepath, excel_file, sheet_name)