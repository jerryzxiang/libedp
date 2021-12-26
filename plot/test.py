import pandas as pd
from pandas import DataFrame
from libplot import libplot as lp

user_filepath = r"/mnt/c/Users/xiang/" #change username based on computer
folder_filepath = "Documents/libedp/test_data/E-chem data (auto backup) - Copy2/"
excel_file = "LCO_Echem_data_summary.xlsx"
sheet_name = "Performance Summary"
save_file_path = user_filepath + folder_filepath
df = pd.ExcelFile(user_filepath + folder_filepath + 
                  excel_file).parse(sheet_name)

labels=[]
labels.append(df['Label'])
file_names=[] # create empty list of file names
file_names.append(df['File Name'])
active_mass=[]
active_mass.append(df['Active mass (g)'])

for item in range(len(file_names[0])):
    if item > 0:
        filename = file_names[0][item]
        label = labels[0][item]
        #numberOfCycles = len(df['Capacity of discharge(mAh)'])
        
        file = save_file_path + str(filename) + '.xls'
        mass = active_mass[0][item]
        
        if filename != 0:
           lp.plotDischargeCapacity(file, filename, mass, label, savePlot=True, saveFilePath=save_file_path)
           lp.plotFirstCycle(file, filename, mass, label,
                                savePlot=True, saveFilePath=save_file_path)