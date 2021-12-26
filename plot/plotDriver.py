"""
Main section
"""

# parsing command line args
#cathode = sys.argv[1]#$print(cathode)

user_filepath = r'C:\\Users\\jerry\\' #change username based on computer

#if cathode == "LCO":
    # get filepaths
folder_filepath = 'Box\\PNE - R&D (Xiaofang)\\R&D Adminstration\\Projects\\LCO\\Experimental Data\\E-chem\\'
results_filepath = user_filepath + folder_filepath + 'Results\\'    
excelsheets_filepath = user_filepath + folder_filepath + 'NDA and Excel files\\Excel\\'
df = pd.ExcelFile(user_filepath + folder_filepath + 
                  'PNE LCO Echem data summary.xlsx').parse('Performance Summary')
#if cathode == "LNO":
#    folder_filepath = 'Box\\PNE - R&D (Xiaofang)\\R&D Adminstration\\Projects\\High Ni\\Experimental Data\\E-chem\\'
#


labels=[]
labels.append(df['Label'])
file_names=[] # create empty list of file names
file_names.append(df['File Name'])
active_mass=[]
active_mass.append(df['Active mass (g)'])

for item in range(len(file_names[0])):
    if item > 270:
        filename = file_names[0][item]
        label = labels[0][item]
        #numberOfCycles = len(df['Capacity of discharge(mAh)'])
        
        file = excelsheets_filepath + str(filename) + '.xls'
        mass = active_mass[0][item]
        
        if filename != 0:
           plotDischargeCapacity(file, filename, mass, label)
           plotFirstCycle(file, filename, mass, label)
           # plotCapacityRetention(file, filename, mass, label)
           # getCapacityRetention(file, filename, mass, label)
           # print(l50)
           # CapacityRetentionExcelExport(label_name, l50, l100, l150, l200)
