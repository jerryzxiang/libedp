"""
Echem data plotting
"""
# import libraries
import pandas as pd
from pandas import DataFrame
import matplotlib.pyplot as plt
import sys

class libplot():
    """
    A spreadsheet is necessary. The spreadsheet is named 'file'.
    The spreadsheet should contain columns with:
    'label': Indicating the sample ID
    'mass': Indicating the active mass of the sample
    'filename': Indicating the name of the NDA file of the sample 
    """
    def __init__(self, file, filename, mass, label, savePlot=True):
        self.file = file
        self.filename = filename
        self.mass = mass
        self.label = label
        self.savePlot = savePlot if savePlot is not True else False

    # function to plot and save mass-normalized discharge capacity of one file
    def plotDischargeCapacity(file, filename, mass, label):
        # parse Excel file into pandas dataframe
        df = pd.ExcelFile(file).parse(sheet_name = 1)
        # get number of cycles from Excel file, which is number of columns
        cycleNumber = df['ToTal of Cycle']
        dischargeCapacity = df['Capacity of discharge(mAh)']
        
        plt.figure()
        plt.plot(cycleNumber, dischargeCapacity/mass, 'ro')
        plt.rcParams['xtick.labelsize']=12
        plt.rcParams['ytick.labelsize']=12
        plt.grid(True)
        plt.xlabel('Cycle Number',fontsize=12)
        plt.ylabel('Discharge Capacity (mAh g${^-1}$)',fontsize=12)
        #plt.title(label + " " + filename + 
                #' Capacity of discharge',fontsize=12)
        plt.xlim(0,len(cycleNumber))
        plt.grid
        plt.savefig(results_filepath + label + "_" + 
                    filename +"_capacity_of_discharge" 
                    + ".png", format="PNG", dpi=300, bbox_inches = "tight")
        
        # clear figs
        plt.cla()
        plt.clf()
        plt.close()

        
    # function to plot and save first cycle charge and discharge curve
    def plotFirstCycle(file, filename, mass, label):
        df = pd.ExcelFile(file).parse(sheet_name = 3)
        firstCycle = df.loc[df['Cycle'] == 1]
        firstCycleVoltage = firstCycle['Voltage(V)']
        firstCycleCapacity = firstCycle['CapaCity(mAh)']
        
        df2 = pd.ExcelFile(file).parse(sheet_name = 1)
        firstCycleDischargeCapacity = df2.iat[0,3] / mass
        firstCycleEfficiency = 100 * df2.iat[0,3] / df2.iat[0,2]
        
        plt.figure()
        plt.plot(firstCycleCapacity/mass, firstCycleVoltage, 'r', 
                linestyle = 'none', marker = '.', markersize = 2)
        plt.rcParams['xtick.labelsize']=16
        plt.rcParams['ytick.labelsize']=16
        plt.xlabel('Capacity (mAh g$^{-1}$)',fontsize=16)
        plt.ylabel('Potential (V)',fontsize=16)
        #plt.title(label + " " + filename 
                # +' First Cycle', fontsize = 16)
        plt.ylim(2.0,4.4) # voltage limits
        # add discharge capacity and efficiency to figure
        #plt.figtext(.15, .2, "1st Cycle Efficiency = "
        #            + str(round(firstCycleEfficiency, 2)) + "%", 
        #         fontsize=6)
        #plt.figtext(.15, .15, "1st Cycle Discharge Capacity = " 
        #            + str(round(firstCycleDischargeCapacity,2))
        #            + " mAh/g", fontsize=6)
        plt.grid()
        plt.savefig(user_filepath + folder_filepath + 'Results\\' + label 
                    + "_" + filename + "_first_cycle" 
                    + ".png", format="PNG", dpi=300, bbox_inches = "tight")
        # clear figs
        plt.cla()
        plt.clf()
        plt.close()


    # get capacity retention (from second cycle)
    def plotCapacityRetention(file, filename, mass, label):
        df = pd.ExcelFile(file).parse(sheet_name = 1)
        cycleNumber = df['ToTal of Cycle']
        dischargeCapacity = df['Capacity of discharge(mAh)'] # mAh
        secondCycleDischargeCapacity = df.iat[1,3] # mAh
        
        # get array of capacity retention as % (using second cycle as base)
        capacityRetention = 100 * dischargeCapacity / secondCycleDischargeCapacity
        
        plt.figure()
        plt.plot(cycleNumber, capacityRetention, 'ro')
        plt.rcParams['xtick.labelsize']=12
        plt.rcParams['ytick.labelsize']=12
        plt.grid(True)
        plt.xlabel('Number of Cycles',fontsize=12)
        plt.ylabel('Capacity Retention (from 2nd Cycle) [%]',fontsize=12)
        plt.title(label + " " + filename + 
                ' capacity retention',fontsize=12)
        plt.xlim(0, len(cycleNumber))
        plt.grid
        plt.savefig(results_filepath + label + "_" + 
                    filename +"_capacity_retention" 
                    + ".png", format="PNG", dpi=300, bbox_inches = "tight")
        
        """
        if cycleNumber > 200:
            fiftyCycleDischargeCapacity = df.iat[49,3] / mass
            oneHundredDischargeCapacity = df.iat[99,3] / mass
            oneHundredFiftyDischargeCap = df.iat[149,3] / mass
            twoHundredDischargeCapacity = df.iat[199,3] / mass
        """
        
        # clear figs
        plt.cla()
        plt.clf()
        plt.close()
        
        
    label_name = []
    l50 = []
    l100 = []
    l150 = []
    l200 = []
    def getCapacityRetention(file, filename, mass, label):
        df = pd.ExcelFile(file).parse(sheet_name = 1)
        cycleNumber = df['ToTal of Cycle']
        dischargeCapacity = df['Capacity of discharge(mAh)'] # mAh
        secondCycleDischargeCapacity = df.iat[1,3] # mAh
        
        # get array of capacity retention as % (using second cycle as base)
        capacityRetention = 100 * dischargeCapacity / secondCycleDischargeCapacity
        numberOfCycles = len(df['Capacity of discharge(mAh)'])
        
        label_name.append(label)
        if numberOfCycles > 200:
            fiftyCycleDischargeCapacity = df.iat[49,3] / mass
            oneHundredDischargeCapacity = df.iat[99,3] / mass
            oneHundredFiftyDischargeCap = df.iat[149,3] / mass
            twoHundredDischargeCapacity = df.iat[199,3] / mass
            
            fiftyCycleRetention = fiftyCycleDischargeCapacity/secondCycleDischargeCapacity
            oneHundredRetention = oneHundredDischargeCapacity/secondCycleDischargeCapacity
            oneHundredFiftyRetention = oneHundredFiftyDischargeCap/secondCycleDischargeCapacity
            twoHundredRetention = twoHundredDischargeCapacity/secondCycleDischargeCapacity
            l50.append(fiftyCycleRetention)
            l100.append(oneHundredRetention)
            l150.append(oneHundredFiftyRetention)
            l200.append(twoHundredRetention)
            
        elif numberOfCycles > 150:
            fiftyCycleDischargeCapacity = df.iat[49,3] / mass
            oneHundredDischargeCapacity = df.iat[99,3] / mass
            oneHundredFiftyDischargeCap = df.iat[149,3] / mass
            
            fiftyCycleRetention = fiftyCycleDischargeCapacity/secondCycleDischargeCapacity
            oneHundredRetention = oneHundredDischargeCapacity/secondCycleDischargeCapacity
            oneHundredFiftyRetention = oneHundredFiftyDischargeCap/secondCycleDischargeCapacity
            l50.append(fiftyCycleRetention)
            l100.append(oneHundredRetention)
            l150.append(oneHundredFiftyRetention)
            l200.append(0)
            
        elif numberOfCycles > 100:
            fiftyCycleDischargeCapacity = df.iat[49,3] / mass
            oneHundredDischargeCapacity = df.iat[99,3] / mass
            
            fiftyCycleRetention = fiftyCycleDischargeCapacity/secondCycleDischargeCapacity
            oneHundredRetention = oneHundredDischargeCapacity/secondCycleDischargeCapacity
            l50.append(fiftyCycleRetention)
            l100.append(oneHundredRetention)
            l150.append(0)
            l200.append(0)
        elif numberOfCycles > 50:
            fiftyCycleDischargeCapacity = df.iat[49,3] / mass
            fiftyCycleRetention = fiftyCycleDischargeCapacity/secondCycleDischargeCapacity
            l50.append(fiftyCycleRetention)
            l100.append(0)
            l150.append(0)
            l200.append(0)

    def CapacityRetentionExcelExport(label_name, l50, l100, l150, l200):
        df = DataFrame({'Label': label_name, '50th cycle cap ret': l50, 
                    '100th cycle cap ret': l100,'150th cycle cap ret': l150,
                    '200th cycle cap ret': l200})
        df = pd.DataFrame.from_dict(df,orient='index')
        df = df.transpose()
        df.to_excel('cap_retention.xlsx', sheet_name='sheet1', index = False)
    
"""
Main section
"""

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
"""