"""
Echem data plotting
"""
# import libraries
import pandas as pd
from pandas import DataFrame
import matplotlib.pyplot as plt

class libplot():
    """
    A spreadsheet is necessary.
    The spreadsheet should contain columns with:
    'label': Indicating the sample ID
    'mass': Indicating the active mass of the sample
    'filename': Indicating the name of the NDA file of the sample 
    'file': Is the excel file of the converted NDA file 
            of the sample. Created by exporting NDA file with 
            settings: Layer Report, File Format: EXCEL, 
            Export Way: Without EXCEL installed

    """
    def __init__(self):
        self = self
        
    def plotDischargeCapacity(file, filename, mass, 
                                label, savePlot=False, saveFilePath=None):
        """
        function to plot and save mass-normalized discharge capacity of one file
        """
        # parse first sheet of Excel file into pandas dataframe
        df = pd.ExcelFile(file).parse(sheet_name = 1)
        # get number of cycles from Excel file, which is number of columns
        cycleNumber = df['ToTal of Cycle']
        dischargeCapacity = df['Capacity of discharge(mAh)']
        
        plt.figure()
        plt.plot(cycleNumber, dischargeCapacity/mass, 'ro')
        plt.grid(True)
        plt.rcParams['xtick.labelsize'] = 12
        plt.rcParams['ytick.labelsize'] = 12
        plt.xlabel('Cycle Number', fontsize = 12)
        plt.ylabel('Discharge Capacity (mAh g${^-1}$)', fontsize = 12)
        plt.title(label + " " + filename + 
                ' Capacity of discharge', fontsize=12)
        plt.xlim(0, len(cycleNumber))
        plt.show()

        if savePlot is True:
            plt.savefig(saveFilePath + label + "_" + 
                        filename +"_capacity_of_discharge" 
                        + ".png", format="PNG", dpi=300, bbox_inches = "tight")
        
        # clear figs
        #plt.cla()
        #plt.clf()
        #plt.close()

    def plotFirstCycle(file, filename, mass, 
                        label, lowerVoltageLimit=2.0,
                        upperVoltageLimit=4.4,
                        savePlot=False, saveFilePath=None):
        """
        function to plot and save first cycle charge and discharge curve
        Default values of lower and upper voltage limits
        are 2.0 and 4.4 V respectively.
        """
        df = pd.ExcelFile(file).parse(sheet_name = 3)
        firstCycle = df.loc[df['Cycle'] == 1]
        firstCycleVoltage = firstCycle['Voltage(V)']
        firstCycleCapacity = firstCycle['CapaCity(mAh)']
        
        df2 = pd.ExcelFile(file).parse(sheet_name = 1)
        firstCycleDischargeCapacity = df2.iat[0,3] / mass
        firstCycleEfficiency = 100 * df2.iat[0,3] / df2.iat[0,2]
        
        plt.figure()
        plt.plot(firstCycleCapacity / mass, firstCycleVoltage, 'r', 
                linestyle = 'none', marker = '.', markersize = 2)
        plt.grid(True)
        plt.rcParams['xtick.labelsize'] = 16
        plt.rcParams['ytick.labelsize'] = 16
        plt.xlabel('Capacity (mAh g$^{-1}$)', fontsize = 16)
        plt.ylabel('Potential (V)', fontsize = 16)
        plt.title(label + " " + filename 
                 +' First Cycle', fontsize = 16)
        plt.ylim(lowerVoltageLimit, upperVoltageLimit) # voltage limits
        plt.show()

        if savePlot is True:
            plt.savefig(user_filepath + folder_filepath + 'Results\\' + label 
                        + "_" + filename + "_first_cycle" 
                        + ".png", format="PNG", dpi=300, bbox_inches = "tight")