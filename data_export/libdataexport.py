import pandas as pd
from pandas import DataFrame

class LIBDataExport():
    """

    """
    def __init__(self):
        self = self
    
    def export2excel(file, filename, mass, label):
        df = DataFrame({'Label': label, 'Active Mass':
                        mass})
        df = pd.DataFrame.from_dict(df, orient='index')
        df = df.transpose()
        df.to_excel(filename + " data_export.xlsx", sheet_name='sheet1',
                       index=False)
    
    def aggData():
        """
        Aggregate data
        """

    def getDataFrame1(file):
        df = pd.ExcelFile(file).parse(sheet_name = 1)

        return df

    def getSpecChrgCapacity(file, mass):
        df = pd.ExcelFile(file).parse(sheet_name = 1)
        # Discharge Capacity
        chrgCap = df['Capacity of charge(mAh)']
        # Specific Discharge Capacity
        specChrgCap = chrgCap / mass 

        return specChrgCap

    def getSpecDisCapacity(file, mass):
        df = pd.ExcelFile(file).parse(sheet_name = 1)
        # Discharge Capacity
        disCap = df['Capacity of discharge(mAh)']
        # Specific Discharge Capacity
        specDisCap = disCap / mass 

        return specDisCap
    
    def getCycNum(file):
        df = pd.ExcelFile(file).parse(sheet_name = 1)
        # Cycle Number
        cycNum = df['ToTal of Cycle']

        return cycNum

    def getFirstCycCapRet(file):
        df = pd.ExcelFile(file).parse(sheet_name = 1)
        firstCycCapRet = df['Cycle Life(%)']

        return firstCycCapRet

    def getSecondCycCapRet(file):
        df = pd.ExcelFile(file).parse(sheet_name = 1)
        # Discharge Capacity of Second Cycle
        secondCycDisCap = df.iat[1,3] 
        secondCycCapRet = df['Capacity of discharge(mAh)'] / secondCycDisCap * 100

        return secondCycCapRet
    
    def getFirstCycEfficiency(file):
        df = pd.ExcelFile(file).parse(sheet_name = 1)
        firstCycleEfficiency = 100 * df2.iat[0,3] / df2.iat[0,2]

        return firstCycleEfficiency
    
    def getFirstCycDisCap(file):
        df = pd.ExcelFile(file).parse(sheet_name = 1)
        firstCycleDischargeCapacity = df2.iat[0,3] / mass

        return firstCycleDischargeCapacity

    def getFirstCycChrgCap(file):
        df = pd.ExcelFile(file).parse(sheet_name = 1)
        firstCycleChargingCapacity = df2.iat[0,2] / mass

        return firstCycleChargingCapacity




    