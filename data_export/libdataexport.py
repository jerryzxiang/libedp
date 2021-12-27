import pandas as pd
from pandas import DataFrame

class libdataexport():
    """

    """
    def __init__(self):
        self = self
    
    def export2csv(file, filename, mass, label):
        df = DataFrame({'Label': label, 'Active Mass':
                        mass})
        df = pd.DataFrame.from_dict(df, orient='index')
        df = df.transpose()
        df.to_excel(filename + " data_export.xlsx", sheet_name='sheet1',
                       index=False)
    
    def getSpecDisCapacity(file, filename, mass, label):
        df = pd.ExcelFile(file).parse(sheet_name = 1)
        # Cycle Number
        cycNum = df['ToTal of Cycle']
        # Discharge Capacity
        disCap = df['Capacity of discharge(mAh)']
        # Specific Discharge Capacity
        specDisCap = disCap / mass 

        return specDisCap
        