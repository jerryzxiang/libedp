import pandas as pd
import matplotlib.pyplot as plt

# declare beginning of file paths as strings here for convenience
user_filepath = 'C:\\Users\\jerry\\' #change username based on computer
folder_filepath = 'Box\\PNE - R&D (Xiaofang)\\R&D Adminstration\\Projects\\LCO\\Experimental Data\\E-chem\\'

# load data
df = pd.ExcelFile(user_filepath + folder_filepath + 'PNE LCO Echem data summary.xlsx').parse('Performance Summary')
labels=[]
labels.append(df['Label'])
file_names=[] # create empty list of file names
file_names.append(df['File Name'])
active_mass=[]
active_mass.append(df['Active mass (g)'])
#print(file_names)

#color=['r', 'b','k','k']
for item in range(len(file_names[0])):
    if item > 230:
        if file_names[0][item] != 0:
    
            file  = user_filepath + folder_filepath + 'NDA and Excel files\\Excel\\' + file_names[0][item] + '.xls'
            mass = active_mass[0][item] # find corresponding active mass to file
            
            df2 = pd.ExcelFile(file).parse(sheet_name=1)
            
            numberOfCycles = len(df2['Capacity of discharge(mAh)'])
            FirstDischarge=df2.iat[0,3]/mass
            FirstEfficiency=100*df2.iat[0,3]/df2.iat[0,2]
            #print('First Cycle Discharge Capacity (mAh/g)=','%.2f' % FirstDischarge)
            #print('First Cycle Discharge Efficiency (%)=','%.2f' % FirstEfficiency)
            #secondDischargeCap = df2.iat[1,3]/mass
            
            xx1=df2['ToTal of Cycle']
            yy1=df2['Capacity of discharge(mAh)']
            plt.figure()
            plt.plot(xx1,yy1/mass, 'ro')
            plt.rcParams['xtick.labelsize']=12
            plt.rcParams['ytick.labelsize']=12
            plt.grid(True)
            plt.xlabel('Number of Cycles',fontsize=12)
            plt.ylabel('Capacity of discharge(mAh/g)',fontsize=12)
            plt.title(labels[0][item] + " " + file_names[0][item] + ' Capacity of discharge',fontsize=12)
            plt.xlim(0,len(xx1))
            plt.grid
            #if not os.path.exists(file_path):
            plt.savefig(user_filepath + folder_filepath + 'Results\\' + labels[0][item] + "_" + file_names[0][item] +"_capacity_of_discharge" + ".png", format="PNG", dpi=300, bbox_inches = "tight")
            #plt.show
            
            #capacityRetention = df2['Capacity of discharge(mAh)'] / secondDischargeCap
            #plt.plot(xx1, capacityRetention)
            
            #if numberOfCycles >= 200
            #    50DischargeCap = df.iat[49,3]/mass
            #    100DischargeCap = df.iat[99,3]/mass
            #    150DischargeCap = df.iat[149,3]/mass
            #    200DischargeCap = df.iat[199,3]/mass
            #    
            #    for 
            #    retention = 
                
            #if numberOfCycle >= 50:   
            #    FiftythDischarge=df.iat[49,3]/activemass[item]
            #    CycleLife=100*df.iat[49,3]/df.iat[1,3]
            #    print('2nd cycle Discharge Capacity (mAh/g)=','%.2f' % SecondDischarge)
            #    print('25th Discharge Capacity (mAh/g)=','%.2f' % FiftythDischarge)
            #    print('Cycle Life(%)=','%.2f' % CycleLife)
    
            df3=pd.ExcelFile(file).parse(sheet_name=3)
            first = df3.loc[df3['Cycle'] == 1]# select the first cycle
            
            vol=first['Voltage(V)']
            cap=first['CapaCity(mAh)']
            plt.figure()
            plt.plot(cap/mass,vol,'r',linestyle='none', marker='.', markersize=2)
            plt.rcParams['xtick.labelsize']=16
            plt.rcParams['ytick.labelsize']=16
            plt.xlabel('Capacity(mAh/g)',fontsize=16)
            plt.ylabel('Voltage(V)',fontsize=16)
            plt.title(labels[0][item] + " " + file_names[0][item] +' First Cycle', fontsize=16)
            
            plt.ylim(2.0,4.4)
            plt.grid()
            plt.savefig(user_filepath + folder_filepath + 'Results\\' + labels[0][item] + "_" + file_names[0][item] + "_first_cycle" + ".png", format="PNG", dpi=300, bbox_inches = "tight")
            #plt.show
        
        else: 
            continue