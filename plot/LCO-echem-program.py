import pandas as pd
import xlrd
import matplotlib.pyplot as plt
import os
import xlsxwriter

sample_path="C:\\Users\\jerry\\Box\\PNE - R&D (Xiaofang)\\R&D Adminstration\\Projects\\LCO\\Experimental Data\\E-chem\\Results\\"
files="C:\\Users\\jerry\\Box\\PNE - R&D (Xiaofang)\\R&D Adminstration\\Projects\\LCO\\Experimental Data\\E-chem\\NDA and Excel files\\Excel\\"
df = pd.ExcelFile("C:\\Users\\jerry\\Box\\PNE - R&D (Xiaofang)\\R&D Adminstration\\Projects\\LCO\\Experimental Data\\E-chem\\PNE LCO Echem data summary.xlsx").parse("Performance Summary")

def getdict():
  rfc=[]
  for i in range(len(df)):
    label=df.iat[i,3]
    name=df.iat[i,6]
    dic=[]
    if rfc.count([label])==0:
      dic.append(label)
      rfc.extend([dic])
    for k in range(len(rfc)):
      if rfc[k][0]==label:
        rfc[k].append(name)
  #for x in range(len(rfc)):
  #  print(rfc[x])
  #print(rfc[10][0])
  return rfc

def getlab():
  labexist=False
  chlab=input("choose label: ")
  for x in range(len(rfc)):
      if rfc[x][0]==chlab:
        labexist=True

  while labexist==False:
    print("label does not exist")
    chlab=input("choose label: ")
    for x in range(len(rfc)):
      if rfc[x][0]==chlab:
        labexist=True
  return chlab

def getcyc():
  input_string = input('select cycles:')
  user_list = input_string.split()
  for i in range(len(user_list)):
    user_list[i] = int(user_list[i])
  return user_list

def getmass():
  masstab=[]
  for i in range(len(df)):
    mass=df.iat[i,9]
    name=df.iat[i,6]
    dic=[]
    dic.append(name)
    dic.append(mass)
    masstab.extend([dic])
  return masstab

def table():
  newpath = sample_path + str(chlab) + "\\"
    
  if not os.path.exists(newpath):
      os.makedirs(newpath)    
      
  filepath1=newpath + str(chlab) + " data.xlsx"

  workbook = xlsxwriter.Workbook(filepath1)
  worksheet = workbook.add_worksheet()

  for x in range(len(rfc)):
    if rfc[x][0]==chlab:
      table = [["label","Name","Capacity of Charge(mAh)","Capacity of Discharge(mAh)","Efficiency(%)","Energy of Charge(mWh)","Energy of Discharge(mWh)"]]
      vrb=0
      for file in os.scandir(files):
        names=os.listdir(files)
        names=' '.join(names).replace('.xls','').split()
        if rfc[x].count(names[vrb])==1:
          data=[]
          data.append(chlab)

          name=names[vrb]
          data.append(name)
          
          for e in range(len(masstab)):
            if masstab[e][0]==name:
              mass=masstab[e][1]
          
          df2 = pd.ExcelFile(file).parse(sheet_name=1)
          FirstCharge=df2.iat[0,2]
          FirstDischarge=df2.iat[0,3]/mass
          FirstEfficiency=100*df2.iat[0,3]/df2.iat[0,2]
          data.append(FirstCharge)
          data.append(FirstDischarge)
          data.append(FirstEfficiency)

          df3=pd.ExcelFile(file).parse(sheet_name=3)

          chargenergy=[]
          dischargenergy=[]
          wb = xlrd.open_workbook(file)
          sheet = wb.sheet_by_index(3)
          for i in range(sheet .nrows):
            if i<sheet.nrows-1:
              cycle=df3.iat[i,3]
              if cycle==1:
                jumpto=df3.iat[i,2]
                energy=df3.iat[i,8]
                if jumpto==2:
                  chargenergy.append(energy)
                elif jumpto==4:
                  dischargenergy.append(energy)

          if len(chargenergy)>0:
            data.append(chargenergy[-1])
          
          if len(dischargenergy)>0:
            data.append(dischargenergy[-1])

          table.extend([data])

        vrb+=1

  for row_num, row_data in enumerate(table):
      for col_num, col_data in enumerate(row_data):
          worksheet.write(row_num, col_num, col_data)

  workbook.close()

  return table
      

def capofdis(file,vrb1,mass,names):
  newpath = sample_path + str(chlab) + "\\"
    
  if not os.path.exists(newpath):
      os.makedirs(newpath)    

  file_path = newpath + str(chlab) + " " + str(names[vrb1]) + " Capacity of discharge.png"
  
  df = pd.ExcelFile(file).parse(sheet_name=1)

  xx1=df['ToTal of Cycle']
  yy1=df['Capacity of discharge(mAh)']
  plt.figure()
  plt.plot(xx1,yy1/mass, 'r')
  plt.rcParams['xtick.labelsize']=14
  plt.rcParams['ytick.labelsize']=14
  plt.grid(True)
  plt.xlabel('Number of Cycles',fontsize=14)
  plt.ylabel('Capacity of discharge(mAh/g)',fontsize=14)
  plt.title('Capacity of discharge',fontsize=14)
  plt.xlim(0,51)
  plt.grid
  plt.savefig(file_path, format="PNG", dpi=300, bbox_inches = "tight")

def CyclesPlot(file,x,vrb1,mass,names):
  newpath = sample_path + str(chlab) + "\\"
    
  if not os.path.exists(newpath):
      os.makedirs(newpath)    
    
  file_path2 = newpath + str(chlab) + " " + str(names[vrb1]) + " cycles " + str(cycles) + " plot.png"
  
  df=pd.ExcelFile(file).parse(sheet_name=3)
  first = df.loc[df['Cycle'] == cycles[x]]# select the first cycle

  vol=first['Voltage(V)']
  cap=first['CapaCity(mAh)']
  style=["r","g","b","c","m","y","k","w"]
  plt.plot(cap/mass,vol,style[x],linestyle='none', marker='.', markersize=2)
  plt.rcParams['xtick.labelsize']=16
  plt.rcParams['ytick.labelsize']=16
  plt.xlabel('Capacity(mAh/g)',fontsize=16)
  plt.ylabel('Voltage(V)',fontsize=16)
  plt.title(str(names[vrb1])+' Cycle #'+ str(cycles), fontsize=16)
  plt.ylim(2.0,4.4)
  plt.grid()
  plt.savefig(file_path2, format="PNG", dpi=300, bbox_inches = "tight")

def graph():
  for k in range(len(rfc)):
    if rfc[k][0]==chlab:
      vrb1=0
      for file in os.scandir(files):
        names=os.listdir(files)
        names=' '.join(names).replace('.xls','').split()
        if rfc[k].count(names[vrb1])==1:
          
          print(names[vrb1])

          for e in range(len(masstab)):
            if masstab[e][0]==names[vrb1]:
              mass=masstab[e][1]
        
          capofdis(file,vrb1,mass,names)

          plt.figure()
          for x in range(len(cycles)):
            CyclesPlot(file,x,vrb1,mass,names)
        
        vrb1+=1
          
  return plt.show()


rfc=getdict()
masstab=getmass()
chlab=getlab()
table()
cycles=getcyc()
graph()

