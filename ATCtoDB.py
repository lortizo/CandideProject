import pandas as pd
import os
import glob

appended_data = []
for xlsfile in glob.glob(os.path.join('.','*.xlsm')):

	df1 = pd.read_excel(xlsfile, 'Dir1', skiprows=12, usecols=[2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,24,25,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,78,79,81,82],
                   headers=None, names=['Date','Period Start','Total','C1','C2','C3','C4','C5','C6','C7','C8','C9','C10','C11','C12','C13','C14','C15','MSPD','MaxSPD','MinSPD','85th% SPD','50th% SPD','S0-10','S10-20','S20-30','S30-40','S40-50','S50-60','S60-70','S70-80','S80-90','S90-100','S100-110','S110-120','S120-130','S130-140','S140-150','S150-160','S160-170','S170-180','S180-190','Sover190','CarAvgSpd','Car85%','TruckAvgSpd','Truck85%'])
    # 85,86,87,88,89
    # 'Car50%','Truck50%','CarMinSpeed','CarMaxSpeed','TruckMinSpeed'
	df1 = df1.dropna(axis=0,how='any')
	df0 = pd.read_excel(xlsfile, 'Dir1')
	Location = df0.iloc[0,3]
	Street = df0.iloc[2,3]
	Suburb = df0.iloc[3,3]
	Direction = df0.iloc[9,3]
	df1.insert(0,"Location", Location)
	df1.insert(1,"Road", Street)
	df1.insert(2,"Suburb", Suburb)
	df1.insert(3,"Direction", Direction)

	df2 = pd.read_excel(xlsfile, 'Dir2', skiprows=12, usecols=[2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,24,25,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,78,79,81,82],
                   headers=None, names=['Date','Period Start','Total','C1','C2','C3','C4','C5','C6','C7','C8','C9','C10','C11','C12','C13','C14','C15','MSPD','MaxSPD','MinSPD','85th% SPD','50th% SPD','S0-10','S10-20','S20-30','S30-40','S40-50','S50-60','S60-70','S70-80','S80-90','S90-100','S100-110','S110-120','S120-130','S130-140','S140-150','S150-160','S160-170','S170-180','S180-190','Sover190','CarAvgSpd','Car85%','TruckAvgSpd','Truck85%'])
	df2 = df2.dropna(axis=0,how='any')
	df0 = pd.read_excel(xlsfile, 'Dir2')
	Location = df0.iloc[0,3]
	Street = df0.iloc[2,3]
	Suburb = df0.iloc[3,3]
	Direction = df0.iloc[9,3]
	df2.insert(0,"Location", Location)
	df2.insert(1,"Road", Street)
	df2.insert(2,"Suburb", Suburb)
	df2.insert(3,"Direction", Direction)

	df99 = df1.append(df2)
	appended_data.append(df99)
appended_data = pd.concat(appended_data)

#df99.to_csv("P://V17200-17299//V172700 Eastern Bendigo Network Enhancements//External//Traffic data//TrafficDataOD//5691_Q45491508EastBendigoTrafficSurvey_ODLocations//x.csv")
#appended_data.to_csv("E://LeoPyScripts//GTA_DB.csv")
appended_data.to_csv('GTA_DB.csv')
