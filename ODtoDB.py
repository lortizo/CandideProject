df1 = pd.read_excel(xlsfile, 'Matrix 6-7', skiprows=7, usecols=[1,4,5,6,7,8,9,10,11,12,13], header=0)
df1 = df1.set_index('Survey Time')
df1 = df1.iloc[2:12]
df1 = df1.stack().reset_index()
df1.columns = ['O','D','V_C1']
df1['OD_id'] = pd.factorize(df1.O+df1.D)[0]
df2 = pd.read_excel(xlsfile, 'Matrix 6-7', skiprows=25, usecols=[1,4,5,6,7,8,9,10,11,12,13], header=0)
df2 = df2.set_index('Survey Time')
df2 = df2.iloc[2:12]
df2 = df2.stack().reset_index()
df2.columns = ['O','D','V_C2']
df2['OD_id'] = pd.factorize(df2.O+df2.D)[0]
df2 = df2[['OD_id','V_C2']]
#######################################
df3 = pd.read_excel(xlsfile, 'Matrix 6-7', skiprows=43, usecols=[1,4,5,6,7,8,9,10,11,12,13], header=0)
df3 = df3.set_index('Survey Time')
df3 = df3.iloc[2:12]
df3 = df3.stack().reset_index()
df3.columns = ['O','D','V_C3']
df3['OD_id'] = pd.factorize(df3.O+df3.D)[0]
df3 = df3[['OD_id','V_C3']]
#######################################
df4 = pd.read_excel(xlsfile, 'Matrix 6-7', skiprows=61, usecols=[1,4,5,6,7,8,9,10,11,12,13], header=0)
df4 = df4.set_index('Survey Time')
df4 = df4.iloc[2:12]
df4 = df4.stack().reset_index()
df4.columns = ['O','D','V_C4']
df4['OD_id'] = pd.factorize(df4.O+df4.D)[0]
df4 = df4[['OD_id','V_C4']]
##3#########################
df10=pd.merge(df1,df2, on='OD_id')
df11=pd.merge(df4,df3, on='OD_id')
df12=pd.merge(df10,df11, on='OD_id')

#### Extracting additional info
df99 = pd.read_excel(xlsfile, 'Matrix 6-7', header=None)
#row, column
Date = df99.loc[2,2]
StartTime = df99.loc[3,2]
EndTime = df99.loc[3,5]
df12.insert(0,"Date", Date)
df12.insert(1,"StartTime", StartTime)
df12.insert(2,"EndTime", EndTime)
df12
