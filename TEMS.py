import pandas as pd
import os
import time
import numpy as np

with open("UE Capability Information (UL-DCCH).txt",'r') as f1,open('2.txt', 'w') as f2:
    for line in f1.readlines():
        line = line.strip()
        a ="bandEUTRA-r10 :"
        b = "BandCombinationParameters-r10"
        if b in line:
            f2.write(line + '\n')
        if a in line:
            f2.write(line + '\n')
df =pd.read_csv("2.txt",header=None, names=['CA_Sub_Group'])
#df['bandEUTRA_r10']=df['CA_Sub_Group']
#print(df)
#df['bandEUTRA_r10'].str.split(' : ',expand=True)
#df['CA_Group'].fillna(value= b ,inplace = True)
#band = "bandEUTRA_r10"
#df_mask=df['bandEUTRA_r10']= band
#positions = np.flatnonzero(df_mask)
#filtered_df=df.iloc[positions]
#print(filtered_df)
df.to_excel("Tems CA ueCapabilityInformation " + str(time.strftime("%Y-%m-%d %H-%M-%S")) + ".xlsx", index=False)
os.remove("2.txt")

