import os,time
import pandas as pd
import numpy as np
start_time=time.time()
dir2 = input('请输入文件夹位置：')  #设置工作路径
dir= dir2.replace("\"", "\\")

#新建列表，存放每个文件数据框（每一个excel读取后存放在数据框,依次读取多个相同结构的Excel文件并创建DataFrame）
DFs = []

for root, dirs, files in os.walk(dir):  #第一个为起始路径，第二个为起始路径下的文件夹，第三个是起始路径下的文件。
    for file in files:
        file_path=os.path.join(root,file)  #将路径名和文件名组合成一个完整路径
        df = pd.read_csv(file_path,low_memory=False,sep=',', header= None) #excel转换成DataFrame
        df['Soure']=file
        df=df.replace(r'\,', np.nan, regex=True)
        DFs.append(df)
#合并所有数据，将多个DataFrame合并为一个
alldata = pd.concat(DFs)  #sort='False'
print(alldata.columns)
alldata2 = alldata[['Soure',0]]
alldata2.to_csv("csv_merge.csv",index = False,encoding="gbk")
ps_attach= pd.read_csv("csv_merge.csv",sep=":|,",skiprows=1,names=None)
ps_attach.columns=['City', 'Time','MODULE','SYSTYPE','SERVTYPE','PROC ','FAILMSG','NE','IMEISV','IMSI','MSISDN','Number','TAC','ExCAUSE','InCAUSE','APN','nn']
ps_attach['Cause ID']=ps_attach['ExCAUSE'].astype(str)+str("--")+ps_attach['InCAUSE'].astype(str)
ps_attach['MCCMNC']=ps_attach['IMSI'].str[0:5]
ps_attach=ps_attach[(ps_attach['PROC '] == "Attach procedure")]
ps_attach.groupby(['Cause ID'])['Cause ID'].count().to_csv("1.csv")
pd_at_su=pd.read_csv("1.csv")
ps_caused=pd.read_csv("Cause.csv",encoding="gbk")
print(ps_caused)
ps_attach_re=pd.merge(pd_at_su,ps_caused,on=['Cause ID'],how='outer')

ps_attach_re.to_excel("Attach Failure Analysis.xlsx",sheet_name="Summary")
res_file = 'Attach Failure Analysis.xlsx'
writer = pd.ExcelWriter(res_file,mode='a',index=False,engine='openpyxl')
ps_attach.groupby(['MCCMNC'])['FAILMSG'].count().to_excel(writer,sheet_name="MNC",index=True)
ps_attach.to_excel(writer,sheet_name="Log",index=False)
writer.save()
end_time=time.time()
times=round(end_time-start_time,2)
print('合并完成，耗时{}秒'.format(times))

