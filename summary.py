import pandas as pd
import xlrd
import numpy as np
from pandas import DataFrame
import matplotlib.pyplot as plt
import os

##os.chdir(r"E:\Portable Python-3.7.5 x64")
##pd.read_csv("sra.csv", delimiter=",").to_excel("sra.xlsx")
cdf=pd.read_excel("SRA_Questionnaire.xlsx")
cdf.drop(['Unnamed: 0','Stakeholder', 'Remark', 'System_User', 'user2', 'Date',],axis=1,inplace=True)

column_indices = range(0,6)
new_names = ['Unique_ID', 'Site_Name', 'GMP_Area', 'Identifier', 'Sub_System','Questions', 'Answer']
old_names=cdf.columns[column_indices]# Change column name
cdf.rename(columns=dict(zip(old_names,new_names)),inplace=True)
cdf.loc[:,'Answer_score'] = cdf['Answer']# Copy a column # copies column Answer into new column Answer_score
cdf["Answer_score"]=cdf["Answer_score"].str.title() # upper letter
cdf["Answer_score"] = cdf["Answer_score"].str.strip()# trim of column
cdf.Answer_score[cdf.Answer_score=="Yes"]=100
cdf.Answer_score[cdf.Answer_score=="No"]=0
cdf.Answer_score[cdf.Answer_score=="Partial"]=50
cdf["Sub_System"].replace("Manufacturing processes and controls","Manufacturing Processes and Controls",inplace=True)
cdf["Lowest_Score"]=cdf.groupby(["Site_Name","GMP_Area","Sub_System"])["Answer_score"].transform("min")
cdf["Final_score"]=cdf["Lowest_Score"]
l=len(cdf)-1


for i in range(0,l,1):
    if cdf.iloc[i,2]==cdf.iloc[i+1,2] and cdf.iloc[i,4]==cdf.iloc[i+1,4]:
        cdf.iloc[i,9]=np.NaN
print(len(cdf))


analysis=pd.pivot_table(cdf,values='Final_score',index=["GMP_Area"],aggfunc=np.mean)

analysis.sort_index(axis=1,ascending=False,inplace=True)# sort column name alphabetically

with pd.ExcelWriter("SRA_final.xlsx") as writer:
    cdf.to_excel(writer,sheet_name="tkinter_All_Site",index=False)
    analysis.to_excel(writer,sheet_name="analysis",float_format="%.2f")# putting decimal format




plt.figure(figsize=(8,4), dpi= 80)
plt.bar(analysis.index,analysis["Final_score"])
plt.title("Process & Procedures")
plt.ylabel("GMP Area Score")
plt.legend()
plt.show()
