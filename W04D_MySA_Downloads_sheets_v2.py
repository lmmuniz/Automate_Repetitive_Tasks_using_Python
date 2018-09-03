# -*- coding: utf-8 -*-
"""
Created on Fri May 18 14:11:42 2018

@author: julio
"""

import pandas as pd
import os
import datetime
import csv
import shutil
import time

working_dir            = r'D:\\brrt\\data'
apple_downloads_file   = 'MySA_Apple_downloads.xlsx'
android_downloads_file = 'MySA_Android_downloads.xlsx'
W04_downloads_file     = 'W04B_MySA_downloads.csv'
GD_flag_file           = 'W04D_inet_gd_flag.csv'
GD_flag = {}

os.chdir(working_dir)

def mask_NONEU(df):
   
    new_df = df.copy()
    new_df.reset_index(inplace=True)

    for ix in range(len(new_df)-1):
        
        iot = new_df.ix[ix,'IOT']
        #country = new_df.loc[ix,['Country']]
        
        if iot == 'Europe IOT':
            new_df.ix[ix,'Email']    = '*Masked E-mail*'
    new_df.drop('index',1,inplace=True)
    return new_df



def read_gd_flag_file(file):
    reader = csv.reader(open(GD_flag_file, 'r'))
    GD_flag = {}
    for row in reader:
        k, v = row
        GD_flag[k] = v
    return GD_flag

def read_downloads_df(file):
    df = pd.read_excel(file)
    df.columns=df.loc[2]
    df=df.drop(df.index[:3])
    return df

def get_gd_flag(row):
    global GD_flag
    iid = row['Email']
    if iid in GD_flag:
        return GD_flag[iid]
    else:
        return ''

def major_group(row):
    if row['HR Group ID']  == 'GBS':
        return('GBS')
    elif row['HR Group ID'] == 'GTS':
        return('GTS')
    else:
        return('Other')
        
def gbs_adoption_group(row):
    if row['GD Flag']  == 'Y' and row['IOT'] not in ('Latin America IOT','Greater China Group IOT'):
        return('CIC')
    elif row['IOT'] == 'ASIA Pacific IOT':
        return('AP IOT')
    elif row['IOT'] == 'Europe IOT':
        return('EU IOT')
    elif row['IOT'] == 'Greater China Group IOT':
        return('GCG IOT')
    elif row['IOT'] == 'Japan IOT':
        return('JP IOT')
    elif row['IOT'] == 'Latin America IOT':
        return('LA IOT')
    elif row['IOT'] == 'Middle East and Africa IOT':
        return('MEA IOT')
    elif row['IOT'] == 'North America IOT':
        return('NA IOT')


def process_W04D():
    #read MySA Apple Downloads file
    df1 = read_downloads_df(apple_downloads_file)
    df1['OS Type']='Apple'
    
    #read MySA Android Downloads file
    df2 = read_downloads_df(android_downloads_file)
    df2['OS Type']='Android'
    
    #Merge MySA Apple and Android Downloads Files in a single dataframe
    df  = pd.concat([df1,df2])
    
    #Make Email field lower for lookup with Metrics list of emails to find the GD_FLG value
    df['Email']=df['Email'].str.lower()
    #assign GD_flag as global variable
    global GD_flag
    GD_flag = read_gd_flag_file(GD_flag_file)
    df['GD Flag']=df.apply(get_gd_flag,axis=1)
    
    #Create extra fields based on analysis of functions below
    df['Major Group'] = df.apply(major_group,axis=1)
    df['GBS Adoption Group'] = df.apply(gbs_adoption_group,axis=1)
    df['Template Version']=1
    df['Date of Data Retrieval']=datetime.date.today().strftime("%Y-%m-%d")
    
    #generate W04D export file for W04D MySA Download report
    #df.to_csv(W04_downloads_file,index=False)
    
    #generate: W04 Troubleshooter MySA Downloads
    df_W04D_TS = df
    df_W04D_TS.to_csv('W04B_MySA_downloads_TS.csv',index=False)
    
    #generate: W04 GBS North America MySA Downloads
    df_W04D_GBS_NA = df[(df['Major Group']=='GBS') & (df['GBS Adoption Group'] =='NA IOT')]
    df_W04D_GBS_NA.to_csv('W04B_MySA_downloads_GBS_NA.csv',index=False)
    
    #generate: W04 GTS North America MySA Downloads
    df_W04D_GTS_NA = df[(df['Major Group']=='GTS') & (df['GBS Adoption Group'] =='NA IOT')]
    df_W04D_GTS_NA.to_csv('W04B_MySA_downloads_GTS_NA.csv',index=False)
    
    #generate: W04 GBS Europe RDM MySA Downloads
    df_W04D_GBS_EP = df[(df['Major Group']=='GBS') & (df['GBS Adoption Group'] =='EU IOT')]
    df_W04D_GBS_EP.to_csv('W04B_MySA_downloads_GBS_EP.csv',index=False)
    
    #generate: W04 GTS Europe RDM MySA Downloads
    df_W04D_GTS_EP = df[(df['Major Group']=='GTS') & (df['GBS Adoption Group'] =='EU IOT')]
    df_W04D_GTS_EP.to_csv('W04B_MySA_downloads_GTS_EP.csv',index=False)    

    #generate: W04 GBS Japan MySA Downloads
    df_W04D_GBS_JP = df[(df['Major Group']=='GBS') & (df['GBS Adoption Group'] =='JP IOT')]
    df_W04D_GBS_JP.to_csv('W04B_MySA_downloads_GBS_JP.csv',index=False)
    
    #generate: W04 GTS Japan MySA Downloads
    df_W04D_GTS_JP = df[(df['Major Group']=='GTS') & (df['GBS Adoption Group'] =='JP IOT')]
    df_W04D_GTS_JP.to_csv('W04B_MySA_downloads_GTS_JP.csv',index=False)  
    
    #generate: W04 GBS Asia Pacific MySA Downloads
    df_W04D_GBS_AP = df[(df['Major Group']=='GBS') & (df['GBS Adoption Group'] =='AP IOT')]
    df_W04D_GBS_AP.to_csv('W04B_MySA_downloads_GBS_AP.csv',index=False)
    
    #generate: W04 GTS Asia Pacific MySA Downloads
    df_W04D_GTS_AP = df[(df['Major Group']=='GTS') & (df['GBS Adoption Group'] =='AP IOT')]
    df_W04D_GTS_AP.to_csv('W04B_MySA_downloads_GTS_AP.csv',index=False)     
    
    #generate: W04 GBS Greater China Group MySA Downloads
    df_W04D_GBS_GCG = df[(df['Major Group']=='GBS') & (df['GBS Adoption Group'] =='GCG IOT')]
    df_W04D_GBS_GCG.to_csv('W04B_MySA_downloads_GBS_GCG.csv',index=False)
    
    #generate: W04 GBS Greater China Group MySA Downloads
    df_W04D_GTS_GCG = df[(df['Major Group']=='GTS') & (df['GBS Adoption Group'] =='GCG IOT')]
    df_W04D_GTS_GCG.to_csv('W04B_MySA_downloads_GTS_GCG.csv',index=False)     
    
    #generate: W04 GBS Latin America MySA Downloads
    df_W04D_GBS_LA = df[(df['Major Group']=='GBS') & (df['GBS Adoption Group'] =='LA IOT')]
    df_W04D_GBS_LA.to_csv('W04B_MySA_downloads_GBS_LA.csv',index=False)
    
    #generate: W04 GTS Latin America MySA Downloads
    df_W04D_GTS_LA = df[(df['Major Group']=='GTS') & (df['GBS Adoption Group'] =='LA IOT')]
    df_W04D_GTS_LA.to_csv('W04B_MySA_downloads_GTS_LA.csv',index=False)       
    
    #generate: W04 GBS Middle East & Africa MySA Downloads
    df_W04D_GBS_MEA = df[(df['Major Group']=='GBS') & (df['GBS Adoption Group'] =='MEA IOT')]
    df_W04D_GBS_MEA.to_csv('W04B_MySA_downloads_GBS_MEA.csv',index=False)
    
    #generate: W04 GTS Middle East & Africa MySA Downloads
    df_W04D_GTS_MEA = df[(df['Major Group']=='GTS') & (df['GBS Adoption Group'] =='MEA IOT')]
    df_W04D_GTS_MEA.to_csv('W04B_MySA_downloads_GTS_MEA.csv',index=False)   
    
    #generate: W04 GBS CIC EU RDM MySA Downloads
    df_W04D_GBS_CIC_EURDM = df[(df['Major Group']=='GBS') & (df['GBS Adoption Group'] =='CIC')]
    df_W04D_GBS_CIC_EURDM.to_csv('W04B_MySA_downloads_GBS_CIC_EURDM.csv',index=False)
    
    #generate: W04 GTS CIC EU RDM MySA Downloads
    df_W04D_GTS_CIC_EURDM = df[(df['Major Group']=='GTS') & (df['GBS Adoption Group'] =='CIC')]
    df_W04D_GTS_CIC_EURDM.to_csv('W04B_MySA_downloads_GTS_CIC_EURDM.csv',index=False)       
    
    #generate masked Email dataframe for CIC NONEU population
    df_W04D_CIC_NONEU = mask_NONEU(df)
    
    #generate: W04 GBS CIC NonEU MySA Downloads
    df_W04D_GBS_CIC_NONEU = df_W04D_CIC_NONEU[(df_W04D_CIC_NONEU['Major Group']=='GBS') & (df_W04D_CIC_NONEU['GBS Adoption Group'] =='CIC')]
    df_W04D_GBS_CIC_NONEU.to_csv('W04B_MySA_downloads_GBS_CIC_NONEU.csv',index=False)
    
    #generate: W04 GTS CIC NonEU MySA Downloads
    df_W04D_GTS_CIC_NONEU = df_W04D_CIC_NONEU[(df_W04D_CIC_NONEU['Major Group']=='GTS') & (df_W04D_CIC_NONEU['GBS Adoption Group'] =='CIC')]
    df_W04D_GTS_CIC_NONEU.to_csv('W04B_MySA_downloads_GTS_CIC_NONEU.csv',index=False)     
    
    
    #move the 3 temp files to backup folder with run date at file name
    rundate = time.strftime("%Y-%m-%d", time.gmtime())
    
    shutil.move('D:\\brrt\\data\\MySA_Apple_downloads.xlsx', 'D:\\brrt\\data\\backup\\MySA_Apple_downloads_'+rundate+'.xlsx') 
    shutil.move('D:\\brrt\\data\\MySA_Android_downloads.xlsx', 'D:\\brrt\\data\\backup\\MySA_Android_downloads_'+rundate+'.xlsx')
    shutil.move('D:\\brrt\\data\\W04D_inet_gd_flag.csv', 'D:\\brrt\\data\\backup\\W04D_inet_gd_flag_'+rundate+'.csv')
    
    print('W04D process completed!')

#Start of W04D execition script
try:
  if (    os.path.isfile("MySA_Android_downloads.xlsx") 
      and os.path.isfile("MySA_Apple_downloads.xlsx") 
      and os.path.isfile("W04D_inet_gd_flag.csv")
     ):
      print('W04D files are present at brrt/data folder, starting the process...')
      process_W04D()
  else:
      print('W04D will not run, the 3 files needed is not at brrt/data foder...')
                    
except:
    print('W04D: Some error happened once looking for the 3 files needed for W04D process to run.')
    