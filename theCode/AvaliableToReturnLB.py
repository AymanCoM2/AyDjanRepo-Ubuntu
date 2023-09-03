#!/usr/bin/env python
# coding: utf-8

# In[1]:


import warnings
warnings.filterwarnings('ignore')

# imports for SQL data part
import pyodbc
from datetime import datetime, timedelta
import pandas as pd
import numpy as np


# In[2]:


query_1 = ("With INV AS("
"SELECT " 
"V0.DocNum 'VDocNum', V0.DocDate 'VDocDate', V0.CardCode, MAX(V0.CardName) 'CardName',MAX(V0.Comments) 'Comments', "
"MAX(V1.ItemCode) 'ItemCode', MAX(V1.Dscription)'Dscription', MAX(V1.Quantity) 'VQty', MAX(T00.U_NAME) 'VSP', MAX(V0.NumAtCard) 'NumAtCard' "


"FROM OINV V0 "
"INNER JOIN INV1 V1 ON V0.DocEntry = V1.DocEntry " 
"LEFT JOIN OUSR T00 ON V0.USERSIGN = T00.INTERNAL_K "
"WHERE V0.CANCELED = 'N' AND ISNULL(V0.Comments,0) NOT LIKE N'%عين%' " 
"GROUP BY V0.CardCode,  V0.DocDate, V0.DocNum, V1.ItemCode), " 

"INR AS("
"SELECT "
"MAX(R0.U_Ref) 'RInvoiceID',R0.DocNum 'RDocNum', R0.DocDate 'RDocDate', R0.CardCode, MAX(R0.CardName) 'CardName',MAX(R0.Comments)'Comments', "
"MAX(R1.ItemCode)'ItemCode',MAX(R1.Dscription) 'Dscription', MAX(R1.Quantity) 'RQty',MAX(T00.U_NAME) 'RSP' "

"FROM ORIN R0 "
"INNER JOIN RIN1 R1 ON R0.DocEntry = R1.DocEntry "
"LEFT JOIN OUSR T00 ON R0.USERSIGN = T00.INTERNAL_K "
"WHERE R0.CANCELED = 'N' AND ISNULL(R0.Comments,0) NOT LIKE N'%عين%' "
"GROUP BY R0.CardCode,  R0.DocDate, R0.DocNum, R1.ItemCode) "

"SELECT R.RDocNum, R.RDocDate, R.CardCode, R.CardName, R.ItemCode,R.Dscription,R.RQty, "
"V.VDocNum, V.VDocDate,V.VQty, V.VSP, V.Comments "
         
"FROM INR R LEFT JOIN INV V "
"ON R.CardCode = V.CardCode AND R.ItemCode = V.ItemCode AND R.RDocDate >= V.VDocDate " 
"WHERE R.CardCode = (?) "
"ORDER BY R.RDocDate ASC,R.RDocNum ASC, R.ItemCode ASC,V.VDocDate DESC, V.VDocNum DESC ")


# In[3]:


query_2 = ("SELECT T0.CardCode,T0.CardName,T0.DocNum, "
"T0.DocDate,T1.ItemCode, T1.Dscription,T1.Quantity,T1.INMPrice,T0.Comments "
"FROM (OINV T0 INNER JOIN INV1 T1 ON T0.DocEntry = T1.DocEntry) "
"WHERE  T0.CardCode = (?) AND T0.CANCELED ='N'AND ISNULL(T0.Comments,0) NOT LIKE N'%عين%'")


# In[4]:


query_3 = ("With INV AS("
"SELECT " 
"V0.DocNum 'VDocNum', V0.DocDate 'VDocDate', V0.CardCode, MAX(V0.CardName) 'CardName',MAX(V0.Comments) 'Comments', "
"MAX(V1.ItemCode) 'ItemCode', MAX(V1.Dscription)'Dscription', MAX(V1.Quantity) 'VQty', MAX(T00.U_NAME) 'VSP', MAX(V0.NumAtCard) 'NumAtCard' "


"FROM OINV V0 "
"INNER JOIN INV1 V1 ON V0.DocEntry = V1.DocEntry " 
"LEFT JOIN OUSR T00 ON V0.USERSIGN = T00.INTERNAL_K "
"WHERE V0.CANCELED = 'N' AND ISNULL(V0.Comments,0) LIKE N'%عين%' " 
"GROUP BY V0.CardCode,  V0.DocDate, V0.DocNum, V1.ItemCode), " 

"INR AS("
"SELECT "
"MAX(R0.U_Ref) 'RInvoiceID',R0.DocNum 'RDocNum', R0.DocDate 'RDocDate', R0.CardCode, MAX(R0.CardName) 'CardName',MAX(R0.Comments)'Comments', "
"MAX(R1.ItemCode)'ItemCode',MAX(R1.Dscription) 'Dscription', MAX(R1.Quantity) 'RQty',MAX(T00.U_NAME) 'RSP' "

"FROM ORIN R0 "
"INNER JOIN RIN1 R1 ON R0.DocEntry = R1.DocEntry "
"LEFT JOIN OUSR T00 ON R0.USERSIGN = T00.INTERNAL_K "
"WHERE R0.CANCELED = 'N' AND ISNULL(R0.Comments,0) LIKE N'%عين%' "
"GROUP BY R0.CardCode,  R0.DocDate, R0.DocNum, R1.ItemCode) "

"SELECT R.RDocNum, R.RDocDate, R.CardCode, R.CardName, R.ItemCode,R.Dscription,R.RQty, "
"V.VDocNum, V.VDocDate,V.VQty, V.VSP, V.Comments "
         
"FROM INR R LEFT JOIN INV V "
"ON R.CardCode = V.CardCode AND R.ItemCode = V.ItemCode AND R.RDocDate >= V.VDocDate " 
"WHERE R.CardCode = (?) "
"ORDER BY R.RDocDate ASC,R.RDocNum ASC, R.ItemCode ASC,V.VDocDate DESC, V.VDocNum DESC ")


# In[5]:


query_4 = ("SELECT T0.CardCode,T0.CardName,T0.DocNum, "
"T0.DocDate,T1.ItemCode, T1.Dscription,T1.Quantity,T1.INMPrice,T0.Comments "
"FROM (OINV T0 INNER JOIN INV1 T1 ON T0.DocEntry = T1.DocEntry) "
"WHERE  T0.CardCode = (?) AND T0.CANCELED ='N'AND ISNULL(T0.Comments,0) LIKE N'%عين%'")


# In[6]:


def QueryData(query, cardcode):
    
    cnxn_str = ("Driver={SQL Server};"
            "Server=192.2.1.182;"
            "Database=LB;"
            "UID=sa;"
            "PWD=Aya@123;")
    cnxn = pyodbc.connect(cnxn_str)
    data = pd.read_sql(query, cnxn, params=[cardcode])
    
    return data    


# In[7]:


def AvaliableToReturn(df):
    
    df.sort_values(by=['RDocDate','RDocNum','ItemCode','VDocDate', 'VDocNum'], ascending=[True,True,True,False,False], inplace=True, ignore_index =True)
    
    df['ATR'] = df['VQty']
    df['Returned'] = 0.0
    
    
    for i in df['RDocNum'].unique():
        UniRID = i
    
        for j in df['ItemCode'][df['RDocNum'] == UniRID].unique():
            UniIT = j
        
            RQty = df['RQty'][(df['RDocNum'] == UniRID) & (df['ItemCode'] == UniIT)]    
            returnQty = max(np.append(RQty,0))
            updateATR = {}
        
            if(np.isnan(df['VDocNum'][(df['RDocNum'] == UniRID) & (df['ItemCode'] == UniIT)]).all()):
                continue
            
            for ind, row in df[(df['RDocNum'] == UniRID) & (df['ItemCode'] == UniIT)].iterrows():
        
                
                if (returnQty <= df['ATR'][ind]):

                    df['ATR'][ind] = df['ATR'][ind] - returnQty
                    df['Returned'][ind] = returnQty
                    returnQty = 0
                

                elif (returnQty > df['ATR'][ind]):
                
                    returnQty = returnQty - df['ATR'][ind]
                    df['Returned'][ind] = df['ATR'][ind]
                    df['ATR'][ind] = 0
                

            
                updateATR[df['VDocNum'][ind]] = df['ATR'][ind]
                    
            
            #update
            idxn = np.where((df['VDocNum'].isin(updateATR.keys())) & (df['ItemCode'] == UniIT))
            for i in df.loc[idxn].index:
                df['ATR'][i] = updateATR[df['VDocNum'][i]]
    
    sub_df = df[['VDocNum', 'VDocDate', 'ItemCode', 'Dscription', 'VQty', 'Returned', 'ATR', 'VSP', 'CardCode', 'CardName', 'Comments']]
    return sub_df


# In[8]:


def update_ARwithATR(df1,df2,cardcode):
    
    df1['ATR'] = df1['Quantity']
    
    for i in df1['DocNum'].unique():
        UniRID = i
    
        for j in df1['ItemCode'][df1['DocNum'] == UniRID].unique():
            UniIT = j

        
            if(df2['ATR'][(df2['VDocNum'] == UniRID) & (df2['ItemCode'] == UniIT)].empty):
                continue
            
            else:
                df1['ATR'][(df1['DocNum'] == UniRID) & (df1['ItemCode'] == UniIT)] = df2['ATR'][(df2['VDocNum'] == UniRID) & (df2['ItemCode'] == UniIT)].values[0]
    
    df1 = df1[df1['ATR'] != 0]
    df1.sort_values(by=['ItemCode','DocNum','DocDate'], ascending=[True,False,False], inplace=True, ignore_index =True)
    return df1.to_excel('LB_AvaliableToReturn.xlsx', index= False)


# In[9]:


import PySimpleGUI as sg

sg.theme('DarkAmber')   # Add a touch of color
# All the stuff inside your window.
layout = [ [sg.Text('Enter Customer Code'), sg.InputText()],
           [sg.Button('Ok'), sg.Button('Cancel'), sg.T("          "), sg.Checkbox('فواتير عينات', default=False, key="-IN-")] ]

# Create the Window
window = sg.Window('LB Avaliable To Return', layout, size=(300,100))
# Event Loop to process "events" and get the "values" of the inputs
while True:
    event, values = window.read()
    
    if event == sg.WIN_CLOSED or event == 'Cancel': # if user closes window or clicks cancel
        break
        
    if values["-IN-"] == True:
        out_df1 = QueryData(query_3, values[0])
        out_ATR = AvaliableToReturn(out_df1)
        out_df2 = QueryData(query_4, values[0])
        update_ARwithATR(out_df2,out_ATR,values[0])
        s = 'Avaliable To Return for Samples ' + values[0] + ' Exported Sucessfully'
        sg.popup(s) 
    else:    
        out_df1 = QueryData(query_1, values[0])
        out_ATR = AvaliableToReturn(out_df1)
        out_df2 = QueryData(query_2, values[0])
        update_ARwithATR(out_df2,out_ATR,values[0])
        s = 'Avaliable To Return for ' + values[0] + ' Exported Sucessfully'
        sg.popup(s) 

window.close()


# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:




