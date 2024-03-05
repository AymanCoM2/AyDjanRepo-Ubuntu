from datetime import datetime, timedelta
import warnings
import pyodbc
from datetime import datetime
import pandas as pd
import numpy as np
from django.shortcuts import render
from django.contrib.auth.forms import AuthenticationForm
from django.shortcuts import render, redirect
from django.contrib.auth import login
from .forms import CustomUserCreationForm  # Import your custom form
from django.contrib.auth.models import User
import datetime
import os
import warnings
warnings.filterwarnings('ignore')
from functools import partial

def clearStaticPath():
    current_directory = os.getcwd()
    docs_folder_path = os.path.join(current_directory, 'AyDjanRepo')
    static_folder_path = os.path.join(docs_folder_path, 'static')
    file_names = []
    for root, dirs, files in os.walk(static_folder_path):
        for file in files:
            if (file.startswith("_")):
                file_path = os.path.join(root, file)
                os.remove(file_path)
            else:
                file_path = os.path.join(root, file)
                file_names.append(file_path)

    file_names.sort()

    if (len(file_names) >= 15):
        for x in range(5):
            os.remove(file_names[x])


def generate_random_string(cardCode, checkValue):
    sampleTxt = ''
    if (checkValue):
        sampleTxt = 'With_Samples'
    else:
        sampleTxt = 'Without_samples'
    # Get the current date and time
    now = datetime.datetime.now()
    # Format the date and time as Y_m_d_h
    timestamp = now.strftime("%Y_%m_%d_%H_%M_%S")
    finalGeneratedName = cardCode+"_"+sampleTxt+"_"+timestamp
    # print(finalGeneratedName)
    return finalGeneratedName


def register(request):
    if request.method == 'POST':
        form = CustomUserCreationForm(request.POST)
        if form.is_valid():
            email = form.cleaned_data.get('email')
            # Add allowed domains here
            allowed_domains = ['lbaik.com',
                               'devo-p.com', 'aljouai.com', '2coom.com']

            if any(email.endswith(domain) for domain in allowed_domains):
                # Check if a user with the same email already exists
                if User.objects.filter(email=email).exists():
                    # Handle the case where the user is already registered
                    return render(request, 'registration/registration_error.html')
                else:
                    # Create a new User instance with the email as the username
                    username = email
                    password = form.cleaned_data.get('password1')
                    User.objects.create_user(
                        username, email=email, password=password)
                    # Redirect to the "home" page after successful registration
                    return redirect('home_page')
            else:
                pass
    else:
        form = CustomUserCreationForm()  # Use your custom form here without errors

    return render(request, 'registration/register.html', {'form': form})


def custom_login(request):
    if request.method == 'POST':
        form = AuthenticationForm(request, request.POST)
        if form.is_valid():
            login(request, form.get_user())
            # Redirect to a success page or another view
            return redirect('success_page_name')
    else:
        form = AuthenticationForm()
    return render(request, 'registration/login.html', {'form': form})


def main_page(request):
    return render(request, 'main_page.html')


def allcodeController(request):
    if request.method == 'GET':
        input_data = request.GET.get('dataInput', 'ALL')
        sample_data = request.GET.get('samples')
        db_name = request.GET.get('db_name')
    else:
        input_data = ''

    if (sample_data == 'on'):
        sample_data = True
    else:
        sample_data = False

    dbParam = str(db_name)

    # fileName = startingPointALL(sample_data)  # ! TODO
    fileName = startingPointALL(sample_data, dbParam)

    context = {
        'input_data': input_data,
        'sample_data': sample_data,
        'fileName': fileName + ".xlsx",
        'db_name': db_name
    }

    return render(request, 'allcode.html', context)


def cardcodeController(request):
    if request.method == 'GET':
        input_data = request.GET.get('dataInput', '')
        sample_data = request.GET.get('samples')
        db_name = request.GET.get('db_name')
    else:
        input_data = ''

    if (sample_data == 'on'):
        sample_data = True
    else:
        sample_data = False

    dbParam = str(db_name)
    fileName = startingPoint(input_data, sample_data, dbParam)
    # if db_name != None:
        
    # else :
    #     fileName = "test.xml"

    context = {
        'input_data': input_data,
        'sample_data': sample_data,
        'fileName': fileName + ".xlsx",
        'db_name': db_name
    }

    return render(request, 'cardcode.html', context)


def cardcodeControllerB(request):
    if request.method == 'GET':
        input_data = request.GET.get('dataInput', '')
        sample_data = request.GET.get('samples')
        db_name = request.GET.get('db_name')
    else:
        input_data = ''

    if (sample_data == 'on'):
        sample_data = True
    else:
        sample_data = False

    dbParam = str(db_name)

    fileName = startingPointB(input_data, sample_data, dbParam)

    context = {
        'input_data': input_data,
        'sample_data': sample_data,
        'fileName': fileName + ".xlsx",
        'db_name': db_name
    }

    return render(request, 'cardcodeB.html', context)



# ==========================================================
warnings.filterwarnings('ignore')

query_1 = ("""With INV AS(
           SELECT 
           V0.DocNum 'VDocNum', V0.DocDate 'VDocDate', V0.CardCode, MAX(V0.CardName) 'CardName',MAX(V0.Comments) 'Comments', 
           MAX(V1.ItemCode) 'ItemCode', MAX(V1.Dscription)'Dscription',
           CASE WHEN MAX(V1.ItemCode)='25090102067' and V0.DocNum='101194' and SUM(V1.Quantity)=500 THEN SUM(V1.Quantity)-99 
           ELSE SUM(V1.Quantity) END AS 'VQty',
           MAX(T00.U_NAME) 'VSP', MAX(V0.NumAtCard) 'NumAtCard' 
           FROM OINV V0 
           INNER JOIN INV1 V1 ON V0.DocEntry = V1.DocEntry 
           LEFT JOIN OUSR T00 ON V0.USERSIGN = T00.INTERNAL_K 
           WHERE V0.CANCELED = 'N' AND ISNULL(V0.Comments,0) NOT LIKE N'%عين%' 
           GROUP BY V0.CardCode,  V0.DocDate, V0.DocNum, V1.ItemCode), 

           INR AS(
           SELECT 
           MAX(R0.U_Ref) 'RInvoiceID',R0.DocNum 'RDocNum', R0.DocDate 'RDocDate', R0.CardCode, MAX(R0.CardName) 'CardName',MAX(R0.Comments)'Comments', 
           MAX(R1.ItemCode)'ItemCode',MAX(R1.Dscription) 'Dscription', SUM(R1.Quantity) 'RQty',MAX(T00.U_NAME) 'RSP' 

           FROM ORIN R0 
           INNER JOIN RIN1 R1 ON R0.DocEntry = R1.DocEntry 
           LEFT JOIN OUSR T00 ON R0.USERSIGN = T00.INTERNAL_K 
           WHERE R0.CANCELED = 'N' AND ISNULL(R0.Comments,0) NOT LIKE N'%عين%' 
           GROUP BY R0.CardCode,  R0.DocDate, R0.DocNum, R1.ItemCode) 

           SELECT R.RDocNum, R.RDocDate, R.CardCode, R.CardName, R.ItemCode,R.Dscription,R.RQty, 
           V.VDocNum, V.VDocDate,V.VQty, V.VSP, V.Comments 

           FROM INR R LEFT JOIN INV V 
           ON R.CardCode = V.CardCode AND R.ItemCode = V.ItemCode AND R.RDocDate >= V.VDocDate 
           WHERE R.CardCode = (?) 
           ORDER BY R.RDocDate ASC,R.RDocNum ASC, R.ItemCode ASC,V.VDocDate DESC, V.VDocNum DESC """)

query_2 = (""" SELECT T0.CardCode,MAX(T0.CardName)'CardName',T0.DocNum, 
           T0.DocDate,T1.ItemCode, MAX(T1.Dscription)'Dscription',
           CASE WHEN T1.ItemCode ='25090102067' and T0.DocNum='101194' and sum(T1.Quantity)=500 THEN SUM(T1.Quantity-99)
           ELSE sum(T1.Quantity) END AS 'Quantity',
           MAX(T1.INMPrice)'INMPrice',MAX(T0.Comments )'Comments' 
           FROM (OINV T0 INNER JOIN INV1 T1 ON T0.DocEntry = T1.DocEntry) 
           WHERE  T0.CardCode = (?) AND  T0.CANCELED ='N'AND ISNULL(T0.Comments,0) NOT LIKE N'%عين%' 
	   	   GROUP BY T0.CardCode,  T0.DocDate, T0.DocNum, T1.ItemCode """)

query_3 = ("""With INV AS(
           SELECT 
           V0.DocNum 'VDocNum', V0.DocDate 'VDocDate', V0.CardCode, MAX(V0.CardName) 'CardName',MAX(V0.Comments) 'Comments', 
           MAX(V1.ItemCode) 'ItemCode', MAX(V1.Dscription)'Dscription',
           CASE WHEN MAX(V1.ItemCode)='25090102067' and V0.DocNum='101194' and SUM(V1.Quantity)=500 THEN SUM(V1.Quantity)-99 
           ELSE SUM(V1.Quantity) END AS 'VQty',           
           MAX(T00.U_NAME) 'VSP', MAX(V0.NumAtCard) 'NumAtCard' 


           FROM OINV V0 
           INNER JOIN INV1 V1 ON V0.DocEntry = V1.DocEntry 
           LEFT JOIN OUSR T00 ON V0.USERSIGN = T00.INTERNAL_K 
           WHERE V0.CANCELED = 'N' AND ISNULL(V0.Comments,0) LIKE N'%عين%' 
           GROUP BY V0.CardCode,  V0.DocDate, V0.DocNum, V1.ItemCode), 

           INR AS(
           SELECT 
           MAX(R0.U_Ref) 'RInvoiceID',R0.DocNum 'RDocNum', R0.DocDate 'RDocDate', R0.CardCode, MAX(R0.CardName) 'CardName',MAX(R0.Comments)'Comments', 
           MAX(R1.ItemCode)'ItemCode',MAX(R1.Dscription) 'Dscription', SUM(R1.Quantity) 'RQty',MAX(T00.U_NAME) 'RSP' 

           FROM ORIN R0 
           INNER JOIN RIN1 R1 ON R0.DocEntry = R1.DocEntry 
           LEFT JOIN OUSR T00 ON R0.USERSIGN = T00.INTERNAL_K 
           WHERE R0.CANCELED = 'N' AND ISNULL(R0.Comments,0) LIKE N'%عين%' 
           GROUP BY R0.CardCode,  R0.DocDate, R0.DocNum, R1.ItemCode) 

           SELECT R.RDocNum, R.RDocDate, R.CardCode, R.CardName, R.ItemCode,R.Dscription,R.RQty, 
           V.VDocNum, V.VDocDate,V.VQty, V.VSP, V.Comments 

           FROM INR R LEFT JOIN INV V 
           ON R.CardCode = V.CardCode AND R.ItemCode = V.ItemCode AND R.RDocDate >= V.VDocDate 
           WHERE R.CardCode = (?) 
           ORDER BY R.RDocDate ASC,R.RDocNum ASC, R.ItemCode ASC,V.VDocDate DESC, V.VDocNum DESC """)

query_4 = ("""SELECT T0.CardCode,MAX(T0.CardName)'CardName',T0.DocNum, 
           T0.DocDate,T1.ItemCode, MAX(T1.Dscription)'Dscription',
           CASE WHEN T1.ItemCode ='25090102067' and T0.DocNum='101194' and sum(T1.Quantity)=500 THEN SUM(T1.Quantity-99)
           ELSE sum(T1.Quantity) END AS 'Quantity',
           MAX(T1.INMPrice)'INMPrice',MAX(T0.Comments )'Comments' 
           FROM (OINV T0 INNER JOIN INV1 T1 ON T0.DocEntry = T1.DocEntry) 
           WHERE  T0.CardCode = (?) AND  T0.CANCELED ='N'AND ISNULL(T0.Comments,0) LIKE N'%عين%' 
	   	   GROUP BY T0.CardCode,  T0.DocDate, T0.DocNum, T1.ItemCode""")

# This is the Line TO CHANGE


def QueryData(query, cardcode, dbParameter):

    if (dbParameter is None):
        dbParameter = "TM"
    cnxn_str = ""

    if (dbParameter == "TM"):
        # cnxn_str = ("Driver={SQL Server};"
        cnxn_str = ("Driver={/opt/microsoft/msodbcsql17/lib64/libmsodbcsql-17.10.so.4.1};"
                    "Server=10.10.10.100;"
                    # "Server=jdry1.ifrserp.net,445;"
                    "Database=TM;"
                    "UID=ayman;"
                    "PWD=admin@1234;")
    elif (dbParameter == "LB"):
        # cnxn_str = ("Driver={SQL Server};"
        cnxn_str = ("Driver={/opt/microsoft/msodbcsql17/lib64/libmsodbcsql-17.10.so.4.1};"
                    "Server=10.10.10.100;"
                    # "Server=jdry1.ifrserp.net,445;"
                    "Database=LB;"
                    "UID=ayman;"
                    "PWD=admin@1234;")
    else:
        # cnxn_str = ("Driver={SQL Server};"
        cnxn_str = ("Driver={/opt/microsoft/msodbcsql17/lib64/libmsodbcsql-17.10.so.4.1};"
                    "Server=10.10.10.100;"
                    # "Server=jdry1.ifrserp.net,445;"
                    "Database=TM;"
                    "UID=ayman;"
                    "PWD=admin@1234;")

    cnxn = pyodbc.connect(cnxn_str)

    data = pd.read_sql(query, cnxn, params=[cardcode])

    return data


def AvaliableToReturn(df):
    df.sort_values(by=['RDocDate', 'RDocNum', 'ItemCode', 'VDocDate', 'VDocNum'], ascending=[
                   True, True, True, False, False], inplace=True, ignore_index=True)

    df['ATR'] = df['VQty']
    df['Returned'] = 0.0

    for i in df['RDocNum'].unique():
        UniRID = i

        for j in df['ItemCode'][df['RDocNum'] == UniRID].unique():
            UniIT = j

            RQty = df['RQty'][(df['RDocNum'] == UniRID) &
                              (df['ItemCode'] == UniIT)]
            returnQty = max(np.append(RQty, 0))
            updateATR = {}
            if (df['VDocNum'][(df['RDocNum'] == UniRID) & (df['ItemCode'] == UniIT)].isnull().all()):
                # if (np.isnan(df['VDocNum'][(df['RDocNum'] == UniRID) & (df['ItemCode'] == UniIT)]).all()):
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

            # update
            idxn = np.where((df['VDocNum'].isin(updateATR.keys())) & (
                df['ItemCode'] == UniIT))
            for i in df.loc[idxn].index:
                df['ATR'][i] = updateATR[df['VDocNum'][i]]

    sub_df = df[['VDocNum', 'VDocDate', 'ItemCode', 'Dscription', 'VQty',
                 'Returned', 'ATR', 'VSP', 'CardCode', 'CardName', 'Comments']]
    return sub_df


def update_ARwithATR(df1, df2, cardcode, pathFinal):
    df1['ATR'] = df1['Quantity']
    for i in df1['DocNum'].unique():
        UniRID = i
        for j in df1['ItemCode'][df1['DocNum'] == UniRID].unique():
            UniIT = j

            if (df2['ATR'][(df2['VDocNum'] == UniRID) & (df2['ItemCode'] == UniIT)].empty):
                continue

            else:
                df1['ATR'][(df1['DocNum'] == UniRID) & (df1['ItemCode'] == UniIT)] = df2['ATR'][(
                    df2['VDocNum'] == UniRID) & (df2['ItemCode'] == UniIT)].values[0]

    df1 = df1[df1['ATR'] != 0]
    df1.sort_values(by=['ItemCode', 'DocNum', 'DocDate'], ascending=[
                    True, False, False], inplace=True, ignore_index=True)
    clearStaticPath()
    return df1.to_excel(pathFinal, index=False)


def startingPoint(cardCode, checkValue, dbParameter):
    finalPath = ""
    # LB_AvaliableToReturn.xlsx
    randomFileName = generate_random_string(cardCode, checkValue)
    finalPath = finalPath + "./AyDjanRepo/static/" + randomFileName + ".xlsx"

    if checkValue == True:
        out_df1 = QueryData(query_3, cardCode, dbParameter)
        out_ATR = AvaliableToReturn(out_df1)
        out_df2 = QueryData(query_4, cardCode, dbParameter)
        update_ARwithATR(out_df2, out_ATR, cardCode, finalPath)

    else:
        out_df1 = QueryData(query_1, cardCode, dbParameter)
        out_ATR = AvaliableToReturn(out_df1)
        out_df2 = QueryData(query_2, cardCode, dbParameter)
        update_ARwithATR(out_df2, out_ATR, cardCode, finalPath)

    return randomFileName


# ==========================================================
warnings.filterwarnings('ignore')

query_1A = (""" With INV AS(
            SELECT 
            V0.DocNum 'VDocNum', V0.DocDate 'VDocDate', V0.CardCode, MAX(V0.CardName) 'CardName',MAX(V0.Comments) 'Comments', 
            MAX(V1.ItemCode) 'ItemCode', MAX(V1.Dscription)'Dscription',
            CASE WHEN MAX(V1.ItemCode)='25090102067' and V0.DocNum='101194' and SUM(V1.Quantity)=500 THEN SUM(V1.Quantity)-99 
            ELSE SUM(V1.Quantity) END AS 'VQty',
            MAX(T00.U_NAME) 'VSP', MAX(V0.NumAtCard) 'NumAtCard' 


            FROM OINV V0 
            INNER JOIN INV1 V1 ON V0.DocEntry = V1.DocEntry 
            LEFT JOIN OUSR T00 ON V0.USERSIGN = T00.INTERNAL_K 
            WHERE V0.CANCELED = 'N' AND ISNULL(V0.Comments,0) NOT LIKE N'%عين%' 
            GROUP BY V0.CardCode,  V0.DocDate, V0.DocNum, V1.ItemCode), 

            INR AS(
            SELECT 
            MAX(R0.U_Ref) 'RInvoiceID',R0.DocNum 'RDocNum', R0.DocDate 'RDocDate', R0.CardCode, MAX(R0.CardName) 'CardName',MAX(R0.Comments)'Comments', 
            MAX(R1.ItemCode)'ItemCode',MAX(R1.Dscription) 'Dscription', SUM(R1.Quantity) 'RQty',MAX(T00.U_NAME) 'RSP' 

            FROM ORIN R0 
            INNER JOIN RIN1 R1 ON R0.DocEntry = R1.DocEntry 
            LEFT JOIN OUSR T00 ON R0.USERSIGN = T00.INTERNAL_K 
            WHERE R0.CANCELED = 'N' AND ISNULL(R0.Comments,0) NOT LIKE N'%عين%' 
            GROUP BY R0.CardCode,  R0.DocDate, R0.DocNum, R1.ItemCode) 

            SELECT R.RDocNum, R.RDocDate, R.CardCode, R.CardName, R.ItemCode,R.Dscription,R.RQty, 
            V.VDocNum, V.VDocDate,V.VQty, V.VSP, V.Comments 

            FROM INR R LEFT JOIN INV V 
            ON R.CardCode = V.CardCode AND R.ItemCode = V.ItemCode AND R.RDocDate >= V.VDocDate 
            --WHERE R.CardCode = (?) 
            ORDER BY R.RDocDate ASC,R.RDocNum ASC, R.ItemCode ASC,V.VDocDate DESC, V.VDocNum DESC """ )


# In[3]:


query_2A = (""" SELECT T0.CardCode,MAX(T0.CardName)'CardName',T0.DocNum, 
           T0.DocDate,T1.ItemCode, MAX(T1.Dscription)'Dscription',
           CASE WHEN T1.ItemCode ='25090102067' and T0.DocNum='101194' and sum(T1.Quantity)=500 THEN SUM(T1.Quantity-99)
           ELSE sum(T1.Quantity) END AS 'Quantity',
           MAX(T1.INMPrice)'INMPrice',MAX(T0.Comments )'Comments' 
           FROM (OINV T0 INNER JOIN INV1 T1 ON T0.DocEntry = T1.DocEntry) 
           WHERE  T0.CardCode = (?) AND  T0.CANCELED ='N' AND ISNULL(T0.Comments,0) NOT LIKE N'%عين%' 
	   	   GROUP BY T0.CardCode,  T0.DocDate, T0.DocNum, T1.ItemCode""")


# In[4]:


query_3A = (""" With INV AS(
            SELECT 
            V0.DocNum 'VDocNum', V0.DocDate 'VDocDate', V0.CardCode, MAX(V0.CardName) 'CardName',MAX(V0.Comments) 'Comments', 
            MAX(V1.ItemCode) 'ItemCode', MAX(V1.Dscription)'Dscription',
            CASE WHEN MAX(V1.ItemCode)='25090102067' and V0.DocNum='101194' and SUM(V1.Quantity)=500 THEN SUM(V1.Quantity)-99 
            ELSE SUM(V1.Quantity) END AS 'VQty',
            MAX(T00.U_NAME) 'VSP', MAX(V0.NumAtCard) 'NumAtCard' 


            FROM OINV V0 
            INNER JOIN INV1 V1 ON V0.DocEntry = V1.DocEntry 
            LEFT JOIN OUSR T00 ON V0.USERSIGN = T00.INTERNAL_K 
            WHERE V0.CANCELED = 'N' AND ISNULL(V0.Comments,0) LIKE N'%عين%' 
            GROUP BY V0.CardCode,  V0.DocDate, V0.DocNum, V1.ItemCode), 

            INR AS(
            SELECT 
            MAX(R0.U_Ref) 'RInvoiceID',R0.DocNum 'RDocNum', R0.DocDate 'RDocDate', R0.CardCode, MAX(R0.CardName) 'CardName',MAX(R0.Comments)'Comments', 
            MAX(R1.ItemCode)'ItemCode',MAX(R1.Dscription) 'Dscription', SUM(R1.Quantity) 'RQty',MAX(T00.U_NAME) 'RSP' 

            FROM ORIN R0 
            INNER JOIN RIN1 R1 ON R0.DocEntry = R1.DocEntry 
            LEFT JOIN OUSR T00 ON R0.USERSIGN = T00.INTERNAL_K 
            WHERE R0.CANCELED = 'N' AND ISNULL(R0.Comments,0) LIKE N'%عين%' 
            GROUP BY R0.CardCode,  R0.DocDate, R0.DocNum, R1.ItemCode) 

            SELECT R.RDocNum, R.RDocDate, R.CardCode, R.CardName, R.ItemCode,R.Dscription,R.RQty, 
            V.VDocNum, V.VDocDate,V.VQty, V.VSP, V.Comments 

            FROM INR R LEFT JOIN INV V 
            ON R.CardCode = V.CardCode AND R.ItemCode = V.ItemCode AND R.RDocDate >= V.VDocDate 
            --WHERE R.CardCode = (?) 
            ORDER BY R.RDocDate ASC,R.RDocNum ASC, R.ItemCode ASC,V.VDocDate DESC, V.VDocNum DESC """)




query_4A = (""" SELECT T0.CardCode,MAX(T0.CardName)'CardName',T0.DocNum, 
           T0.DocDate,T1.ItemCode, MAX(T1.Dscription)'Dscription',
           CASE WHEN T1.ItemCode ='25090102067' and T0.DocNum='101194' and sum(T1.Quantity)=500 THEN SUM(T1.Quantity-99)
           ELSE sum(T1.Quantity) END AS 'Quantity',
           MAX(T1.INMPrice)'INMPrice',MAX(T0.Comments )'Comments' 
           FROM (OINV T0 INNER JOIN INV1 T1 ON T0.DocEntry = T1.DocEntry) 
           WHERE  T0.CardCode = (?) AND  T0.CANCELED ='N' AND ISNULL(T0.Comments,0) LIKE N'%عين%' 
	   	   GROUP BY T0.CardCode,  T0.DocDate, T0.DocNum, T1.ItemCode """)


def QueryDataALL(query , dbParameter):

    if (dbParameter is None):
        dbParameter = "TM"
    cnxn_str = ""

    if (dbParameter == "TM"):
        # cnxn_str = ("Driver={SQL Server};"
        cnxn_str = ("Driver={/opt/microsoft/msodbcsql17/lib64/libmsodbcsql-17.10.so.4.1};"
                    "Server=10.10.10.100;"
                    # "Server=jdry1.ifrserp.net,445;"
                    "Database=TM;"
                    "UID=ayman;"
                    "PWD=admin@1234;")
    elif (dbParameter == "LB"):
        # cnxn_str = ("Driver={SQL Server};"
        cnxn_str = ("Driver={/opt/microsoft/msodbcsql17/lib64/libmsodbcsql-17.10.so.4.1};"
                    "Server=10.10.10.100;"
                    # "Server=jdry1.ifrserp.net,445;"
                    "Database=LB;"
                    "UID=ayman;"
                    "PWD=admin@1234;")
    else:
        # cnxn_str = ("Driver={SQL Server};"
        cnxn_str = ("Driver={/opt/microsoft/msodbcsql17/lib64/libmsodbcsql-17.10.so.4.1};"
                    "Server=10.10.10.100;"
                    # "Server=jdry1.ifrserp.net,445;"
                    "Database=TM;"
                    "UID=ayman;"
                    "PWD=admin@1234;")

    cnxn = pyodbc.connect(cnxn_str)
    data = pd.read_sql(query, cnxn, params=[])

    return data


def AvaliableToReturnALL(df):

    df.sort_values(by=['RDocDate', 'RDocNum', 'ItemCode', 'VDocDate', 'VDocNum'], ascending=[
                   True, True, True, False, False], inplace=True, ignore_index=True)

    df['ATR'] = df['VQty']
    df['Returned'] = 0.0

    for i in df['RDocNum'].unique():
        UniRID = i

        for j in df['ItemCode'][df['RDocNum'] == UniRID].unique():
            UniIT = j

            RQty = df['RQty'][(df['RDocNum'] == UniRID) &
                              (df['ItemCode'] == UniIT)]
            returnQty = max(np.append(RQty, 0))
            updateATR = {}

            if (np.isnan(df['VDocNum'][(df['RDocNum'] == UniRID) & (df['ItemCode'] == UniIT)]).all()):
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

            # update
            idxn = np.where((df['VDocNum'].isin(updateATR.keys())) & (
                df['ItemCode'] == UniIT))
            for i in df.loc[idxn].index:
                df['ATR'][i] = updateATR[df['VDocNum'][i]]

    sub_df = df[['VDocNum', 'VDocDate', 'ItemCode', 'Dscription', 'VQty',
                 'Returned', 'ATR', 'VSP', 'CardCode', 'CardName', 'Comments']]
    return sub_df


def update_ARwithATRALL(df1, df2, pathFinal):
    df1['ATR'] = df1['Quantity']
    for i in df1['DocNum'].unique():
        UniRID = i
        for j in df1['ItemCode'][df1['DocNum'] == UniRID].unique():
            UniIT = j

            if (df2['ATR'][(df2['VDocNum'] == UniRID) & (df2['ItemCode'] == UniIT)].empty):
                continue

            else:
                df1['ATR'][(df1['DocNum'] == UniRID) & (df1['ItemCode'] == UniIT)] = df2['ATR'][(
                    df2['VDocNum'] == UniRID) & (df2['ItemCode'] == UniIT)].values[0]

    df1 = df1[df1['ATR'] != 0]
    df1.sort_values(by=['ItemCode', 'DocNum', 'DocDate'], ascending=[
                    True, False, False], inplace=True, ignore_index=True)
    return df1.to_excel(pathFinal, index=False)


def startingPointALL(checkValue , dbParamerter):
    finalPath = ""
    # LB_AvaliableToReturn.xlsx
    randomFileName = generate_random_string("ALL", 7)
    finalPath = finalPath + "./AyDjanRepo/static/" + randomFileName + ".xlsx"

    if checkValue == True:
        out_df1 = QueryDataALL(query_3A,dbParamerter)
        out_ATR = AvaliableToReturnALL(out_df1)
        out_df2 = QueryDataALL(query_4A,dbParamerter)
        update_ARwithATRALL(out_df2, out_ATR, finalPath)

    else:
        out_df1 = QueryDataALL(query_1A,dbParamerter)
        out_ATR = AvaliableToReturnALL(out_df1)
        out_df2 = QueryDataALL(query_2A,dbParamerter)
        update_ARwithATRALL(out_df2, out_ATR, finalPath)

    return randomFileName


query_1B = ("""With INV AS(
SELECT  
V0.DocNum 'VDocNum', V0.DocDate 'VDocDate',D1.CardFName 'Foreign Name' ,MAX(V0.CardCode)'Card Code', MAX(V0.CardName) 'CardName',MAX(V0.Comments) 'Comments', 
MAX(V1.ItemCode) 'ItemCode', MAX(V1.Dscription)'Dscription',
CASE WHEN MAX(V1.ItemCode)='25090102067' and V0.DocNum='101194' and SUM(V1.Quantity)=500 THEN SUM(V1.Quantity)-99
else SUM(V1.Quantity) END AS 'VQty',
MAX(T00.U_NAME) 'VSP', MAX(V0.NumAtCard) 'NumAtCard' 


FROM OINV V0 
INNER JOIN INV1 V1 ON V0.DocEntry = V1.DocEntry  
LEFT JOIN OUSR T00 ON V0.USERSIGN = T00.INTERNAL_K 
LEFT JOIN OCRD D1 ON V0.CardCode = D1.CardCode
WHERE V0.CANCELED = 'N' AND ISNULL(V0.Comments,0) NOT LIKE N'%عين%'  
GROUP BY D1.CardFName,  V0.DocDate, V0.DocNum, V1.ItemCode),  

INR AS(
SELECT 
MAX(R0.U_Ref) 'RInvoiceID',R0.DocNum 'RDocNum', R0.DocDate 'RDocDate',D1.CardFName 'Foreign Name', MAX(R0.CardCode)'CardCode', MAX(R0.CardName) 'CardName',MAX(R0.Comments)'Comments', 
MAX(R1.ItemCode)'ItemCode',MAX(R1.Dscription) 'Dscription', SUM(R1.Quantity) 'RQty',MAX(T00.U_NAME) 'RSP' 

FROM ORIN R0 
INNER JOIN RIN1 R1 ON R0.DocEntry = R1.DocEntry 
LEFT JOIN OUSR T00 ON R0.USERSIGN = T00.INTERNAL_K 
LEFT JOIN OCRD D1 ON R0.CardCode = D1.CardCode
WHERE R0.CANCELED = 'N' AND ISNULL(R0.Comments,0) NOT LIKE N'%عين%' 
GROUP BY D1.CardFName,  R0.DocDate, R0.DocNum, R1.ItemCode) 

SELECT R.RDocNum, R.RDocDate, R.[Foreign Name] , R.CardCode, R.CardName, R.ItemCode,R.Dscription,R.RQty, 
V.VDocNum, V.VDocDate,V.VQty, V.VSP, V.Comments 
         
FROM INR R LEFT JOIN INV V 
ON R.[Foreign Name] = V.[Foreign Name] AND R.ItemCode = V.ItemCode AND R.RDocDate >= V.VDocDate  
WHERE R.[Foreign Name] =(?) 
ORDER BY R.RDocDate ASC,R.RDocNum ASC, R.ItemCode ASC,V.VDocDate DESC, V.VDocNum DESC """)


query_2B = ("""SELECT MAX(D1.CardFName)'Foreign Name',T0.CardCode,MAX(T0.CardName)'CardName',T0.DocNum, 
            T0.DocDate,T1.ItemCode, MAX(T1.Dscription)'Dscription',
	    CASE WHEN T1.ItemCode ='25090102067' and T0.DocNum='101194' and sum(T1.Quantity)=500 THEN SUM(T1.Quantity-99)
           ELSE sum(T1.Quantity) END AS 'Quantity',
           MAX(T1.INMPrice)'INMPrice',MAX(T0.Comments )'Comments' 
            FROM (OINV T0 INNER JOIN INV1 T1 ON T0.DocEntry = T1.DocEntry LEFT JOIN OCRD D1 ON T0.CardCode = D1.CardCode) 
            WHERE  D1.CardFName = (?) AND T0.CANCELED ='N'AND ISNULL(T0.Comments,0) NOT LIKE N'%عين%'
			GROUP BY T0.CardCode,  T0.DocDate, T0.DocNum, T1.ItemCode """)

query_3B = (""" With INV AS(
            SELECT 
            V0.DocNum 'VDocNum', V0.DocDate 'VDocDate', D1.CardFName 'Foreign Name' ,MAX(V0.CardCode)'Card Code', MAX(V0.CardName) 'CardName',MAX(V0.Comments) 'Comments', 
            MAX(V1.ItemCode) 'ItemCode', MAX(V1.Dscription)'Dscription',
            CASE WHEN MAX(V1.ItemCode)='25090102067' and V0.DocNum='101194' and SUM(V1.Quantity)=500 THEN SUM(V1.Quantity)-99
            else SUM(V1.Quantity) END AS 'VQty',
            MAX(T00.U_NAME) 'VSP', MAX(V0.NumAtCard) 'NumAtCard' 


            FROM OINV V0 
            INNER JOIN INV1 V1 ON V0.DocEntry = V1.DocEntry 
            LEFT JOIN OUSR T00 ON V0.USERSIGN = T00.INTERNAL_K 
            LEFT JOIN OCRD D1 ON V0.CardCode = D1.CardCode
            WHERE V0.CANCELED = 'N' AND ISNULL(V0.Comments,0) LIKE N'%عين%' 
            GROUP BY D1.CardFName,  V0.DocDate, V0.DocNum, V1.ItemCode), 

            INR AS(
            SELECT 
            MAX(R0.U_Ref) 'RInvoiceID',R0.DocNum 'RDocNum', D1.CardFName 'Foreign Name',R0.DocDate 'RDocDate', MAX(R0.CardCode)'CardCode', MAX(R0.CardName) 'CardName',MAX(R0.Comments)'Comments', 
            MAX(R1.ItemCode)'ItemCode',MAX(R1.Dscription) 'Dscription', SUM(R1.Quantity) 'RQty',MAX(T00.U_NAME) 'RSP' 

            FROM ORIN R0 
            INNER JOIN RIN1 R1 ON R0.DocEntry = R1.DocEntry 
            LEFT JOIN OUSR T00 ON R0.USERSIGN = T00.INTERNAL_K 
            LEFT JOIN OCRD D1 ON R0.CardCode = D1.CardCode
            WHERE R0.CANCELED = 'N' AND ISNULL(R0.Comments,0) LIKE N'%عين%' 
            GROUP BY D1.CardFName,  R0.DocDate, R0.DocNum, R1.ItemCode) 

            SELECT R.RDocNum, R.RDocDate, R.CardCode, R.CardName,R.[Foreign Name], R.ItemCode,R.Dscription,R.RQty, 
            V.VDocNum, V.VDocDate,V.VQty, V.VSP, V.Comments 

            FROM INR R LEFT JOIN INV V 
            ON R.[Foreign Name] = V.[Foreign Name] AND R.ItemCode = V.ItemCode AND R.RDocDate >= V.VDocDate 
            WHERE R.[Foreign Name] = (?) 
            ORDER BY R.RDocDate ASC,R.RDocNum ASC, R.ItemCode ASC,V.VDocDate DESC, V.VDocNum DESC """ )

query_4B = (""" SELECT MAX(D1.CardFName)'Foreign Name',T0.CardCode,MAX(T0.CardName)'CardName',T0.DocNum, 
            T0.DocDate,T1.ItemCode, MAX(T1.Dscription)'Dscription',
	    CASE WHEN T1.ItemCode ='25090102067' and T0.DocNum='101194' and sum(T1.Quantity)=500 THEN SUM(T1.Quantity-99)
           ELSE sum(T1.Quantity) END AS 'Quantity',
           MAX(T1.INMPrice)'INMPrice',MAX(T0.Comments )'Comments' 
            FROM (OINV T0 INNER JOIN INV1 T1 ON T0.DocEntry = T1.DocEntry LEFT JOIN OCRD D1 ON T0.CardCode = D1.CardCode) 
            WHERE  D1.CardFName = (?) AND T0.CANCELED ='N'AND ISNULL(T0.Comments,0) LIKE N'%عين%'
			GROUP BY T0.CardCode,  T0.DocDate, T0.DocNum, T1.ItemCode""")


def QueryDataB(query,Foreignname,dbParameter):
    if (dbParameter is None):
        dbParameter = "TM"
    cnxn_str = ""

    if (dbParameter == "TM"):
        # cnxn_str = ("Driver={SQL Server};"
        cnxn_str = ("Driver={/opt/microsoft/msodbcsql17/lib64/libmsodbcsql-17.10.so.4.1};"
                    "Server=10.10.10.100;"
                    # "Server=jdry1.ifrserp.net,445;"
                    "Database=TM;"
                    "UID=ayman;"
                    "PWD=admin@1234;")
    elif (dbParameter == "LB"):
        # cnxn_str = ("Driver={SQL Server};"
        cnxn_str = ("Driver={/opt/microsoft/msodbcsql17/lib64/libmsodbcsql-17.10.so.4.1};"
                    "Server=10.10.10.100;"
                    # "Server=jdry1.ifrserp.net,445;"
                    "Database=LB;"
                    "UID=ayman;"
                    "PWD=admin@1234;")
    else:
        # cnxn_str = ("Driver={SQL Server};"
        cnxn_str = ("Driver={/opt/microsoft/msodbcsql17/lib64/libmsodbcsql-17.10.so.4.1};"
                    "Server=10.10.10.100;"
                    # "Server=jdry1.ifrserp.net,445;"
                    "Database=TM;"
                    "UID=ayman;"
                    "PWD=admin@1234;")

    cnxn = pyodbc.connect(cnxn_str)
    data = pd.read_sql(query, cnxn, params=[Foreignname])

    return data


def AvaliableToReturnB(df):
    df.sort_values(by=['RDocDate', 'RDocNum', 'ItemCode', 'VDocDate', 'VDocNum'], ascending=[
                   True, True, True, False, False], inplace=True, ignore_index=True)

    df['ATR'] = df['VQty']
    df['Returned'] = 0.0

    for i in df['RDocNum'].unique():
        UniRID = i

        for j in df['ItemCode'][df['RDocNum'] == UniRID].unique():
            UniIT = j

            RQty = df['RQty'][(df['RDocNum'] == UniRID) &
                              (df['ItemCode'] == UniIT)]
            returnQty = max(np.append(RQty, 0))
            updateATR = {}

            if (np.isnan(df['VDocNum'][(df['RDocNum'] == UniRID) & (df['ItemCode'] == UniIT)]).all()):
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

            # update
            idxn = np.where((df['VDocNum'].isin(updateATR.keys())) & (
                df['ItemCode'] == UniIT))
            for i in df.loc[idxn].index:
                df['ATR'][i] = updateATR[df['VDocNum'][i]]

    sub_df = df[['VDocNum', 'VDocDate', 'ItemCode', 'Dscription', 'VQty',
                 'Returned', 'ATR', 'VSP', 'CardCode', 'CardName', 'Comments']]
    return sub_df


def update_ARwithATRB(df1, df2, pathFinal):
    df1['ATR'] = df1['Quantity']
    for i in df1['DocNum'].unique():
        UniRID = i
        for j in df1['ItemCode'][df1['DocNum'] == UniRID].unique():
            UniIT = j

            if (df2['ATR'][(df2['VDocNum'] == UniRID) & (df2['ItemCode'] == UniIT)].empty):
                continue

            else:
                df1['ATR'][(df1['DocNum'] == UniRID) & (df1['ItemCode'] == UniIT)] = df2['ATR'][(
                    df2['VDocNum'] == UniRID) & (df2['ItemCode'] == UniIT)].values[0]

    df1 = df1[df1['ATR'] != 0]
    df1.sort_values(by=['ItemCode', 'DocNum', 'DocDate'], ascending=[
                    True, False, False], inplace=True, ignore_index=True)
    return df1.to_excel(pathFinal, index=False)


def startingPointB(input_data,checkValue,dbParameter):
    finalPath = ""
    # LB_AvaliableToReturn.xlsx
    randomFileName = generate_random_string("ALL", 7)
    finalPath = finalPath + "./AyDjanRepo/static/" + randomFileName + ".xlsx"

    if checkValue == True:
        out_df1 = QueryDataB(query_3B,input_data,dbParameter)
        out_ATR = AvaliableToReturnB(out_df1)
        out_df2 = QueryDataB(query_4B,input_data,dbParameter)
        update_ARwithATRB(out_df2, out_ATR, finalPath)

    else:
        out_df1 = QueryDataB(query_1B,input_data,dbParameter)
        out_ATR = AvaliableToReturnB(out_df1)
        out_df2 = QueryDataB(query_2B,input_data,dbParameter)
        update_ARwithATRB(out_df2, out_ATR, finalPath)

    return randomFileName


