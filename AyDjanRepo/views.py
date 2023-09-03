import warnings
import pyodbc
from datetime import datetime, timedelta
import pandas as pd
import numpy as np
from django.shortcuts import render
from django.http import HttpResponse
import random
import string


def generate_random_string(length):
    characters = string.ascii_letters + string.digits
    random_string = ''.join(random.choice(characters) for _ in range(length))
    return random_string

# random_string = generate_random_string(7)


def home(request):
    if request.method == 'GET':
        input_data = request.GET.get('dataInput', '')
        sample_data = request.GET.get('samples')
    else:
        input_data = ''

    if (sample_data == 'on'):
        sample_data = True
    else:
        sample_data = False

    fileName = startingPoint(input_data, sample_data)

    context = {
        'input_data': input_data,
        'sample_data': sample_data,
        'fileName': fileName + ".xlsx"
    }

    return render(request, 'home.html', context)


# ==========================================================
warnings.filterwarnings('ignore')

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

query_2 = ("SELECT T0.CardCode,T0.CardName,T0.DocNum, "
           "T0.DocDate,T1.ItemCode, T1.Dscription,T1.Quantity,T1.INMPrice,T0.Comments "
           "FROM (OINV T0 INNER JOIN INV1 T1 ON T0.DocEntry = T1.DocEntry) "
           "WHERE  T0.CardCode = (?) AND T0.CANCELED ='N'AND ISNULL(T0.Comments,0) NOT LIKE N'%عين%'")

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

query_4 = ("SELECT T0.CardCode,T0.CardName,T0.DocNum, "
           "T0.DocDate,T1.ItemCode, T1.Dscription,T1.Quantity,T1.INMPrice,T0.Comments "
           "FROM (OINV T0 INNER JOIN INV1 T1 ON T0.DocEntry = T1.DocEntry) "
           "WHERE  T0.CardCode = (?) AND T0.CANCELED ='N'AND ISNULL(T0.Comments,0) LIKE N'%عين%'")


def QueryData(query, cardcode):
    cnxn_str = ("Driver={/opt/microsoft/msodbcsql17/lib64/libmsodbcsql-17.10.so.4.1};"
                "Server=10.10.10.100;"
                "Database=LB;"
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
    return df1.to_excel(pathFinal, index=False)


def startingPoint(cardCode, checkValue):
    finalPath = ""
    # LB_AvaliableToReturn.xlsx
    randomFileName = generate_random_string(7)
    finalPath = finalPath + "./AyDjanRepo/static/" + randomFileName + ".xlsx"

    if checkValue == True:
        out_df1 = QueryData(query_3, cardCode)
        out_ATR = AvaliableToReturn(out_df1)
        out_df2 = QueryData(query_4, cardCode)
        update_ARwithATR(out_df2, out_ATR, cardCode, finalPath)

    else:
        out_df1 = QueryData(query_1, cardCode)
        out_ATR = AvaliableToReturn(out_df1)
        out_df2 = QueryData(query_2, cardCode)
        update_ARwithATR(out_df2, out_ATR, cardCode, finalPath)

    return randomFileName
