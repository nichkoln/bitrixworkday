#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Jan 23 16:26:29 2024

@author: nichkoln
"""

import requests
from datetime import datetime, timedelta
from google.oauth2 import service_account
from googleapiclient.discovery import build
import time
import traceback
import pandas as pd

SERVICE_ACCOUNT_FILE = '1'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SPREADSHEET_ID = 'у'
SHEET_NAME = 'у'
SHEET_NAME1 = 'у'


def getsheet(service,sheet_id, sheet_name):
    result = service.spreadsheets().values().get(spreadsheetId=sheet_id, range=sheet_name).execute()
    values = result.get('values', [])
    return values

def pivot(data):
    
    # Создаем DataFrame из массива
    df = pd.DataFrame(data[2:],columns=data[0])
    df_cut = df.iloc[:, 2::3]
    df_cut = pd.concat([df[['ID', 'NAME']], df_cut], axis = 1)

    column_dates = df_cut.columns.values.tolist()
    for i, col in enumerate(column_dates[2:]):
        the_date_formatted = datetime.strptime(col,"%Y-%m-%d").strftime("%Y-%m")
        column_dates[i+2] = the_date_formatted

    df_cut.columns = column_dates
    unpivot_df = pd.melt(df_cut, 
                     id_vars=['ID','NAME'], 
                     value_vars=list(df_cut.columns[1:]),
                     var_name='Дата', value_name='Время')
    unpivot_df = unpivot_df.rename(columns = {'NAME' : 'NAME'}).dropna(subset=['Время'], inplace=False)
    pivot_df = pd.pivot_table(unpivot_df, 
                          values = 'Время',
                          index = ['ID','NAME'],
                          columns = 'Дата',
                          aggfunc = 'count')
 

    return(pivot_df)

def updateggl(values,service,sheet_id, sheet_name):
    service.spreadsheets().values().update(spreadsheetId=sheet_id, range=sheet_name,
                                           body={'values': values}, valueInputOption='RAW').execute()

def make2d(df):
    values = [list(map(str, df.index.names)) + list(map(str, df.columns.tolist()))]

    for i in range(len(df)):
        row = list(map(str, df.index[i])) + list(map(str, df.iloc[i].tolist()))
        values.append(row)
    values.append(['','Дата Обновления:',datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
    return values

def main():
    creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    service = build('sheets', 'v4', credentials=creds)
    #while True:
    try:
        sheet=getsheet(service,SPREADSHEET_ID, SHEET_NAME)
        updateggl(make2d(pivot(sheet)),service,SPREADSHEET_ID, SHEET_NAME1)
        
    except Exception as e:
            print('5 sec')
            print(traceback.format_exc())
    # Пауза в одну минуту перед следующим обновлением
    time.sleep(5)
    print( datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

if __name__ == '__main__':
    main()
