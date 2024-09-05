#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Dec 26 19:05:06 2023

@author: nichkoln
"""

import requests
from datetime import datetime, timedelta
from google.oauth2 import service_account
from googleapiclient.discovery import build
import time
import traceback

SERVICE_ACCOUNT_FILE = '.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SPREADSHEET_ID = ''
SHEET_NAME = ''
SHEET_NAME1 = ''

def getworkday(id):
    url = ''
    params = {'USER_ID': id}
    resp = requests.get(url=url, params=params)
    response = resp.json()
    return response['result']

def pushtoya(id,name,args, service, sheet_id, sheet_name):
    today = datetime.now().strftime("%Y-%m-%d")
    status=args['STATUS']
    Tstrt='1970-01-01'
    if args.get('TIME_START'):
        Tstrt = datetime.strptime(str(args['TIME_START'])[:10], "%Y-%m-%d").strftime("%Y-%m-%d")
    if status=='CLOSED'  and args.get('TIME_FINISH'):
        Tfnsh = datetime.strptime(str(args['TIME_FINISH'])[:10], "%Y-%m-%d").strftime("%Y-%m-%d")
        
        if Tfnsh == today:
            tstart = datetime.strptime(str(args['TIME_START'])[11:19], '%H:%M:%S').strftime('%H:%M')
            tfnsh = str(args['TIME_FINISH'])[11:16]
            dsecs1 = int(str(args['DURATION'])[:2]) * 3600 + int(str(args['DURATION'])[3:5]) * 60 + int(str(args['DURATION'])[-2:])
            
            # Получение текущих данных из таблицы
            result = service.spreadsheets().values().get(spreadsheetId=sheet_id, range=sheet_name).execute()
            values = result.get('values', [])
            
            if not values:
                headers = ['ID', 'NAME'] + [today, ' ', ' ']
                values.append(headers)
                values.append([' 1', ' 2',])
            
            # Определение текущего столбца (день недели)
            current_column = None
            for i, header in enumerate(values[0]):
                if header == today:
                    current_column = i
                    break

            # Если текущей день недели еще не добавлен в таблицу, добавляем новый столбец
            if current_column is None:
                new_day = datetime.now().strftime("%Y-%m-%d")
                values[0].append(new_day)
                current_column = len(values[0])-1
                values[0].append(' ')
                values[0].append(' ')
                values[1].append('начало')
                values[1].append('конец')
                values[1].append('длительность')
                

                # Заголовки для нового столбца
                #values[0].extend(['', '', ''])
                service.spreadsheets().values().update(spreadsheetId=sheet_id, range=sheet_name,
                                                       body={'values': values}, valueInputOption='RAW').execute()
                        
            # Определение текущей строки пользователя
            user_row = None
            for i, row in enumerate(values):
                if len(row) > 0 and row[0] == str(id):
                    user_row = i
                    break
            
            if len(values[user_row])<len(values[1]):
                for i in range(len(values[1])-len(values[user_row])):
                    values[user_row].append('')
            
            # Если пользователя еще нет в таблице, добавляем новую строку
            if user_row is None:
                user_row = len(values)
                values.append([str(id), name] + [''] * (len(values[0]) - 5))
            
            # Запись новых данных в таблицу
            values[user_row][current_column:current_column+3] = [tstart, tfnsh, str(args['DURATION'])]
            service.spreadsheets().values().update(spreadsheetId=sheet_id, range=sheet_name,
                                                   body={'values': values}, valueInputOption='RAW').execute()
            
    elif status=='OPENED' and Tstrt == today:
                tstart = datetime.strptime(str(args['TIME_START'])[11:19], '%H:%M:%S').strftime('%H:%M')
                #tfnsh = str(args['TIME_FINISH'])[11:16]
                #dsecs1 = int(str(args['DURATION'])[:2]) * 3600 + int(str(args['DURATION'])[3:5]) * 60 + int(str(args['DURATION'])[-2:])
                
                # Получение текущих данных из таблицы
                result = service.spreadsheets().values().get(spreadsheetId=sheet_id, range=sheet_name).execute()
                values = result.get('values', [])
                
                if not values:
                    headers = ['ID', 'NAME'] + [today, ' ', ' ']
                    values.append(headers)
                    values.append([' 1', ' 2'])
                # Определение текущего столбца (день недели)
                current_column = None
                for i, header in enumerate(values[0]):
                    if header == today:
                        current_column = i
                        break

                # Если текущей день недели еще не добавлен в таблицу, добавляем новый столбец
                if current_column is None:
                    new_day = datetime.now().strftime("%Y-%m-%d")
                    values[0].append(new_day)
                    current_column = len(values[0])-1
                    values[0].append(' ')
                    values[0].append(' ')
                    values[1].append('начало')
                    values[1].append('конец')
                    values[1].append('длительность')

                    # Заголовки для нового столбца
                    #values[0].extend(['', '', ''])
                    service.spreadsheets().values().update(spreadsheetId=sheet_id, range=sheet_name,
                                                           body={'values': values}, valueInputOption='RAW').execute()
                            
                # Определение текущей строки пользователя
                user_row = None
                for i, row in enumerate(values):
                    if len(row) > 0 and row[0] == str(id):
                        user_row = i
                        break
                
                if len(values[user_row])<len(values[1]):
                    for i in range(len(values[1])-len(values[user_row])):
                        values[user_row].append('')
                
                # Если пользователя еще нет в таблице, добавляем новую строку
                if user_row is None:
                    user_row = len(values)
                    values.append([str(id), name] + [''] * (len(values[0]) - 5))
                
                # Запись новых данных в таблицу
                values[user_row][current_column:current_column+3] = [tstart, ' ', " "]
                service.spreadsheets().values().update(spreadsheetId=sheet_id, range=sheet_name,
                                                       body={'values': values}, valueInputOption='RAW').execute()
 
            
def push(id,name,args, service, sheet_id, sheet_name):
    today = datetime.now().strftime("%Y-%m-%d")
    status=args['STATUS']
    Tstrt='1970-01-01'
    if args.get('TIME_START'):
        Tstrt = datetime.strptime(str(args['TIME_START'])[:10], "%Y-%m-%d").strftime("%Y-%m-%d")
    if status=='CLOSED'  and args.get('TIME_FINISH'):
        Tfnsh = datetime.strptime(str(args['TIME_FINISH'])[:10], "%Y-%m-%d").strftime("%Y-%m-%d")
        
        if Tfnsh == today:
            tstart = datetime.strptime(str(args['TIME_START'])[11:19], '%H:%M:%S').strftime('%H:%M')
            tfnsh = str(args['TIME_FINISH'])[11:16]
            dsecs1 = int(str(args['DURATION'])[:2]) * 3600 + int(str(args['DURATION'])[3:5]) * 60 + int(str(args['DURATION'])[-2:])
            
            # Получение текущих данных из таблицы
            result = service.spreadsheets().values().get(spreadsheetId=sheet_id, range=sheet_name).execute()
            values = result.get('values', [])
            
            if not values:
                headers = ['ID', 'NAME']
                values.append(headers)
                values.append([' 1', ' 2',])
            
            # Определение текущего столбца (день недели)
            current_column = None
            for i, header in enumerate(values[0]):
                if header == today:
                    current_column = i
                    break
                
            # Если текущей день недели еще не добавлен в таблицу, добавляем новый столбец
            if current_column is None:
                new_day = datetime.now().strftime("%Y-%m-%d")
                values[0].append(new_day)
                current_column = len(values[0])-1
                values[0].append(' ')
                #values[0].append(' ')
                values[1].append('начало')
                values[1].append('конец')
                #values[1].append('длительность')
                

                # Заголовки для нового столбца
                #values[0].extend(['', '', ''])
                service.spreadsheets().values().update(spreadsheetId=sheet_id, range=sheet_name,
                                                       body={'values': values}, valueInputOption='RAW').execute()
                        
            # Определение текущей строки пользователя
            user_row = None
            for i, row in enumerate(values):
                if len(row) > 0 and row[0] == str(id):
                    user_row = i
                    break
            
            if len(values[user_row])<len(values[1]):
                for i in range(len(values[1])-len(values[user_row])):
                    values[user_row].append('')
            
            # Если пользователя еще нет в таблице, добавляем новую строку
            if user_row is None:
                user_row = len(values)
                values.append([str(id), name] + [''] * (len(values[0]) - 4))
            
            # Запись новых данных в таблицу
            values[user_row][current_column:current_column+2] = [tstart, tfnsh]
            service.spreadsheets().values().update(spreadsheetId=sheet_id, range=sheet_name,
                                                   body={'values': values}, valueInputOption='RAW').execute()
            
    elif status=='OPENED' and Tstrt == today:
                tstart = datetime.strptime(str(args['TIME_START'])[11:19], '%H:%M:%S').strftime('%H:%M')
                #tfnsh = str(args['TIME_FINISH'])[11:16]
                #dsecs1 = int(str(args['DURATION'])[:2]) * 3600 + int(str(args['DURATION'])[3:5]) * 60 + int(str(args['DURATION'])[-2:])
                
                # Получение текущих данных из таблицы
                result = service.spreadsheets().values().get(spreadsheetId=sheet_id, range=sheet_name).execute()
                values = result.get('values', [])
                
                if not values:
                    headers = ['ID', 'NAME'] 
                    values.append(headers)
                    values.append([' 1', ' 2'])
                # Определение текущего столбца (день недели)
                current_column = None
                for i, header in enumerate(values[0]):
                    if header == today:
                        current_column = i
                        
                        break

                # Если текущей день недели еще не добавлен в таблицу, добавляем новый столбец
                if current_column is None:
                    new_day = datetime.now().strftime("%Y-%m-%d")
                    values[0].append(new_day)
                    current_column = len(values[0])-1
                    values[0].append(' ')
                    #values[0].append(' ')
                    values[1].append('начало')
                    values[1].append('конец')
                    #values[1].append('длительность')

                    # Заголовки для нового столбца
                    #values[0].extend(['', '', ''])
                    service.spreadsheets().values().update(spreadsheetId=sheet_id, range=sheet_name,
                                                           body={'values': values}, valueInputOption='RAW').execute()
                            
                # Определение текущей строки пользователя
                user_row = None
                for i, row in enumerate(values):
                    if len(row) > 0 and row[0] == str(id):
                        user_row = i
                        break
                if len(values[user_row])<len(values[1]):
                    for i in range(len(values[1])-len(values[user_row])):
                        values[user_row].append('')
                    
                    
                # Если пользователя еще нет в таблице, добавляем новую строку
                if user_row is None:
                    user_row = len(values)
                    values.append([str(id), name] + [''] * (len(values[0]) - 4))
                
                # Запись новых данных в таблицу
                values[user_row][current_column:current_column+2] = [tstart, ' ']
                service.spreadsheets().values().update(spreadsheetId=sheet_id, range=sheet_name,
                                                       body={'values': values}, valueInputOption='RAW').execute()
            
def main():
    creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    service = build('sheets', 'v4', credentials=creds)
    o=11076
    while True:
        try:
            o+=1
            users_request = requests.get('')
            users = users_request.json()
            
            for i in users['result']:
                user_id = int(i['ID'])
                #print(i['ID'], i['NAME'], i['LAST_NAME'], getworkday(user_id)['STATUS'])
                pushtoya(user_id,str(i['LAST_NAME']+' '+ i['NAME']),getworkday(user_id), service, SPREADSHEET_ID, SHEET_NAME)
                push(user_id,str(i['LAST_NAME']+' '+ i['NAME']),getworkday(user_id), service, SPREADSHEET_ID, SHEET_NAME1)

        except Exception as e:
            print('5  sec')
            print(traceback.format_exc())
        # Пауза в одну минуту перед следующим обновлением
        time.sleep(5)
        print(o, datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

if __name__ == '__main__':
    main()
