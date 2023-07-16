import pandas as pd
import numpy as np
import pyodbc
import openpyxl
from sqlalchemy import create_engine
from sqlalchemy.engine import URL
from pathlib import Path
from openpyxl import load_workbook, styles, formatting
import sys
import os
import smartsheet
from datetime import datetime, timedelta
import urllib.request

server = '$$$$' 
database = '$$$' 
username = '@@@@@' 
password = '!!!!!!' 
connection_string = 'DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password
connection_url = URL.create("mssql+pyodbc", query={"odbc_connect": connection_string})
engine = create_engine(connection_url)
print("Connection Established:")

df=pd.read_sql('''
select a.shiptoparty as [Account Number], b.shiptoparty_dba as [Account Name], b.address as [Address], b.city as [City], b.state_key as [State], b.zipcode as [Zipcode],a.salesperson_key as [Salesman_Key], a.salesperson_text as [Salesman], a.salesteam_text as [Salesteam] from [ivy.mm.dim.sales_master_tmp] a
left join [ivy.mm.dim.shiptoparty] b on a.shiptoparty = b.shiptoparty
where b.active = 'active' and a.salesteam_text in ('RED TEAM 1','RED TEAM 2')
group by a.shiptoparty, b.shiptoparty_dba, b.address, b.city, b.state_key, b.zipcode, a.salesperson_key, a.salesperson_text, a.salesteam_text
order by salesteam_text asc
''',con=engine)
df=df.astype({"Account Number":"str", "Account Name":"str","Address":"str","City":"str","State":"str","Zipcode":"str","Salesman_Key":"str","Salesman":"str","Salesteam":"str"}).sort_values(by=["Salesteam"],ascending=True)


today = datetime.today().strftime("%m%d%Y")
data = pd.DataFrame(df)
data['Zipcode']= data['Zipcode'].str.split("-").str[0]

SalesMaster = data.to_csv(r'\\nas2\SOM Content DB\Ivykiss Artwork\ZZ_TEMP\SOM\14_Salesman Account Change Request Form\Salesman - Salesteam Master\Salesman - Salesteam Master RED.csv')
print('Salesman - Salesteam Master RED has been successfully updated!')

smartsheet_client = smartsheet.Smartsheet("$$$$$$$$$$$$")
sheet = smartsheet_client.Sheets.get_sheet(684744041097092)
newRow = smartsheet_client.models.Row()
newCell = smartsheet_client.models.Cell()
smartsheet_client.Sheets.add_rows(684744041097092, newRow)
smartsheet_client.Sheets.get_sheet(684744041097092)
for row in sheet.rows:
    thisRow=  row.id
smartsheet_client.Attachments.attach_file_to_row(684744041097092, thisRow, ('SalesMaster.csv', open(r'\\nas2\SOM Content DB\Ivykiss Artwork\ZZ_TEMP\SOM\14_Salesman Account Change Request Form\Salesman - Salesteam Master\Salesman - Salesteam Master RED.csv'))) 
   
print('Attached Successfully')

df=pd.read_sql('''
select a.shiptoparty as [Account Number], b.shiptoparty_dba as [Account Name], b.address as [Address], b.city as [City], b.state_key as [State], b.zipcode as [Zipcode],a.salesperson_key as [Salesman_Key], a.salesperson_text as [Salesman], a.salesteam_text as [Salesteam] from [ivy.mm.dim.sales_master_tmp] a
left join [ivy.mm.dim.shiptoparty] b on a.shiptoparty = b.shiptoparty
where b.active = 'active' and a.salesteam_text in ('IVY TEAM 1', 'IVY TEAM 2', 'IVY TEAM 3', 'IVY TEAM 4','IVY TEAM 5', 'IVY TEAM 6')
group by a.shiptoparty, b.shiptoparty_dba, b.address, b.city, b.state_key, b.zipcode, a.salesperson_key, a.salesperson_text, a.salesteam_text
order by salesteam_text asc
''',con=engine)
df=df.astype({"Account Number":"str", "Account Name":"str","Address":"str","City":"str","State":"str","Zipcode":"str","Salesman_Key":"str","Salesman":"str","Salesteam":"str"}).sort_values(by=["Salesteam"],ascending=True)

today = datetime.today().strftime("%m%d%Y")
data = pd.DataFrame(df)
data['Zipcode']= data['Zipcode'].str.split("-").str[0]
SalesMaster= data.to_csv(r'\\nas2\SOM Content DB\Ivykiss Artwork\ZZ_TEMP\SOM\14_Salesman Account Change Request Form\Salesman - Salesteam Master\Salesman - Salesteam Master IVY.csv')
print("updated!")

import smartsheet
smartsheet_client = smartsheet.Smartsheet("$$$$$$$$$$$")
sheet = smartsheet_client.Sheets.get_sheet(4625393715046276)
newRow = smartsheet_client.models.Row()
newCell = smartsheet_client.models.Cell()
smartsheet_client.Sheets.add_rows(4625393715046276, newRow)
smartsheet_client.Sheets.get_sheet(4625393715046276)
for row in sheet.rows:
    thisRow=  row.id
smartsheet_client.Attachments.attach_file_to_row(4625393715046276, thisRow, ('SalesMaster.csv', open(r'\\nas2\SOM Content DB\Ivykiss Artwork\ZZ_TEMP\SOM\14_Salesman Account Change Request Form\Salesman - Salesteam Master\Salesman - Salesteam Master IVY.csv', 'rb'), 'application/vnd.ms-excel')) 
   
print('Attached Successfully')

df=pd.read_sql('''
select distinct(salesperson_text),salesperson_key from [ivy.mm.dim.sales_master]
''',con=engine)
Salesman_Dropdown = df.to_csv(r'\\nas2\SOM Content DB\Ivykiss Artwork\ZZ_TEMP\SOM\14_Salesman Account Change Request Form\Salesman - Salesteam Master\SalesmanList.csv')

import smartsheet
smartsheet_client = smartsheet.Smartsheet("$$$$$$$$$$$")
sheet = smartsheet_client.Sheets.get_sheet(5274590839629700)
newRow = smartsheet_client.models.Row()
newCell = smartsheet_client.models.Cell()
smartsheet_client.Sheets.add_rows(5274590839629700,newRow)
smartsheet_client.Sheets.get_sheet(5274590839629700)
for row in sheet.rows:
    thisRow = row.id
smartsheet_client.Attachments.attach_file_to_row(5274590839629700, thisRow, ('SalesmanList.csv', open(r'\\nas2\SOM Content DB\Ivykiss Artwork\ZZ_TEMP\SOM\14_Salesman Account Change Request Form\Salesman - Salesteam Master\SalesmanList.csv','rb'),'application/vnd.ms-excel'))
print('Attached Successfully')

import win32com.client
ol = win32com.client.Dispatch('Outlook.Application')
olmailitem = 0x0
newmail = ol.CreateItem(olmailitem)
newmail.Subject = 'Salesman List, IVY Salesmaster Table, RED Salesmaster Table was uploaded on smartsheet successfully!'
newmail.To = '%%%%%%ivyent.com'
newmail.CC = '%%%%%ivyent.com'
newmail.Send()

