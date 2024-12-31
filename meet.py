
import xlrd
import pandas as pd
import aiohttp

dataframe = pd.read_excel('time-checking-analysis.xlsx')
df = dataframe
df = df.dropna()
df = df.rename(columns=df.iloc[0]).drop(df.index[0]).drop('Sr.No.', axis=1)
print(df.head())

#exporting datafram to csv file 
export_df = df.to_csv(r'newnewfilefile.csv')

# exporting datafram to excel file
export_df = df.to_excel(r'newnewfilefile.csv')


# Python code to illustrate Sending mail from
# your Gmail account
import smtplib, ssl

port = 587
smtp_server_gmail = "smtp.gmail.com"
sender_email = "ramondjemon@gmail.com"
receiver_email = "rahul.rachh3005@gmail.com"
password = "newton1234"
message = "Hello teesting gmail smtp server. Heres to the awesome automation " \
          "that happens in python."

context=ssl.create_default_context()
# creates SMTP session
with smtplib.SMTP(smtp_server_gmail, port) as s:
    s.ehlo()
# start TLS for security
    s.starttls(context=context)
    s.ehlo()
# Authentication
    s.login(sender_email, password)
# message to be sent
# sending the mail
    s.sendmail(sender_email, receiver_email, message)
# terminating the session
# s.quit()

print('Done!!!!!!!!!!')

# ABC
# ABC

# ABC

# ABC

# ABC
import pandas as pd
# ABC
try:
    # read excel file
    f = pd.read_excel("36_Month_Sales_Forecast_13.xls", sheet_name='Sheet1,Singham')
    print(f)

except Exception as e: print(e)

# read text file
with open('new.txt') as f:
    print(f.readlines())



# uses python3
# rdbms used id SQL Server by MS
import pyodbc
cnxn = pyodbc.connect('DRIVER={SQL Server};'
                      'SERVER=HARSH;'
                      # 'DATABASE=StudentDb'
                      'UID=sa;'
                      'PWD=cimcon#123;'
                      )

cursor = cnxn.cursor()
# nocount = "SET NOCOUNT ON "
sql_insert_query = """
                    INSERT INTO StudentDb.dbo.StudentReg (Id,Name,Address,City) 
                    VALUES 
                    (889, 'Stevie Gerrard ', 'Anfield', 'Liverpoool, England(Not in Europe Anymore)')
                    """

# cursor.execute("SET IDENTITY_INSERT StudentDb.dbo.StudentReg ON")
# cursor.execute(sql_insert_query)
# cnxn.commit()
cursor.execute("select * from StudentDb.dbo.StudentReg")
# cursor.fetchall()
print('The rows of this table are: ')
print(tuple(i[0] for i in cursor.description))
# print(cursor.fetchall())
# print(cursor)

for row in cursor.fetchall():
    print(row)

# print('\nThe columns of this table are: ', [i[0] for i in cursor.description])


cnxn.close()


#uses python3
# fetch NIFTY50 price and its changes from moneycontrol.com/stocksmarketindia/

import urllib.request as ur
from bs4 import BeautifulSoup


quote_page = "https://www.moneycontrol.com/stocksmarketsindia/"

page = ur.urlopen(quote_page)

soup = BeautifulSoup(page, 'lxml')

name_box = soup.find('a', attrs={'class': 'robo_medium'}).text
price_box = soup.find('td', attrs={'class': 'tbl_redtxt'}).find_previous_sibling().text
change_box = soup.find('td', attrs={'class': 'tbl_redtxt'}).text
perchange_box = soup.find('td', attrs={'class': 'tbl_redtxt'}).find_next_sibling().text
# price_box = soup.find('td', attrs={'class': 'tdred'}).next

# name =
# print(name)
print("Index: ", name_box)
print("Price: ", price_box)
print("Change: ", change_box)
print("%Chng: ", perchange_box)

# uses Python3
# writing different types of files from python

with open('newwritingfile', 'w') as f:
    f.write('Helloe')

with open('newwritingfile', 'a') as f:
    f.write('Helloe')

with open('newwritingfile', 'r') as f:
    f.write('Helloe')

with open('newwritingfile', 'r+') as f:
    f.write('Helloe')

with open('newwritingfile', 'w+') as f:
    f.write('Helloe')

# import aiohttp
# SSNSelect * from Repository 
