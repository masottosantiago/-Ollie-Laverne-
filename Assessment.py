# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import time
import pandas as pd
import glob
import os
import sqlite3

#Function to automate webpage
def download_file (link, download_directory):
    s = Service(r"C:\Program Files (x86)\chromedriver.exe")
    options = webdriver.ChromeOptions() ;

    prefs = {"download.default_directory" : download_directory};

    options.add_experimental_option("prefs",prefs);

    driver = webdriver.Chrome(service=s, options = options)
    
    driver.get(link)
    driver.find_element(By.CLASS_NAME, 'wp-block-button__link').click()

    time.sleep(5) #give time for download to finish before closing the webpage
    driver.close()

#Function to get the latest file in the download directory
def latest_file(download_directory):
    list_of_files = glob.glob(download_directory + '\*.xlsx') 
    latest_file = max(list_of_files, key=os.path.getctime)
    return latest_file
    
link = "https://jobs.homesteadstudio.co/data-engineer/assessment/download/"
download_directory = "C:\Assessment"

download_file(link, download_directory) #access the webpage stated in the link and save the file in the download directory
file = latest_file(download_directory) #get the latest file downloaded


xls = pd.ExcelFile(file) #read all the worksheets
df1 = pd.read_excel(xls, 'data') #create worksheet for data worksheet
df2 = pd.read_excel(xls, 'pivot_table') #create worksheet for pivot_table worksheet

#Create pivot table with index set to Platform(Northbeam) and affregated sum of some columns as stated
pivot = df1.pivot_table(index=['Platform (Northbeam)'], values=['Spend','Attributed Rev (1d)','Imprs','Visits',
                                                                'New Visits', 'Transactions (1d)',
                                                                'Email Signups (1d)'], aggfunc='sum')

#Sort pivot table based on descending values of the Attributed Rev column
df_sorted = pivot.sort_values(by=['Attributed Rev (1d)'],  ascending=False)

#Add the index [Platform(Northbeam)] back into the pivot table
pivot_table_final = df_sorted.rename_axis(None, axis=1).reset_index()

#Change the Column Headers to match the Database Headers
pivot_table_final.columns = ['Row_Labels','Spend','Attributed_Rev','Imprs', 'Visits', 'New_Visits', 'Transactions', 'Email_Signups']

#Connect to database and transfer Pivot Table Data
con = sqlite3.connect("pivot.db") #create pivot database
cur = con.cursor()

#If table does not exist, create table named pivot table with the following columns and data type
cur.execute('''CREATE TABLE IF NOT EXISTS pivot_table
            (ID INT PRIMARY KEY NOT NULL,
             Row_Labels TEXT NOT NULL, 
             Spend REAL,
             Attributed_Rev REAL, 
             Imprs INT,
             Visits INT, 
             New_Visits INT, 
             Transactions REAL,
             Email_Signups REAL);''')

#Send Pivot Table Data to Database. If the table exist, replace the data with the new data 
pivot_table_final.to_sql('pivot_table',con,if_exists='replace',index=False)
cur.close()
con.close()

#%% Test if data was sent to database
# Read sqlite query results into a pandas DataFrame
con = sqlite3.connect("pivot.db")
cur = con.cursor()
dftest = pd.read_sql_query("SELECT * from pivot_table", con)

# Verify that result of SQL query is stored in the dataframe
print(dftest.head())

con.close()