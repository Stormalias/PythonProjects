#Changlog for csvExtractor.py will be on FoundryMonitoringScript.py file

## Import packages
import csv
import datetime
import os
import pandas as pd
import pyodbc
import sys

#Variable Declarations
DataSource = 'DSN=foundryTest' + ';' #Change this name to whatever you named your datasource
csvDir = sys.path[0] 
header_list = ["Valuation_Date", "Bench", "Line_Item_Type", "Portfolio", "DTD", "MTD", "YTD", "Prev_MTD", "Prev_YTD", "MTD_Diff",\
                "YTD_Diff", "MTD_Check", "YTD_Check", "Overall_Check", "MTD_Check_Text", "YTD_Check_Text", "Portfolio_Details",\
                "Error_Details", "RefData_Datetime", "Run_Datetime"]
                

#Function to retieve last business day                
def get_previous_business_day():
    today = pd.Timestamp.today()
    previous_day = today - pd.DateOffset(days=1)
    previous_business_day = pd.bdate_range(end=previous_day, periods=1)[0]
    return previous_business_day.strftime("%Y-%m-%d")


#Runs SQL query on server. Change DSN name for name given on ODBC
def SQL_Query_CSV():
    cnxn = pyodbc.connect(DataSource)
    cursor = cnxn.cursor()

    sql_call = (f"""SELECT *
            FROM "/BP/IST-IG-DD/technical/dashboard_objects/pnl/Integrity_Checks/PNL050_Portfolio_DTD_Check"
            WHERE Valuation_Date = '{str(get_previous_business_day())}'
            """)
                                         
    print(sql_call)
    cursor.execute(sql_call)
    rows = cursor.fetchall()
    print("SQL Query success")
    return rows


def AfternoonCsv():
    with open(csvDir + '/Afternoon/COB_' + str(get_previous_business_day()) + '.csv', 'w', newline='') as testCsv:
        write = csv.writer(testCsv)

        write.writerow(header_list)
        write.writerows(SQL_Query_CSV())


def MorningCsv():
    
    with open(csvDir + '/Morning/COB_' + str(get_previous_business_day()) + '.csv', 'w', newline='') as testCsv:
        write = csv.writer(testCsv)
    
        write.writerow(header_list)
        write.writerows(SQL_Query_CSV())
