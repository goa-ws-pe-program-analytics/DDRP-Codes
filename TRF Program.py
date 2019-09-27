# -*- coding: utf-8 -*-
"""
Created on Tue Sep 24 09:29:54 2019

@author: isaac.nyamekye
"""

# # Import Required Libraries 



import pandas as pd
import datetime as dt
import numpy as np
import pypyodbc


# # Extract Data from TRF SQL Server


#Connects to TRF data located on SQL Server
connection = pypyodbc.connect('Driver={SQL Server};'
                              'Server=EDM-GOA-SQL-452;'
                              'Database=A_STAGING;')
#Defines and Executes SQL code to pull required variables from SQL Server
sql = ('SELECT trf_id,'
            'MAX(referraldate) AS referral_date,'
            'MAX(completeddatetime) AS completed_date,'
            'MAX(noc_noc) AS noc,'
            'MAX(age) AS age,'
            'MAX(postalcode) AS postal_Code,'
            'MAX(gender_fmt) AS gender,'
            'MAX(esdcfeedback) AS esdc_feedback_num,'
            'MAX(esdcfeedback_fmt) AS esdc_feedback,'
            'MAX(serviceproviderfeedback_fmt) AS service_provider_feedback'
            ' FROM trf.referral GROUP BY trf_id ORDER BY trf_id;')
df1 = pd.read_sql(sql, connection)
connection.close()


# # Import Postal Code Translator File (PCTF)


#Imports PCTF from shared drive
PCTF1 = pd.read_excel('M:/WS/Program Effectiveness/Program Analytics/DDRP/Monthly_Report/2019-20/_Latest_Month/Postal Code Translator File.xlsx')

#Subset only the columns of interest
PCTF2 = PCTF1[['POSTALCODE','ERNAME_2016']]

#Renames columns
PCTF= PCTF2.rename(index=str, columns={"POSTALCODE": "postal_code", "ERNAME_2016": "Economic_Region"})


# # Add Economic Region from PCTF to TRF Data


df = pd.merge(df1, PCTF, how='left', on=['postal_code'],sort=False,validate='many_to_one')


# # Manipulation & Cleaning via Python (Pandas)


# Converts raw data into Excel Required Date Format (Vlookups)
df['ref_date_text'] = df['referral_date'].dt.strftime('%b%Y')
df['comp_date_text'] = df['completed_date'].dt.strftime('%b%Y')

# Converts Ages into Age Brackets
    ## IDEAL BRACKETS ##
#labels = ["0 - 14","15 - 24","25 - 54","55 - 64","65+"]
#df['age_bracket'] = pd.cut(df['age'],bins=[0,15,25,55,65,300], right=False,labels=labels)

    ## OLD BRACKETS ##
labels = ["Under 25","25-44","45-54","55+"]
df['age_bracket'] = pd.cut(df['age'],bins=[0,25,45,55,300], right=False,labels=labels)

# Converts NOC Codes into Skill Levels
def skill_level(noc_code):
    first = str(int(str(noc_code)[:1]))
    second = str(int(str(noc_code)[1:2]))
    if first == '0':
        skill_level = 'NOC 0'
    elif second == '0':
        skill_level = 'NOC A'
    elif second == '1':
        skill_level = 'NOC A'
    elif second == '2':
        skill_level = 'NOC B'
    elif second == '3':
        skill_level = 'NOC B'    
    elif second == '4':
        skill_level = 'NOC C'    
    elif second == '5':
        skill_level = 'NOC C' 
    elif second == '6':
        skill_level = 'NOC D'       
    elif second == '7':
        skill_level = 'NOC D'
    else: 
        skill_level = 'Error'
    return skill_level

df['skill_level'] = df['noc'].apply(skill_level)

# Converts NOC Codes into Broad Occupational Category
def skill_type(noc_code):
    first = str(int(str(noc_code)[:1]))
    if first == '0':
        skill_type = '0'
    elif first == '1':
        skill_type = '1'
    elif first == '2':
        skill_type = '2'
    elif first == '3':
        skill_type = '3'
    elif first == '4':
        skill_type = '4'  
    elif first == '5':
        skill_type = '5'    
    elif first == '6':
        skill_type = '6'
    elif first == '7':
        skill_type = '7'   
    elif first == '8':
        skill_type = '8'
    elif first == '9':
        skill_type = '9'
    else: 
        skill_type= 'Error'
    return skill_type

df['skill_type'] = df['noc'].apply(skill_type)

# Creates new column called completed which has a '2' if the completed_date is later than 1900-01-00. Else it has a '1'

def completed (completed_date):
    success = 2
    if str(completed_date) <= '1900-01-01 00:00:00':
        success = 0
    else:
        success = 1
    return success
    
df['completed'] = df['completed_date'].apply(completed)

# Creates a new column called "feedback" which contains a '2' if the esdc_feedback falls into any of the following categories:
    # 100000000 Client not interested – Other
    # 100000001 Client not interested – Employed
    # 100000002 Client interested – Client has identified desired services
    # 100000003 Client interested – Client undecided about desired service(s)
    # 100000004 Client not interested – In training/school
    # 100000007 Client's case transferred
    # 100000008 Client not interested – Family responsibilities
    # 100000009 Client not interested – Health
    # 100000010 Client not interested – Moving
    # 100000011 Client not interested – No Help Required


def feedback (esdc_feedback_num):
    if 100000000 <= esdc_feedback_num <=100000004:
         success = 1
    elif 100000005 <= esdc_feedback_num <=100000006:
         success = 0
    elif 100000007 <= esdc_feedback_num <=100000011:
         success = 1
    else:
         success = 0
    return success

df['feedback'] = df['esdc_feedback_num'].apply(feedback)

# Creates a new column called "successful_contact" which contains a '2' if both the '2' conditions above have been met
df['successful_contact'] = df.apply(lambda row: row.feedback + row.completed, axis = 1)

def success (num):
    if num == 2:
        success = True
    else:
        success = False
    return success
  
df['successful_contact'] = df['successful_contact'].apply(success)
df = df.drop(['completed','feedback'],axis = 1)


# # Export to Data Metrics Excel File

export_file = 'M:/WS/Program Effectiveness/Program Analytics/DDRP/Monthly_Report/2019-20/_Latest_Month/Raw_Data/TRF_RAW.xlsx'
df.to_excel(export_file,sheet_name='TRF_CRM_Extract',index=False,startrow=1, startcol=1)






