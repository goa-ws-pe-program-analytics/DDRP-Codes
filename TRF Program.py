# -*- coding: utf-8 -*-
"""
Created on Tue Sep 24 09:29:54 2019

@author: isaac.nyamekye
"""

# Import Required Libraries
import pandas as pd
import numpy as np
import pypyodbc
import datetime

# Extract Data from TRF SQL Server

# Connects to TRF data located on SQL Server
connection = pypyodbc.connect('Driver={SQL Server};'
                              'Server=EDM-GOA-SQL-452;'
                              'Database=A_STAGING;')
# Defines and Executes SQL code to pull required variables from SQL Server
sql = ('SELECT trf_id,'
       'MAX(referraldate) AS referral_date,'
       'MAX(completeddatetime) AS completed_date,'
       'MAX(noc_noc) AS noc,'
       'MAX(age) AS age,'
       'MAX(postalcode) AS postal_Code,'
       'MAX(gender_fmt) AS gender,'
       'MAX(referralstatus_fmt) AS referral_status,'
       'MAX(esdcfeedback) AS esdc_feedback_num,'
       'MAX(esdcfeedback_fmt) AS esdc_feedback,'
       'MAX(serviceproviderfeedback_fmt) AS service_provider_feedback'
       ' FROM trf.referral GROUP BY trf_id ORDER BY trf_id;')
df1 = pd.read_sql(sql, connection)
connection.close()

# Import Postal Code Translator File (PCTF)
PCTF1 = pd.read_excel('M:/WS/Program Effectiveness/Program Analytics/DDRP/Monthly_Report/2019-20/_Latest_Month/Postal Code Translator File.xlsx')

# Subset only the columns of interest
PCTF2 = PCTF1[['POSTALCODE', 'ERNAME_2016']]

# Renames columns
PCTF = PCTF2.rename(index=str, columns={"POSTALCODE": "postal_code", "ERNAME_2016": "Economic_Region"})

# Add Economic Region from PCTF to TRF Data
df = pd.merge(df1, PCTF, how='left', on=['postal_code'], sort=False, validate='many_to_one')

# Capitalize only the first letter of word in Economic Region. The original Economic Region values are all upper case
df['Economic_Region'] = df['Economic_Region'].str.title()

# Creating Referral and Completion Year, Fiscal Year and Month
## Note: The fiscal year (Aor-Mar) is represented by the ending year. E.g. 2018-19 would be 2019

df['ref_year'] = df['referral_date'].dt.strftime('%Y')
df['comp_year'] = df['completed_date'].dt.strftime('%Y')

df['ref_fyear'] = pd.to_datetime(df['referral_date']).apply(pd.Period, freq='A-MAR')
df['comp_fyear'] = pd.to_datetime(df['completed_date']).apply(pd.Period, freq='A-MAR')

df['ref_month'] = df['referral_date'].dt.strftime('%b').apply(lambda x: x.upper())
df['comp_month'] = df['completed_date'].dt.strftime('%b').apply(lambda x: x.upper())

# Converts Ages into Age Brackets
## IDEAL BRACKETS ##
# labels = ["0 - 14","15 - 24","25 - 54","55 - 64","65+"]
# df['age_bracket'] = pd.cut(df['age'],bins=[0,15,25,55,65,300], right=False,labels=labels)

## OLD BRACKETS ##
labels = ["< 25", "25-44", "45-54", "55+"]
df['age_bracket'] = pd.cut(df['age'], bins=[0, 25, 45, 55, 300], right=False, labels=labels)


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
        skill_type = '1: Business, finance and administration'
    elif first == '2':
        skill_type = '2: Natural and applied sciences'
    elif first == '3':
        skill_type = '3: Health'
    elif first == '4':
        skill_type = '4: Education, law and social, community and government services'
    elif first == '5':
        skill_type = '5: Art, culture, recreation and sport'
    elif first == '6':
        skill_type = '6: Sales and service'
    elif first == '7':
        skill_type = '7: Trades, transport and equipment operators'
    elif first == '8':
        skill_type = '8: Natural resources, agriculture'
    elif first == '9':
        skill_type = '9: Manufacturing and utilities'
    else:
        skill_type = 'Error'
    return skill_type


df['skill_type'] = df['noc'].apply(skill_type)


# Creates a new column called "successful_contact" which contains a 'Yes' if the esdc_feedback falls into any of the
# following categories:
## This shows whether providers were able to get in touch with referral

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

def feedback(esdc_feedback_num):
    if 100000000 <= esdc_feedback_num <= 100000004:
        success = 'Yes'
    elif 100000005 <= esdc_feedback_num <= 100000006:
        success = 'No'
    elif 100000007 <= esdc_feedback_num <= 100000011:
        success = 'Yes'
    else:
        success = 'No'
    return success


df['successful_contact'] = df['esdc_feedback_num'].apply(feedback)


def further_service(service_provider_feedback):
    if service_provider_feedback == 'Attended meeting – Received employment services':
        interested_further_service = 'Clients Interested in Further Services'
    elif service_provider_feedback == 'Attended meeting – Will participate in employment program':
        interested_further_service = 'Clients Interested in Further Services'
    elif service_provider_feedback == 'Interested – Referred to Another Service Provider':
        interested_further_service = 'Clients Interested in Further Services'
    elif service_provider_feedback == 'Interested – Meeting Scheduled':
        interested_further_service = 'Clients Interested in Further Services'
    elif service_provider_feedback == 'Interested – Will call back to schedule meeting':
        interested_further_service = 'Clients Interested in Further Services'
    else:
        interested_further_service = 'Clients Not Interested in Further Services'
    return interested_further_service


df['interested_further_service'] = df['service_provider_feedback'].apply(further_service)

TRF = pd.read_excel(r"\\goa\desktop\E_J\isaac.nyamekye\Desktop\DDRP - Monthly_Report.xlsx", index_col=0, sheet_name='TRF_T1_19-20', usecols='A:A')

# Extracting Month for Today's Date
ThisMonth = datetime.date.today().replace(day=1).strftime("%b").upper()

df_empty = pd.DataFrame(
    columns=['APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC', 'JAN', 'FEB', 'MAR', 'YTD', '', 'YTD1', '%∆1', 'YTD2', '%∆2'])


def func_df(fy, fy1, out_df, val, ind, col, ind_n):
    odf = df[fy]

    out_df = pd.pivot_table(odf, values=val, index=ind, columns=col, aggfunc=len)

    out_df = out_df.rename(index={val: ind_n})

    if ThisMonth != 'APR':
        del out_df[ThisMonth]

    out_df = pd.concat([df_empty, out_df], sort=False)

    out_df['YTD'] = out_df.sum(axis=1)

    odf1 = df[fy1]

    df2 = pd.pivot_table(odf1, values=val, index=ind, columns=col, aggfunc=len)

    df2 = df2.rename(index={val: ind_n})

    df2 = pd.concat([df_empty, df2], sort=False)

    df2 = pd.DataFrame(df2[['APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC', 'JAN', 'FEB', 'MAR']].where(
        out_df[['APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC', 'JAN', 'FEB', 'MAR']].notnull(),
        np.nan))

    df2['YTD1'] = df2.sum(axis=1)

    out_df['YTD1'] = df2['YTD1']

    out_df['%∆1'] = out_df['YTD'] / out_df['YTD1'] - 1

    return out_df


files_received, contact_Attempts, interactions, by_gender, by_age, by_economic_region, by_skill_level, by_interested_further_service, \
by_interested_feedback, by_not_interested_feedback = [], [], [], [], [], [], [], [], [], []

files_received = func_df(df['ref_fyear'] == '2020', df['ref_fyear'] == '2019', files_received, 'ref_year', None, 'ref_month', 'Client Files Received')

contact_Attempts = func_df(df['comp_fyear'] == '2020', df['comp_fyear'] == '2019', contact_Attempts, 'comp_year', None, 'comp_month',
                           'Client Contact Attempts')

interactions = func_df((df['comp_fyear'] == '2020') & (df['successful_contact'] == 'Yes'),
                       (df['comp_fyear'] == '2019') & (df['successful_contact'] == 'Yes'), interactions, 'comp_year', None, 'comp_month',
                       'Client Interactions')

by_gender = func_df((df['comp_fyear'] == '2020') & (df['successful_contact'] == 'Yes'),
                    (df['comp_fyear'] == '2019') & (df['successful_contact'] == 'Yes'), by_gender, 'comp_year', 'gender', 'comp_month', 'By Gender')

gender = TRF[15:19]
by_gender = by_gender.join(gender, how='outer')

by_age = func_df((df['comp_fyear'] == '2020') & (df['successful_contact'] == 'Yes'),
                 (df['comp_fyear'] == '2019') & (df['successful_contact'] == 'Yes'), by_age, 'comp_year', 'age_bracket', 'comp_month', 'By Age')

age = TRF[20:24]
by_age = by_age.join(age, how='outer')

by_economic_region = func_df((df['comp_fyear'] == '2020') & (df['successful_contact'] == 'Yes'),
                             (df['comp_fyear'] == '2019') & (df['successful_contact'] == 'Yes'), by_economic_region, 'comp_year', 'Economic_Region',
                             'comp_month', 'By Age')

economic_region = TRF[25:35]
by_economic_region = pd.concat([by_economic_region, economic_region], axis=1, sort=False)

by_skill_level = func_df((df['comp_fyear'] == '2020') & (df['successful_contact'] == 'Yes'),
                         (df['comp_fyear'] == '2019') & (df['successful_contact'] == 'Yes'), by_skill_level, 'comp_year', 'skill_level', 'comp_month',
                         'By Skill Level')

skill_level = TRF[36:41]
by_skill_level = by_skill_level.join(skill_level, how='outer')

by_skill_type = func_df((df['comp_fyear'] == '2020') & (df['successful_contact'] == 'Yes'),
                        (df['comp_fyear'] == '2019') & (df['successful_contact'] == 'Yes'), skill_type, 'comp_year', 'skill_type', 'comp_month',
                        'By Skill Type')

skill_type = TRF[42:51]
by_skill_type = by_skill_type.join(skill_type, how='outer')

by_interested_further_service = func_df((df['comp_fyear'] == '2020') & (df['successful_contact'] == 'Yes') & (
        df['interested_further_service'] == 'Clients Interested in Further Services'),
                                        (df['comp_fyear'] == '2019') & (df['successful_contact'] == 'Yes'), by_interested_further_service,
                                        'comp_year', 'interested_further_service', 'comp_month', 'Clients Interested in Further Services')

by_interested_feedback = func_df((df['comp_fyear'] == '2020') & (df['successful_contact'] == 'Yes') & (
        df['interested_further_service'] == 'Clients Interested in Further Services'),
                                 (df['comp_fyear'] == '2019') & (df['successful_contact'] == 'Yes'), by_interested_feedback, 'comp_year',
                                 'service_provider_feedback', 'comp_month', 'Clients Interested in Further Services')

interested_feedback = TRF[53:58]
by_interested_feedback = pd.concat([by_interested_feedback, interested_feedback], axis=1, sort=False)

by_not_interested_further_service = func_df((df['comp_fyear'] == '2020') & (df['successful_contact'] == 'Yes') & (
        df['interested_further_service'] == 'Clients Not Interested in Further Services'),
                                            (df['comp_fyear'] == '2019') & (df['successful_contact'] == 'Yes'), by_interested_further_service,
                                            'comp_year', 'interested_further_service', 'comp_month', 'Clients Not Interested in Further Services')

by_not_interested_feedback = func_df((df['comp_fyear'] == '2020') & (df['successful_contact'] == 'Yes') & (
        df['interested_further_service'] == 'Clients Not Interested in Further Services'),
                                     (df['comp_fyear'] == '2019') & (df['successful_contact'] == 'Yes'), by_not_interested_feedback, 'comp_year',
                                     'esdc_feedback', 'comp_month', 'Clients Not Interested in Further Services')

not_interested_feedback = TRF[60:67]
by_not_interested_feedback = pd.concat([by_not_interested_feedback, not_interested_feedback], axis=1, sort=False)


# Converting by gender, age, economic region, skill level and skill type values to percentage
def percent(var):
    var = (var / var.sum())
    var['%∆1'] = var['YTD'] - var['YTD1']
    return var


by_gender = percent(by_gender)  # gender to percentage
by_age = percent(by_age)  # age to percentage
by_economic_region = percent(by_economic_region)  # economic region to percentage
by_skill_level = percent(by_skill_level)  # skill level to percentage
by_skill_type = percent(by_skill_type)  # skill type to percentage


# Convert nan to "-"
def conv(out_df):
    for columns in out_df:
        if out_df[columns].sum() > 0:
            out_df[columns] = out_df[columns].fillna(0)
        else:
            out_df[columns] = out_df[columns].fillna("-")
    out_df[''] = np.nan
    out_df['YTD2'] = "N/A"
    out_df['%∆2'] = "N/A"
    out_df['%∆1'] = out_df['%∆1'].replace("-", 0)
    out_df = out_df.replace([np.inf, -np.inf], "-")
    return out_df


files_received = conv(files_received)
contact_Attempts = conv(contact_Attempts)
interactions = conv(interactions)
by_gender = conv(by_gender)
by_age = conv(by_age)
by_economic_region = conv(by_economic_region)
by_skill_level = conv(by_skill_level)
by_skill_type = conv(by_skill_type)
by_interested_further_service = conv(by_interested_further_service)
by_interested_feedback = conv(by_interested_feedback)
by_not_interested_further_service = conv(by_not_interested_further_service)
by_not_interested_feedback = conv(by_not_interested_feedback)

# Export to Data Metrics Excel File

from openpyxl import load_workbook

# from openpyxl.utils.dataframe import dataframe_to_rows

# Load Workbook
wb = load_workbook(r"\\goa\desktop\E_J\isaac.nyamekye\Desktop\DDRP - Monthly_Report.xlsx")

xl_writer = pd.ExcelWriter(r"\\goa\desktop\E_J\isaac.nyamekye\Desktop\DDRP - Monthly_Report.xlsx", engine='openpyxl')

xl_writer.book = wb

# Read all sheets in workbook
xl_writer.sheets = dict((ws.title, ws) for ws in wb.worksheets)

# Export dataframes to excel
files_received.to_excel(xl_writer, 'TRF_T1_19-20', header=False, index=False, startcol=2, startrow=9)
contact_Attempts.to_excel(xl_writer, 'TRF_T1_19-20', header=False, index=False, startcol=2, startrow=11)
interactions.to_excel(xl_writer, 'TRF_T1_19-20', header=False, index=False, startcol=2, startrow=13)
by_gender.to_excel(xl_writer, 'TRF_T1_19-20', header=False, index=False, startcol=2, startrow=16)
by_age.to_excel(xl_writer, 'TRF_T1_19-20', header=False, index=False, startcol=2, startrow=21)
by_economic_region.to_excel(xl_writer, 'TRF_T1_19-20', header=False, index=False, startcol=2, startrow=26)
by_skill_level.to_excel(xl_writer, 'TRF_T1_19-20', header=False, index=False, startcol=2, startrow=37)
by_skill_type.to_excel(xl_writer, 'TRF_T1_19-20', header=False, index=False, startcol=2, startrow=43)
by_interested_further_service.to_excel(xl_writer, 'TRF_T1_19-20', header=False, index=False, startcol=2, startrow=53)
by_interested_feedback.to_excel(xl_writer, 'TRF_T1_19-20', header=False, index=False, startcol=2, startrow=54)
by_not_interested_further_service.to_excel(xl_writer, 'TRF_T1_19-20', header=False, index=False, startcol=2, startrow=59)
by_not_interested_feedback.to_excel(xl_writer, 'TRF_T1_19-20', header=False, index=False, startcol=2, startrow=61)

# Save Workbook
xl_writer.save()

# Close Workbook
wb.close()
