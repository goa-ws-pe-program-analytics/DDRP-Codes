# -*- coding: utf-8 -*-
"""
Created on Mon May 27 10:55:15 2019

@author: isaac.nyamekye
"""

import pandas as pd
import numpy as np

worksheets = pd.ExcelFile(
    r"M:/WS/Program Effectiveness/Program Analytics/DDRP/Monthly_Report/2019-20/_Latest_Month/Raw_Data/Raw Data AB LM Surveys 2011-2017.xlsx")

"""
# Reading all sheets in workbook into one dataframe
wslists = worksheets.sheet_names

listings = []

for wslist in wslists:
    listing = pd.read_excel(worksheets,sheet_name=wslist, na_values=['N/A','n/a','na','NA',' '])
    listing['Wslist'] = wslist
    listings.append(listing)

combined_data = pd.concat(listings)
"""


# Read Excel Sheets Separately
def read(df, sheet):
    df = pd.read_excel(worksheets, sheet_name=sheet, na_values=['N/A', 'n/a', 'na', 'NA', ' '])

    # Selecting Variables of Interest
    df = pd.DataFrame(df, columns=['Organization', 'Occupational Group ', 'Occupation Title', 'Regulated in Alberta',
                                   'Applications received from Albertans', 'Out-of-province Applications received ',
                                   'Total Applications received', 'Processing Time for Alberta Applications',
                                   'Processing Time for out-of-province Applications', 'Year'])

    # Renaming Occupational Group
    df['Occupational Group '] = df['Occupational Group '].replace(
        {'B': 'Business, Finance and Real Estate', 'E': 'Engineering, Architecture, Science and Technology',
         'H': 'Health and Social Services', 'L': 'Legal, Education and Government', 'O': 'Other', 'T': 'Other'})

    return df


# 2017 data
df2017 = []
df2017 = read(df2017, '2017 Data')

# 2016 data
df2016 = []
df2016 = read(df2016, '2016 Data')

# 2015 data
df2015 = []
df2015 = read(df2015, '2015 Data')

# 2014 data
df2014 = []
df2014 = read(df2014, '2014 Data')

# 2013 data
df2013 = []
df2013 = read(df2013, '2013 Data')

# 2012 data
df2012 = []
df2012 = read(df2012, '2012 Data')

# Dropping Excluded Occupations - 2017

Excluded_Occu_2017 = list(['Automotive Salesperson', 'Building Operator A&B', 'Chiropractors', 'Fireman (FIR)',
                           'Forest Technologists (Registered)', 'Forester (Registered Professional)',
                           'Horse Racing occupations', 'Hunting & Fishing Guides', 'Land Surveyors', 'Locksmiths',
                           'Occupational Therapists', 'Private Investigators', 'Security Workers',
                           'Shorthand Reporters', 'Speech Language Pathologist'])

df2017 = df2017.drop(df2017[(df2017['Occupation Title'].isin(Excluded_Occu_2017))].index)

# Dropping Excluded Occupations - 2016

Excluded_Occu_2016 = list(
    ['Asbestos Worker', 'Bridge Inspector and Maintenance System Inspector', 'Building Operator A&B', 'Chiropractors',
     'Composting Facility Operator', 'Driver Examiner', 'Fireman (FIR)', 'Home Inspector', 'Hunting & Fishing Guides',
     'Land Surveyors', 'Local Government Manager', 'Locksmiths', 'Occupational Therapists', 'Optometrists',
     'Podiatrists', 'Private Investigators', 'Security Workers'])

df2016 = df2016.drop(df2016[(df2016['Occupation Title'].isin(Excluded_Occu_2016))].index)

# Dropping Excluded Occupations - 2015

Excluded_Occu_2015 = list(
    ['Asbestos Worker', 'Automotive Salesperson', 'Biologist', 'Groom', 'Horse Racing occupations',
     'Private Investigators', 'Security Guard', 'Locksmiths', 'Occupational Therapists',
     'Information Systems Professional', 'Shorthand Reporter', 'Water and/or Wastewater Operator'])

df2015 = df2015.drop(df2015[(df2015['Occupation Title'].isin(Excluded_Occu_2015))].index)

# Dropping Excluded Occupations - 2014'

Excluded_Occu_2014 = list(['Fireman (FIR)'])

df2014 = df2014.drop(df2014[(df2014['Occupation Title'].isin(Excluded_Occu_2014))].index)

# Dropping Excluded Occupations - 2013

Excluded_Occu_2013 = list(
    ['Chartered Accountants', 'Chiropractors', 'Driver Examiner', 'Home Economists', 'Home Inspectors',
     'Horse Jockeys/ Horse Racing Standard bred Drivers', 'Hunting & Fishing Guides', 'Landscape Architect',
     'Occupational Therapists', 'Podiatrists', 'Physicians and Surgeons', 'Respiratory Therapists'])

df2013 = df2013.drop(df2013[(df2013['Occupation Title'].isin(Excluded_Occu_2013))].index)


# Changing Occupation Group for some Occupations

def change_occu(df, ot, og):
    df.loc[df['Occupation Title'].isin(ot), 'Occupational Group '] = og

    return df


# Changing Occupational Group for some Occupations - 2017
df2017 = change_occu(df2017, ['Asbestos Worker', 'Home Economist/ Human Ecologist'],
                     "Engineering, Architecture, Science and Technology")

df2017 = change_occu(df2017, ['Vehicle Inspection Technician'], "Legal, Education and Government")

# Dropping Other Occupation Group from 2017 data
df2017 = df2017.drop(df2017[df2017['Occupational Group '] == 'Other'].index)

# Changing Occupational Group for some Occupations - 2013
df2013 = change_occu(df2013, ['Locksmith'], "Engineering, Architecture, Science and Technology")

# Dropping Other Occupation Group from 2017 data
df2013 = df2013.drop(df2013[df2013['Occupational Group '] == 'Other'].index)


# Occupational Groups Summary

def func_OG(odf, sumdata):
    sumdata = odf.groupby('Occupational Group ')[
        'Applications received from Albertans', 'Out-of-province Applications received '].sum()

    sumdata['Total Applications received'] = sumdata['Applications received from Albertans'] + sumdata[
        'Out-of-province Applications received ']

    sumdata.loc['Total'] = sumdata.sum()

    sumdata['% Alberta'] = sumdata['Applications received from Albertans'] / sumdata['Total Applications received']

    sumdata['% Out-of-province'] = sumdata['Out-of-province Applications received '] / sumdata[
        'Total Applications received']

    # Rearrange columns
    sumdata = sumdata[['Applications received from Albertans', '% Alberta', 'Out-of-province Applications received ',
                       '% Out-of-province', 'Total Applications received']]

    # Format columns
    #    sumdata[['% Alberta','% Out-of-province']] = sumdata[['% Alberta','% Out-of-province']].applymap(lambda x: "{0:.0f}%".format(x*100))

    #    sumdata[['Applications received from Albertans','Out-of-province Applications received ','Total Applications received']] = sumdata[['Applications received from Albertans','Out-of-province Applications received ','Total Applications received']].applymap(lambda x: "{:,.0f}".format(x))

    return sumdata


# Occupational Groups Summary -2017
Occu_grp_smry_2017 = []
Occu_grp_smry_2017 = func_OG(df2017, Occu_grp_smry_2017)

# Occupational Groups Summary -2016
Occu_grp_smry_2016 = []
Occu_grp_smry_2016 = func_OG(df2016, Occu_grp_smry_2016)

# Occupational Groups Summary -2015
Occu_grp_smry_2015 = []
Occu_grp_smry_2015 = func_OG(df2015, Occu_grp_smry_2015)

# Occupational Groups Summary -2014
Occu_grp_smry_2014 = []
Occu_grp_smry_2014 = func_OG(df2014, Occu_grp_smry_2014)

# Occupational Groups Summary -2013
Occu_grp_smry_2013 = []
Occu_grp_smry_2013 = func_OG(df2013, Occu_grp_smry_2013)

# Occupational Groups Summary -2012
Occu_grp_smry_2012 = []
Occu_grp_smry_2012 = func_OG(df2012, Occu_grp_smry_2012)

# Comparison of Applications 2012-2017
df_all = pd.concat([df2017, df2016, df2015, df2014, df2013, df2012])

Applications = pd.DataFrame(
    df_all.groupby('Year')['Applications received from Albertans', 'Out-of-province Applications received '].sum()).T

Applications.index.name = ''
Applications.loc['Total Applications received'] = Applications.sum()

Applications.loc['% Alberta'] = Applications.loc['Applications received from Albertans'] / Applications.loc[
    'Total Applications received']

Applications.loc['% Out-of-province'] = Applications.loc['Out-of-province Applications received '] / Applications.loc[
    'Total Applications received']

# Rearrange rows
Applications = Applications.reindex(index=(
    'Applications received from Albertans', '% Alberta', 'Out-of-province Applications received ', '% Out-of-province',
    'Total Applications received'))

# Format rows
# Applications.loc[['% Alberta','% Out-of-province']] = Applications.loc[['% Alberta','% Out-of-province']].applymap(lambda x: "{0:.0f}%".format(x*100))

# Applications.loc[['Applications received from Albertans','Out-of-province Applications received ','Total Applications received']] = Applications.loc[['Applications received from Albertans','Out-of-province Applications received ','Total Applications received']].applymap(lambda x: "{:,.0f}".format(x))

# Application by Occupational Group and Year
Applications_OG_AB = pd.DataFrame(
    df_all.groupby(['Year', 'Occupational Group '])['Applications received from Albertans'].sum()).unstack('Year',
                                                                                                           fill_value=0)

Applications_OG_AB.index.name = ''

Applications_OG_OP = pd.DataFrame(
    df_all.groupby(['Year', 'Occupational Group '])['Out-of-province Applications received '].sum()).unstack('Year',
                                                                                                             fill_value=0)

Applications_OG_OP.index.name = ' '


# Labour Mobility Rate - Occupational Groups
def func_LMR(odf, sumdata, OG):
    odf = odf[odf['Occupational Group '] == OG]

    sumdata = odf.groupby('Occupation Title')[
        'Applications received from Albertans', 'Out-of-province Applications received '].sum()

    sumdata['Total Applications received'] = sumdata['Applications received from Albertans'] + sumdata[
        'Out-of-province Applications received ']

    sumdata['Labour Mobility Rate'] = sumdata['Out-of-province Applications received '] / sumdata[
        'Total Applications received']

    # Drop NAN values
    sumdata = sumdata.dropna()

    # Selecting Top 5 Occupations

    sumdata = sumdata[['Labour Mobility Rate']].nlargest(5, 'Labour Mobility Rate')

    # Drop if total application is zero
    sumdata = sumdata.loc[(sumdata['Labour Mobility Rate'] != 0)]

    # Format Labour Mobility Rate
    #    sumdata[['Labour Mobility Rate']] =sumdata[['Labour Mobility Rate']].applymap(lambda x: "{0:.0f}%".format(x*100))

    return sumdata


# 2017 Labour Mobility Rate - Business, Finance and Real Estate
B_Occupational_summary_2017 = []
B_Occupational_summary_2017 = func_LMR(df2017, B_Occupational_summary_2017, 'Business, Finance and Real Estate')

# 2017 Labour Mobility Rate - Engineering, Architecture, Science and Technology
E_Occupational_summary_2017 = []
E_Occupational_summary_2017 = func_LMR(df2017, E_Occupational_summary_2017,
                                       'Engineering, Architecture, Science and Technology')

# 2017 Labour Mobility Rate - Health and Social Services
H_Occupational_summary_2017 = []
H_Occupational_summary_2017 = func_LMR(df2017, H_Occupational_summary_2017, 'Health and Social Services')

# 2017 Labour Mobility Rate - Legal, Education and Government
L_Occupational_summary_2017 = []
L_Occupational_summary_2017 = func_LMR(df2017, L_Occupational_summary_2017, 'Legal, Education and Government')

# 2016 Labour Mobility Rate - Business, Finance and Real Estate
B_Occupational_summary_2016 = []
B_Occupational_summary_2016 = func_LMR(df2016, B_Occupational_summary_2016, 'Business, Finance and Real Estate')

# 2016 Labour Mobility Rate - Engineering, Architecture, Science and Technology
E_Occupational_summary_2016 = []
E_Occupational_summary_2016 = func_LMR(df2016, E_Occupational_summary_2016,
                                       'Engineering, Architecture, Science and Technology')

# 2016 Labour Mobility Rate - Health and Social Services
H_Occupational_summary_2016 = []
H_Occupational_summary_2016 = func_LMR(df2016, H_Occupational_summary_2016, 'Health and Social Services')

# 2016 Labour Mobility Rate - Legal, Education and Government
L_Occupational_summary_2016 = []
L_Occupational_summary_2016 = func_LMR(df2016, L_Occupational_summary_2016, 'Legal, Education and Government')

# 2015 Labour Mobility Rate - Business, Finance and Real Estate
B_Occupational_summary_2015 = []
B_Occupational_summary_2015 = func_LMR(df2015, B_Occupational_summary_2015, 'Business, Finance and Real Estate')

# 2015 Labour Mobility Rate - Engineering, Architecture, Science and Technology
E_Occupational_summary_2015 = []
E_Occupational_summary_2015 = func_LMR(df2015, E_Occupational_summary_2015,
                                       'Engineering, Architecture, Science and Technology')

# 2015 Labour Mobility Rate - Health and Social Services
H_Occupational_summary_2015 = []
H_Occupational_summary_2015 = func_LMR(df2015, H_Occupational_summary_2015, 'Health and Social Services')

# 2015 Labour Mobility Rate - Legal, Education and Government
L_Occupational_summary_2015 = []
L_Occupational_summary_2015 = func_LMR(df2015, L_Occupational_summary_2015, 'Legal, Education and Government')

# 2014 Labour Mobility Rate - Business, Finance and Real Estate
B_Occupational_summary_2014 = []
B_Occupational_summary_2014 = func_LMR(df2014, B_Occupational_summary_2014, 'Business, Finance and Real Estate')

# 2014 Labour Mobility Rate - Engineering, Architecture, Science and Technology
E_Occupational_summary_2014 = []
E_Occupational_summary_2014 = func_LMR(df2014, E_Occupational_summary_2014,
                                       'Engineering, Architecture, Science and Technology')

# 2014 Labour Mobility Rate - Health and Social Services
H_Occupational_summary_2014 = []
H_Occupational_summary_2014 = func_LMR(df2014, H_Occupational_summary_2014, 'Health and Social Services')

# 2014 Labour Mobility Rate - Legal, Education and Government
L_Occupational_summary_2014 = []
L_Occupational_summary_2014 = func_LMR(df2014, L_Occupational_summary_2014, 'Legal, Education and Government')

# 2013 Labour Mobility Rate - Business, Finance and Real Estate
B_Occupational_summary_2013 = []
B_Occupational_summary_2013 = func_LMR(df2013, B_Occupational_summary_2013, 'Business, Finance and Real Estate')

# 2013 Labour Mobility Rate - Engineering, Architecture, Science and Technology
E_Occupational_summary_2013 = []
E_Occupational_summary_2013 = func_LMR(df2013, E_Occupational_summary_2013,
                                       'Engineering, Architecture, Science and Technology')

# 2013 Labour Mobility Rate - Health and Social Services
H_Occupational_summary_2013 = []
H_Occupational_summary_2013 = func_LMR(df2013, H_Occupational_summary_2013, 'Health and Social Services')

# 2013 Labour Mobility Rate - Legal, Education and Government
L_Occupational_summary_2013 = []
L_Occupational_summary_2013 = func_LMR(df2013, L_Occupational_summary_2013, 'Legal, Education and Government')

# 2012 Labour Mobility Rate - Business, Finance and Real Estate
B_Occupational_summary_2012 = []
B_Occupational_summary_2012 = func_LMR(df2012, B_Occupational_summary_2012, 'Business, Finance and Real Estate')

# 2012 Labour Mobility Rate - Engineering, Architecture, Science and Technology
E_Occupational_summary_2012 = []
E_Occupational_summary_2012 = func_LMR(df2012, E_Occupational_summary_2012,
                                       'Engineering, Architecture, Science and Technology')

# 2012 Labour Mobility Rate - Health and Social Services
H_Occupational_summary_2012 = []
H_Occupational_summary_2012 = func_LMR(df2012, H_Occupational_summary_2012, 'Health and Social Services')

# 2012 Labour Mobility Rate - Legal, Education and Government
L_Occupational_summary_2012 = []
L_Occupational_summary_2012 = func_LMR(df2012, L_Occupational_summary_2012, 'Legal, Education and Government')


# Labour Mobility Rate

def func_LMR_1(sumdata, Grp):
    sumdata = pd.DataFrame(
        df_all.groupby(Grp)['Applications received from Albertans', 'Out-of-province Applications received '].sum())

    sumdata['Total Applications received'] = sumdata['Applications received from Albertans'] + sumdata[
        'Out-of-province Applications received ']

    sumdata['Labour Mobility Rate'] = sumdata['Out-of-province Applications received '] / sumdata[
        'Total Applications received']
    sumdata.T

    sumdata.index.name = ''

    return sumdata


# Labour Mobility Rate by Year
LMR_YR, LMR_YR_OG = [], []
LMR_YR = func_LMR_1(LMR_YR, 'Year')

# Labour Mobility Rate by Year and Occupational Group
LMR_YR_OG = func_LMR_1(LMR_YR_OG, ['Year', 'Occupational Group '])
LMR_YR_OG = pd.pivot_table(LMR_YR_OG, values='Labour Mobility Rate', index=['Occupational Group '], columns=['Year'],
                           aggfunc=np.sum)


# Regulated Occupations Reporting Highest Number of Out-of-Province Applicants

def func_OG_T(odf, sumdata):
    # Selecting Top 10 Out-of-province Occupations
    sumdata = odf.groupby('Occupation Title')[
        'Occupation Title', 'Out-of-province Applications received '].sum().nlargest(10,
                                                                                     'Out-of-province Applications received ')

    sumdata.loc['Total'] = sumdata.sum()

    # Format Out-of-province Applications received columns
    #    sumdata[['Out-of-province Applications received ']] = sumdata[['Out-of-province Applications received ']].applymap(lambda x: "{:,.0f}".format(x))

    return sumdata


# 2017 Top 10 Out-of-Province Applicants Regulated Occupations

Top_Occu_2017 = []
Top_Occu_2017 = func_OG_T(df2017, Top_Occu_2017)

# 2016 Top 10 Out-of-Province Applicants Regulated Occupations

Top_Occu_2016 = []
Top_Occu_2016 = func_OG_T(df2016, Top_Occu_2016)

# 2015 Top 10 Out-of-Province Applicants Regulated Occupations

Top_Occu_2015 = []
Top_Occu_2015 = func_OG_T(df2015, Top_Occu_2015)

# 2014 Top 10 Out-of-Province Applicants Regulated Occupations

Top_Occu_2014 = []
Top_Occu_2014 = func_OG_T(df2014, Top_Occu_2014)

# 2013 Top 10 Out-of-Province Applicants Regulated Occupations

Top_Occu_2013 = []
Top_Occu_2013 = func_OG_T(df2013, Top_Occu_2013)

# 2012 Top 10 Out-of-Province Applicants Regulated Occupations

Top_Occu_2012 = []
Top_Occu_2012 = func_OG_T(df2012, Top_Occu_2012)


# Occupational Groups Summary

def func_OG_All(odf, sumdata, sumdata1):
    sumdata = odf.groupby(['Occupational Group ', 'Occupation Title'])[
        'Applications received from Albertans', 'Out-of-province Applications received '].sum()

    sumdata1 = odf.groupby(['Occupational Group '])[
        'Applications received from Albertans', 'Out-of-province Applications received '].sum()

    sumdata1.index = [sumdata1.index.get_level_values(0),
                      ['Total'] * len(sumdata1)]

    sumdata = pd.concat([sumdata, sumdata1]).sort_index(level=[0])

    sumdata['Total Applications received'] = sumdata['Applications received from Albertans'] + sumdata[
        'Out-of-province Applications received ']

    sumdata = sumdata.loc[(sumdata['Total Applications received'] != 0)]

    sumdata = sumdata.sort_values(['Occupational Group ', 'Total Applications received'], ascending=[True, False])

    sumdata['% Alberta'] = sumdata['Applications received from Albertans'] / sumdata['Total Applications received']

    sumdata['% Out-of-province'] = sumdata['Out-of-province Applications received '] / sumdata[
        'Total Applications received']

    # Rearrange columns
    sumdata = sumdata[['Applications received from Albertans', '% Alberta', 'Out-of-province Applications received ',
                       '% Out-of-province', 'Total Applications received']]

    # Format columns
    #    sumdata[['% Alberta','% Out-of-province']] = sumdata[['% Alberta','% Out-of-province']].applymap(lambda x: "{0:.0f}%".format(x*100))

    #    sumdata[['Applications received from Albertans','Out-of-province Applications received ','Total Applications received']] = sumdata[['Applications received from Albertans','Out-of-province Applications received ','Total Applications received']].applymap(lambda x: "{:,.0f}".format(x))

    return sumdata


# Occupational Groups Summary -2017
All_Occu_smry_2017 = []
All_Occu_smry_2017a = []
All_Occu_smry_2017 = func_OG_All(df2017, All_Occu_smry_2017, All_Occu_smry_2017a)

# Occupational Groups Summary -2016
All_Occu_smry_2016 = []
All_Occu_smry_2016a = []
All_Occu_smry_2016 = func_OG_All(df2016, All_Occu_smry_2016, All_Occu_smry_2016a)

# Occupational Groups Summary -2015
All_Occu_smry_2015 = []
All_Occu_smry_2015a = []
All_Occu_smry_2015 = func_OG_All(df2015, All_Occu_smry_2015, All_Occu_smry_2015a)

# Occupational Groups Summary -2014
All_Occu_smry_2014 = []
All_Occu_smry_2014a = []
All_Occu_smry_2014 = func_OG_All(df2014, All_Occu_smry_2014, All_Occu_smry_2014a)

# Occupational Groups Summary -2013
All_Occu_smry_2013 = []
All_Occu_smry_2013a = []
All_Occu_smry_2013 = func_OG_All(df2013, All_Occu_smry_2013, All_Occu_smry_2013a)

# Occupational Groups Summary -2012
All_Occu_smry_2012 = []
All_Occu_smry_2012a = []
All_Occu_smry_2012 = func_OG_All(df2012, All_Occu_smry_2012, All_Occu_smry_2012a)

# Application Processing Time

from numpy import mean
from numpy import std


def func_Outliers(odf):
    # Calculating Mean and STD for Processing Time

    ABmean_var, ABstd_var = mean(odf['Processing Time for Alberta Applications']), std(
        odf['Processing Time for Alberta Applications'])

    OPmean_var, OPstd_var = mean(odf['Processing Time for out-of-province Applications']), std(
        odf['Processing Time for out-of-province Applications'])

    ABcut_off = ABstd_var * 3

    OPcut_off = OPstd_var * 3

    ABlower, ABupper = ABmean_var - ABcut_off, ABmean_var + ABcut_off

    OPlower, OPupper = OPmean_var - OPcut_off, OPmean_var + OPcut_off

    # Identifying Outliers that are 3 std away from the mean process time
    ABoutliers = [x for x in odf['Processing Time for Alberta Applications'] if x < ABlower or x > ABupper]
    print(' AB Identified outliers: %d' % len(ABoutliers))

    OPoutliers = [x for x in odf['Processing Time for out-of-province Applications'] if x < OPlower or x > OPupper]
    print('OP Identified outliers: %d' % len(OPoutliers))

    # Replaceing outliers with NAN (missing values)
    odf.loc[odf['Processing Time for Alberta Applications'].isin(
        ABoutliers), 'Processing Time for Alberta Applications'] = np.nan

    odf.loc[odf['Processing Time for out-of-province Applications'].isin(
        OPoutliers), 'Processing Time for out-of-province Applications'] = np.nan

    return odf


# Replacing outliers with NAN (missing values) =2017
df2017 = func_Outliers(df2017)

# Replacing outliers with NAN (missing values) =2016
df2016 = func_Outliers(df2016)

# Replacing outliers with NAN (missing values) =2015
df2015 = func_Outliers(df2015)

# Replacing outliers with NAN (missing values) =2014
df2014 = func_Outliers(df2014)

# Replacing outliers with NAN (missing values) =2013
df2013 = func_Outliers(df2013)

# Replacing outliers with NAN (missing values) =2012
df2012 = func_Outliers(df2012)

All_PT = pd.concat([df2017, df2016, df2015, df2014, df2013, df2012])

#All_PT.to_excel("//GOA/MyDocs/I/isaac.nyamekye/Projects/Labour Mobility/LMS_All_Year.xlsx", index=False)

# Comparison of Processing Time 2012-2017
Processing_T = pd.DataFrame(All_PT.groupby('Year')[
                                'Processing Time for Alberta Applications', 'Processing Time for out-of-province Applications'].mean()).T
Processing_T.index.name = ' '

# Processing Time by Occupational Group
Processing_T_OG_AB = pd.DataFrame(
    All_PT.groupby(['Year', 'Occupational Group '])['Processing Time for Alberta Applications'].mean()).unstack('Year',
                                                                                                                fill_value=0)

Processing_T_OG_AB.index.name = ' '

Processing_T_OG_OP = pd.DataFrame(
    All_PT.groupby(['Year', 'Occupational Group '])['Processing Time for out-of-province Applications'].mean()).unstack(
    'Year', fill_value=0)

Processing_T_OG_OP.index.name = ' '

# Processing Time by Occupational Group and Title
# Pivot table for Processing time by year, occupation group and title
Processing_T_OT_AB = pd.pivot_table(All_PT, values=['Processing Time for Alberta Applications'],
                                    index=['Occupational Group ', 'Occupation Title'], columns=['Year'],
                                    aggfunc=np.mean)
Processing_T_OT_AB = Processing_T_OT_AB.reindex(
    Processing_T_OT_AB['Processing Time for Alberta Applications'].sort_values(
        by=['Occupational Group ', 2017]).index)  # Sorting by 2017 values

# Pivot table for Processing time by year, occupation group and title
Processing_T_OT_OP = pd.pivot_table(All_PT, values=['Processing Time for out-of-province Applications'],
                                    index=['Occupational Group ', 'Occupation Title'], columns=['Year'],
                                    aggfunc=np.mean)
Processing_T_OT_OP = Processing_T_OT_OP.reindex(
    Processing_T_OT_OP['Processing Time for out-of-province Applications'].sort_values(
        by=['Occupational Group ', 2017]).index)  # Sorting by 2017 values

"""
def output(out, df1, df2, df3, df4, df5, df6, df7):
    # Creating Excel Workbook for tables and charts
    writer = pd.ExcelWriter(r"//GOA/MyDocs/I/isaac.nyamekye/Projects/Labour Mobility/" + out + ".xlsx",
                            engine='xlsxwriter')  # Creating Excel Writer Object from Pandas
    workbook = writer.book
    worksheet = workbook.add_worksheet('Occupational Groups Summary')
    writer.sheets['Occupational Groups Summary'] = worksheet

    worksheet1 = workbook.add_worksheet('Labour Mobility Rate')
    writer.sheets['Labour Mobility Rate'] = worksheet1

    worksheet2 = workbook.add_worksheet('Processing Time Summary')
    writer.sheets['Processing Time Summary'] = worksheet2

    worksheet3 = workbook.add_worksheet('All Occupations')
    writer.sheets['All Occupations'] = worksheet3

    worksheet4 = workbook.add_worksheet('Processing - All Occupations')
    writer.sheets['Processing - All Occupations'] = worksheet4

    merge_format = workbook.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': 'blue',
        'font_color': 'white'})

    format1 = workbook.add_format({'num_format': '0%'})

    format2 = workbook.add_format({'num_format': '#,##0'})

    worksheet.merge_range('B2:G3',
                          'Alberta and Out-of-Province Applicants Entering a Regulated Occupation by Occupational Groups',
                          merge_format)

    worksheet.add_table('B5:G9')

    worksheet.set_column(1, 1, 45)
    worksheet.set_column(2, 7, 28)

    df1.to_excel(writer, sheet_name='Occupational Groups Summary', startrow=4, startcol=1)

    worksheet.merge_range('B13:H14',
                          'Comparison of Number of Alberta and Out-of-Province Certification/Licensure Applications Between 2012-2017',
                          merge_format)

    worksheet.add_table('B16:H21')

    Applications.to_excel(writer, sheet_name='Occupational Groups Summary', startrow=15, startcol=1)

    worksheet.merge_range('B24:C25', 'Regulated Occupations with Highest Number of Out-of-Province Applicants',
                          merge_format)

    worksheet.add_table('B27:C38')

    df2.to_excel(writer, sheet_name='Occupational Groups Summary', startrow=26, startcol=1)

    worksheet1.merge_range('B2:G3', 'Labour Mobility Rates within Occupational Groups', merge_format)

    worksheet1.merge_range('B6:G6', 'Business, Finance and Real Estate', merge_format)

    worksheet1.add_table('B8:C13')

    worksheet1.set_column('C:C', 25, format1)

    worksheet1.set_column(1, 1, 45)

    df3.to_excel(writer, sheet_name='Labour Mobility Rate', startrow=7, startcol=1)

    B_Chart = workbook.add_chart({"type": "bar"})

    B_Chart.add_series({

        "values": "='Labour Mobility Rate'!C8:C13",
        "categories": "='Labour Mobility Rate'!B8:B13",
        "data_labels": {"value": True},
        "gap": 50,
    })

    B_Chart.set_x_axis({'visible': False,
                        'major_gridlines': {'visible': False}})

    B_Chart.set_y_axis({'major_tick_mark': 'none',
                        'reverse': True,
                        'major_gridlines': {'visible': False}})

    B_Chart.set_title({'name': 'Business, Finance and Real Estate',
                       'font': {'size': 8, 'bold': True}})

    B_Chart.set_legend({'none': True})

    worksheet1.insert_chart("E8", B_Chart)

    # Engineering, Architure, Science and Technology table and chart
    worksheet1.merge_range('B24:G24', 'Engineering, Architure, Science and Technology', merge_format)

    worksheet1.add_table('B26:C31')

    df4.to_excel(writer, sheet_name='Labour Mobility Rate', startrow=25, startcol=1)

    E_Chart = workbook.add_chart({"type": "bar"})

    E_Chart.add_series({

        "values": "='Labour Mobility Rate'!C26:C31",
        "categories": "='Labour Mobility Rate'!B26:B31",
        "data_labels": {"value": True},
        "gap": 50,
    })

    E_Chart.set_x_axis({'visible': False,
                        'major_gridlines': {'visible': False}})

    E_Chart.set_y_axis({'major_tick_mark': 'none',
                        'reverse': True,
                        'major_gridlines': {'visible': False}})

    E_Chart.set_title({'name': 'Engineering, Architure, Science and Technology',
                       'font': {'size': 8, 'bold': True}})

    E_Chart.set_legend({'none': True})

    worksheet1.insert_chart("E26", E_Chart)

    # Health and Social Services table and chart
    worksheet1.merge_range('B42:G42', 'Health and Social Services', merge_format)

    worksheet1.add_table('B44:C49')

    df5.to_excel(writer, sheet_name='Labour Mobility Rate', startrow=43, startcol=1)

    H_Chart = workbook.add_chart({"type": "bar"})

    H_Chart.add_series({

        "values": "='Labour Mobility Rate'!C44:C49",
        "categories": "='Labour Mobility Rate'!B44:B49",
        "data_labels": {"value": True},
        "gap": 50,
    })

    H_Chart.set_x_axis({'visible': False,
                        'major_gridlines': {'visible': False}})

    H_Chart.set_y_axis({'major_tick_mark': 'none',
                        'reverse': True,
                        'major_gridlines': {'visible': False}})

    H_Chart.set_title({'name': 'Health and Social Services',
                       'font': {'size': 8, 'bold': True}})

    H_Chart.set_legend({'none': True})

    worksheet1.insert_chart("E44", H_Chart)

    # Legal, Education and Government table and chart
    worksheet1.merge_range('B60:G60', 'Legal, Education and Government', merge_format)

    worksheet1.add_table('B62:C67')

    df6.to_excel(writer, sheet_name='Labour Mobility Rate', startrow=61, startcol=1)

    L_Chart = workbook.add_chart({"type": "bar"})

    L_Chart.add_series({

        "values": "='Labour Mobility Rate'!C62:C67",
        "categories": "='Labour Mobility Rate'!B62:B67",
        "data_labels": {"value": True},
        "gap": 50,
    })

    L_Chart.set_x_axis({'visible': False,
                        'major_gridlines': {'visible': False}})

    L_Chart.set_y_axis({'major_tick_mark': 'none',
                        'reverse': True,
                        'major_gridlines': {'visible': False}})

    L_Chart.set_title({'name': 'Legal, Education and Government',
                       'font': {'size': 8, 'bold': True}})

    L_Chart.set_legend({'none': True})

    worksheet1.insert_chart("E62", L_Chart)

    worksheet2.merge_range('B2:H3', 'Average Processing Time (2012-2017)', merge_format)

    worksheet2.add_table('B5:H7')

    worksheet2.set_column('C:J', 20, format2)

    worksheet2.set_column(1, 1, 45)

    Processing_T.to_excel(writer, sheet_name='Processing Time Summary', startrow=4, startcol=1)

    worksheet2.merge_range('B10:H11', 'Average Processing Time for Alberta Applicants by Occupation Group (2012-2017)',
                           merge_format)

    Processing_T_OG_AB.to_excel(writer, sheet_name='Processing Time Summary', startrow=12, startcol=1)

    worksheet2.merge_range('B22:H23',
                           'Average Processing Time for Out-of-Province Applicants by Occupation Group (2012-2017)',
                           merge_format)

    Processing_T_OG_OP.to_excel(writer, sheet_name='Processing Time Summary', startrow=24, startcol=1)

    worksheet3.merge_range('B2:H3',
                           'Alberta and Out-of-Province Applicants Entering a Regulated Occupation by Occupational Groups',
                           merge_format)

    worksheet3.add_table('B5:H112')

    worksheet3.set_column(1, 1, 44)
    worksheet3.set_column(2, 2, 44)
    worksheet3.set_column(3, 7, 26)
    worksheet3.set_column('D:D', 26, format2)
    worksheet3.set_column('E:E', 26, format1)
    worksheet3.set_column('F:F', 26, format2)
    worksheet3.set_column('G:G', 26, format1)
    worksheet3.set_column('H:H', 26, format2)

    df7.to_excel(writer, sheet_name='All Occupations', startrow=4, startcol=1)

    worksheet4.merge_range('B2:H3', 'Average Processing Time for Alberta Applicants by Occupation Group (2012-2017)',
                           merge_format)

    worksheet4.set_column('D:H', 8, format2)

    worksheet4.set_column(1, 2, 30)

    Processing_T_OT_AB.to_excel(writer, sheet_name='Processing - All Occupations', startrow=4, startcol=1)

    worksheet4.merge_range('J2:P3',
                           'Average Processing Time for Out-of-Province Applicants by Occupation Group (2012-2017)',
                           merge_format)

    worksheet4.set_column('M:P', 8, format2)

    worksheet4.set_column('K:J', 30)

    Processing_T_OT_OP.to_excel(writer, sheet_name='Processing - All Occupations', startrow=4, startcol=9)

    writer.save()

    return df1, df2, df3, df4, df5, df6, df7


# Creating Excel Workbook for tables and charts -2017
LMS_2017 = output('LMS_2017', Occu_grp_smry_2017, Top_Occu_2017, B_Occupational_summary_2017,
                  E_Occupational_summary_2017, H_Occupational_summary_2017, L_Occupational_summary_2017,
                  All_Occu_smry_2017)

# Creating Excel Workbook for tables and charts -2016
LMS_2016 = output('LMS_2016', Occu_grp_smry_2016, Top_Occu_2016, B_Occupational_summary_2016,
                  E_Occupational_summary_2016, H_Occupational_summary_2016, L_Occupational_summary_2016,
                  All_Occu_smry_2016)

# Creating Excel Workbook for tables and charts -2015
LMS_2015 = output('LMS_2015', Occu_grp_smry_2015, Top_Occu_2015, B_Occupational_summary_2015,
                  E_Occupational_summary_2015, H_Occupational_summary_2015, L_Occupational_summary_2015,
                  All_Occu_smry_2015)

# Creating Excel Workbook for tables and charts -2014
LMS_2014 = output('LMS_2014', Occu_grp_smry_2014, Top_Occu_2014, B_Occupational_summary_2014,
                  E_Occupational_summary_2014, H_Occupational_summary_2014, L_Occupational_summary_2014,
                  All_Occu_smry_2014)

# Creating Excel Workbook for tables and charts -2013
LMS_2013 = output('LMS_2013', Occu_grp_smry_2013, Top_Occu_2013, B_Occupational_summary_2013,
                  E_Occupational_summary_2013, H_Occupational_summary_2013, L_Occupational_summary_2013,
                  All_Occu_smry_2013)

# Creating Excel Workbook for tables and charts -2012
LMS_2012 = output('LMS_2012', Occu_grp_smry_2012, Top_Occu_2012, B_Occupational_summary_2012,
                  E_Occupational_summary_2012, H_Occupational_summary_2012, L_Occupational_summary_2012,
                  All_Occu_smry_2012)
"""

from openpyxl import load_workbook

# from openpyxl.utils.dataframe import dataframe_to_rows

# Load Workbook
wb = load_workbook("M:/WS/Program Effectiveness/Program Analytics/DDRP/Monthly_Report/2019-20/_Latest_Month/DDRP - Monthly_Report.xlsx")

xl_writer = pd.ExcelWriter("M:/WS/Program Effectiveness/Program Analytics/DDRP/Monthly_Report/2019-20/_Latest_Month/DDRP - Monthly_Report.xlsx",
                           engine='openpyxl')

xl_writer.book = wb
# Read all sheets in workbook
xl_writer.sheets = dict((ws.title, ws) for ws in wb.worksheets)

# Export dataframes to excel

Applications.to_excel(xl_writer, 'LMS', index=True, startcol=1, startrow=3)

Applications_OG_AB.to_excel(xl_writer, 'LMS', index=True, startcol=1, startrow=12)

Applications_OG_OP.to_excel(xl_writer, 'LMS', index=True, startcol=1, startrow=23)

Processing_T.to_excel(xl_writer, 'LMS', index=True, startcol=1, startrow=34)

Processing_T_OG_AB.to_excel(xl_writer, 'LMS', index=True, startcol=1, startrow=40)

Processing_T_OG_OP.to_excel(xl_writer, 'LMS', index=True, startcol=1, startrow=51)

LMR_YR.to_excel(xl_writer, 'LMS', index=True, startcol=1, startrow=62)

LMR_YR_OG.to_excel(xl_writer, 'LMS', index=True, startcol=1, startrow=73)

# Save Workbook
xl_writer.save()

# Close Workbook
wb.close()
