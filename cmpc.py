"""
Module Docstring
"""

__author__ = "MiKe Howard"
__version__ = "0.1.0"
__license__ = "MIT"

import logging
import pandas as pd
import glob
import os
from datetime import date

# from xlsxwriter.utility import xl_rowcol_to_cell
from dateutil.tz import tzutc
import numpy as np


# OS Functions
def filesearch(word=""):
    """Returns a list with all files with the word/extension in it"""
    logger.info('Starting filesearch')
    file = []
    for f in glob.glob("*"):
        if word[0] == ".":
            if f.endswith(word):
                file.append(f)

        elif word in f:
            file.append(f)
            # return file
    logger.debug(file)
    return file


def Change_Working_Path(path):
    # Check if New path exists
    if os.path.exists(path):
        # Change the current working Directory
        try:
            os.chdir(path)  # Change the working directory
        except OSError:
            logger.error("Can't change the Current Working Directory", exc_info=True)
    else:
        print("Can't change the Current Working Directory because this path doesn't exits")


# Pandas Functions
def Excel_to_Pandas(filename):
    logger.info('importing file ' + filename)
    df = []
    try:
        df = pd.read_excel(filename)
    except:
        logger.error("Error importing file " + filename, exc_info=True)

    df = Cleanup_Dataframe(df)
    logger.debug(df.info(verbose=True))
    return df


def Cleanup_Dataframe(df):
    logger.info('Starting Cleanup_Dataframe')
    logger.debug(df.info(verbose=True))
    # Remove whitespace on both ends of column headers
    df.columns = df.columns.str.strip()

    # Replace whitespace in column header with _
    df.columns = df.columns.str.replace(' ', '_')

    return df

def output_Merto_West_Data_current_year(current_year_tuple):
    strlist = ['Hold',
               'In Scoping',
               'In-Service',
               'Released']

    print(str(date.today().year) + ':')
    print(' '.join([str(len(current_year_tuple[0].index)), strlist[3]]))
    print(' '.join([str(len(current_year_tuple[1].index)), strlist[2]]))
    print(' '.join([str(len(current_year_tuple[2].index)), strlist[1]]))
    print(' '.join([str(len(current_year_tuple[3].index)), strlist[0]]))
    print('')

def output_Merto_West_Data_next_year(next_year_tuple):
    strlist = ['Hold',
               'In Scoping',
               'In-Service',
               'Released']

    print(str(date.today().year + 1) + ':')
    print(' '.join([str(len(next_year_tuple[0].index)), strlist[3]]))
    print(' '.join([str(len(next_year_tuple[1].index)), strlist[2]]))
    print(' '.join([str(len(next_year_tuple[2].index)), strlist[1]]))
    print(' '.join([str(len(next_year_tuple[3].index)), strlist[0]]))

def query_Merto_West_Data_current_year(Project_Data_df):
    currentyeardreleasedf = Project_Data_df[
        (Project_Data_df['Estimated_In-Service_Date'].dt.year == date.today().year) &
        (Project_Data_df['PROJECTSTATUS'] == 'Released') &
        (Project_Data_df['REGIONNAME'] == 'METRO WEST') &
        (Project_Data_df['PROJECTCATEGORY'] == 'CMPC')]

    currentyeardInServicedf = Project_Data_df[
        (Project_Data_df['Estimated_In-Service_Date'].dt.year == date.today().year) &
        (Project_Data_df['PROJECTSTATUS'] == 'In-Service') &
        (Project_Data_df['REGIONNAME'] == 'METRO WEST') &
        (Project_Data_df['PROJECTCATEGORY'] == 'CMPC')]

    currentyeardInScopingdf = Project_Data_df[
        (Project_Data_df['Estimated_In-Service_Date'].dt.year == date.today().year) &
        (Project_Data_df['PROJECTSTATUS'] == 'In Scoping') &
        (Project_Data_df['REGIONNAME'] == 'METRO WEST') &
        (Project_Data_df['PROJECTCATEGORY'] == 'CMPC')]

    currentyeardOnHolddf = Project_Data_df[
        (Project_Data_df['Estimated_In-Service_Date'].dt.year == date.today().year) &
        (Project_Data_df['PROJECTSTATUS'] == 'On Hold') &
        (Project_Data_df['REGIONNAME'] == 'METRO WEST') &
        (Project_Data_df['PROJECTCATEGORY'] == 'CMPC')]

    return currentyeardreleasedf, currentyeardInServicedf, currentyeardInScopingdf, currentyeardOnHolddf

def query_Merto_West_Data_next_year(Project_Data_df):
    nextyeardreleasedf = Project_Data_df[
        (Project_Data_df['Estimated_In-Service_Date'].dt.year == date.today().year + 1) &
        (Project_Data_df['PROJECTSTATUS'] == 'Released') &
        (Project_Data_df['REGIONNAME'] == 'METRO WEST') &
        (Project_Data_df['PROJECTCATEGORY'] == 'CMPC')]

    nextyeardInServicedf = Project_Data_df[
        (Project_Data_df['Estimated_In-Service_Date'].dt.year == date.today().year + 1) &
        (Project_Data_df['PROJECTSTATUS'] == 'In-Service') &
        (Project_Data_df['REGIONNAME'] == 'METRO WEST') &
        (Project_Data_df['PROJECTCATEGORY'] == 'CMPC')]

    nextyeardInScopingdf = Project_Data_df[
        (Project_Data_df['Estimated_In-Service_Date'].dt.year == date.today().year + 1) &
        (Project_Data_df['PROJECTSTATUS'] == 'In Scoping') &
        (Project_Data_df['REGIONNAME'] == 'METRO WEST') &
        (Project_Data_df['PROJECTCATEGORY'] == 'CMPC')]

    nextyeardOnHolddf = Project_Data_df[
        (Project_Data_df['Estimated_In-Service_Date'].dt.year == date.today().year + 1) &
        (Project_Data_df['PROJECTSTATUS'] == 'On Hold') &
        (Project_Data_df['REGIONNAME'] == 'METRO WEST') &
        (Project_Data_df['PROJECTCATEGORY'] == 'CMPC')]

    return nextyeardreleasedf, nextyeardInServicedf, nextyeardInScopingdf, nextyeardOnHolddf

def current_year_data_drive_output(current_year_tuple):
    strlist = ['Hold',
               'In Scoping',
               'In-Service',
               'Released',
               ': Number of ',
               ' Metro West CMPC Engineering projects',
               'Total number of the ',
               'Metro West CMPM Engineering projects that have approved WA > $200,000',
               'of those projects are Distribution Automation',
               'Excluding the DA projects,',
               'CMPC Engineering projects have approved WA amounts of > $200,000.',
               'of the',
               'have been placed In-Service ',
               'CMPC Engineering projects are scheduled to be completed during Quarter 4 of 2020']

    currentyeartotalproject = sum(
        [len(current_year_tuple[0].index), len(current_year_tuple[1].index), len(current_year_tuple[2].index),
         len(current_year_tuple[3].index)])
    print('')
    print(''.join([str(currentyeartotalproject), strlist[4], str(date.today().year), strlist[5]]))

    bigprojects = len(current_year_tuple[0][current_year_tuple[0]['Approved_WA_Amount'] >= 200000]) + len(
        current_year_tuple[1][current_year_tuple[1]['Approved_WA_Amount'] >= 200000])

    bigprojectsda = len(current_year_tuple[0][(current_year_tuple[0]['Approved_WA_Amount'] >= 200000) &
                                              (current_year_tuple[0]['BUDGETITEMNUMBER'] == '00003407')]) + len(
        current_year_tuple[1][(current_year_tuple[1]['Approved_WA_Amount'] >= 200000) &
                                (current_year_tuple[1]['BUDGETITEMNUMBER'] == '00003407')])

    bigprojectsnotdareleased = len(current_year_tuple[0][(current_year_tuple[0]['Approved_WA_Amount'] >= 200000) &
                                                         (current_year_tuple[0]['BUDGETITEMNUMBER'] != '00003407')])

    bigprojectsnotdainservice = len(current_year_tuple[1][(current_year_tuple[1]['Approved_WA_Amount'] >= 200000) &
                                                            (current_year_tuple[1][
                                                                 'BUDGETITEMNUMBER'] != '00003407')])

    print(' '.join([str(bigprojects), strlist[7]]))
    print(' '.join([str(bigprojectsda), strlist[8]]))

    print(' '.join([strlist[9], str(bigprojectsnotdareleased + bigprojectsnotdainservice), strlist[10]]))

    print(
        ' '.join([str(bigprojectsnotdainservice), strlist[11], str(bigprojectsnotdareleased + bigprojectsnotdainservice)
                     , strlist[12]]))

    print('')

    bigprojectsnotdareleaseddf = current_year_tuple[0][(current_year_tuple[0]['Approved_WA_Amount'] >= 200000) &
                                                       (current_year_tuple[0]['BUDGETITEMNUMBER'] != '00003407')]
    return bigprojectsnotdareleaseddf


def Programs_data_query(Project_Data_df):

    EHVbreakerreplacementscurrentyeardf = pd.DataFrame().reindex_like(Project_Data_df)
    HVbreakerreplacementscurrentyeardf = pd.DataFrame().reindex_like(Project_Data_df)
    FIDcurrentyearddf = pd.DataFrame().reindex_like(Project_Data_df)
    MDbreakerreplacementscurrentyearddf = pd.DataFrame().reindex_like(Project_Data_df)
    Linehardingcurrentyeardf = pd.DataFrame().reindex_like(Project_Data_df)
    Watercrossingscurrentyeardf =  pd.DataFrame().reindex_like(Project_Data_df)
    EHVbreakerreplacementsnextyeardf = pd.DataFrame().reindex_like(Project_Data_df)
    HVbreakerreplacementsnextyeardf = pd.DataFrame().reindex_like(Project_Data_df)
    FIDnextyearddf = pd.DataFrame().reindex_like(Project_Data_df)
    MDbreakerreplacementsnextyearddf = pd.DataFrame().reindex_like(Project_Data_df)
    Linehardingnextyeardf = pd.DataFrame().reindex_like(Project_Data_df)
    Watercrossingsnextyeardf = pd.DataFrame().reindex_like(Project_Data_df)

    budget_items_list = ['00003201', '00003202', '00003206', '00003203', '00003212', '00003226']
    Year_list = [date.today().year, date.today().year + 1]
    for bugget_item in budget_items_list:
        for year in Year_list:
            df = Project_Data_df[
                (Project_Data_df['Estimated_In-Service_Date'].dt.year == year) &
                (Project_Data_df['BUDGETITEMNUMBER'] == bugget_item)]
            if bugget_item == '00003201' and year == date.today().year:
                EHVbreakerreplacementscurrentyeardf = df
            elif bugget_item == '00003202' and year == date.today().year:
                HVbreakerreplacementscurrentyeardf = df
            elif bugget_item == '00003206' and year == date.today().year:
                FIDcurrentyearddf = df
            elif bugget_item == '00003203' and year == date.today().year:
                MDbreakerreplacementscurrentyearddf = df
            elif bugget_item == '00003212' and year == date.today().year:
                Linehardingcurrentyeardf = df
            elif bugget_item == '00003226' and year == date.today().year:
                Watercrossingscurrentyeardf = df


            elif bugget_item == '00003201' and year == date.today().year + 1:
                EHVbreakerreplacementsnextyeardf = df
            elif bugget_item == '00003202' and year == date.today().year + 1:
                HVbreakerreplacementsnextyeardf = df
            elif bugget_item == '00003206' and year == date.today().year + 1:
                FIDnextyearddf = df
            elif bugget_item == '00003203' and year == date.today().year + 1:
                MDbreakerreplacementsnextyearddf = df
            elif bugget_item == '00003212' and year == date.today().year + 1:
                Linehardingnextyeardf = df
            elif bugget_item == '00003226' and year == date.today().year + 1:
                Watercrossingsnextyeardf = df



    tuple = EHVbreakerreplacementscurrentyeardf,\
            HVbreakerreplacementscurrentyeardf,\
            FIDcurrentyearddf,\
            MDbreakerreplacementscurrentyearddf, \
            Linehardingcurrentyeardf, \
            Watercrossingscurrentyeardf,\
            EHVbreakerreplacementsnextyeardf,\
            HVbreakerreplacementsnextyeardf,\
            FIDnextyearddf,\
            MDbreakerreplacementsnextyearddf,\
            Linehardingnextyeardf,\
            Watercrossingsnextyeardf

    return tuple

def Programs_data_output(Programs_data_df):
    EHVbreakerreplacementscurrentyeardf = Programs_data_df[0]
    HVbreakerreplacementscurrentyeardf = Programs_data_df[1]
    FIDcurrentyearddf = Programs_data_df[2]
    MDbreakerreplacementscurrentyearddf = Programs_data_df[3]
    Linehardingcurrentyeardf = Programs_data_df[4]
    Watercrossingscurrentyeardf = Programs_data_df[5]
    EHVbreakerreplacementsnextyeardf = Programs_data_df[6]
    HVbreakerreplacementsnextyeardf = Programs_data_df[7]
    FIDnextyearddf = Programs_data_df[8]
    MDbreakerreplacementsnextyearddf = Programs_data_df[9]
    Linehardingnextyeardf = Programs_data_df[10]
    Watercrossingsnextyeardf = Programs_data_df[11]

    strlist = ['- 138KV Breaker Replacements',
               '- FID Replacements',
               '- 69 KV Breaker Replacements',
               'Released',
               'In-Service',
               'No Capital Spends',
               '- 345KV Breaker Replacements']
#PEC
    print(str(date.today().year) + ':')
    print(' '.join([str(len(EHVbreakerreplacementscurrentyeardf.index)), strlist[6]]))
    print(' '.join([str(len(HVbreakerreplacementscurrentyeardf.index)), strlist[0]]))
    print(' '.join([str(len(FIDcurrentyearddf.index)), strlist[1]]))
    print(' '.join([str(len(MDbreakerreplacementscurrentyearddf.index)), strlist[2]]))

    print('')
    print(str(date.today().year + 1) + ':')
    print(' '.join([str(len(EHVbreakerreplacementsnextyeardf.index)), strlist[6]]))
    print(' '.join([str(len(HVbreakerreplacementsnextyeardf.index)), strlist[0]]))
    print(' '.join([str(len(FIDnextyearddf.index)), strlist[1]]))
    print(' '.join([str(len(MDbreakerreplacementsnextyearddf.index)), strlist[2]]))
    print('')

    print(strlist[0])
    print(' '.join([str(
        len(Programs_data_df[1][HVbreakerreplacementscurrentyeardf['PROJECTSTATUS'] == 'Released'])),
                    strlist[3]]))
    print(' '.join([str(
        len(HVbreakerreplacementscurrentyeardf[HVbreakerreplacementscurrentyeardf['PROJECTSTATUS'] == 'In-Service'])),
        strlist[4]]))
    print(' '.join([str(
        len(HVbreakerreplacementscurrentyeardf[
                HVbreakerreplacementscurrentyeardf['PROJECTSTATUS'] == 'No Capital Spend'])),
        strlist[5]]))

    print('')
    print(strlist[1])
    print(' '.join([str(
        len(FIDcurrentyearddf[FIDcurrentyearddf['PROJECTSTATUS'] == 'Released'])),
        strlist[3]]))
    print(' '.join([str(
        len(FIDcurrentyearddf[FIDcurrentyearddf['PROJECTSTATUS'] == 'In-Service'])),
        strlist[4]]))
    print(' '.join([str(
        len(FIDcurrentyearddf[
                FIDcurrentyearddf['PROJECTSTATUS'] == 'No Capital Spend'])),
        strlist[5]]))

    print('')
    print(strlist[2])
    print(' '.join([str(
        len(MDbreakerreplacementscurrentyearddf[MDbreakerreplacementscurrentyearddf['PROJECTSTATUS'] == 'Released'])),
        strlist[3]]))
    print(' '.join([str(
        len(MDbreakerreplacementscurrentyearddf[MDbreakerreplacementscurrentyearddf['PROJECTSTATUS'] == 'In-Service'])),
        strlist[4]]))
    print(' '.join([str(
        len(MDbreakerreplacementscurrentyearddf[
                MDbreakerreplacementscurrentyearddf['PROJECTSTATUS'] == 'No Capital Spend'])),
        strlist[5]]))
    print('')

    print('T Line Harding')
    print(str(date.today().year) + ':')
    print(' '.join([str(
        len(Linehardingcurrentyeardf[Linehardingcurrentyeardf['PROJECTSTATUS'] == 'Released'])),
        strlist[3]]))

    print(' '.join([str(
        len(Linehardingcurrentyeardf[Linehardingcurrentyeardf['PROJECTSTATUS'] == 'Engineering Only'])),
        'Engineering Only']))

    print(' '.join([str(
        len(Linehardingcurrentyeardf[Linehardingcurrentyeardf['PROJECTSTATUS'] == 'In-Service'])),
        'In-Service']))

    print(str(date.today().year + 1) + ':')
    print(' '.join([str(
        len(Linehardingnextyeardf[Linehardingnextyeardf['PROJECTSTATUS'] == 'Released'])),
        strlist[3]]))

    print(' '.join([str(
        len(Linehardingnextyeardf[Linehardingnextyeardf['PROJECTSTATUS'] == 'Engineering Only'])),
        'Engineering Only']))

    print(' '.join([str(
        len(Linehardingnextyeardf[Linehardingnextyeardf['PROJECTSTATUS'] == 'Draft'])),
        'Draft']))

    print('')
    print('Water Crossing')
    print(str(date.today().year) + ':')
    print(' '.join([str(
        len(Watercrossingscurrentyeardf[Watercrossingscurrentyeardf['PROJECTSTATUS'] == 'Released'])),
        strlist[3]]))

    print(' '.join([str(
        len(Watercrossingscurrentyeardf[Watercrossingscurrentyeardf['PROJECTSTATUS'] == 'Engineering Only'])),
        'Engineering Only']))

    print(' '.join([str(
        len(Watercrossingscurrentyeardf[Watercrossingscurrentyeardf['PROJECTSTATUS'] == 'In-Service'])),
        'In-Service']))

    print('')

    print(str(date.today().year + 1) + ':')
    print(' '.join([str(
        len(Watercrossingsnextyeardf[Watercrossingsnextyeardf['PROJECTSTATUS'] == 'Released'])),
        strlist[3]]))

    print(' '.join([str(
        len(Watercrossingsnextyeardf[Watercrossingsnextyeardf['PROJECTSTATUS'] == 'Engineering Only'])),
        'Engineering Only']))

    print(' '.join([str(
        len(Watercrossingsnextyeardf[Watercrossingsnextyeardf['PROJECTSTATUS'] == 'Draft'])),
        'Draft']))

def main():
    PAT_Filename = 'PAT Grand Summary Report.xlsx'
    Project_Data_Filename = 'All Project Data Report Metro West or Mike.xlsx'
    Schedules_Filename = 'Metro West PETE Schedules.xlsx'
    """ Main entry point of the app """
    logger.info("Starting Pete Maintenance Helper")
    Change_Working_Path('./Data')
    try:
        Project_Data_df = Excel_to_Pandas(Project_Data_Filename)
    except:
        logger.error('Can not find Project Data file')
        raise
    try:
        PScedules_df = Excel_to_Pandas(Schedules_Filename)
    except:
        logger.error('Can not find Project Data file')
        raise

    try:
        patdf = Excel_to_Pandas(PAT_Filename)
    except:
        logger.error('Can not find Project Data file')
        raise
    # patdf= patdf['PETE_ID','WA_Amount_Grand_Summary']
    Project_Data_df = pd.merge(Project_Data_df, patdf, on='PETE_ID', how='outer')

    Project_Data_df.info()
    current_year_tuple = query_Merto_West_Data_current_year(Project_Data_df)
    next_year_tuple = query_Merto_West_Data_next_year(Project_Data_df)


    output_Merto_West_Data_current_year(current_year_tuple)
    output_Merto_West_Data_next_year(next_year_tuple)

    bigprojectsnotdareleaseddf = current_year_data_drive_output(current_year_tuple)
   # bigprojectsnotdareleaseddf.to_csv('metro_west_large_no_da_released.csv')

    #print(bigprojectsnotdareleaseddf)

    # bigprojectsnotdawithschedules = pd.merge(PScedules_df, bigprojectsnotdainservicedf , on='PETE_ID', how='inner')
   # bigprojectsnotdawPE = PScedules_df[
     #   (PScedules_df['Grandchild'] == 'Project Energization') &
     #   (PScedules_df['PETE_ID'].isin(bigprojectsnotdareleaseddf['PETE_ID']))]

   # bigprojectsnotdawPE.to_csv('metro_west_large_PE.csv')

    Programs_data_df=Programs_data_query(Project_Data_df)
    Programs_data_output(Programs_data_df)




if __name__ == "__main__":
    """ This is executed when run from the command line """
    # Setup Logging
    logger = logging.getLogger('root')
    FORMAT = "[%(filename)s:%(lineno)s - %(funcName)20s() ] %(message)s"
    logging.basicConfig(format=FORMAT)
    logger.setLevel(logging.INFO)

    main()
