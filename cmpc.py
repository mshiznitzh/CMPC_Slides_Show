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

    print(str(date.today().year) + ':')
    print(' '.join([str(len(currentyeardreleasedf.index)), strlist[3]]))
    print(' '.join([str(len(currentyeardInServicedf.index)), strlist[2]]))
    print(' '.join([str(len(currentyeardInScopingdf.index)), strlist[1]]))
    print(' '.join([str(len(currentyeardOnHolddf.index)), strlist[0]]))
    print('')
    print(str(date.today().year + 1) + ':')
    print(' '.join([str(len(nextyeardreleasedf.index)), strlist[3]]))
    print(' '.join([str(len(nextyeardInServicedf.index)), strlist[2]]))
    print(' '.join([str(len(nextyeardInScopingdf.index)), strlist[1]]))
    print(' '.join([str(len(nextyeardOnHolddf.index)), strlist[0]]))

    currentyeartotalproject = sum(
        [len(currentyeardreleasedf.index), len(currentyeardInServicedf.index), len(currentyeardInScopingdf.index),
         len(currentyeardOnHolddf.index)])
    print('')
    print(''.join([str(currentyeartotalproject), strlist[4], str(date.today().year), strlist[5]]))

    bigprojects = len(currentyeardreleasedf[currentyeardreleasedf['Approved_WA_Amount'] >= 200000]) + len(
        currentyeardInServicedf[currentyeardInServicedf['Approved_WA_Amount'] >= 200000])

    bigprojectsda = len(currentyeardreleasedf[(currentyeardreleasedf['Approved_WA_Amount'] >= 200000) &
                                              (currentyeardreleasedf['BUDGETITEMNUMBER'] == '00003407')]) + len(
        currentyeardInServicedf[(currentyeardInServicedf['Approved_WA_Amount'] >= 200000) &
                                (currentyeardInServicedf['BUDGETITEMNUMBER'] == '00003407')])

    bigprojectsnotdareleased = len(currentyeardreleasedf[(currentyeardreleasedf['Approved_WA_Amount'] >= 200000) &
                                                         (currentyeardreleasedf['BUDGETITEMNUMBER'] != '00003407')])

    bigprojectsnotdainservice = len(currentyeardInServicedf[(currentyeardInServicedf['Approved_WA_Amount'] >= 200000) &
                                                            (currentyeardInServicedf[
                                                                 'BUDGETITEMNUMBER'] != '00003407')])

    print(' '.join([str(bigprojects), strlist[7]]))
    print(' '.join([str(bigprojectsda), strlist[8]]))

    print(' '.join([strlist[9], str(bigprojectsnotdareleased + bigprojectsnotdainservice), strlist[10]]))

    print(
        ' '.join([str(bigprojectsnotdainservice), strlist[11], str(bigprojectsnotdareleased + bigprojectsnotdainservice)
                     , strlist[12]]))

    print('')

    bigprojectsnotdareleaseddf = currentyeardreleasedf[(currentyeardreleasedf['Approved_WA_Amount'] >= 200000) &
                                                       (currentyeardreleasedf['BUDGETITEMNUMBER'] != '00003407')]

    print(bigprojectsnotdareleaseddf)

    # bigprojectsnotdawithschedules = pd.merge(PScedules_df, bigprojectsnotdainservicedf , on='PETE_ID', how='inner')
    bigprojectsnotdawPE = PScedules_df[
        (PScedules_df['Grandchild'] == 'Project Energization') &
        (PScedules_df['PETE_ID'].isin(bigprojectsnotdareleaseddf['PETE_ID']))]

    bigprojectsnotdawPE.to_csv('metro_west_large_PE.csv')

    # PEC Slide
    strlist = ['- 138KV Breaker Replacements',
               '- FID Replacements',
               '- 69 KV Breaker Replacements',
               'Released',
               'In-Service',
               'No Capital Spends']

    HVbreakerreplacementscurrentyeardf = Project_Data_df[
        (Project_Data_df['Estimated_In-Service_Date'].dt.year == date.today().year) &
        (Project_Data_df['BUDGETITEMNUMBER'] == '00003202')]

    FIDcurrentyearddf = Project_Data_df[
        (Project_Data_df['Estimated_In-Service_Date'].dt.year == date.today().year) &
        (Project_Data_df['BUDGETITEMNUMBER'] == '00003206')]

    MDbreakerreplacementscurrentyearddf = Project_Data_df[
        (Project_Data_df['Estimated_In-Service_Date'].dt.year == date.today().year) &
        (Project_Data_df['BUDGETITEMNUMBER'] == '00003203')]

    HVbreakerreplacementsnextyeardf = Project_Data_df[
        (Project_Data_df['Estimated_In-Service_Date'].dt.year == date.today().year + 1) &
        (Project_Data_df['BUDGETITEMNUMBER'] == '00003202')]

    FIDnextyearddf = Project_Data_df[
        (Project_Data_df['Estimated_In-Service_Date'].dt.year == date.today().year + 1) &
        (Project_Data_df['BUDGETITEMNUMBER'] == '00003206')]

    MDbreakerreplacementsnextyearddf = Project_Data_df[
        (Project_Data_df['Estimated_In-Service_Date'].dt.year == date.today().year + 1) &
        (Project_Data_df['BUDGETITEMNUMBER'] == '00003203')]

    print(str(date.today().year) + ':')
    print(' '.join([str(len(HVbreakerreplacementscurrentyeardf.index)), strlist[0]]))
    print(' '.join([str(len(FIDcurrentyearddf.index)), strlist[1]]))
    print(' '.join([str(len(MDbreakerreplacementscurrentyearddf.index)), strlist[2]]))

    print('')
    print(str(date.today().year + 1) + ':')
    print(' '.join([str(len(HVbreakerreplacementsnextyeardf.index)), strlist[0]]))
    print(' '.join([str(len(FIDnextyearddf.index)), strlist[1]]))
    print(' '.join([str(len(MDbreakerreplacementsnextyearddf.index)), strlist[2]]))
    print('')

    print(strlist[0])
    print(' '.join([str(
        len(HVbreakerreplacementscurrentyeardf[HVbreakerreplacementscurrentyeardf['PROJECTSTATUS'] == 'Released'])),
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

    # Tline Harding

    Watercrossingscurrentyeardf = Project_Data_df[
        (Project_Data_df['Estimated_In-Service_Date'].dt.year == date.today().year) &
        (Project_Data_df['BUDGETITEMNUMBER'] == '00003212')]

    Watercrossingsnextyeardf = Project_Data_df[
        (Project_Data_df['Estimated_In-Service_Date'].dt.year == date.today().year + 1) &
        (Project_Data_df['BUDGETITEMNUMBER'] == '00003212')]

    print('T Line Harding')
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

    # Water Crossing

    Watercrossingscurrentyeardf = Project_Data_df[
        (Project_Data_df['Estimated_In-Service_Date'].dt.year == date.today().year) &
        (Project_Data_df['BUDGETITEMNUMBER'] == '00003226')]

    Watercrossingsnextyeardf = Project_Data_df[
        (Project_Data_df['Estimated_In-Service_Date'].dt.year == date.today().year + 1) &
        (Project_Data_df['BUDGETITEMNUMBER'] == '00003226')]

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


if __name__ == "__main__":
    """ This is executed when run from the command line """
    # Setup Logging
    logger = logging.getLogger('root')
    FORMAT = "[%(filename)s:%(lineno)s - %(funcName)20s() ] %(message)s"
    logging.basicConfig(format=FORMAT)
    logger.setLevel(logging.DEBUG)

    main()
