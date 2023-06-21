# Script to classify the issues
import pandas as pd
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches
import numpy as np
import dataframe_image as dfi
import time
from datetime import date
import seaborn as sns
import statistics
from tqdm import tqdm
from time import sleep
from termcolor import colored


## SETUP ##
# file with prices before check
data = 'prices_excel_url'   
# country
country = 'Australia'
# folder path
path = f'{country}/{country} Aug'


## Files ##
#file with issues
issue_type = pd.read_csv('issue_type.csv')
issue_type = issue_type.loc[(issue_type['Country'] == country)]
issue_list = list(issue_type['Price_list'])
# Reading file with checked prices
pricesdf = pd.read_csv(f'{path}/{data}.csv')

# Preparing columns to be filled
pricesdf['difference'] = ''
pricesdf['Are prices matched?'] = ''
pricesdf['Difference type'] = ''
pricesdf['Is Contract ID the same?'] = ''
pricesdf['Is Price Program Name the same?'] = ''
pricesdf.fillna('None', inplace=True)

# Removing 'C' and 'D' from Contract ID
pricesdf['Contract ID from URL'] = pricesdf['Contract ID from URL'].astype('str')
pricesdf['Contract ID from URL'] = pricesdf['Contract ID from URL'].str.extract('(\d+)')
pricesdf['Contract ID from Excel'] = pricesdf['Contract ID from Excel'].astype('str')
pricesdf['Contract ID from Excel'] = pricesdf['Contract ID from Excel'].str.extract('(\d+)')
pricesdf.fillna('None', inplace=True)

## Check ##
# Loop for checking if the prices, Contract ID and Price Program Name are the same
for index in tqdm(pricesdf.index, desc = colored("Update on check:", "red"), colour="green"):
    # Checking if the prices are the same
    # Both None values 
    if pricesdf['Price from URL'][index] == 'None' and pricesdf['Price from Excel'][index] == 'None':
        pricesdf['Difference type'][index] = 'none value'
        pricesdf['Are prices matched?'][index] = 'no'
        pricesdf['difference'][index] = 'none value'
    # One None Value 
    elif pricesdf['Price from URL'][index] == 'None' or pricesdf['Price from Excel'][index] == 'None':
        pricesdf['Difference type'][index] = 'none value' 
        pricesdf['Are prices matched?'][index] = 'no'
        pricesdf['difference'][index] = 'none value'
    # Error Value 
    elif pricesdf['Price from URL'][index] == 'error':
        pricesdf['Difference type'][index] = 'not found in URL' 
        pricesdf['Are prices matched?'][index] = 'no'
        pricesdf['difference'][index] = 'not found'
    # Both numeric values
    else:
        pricesdf['Price from URL'][index] =  pd.to_numeric(pricesdf['Price from URL'][index])
        pricesdf['Price from Excel'][index] = pd.to_numeric(pricesdf['Price from Excel'][index])
        # The same values
        if pricesdf['Price from URL'][index] == pricesdf['Price from Excel'][index]:
            pricesdf['Difference type'][index] = 'no difference'
            pricesdf['Are prices matched?'][index] = 'yes'
            pricesdf['difference'][index] = '0'
        # Different values
        else: 
            pricesdf['difference'][index] = round(abs(pricesdf['Price from URL'][index] - pricesdf['Price from Excel'][index]),4)
            # Difference < 0.01- doesn't count as mismatch
            if pd.to_numeric(pricesdf['difference'][index]) < 0.01:
                pricesdf['Difference type'][index] = 'no difference'
                pricesdf['Are prices matched?'][index] = 'yes'
                pricesdf['difference'][index] = '0'
            # Difference > 0.01
            else:
                pricesdf['Are prices matched?'][index] = 'no'
                pricesdf['Difference type'][index] = 'different values'

    # Checking if Contract ID, Price Program Name are the same
    # Contract ID check
    # Both None values
    if pricesdf['Contract ID from URL'][index] == 'None' and pricesdf['Contract ID from Excel'][index]=='None':
        pricesdf['Is Contract ID the same?'][index] = 'none value'
    # One None value
    elif pricesdf['Contract ID from URL'][index] == 'None'and pricesdf['Contract ID from Excel'][index] != 'None':
        pricesdf['Is Contract ID the same?'][index] = 'none value'
    # One None value
    elif pricesdf['Contract ID from URL'][index] != 'None'and pricesdf['Contract ID from Excel'][index] == 'None':
        pricesdf['Is Contract ID the same?'][index] = 'none value excel'
    # Test Load Issue in Contract ID
    elif pricesdf['Contract ID from URL'][index] == 'TestLoadIssue':
        pricesdf['Is Contract ID the same?'][index] = 'none value'
    # Conract IDs the same
    elif pd.to_numeric(pricesdf['Contract ID from URL'][index]) == pd.to_numeric(pricesdf['Contract ID from Excel'][index]):
        pricesdf['Is Contract ID the same?'][index] = 'yes'
    # Contract IDs different 
    else:
        pricesdf['Is Contract ID the same?'][index] = 'no'

    # Price Program Name check
    # Both None values
    if pricesdf['Price Program from URL'][index] == 'None' and pricesdf['Price Program from Excel'][index] ==  'None':
        pricesdf['Is Price Program Name the same?'][index] = 'none value'  
    # One None value
    if pricesdf['Price Program from URL'][index] == 'None' or pricesdf['Price Program from Excel'][index] ==  'None':
        pricesdf['Is Price Program Name the same?'][index] = 'none value'
    # Price Program Names the same
    elif pricesdf['Price Program from URL'][index] == pricesdf['Price Program from Excel'][index]:
        pricesdf['Is Price Program Name the same?'][index] = 'yes'
    # Price Program Names different
    else:
        pricesdf['Is Price Program Name the same?'][index] = 'no'

# Mismatched contracts
mismatched = pricesdf.loc[(pricesdf['Are prices matched?'] == 'no')]
mismatched['Issue type'] = ''

# Check on the type of issue
for index in tqdm(mismatched.index, desc = colored("Update on type issue check:", "red"), colour="green"):
    # Different contracts
    if mismatched['Is Contract ID the same?'][index] == 'no':
        # Contract from URL is the Price List
        if pd.to_numeric(mismatched['Contract ID from URL'][index]) in issue_list:
            mismatched['Issue type'][index] = 'List Price vs Contract ID'
        # Different contract (not from Price List)
        else:
            mismatched['Issue type'][index] = 'Different Contract IDs'
    # The same Contract IDs (backdated contract)
    elif mismatched['Is Contract ID the same?'][index] == 'yes':
    #    if pd.to_numeric(mismatched['Contract ID from URL'][index]) in issue_list and pd.to_numeric(mismatched['Contract ID from Excel'][index]) in issue_list:
        #    mismatched['Issue type'][index] = 'Both List Price'
     #   else:
        mismatched['Issue type'][index] = 'Backdating issue'
    # Other issue
    elif mismatched['Is Contract ID the same?'][index] == 'none value excel':
        if pd.to_numeric(mismatched['Contract ID from URL'][index]) in issue_list:
            mismatched['Issue type'][index] = 'List Price vs Contract ID'
        else:
            mismatched['Issue type'][index] = 'Other issue'
    elif mismatched['Is Contract ID the same?'][index] == 'none value':
        mismatched['Issue type'][index] = 'Other issue'

# Exporting all the contract and mismatched contracts to csv
mismatched = mismatched.drop(['Are prices matched?', 'Difference type'], axis = 1)
mismatched.to_csv(f'{path}/files/mismatched prices {country}.csv') 
mismatched.to_excel(f'{path}/files/mismatched prices {country}.xlsx') 
pricesdf.to_csv(f'{path}/files/all transactions {country}.csv') 
pricesdf.to_excel(f'{path}/files/all transactions {country}.xlsx') 



