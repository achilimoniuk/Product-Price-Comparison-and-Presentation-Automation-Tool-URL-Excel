# Script to run multiple rest call to Model N
import http.client
import json
import ssl
import requests
import json
import pandas as pd
import ast
import numpy as np
from tqdm import tqdm
from time import sleep
import datetime
import inquirer
from csv import writer

## SETUP
# Choosing excel file and country
country = 'Brazil'
file_chosen = f'{country}/{country} Aug/Historical Sales - Brazil - QA - from 20220101.xlsx'
csv_file = f'{country}/{country} Aug/prices_excel_modeln.csv'

# URL
request_url='https://abbott-coredx-qa.cloud.modeln.com/modeln/rest/cdxprice/cdxresolve'
headers = {'content-type': 'application/json'}

# Defining certificate related stuff and host of endpoint
certificate_file = 'certificates/abbott-coredx-qa.cloud.modeln.com.pem'
certificate_secret= 'your_certificate_secret'
host = 'abbott-coredx-qa.cloud.modeln.com'
 
request_headers = {
    'Content-Type': 'application/json'
}

# Reading excel file 
excel = pd.read_excel(f'{file_chosen}', header=0)
excel = pd.DataFrame(excel)

# Removing service contracts
excel = excel[excel['PRC_METHOD'] != 'RENTAL']

# Removing '-', adding new contract id column
excel['EXTERNAL_ITEM_ID'] = excel['EXTERNAL_ITEM_ID'].str.replace("-", "")

# Removing the ending from Customer ID
excel['SHIP_TO_CUST_NUM'] = excel['SHIP_TO_CUST_NUM'].astype('str')
excel['SHIP_TO_CUST_NUM'] = excel['SHIP_TO_CUST_NUM'].str.extract('(\d+)')

# If the program stopes
# Starting with the last index from prices_modeln_excel
#excel = excel[3755:]

# Adding headers
with open(f'{csv_file}', 'a+', newline='') as write_obj:
    # Create a writer object from csv module
    csv_writer = writer(write_obj)
    # Add contents of list as last row in the csv file
    csv_writer.writerow(['Customer Number', 'Reference number', 'Product Number',
'Invoice date', 
'Contract type',
'Price Program from Model N', 'Price Program from Excel', 
'Contract ID from Model N', 'Contract ID from Excel',
'Price from Model N', 'Price from Excel'])
    write_obj.close()


# Loop for getting prices
for index in tqdm(excel.index):
    try:
        request_body_dict={
            "customerNumber": f"{excel['SHIP_TO_CUST_NUM'][index]}",
            "productNumbers": [f"{excel['EXTERNAL_ITEM_ID'][index]}"],
            "currency": f"{excel['CURRENCY'][index]}",
            "maxPrices": "1",
            "org": f"{excel['EXTERNAL_ORG_ID'][index]}",
            "pricingDate" : f"{pd.to_datetime(excel['EXTERNAL_INV_DATE'][index])}",
            "sourceSys" : "SYMPHONY"
        }

        # Define the client certificate settings for https connection
        context = ssl.SSLContext(ssl.PROTOCOL_SSLv23)
        context.load_cert_chain(certfile=certificate_file)
        
        # Create a connection to submit HTTP requests
        connection = http.client.HTTPSConnection(host, port=443, context=context)
        
        # Use connection to submit a HTTP POST request
        connection.request(method="POST", url=request_url, headers=request_headers, body=json.dumps(request_body_dict))
        
        # Print the HTTP response from the IOT service endpoint
        response = connection.getresponse()
        data = response.read()

        # Decoding Byte type to json type
        my_json = data.decode('utf8')

        # Converting json to dictionary
        data = json.loads(my_json)
    
        # Getting the values and converting list to Dataframe
        dict = data['resolvedPrices']
        df = pd.DataFrame(dict, index=[0])

        # Changing the item id
        excel['EXTERNAL_ITEM_ID'][index] = excel['EXTERNAL_ITEM_ID'][index][:-2]+ '-' + excel['EXTERNAL_ITEM_ID'][index][-2:]

        # Adding lines to csv file
        with open(f'{csv_file}', 'a+', newline='') as write_obj:
            # Create a writer object from csv module
            csv_writer = writer(write_obj)
            # Add contents of list as last row in the csv file
            csv_writer.writerow([excel['SHIP_TO_CUST_NUM'][index][2:],excel['LINE_REF_NUM'][index],excel['EXTERNAL_ITEM_ID'][index],
            excel['EXTERNAL_INV_DATE'][index], 
            excel['PRC_METHOD'][index],
            df['priceProgram'][0], excel['PRICE_PROGRAM'][index],
            df['pricingDocId'][0], excel['EXTERNAL_CONTRACT_ID'][index],
            df['resolvedPrice'][0], excel['CONTRACT_AMT'][index]])
            write_obj.close()
    except:
        # Changing the item id
        excel['EXTERNAL_ITEM_ID'][index] = excel['EXTERNAL_ITEM_ID'][index][:-2]+ '-' + excel['EXTERNAL_ITEM_ID'][index][-2:]

         # Adding lines to csv file
        with open(f'{csv_file}', 'a+', newline='') as write_obj:
            # Create a writer object from csv module
            csv_writer = writer(write_obj)
            # Add contents of list as last row in the csv file
            csv_writer.writerow([excel['SHIP_TO_CUST_NUM'][index][2:],excel['LINE_REF_NUM'][index],excel['EXTERNAL_ITEM_ID'][index],
            excel['EXTERNAL_INV_DATE'][index],
            excel['PRC_METHOD'][index],
            '', excel['PRICE_PROGRAM'][index],
            '', excel['EXTERNAL_CONTRACT_ID'][index],
            'error', excel['CONTRACT_AMT'][index]])
            write_obj.close()



