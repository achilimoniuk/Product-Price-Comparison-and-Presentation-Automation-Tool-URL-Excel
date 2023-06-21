# Script to run multiple rest call to URL
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
country = 'country'
file_chosen = f'{country}//file.xlsx'
csv_file = f'{country}/prices_country.csv'

# URL
request_url='requested url'
headers = {'content-type': 'application/json'}

# Defining certificate related stuff and host of endpoint
certificate_file = 'certificates/certificate_file'
certificate_secret= 'secret'
host = 'host.com'
 
request_headers = {
    'Content-Type': 'application/json'
}

# Reading excel file 
excel = pd.read_excel(f'{file_chosen}', header=0)
excel = pd.DataFrame(excel)

# Removing chosen contracts
excel = excel[excel['Method'] != 'method_chosen']

# Removing '-' from id
excel['ID'] = excel['ID'].str.replace("-", "")

# Removing the ending from Customer
excel['Customer'] = excel['Customer'].astype('str')
excel['Customer'] = excel['Customer'].str.extract('(\d+)')

# Excel with the last stop
#excel = excel[22413:]

# Adding headers
with open(f'{csv_file}', 'a+', newline='') as write_obj:
    # Create a writer object from csv module
    csv_writer = writer(write_obj)
    # Add contents of list as last row in the csv file
    csv_writer.writerow(['Customer', 'Reference number', 'Product',
'Date', 
'Contract type',
'Price Program from URL', 'Price Program from Excel', 
'Contract ID from URL', 'Contract ID from Excel',
'Price from URL', 'Price from Excel'])
    write_obj.close()


# Loop for getting prices from URL based on excel info
for index in tqdm(excel.index):
    try:
        request_body_dict={
            "customer": f"{excel['Customer'][index]}",
            "product": [f"{excel['Product'][index]}"],
            "currency": f"{excel['Currency'][index]}",
            "organization": f"{excel['Organization'][index]}",
            "date" : f"{pd.to_datetime(excel['Date'][index])}",
            "system" : "SYMPHONY"
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
        dict = data['prices']
        df = pd.DataFrame(dict, index=[0])

        # Changing the id
        excel['ID'][index] = excel['ID'][index][:-2]+ '-' + excel['ID'][index][-2:]

        # Adding lines to csv file
        with open(f'{csv_file}', 'a+', newline='') as write_obj:
            # Create a writer object from csv module
            csv_writer = writer(write_obj)
            # Add contents of list as last row in the csv file
            csv_writer.writerow([excel['Customer'][index][2:],excel['Reference number'][index],
            excel['Product'][index],
            excel['Date'][index], 
            excel['Method'][index],
            df['Program'][0], 
            excel['Program'][index],
            df['ID'][0], excel['ID'][index],
            df['Price'][0], excel['Price'][index]])
            write_obj.close()

    # Adding lines of products that haven't been found in URL
    except:
        # Changing the item id
        excel['ID'][index] = excel['ID'][index][:-2]+ '-' + excel['ID'][index][-2:]

         # Adding lines to csv file
        with open(f'{csv_file}', 'a+', newline='') as write_obj:
            # Create a writer object from csv module
            csv_writer = writer(write_obj)
            # Add contents of list as last row in the csv file
            csv_writer.writerow([excel['Customer'][index][2:],excel['Reference number'][index],
            excel['Product'][index],
            excel['Date'][index], 
            excel['Method'][index],
            '', 
            excel['Program'][index],
            '', excel['ID'][index],
            'error', excel['Price'][index]])
            write_obj.close()



