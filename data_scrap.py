#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import requests
from bs4 import BeautifulSoup
import pandas as pd

url = 'https://docs.google.com/spreadsheets/u/0/d/e/2PACX-1vR-Ozx9mQGe6Tf9_Lg3QGWnLIbsetmUjLxVgLzRcmHfKopMysK_90pTdoZkagEJ_9C1U2pG91ORIOED/pubhtml?pli=1'

# Send a GET request to the website
response = requests.get(url)

# Check if the request was successful (status code 200)
if response.status_code == 200:
    # Parse the HTML content of the page
    soup = BeautifulSoup(response.text, 'html.parser')

    # Extract data from the table
    table = soup.find('table')
    table_data = []
    if table:
        rows = table.find_all('tr')
        for row in rows:
            columns = row.find_all(['th', 'td'])
            row_data = [col.text.strip() for col in columns]
            table_data.append(row_data)

        # If there are headers and data
        if len(table_data) > 1:
            headers = ['Year', 'Number of deaths', 'Vehicules(millions)', 'Vehicle miles (billions)	', 'Drivers (millions)', 'Per 10,000 motor vehicles', 'Per 100,000,000 vehicle miles', 'Per 100,000 population']  # Replace with your actual column names
            data = table_data[1:]
            modified_data = [subarray[1:] for subarray in data]
            modified_data = modified_data[3:-5]

            print(modified_data)
            # Create a DataFrame from the extracted data
            df = pd.DataFrame(modified_data, columns=headers)
            
            # Export the DataFrame to an Excel file with better formatting
            excel_file_path = 'C:/Tools/scraped_data11.xlsx'
            df.to_excel(excel_file_path, index=False)
            print(f'Data has been exported to {excel_file_path}')
        else:
            print('Table has no data.')

else:
    print(f"Failed to retrieve the page. Status code: {response.status_code}")

