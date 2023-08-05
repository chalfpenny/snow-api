import openpyxl
import json

# Load workbook
wb = openpyxl.load_workbook(r'C:\Users\Charles\OneDrive\Personal Documents\Glenhaven\sites.xlsx')

# Select active sheet
sheet = wb.active

# Read data from spreadsheet
data = []
for idx, row in enumerate(sheet.iter_rows(min_row=2, values_only=True)):  # min_row=2 to skip header
    customer_name, site_name, street, city, state, zip_code = row
    data.append(
        {
            'SiteId': idx+1, 
            'SiteName': site_name, 
            'Street': street,
            'City': city,
            'State': state,
            'ZipCode': str(zip_code),
            'CustomerName': customer_name
        })
    
    if idx == 30:
        break

# Write data to JSON file
#with open('output.json', 'w') as f:
#    json.dump(data, f, indent=None)

with open('output.json', 'w') as f:
    for record in data:
        f.write(json.dumps(record))
        f.write(',\n')


