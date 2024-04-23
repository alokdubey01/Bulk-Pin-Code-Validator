import requests
import pandas as pd

# Function to fetch details for a given PIN code using the API
def fetch_pin_details(pin_code):
    url = f"https://api.postalpincode.in/pincode/{pin_code}"
    response = requests.get(url)
    if response.status_code == 200:
        data = response.json()
        if data[0]['Status'] == 'Success' and len(data[0]['PostOffice']) > 0:
            # Using only the first available entry in the PostOffice list
            pin_info = data[0]['PostOffice'][0]  # Changed from 1 to 0 for safe access to the first item
            return {
                'PIN Code': pin_code,
                'Office Name': pin_info['Name'],
                'Branch Type': pin_info['BranchType'],
                'Delivery Status': pin_info['DeliveryStatus'],
                'Circle': pin_info['Circle'],
                'District': pin_info['District'],
                'Division': pin_info['Division'],
                'Region': pin_info['Region'],
                'Block': pin_info['Block'],
                'State': pin_info['State'],
                'Country': pin_info['Country']
            }
        else:
            return {'PIN Code': pin_code, 'Error': 'No data found or invalid PIN code'}
    else:
        print(f"Failed to fetch data for PIN code: {pin_code} with status code {response.status_code}")
        return None

# List of example PIN codes to query
pin_codes = [
    110001, 700001, 600001, 400001, 500001, 
    560001, 110002, 122001, 380001, 411001, 
    302001, 682001, 700091, 834001, 141001,
    160001, 781001, 305001, 452001, 144001,
    751001, 641001, 302016, 226001, 390001,
    482001, 248001, 273001, 361001, 250001,
    201001, 517501, 403001, 248003, 600007,
    834002, 713101, 713203, 828111, 700027
]

# Efficient fetching: Reduce API calls by checking result before re-calling the function
pin_details = []
for pin in pin_codes:
    result = fetch_pin_details(pin)
    if result:
        pin_details.append(result)

# Convert to DataFrame
result_df = pd.DataFrame(pin_details)

# Write to Excel
result_df.to_excel('pin_code_details.xlsx', index=False, engine='openpyxl')

print("PIN code data written to Excel successfully!")
