import requests
import pandas as pd
import time
from datetime import datetime

# Set your Webex API access token here
access_token = 'YOUR_ACCESS_TOKEN'

# Define the headers for authentication
headers = {
    'Authorization': f'Bearer {access_token}',
    'Content-Type': 'application/json'
}

# Function to make a request with throttling logic
def make_request(url):
    while True:
        response = requests.get(url, headers=headers)
        if response.status_code == 429:
            print("Rate limit exceeded. Sleeping for 10 seconds...")
            time.sleep(10)
        else:
            response.raise_for_status()
            return response.json()

# Function to get all devices
def get_all_devices():
    url = 'https://webexapis.com/v1/devices'
    return make_request(url)['items']

# Function to get workspace details by workspace ID
def get_workspace_details(workspace_id):
    url = f'https://webexapis.com/v1/workspaces/{workspace_id}'
    return make_request(url)

# Function to get location details by location ID
def get_location_details(location_id):
    url = f'https://webexapis.com/v1/locations/{location_id}'
    return make_request(url)

# Function to get floor details by location ID and floor ID
def get_floor_details(location_id, floor_id):
    url = f'https://webexapis.com/v1/locations/{location_id}/floors/{floor_id}'
    return make_request(url)

# Function to format address
def format_address(address_details):
    address_parts = [
        address_details.get('address1', ''),
        address_details.get('city', ''),
        address_details.get('state', ''),
        address_details.get('postalCode', ''),
        address_details.get('country', '')
    ]
    return ', '.join(part for part in address_parts if part)

# Main function to gather data and create an Excel spreadsheet
def main():
    devices = get_all_devices()
    data = []
    
    for device in devices:
        workspace_id = device.get('workspaceId')
        if workspace_id:
            workspace_details = get_workspace_details(workspace_id)
            location_id = workspace_details.get('locationId', '')
            floor_id = workspace_details.get('floorId', '')

            location_name = ''
            floor_number = ''
            address = ''

            if location_id:
                location_details = get_location_details(location_id)
                location_name = location_details.get('name', '')
                address = format_address(location_details.get('address', {}))

            if location_id and floor_id:
                floor_details = get_floor_details(location_id, floor_id)
                floor_number = floor_details.get('floorNumber', '')

            data.append({
                'Device ID': device.get('id'),
                'Device Name': device.get('displayName'),
                'Workspace ID': workspace_id,
                'Workspace Name': workspace_details.get('displayName'),
                'Location Name': location_name,
                'Location Address': address,
                'Floor Number': floor_number
            })

    # Create a DataFrame and save it to an Excel file with a timestamp
    df = pd.DataFrame(data)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f'webex_devices_and_workspaces_{timestamp}.xlsx'
    df.to_excel(filename, index=False)

    print(f"Data has been written to {filename}")

if __name__ == "__main__":
    main()
