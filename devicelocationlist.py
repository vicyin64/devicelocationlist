import requests
import pandas as pd
import time
import logging
from datetime import datetime
from tqdm import tqdm

# Set your Webex API access token here
access_token = 'YOUR_ACCESS_TOKEN'

# Define the headers for authentication
headers = {
    'Authorization': f'Bearer {access_token}',
    'Content-Type': 'application/json'
}

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Function to make a request with throttling logic and retry mechanism
def make_request(url):
    max_retries = 5
    backoff_factor = 2
    retries = 0

    while retries < max_retries:
        try:
            response = requests.get(url, headers=headers, timeout=10)
            if response.status_code == 429:
                logging.warning("Rate limit exceeded. Sleeping for 10 seconds...")
                time.sleep(10)
            else:
                response.raise_for_status()
                return response.json()
        except requests.RequestException as e:
            logging.error(f"Request failed: {e}. Retrying in {backoff_factor ** retries} seconds...")
            time.sleep(backoff_factor ** retries)
            retries += 1

    raise Exception(f"Failed to make request after {max_retries} retries")

# Function to get all workspaces with pagination
def get_all_workspaces():
    url_template = 'https://webexapis.com/v1/workspaces?max=200'
    workspaces = []
    start = 0

    while True:
        url = f"{url_template}&start={start}"
        response = make_request(url)
        items = response['items']
        workspaces.extend(items)
        
        # Check if there are more items to fetch
        if len(items) < 200:
            break
        
        start += len(items)

    return workspaces

# Function to get devices in a workspace by workspace ID
def get_devices_in_workspace(workspace_id):
    url = f'https://webexapis.com/v1/devices?workspaceId={workspace_id}'
    return make_request(url)['items']

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
    try:
        workspaces = get_all_workspaces()
        logging.info(f"Total workspaces processed: {len(workspaces)}")
        
        data = []
        for workspace in tqdm(workspaces, desc="Processing Workspaces", unit="workspace"):
            workspace_id = workspace['id']
            devices = get_devices_in_workspace(workspace_id)
            
            location_id = workspace.get('locationId', '')
            floor_id = workspace.get('floorId', '')

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

            for device in devices:
                data.append({
                    'Device ID': device.get('id'),
                    'Device Name': device.get('displayName'),
                    'Workspace ID': workspace_id,
                    'Workspace Name': workspace.get('displayName'),
                    'Location Name': location_name,
                    'Location Address': address,
                    'Floor Number': floor_number
                })

        # Create a DataFrame and save it to an Excel file with a timestamp
        df = pd.DataFrame(data)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f'webex_devices_and_workspaces_{timestamp}.xlsx'
        df.to_excel(filename, index=False)

        logging.info(f"Data has been written to {filename}")

    except Exception as e:
        logging.error(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
