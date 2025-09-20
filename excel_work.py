from openpyxl import load_workbook
from dotenv import load_dotenv
import os
import io
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

# Scopes required for accessing Google Drive
SCOPES = ['https://www.googleapis.com/auth/drive']


def authenticate_google_drive():
    """
    Authenticate and return Google Drive service object.

    Returns:
        googleapiclient.discovery.Resource: Authenticated Drive service
    """
    creds = None

    # Check if token.json exists (stored credentials)
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

    # If no valid credentials, get new ones
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            # You need to download credentials.json from Google Cloud Console
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)

        # Save credentials for next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    return build('drive', 'v3', credentials=creds)


def download_file_by_id(file_id, destination_path, service=None):
    """
    Download a file from Google Drive by its file ID.

    Args:
        file_id (str): Google Drive file ID
        destination_path (str): Local path where file will be saved
        service: Optional pre-authenticated Drive service object

    Returns:
        bool: True if successful, False otherwise
    """
    try:
        if service is None:
            service = authenticate_google_drive()

        # Get file metadata
        file_metadata = service.files().get(fileId=file_id).execute()
        file_name = file_metadata.get('name', 'unknown_file')

        print(f"Downloading: {file_name}")

        # Request file content
        request = service.files().get_media(fileId=file_id)

        # Create file stream
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)

        done = False
        while done is False:
            status, done = downloader.next_chunk()
            print(f"Download progress: {int(status.progress() * 100)}%")

        # Write to file
        with open(destination_path, 'wb') as f:
            f.write(fh.getvalue())

        print(f"File downloaded successfully to: {destination_path}")
        return True

    except Exception as e:
        print(f"Error downloading file: {str(e)}")
        return False


def download_file_by_name(file_name, destination_path, service=None):
    """
    Download a file from Google Drive by searching for its name.

    Args:
        file_name (str): Name of the file to search for
        destination_path (str): Local path where file will be saved
        service: Optional pre-authenticated Drive service object

    Returns:
        bool: True if successful, False otherwise
    """
    try:
        if service is None:
            service = authenticate_google_drive()

        # Search for file by name
        results = service.files().list(
            q=f"name='{file_name}'",
            fields="files(id, name)"
        ).execute()

        files = results.get('files', [])

        if not files:
            print(f"No file found with name: {file_name}")
            return False

        if len(files) > 1:
            print(f"Multiple files found with name '{file_name}'. Using the first one.")

        file_id = files[0]['id']
        return download_file_by_id(file_id, destination_path, service)

    except Exception as e:
        print(f"Error searching/downloading file: {str(e)}")
        return False


def list_drive_files(service=None, max_results=10):
    """
    List files in Google Drive (useful for finding file IDs).

    Args:
        service: Optional pre-authenticated Drive service object
        max_results (int): Maximum number of files to return

    Returns:
        list: List of file dictionaries with id, name, and mimeType
    """
    try:
        if service is None:
            service = authenticate_google_drive()

        results = service.files().list(
            pageSize=max_results,
            fields="files(id, name, mimeType, size)"
        ).execute()

        files = results.get('files', [])

        if not files:
            print('No files found.')
            return []

        print(f"Found {len(files)} files:")
        for file in files:
            size = file.get('size', 'Unknown size')
            print(f"- {file['name']} (ID: {file['id']}) [{file['mimeType']}] - {size} bytes")

        return files

    except Exception as e:
        print(f"Error listing files: {str(e)}")
        return []


def search_files(file_name=None, query=None, service=None, max_results=10):
    """
    Search for files in Google Drive and return their IDs.

    Args:
        file_name (str): Exact file name to search for
        query (str): Custom search query (e.g., "name contains 'report'")
        service: Optional pre-authenticated Drive service object
        max_results (int): Maximum number of files to return

    Returns:
        list: List of dictionaries with file info (id, name, mimeType)
    """
    try:
        if service is None:
            service = authenticate_google_drive()

        # Build search query
        if file_name:
            search_query = f"name='{file_name}'"
        elif query:
            search_query = query
        else:
            search_query = ""  # Return all files

        results = service.files().list(
            q=search_query,
            pageSize=max_results,
            fields="files(id, name, mimeType, size, modifiedTime)"
        ).execute()

        files = results.get('files', [])

        if not files:
            print('No files found.')
            return []

        print(f"Found {len(files)} file(s):")
        for file in files:
            size = file.get('size', 'Unknown size')
            modified = file.get('modifiedTime', 'Unknown date')
            print(f"- {file['name']}")
            print(f"  ID: {file['id']}")
            print(f"  Type: {file['mimeType']}")
            print(f"  Size: {size} bytes")
            print(f"  Modified: {modified}")
            print()

        return files

    except Exception as e:
        print(f"Error searching files: {str(e)}")
        return []


def get_file_id_by_name(file_name, service=None):
    """
    Get the file ID of a file by its name (returns first match).
    
    Args:
        file_name (str): Name of the file to find
        service: Optional pre-authenticated Drive service object
    
    Returns:
        str: File ID if found, None otherwise
    """
    files = search_files(file_name=file_name, service=service, max_results=1)
    return files[0]['id'] if files else None


#TODO Get the excel file from google drive
# def read_excel_from_drive(file_id: str, sheet: str, cell: str):

#     wb = load_workbook(file)
#     ws = wb[sheet]

#     value = ws[cell].value
#     wb.close()
#     return value

# # Usage: Pass the Google Drive file ID (the long string in the link)
# val = read_excel_from_drive("YOUR_FILE_ID", "Sheet1", "A1")
# print("Value in A1:", val)

# def read_cells(file_path:str, sheet_name: str, cell_x: str, cell_y: str):
#     wb = load_workbook(file_path)
#     sheet = wb[sheet_name]
    
#     #TODO Get the values to calculate from the correct months
#     value_x = sheet[cell_x].value
#     value_y = sheet[cell_y].value

#     #TODO Calculate all the expenses

#     #TODO Call the write function to write the final value

#     wb.close()
#     return value_x, value_y

# def write_cell(file_path:str, sheet_name: str, cell: str, value):
#     #TODO Get the value to write and write
#     wb = load_workbook("file_path")
#     sheet = wb[""]

#     sheet[cell] = value  # write value into the cell

#     wb.save(FILE_PATH)
#     wb.close()

def main():
    load_dotenv()
    # file_path = os.getenv("FILE_PATH")
    # sheet_name = "Despesas David 25"

    #TODO Get the values from first person

    #TODO Get the values from the second person

    #TODO Calculate who spent more and how much owes the other person

    #TODO Write how much it owes and who owes who 
    #v_x, v_y = read_cells(file_path, sheet_name, "D16", "E16")	
    #print(v_x)
    #print(v_y)

    # Get file id
    file_id = get_file_id_by_name("CalculadoraDespesas.xlsx")
    print(file_id)

if __name__ == "__main__":
    main()
