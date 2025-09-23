import os
import io
from dotenv import load_dotenv
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from openpyxl import load_workbook

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


def get_file_for_editing(file_id, local_filename=None, service=None):
    """
    Download a file from Google Drive for local editing.

    Args:
        file_id (str): Google Drive file ID
        local_filename (str): Optional local filename (defaults to original name)
        service: Optional pre-authenticated Drive service object

    Returns:
        str: Path to the downloaded file, or None if failed
    """
    try:
        if service is None:
            service = authenticate_google_drive()

        # Get file metadata first
        file_metadata = service.files().get(fileId=file_id).execute()
        original_name = file_metadata.get('name', f'file_{file_id}')

        # Use provided filename or original name
        if local_filename is None:
            local_filename = original_name

        print(f"Getting file: {original_name}")

        # Download the file
        success = download_file_by_id(file_id, local_filename, service)

        if success:
            print(f"File ready for editing: {local_filename}")
            return local_filename
        else:
            return None

    except Exception as e:
        print(f"Error getting file: {str(e)}")
        return None


def save_file_back_to_drive(file_id, local_filename, service=None):
    """
    Save your edited local file back to Google Drive (updates the same file ID).

    Args:
        file_id (str): Google Drive file ID to update
        local_filename (str): Path to your edited local file
        service: Optional pre-authenticated Drive service object

    Returns:
        bool: True if successful, False otherwise
    """
    try:
        if service is None:
            service = authenticate_google_drive()

        success = update_file_content(file_id, local_filename, service)

        if success:
            print(f"File saved back to Google Drive! File ID: {file_id}")

        return success

    except Exception as e:
        print(f"Error saving file back to Drive: {str(e)}")
        return False


def update_file_content(file_id, new_file_path, service=None):
    """
    Update the content of an existing file in Google Drive.

    Args:
        file_id (str): Google Drive file ID to update
        new_file_path (str): Local path to the new content
        service: Optional pre-authenticated Drive service object

    Returns:
        bool: True if successful, False otherwise
    """
    try:
        if service is None:
            service = authenticate_google_drive()

        media = MediaFileUpload(new_file_path)
        updated_file = service.files().update(
            fileId=file_id,
            media_body=media
        ).execute()

        print(f"File content updated successfully. File ID: {updated_file.get('id')}")
        return True

    except Exception as e:
        print(f"Error updating file: {str(e)}")
        return False


def edit_file_workflow(file_id, service=None):
    """
    Complete workflow: download file, let you edit it, then save it back.

    Args:
        file_id (str): Google Drive file ID
        service: Optional pre-authenticated Drive service object

    Returns:
        str: Path to local file for editing
    """
    print("=" * 50)
    print("FILE EDITING WORKFLOW")
    print("=" * 50)

    # Download the file
    local_file = get_file_for_editing(file_id, service=service)

    if local_file:
        print(f"\n✅ File downloaded: {local_file}")
        # print(f"save_file_back_to_drive('{file_id}', '{local_file}')")
        print("\n" + "=" * 50)

        return local_file
    else:
        print("❌ Failed to download file")
        return None


def read_cells(file_path: str, sheet_name: str):
    wb = load_workbook(file_path, data_only=True)
    sheet = wb[sheet_name]
    # From C to M line 16, 24, 50, 39

    #TODO Get the values to calculate from the correct months
    value_x = sheet[cell_x].value
    value_y = sheet[cell_y].value

    #TODO Calculate all the expenses

    #TODO Call the write function to write the final value

    wb.close()
    return value_x, value_y

# def write_cell(file_path:str, sheet_name: str, cell: str, value):
#     #TODO Get the value to write and write
#     wb = load_workbook("file_path")
#     sheet = wb[""]

#     sheet[cell] = value  # write value into the cell

#     wb.save(FILE_PATH)
#     wb.close()


def main():
    load_dotenv()
    file_id = os.getenv("FILE_ID")
    sheet_1 = os.getenv("SHEET_1")
    sheet_2 = os.getenv("SHEET_2")

    # Get file
    local_file = edit_file_workflow(file_id)
    print(local_file)

    #TODO Check last month with data and start on the month before until paid month found


    #TODO Get the values from first person
    v_1 = read_cells(local_file, sheet_1)
    print(v_1)

    #TODO Get the values from the second person
    v_2 = read_cells(local_file, sheet_2)
    print(v_2)

    #TODO Calculate who spent more and how much owes the other person

    #TODO Write how much it owes and who owes who
    # v_x, v_y = read_cells(file_path, sheet_name, "D16", "E16")
    # print(v_x)
    # print(v_y)


if __name__ == "__main__":
    main()
