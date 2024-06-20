import dropbox

from datetime import datetime
from dotenv import load_dotenv
import logging
import os
from pathlib import Path

load_dotenv()

APP_KEY=os.environ['APP_KEY']
REFRESH_TOKEN=os.environ['REFRESH_TOKEN']
EXTRACTED_FOLDER=os.environ['EXTRACTED_FILES_FOLDER_NAME']

logger = logging.getLogger(__name__)
logging.basicConfig(filename='download.log', encoding='utf-8', level=logging.DEBUG, format='%(asctime)s %(message)s')

def download_file(dropbox_object, filename):
	now = datetime.now()
	
	filename_no_extension = Path(filename).stem + '_' + now.astimezone().isoformat()
	extension = os.path.splitext(filename)[1]
	output_path = output_path = os.path.join(EXTRACTED_FOLDER, filename_no_extension + extension)
	
	dropbox_object.files_download_to_file(output_path, '/' + filename)
	
	print(filename + " saved in " + output_path)

	return filename_no_extension + extension


with dropbox.Dropbox(oauth2_refresh_token=REFRESH_TOKEN, app_key=APP_KEY) as dbx:
	result = dbx.files_list_folder('', recursive=False)
	# TODO: use files_list_folder_continue 
	successful_entries = {}
	unsuccessful_entries = []
	for entry in result.entries:
		if(isinstance(entry, dropbox.files.FileMetadata)):
			try:
				successful_index = download_file(dbx, entry.name)
			except dropbox.exceptions.ApiError as e:
				logging.error('Error when downloading file: ' + entry.name + ':' + e.error)
				unsuccessful_entries[entry]
			except Exception as e:
				raise e
				unsuccessful_entries.append(entry)
			else:
				logging.info(f"LOCAL: {entry.name} extracted to {EXTRACTED_FOLDER}/{successful_index}")
				successful_entries[successful_index] = entry

	for filename, entry in successful_entries.items():
		try:
			dbx.files_move('/'+entry.name, '/' + EXTRACTED_FOLDER + '/' + filename)
		except Exception as e:
			logging.error('Error when moving extracted file: ' + entry.name)
			logging.error(e)
		else:
			logging.info(f"DROPBOX: {entry.name} moved to {EXTRACTED_FOLDER}/{filename}")