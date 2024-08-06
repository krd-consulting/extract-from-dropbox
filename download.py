import dropbox

from datetime import datetime
from dotenv import load_dotenv
import logging
import os
from pathlib import Path

load_dotenv()

APP_KEY=os.environ['APP_KEY']
REFRESH_TOKEN=os.environ['REFRESH_TOKEN']
LOCAL_EXTRACTED_FOLDER=os.environ['LOCAL_EXTRACTED_FOLDER']
DROPBOX_EXTRACTED_FOLDER=os.environ['DROPBOX_EXTRACTED_FOLDER']

logger = logging.getLogger(__name__)
logging.basicConfig(filename='download.log', encoding='utf-8', level=logging.DEBUG, format='%(asctime)s %(message)s')

def download_file(dropbox_object, filename):
	now = datetime.now()
	
	timestamp = now.astimezone().strftime('%Y%m%d_%H%M')
	filename_no_extension = Path(filename).stem + '_' + timestamp
	extension = os.path.splitext(filename)[1]
	output_folder = os.path.join(LOCAL_EXTRACTED_FOLDER, timestamp)
	output_path =  os.path.join(output_folder, filename_no_extension + extension)
	if not os.path.exists(output_folder):
		os.makedirs(output_folder)
	path_to_download = os.path.join('/', filename)
	dropbox_object.files_download_to_file(output_path, path_to_download)

	return filename_no_extension + extension

def download_files_from_entries(dropbox_object, entries):
	successful_entries = {}
	unsuccessful_entries = []


	for entry in entries:
		if(not isinstance(entry, dropbox.files.FileMetadata)): 
			continue

		try:
			successful_index = download_file(dbx, entry.name)
		except dropbox.exceptions.ApiError as e:
			logging.error('Error when downloading file: ' + entry.name + ':' + e.error)
			unsuccessful_entries[entry]
		except Exception as e:
			raise e
			unsuccessful_entries.append(entry)
		else:
			logging.info(f"LOCAL: {entry.name} extracted to {LOCAL_EXTRACTED_FOLDER}/{successful_index}")
			successful_entries[successful_index] = entry

	return (successful_entries, unsuccessful_entries)

def move_extracted_files(dropbox_object, successful_entries):
	for filename, entry in successful_entries.items():
		try:
			file_path = '/' + entry.name
			output_path = '/' + DROPBOX_EXTRACTED_FOLDER + '/' + filename
			dbx.files_move(file_path, output_path)
		except Exception as e:
			logging.error('Error when moving extracted file: ' + entry.name + 'to: ' + str(output_path))
			logging.error(e)
		else:
			logging.info(f"DROPBOX: {file_path} moved to {output_path}")


with dropbox.Dropbox(oauth2_refresh_token=REFRESH_TOKEN, app_key=APP_KEY) as dbx:
	result = dbx.files_list_folder('', recursive=False)
	entries = result.entries
	logging.info(f"seeing { len(entries)} files in Dropbox")

	while(result.has_more):
		result = dbx.files_list_folder_continue(result.cursor)
		entries += result.entries
		logging.info(f"seeing { len(entries)} files in Dropbox")
	successful_entries, unsuccessful_entries = download_files_from_entries(dbx, entries)
	move_extracted_files(dbx, successful_entries)


	