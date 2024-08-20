import dropbox

from datetime import datetime
from dotenv import load_dotenv
import logging
import os
from pathlib import Path

def download_file(dropbox_object, filename, local_extracted_folder):
	now = datetime.now()
	
	timestamp = now.astimezone().strftime('%Y%m%d_%H%M')
	new_filename_no_extension = Path(filename).stem + '_' + timestamp
	extension = os.path.splitext(filename)[1]
	output_folder = os.path.join(local_extracted_folder, timestamp)
	output_path =  os.path.join(output_folder, new_filename_no_extension + extension)
	if not os.path.exists(output_folder):
		os.makedirs(output_folder)
	path_to_download = os.path.join('/', filename)
	dropbox_object.files_download_to_file(output_path, path_to_download)

	return new_filename_no_extension + extension

def download_files_from_entries(dropbox_object, entries, local_extracted_folder):
	successful_entries = {}
	unsuccessful_entries = []

	for entry in entries:
		if(not isinstance(entry, dropbox.files.FileMetadata)): 
			continue

		try:
			successful_index = download_file(dbx, entry.name, local_extracted_folder)
		except dropbox.exceptions.ApiError as e:
			logging.error('Error when downloading file: ' + entry.name + ':' + e.error)
			unsuccessful_entries[entry]
		except Exception as e:
			raise e
			unsuccessful_entries.append(entry)
		else:
			logging.info(f"LOCAL: {entry.name} extracted to {local_extracted_folder}/{successful_index}")
			successful_entries[successful_index] = entry

	return (successful_entries, unsuccessful_entries)

def move_extracted_files(dropbox_object, successful_entries, dropbox_extracted_folder):
	for filename, entry in successful_entries.items():
		try:
			file_path = '/' + entry.name
			output_path = '/' + dropbox_extracted_folder + '/' + filename
			dbx.files_move(file_path, output_path)
		except Exception as e:
			logging.error('Error when moving extracted file: ' + entry.name + 'to: ' + str(output_path))
			logging.error(e)
		else:
			logging.info(f"DROPBOX: {file_path} moved to {output_path}")


def extract_from_dropbox(app_key=app_key, refresh_token=refresh_token, dropbox_extracted_folder=dropbox_extracted_folder, local_extracted_folder=local_extracted_folder):
	with dropbox.Dropbox(oauth2_refresh_token=refresh_token, app_key=app_key) as dbx:
		result = dbx.files_list_folder('', recursive=False)
		entries = result.entries
		logging.info(f"seeing { len(entries)} files in Dropbox")

		while(result.has_more):
			result = dbx.files_list_folder_continue(result.cursor)
			entries += result.entries
			logging.info(f"seeing { len(entries)} files in Dropbox")
		successful_entries, unsuccessful_entries = download_files_from_entries(dbx, entries, local_extracted_folder)
		move_extracted_files(dbx, successful_entries, dropbox_extracted_folder)

if __name__ == '__main__':
	load_dotenv()

	APP_KEY=os.environ['APP_KEY']
	REFRESH_TOKEN=os.environ['REFRESH_TOKEN']
	LOCAL_EXTRACTED_FOLDER=os.environ['LOCAL_EXTRACTED_FOLDER']
	DROPBOX_EXTRACTED_FOLDER=os.environ['DROPBOX_EXTRACTED_FOLDER']

	logger = logging.getLogger(__name__)
	logging.basicConfig(filename='extract.log', encoding='utf-8', level=logging.DEBUG, format='%(asctime)s %(message)s')

	extract_from_dropbox(app_key=APP_KEY, refresh_token=REFRESH_TOKEN, dropbox_extracted_folder=DROPBOX_EXTRACTED_FOLDER, local_extracted_folder=LOCAL_EXTRACTED_FOLDER)