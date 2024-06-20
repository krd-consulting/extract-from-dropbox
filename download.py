import dropbox

from dotenv import load_dotenv

import os

load_dotenv()

APP_KEY=os.environ['APP_KEY']
REFRESH_TOKEN=os.environ['REFRESH_TOKEN']

file_to_download = 'file.txt'

with dropbox.Dropbox(oauth2_refresh_token=REFRESH_TOKEN, app_key=APP_KEY) as dbx:
	metadata, f = dbx.files_download('/' + file_to_download)
	output_path = os.path.join(os.environ['DOWNLOADS_FOLDER_NAME'], file_to_download);
	out = open(output_path, 'wb')
	out.write(f.content)
	print(file_to_download + " saved in " + output_path + "\n")
	out.close()