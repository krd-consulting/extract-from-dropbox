System Requirements
- Python >= 3.8
- pip

Instructions:
1. `$ python -m pip install -r requirements.txt`
2. Create a `.env` file and add your Dropbox APP_KEY. The `.env` file should look like:
```
APP_KEY=[your Dropbox App Key]
```
3. Run `python get_refresh_token.py` and complete OAuth flow. This retrieves a refresh token and saves it to `.env`.
4. To download files, run `python download.py`.