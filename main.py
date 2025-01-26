import sqlalchemy
import extract
import load
import logging
import os

from dotenv import load_dotenv

if __name__ == '__main__':
    load_dotenv()

    APP_KEY=os.environ['APP_KEY']
    REFRESH_TOKEN=os.environ['REFRESH_TOKEN']
    LOCAL_EXTRACTED_FOLDER=os.environ['LOCAL_EXTRACTED_FOLDER']
    DROPBOX_EXTRACTED_FOLDER=os.environ['DROPBOX_EXTRACTED_FOLDER']
    SERVER = os.environ['DATABASE_SERVER']
    CORE_DATABASE = os.environ['CORE_DATABASE']
    DATABASE = os.environ['DATABASE']
    CORE_ENGINE = sqlalchemy.create_engine('mssql+pyodbc://' + SERVER + '/' + CORE_DATABASE + '?TrustServerCertificate=yes&trusted_connection=yes&driver=ODBC+Driver+18+for+SQL+Server')
    ENGINE = sqlalchemy.create_engine('mssql+pyodbc://' + SERVER + '/' + DATABASE + '?TrustServerCertificate=yes&trusted_connection=yes&driver=ODBC+Driver+18+for+SQL+Server')

    logger = logging.getLogger(__name__)
    logging.basicConfig(filename='extract.log', encoding='utf-8', level=logging.DEBUG, format='%(asctime)s %(message)s')

    successfully_downloaded_entries = extract.extract_from_dropbox(app_key=APP_KEY, refresh_token=REFRESH_TOKEN, dropbox_extracted_folder=DROPBOX_EXTRACTED_FOLDER, local_extracted_folder=LOCAL_EXTRACTED_FOLDER)[0]

    for local_path, entry in successfully_downloaded_entries.items():
        filename = load.FileName(local_path)

        if(filename.document_type == load.DocumentType.BUDGET):
            budget_items = load.read_new_budget_items(local_path)
            with sqlalchemy.orm.Session(ENGINE) as session:
                load.populate_budget_items_table(session, sqlalchemy.orm.Session(CORE_ENGINE), budget_items, filename)
                session.commit()
        elif(filename.document_type == load.DocumentType.FINANCIAL_REPORT):
            with sqlalchemy.orm.Session(ENGINE) as session:
                report_filename = load.FinancialReportFileName(filename.filename)
                expense_allocation_items = load.read_financial_report(report_filename.filename)
                load.populate_expense_allocations_table(session, sqlalchemy.orm.Session(CORE_ENGINE), expense_allocation_items, report_filename)
                session.commit()