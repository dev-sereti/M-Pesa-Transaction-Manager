import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent

DATABASE_PATH = os.path.join(BASE_DIR, 'mpesa_transactions.db')

# DEFAULT_EXCEL_PATH = os.path.join(BASE_DIR, 'transactions.xlsx')
DEFAULT_EXCEL_PATH = os.path.join(BASE_DIR, 'Expenditure Tracker v8.xlsx')

LOG_LEVEL = 'INFO'
LOG_FILE = os.path.join(BASE_DIR, 'mpesa_app.log')