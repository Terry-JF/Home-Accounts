# src/config.py
# Centralized configuration for Home Accounts project

import os
from dotenv import load_dotenv
import sqlite3
import logging
from logging.handlers import TimedRotatingFileHandler

# Load .env file from project root
load_dotenv()

# Base directory for the project
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Configuration dictionaries
CONFIG = {
    'APP_ENV': os.getenv('APP_ENV', 'test'),
    'APP_DEBUG': os.getenv('APP_DEBUG', 'True').lower() == 'true',
    'APP_LOG_LEVEL': os.getenv('APP_LOG_LEVEL', 'debug').lower(),
    'LOG_DAYS_TO_KEEP': int(os.getenv('LOG_DAYS_TO_KEEP', 10)),
    'DATA_DIR': os.getenv('DATA_DIR', 'C:/HA-Data/database'),
    'LOG_DIR': os.getenv('LOG_DIR', 'C:/HA-Data/LOGS'),
    'BANK_DIR': os.getenv('BANK_DIR', 'C:/HA-Data/BANK_TRANSACTIONS'),
    'DB_PATH': os.getenv('DB_PATH', 'C:/HA-Data/database/HAtest.db'),
    'DB_PATH_TEST': os.getenv('DB_PATH_TEST', 'C:/HA-Data/database/HAtest.db'),
    'API_BASE_URL': os.getenv('API_BASE_URL', 'https://bankaccountdata.gocardless.com/api/v2'),
    'TOKEN_MARGIN': int(os.getenv('TOKEN_MARGIN', 60)),
    'REQUISITION_VALIDITY_DAYS': int(os.getenv('REQUISITION_VALIDITY_DAYS', 90)),
    'VENV_DIR': os.getenv('VENV_DIR', 'C:/HA-Project/.venv')
}

# Colour palette dictionaries
COLORS = {
    "white": "#FFFFFF",             # Weekday BG, Active Tab Text, Daily Totals Text, Non-Active Tab BG
    "oldlace": "#DEE7DF",           # Weekend BG
    "yellow": "#FFFF80",            # Flag 1
    "green": "#80FF80",             # Flag 2
    "cyan": "#B9FFFF",              # Flag 4
    "marker": "#ED87ED",            # Row Marker
    "pale_blue": "#ADD8E6",         # was "#DFFFFF" - Exit Button
    "very_pale_blue": "#DFFFFF",    # Main Form BG - was #E6F0FA
    "very_pale_pink": "#F0D0D0",    # Main Form BG - when in TEST mode
    "red": "#DD0000",               # was "#FAA0A0" - Daily Totals OD Text
    "dark_grey": "#5D5D5D",         # Daily Totals BG
    "pink": "#FADBFD",              # Daily Totals OD BG
    "black": "#000000",             # Non-Active Tab Text
    "dark_brown": "#803624",        # Active Tab BG
    "orange": "#FFC993",            # Drill Down Marker
    "grey": "#E0E0E0",              # AHB row 8 background
    "pale_green": "#E0FFE0"         # AHB row 3 background
}

TEXT_COLORS = {                     # used for transaction status
    "Unknown": "gray",
    "Forecast": "#800040",          # Brown
    "Processing": "#0000FF",        # Blue
    "Complete": "#000000"           # Black 
}

# Set up logging
log_levels = {
    'debug': logging.DEBUG,
    'info': logging.INFO,
    'warning': logging.WARNING,
    'error': logging.ERROR
}

def load_db_settings():
    global master_bg, COLORS

    # Set up logging
    logger = logging.getLogger('HA.db_setup')
    
    # Get correct DB and set master_bg colour
    if CONFIG['APP_ENV'] == 'test':
        dbpath = get_config('DB_PATH_TEST')
        master_bg=COLORS["very_pale_pink"]
    else:
        dbpath = get_config('DB_PATH')
        master_bg=COLORS["very_pale_blue"]
        
    logger.debug(f"Opening database: {dbpath}")
    #logger.debug(f"master_bg: {master_bg}")

    """Load settings from the database."""
    try:
        with sqlite3.connect(dbpath) as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT key, value FROM settings")
            for key, value in cursor.fetchall():
                try:
                    if key in ['EXPIRY_WARNING_DAYS', 'LOG_DAYS_TO_KEEP', 'TOKEN_MARGIN', 'REQUISITION_VALIDITY_DAYS']:
                        CONFIG[key] = int(value)
                    else:
                        CONFIG[key] = value
                except ValueError:
                    logger.error(f"Invalid value for {key}: {value}")
    except sqlite3.Error as e:
        logger.error(f"Failed to load settings from database: {e}")

def setup_logging():
    """Configure logging with file and optional console handlers."""
    logger = logging.getLogger('HA')
    logger.setLevel(log_levels.get(CONFIG['APP_LOG_LEVEL'], logging.DEBUG))

    # Clear any existing handlers to prevent duplicates
    logger.handlers.clear()

    # File handler with rotation
    log_file = os.path.join(CONFIG['LOG_DIR'], 'app.log')
    file_handler = TimedRotatingFileHandler(
        log_file,
        when='midnight',
        interval=1,
        backupCount=CONFIG['LOG_DAYS_TO_KEEP']
    )
    file_handler.setFormatter(logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s'))
    logger.addHandler(file_handler)

    # Console handler (optional, only for interactive runs)
    if CONFIG['APP_DEBUG']:
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s'))
        logger.addHandler(console_handler)

def get_config(key):
    """Get a configuration value by key."""
    return CONFIG.get(key)

def init_config():
    """Initialize configuration and ensure directories exist."""
    for dir_path in [CONFIG['DATA_DIR'], CONFIG['LOG_DIR'], CONFIG['BANK_DIR']]:
        os.makedirs(dir_path, exist_ok=True)
    setup_logging()
    
    # Set up logging
    logger = logging.getLogger('HA.init')
    logger.debug("++++++++++++++++++++++++ START OF NEW SESSION ++++++++++++++++++++++++")
    logger.debug(f"LOG_DIR: {CONFIG['LOG_DIR']}")
    
    load_db_settings()
    logger.debug("Configuration initialized")
    