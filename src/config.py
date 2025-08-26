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

# Configuration dictionary
CONFIG = {
    'APP_ENV': os.getenv('APP_ENV', 'production'),
    'APP_DEBUG': os.getenv('APP_DEBUG', 'True').lower() == 'true',
    'APP_LOG_LEVEL': os.getenv('APP_LOG_LEVEL', 'debug').lower(),
    'LOG_DAYS_TO_KEEP': int(os.getenv('LOG_DAYS_TO_KEEP', 10)),
    'DATA_DIR': os.getenv('DATA_DIR', 'C:/HA-Data/database'),
    'LOG_DIR': os.getenv('DATA_DIR', 'C:/HA-Data/LOGS'),
    'BANK_DIR': os.getenv('BANK_DIR', 'C:/HA-Data/BANK_TRANSACTIONS'),
    'DB_PATH': os.getenv('DB_PATH', 'C:/HA-Data/database/HAdata.db'),
    'API_BASE_URL': os.getenv('API_BASE_URL', 'https://bankaccountdata.gocardless.com/api/v2'),
    'TOKEN_MARGIN': int(os.getenv('TOKEN_MARGIN', 60)),
    'REQUISITION_VALIDITY_DAYS': int(os.getenv('REQUISITION_VALIDITY_DAYS', 90)),
    'VENV_DIR': os.getenv('VENV_DIR', 'C:/HA-Project/.venv')
}

def load_db_settings():
    """Load settings from the database."""
    try:
        with sqlite3.connect(CONFIG['DB_PATH']) as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT key, value FROM settings")
            for key, value in cursor.fetchall():
                try:
                    if key in ['EXPIRY_WARNING_DAYS', 'LOG_DAYS_TO_KEEP', 'TOKEN_MARGIN', 'REQUISITION_VALIDITY_DAYS']:
                        CONFIG[key] = int(value)
                    else:
                        CONFIG[key] = value
                except ValueError:
                    logging.error(f"Invalid value for {key}: {value}")
    except sqlite3.Error as e:
        logging.error(f"Failed to load settings from database: {e}")

# Set up logging
log_levels = {
    'debug': logging.DEBUG,
    'info': logging.INFO,
    'notice': logging.INFO,
    'warning': logging.WARNING,
    'error': logging.ERROR
}

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
    load_db_settings()
    logging.getLogger('HA').debug("Configuration initialized")