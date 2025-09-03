# src/config.py
# Centralized configuration for Home Accounts project

import os
from dotenv import load_dotenv
import sqlite3
import logging
from logging.handlers import TimedRotatingFileHandler

# Load .env file from project root
load_dotenv()

# Set up logging
logger = logging.getLogger('HA.config')

# Set up logging
log_levels = {
    'debug': logging.DEBUG,
    'info': logging.INFO,
    'warning': logging.WARNING,
    'error': logging.ERROR
}
# Set up Fonts - standardise use of fonts across app
ha_normal = "Arial", 10
ha_normal_bold = "Arial", 10, "bold"
ha_normal_list = '("Arial", 10)'
ha_button = "Arial", 11
ha_large = "Arial", 12
ha_vlarge = "Arial", 14
ha_help = "Arial", 16
ha_note = "Arial", 9, "italic"
# Headings
ha_head10 = "Arial", 10, "bold"
ha_head11 = "Arial", 11, "bold"
ha_head12 = "Arial", 12, "bold"
ha_head14 = "Arial", 14, "bold"
ha_head16 = "Arial", 16, "bold"


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
DEFAULT_COLORS = {                    # Should never change - used to reset the pallette if the user requests this
    "home_bg": "#DFFFFF",           # Lup_LupT_ID=2, Lup_Seq=1, Main Form BG
    "home_test_bg": "#F0D0D0",      # Lup_LupT_ID=2, Lup_Seq=2, Main Form BG - TEST mode
    "flag_y_bg": "#FFFF80",         # Lup_LupT_ID=2, Lup_Seq=3, Flag BG - Yellow
    "flag_g_bg": "#80FF80",         # Lup_LupT_ID=2, Lup_Seq=4, Flag BG - Green
    "flag_b_bg": "#89FFFF",         # Lup_LupT_ID=2, Lup_Seq=5, Flag BG - Blue
    "flag_dd_bg": "#FFC993",        # Lup_LupT_ID=2, Lup_Seq=6, Flag BG - Drill Down Marker
    "flag_mk_bg": "#ED87ED",        # Lup_LupT_ID=2, Lup_Seq=7, Flag BG - Reconciliation Marker
    "tran_wk_bg": "#FFFFFF",        # Lup_LupT_ID=2, Lup_Seq=8, Transaction Treeview - Weekday row BG
    "tran_we_bg": "#DEE7DF",        # Lup_LupT_ID=2, Lup_Seq=9, Transaction Treeview - Weekend row BG
    "dtot_bg": "#5D5D5D",           # Lup_LupT_ID=2, Lup_Seq=10, Daily Totals row BG
    "dtot_ol_bg": "#FADBFD",        # Lup_LupT_ID=2, Lup_Seq=11, Daily Totals row - Over Limit - BG
    "tab_bg": "#FFFFFF",            # Lup_LupT_ID=2, Lup_Seq=12, Not Selected Tab BG
    "tab_act_bg": "#803624",        # Lup_LupT_ID=2, Lup_Seq=13, Selected Tab BG
    "act_but_bg": "#E0FFE0",        # Lup_LupT_ID=2, Lup_Seq=14, Active Button BG
    "act_but_hi_bg": "#98FB98",     # Lup_LupT_ID=2, Lup_Seq=15, Active Button - Highlighted - BG
    "last_stat_bg": "#E0E0E0",      # Lup_LupT_ID=2, Lup_Seq=16, Last Statement Balance field BG
    "title1_bg": "#5D5D5D",         # Lup_LupT_ID=2, Lup_Seq=17, Title/Header BG - AHB Header row
    "title2_bg": "#E0E0E0",         # Lup_LupT_ID=2, Lup_Seq=18, Title/Header BG - Transaction Treeview Header row
    "title3_bg": "#000000",         # Lup_LupT_ID=2, Lup_Seq=19, Title/Header BG
    "exit_but_bg": "#ADD8E6",       # Lup_LupT_ID=2, Lup_Seq=20, Exit/Close Buttons BG
    "del_but_bg": "#DD0000",        # Lup_LupT_ID=2, Lup_Seq=21, Delete Button BG
    
    "complete_tx": "#000000",       # Lup_LupT_ID=3, Lup_Seq=3, Transaction Treeview text - Completed transactions
    "pending_tx": "#0000FF",        # Lup_LupT_ID=3, Lup_Seq=4, Transaction Treeview text - Pending/Processing transactions
    "forecast_tx": "#800000",       # Lup_LupT_ID=3, Lup_Seq=5, Transaction Treeview text - Forecast transactions
    "dtot_tx": "#FFFFFF",           # Lup_LupT_ID=3, Lup_Seq=6, Daily Totals row text
    "dtot_ol_tx": "#DD0000",        # Lup_LupT_ID=3, Lup_Seq=7, Daily Totals row - Over Limit - text
    "tab_tx": "#000000",            # Lup_LupT_ID=3, Lup_Seq=8, Not Selected Tab text
    "tab_act_tx": "#FFFFFF",        # Lup_LupT_ID=3, Lup_Seq=9, Selected Tab text
    "act_but_tx": "#000000",        # Lup_LupT_ID=3, Lup_Seq=10, Active Button text
    "act_but_hi_tx": "#000000",     # Lup_LupT_ID=3, Lup_Seq=11, Active Button - Highlighted - text
    "last_stat_tx": "#000000",      # Lup_LupT_ID=3, Lup_Seq=12, Last Statement Balance field text
    "title1_tx": "#FFFFFF",         # Lup_LupT_ID=3, Lup_Seq=13, Title/Header text
    "title2_tx": "#000000",         # Lup_LupT_ID=3, Lup_Seq=14, Title/Header text
    "title3_tx": "#FFFFFF",         # Lup_LupT_ID=3, Lup_Seq=15, Title/Header text
    "exit_but_tx": "#000000",       # Lup_LupT_ID=3, Lup_Seq=16, Exit/Close Buttons text
    "del_but_tx": "#000000",        # Lup_LupT_ID=3, Lup_Seq=17, Delete Button text
    # Legacy keys - to be kept as not user configurable
    "white": "#FFFFFF",             # default widget background
    "black": "#000000",             # default enabled widget text
    "grey": "#5D5D5D",              # default disabled widget text
}

COLORS = {                            # The active pallette - loaded from database at startup, can be edited by user, and saved to database
    # Background colours
    "home_bg": "#DFFFFF",           # very_pale_blue    Main Form BG                            seashell
    "home_test_bg": "#F0D0D0",      # very_pale_pink    Main Form BG - TEST mode
    "flag_y_bg": "#FFFF80",         # yellow            Flag BG - Yellow
    "flag_g_bg": "#80FF80",         # green             Flag BG - Green
    "flag_b_bg": "#89FFFF",         # cyan              Flag BG - Blue
    "flag_dd_bg": "#FFC993",        # orange            Flag BG - Drill Down Marker             bright orange
    "flag_mk_bg": "#ED87ED",        # marker            Flag BG - Reconciliation Marker         bright_pink
    "tran_wk_bg": "#FFFFFF",        # white             Transaction Treeview - Weekday row BG
    "tran_we_bg": "#DEE7DF",        # oldlace           Transaction Treeview - Weekend row BG
    "dtot_bg": "#5D5D5D",           # dark_grey         Daily Totals row BG
    "dtot_ol_bg": "#FADBFD",        # pink              Daily Totals row - Over Limit - BG
    "tab_bg": "#FFFFFF",            # pale_grey         Not Selected Tab BG
    "tab_act_bg": "#803624",        # dark_brown        Selected Tab BG
    "act_but_bg": "#E0FFE0",        # pale_green        Active Button BG
    "act_but_hi_bg": "#98FB98",     # pale_green2       Active Button - Highlighted - BG
    "last_stat_bg": "#E0E0E0",      # grey              Last Statement Balance field BG
    "title1_bg": "#5D5D5D",         # Lup_LupT_ID=2, Lup_Seq=17, Title/Header BG - AHB Header row
    "title2_bg": "#E0E0E0",         # Lup_LupT_ID=2, Lup_Seq=18, Title/Header BG - Transaction Treeview Header row
    "title3_bg": "#000000",         # Lup_LupT_ID=2, Lup_Seq=19, Title/Header BG
    "exit_but_bg": "#ADD8E6",       # pale_blue         Exit/Close Buttons BG
    "del_but_bg": "#DD0000",        # red               Delete Button BG
    # Text colours
    "complete_tx": "#000000",       # Lup_LupT_ID=3, Lup_Seq=3, Transaction Treeview text - Completed transactions
    "pending_tx": "#0000FF",        # blue              Transaction Treeview text - Pending/Processing transactions
    "forecast_tx": "#800000",       # maroon            Transaction Treeview text - Forecast transactions
    "dtot_tx": "#FFFFFF",           # Lup_LupT_ID=3, Lup_Seq=6, Daily Totals row text
    "dtot_ol_tx": "#DD0000",        # red2              Daily Totals row - Over Limit - text
    "tab_tx": "#000000",            # Lup_LupT_ID=3, Lup_Seq=8, Not Selected Tab text
    "tab_act_tx": "#FFFFFF",        # Lup_LupT_ID=3, Lup_Seq=9, Selected Tab text
    "act_but_tx": "#000000",        # Lup_LupT_ID=3, Lup_Seq=10, Active Button text
    "act_but_hi_tx": "#000000",     # Lup_LupT_ID=3, Lup_Seq=11, Active Button - Highlighted - text
    "last_stat_tx": "#000000",      # Lup_LupT_ID=3, Lup_Seq=12, Last Statement Balance field text
    "title1_tx": "#FFFFFF",         # Lup_LupT_ID=3, Lup_Seq=13, Title/Header text
    "title2_tx": "#000000",         # Lup_LupT_ID=3, Lup_Seq=14, Title/Header text
    "title3_tx": "#FFFFFF",         # Lup_LupT_ID=3, Lup_Seq=15, Title/Header text
    "exit_but_tx": "#000000",       # Lup_LupT_ID=3, Lup_Seq=16, Exit/Close Buttons text
    "del_but_tx": "#000000",        # Lup_LupT_ID=3, Lup_Seq=17, Delete Button text
    # Legacy keys - to be kept as not user configurable
    "white": "#FFFFFF",
    "black": "#000000",
    # Legacy keys (kept during migration)
    "oldlace": "#DEE7DF",
    "yellow": "#FFFF80",
    "green": "#80FF80",
    "cyan": "#B9FFFF",
    "marker": "#ED87ED",
    "pale_blue": "#ADD8E6",
    "very_pale_blue": "#DFFFFF",
    "very_pale_pink": "#F0D0D0",
    "red": "#DD0000",
    "dark_grey": "#5D5D5D",
    "pink": "#FADBFD",
    "dark_brown": "#803624",
    "orange": "#FFC993",
    "grey": "#E0E0E0",
    "pale_green": "#E0FFE0",
    "pale_grey": "#F0F0F0",
    "blue": "#0000FF",
    "maroon": "#800000",
    "red2": "#FF0000",
    "darkgray": "#707070",
    "pink2": "#FFE0EB",
    "darkbrown2": "#802020",
    "seashell": "#F0F8F8",
    "yellow2": "#FFFF00",
    "pale_green2": "#98FB98",
    "cyan2": "#A6FFFF",
    "bright_orange": "#FFC993",
    "bright_pink": "#ED87ED",
    "oldlace2": "#F2EDE0"
}

# Legacy color mappings (reference for migration, not used at runtime)
LEGACY_COLOR_MAP = {
    "white": "tran_wk_bg",
    "oldlace": "tran_we_bg",
    "yellow": "flag_y_bg",
    "green": "flag_g_bg",
    "cyan": "flag_b_bg",
    "marker": "flag_mk_bg",
    "pale_blue": "exit_but_bg",
    "very_pale_blue": "home_bg",
    "very_pale_pink": "home_test_bg",
    "red": "del_but_bg",
    "dark_grey": "dtot_bg",
    "pink": "dtot_ol_bg",
    "black": "normal_tx",
    "dark_brown": "tab_act_bg",
    "orange": "flag_dd_bg",
    "grey": "last_stat_bg",
    "pale_green": "act_but_bg",
    "pale_grey": "tab_bg",
    "blue": "pending_tx",
    "maroon": "forecast_tx",
    "red2": "dtot_ol_tx",
    "darkgray": "disabled_tx",
    "pink2": "dtot_ol_bg",
    "darkbrown2": "tab_act_bg",
    "seashell": "home_bg",
    "yellow2": "flag_y_bg",
    "pale_green2": "act_but_hi_bg",
    "cyan2": "flag_b_bg",
    "bright_orange": "flag_dd_bg",
    "bright_pink": "flag_mk_bg",
    "oldlace2": "tran_we_bg"
}

#TEXT_COLORS = {
#    "Unknown": "gray",                              # "disabled_tx"
#    "Forecast": "#800040",          # Brown       #"forecast_tx"
#    "Processing": "#0000FF",        # Blue        #"pending_tx"
#    "Complete": "#000000"           # Black       #"complete_tx"
#}

def load_colors_from_db(db_path):
    """
    Load colors from Lookups table and update COLORS dictionary, falling back to DEFAULT_COLORS if records are missing.
    
    Args:
        db_path (str): Path to SQLite database
    """
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        # Expected counts
        expected_counts = {2: 21, 3: 17}  # Lup_LupT_ID: expected Lup_Seq count
        color_map = {
            (2, 1): "home_bg",
            (2, 2): "home_test_bg",
            (2, 3): "flag_y_bg",
            (2, 4): "flag_g_bg",
            (2, 5): "flag_b_bg",
            (2, 6): "flag_dd_bg",
            (2, 7): "flag_mk_bg",
            (2, 8): "tran_wk_bg",
            (2, 9): "tran_we_bg",
            (2, 10): "dtot_bg",
            (2, 11): "dtot_ol_bg",
            (2, 12): "tab_bg",
            (2, 13): "tab_act_bg",
            (2, 14): "act_but_bg",
            (2, 15): "act_but_hi_bg",
            (2, 16): "last_stat_bg",
            (2, 17): "title1_bg",
            (2, 18): "title2_bg",
            (2, 19): "title3_bg",
            (2, 20): "exit_but_bg",
            (2, 21): "del_but_bg",
            (3, 1): "normal_tx",
            (3, 2): "disabled_tx",
            (3, 3): "complete_tx",
            (3, 4): "pending_tx",
            (3, 5): "forecast_tx",
            (3, 6): "dtot_tx",
            (3, 7): "dtot_ol_tx",
            (3, 8): "tab_tx",
            (3, 9): "tab_act_tx",
            (3, 10): "act_but_tx",
            (3, 11): "act_but_hi_tx",
            (3, 12): "last_stat_tx",
            (3, 13): "title1_tx",
            (3, 14): "title2_tx",
            (3, 15): "title3_tx",
            (3, 16): "exit_but_tx",
            (3, 17): "del_but_tx"
        }

        for lup_type_id in (2, 3):
            cursor.execute("SELECT Lup_Seq, Lup_Desc FROM Lookups WHERE Lup_LupT_ID = ? ORDER BY Lup_Seq", (lup_type_id,))
            records = cursor.fetchall()
            found_seqs = {seq for seq, _ in records}
            expected_seqs = set(range(1, expected_counts[lup_type_id] + 1))
            missing_seqs = expected_seqs - found_seqs
            if missing_seqs:
                logger.warning(f"Missing Lookups records for Lup_LupT_ID={lup_type_id}, Lup_Seq={missing_seqs}")

            for lup_seq, lup_desc in records:
                key = color_map.get((lup_type_id, lup_seq))
                if key:
                    # Normalize color format: convert 0xRRGGBB to #RRGGBB
                    color = lup_desc.replace("0x", "#") if lup_desc.startswith("0x") else lup_desc
                    COLORS[key] = color
                    #logger.debug(f"Updated COLORS[{key}] = {color}")
                else:
                    logger.warning(f"No mapping for Lup_LupT_ID={lup_type_id}, Lup_Seq={lup_seq}")

            # Fallback to DEFAULT_COLORS for missing records
            for lup_seq in missing_seqs:
                key = color_map.get((lup_type_id, lup_seq))
                if key:
                    COLORS[key] = DEFAULT_COLORS[key]
                    # Update legacy key if it exists
                    legacy_keys = [k for k, v in LEGACY_COLOR_MAP.items() if v == key]
                    for legacy_key in legacy_keys:
                        COLORS[legacy_key] = DEFAULT_COLORS[key]
                        logger.debug(f"Fallback to DEFAULT_COLORS[{legacy_key}] = {DEFAULT_COLORS[key]}")
                    logger.debug(f"Fallback to DEFAULT_COLORS[{key}] = {DEFAULT_COLORS[key]}")

        conn.close()
    except Exception as e:
        logger.error(f"Failed to load colors from database: {e}")
        COLORS.update(DEFAULT_COLORS)  # Fallback to defaults
        logger.debug("Reverted to DEFAULT_COLORS")

def load_db_settings():
    global master_bg, COLORS, scaling_factor

    # Set up logging
    logger = logging.getLogger('HA.db_setup')
    
    # Get correct DB and set master_bg colour
    if CONFIG['APP_ENV'] == 'test':
        dbpath = get_config('DB_PATH_TEST')
    else:
        dbpath = get_config('DB_PATH')
        
    logger.debug(f"Opening database: {dbpath}")

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
        
    """Load master background colour setting from the database."""
    if CONFIG['APP_ENV'] == 'test':
        cursor.execute("SELECT Lup_Desc FROM Lookups WHERE Lup_LupT_ID=2 AND Lup_Seq=2")
        master_bg = cursor.fetchone()
    else:
        cursor.execute("SELECT Lup_Desc FROM Lookups WHERE Lup_LupT_ID=2 AND Lup_Seq=1")
        master_bg = cursor.fetchone()
    logger.debug(f"master_bg: {master_bg}")
    
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
    