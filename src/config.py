# src/config.py
# Centralized configuration for Home Accounts project

import os
#from dotenv import load_dotenv
import tkinter as tk
import sqlite3
import logging
#from dotenv import load_dotenv
from logging.handlers import TimedRotatingFileHandler
from datetime import datetime, timedelta
import shutil
import glob

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

scroll_row = 1 # Global - default, will be overwritten by DB value

# Base directory for the project
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Configuration dictionary
CONFIG = {}

# Colour palette dictionaries
DEFAULT_COLORS = {                    # Should never change - used to reset the pallette if the user requests this
    "home_bg": "#DFFFFF",           # Lup_LupT_ID=2, Lup_Seq=1, Main Form BG
    "home_test_bg": "#F7E3E3",      # Lup_LupT_ID=2, Lup_Seq=2, Main Form BG - TEST mode
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
    "act_but_bg": "#C7DFFE",        # Lup_LupT_ID=2, Lup_Seq=14, Active Button BG
    "act_but_hi_bg": "#98FB98",     # Lup_LupT_ID=2, Lup_Seq=15, Active Button - Highlighted - BG
    "last_stat_bg": "#E0E0E0",      # Lup_LupT_ID=2, Lup_Seq=16, Last Statement Balance field BG
    "title1_bg": "#5D5D5D",         # Lup_LupT_ID=2, Lup_Seq=17, Title/Header BG - AHB Header row
    "title2_bg": "#E0E0E0",         # Lup_LupT_ID=2, Lup_Seq=18, Title/Header BG - Transaction Treeview Header row
    "title3_bg": "#000000",         # Lup_LupT_ID=2, Lup_Seq=19, Title/Header BG
    "exit_but_bg": "#ADD8E6",       # Lup_LupT_ID=2, Lup_Seq=20, Exit/Close Buttons BG
    "del_but_bg": "#DD4040",        # Lup_LupT_ID=2, Lup_Seq=21, Delete Button BG
    
    "normal_tx": "#FFFFFF",         # Lup_LupT_ID=3, Lup_Seq=1, Enabled widget text colour
    "disabled_tx": "#5D5D5D",       # Lup_LupT_ID=3, Lup_Seq=2, Disabled widget text colour
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

COLORS = {                            # The active pallette - loaded from database at startup, can be edited by user, and re-saved to database
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
    "normal_tx": "#000000",         # Lup_LupT_ID=3, Lup_Seq=1, Enabled widget text colour
    "disabled_tx": "#5D5D5D",       # Lup_LupT_ID=3, Lup_Seq=2, Disabled widget text colour
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
    "white": "#FFFFFF",             # default widget background
    "black": "#000000",             # default enabled widget text
    "grey": "#5D5D5D",              # default disabled widget text
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
    #"grey": "#E0E0E0",
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

# Global icon cache
ICON_CACHE = {}

# Legacy color mappings (reference for migration, not used at runtime) - is this used anywhere?
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
                        #logger.debug(f"Fallback to DEFAULT_COLORS[{legacy_key}] = {DEFAULT_COLORS[key]}")
                    #logger.debug(f"Fallback to DEFAULT_COLORS[{key}] = {DEFAULT_COLORS[key]}")

        conn.close()
    except Exception as e:
        logger.error(f"Failed to load colors from database: {e}")
        COLORS.update(DEFAULT_COLORS)  # Fallback to defaults
        #logger.debug("Reverted to DEFAULT_COLORS")

def load_db_settings():
    global master_bg, COLORS, scaling_factor
    global scroll_row

    # Set up logging
    logger = logging.getLogger('HA.db_setup')
    
    # Get correct DB and set master_bg colour
    dbpath = get_config('DB_PATH')
    #logger.debug(f"Opening database: {dbpath}")
    
    """Backup the current database."""
    backup_dir = get_config('BACKUP_PATH')
    if os.path.exists(dbpath):
        try:
            # Create timestamped backup path: BACKUP/YYYY/MM/HAtest.db-YYYY.MM.DD.HH.MM.SS
            now = datetime.now()
            year = now.strftime('%Y')
            month = now.strftime('%m')
            timestamp = now.strftime('%Y.%m.%d.%H.%M.%S')
            db_filename = os.path.basename(dbpath)  # e.g., HAtest.db
            backup_subdir = os.path.join(backup_dir, year, month)
            os.makedirs(backup_subdir, exist_ok=True)
            backup_filepath = os.path.join(backup_subdir, f"{db_filename}-{timestamp}")
            
            # Copy the database file
            shutil.copy2(dbpath, backup_filepath)
            logger.debug(f"Database backed up to: {backup_filepath}")
            
            # Verify backup file exists and is not empty
            if os.path.exists(backup_filepath) and os.path.getsize(backup_filepath) > 0:
                logger.debug(f"Backup verified: {backup_filepath}, size: {os.path.getsize(backup_filepath)} bytes")
            else:
                logger.error(f"Backup failed or empty: {backup_filepath}")
        except (OSError, shutil.Error) as e:
            logger.error(f"Failed to backup database to {backup_filepath}: {e}")
    else:
        logger.warning(f"Database file not found for backup: {dbpath}")
    
    """Load settings from the database."""
    try:
        with sqlite3.connect(dbpath) as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT key, value FROM settings")
            for key, value in cursor.fetchall():
                try:
                    if key in ['EXPIRY_WARNING_DAYS', 'LOG_DAYS_TO_KEEP', 'TOKEN_MARGIN', 'REQUISITION_VALIDITY_DAYS', 'IMPORT_DAYS_TO_KEEP']:
                        CONFIG[key] = int(value)
                    else:
                        CONFIG[key] = value
                except ValueError:
                    logger.error(f"Invalid value for {key}: {value}")
    except sqlite3.Error as e:
        logger.error(f"Failed to load settings from database: {e}")
        
    cursor.execute("SELECT value FROM settings WHERE key = 'SCROLL_ROW'")
    row = cursor.fetchone()
    if row:
        scroll_row = int(row[0])   # Updates the global
    #logger.debug(f"load_db_settings scroll_row = {scroll_row}")
        
    """Load master background colour setting from the database."""
    try:
        with sqlite3.connect(dbpath) as conn:
            cursor = conn.cursor()
            if CONFIG['APP_ENV'] == 'test':
                cursor.execute("SELECT Lup_Desc FROM Lookups WHERE Lup_LupT_ID=2 AND Lup_Seq=2")
            else:
                cursor.execute("SELECT Lup_Desc FROM Lookups WHERE Lup_LupT_ID=2 AND Lup_Seq=1")
            master_bg_result = cursor.fetchone()
            global master_bg
            master_bg = master_bg_result[0] if master_bg_result else ('#F0D0D0',)
            #logger.debug(f"master_bg: {master_bg}")
    except sqlite3.Error as e:
        logger.error(f"Failed to load master background color: {e}")
        master_bg = ('#F0D0D0',)
    
def setup_logging():
    """Configure logging with file and optional console handlers, and clean up old backups."""
    logger = logging.getLogger('HA')
    logger.setLevel(log_levels.get(get_config('APP_LOG_LEVEL'), logging.DEBUG))

    # Clear any existing handlers to prevent duplicates
    logger.handlers.clear()

    # File handler with rotation
    log_file = os.path.join(get_config('LOG_PATH'), 'app.log')
    file_handler = TimedRotatingFileHandler(
        log_file,
        when='midnight',
        interval=1,
        backupCount=get_config('LOG_DAYS_TO_KEEP')
    )
    file_handler.setFormatter(logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s'))
    logger.addHandler(file_handler)

    # Console handler (optional, only for interactive runs)
    if get_config('APP_DEBUG'):
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s'))
        logger.addHandler(console_handler)

    # Clean up old log files and database backups
    cutoff_date = datetime.now() - timedelta(days=get_config('LOG_DAYS_TO_KEEP')+1)
    #logger.debug(f"Cleaning up logs and backups older than {cutoff_date.strftime('%Y-%m-%d')}")

    # Clean up old log files
    log_dir = get_config('LOG_PATH')
    for log_file in glob.glob(os.path.join(log_dir, 'app.log.*')):
        try:
            file_mtime = datetime.fromtimestamp(os.path.getmtime(log_file))
            if file_mtime < cutoff_date:
                os.remove(log_file)
                #logger.debug(f"Deleted old log file: {log_file}")
        except (OSError, ValueError) as e:
            logger.error(f"Failed to delete old log file {log_file}: {e}")

def load_icon_cache():
    """Load and cache all icons used in the HA program to prevent garbage collection."""
    icon_definitions = [
        # Map icon key to filename (replace with actual icon filenames)
        ("drag",        "1_drag-48.png"),
        ("edit",        "2_edit-48.png"),
        ("trash",       "3_trash-48.png"),
        ("up_l",        "3_up-48.png"),
        ("down_l",      "4_down-48.png"),
        ("see_match",   "4_see_match-48.png"),
        ("apply_rule",  "5_apply_rule-48.png"),
        ("duplicate",   "6_duplicate-48.png"),
        ("up_s",        "7_up-24.png"),
        ("down_s",      "8_down-24.png"),
        ("checked",     "checked_16.png"),
        ("unchecked",   "unchecked_16.png"),
        ("radio_0_l",   "radio_0_32.png"),
        ("radio_1_l",   "radio_1_32.png"),
        ("radio_0_s",   "radio_0_16.png"),
        ("radio_1_s",   "radio_1_16.png")
    ]
    
    global ICON_CACHE
    ICON_CACHE.clear()  # Clear any existing cache
    for key, filename in icon_definitions:
        try:
            relative_path = (f"src/icons/{filename}")
            icon_path = os.path.join(BASE_DIR, relative_path)
            ICON_CACHE[key] = tk.PhotoImage(file=icon_path)
            #logger.debug(f"Loaded icon: {key} ({filename}) - icon_path = {icon_path}")
        except tk.TclError as e:
            logger.error(f"Failed to load icon {filename}: {e}")
            ICON_CACHE[key] = None  # Store None to avoid repeated attempts
    
    #logger.debug(f"Icon cache initialized with {len(ICON_CACHE)} icons")

# Clean up old database backups
def cleanup_old_files():
    """Clean up old database backups."""
    # Set cutoff date to the first of the current month
    current_date = datetime.now()
    cutoff_date = datetime(current_date.year, current_date.month, 1)
    #logger.debug(f"Cleaning up database backups older than {cutoff_date.strftime('%Y-%m-%d')}")
    
    # Clean up old database backups
    backup_dir = CONFIG['BACKUP_DIR']
    if os.path.exists(backup_dir):
        for year_dir in glob.glob(os.path.join(backup_dir, '[0-9][0-9][0-9][0-9]')):
            try:
                year = int(os.path.basename(year_dir))
                for month_dir in glob.glob(os.path.join(year_dir, '[0-1][0-9]')):
                    month = int(os.path.basename(month_dir))
                    dir_date = datetime(year, month, 1)
                    # Only delete directories strictly older than the cutoff month
                    if dir_date < cutoff_date and (year < cutoff_date.year or (year == cutoff_date.year and month < cutoff_date.month)):
                        #logger.debug(f"Deleting old backup directory: {month_dir} (dir_date={dir_date})")
                        shutil.rmtree(month_dir)
                    else:
                        # Check individual backup files in the month directory
                        for backup_file in glob.glob(os.path.join(month_dir, '*.db-*')):
                            try:
                                # Extract date from filename (e.g., HAtest.db-2025.09.08.21.33.26)
                                timestamp_str = os.path.basename(backup_file).split('-')[-1]
                                file_date = datetime.strptime(timestamp_str, '%Y.%m.%d.%H.%M.%S')
                                if file_date < cutoff_date:
                                    #logger.debug(f"Deleting old backup file: {backup_file} (file_date={file_date})")
                                    os.remove(backup_file)
                            except (ValueError, OSError) as e:
                                logger.error(f"Failed to delete old backup file {backup_file}: {e}")
            except (ValueError, OSError) as e:
                logger.error(f"Failed to process backup directory {year_dir}: {e}")

def get_config(key):
    """Get a configuration value by key."""
    value = CONFIG.get(key)
    if value is None:
        logger.warning(f"CONFIG key not found: {key}")
    return value

def init_config(source, gc_env=None):
    """Initialize configuration - FORCE .env reload every time"""
    # CRITICAL: Force reload .env from disk EVERY STARTUP
    from dotenv import load_dotenv
    env_path = r'C:\HA-Project\.env'
    load_dotenv(dotenv_path=env_path, override=True, verbose=True)
    print("---First line in log after restart---")
    print(f"RELOADED .env: APP_ENV={os.getenv('APP_ENV')}")
    
    CONFIG.clear
    CONFIG.update({
        'APP_ENV': os.getenv('APP_ENV', 'test'),
        'APP_DEBUG': os.getenv('APP_DEBUG', 'True').lower() == 'true',
        'APP_LOG_LEVEL': os.getenv('APP_LOG_LEVEL', 'debug').lower(),
        'PROG_DIR': os.getenv('PROG_DIR', 'C:/HA-Project/'),
        'DATA_DIR': os.getenv('DATA_DIR', 'C:/HA-Data/'),
        'TEST_DIR': os.getenv('TEST_DIR', 'TEST/'),
        'LIVE_DIR': os.getenv('LIVE_DIR', 'LIVE/'),
        'DB_DIR': os.getenv('DB_DIR', 'DATABASE/'),
        'LOG_DIR': os.getenv('LOG_DIR', 'LOGS/'),
        'ILOG_DIR': os.getenv('ILOG_DIR', 'LOGS/IMPORT/'),
        'FLOG_DIR': os.getenv('FLOG_DIR', 'LOGS/FETCH/'),
        'BANK_DIR': os.getenv('BANK_DIR', 'BANK_TRANSACTIONS/'),
        'BACKUP_DIR': os.getenv('BACKUP_DIR', 'BACKUP/'),
        'DB_LIVE': os.getenv('DB_LIVE', 'HAdata.db'),
        'DB_TEST': os.getenv('DB_TEST', 'HAtest.db'),
        'API_BASE_URL': os.getenv('API_BASE_URL', 'https://bankaccountdata.gocardless.com/api/v2'),
        'VENV_DIR': os.getenv('VENV_DIR', 'C:/HA-Project/.venv'),
        # Below values are defaults - actuals are loaded from Settings table in db, but prob the same values
        'EXPIRY_WARNING_DAYS': 10,
        'LOG_DAYS_TO_KEEP': 10,
        'TOKEN_MARGIN': 60,
        'REQUISITION_VALIDITY_DAYS': 90,
        'IMPORT_DAYS_TO_KEEP': 30
    })
    
    # Check if called from main.py or fetch_bank_trans.py
    if source == "GC":
        print(f"Command-line override: forcing APP_ENV = {gc_env}")
        # Override both environment and CONFIG
        os.environ['APP_ENV'] = gc_env
        CONFIG['APP_ENV'] = gc_env
    
    # Compute paths based on APP_ENV
    data_dir = get_config('DATA_DIR')
    if get_config('APP_ENV') == 'test':
        env_dir = get_config('TEST_DIR')
        db_filename = get_config('DB_TEST')
    else:
        env_dir = get_config('LIVE_DIR')
        db_filename = get_config('DB_LIVE')

    # Build paths
    base_path = os.path.join(data_dir, env_dir)
    CONFIG['DB_PATH'] = os.path.join(base_path, get_config('DB_DIR'), db_filename)
    CONFIG['LOG_PATH'] = os.path.join(base_path, get_config('LOG_DIR'))
    CONFIG['ILOG_PATH'] = os.path.join(base_path, get_config('ILOG_DIR'))
    CONFIG['FLOG_PATH'] = os.path.join(base_path, get_config('FLOG_DIR'))
    CONFIG['BANK_PATH'] = os.path.join(base_path, get_config('BANK_DIR'))
    CONFIG['BACKUP_PATH'] = os.path.join(base_path, get_config('DB_DIR'), get_config('BACKUP_DIR'))

    # Ensure directories exist
    for dir_path in [CONFIG['DB_PATH'], CONFIG['LOG_PATH'], CONFIG['ILOG_PATH'], CONFIG['BANK_PATH']]:
        os.makedirs(os.path.dirname(dir_path) if dir_path.endswith('.db') else dir_path, exist_ok=True)

    # Set up logging
    setup_logging()
    
    logger = logging.getLogger('HA.init')
    #logger.debug("++++++++++++++++++++++++ START OF NEW SESSION ++++++++++++++++++++++++")
    #logger.debug(f"DB_PATH: {CONFIG['DB_PATH']}")
    #logger.debug(f"LOG_PATH: {CONFIG['LOG_PATH']}")
    #logger.debug(f"ILOG_PATH: {CONFIG['ILOG_PATH']}")
    #logger.debug(f"FLOG_PATH: {CONFIG['FLOG_PATH']}")
    #logger.debug(f"BANK_PATH: {CONFIG['BANK_PATH']}")
    
    load_db_settings()
    #logger.debug("Configuration initialized")
    
def update_master_bg(colour):
    global master_bg
    master_bg = colour
    
    