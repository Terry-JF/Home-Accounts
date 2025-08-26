# src/fetch_bank_trans.py
# Script to fetch bank transactions via GoCardless and process them

import logging
import os
from datetime import datetime
from src.config import get_config, init_config
from src.db import open_db
from src.gc_utils import get_access_token, fetch_transactions
from src.rules_engine import process_transactions

# Initialize configuration (sets up logging and directories)
init_config()

# Get logger
logger = logging.getLogger('HA.transactions')

def main():
    """Fetch and process bank transactions."""
    logger.debug("Starting fetch_bank_trans.py")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    try:
        # Open database using db.py's logic
        conn, cur = open_db()
        logger.info(f"Database opened at {get_config('DB_PATH')}")
        
        cur.execute("""
            SELECT Acc_ID, Requisition_ID, Fetch_Days
            FROM GC_Account
            WHERE Active = 1 AND Link_Status = 'Linked'
        """)
        accounts = cur.fetchall()
        
        access_token = get_access_token(conn)
        if not access_token:
            logger.error("Failed to get access token.")
            return
        
        for acc_id, requisition_id, fetch_days in accounts:
            json_path = os.path.join(get_config('BANK_DIR'), f"{timestamp}_transactions_{acc_id}.json")
            fetch_transactions(access_token, requisition_id, json_path, fetch_days, conn, acc_id)
            logger.debug(f"Processing transactions for account {acc_id}")
            process_transactions(json_path, conn, acc_id)
            logger.debug(f"Completed processing for account {acc_id}")
            
        conn.close()
    except Exception as e:
        logger.error(f"Error in fetch_bank_trans: {str(e)}")
        raise

if __name__ == "__main__":
    main()
    
    
    
    