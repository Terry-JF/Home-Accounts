import os
from datetime import datetime
import sqlite3
import logging
from gc_utils import get_access_token, fetch_transactions, OUTPUT_DIR
from rules_engine import process_transactions
from config import get_config

# Runs in background when launched by Windows Scheduler

os.makedirs(get_config('LOG_DIR'), exist_ok=True)
os.makedirs(get_config('OUTPUT_DIR'), exist_ok=True)

def main():
    logger = logging.getLogger('HA.transactions')
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = os.path.join(get_config('LOG_DIR'), f"{timestamp}_fetch_transactions.log")
    file_handler = logging.FileHandler(log_file)
    file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
    logger.addHandler(file_handler)
    logger.setLevel(logging.DEBUG)    
    
    try:
        # Running as script
        base_path = os.path.dirname(os.path.dirname(__file__))
        db_path = os.path.join(base_path, "data", "HAdata.db")
        conn = sqlite3.connect(db_path)
        cur = conn.cursor()
        logger.info(f"Database opened {db_path}")
        
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
            json_path = os.path.join(OUTPUT_DIR, f"{timestamp}_transactions_{acc_id}.json")
            fetch_transactions(access_token, requisition_id, json_path, fetch_days, conn, acc_id)
            logger.debug(f"Processing transactions for account {acc_id}")
            process_transactions(json_path, conn, acc_id)
            logger.debug(f"Completed processing for account {acc_id}")
            
        conn.close()
    except Exception as e:
        logger.error(f"Error in fetch_bank_transactions: {str(e)}")
    finally:
        logger.removeHandler(file_handler)
        file_handler.close()
        
if __name__ == "__main__":
    main()
