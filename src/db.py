import sqlite3
import os
import sys
import shutil
from datetime import datetime
from calendar import monthrange
import logging
from config import CONFIG, get_config

# Set up logging
logger = logging.getLogger('HA.db')

# Open and Close DB connection
def open_db():
    """Open the SQLite database, handling both script and .exe modes."""
    if getattr(sys, 'frozen', False):
        # Running as .exe
        base_path = sys._MEIPASS
        db_source = os.path.join(base_path, "data", "HAdata.db")
        # Use user data folder for persistence
        user_data_dir = os.path.join(os.getenv('APPDATA'), 'HA2', 'data')
        os.makedirs(user_data_dir, exist_ok=True)
        db_path = os.path.join(user_data_dir, "HAdata.db")
        # Copy template HAdata.db if it doesn't exist
        if not os.path.exists(db_path):
            shutil.copy(db_source, db_path)
    else:
        # Running as script
        # Get correct DB
        if CONFIG['APP_ENV'] == 'test':
            db_path = get_config('DB_PATH_TEST')
        else:
            db_path = get_config('DB_PATH')
    
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        logger.debug(f"Database open: {db_path}")
        return conn, cursor
    except sqlite3.Error as e:
        raise Exception(f"Failed to open database: {e}")    
    
def close_db(conn):
    if conn:
        conn.close()
    logger.debug("Database closed:")

# Transaction Table
def insert_transaction(cursor, conn, tr_type, day, month, year, status, flag, amount, desc, acc_from, acc_to, cat_pid=None, subcat_cid=None, reg_id=0):
    cursor.execute("INSERT INTO Trans (Tr_Type, Tr_Reg_ID, Tr_DOW, Tr_Day, Tr_Month, Tr_Year, Tr_Stat, Tr_Query_Flag, Tr_Amount, Tr_Desc, Tr_Acc_From, Tr_Acc_To, Tr_Exp_ID, Tr_ExpSub_ID) "
                "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                (tr_type, reg_id, 0, day, month, year, status, flag, amount, desc, acc_from, acc_to, cat_pid, subcat_cid))
    conn.commit()
    return cursor.lastrowid

def update_transaction(cursor, conn, tr_id, tr_type, day, month, year, status, flag, amount, desc, acc_from, acc_to, cat_pid=None, subcat_cid=None):
    cursor.execute("UPDATE Trans SET Tr_Type=?, Tr_Reg_ID=?, Tr_DOW=?, Tr_Day=?, Tr_Month=?, Tr_Year=?, Tr_Stat=?, Tr_Query_Flag=?, Tr_Amount=?, Tr_Desc=?, Tr_Acc_From=?, Tr_Acc_To=?, Tr_Exp_ID=?, Tr_ExpSub_ID=? "
                "WHERE Tr_ID=?",
                (tr_type, 0, 0, day, month, year, status, flag, amount, desc, acc_from, acc_to, cat_pid, subcat_cid, tr_id))
    conn.commit()

def delete_transaction(cursor, conn, tr_id):
    cursor.execute("DELETE FROM Trans WHERE Tr_ID=?", (tr_id,))
    conn.commit()

def fetch_month_rows(cursor, month, year, accounts, account_data):
    # Get FF checkbox setting
    cursor.execute("SELECT Lup_seq FROM Lookups WHERE Lup_LupT_ID = 8")
    ff_rows = cursor.fetchone()
    ff_flag = ff_rows[0]

    cursor.execute("SELECT Tr_ID, Tr_Type, Tr_Reg_ID, Tr_Day, Tr_Month, Tr_Year, Tr_Stat, Tr_Query_Flag, Tr_Amount, "
                "Tr_Desc, Tr_Exp_ID, Tr_ExpSub_ID, Tr_Acc_From, Tr_Acc_To, Tr_FF_Journal_ID "
                "FROM Trans WHERE Tr_Year = ? AND Tr_Month = ? ORDER BY Tr_Day ASC, Tr_Stat DESC", (year, month))
    db_rows = cursor.fetchall()
    rows = []
    daily_balances = [float(acc[2]) if acc[2] is not None else 0.0 for acc in account_data]
    for m in range(month - 1):
        for i, acc in enumerate(account_data):
            daily_balances[i] += float(acc[3 + m]) if acc[3 + m] is not None else 0.0
    current_day = None
    dow_map = {0: "Mon", 1: "Tue", 2: "Wed", 3: "Thu", 4: "Fri", 5: "Sat", 6: "Sun"}

    # Check if month is empty
    if len(db_rows) == 0:
        total_values = ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""] 
        rows.append({"values": tuple(total_values), "status": "Total", "flag": 0, "type": "Total"})
        total_values = ["", "", "", "", "",  "NO Transactions found for this month", "", "", "", "", "", "", "", "", "", "", "", "", "", ""]
        rows.append({"values": tuple(total_values), "status": "Total", "flag": 0, "type": "Total"})
        total_values = ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""] 
        rows.append({"values": tuple(total_values), "status": "Total", "flag": 0, "type": "Total"})
    else:
        for row in db_rows:
            tr_id, tr_type, tr_reg_id, tr_day, tr_month, tr_year, tr_stat, tr_query_flag, tr_amount, tr_desc, tr_exp_id, tr_expsub_id, tr_acc_from, tr_acc_to, tr_ff_id = row
            type_map = {1: "Income", 2: "Expense", 3: "Transfer"}
            trans_type = type_map.get(tr_type, "Expense")
            status_map = {0: "Unknown", 1: "Forecast", 2: "Processing", 3: "Complete"}
            status = status_map.get(tr_stat, "Unknown")
            flag = tr_query_flag if tr_query_flag in [0, 1, 2, 4, 8, 9, 10, 12] else 0

            date_obj = datetime(tr_year, tr_month, tr_day)
            day_name = dow_map[date_obj.weekday()]

            # Daily Totals row
            if current_day is not None and current_day != tr_day:   
                total_values = ["", "", "", "", "", "End of Day Balance"] + ["" if bal == 0.0 else "{:,.2f}".format(bal) for bal in daily_balances]
                rows.append({"values": tuple(total_values), "status": "Total", "flag": 0, "type": "Total"})

            current_day = tr_day

            # If Show FF ID flag set then prefix FF ID to description - for crosschecking transaction with FireFly
            if ff_flag == 1:
                if tr_ff_id is not None and tr_ff_id != 0 and tr_ff_id != "":
                    x = str(tr_ff_id)
                else:
                    x = " "
                tr_ff_desc = f"({x}) {tr_desc}"
            else:
                tr_ff_desc = tr_desc
                
            # If has a Regular Transaction ID exists
            if tr_reg_id > 0:
                str_reg = "**"
            else:
                str_reg = ""

            values = [day_name, str(tr_day), str_reg, "", "", tr_ff_desc] + [""] * 14
            amount = float(tr_amount) if tr_amount else 0.0
            if trans_type == "Income":
                if tr_exp_id == 99:
                    values[3] = "{:,.2f} ??".format(amount)
                else:
                    values[3] = "{:,.2f}".format(amount)
                if tr_acc_to and 0 <= tr_acc_to - 1 < len(accounts):
                    values[6 + (tr_acc_to - 1)] = "{:,.2f}".format(amount)
                    daily_balances[tr_acc_to - 1] += amount
            elif trans_type == "Expense":
                if tr_exp_id == 99:
                    values[4] = "{:,.2f} ??".format(-amount)
                else:
                    values[4] = "{:,.2f}".format(-amount)
                if tr_acc_from and 0 <= tr_acc_from - 1 < len(accounts):
                    values[6 + (tr_acc_from - 1)] = "{:,.2f}".format(-amount)
                    daily_balances[tr_acc_from - 1] -= amount
            elif trans_type == "Transfer":
                if tr_acc_from and tr_acc_to and 0 <= tr_acc_from - 1 < len(accounts) and 0 <= tr_acc_to - 1 < len(accounts):
                    values[6 + (tr_acc_from - 1)] = "{:,.2f}".format(-amount)
                    values[6 + (tr_acc_to - 1)] = "{:,.2f}".format(amount)
                    daily_balances[tr_acc_from - 1] -= amount
                    daily_balances[tr_acc_to - 1] += amount

            rows.append({"values": tuple(values), "status": status, "flag": flag, "type": trans_type, "tr_id": tr_id})

    if current_day is not None:
        total_values = ["", "", "", "", "",  "End of Day Balance"] + ["" if bal == 0.0 else "{:,.2f}".format(bal) for bal in daily_balances]
        rows.append({"values": tuple(total_values), "status": "Total", "flag": 0, "type": "Total"})

    return rows

def fetch_actuals(cursor, pid, cid, month, year):
    cursor.execute("""
        SELECT SUM(Tr_Amount)
        FROM Trans
        WHERE Tr_Year = ? AND Tr_Month = ? AND Tr_Exp_ID = ? AND Tr_ExpSub_ID = ?
    """, (year, month, pid, cid))
    result = cursor.fetchone()
    return result[0] or 0.0


# Lookups
def fetch_notes(cursor):
    cursor.execute("SELECT Lup_Desc FROM Lookups WHERE Lup_LupT_ID = 7 ORDER BY Lup_ID")  # Assuming Lup_ID for order
    rows = cursor.fetchall()
    return "\n".join(row[0] for row in rows) if rows else "No notes available"

def fetch_years(cursor):
    cursor.execute("SELECT Lup_Desc FROM Lookups WHERE Lup_LupT_ID = 1 ORDER BY Lup_Seq")
    return [row[0] for row in cursor.fetchall()]

def fetch_lookup_values(cursor, lup_type_id):
    cursor.execute("SELECT Lup_Desc FROM Lookups WHERE Lup_LupT_ID = ? ORDER BY Lup_Seq", (lup_type_id,))
    return [row[0] for row in cursor.fetchall()]


################### Category (IE_Cata) Table (Remember to use Year in queries)

# Fetch Parent Categories
def fetch_categories(cursor, year, is_income=False):
    if is_income:
        cursor.execute("SELECT IE_PID, IE_Desc FROM IE_Cata WHERE IE_CID = 0 AND IE_Year = ? AND IE_PID = 1 ORDER BY IE_Seq", (year,))
    else:
        cursor.execute("SELECT IE_PID, IE_Desc FROM IE_Cata WHERE IE_CID = 0 AND IE_Year = ? AND IE_PID > 1 ORDER BY IE_Seq", (year,))
    return cursor.fetchall()  # [(IE_PID, IE_Desc), ...]

# Fetch ALL Parent Categories
def fetch_all_categories(cursor, year):
    cursor.execute("SELECT IE_PID, IE_Desc FROM IE_Cata WHERE IE_CID = 0 AND IE_Year = ? ORDER BY IE_Seq", (year,))
    return cursor.fetchall()  # [(IE_PID, IE_Desc), ...]

# Fetch Sub-Categories for a Parent Category
def fetch_subcategories(cursor, pid, year):
    cursor.execute("SELECT IE_CID, IE_Desc FROM IE_Cata WHERE IE_Year = ? AND IE_PID = ? AND IE_CID > 0 ORDER BY IE_Seq", (year, pid))
    return cursor.fetchall()  # [(IE_CID, IE_Desc), ...]

def fetch_exp_categories(cursor, year):
    cursor.execute("SELECT IE_Desc FROM IE_Cata WHERE IE_Year=? AND IE_CID=0 AND IE_PID>1", (year,))
    return [row[0] for row in cursor.fetchall()]

def fetch_inc_categories(cursor, year):
    cursor.execute("SELECT IE_Desc FROM IE_Cata WHERE IE_Year=? AND IE_CID=0 AND IE_PID=1", (year,))
    return [row[0] for row in cursor.fetchall()]

def fetch_category_id(cursor, cat_desc, year):
    cursor.execute("SELECT IE_PID FROM IE_Cata WHERE IE_Desc=? AND IE_Year=? AND IE_CID=0", (cat_desc, year))
    row = cursor.fetchone()
    return row[0] if row else 0

def fetch_subcategory_id(cursor, cat_id, subcat_desc, year):
    cursor.execute("SELECT IE_CID FROM IE_Cata WHERE IE_PID=? AND IE_Desc=? AND IE_Year=?", (cat_id, subcat_desc, year))
    row = cursor.fetchone()
    return row[0] if row else 0

def fetch_category_name(cursor, cat_id, year):
    cursor.execute("SELECT IE_Desc FROM IE_Cata WHERE IE_PID=? AND IE_CID=0 AND IE_Year=?", (cat_id, year))
    row = cursor.fetchone()
    return row[0] if row else ""

def fetch_subcategory_name(cursor, cat_id, subcat_id, year):
    cursor.execute("SELECT IE_Desc FROM IE_Cata WHERE IE_PID=? AND IE_CID=? AND IE_Year=?", (cat_id, subcat_id, year))
    row = cursor.fetchone()
    return row[0] if row else ""


# Account Table
def fetch_account_full_names(cursor, year):
    cursor.execute("SELECT Acc_Name FROM Account WHERE Acc_Year = ? ORDER BY Acc_ID", (year,))
    return [row[0] for row in cursor.fetchall()]

# Fetch Account Names for use by save rule
def fetch_account_names(cursor, year):
    cursor.execute("SELECT Acc_ID, Acc_Name FROM Account WHERE Acc_Year = ? AND Acc_ID BETWEEN 1 AND 12 ORDER BY Acc_ID", (year,))
    return cursor.fetchall()

# Fetch Account Names for use by Triggers/Actions Combo
def fetch_account_c_names(cursor, year):
    cursor.execute("SELECT Acc_Name FROM Account WHERE Acc_Year = ? AND Acc_ID BETWEEN 1 AND 12 ORDER BY Acc_ID", (year,))
    return [row[0] for row in cursor.fetchall()]

def fetch_account_short_names(cursor, year):
    cursor.execute("SELECT Acc_Short_Name FROM Account WHERE Acc_Year = ? ORDER BY Acc_ID", (year))
    result = cursor.fetchone()
    return result[0] if result else ""

def fetch_account_full_name(cursor, acc_id, year):
    cursor.execute("SELECT Acc_Name FROM Account WHERE Acc_ID = ? AND Acc_Year = ?", (acc_id, year))
    result = cursor.fetchone()
    return result[0] if result else ""

def fetch_account_short_name(cursor, acc_id, year):
    cursor.execute("SELECT Acc_Short_Name FROM Account WHERE Acc_ID = ? AND Acc_Year = ?", (acc_id, year))
    result = cursor.fetchone()
    return result[0] if result else ""

def fetch_account_id_by_name(cursor, acc_name, year):
    cursor.execute("SELECT Acc_ID FROM Account WHERE Acc_Name=? AND Acc_Year=?", (acc_name, year))
    row = cursor.fetchone()
    return row[0] if row else ""

def fetch_accounts_by_year(cursor, year):
    cursor.execute("""
        SELECT Acc_ID, Acc_Type, Acc_Name, Acc_Short_Name, Acc_Last4, Acc_Credit_Limit, 
            Acc_Colour, Acc_Open, Acc_Statement_Date, Acc_Prev_Month
        FROM Account WHERE Acc_Year = ? ORDER BY Acc_ID
    """, (year,))
    return cursor.fetchall()

def insert_account(cursor, conn, year, acc_type, name, short_name, last4, credit_limit, colour, open_balance, stmt_date, prev_month):
    cursor.execute("""
        INSERT INTO Account (Acc_Year, Acc_Type, Acc_Name, Acc_Short_Name, Acc_Last4, Acc_Credit_Limit, 
                            Acc_Colour, Acc_Open, Acc_Statement_Date, Acc_Prev_Month)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (year, acc_type, name, short_name, last4, credit_limit, colour, open_balance, stmt_date, prev_month))
    conn.commit()
    return cursor.lastrowid

def update_account(cursor, conn, acc_id, year, acc_type, name, short_name, last4, credit_limit, colour, open_balance, stmt_date, prev_month):
    cursor.execute("""
        UPDATE Account SET Acc_Type = ?, Acc_Name = ?, Acc_Short_Name = ?, Acc_Last4 = ?, 
                        Acc_Credit_Limit = ?, Acc_Colour = ?, Acc_Open = ?, 
                        Acc_Statement_Date = ?, Acc_Prev_Month = ?
        WHERE Acc_ID = ? AND Acc_Year = ?
    """, (acc_type, name, short_name, last4, credit_limit, colour, open_balance, stmt_date, prev_month, acc_id, year))
    conn.commit()

def copy_accounts_from_previous_year(cursor, conn, target_year):
    prev_year = target_year - 1
    cursor.execute("""
        INSERT INTO Account (Acc_ID, Acc_Year, Acc_Type, Acc_Name, Acc_Short_Name, Acc_Last4, Acc_Credit_Limit, 
                            Acc_Colour, Acc_Open, Acc_Statement_Date, Acc_Prev_Month)
        SELECT Acc_ID, ?, Acc_Type, Acc_Name, Acc_Short_Name, Acc_Last4, Acc_Credit_Limit, 
                            Acc_Colour, Acc_Open, Acc_Statement_Date, Acc_Prev_Month
        FROM Account WHERE Acc_Year = ?
    """, (target_year, prev_year))
    conn.commit()
    return cursor.rowcount

def bring_forward_opening_balances(cursor, conn, current_year):
    prev_year = current_year - 1
    # Fetch prior year's data
    cursor.execute("""
        SELECT Acc_ID, Acc_Open, Acc_Jan, Acc_Feb, Acc_Mar, Acc_Apr, Acc_May, Acc_Jun, 
                Acc_Jul, Acc_Aug, Acc_Sep, Acc_Oct, Acc_Nov, Acc_Dec
        FROM Account WHERE Acc_Year = ?
    """, (prev_year,))
    prior_accounts = cursor.fetchall()
    
    if not prior_accounts:
        return 0  # No prior data to process
    
    # Calculate new opening balances and update current year
    for acc in prior_accounts:
        acc_id, acc_open, *months = acc
        acc_open = acc_open or 0.0
        month_sum = sum(m or 0.0 for m in months)  # Handle NULL as 0
        new_open = acc_open + month_sum
        cursor.execute("""
            UPDATE Account 
            SET Acc_Open = ?
            WHERE Acc_Year = ? AND Acc_ID = ?
        """, (new_open, current_year, acc_id))
    conn.commit()
    return len(prior_accounts)  # Number of accounts updated

def update_account_month_transaction_total(cursor, conn, month, year, accounts):

    month_map = {
        1: "Acc_Jan", 2: "Acc_Feb", 3: "Acc_Mar", 4: "Acc_Apr",
        5: "Acc_May", 6: "Acc_Jun", 7: "Acc_Jul", 8: "Acc_Aug",
        9: "Acc_Sep", 10: "Acc_Oct", 11: "Acc_Nov", 12: "Acc_Dec"
    }
    month_col = month_map.get(month, 0)  # Default to 0 if invalid    

    # Fetch Acc_ID and Opening balances
    cursor.execute("SELECT Acc_ID, Acc_Open FROM Account WHERE Acc_Year = ? ORDER BY Acc_ID", (year,))
    acc_map = {row[0]: row[1] for row in cursor.fetchall()}

    # Initialize account totals (assuming Acc_ID starts at 1 and goes to 14)
    mtotal = [0.0] * len(accounts)  # Index 0-13 maps to Acc_ID 1-14

    # Fetch transactions for the provided month & year
    cursor.execute("""
        SELECT Tr_Acc_From, Tr_Acc_To, Tr_Amount
        FROM Trans
        WHERE Tr_Month = ? AND Tr_Year = ?
    """, (month, year))
    
    # Calc change for each account
    for from_acc, to_acc, amount in cursor.fetchall():
        amount = float(amount or 0.0)

    # Adjust for None and valid Acc_IDs (1-14)
        if from_acc is not None and from_acc != 0:
            mtotal[from_acc - 1] -= amount  # Debit from_acc
        if to_acc is not None and to_acc != 0:
            mtotal[to_acc - 1] += amount    # Credit to_acc

    # Update Account table with monthly totals
    for acc_id in acc_map.keys():  # Only update accounts in acc_map
        idx = acc_id - 1  # Adjust for 0-based indexing
        cursor.execute(f""" UPDATE Account SET {month_col} = ? WHERE Acc_ID = ? AND Acc_Year = ? """, (mtotal[idx], acc_id, year))
        conn.commit()

def update_account_year_transactions(cursor, conn, year, accounts):
    for month in range(12):
        update_account_month_transaction_total(cursor, conn, month+1, year, accounts)


# Window Positions
def get_window_position(cursor, win_id):
    cursor.execute("SELECT Win_Name, Win_Left, Win_Top FROM Windows WHERE Win_ID = ?", (win_id,))
    result = cursor.fetchone()
    return result if result else (None, None, None)  # (name, left, top) or (None, None, None)

def save_window_position(cursor, conn, win_id, win_name, left, top):
    cursor.execute(""" INSERT OR REPLACE INTO Windows (Win_ID, Win_Name, Win_Left, Win_Top) VALUES (?, ?, ?, ?) """,
                    (win_id, win_name, left, top))
    conn.commit()


# AHB support
def fetch_transaction_sums(cursor, month, year, accounts):
    # Map account short names to Acc_ID
    cursor.execute("SELECT Acc_ID, Acc_Short_Name FROM Account WHERE Acc_Year = ? ORDER BY Acc_ID", (year,))
    acc_map = {row[1]: row[0] for row in cursor.fetchall()}
    
    # Initialize sums
    completed = [0.0] * len(accounts)
    processing = [0.0] * len(accounts)
    forecast = [0.0] * len(accounts)

    # Initialise Transaction Counters
    tc_complete = 0
    tc_processing = 0
    tc_forecast = 0
    tc_total = 0
    
    # Fetch transactions for the provided month & year
    cursor.execute("""
        SELECT Tr_Acc_From, Tr_Acc_To, Tr_Stat, Tr_Amount
        FROM Trans
        WHERE Tr_Month = ? AND Tr_Year = ?
    """, (month, year))
    
    for from_acc, to_acc, status, amount in cursor.fetchall():
        amount = float(amount or 0.0)
        tc_total += 1

        # Map Acc_ID to index, skip if not found
        from_idx = None
        if from_acc != 0:
            from_name = next((name for name, id in acc_map.items() if id == from_acc), None)
            if from_name and from_name in accounts:
                from_idx = accounts.index(from_name)
            else:
                logger.warning(f"Warning: Tr_Acc_From={from_acc} not found in accounts for year {year}")
        
        to_idx = None
        if to_acc != 0:
            to_name = next((name for name, id in acc_map.items() if id == to_acc), None)
            if to_name and to_name in accounts:
                to_idx = accounts.index(to_name)
            else:
                logger.warning(f"Warning: Tr_Acc_To={to_acc} not found in accounts for year {year}")
        
        if status == 3:  # Completed
            tc_complete += 1
            if from_idx is not None and to_idx is not None:  # Transfer
                completed[from_idx] -= amount
                completed[to_idx] += amount
            elif from_idx is not None:  # Expenditure
                completed[from_idx] -= amount
            elif to_idx is not None:  # Income
                completed[to_idx] += amount
        elif status == 2:  # Processing
            tc_processing += 1
            if from_idx is not None and to_idx is not None:
                processing[from_idx] -= amount
                processing[to_idx] += amount
            elif from_idx is not None:
                processing[from_idx] -= amount
            elif to_idx is not None:
                processing[to_idx] += amount
        elif status == 1:  # Forecast
            tc_forecast += 1
            if from_idx is not None and to_idx is not None:
                forecast[from_idx] -= amount
                forecast[to_idx] += amount
            elif from_idx is not None:
                forecast[from_idx] -= amount
            elif to_idx is not None:
                forecast[to_idx] += amount

    return completed, processing, forecast, tc_complete, tc_processing, tc_forecast, tc_total

def fetch_statement_balances(cursor, month, year, accounts):
    cursor.execute("SELECT Acc_ID, Acc_Short_Name, Acc_Open, Acc_Statement_Date, Acc_Prev_Month, "
                "Acc_Jan, Acc_Feb, Acc_Mar, Acc_Apr, Acc_May, Acc_Jun, "
                "Acc_Jul, Acc_Aug, Acc_Sep, Acc_Oct, Acc_Nov, Acc_Dec "
                "FROM Account WHERE Acc_Year = ? ORDER BY Acc_ID", (year,))
    account_data = cursor.fetchall()
    acc_map = {row[1]: row[0] for row in account_data}
    
    target_year = year if month > 1 else year - 1
    target_month = month - 1 if month > 1 else 12
    days_in_prev_month = monthrange(target_year, target_month)[1]
    target_year = year
    target_month = month
    days_in_month = monthrange(target_year, target_month)[1]
    
    balances = [None] * len(accounts)
    
    for idx, (acc_id, acc_name, acc_open, stmt_date, prev_month, *monthly) in enumerate(account_data):
        if not stmt_date or stmt_date == 0:
            continue
        
        calc_month = month if prev_month == 0 else (month - 1 if month > 1 else 12)
        calc_year = year if prev_month == 0 else (year - 1 if month == 1 else year)
        max_days = days_in_month if prev_month == 0 else days_in_prev_month
        stmt_day = min(stmt_date, max_days)
        
        som = float(acc_open or 0.0)
        for m in range(calc_month - 1):
            som += float(monthly[m] or 0.0)
        
        cursor.execute("""
            SELECT Tr_Acc_From, Tr_Acc_To, Tr_Stat, Tr_Amount
            FROM Trans
            WHERE Tr_Month = ? AND Tr_Year = ? AND Tr_Day <= ?
        """, (calc_month, calc_year, stmt_day))
        
        balance = som
        for from_acc, to_acc, status, amount in cursor.fetchall():
            amount = float(amount or 0.0)
            from_name = next((name for name, id in acc_map.items() if id == from_acc), None) if from_acc != 0 else None
            to_name = next((name for name, id in acc_map.items() if id == to_acc), None) if to_acc != 0 else None
            from_idx = accounts.index(from_name) if from_name else None
            to_idx = accounts.index(to_name) if to_name else None
            
            if from_idx == idx and to_idx is not None:  # Transfer out
                balance -= amount
            elif to_idx == idx and from_idx is not None:  # Transfer in
                balance += amount
            elif from_idx == idx:  # Expenditure
                balance -= amount
            elif to_idx == idx:  # Income
                balance += amount
        
        balances[idx] = balance
    
    return balances


# Regular Transactions Table
def fetch_regular_for_year(cursor, year):
    cursor.execute("""
        SELECT Reg_ID, Reg_Year, Reg_Frequency, Reg_Day, Reg_Month, Reg_Type, Reg_Amount, Reg_Desc, 
            Reg_Start, Reg_Stop, Reg_Exp_ID, Reg_ExpSub_ID, Reg_Acc_From, Reg_Acc_To, Reg_Query_Flag
        FROM Regular WHERE Reg_Year = ? ORDER BY Reg_Frequency, Reg_Day
    """, (year,))
    rows = cursor.fetchall()
    return [{"Reg_ID": r[0], "Reg_Year": r[1], "Reg_Frequency": r[2], "Reg_Day": r[3], "Reg_Month": r[4],
            "Reg_Type": r[5], "Reg_Amount": r[6], "Reg_Desc": r[7], "Reg_Start": r[8], "Reg_Stop": r[9],
            "Reg_Exp_ID": r[10], "Reg_ExpSub_ID": r[11], "Reg_Acc_From": r[12], "Reg_Acc_To": r[13],
            "Reg_Query_Flag": r[14]} for r in rows]

def fetch_regular_by_id(cursor, reg_id):
    cursor.execute("SELECT * FROM Regular WHERE Reg_ID=?", (reg_id,))
    row = cursor.fetchone()
    if row:
        columns = [desc[0] for desc in cursor.description]
        return dict(zip(columns, row))
    return None


# Budget Table
def fetch_budget(cursor, pid, cid, year):
    cursor.execute("""
        SELECT Bud_M1, Bud_M2, Bud_M3, Bud_M4, Bud_M5, Bud_M6, 
            Bud_M7, Bud_M8, Bud_M9, Bud_M10, Bud_M11, Bud_M12
        FROM Budget WHERE Bud_Year = ? AND Bud_PID = ? AND Bud_CID = ?
    """, (year, pid, cid))
    row = cursor.fetchone()
    if row:
        return {i+1: row[i] or 0.0 for i in range(12)}
    return {i+1: 0.0 for i in range(12)}


###############  RULES ENGINE RELATED TABELS  ############################

# Rule Group Table 
def create_rule_group(cursor, conn, name):
    cursor.execute("INSERT INTO RuleGroups (Group_Name, Group_Sequence, Group_Enabled) VALUES (?, ?, 1)",
                    (name, cursor.execute("SELECT COALESCE(MAX(Group_Sequence), 0) + 1 FROM RuleGroups").fetchone()[0]))
    conn.commit()

def fetch_rule_group_names(cursor):
    cursor.execute("SELECT Group_Name FROM RuleGroups WHERE Group_Enabled = 1 ORDER BY Group_Sequence")
    return [row[0] for row in cursor.fetchall()]


# Triggers Table
def delete_trigger(cursor, conn, trigger_id):
    logger.debug(f"Deleting trigger ID: {trigger_id}")
    cursor.execute("DELETE FROM Triggers WHERE Trigger_ID = ?", (trigger_id,))
    conn.commit()

# Actions Table
def delete_action(cursor, conn, action_id):
    logger.debug(f"Deleting action ID: {action_id}")
    cursor.execute("DELETE FROM Actions WHERE Action_ID = ?", (action_id,))
    conn.commit()

# Trig_Options Table
def fetch_trigger_options(cursor):                  # fetches ALL records
    cursor.execute("SELECT TrigO_ID, TrigO_Description FROM Trig_Options ORDER BY TrigO_Seq")
    return cursor.fetchall()

def fetch_trigger_option(cursor, trigo_id):         # fetches one record
    cursor.execute("SELECT TrigO_Description FROM Trig_Options WHERE TrigO_ID = ?", (trigo_id,))
    result = cursor.fetchone()
    return result if result else ()

# Act_Options Table
def fetch_action_options(cursor):                   # fetches ALL records
    cursor.execute("SELECT ActO_ID, ActO_Description FROM Act_Options ORDER BY ActO_Seq")
    result = cursor.fetchall()
    return result if result else ()

def fetch_action_option(cursor, acto_id):           # fetches one record
    cursor.execute("SELECT ActO_Description FROM Act_Options WHERE ActO_ID = ?", (acto_id,))
    result = cursor.fetchone()
    return result if result else ()







