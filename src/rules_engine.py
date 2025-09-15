# src/rules_engine.py
# Transaction processing and rule application

import json
import logging
import re
import os
import sqlite3
import traceback
from datetime import datetime, timedelta
from config import CONFIG, get_config

# Get logger
logger = logging.getLogger('HA.rules_engine')

def process_transactions(json_path, conn, acc_id):
    """
    Process transactions from a JSON file into HA_Import and then Trans table.
    """
    logger.debug(f"Processing transactions from {json_path} for account {acc_id}")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = os.path.join(get_config('LOG_DIR'), f"{timestamp}_acc{acc_id}_import_rules.log")
    logger.info(f"Switching log output to {log_file}")
    
    # Ensure log directory exists
    log_dir = get_config('LOG_DIR')
    os.makedirs(log_dir, exist_ok=True)

    # Save current handlers and create new FileHandler
    original_handlers = logger.handlers[:]
    logger.handlers = []  # Clear existing handlers
    try:
        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        logger.addHandler(file_handler)
        logger.setLevel(logging.DEBUG)
        logger.debug(f"Configured logger with new FileHandler for {log_file}")
    except Exception as e:
        logger.error(f"Failed to configure log file {log_file}: {e}")
        # Restore original handlers and re-raise
        logger.handlers = original_handlers
        raise
    
    try:
        #year = int(root.year_var.get())
        cur = conn.cursor()
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        # Sort transactions by date (oldest first)
        transactions = []
        for account_data in data:
            for trans_type in ["booked", "pending"]:
                for trans in account_data.get("transactions", {}).get(trans_type, []):
                    date = trans.get("bookingDate", "") or trans.get("valueDate", "") or datetime.now().strftime("%Y-%m-%d")
                    transactions.append((date, trans_type, trans, account_data["account_id"]))
        transactions.sort(key=lambda x: x[0])  # Sort by date
        logger.debug(f"Sorted {len(transactions)} transactions by date")
        #logger.debug(f"transactions={transactions}")

        for date, trans_type, trans, account_id in transactions:
            #logger.debug(f"date={date}, trans_type={trans_type}")
            # Create HAI_UID
            transaction_id = trans.get("transactionId", "")
            if transaction_id == "":
                transaction_id = trans.get("internalTransactionId", "")
            hai_uid = f"{acc_id}-{transaction_id}"
            logger.debug(f"++++++++++++ Processing transaction {hai_uid}")
            
            # Map JSON to HA_Import
            amount = float(trans["transactionAmount"]["amount"])
            hai_type = 1 if amount > 0 else 2
            hai_stat = 3 if trans_type == "booked" else 2
            hai_amount = abs(amount)
            date_parts = date.split("-")
            #logger.debug(f"date_parts={date_parts}")
            if len(date_parts) != 3:
                logger.warning(f"Invalid date format for transaction {transaction_id}: {date}")
                continue
            hai_year, hai_month, hai_day = map(int, date_parts)
            hai_desc = trans.get("remittanceInformationUnstructured", "")
            hai_acc_from = acc_id if amount < 0 else 0
            hai_acc_to = acc_id if amount > 0 else 0
            
            # Prepare json data for saving to Trans_Rules
            source_data = date + " " + trans_type + " "
            if hai_type == 1:
                source_data += "Deposit"
            else:
                source_data += "Withdrawal"
            source_data += " £" + str(hai_amount) + " '" + hai_desc + "'"
            #logger.debug(f"source_data={source_data}")
            
            # Check for existing pending transaction with same HAI_UID
            cur.execute("""
                SELECT HAI_ID, HAI_Stat FROM HA_Import WHERE HAI_UID = ?
            """, (hai_uid,))
            existing_record = cur.fetchone()
            
            if existing_record and hai_stat == 3 and existing_record[1] == 2:
                # Update existing pending transaction to booked
                hai_id = existing_record[0]
                logger.debug(f"Updating existing pending HA_Import record: HAI_ID={hai_id}, HAI_UID={hai_uid}")
                try:
                    cur.execute("""
                        UPDATE HA_Import SET
                            HAI_Type = ?, HAI_Day = ?, HAI_Month = ?, HAI_Year = ?, HAI_Stat = ?, HAI_Amount = ?,
                            HAI_Desc = ?, HAI_Acc_From = ?, HAI_Acc_To = ?, HAI_Disp = 'pending_updated'
                        WHERE HAI_ID = ?
                    """, (
                        hai_type, hai_day, hai_month, hai_year, hai_stat, hai_amount,
                        hai_desc, hai_acc_from, hai_acc_to, hai_id
                    ))
                    logger.debug(f"Updated HA_Import: HAI_ID={hai_id}, HAI_UID={hai_uid} to booked")
                except sqlite3.Error as e:
                    logger.error(f"Error updating HA_Import record {hai_id}: {e}")
                    continue
            else:
                # Insert new HA_Import record
                try:
                    cur.execute("""
                        INSERT INTO HA_Import (
                            HAI_UID, HAI_Type, HAI_Day, HAI_Month, HAI_Year, HAI_Stat, 
                            HAI_Amount, HAI_Desc, HAI_Acc_From, HAI_Acc_To, 
                            HAI_DOW, HAI_Query_Flag, HAI_Exp_ID, HAI_ExpSub_ID, HAI_Disp
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0, 0, 0, 0, 'pending')
                    """, (
                        hai_uid, hai_type, hai_day, hai_month, hai_year, hai_stat,
                        hai_amount, hai_desc, hai_acc_from, hai_acc_to
                    ))
                    hai_id = cur.lastrowid
                    logger.debug(f"Inserted into HA_Import: HAI_ID={hai_id}, HAI_UID={hai_uid}")
                except sqlite3.IntegrityError:
                    logger.warning(f"Duplicate transaction skipped: {hai_uid}")
                    continue
            
                # Insert record into Trans_Rules but use only hai_id for now - no Trans record exists yet
                try:
                    cur.execute("INSERT INTO Trans_Rules (HAI_ID, Rule_Desc) VALUES (?, ?)", 
                                (hai_id, source_data))
                    logger.debug(f"Inserted into Trans_Rules: HAI_ID={hai_id}, Rule_Desc={source_data}")
                except sqlite3.Error as e:
                    logger.error(f"Error inserting into Trans_Rules: {e}\n{traceback.format_exc()}")
                    raise            
            logger.debug(f"Applying rules: ID={hai_id}, Type={hai_type}, Date={hai_day}/{hai_month}/{hai_year}, Stat={hai_stat}, £={hai_amount}, Desc={hai_desc}, From={hai_acc_from}, To={hai_acc_to}, Bank={acc_id}")
            
            # Apply rules to set HAI_Type, HAI_Desc, etc.
            deleted = apply_rules(conn, hai_id, acc_id)
            logger.debug(f"deleted={deleted}")
            if deleted is True:
                continue  # Transaction deleted by rule
            
            # Fetch updated HA_Import record
            cur.execute("""
                SELECT HAI_Type, HAI_Day, HAI_Month, HAI_Year, HAI_Amount, 
                    HAI_Acc_From, HAI_Acc_To, HAI_Stat, HAI_Desc, HAI_Exp_ID, HAI_ExpSub_ID
                FROM HA_Import WHERE HAI_ID = ?
            """, (hai_id,))
            result = cur.fetchone()
            if not result:
                logger.error(f"HA_Import record {hai_id} not found after apply_rules")
                continue
            hai_type, hai_day, hai_month, hai_year, hai_amount, hai_acc_from, hai_acc_to, hai_stat, hai_desc, hai_exp, hai_expsub = result
            
            # Matching logic
            match_tr_id = match_transaction(cur, hai_type, hai_day, hai_month, hai_year, hai_amount, hai_acc_from, hai_acc_to, hai_stat)
            
            if match_tr_id:
                # Update existing Trans record
                dow = datetime(hai_year, hai_month, hai_day).weekday() + 1
                if hai_exp == 0 or hai_exp == None:
                    hai_exp = 99
                cur.execute("""
                    UPDATE Trans SET Tr_DOW = ?, Tr_Day = ?, Tr_Month = ?, Tr_Year = ?, Tr_Stat = ?, Tr_Amount = ?,
                        Tr_FF_Journal_ID = ?, Tr_Desc = ?, Tr_Acc_From = ?, Tr_Acc_To = ?, Tr_Exp_ID = ?, Tr_ExpSub_ID = ?
                    WHERE Tr_ID = ?
                """, (dow, hai_day, hai_month, hai_year, hai_stat, hai_amount, hai_id, hai_desc, hai_acc_from, hai_acc_to, hai_exp, hai_expsub, match_tr_id))
                cur.execute("UPDATE HA_Import SET HAI_Disp = 'updated', HAI_Tr_ID = ? WHERE HAI_ID = ?", (match_tr_id, hai_id))
                logger.debug(f"Updated Trans record: Tr_ID={match_tr_id}, HAI_ID={hai_id}")
                
                # Update Trans_Rules history and add matched tr_id
                try:
                    cur.execute("SELECT Rule_Desc FROM Trans_Rules WHERE HAI_ID = ?", (hai_id,))
                    curr_desc = cur.fetchone()[0] 
                    logger.debug(f"Fetched Rule_Desc={curr_desc}")
                    rule_desc = curr_desc + f"\nMatched to existing transaction, Tr_ID={match_tr_id}" 
                    logger.debug(f"updated rule_desc={rule_desc}")
                    cur.execute("UPDATE Trans_Rules SET Rule_Desc = ?, Tr_ID = ? WHERE HAI_ID = ?", (rule_desc, match_tr_id, hai_id))
                    logger.debug(f"Updated Trans_Rules: HAI_ID={hai_id}, Tr_ID={match_tr_id}, Rule_Desc={rule_desc}")
                except sqlite3.Error as e:
                    logger.error(f"Error updating Trans_Rules: {e}\n{traceback.format_exc()}")
                    raise            
                
            else:
                # Create new Trans record
                dow = datetime(hai_year, hai_month, hai_day).weekday() + 1
                if hai_exp == 0 or hai_exp == None:
                    hai_exp = 99
                cur.execute("""
                    INSERT INTO Trans (
                        Tr_Type, Tr_Reg_ID, Tr_DOW, Tr_Day, Tr_Month, Tr_Year, Tr_Stat, 
                        Tr_Query_Flag, Tr_Amount, Tr_Desc, Tr_Exp_ID, Tr_ExpSub_ID, 
                        Tr_Acc_From, Tr_Acc_To, Tr_FF_Journal_ID
                    ) VALUES (?, 0, ?, ?, ?, ?, ?, 0, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    hai_type, dow, hai_day, hai_month, hai_year, hai_stat,
                    hai_amount, hai_desc, hai_exp, hai_expsub, hai_acc_from, hai_acc_to, hai_id
                ))
                match_tr_id = cur.lastrowid
                cur.execute("UPDATE HA_Import SET HAI_Disp = 'created', HAI_Tr_ID = ? WHERE HAI_ID = ?", (match_tr_id, hai_id))
                logger.debug(f"No match found for {hai_id}, new transaction created as {match_tr_id}")
            
                # Update Trans_Rules history and add new tr_id
                try:
                    cur.execute("SELECT Rule_Desc FROM Trans_Rules WHERE HAI_ID = ?", (hai_id,))
                    curr_desc = cur.fetchone()[0] 
                    rule_desc = curr_desc + f"\nNo Match- New transaction, Tr_ID={match_tr_id}" 
                    cur.execute("UPDATE Trans_Rules SET Rule_Desc = ?, Tr_ID = ? WHERE HAI_ID = ?", (rule_desc, match_tr_id, hai_id))
                    logger.debug(f"Updated Trans_Rules: HAI_ID={hai_id}, Tr_ID={match_tr_id}, Rule_Desc={rule_desc}")
                except sqlite3.Error as e:
                    logger.error(f"Error updating Trans_Rules: {e}\n{traceback.format_exc()}")
                    raise            
            
    except Exception as e:
        logger.error(f"Error processing transactions: {str(e)}")
        raise
    finally:
        # Restore original handlers
        logger.handlers = original_handlers
        logger.debug(f"Restored original logger handlers: {[h.__class__.__name__ for h in original_handlers]}")    

def match_transaction(cur, hai_type, hai_day, hai_month, hai_year, hai_amount, hai_acc_from, hai_acc_to, hai_stat):
    """
    Match a transaction against Trans table using 31 steps.
    Returns Tr_ID if matched, None if no match.
    """
    logger.debug(f"Matching transaction: Type={hai_type}, Date={hai_year}-{hai_month}-{hai_day}, Amount={hai_amount}")
    
    # Step 1: Exact match
    query = """
        SELECT Tr_ID FROM Trans
        WHERE Tr_Day = ? AND Tr_Month = ? AND Tr_Year = ? 
        AND Tr_Amount = ? AND Tr_Type = ? 
    """
    if hai_stat == 2:   # Pending - just look for Forecast transactions
        query = query + " AND Tr_Stat = 1 "
    else:               # Booked - look for Forecast and Pending
        query = query + " AND Tr_Stat < 3 "
    if hai_type == 1:   # Income
        query = query + " AND Tr_Acc_to = ? "
        hai_acc = hai_acc_to
    elif hai_type == 2: # Expenditure
        query = query + " AND Tr_Acc_From = ? "
        hai_acc = hai_acc_from
    elif hai_type == 3: # Transfer
        query = query + " AND Tr_Acc_From = ? AND Tr_Acc_To = ? "
        hai_acc = 0
    if hai_acc == 0:
        cur.execute(query, (hai_day, hai_month, hai_year, hai_amount, hai_type, hai_acc_from, hai_acc_to))
    else:
        cur.execute(query, (hai_day, hai_month, hai_year, hai_amount, hai_type, hai_acc))
    result = cur.fetchall()
    if len(result) == 1:
        logger.debug(f"Match found at Step 1 - Date={hai_day}-{hai_month}-{hai_year}, £={hai_amount}, Type={hai_type}, From={hai_acc_from}, To={hai_acc_to}, Tr_id={result[0][0]}")
        return result[0][0]
    if len(result) > 1:
        logger.debug(f"Multiple matches found at Step 1 - Date={hai_day}-{hai_month}-{hai_year}, £={hai_amount}, Type={hai_type}, From={hai_acc_from}, To={hai_acc_to}")
        #for row in result:
            #logger.debug(f"  Matching Tr_ID: {row[0]}")
        return None
    logger.debug("Step 1: No match")
    
    # Steps 2–11: Booked transaction matching pending (up to 10 days back)
    if hai_stat == 3:  # Booked
        for i in range(1, 11):
            check_date = datetime(hai_year, hai_month, hai_day) - timedelta(days=i)
            check_day, check_month, check_year = check_date.day, check_date.month, check_date.year
            query = """
                SELECT Tr_ID FROM Trans
                WHERE Tr_Day = ? AND Tr_Month = ? AND Tr_Year = ? 
                AND Tr_Amount = ? AND Tr_Type = ? AND Tr_Stat = 2
            """
            if hai_type == 1:   # Income
                query = query + " AND Tr_Acc_to = ? "
                hai_acc = hai_acc_to
            elif hai_type == 2: # Expenditure
                query = query + " AND Tr_Acc_From = ? "
                hai_acc = hai_acc_from
            elif hai_type == 3: # Transfer
                query = query + " AND Tr_Acc_From = ? AND Tr_Acc_To = ? "
                hai_acc = 0
            if hai_acc == 0:
                cur.execute(query, (check_day, check_month, check_year, hai_amount, hai_type, hai_acc_from, hai_acc_to))
            else:
                cur.execute(query, (check_day, check_month, check_year, hai_amount, hai_type, hai_acc))
            result = cur.fetchall()
            if len(result) == 1:
                logger.debug(f"Match found at Step {i+1} - Date={check_day}-{check_month}-{check_year}, £={hai_amount}, Type={hai_type}, From={hai_acc_from}, To={hai_acc_to}, Tr_id={result[0][0]}")
                return result[0][0]
            if len(result) > 1:
                logger.debug(f"Multiple matches found at Step {i+1} - Date={hai_day}-{hai_month}-{hai_year}, £={hai_amount}, Type={hai_type}, From={hai_acc_from}, To={hai_acc_to}")
                #for row in result:
                    #logger.debug(f"  Matching Tr_ID: {row[0]}")
                return None
            logger.debug(f"Step {i+1}: No match")
    
    # Steps 12–21: Check ±3 days, then -4 to -7 days for forecast/pending
    for i, offset in enumerate([-1, 1, -2, 2, -3, 3, -4, -5, -6, -7], 12):
        check_date = datetime(hai_year, hai_month, hai_day) + timedelta(days=offset)
        check_day, check_month, check_year = check_date.day, check_date.month, check_date.year
        query = """
            SELECT Tr_ID FROM Trans
            WHERE Tr_Day = ? AND Tr_Month = ? AND Tr_Year = ? 
            AND Tr_Amount = ? AND Tr_Type = ? AND Tr_Stat < 3
        """
        if hai_type == 1:   # Income
            query = query + " AND Tr_Acc_to = ? "
            hai_acc = hai_acc_to
        elif hai_type == 2: # Expenditure
            query = query + " AND Tr_Acc_From = ? "
            hai_acc = hai_acc_from
        elif hai_type == 3: # Transfer
            query = query + " AND Tr_Acc_From = ? AND Tr_Acc_To = ? "
            hai_acc = 0
        if hai_acc == 0:
            cur.execute(query, (check_day, check_month, check_year, hai_amount, hai_type, hai_acc_from, hai_acc_to))
        else:
            cur.execute(query, (check_day, check_month, check_year, hai_amount, hai_type, hai_acc))
        result = cur.fetchall()
        if len(result) == 1:
            logger.debug(f"Match found at Step {i} - Date={check_day}-{check_month}-{check_year}, £={hai_amount}, Type={hai_type}, From={hai_acc_from}, To={hai_acc_to}, Tr_id={result[0][0]}")
            return result[0][0]
        if len(result) > 1:
            logger.debug(f"Multiple matches found at Step {i} - Date={check_day}-{check_month}-{check_year}, £={hai_amount}, Type={hai_type}, From={hai_acc_from}, To={hai_acc_to}")
            #for row in result:
                #logger.debug(f"  Matching Tr_ID: {row[0]}")
            return None
        logger.debug(f"Step {i}: No match")
        #logger.debug(f"Step {i}: No match - {query} \n check_day={check_day}, check_month={check_month}, check_year={check_year}, hai_amount={hai_amount}, hai_type={hai_type}, hai_acc_from{hai_acc_from}, hai_acc_to={hai_acc_to}, hai_acc={hai_acc}")
            
    # Steps 22–28: Expense description matching (up to 7 days back)
    if hai_type == 2:  # Expense
        cur.execute("SELECT Pattern FROM Match_Rules WHERE MRule_Type = 'description'")
        patterns = [row[0] for row in cur.fetchall()]
        for i in range(1, 8):
            check_date = datetime(hai_year, hai_month, hai_day) - timedelta(days=i)
            check_day, check_month, check_year = check_date.day, check_date.month, check_date.year
            for pattern in patterns:
                cur.execute("""
                    SELECT Tr_ID FROM Trans
                    WHERE Tr_Day = ? AND Tr_Month = ? AND Tr_Year = ? 
                    AND Tr_Type = 2 AND Tr_Stat = 2
                    AND Tr_Acc_From = ? AND Tr_Desc LIKE ?
                """, (check_day, check_month, check_year, hai_acc_from, f"%{pattern}%"))
                result = cur.fetchall()
                if len(result) == 1:
                    logger.debug(f"Match found at Pattern match Description -{i+21} - Date={check_day}-{check_month}-{check_year}, From={hai_acc_from}, Pattern={pattern}, Tr_id={result[0][0]}")
                    return result[0][0]
                if len(result) > 1:
                    logger.debug(f"Multiple matches found at Pattern match Description -{i+21}, Pattern={pattern}")
                    #logger.debug(f"Multiple matches found at Step {i+21} - Date={hai_day}-{hai_month}-{hai_year}, £={hai_amount}, Type={hai_type}, From={hai_acc_from}, To={hai_acc_to}")
                    #for row in result:
                        #logger.debug(f"  Matching Tr_ID: {row[0]}")
                    return None
                logger.debug(f"Step {i+21}: No match with pattern {pattern}")
    
    # Step 29: Forecast with zero amount
    query = """
        SELECT Tr_ID FROM Trans
        WHERE Tr_Day = ? AND Tr_Month = ? AND Tr_Year = ? 
        AND Tr_Amount = 0 AND Tr_Type = ? AND Tr_Stat = 1
    """
    if hai_type == 1:   # Income
        query = query + " AND Tr_Acc_to = ? "
        hai_acc = hai_acc_to
    elif hai_type == 2: # Expenditure
        query = query + " AND Tr_Acc_From = ? "
        hai_acc = hai_acc_from
    elif hai_type == 3: # Transfer
        query = query + " AND Tr_Acc_From = ? AND Tr_Acc_To = ? "
        hai_acc = 0
    if hai_acc == 0:
        cur.execute(query, (hai_day, hai_month, hai_year, hai_type, hai_acc_from, hai_acc_to))
    else:
        cur.execute(query, (hai_day, hai_month, hai_year, hai_type, hai_acc))
    result = cur.fetchall()
    if len(result) == 1:
        logger.debug(f"Match found at Step 29 - Date={hai_day}-{hai_month}-{hai_year}, Type={hai_type}, From={hai_acc_from}, To={hai_acc_to}, Tr_id={result[0][0]}")
        return result[0][0]
    if len(result) > 1:
        logger.debug(f"Multiple matches found at Step 29 - Forecast with zero amount - Date={hai_day}-{hai_month}-{hai_year}, £={hai_amount}, Type={hai_type}, From={hai_acc_from}, To={hai_acc_to}")
        #for row in result:
            #logger.debug(f"  Matching Tr_ID: {row[0]}")
        return None
    logger.debug("Step 29: No match")
    
    # Step 30: Forecast with amount mismatch
    tolerance = 1 if hai_amount < 10 else hai_amount * 0.1
    query = """
        SELECT Tr_ID FROM Trans
        WHERE Tr_Day = ? AND Tr_Month = ? AND Tr_Year = ? 
        AND Tr_Amount BETWEEN ? AND ? AND Tr_Type = ? AND Tr_Stat = 1
    """
    if hai_type == 1:   # Income
        query = query + " AND Tr_Acc_to = ? "
        hai_acc = hai_acc_to
    elif hai_type == 2: # Expenditure
        query = query + " AND Tr_Acc_From = ? "
        hai_acc = hai_acc_from
    elif hai_type == 3: # Transfer
        query = query + " AND Tr_Acc_From = ? AND Tr_Acc_To = ? "
        hai_acc = 0
    if hai_acc == 0:
        cur.execute(query, (hai_day, hai_month, hai_year, hai_amount - tolerance, hai_amount + tolerance, hai_type, hai_acc_from, hai_acc_to))
    else:
        cur.execute(query, (hai_day, hai_month, hai_year, hai_amount - tolerance, hai_amount + tolerance, hai_type, hai_acc))
    result = cur.fetchall()
    if len(result) == 1:
        logger.debug(f"Match found at Step 30 - Date={hai_day}-{hai_month}-{hai_year}, £ from {hai_amount - tolerance} to {hai_amount + tolerance}, Type={hai_type}, From={hai_acc_from}, To={hai_acc_to}, Tr_id={result[0][0]}")
        return result[0][0]
    if len(result) > 1:
        logger.debug("Step 30: Multiple matches found")
        #logger.debug(f"Multiple matches found at Step 30 - Forecast with mismatch amount - Date={hai_day}-{hai_month}-{hai_year}, £={hai_amount}, Type={hai_type}, From={hai_acc_from}, To={hai_acc_to}")
        #for row in result:
            #logger.debug(f"  Matching Tr_ID: {row[0]}")
        return None
    logger.debug("Step 30: No match")    
    
    # Step 31: No match
    logger.debug("At Step 31 - No Match Found, creating new record")
    return None

def apply_rules(conn, hai_id, acc_id):
    """
    Apply rules to categorize and modify HA_Import transactions.
    Returns deleted=True if transaction deleted.
    """
    logger.debug(f"Applying rules to HA_Import HAI_ID={hai_id}, Acc_ID={acc_id}")
    deleted = False
    try:
        #year = int(root.year_var.get())
        cur = conn.cursor()
        # Fetch HA_Import record
        cur.execute("""
            SELECT HAI_Type, HAI_Amount, HAI_Desc, HAI_Acc_From, HAI_Acc_To, 
                HAI_Day, HAI_Month, HAI_Year, HAI_Stat
            FROM HA_Import WHERE HAI_ID = ?
        """, (hai_id,))
        result = cur.fetchone()
        if not result:
            logger.error(f"HA_Import record {hai_id} not found")
            return None
        hai_type, hai_amount, hai_desc, hai_acc_from, hai_acc_to, hai_day, hai_month, hai_year, hai_stat = result
        
        # Fetch rule groups
        cur.execute("""
            SELECT Group_ID, Group_Sequence
            FROM RuleGroups
            WHERE Group_Enabled = 1
            ORDER BY Group_Sequence
        """)
        groups = cur.fetchall()
        
        for group_id, group_sequence in groups:
            logger.debug(f"  >>>>  Processing Rule Group {group_id} (Group Sequence {group_sequence})")
            # Fetch rules in group
            cur.execute("""
                SELECT Rule_ID, Rule_Sequence, Rule_Enabled, Rule_Active, 
                    Rule_Trigger_Mode, Rule_Proceed
                FROM Rules
                WHERE Group_ID = ? AND Rule_Active = 1 AND Rule_Enabled = 1
                ORDER BY Rule_Sequence
            """, (group_id,))
            rules = cur.fetchall()
            
            group_proceed = True
            for rule_id, rule_sequence, rule_enabled, rule_active, trigger_mode, rule_proceed in rules:
                if not group_proceed:
                    break
                logger.debug(f"  Processing Rule {rule_id} (Rule Sequence {rule_sequence})")
                
                # Fetch triggers
                cur.execute("""
                    SELECT Trigger_ID, TrigO_ID, Value, Trigger_Sequence
                    FROM Triggers
                    WHERE Rule_ID = ?
                    ORDER BY Trigger_Sequence
                """, (rule_id,))
                triggers = cur.fetchall()
                
                if not triggers:
                    all_triggers_matched = False
                    logger.debug(f"Skipping Rule ID {rule_id}: No triggers defined")
                    continue
                
                all_triggers_matched = True
                any_trigger_matched = False
                for trigger_id, trigo_id, value, trigger_sequence in triggers:
                    #logger.debug(f"Evaluating Trigger {trigger_id} (TrigO_ID={trigo_id}, Value={value})")
                    trigger_matched = False
                    
                    # Amount triggers
                    if trigo_id == 1:  # Amount is exactly
                        trigger_matched = float(hai_amount) == float(value)
                    elif trigo_id == 2:  # Amount is less than
                        trigger_matched = float(hai_amount) < float(value)
                    elif trigo_id == 3:  # Amount is more than
                        trigger_matched = float(hai_amount) > float(value)
                        
                    # Tag triggers
                    elif trigo_id in (4, 5, 6, 7, 8, 9, 10, 11):
                        cur.execute("SELECT Tag_Text FROM Rule_Tags WHERE HAI_ID = ?", (hai_id,))
                        tags = [row[0] for row in cur.fetchall()]
                        if trigo_id == 4:  # Any Tag is exactly
                            trigger_matched = any(tag == value for tag in tags)
                        elif trigo_id == 5:  # Any Tag starts with
                            trigger_matched = any(tag.startswith(value) for tag in tags)
                        elif trigo_id == 6:  # Any Tag ends with
                            trigger_matched = any(tag.endswith(value) for tag in tags)
                        elif trigo_id == 7:  # Any Tag contains
                            trigger_matched = any(value in tag for tag in tags)
                        elif trigo_id == 8:  # No Tag is exactly
                            trigger_matched = not any(tag == value for tag in tags)
                        elif trigo_id == 9:  # No Tag starts with
                            trigger_matched = not any(tag.startswith(value) for tag in tags)
                        elif trigo_id == 10:  # No Tag ends with
                            trigger_matched = not any(tag.endswith(value) for tag in tags)
                        elif trigo_id == 11:  # No Tag contains
                            trigger_matched = not any(value in tag for tag in tags)
                    
                    # Description triggers - ARE THESE CASE SENSITIVE???????
                    elif trigo_id in (12, 13, 14, 15):
                        trigger_matched = hai_desc == value if trigo_id == 12 else \
                                        hai_desc.startswith(value) if trigo_id == 13 else \
                                        hai_desc.endswith(value) if trigo_id == 14 else \
                                        value in hai_desc
                    # Account triggers
                    elif trigo_id in (16, 17, 18, 19):
                        trigger_matched = int(hai_acc_to) == int(value) if trigo_id == 16 else \
                                        int(hai_acc_to) != int(value) if trigo_id == 17 else \
                                        int(hai_acc_from) == int(value) if trigo_id == 18 else \
                                        int(hai_acc_from) != int(value)
                    # Date triggers
                    elif trigo_id in (20, 21, 22):  # Booked Date is on/before/after
                        try:
                            trans_date = datetime(hai_year, hai_month, hai_day)
                            value_date = datetime.strptime(value, "%Y-%m-%d")
                            trigger_matched = trans_date.date() == value_date.date() if trigo_id == 20 else \
                                            trans_date.date() < value_date.date() if trigo_id == 21 else \
                                            trans_date.date() > value_date.date()
                        except ValueError as e:
                            logger.error(f"Invalid date format in Trigger {trigger_id}: {value} ({e})")
                            trigger_matched = False
                        # Transaction type triggers
                    elif trigo_id in (23, 24, 25):
                        # 23=Withdrawal, 24=Deposit. 25=Transfer
                        trigger_matched = hai_type == 2 if trigo_id == 23 else \
                                        hai_type == 1 if trigo_id == 24 else \
                                        hai_type == 3
                        # Status triggers
                    elif trigo_id in (26, 27):
                        # 26=Pending, 27=Booked
                        trigger_matched = hai_stat == 2 if trigo_id == 26 else hai_stat == 3
                    
                    logger.debug(f"    Evaluated Trigger {trigger_id} (TrigO_ID={trigo_id}, Value={value}) = {'>>> MATCHED <<<' if trigger_matched else 'did not match'}")
                    
                    if trigger_mode == 'ALL' and not trigger_matched:
                        all_triggers_matched = False
                        break
                    elif trigger_mode == 'Any' and trigger_matched:
                        any_trigger_matched = True
                        
                #logger.debug(f">> Trigger={trigger_id} Trigger Mode={trigger_mode} Any trigger matched={any_trigger_matched} All triggers matched={all_triggers_matched}")
                        
                # Process actions if triggers fire
                if (trigger_mode == 'Any' and any_trigger_matched) or (trigger_mode == 'ALL' and all_triggers_matched):
                    logger.debug(f"Rule {rule_id} triggered")
                    cur.execute("""
                        SELECT Action_ID, ActO_ID, Value, Action_Sequence
                        FROM Actions
                        WHERE Rule_ID = ?
                        ORDER BY Action_Sequence
                    """, (rule_id,))
                    actions = cur.fetchall()
                    
                    # Generate Rule_Desc
                    action_descs = []
                    deleted = False
                    for action_id, acto_id, value, action_sequence in actions:
                        if acto_id == 1:
                            action_descs.append(f"Added tag '{value}'")
                        elif acto_id == 2:
                            action_descs.append(f"Removed tag '{value}'")
                        elif acto_id == 3:
                            action_descs.append("Removed all tags")
                        elif acto_id == 4:
                            action_descs.append("Converted to Deposit")
                        elif acto_id == 5:
                            action_descs.append("Converted to Transfer")
                        elif acto_id == 6:
                            action_descs.append("Converted to Withdrawal")
                        elif acto_id == 7:
                            action_descs.append("Deleted transaction")
                        elif acto_id == 8:
                            action_descs.append(f"Set category to {value}")
                        elif acto_id == 10:
                            action_descs.append(f"Set description to '{value}'")
                        elif acto_id == 11:
                            action_descs.append(f"Appended description with '{value}'")
                        elif acto_id == 12:
                            action_descs.append(f"Prepended description with '{value}'")
                        elif acto_id == 13:
                            action_descs.append(f"Set destination account to {value}")
                        elif acto_id == 14:
                            action_descs.append(f"Set source account to {value}")
                            
                    no_actions = False
                    if len(action_descs) == 0:
                        no_actions = True
                    rule_desc = f"\n  Applied actions:\n    {'\n    '.join(action_descs)}"
                    
                    for action_id, acto_id, value, action_sequence in actions:
                        logger.debug(f"Executing Action {action_id} (ActO_ID={acto_id}, Value={value})")
                        
                        if acto_id == 1:  # Add Tag
                            cur.execute("SELECT Tag_ID FROM Rule_Tags WHERE HAI_ID = ? AND Tag_Text = ?", (hai_id, value))
                            if not cur.fetchone():
                                cur.execute("INSERT INTO Rule_Tags (HAI_ID, Tag_Text, Created_At) VALUES (?, ?, ?)", 
                                            (hai_id, value, int(datetime.now().timestamp())))
                                logger.debug(f"Added tag '{value}' for HAI_ID {hai_id}")
                        elif acto_id == 2:  # Remove Tag
                            cur.execute("DELETE FROM Rule_Tags WHERE HAI_ID = ? AND Tag_Text = ?", (hai_id, value))
                            logger.debug(f"Removed tag '{value}' for HAI_ID {hai_id}")
                        elif acto_id == 3:  # Remove all Tags
                            cur.execute("DELETE FROM Rule_Tags WHERE HAI_ID = ?", (hai_id,))
                            logger.debug(f"Removed all tags for HAI_ID {hai_id}")
                        elif acto_id == 4:  # Convert to Deposit
                            cur.execute("UPDATE HA_Import SET HAI_Type = 1 WHERE HAI_ID = ?", (hai_id,))
                            hai_type = 1
                            logger.debug(f"Converted to Deposit for HAI_ID={hai_id}")
                        elif acto_id == 5:  # Convert to Transfer
                            cur.execute("UPDATE HA_Import SET HAI_Type = 3 WHERE HAI_ID = ?", (hai_id,))
                            hai_type = 3
                            logger.debug(f"Converted to Transfer for HAI_ID={hai_id}")
                        elif acto_id == 6:  # Convert to Withdrawal
                            cur.execute("UPDATE HA_Import SET HAI_Type = 2 WHERE HAI_ID = ?", (hai_id,))
                            hai_type = 2
                            logger.debug(f"Converted to Withdrawal for HAI_ID={hai_id}")
                        elif acto_id == 7:  # Delete Transaction
                            cur.execute("DELETE FROM HA_Import WHERE HAI_ID = ?", (hai_id,))
                            logger.debug(f"Deleted transaction HAI_ID={hai_id}")
                            conn.commit()
                            deleted = True
                            return deleted # Skip further processing
                        elif acto_id == 8:  # Set Category to
                            try:
                                exp_id, expsub_id = map(int, value.split(','))
                                cur.execute("UPDATE HA_Import SET HAI_Exp_ID = ?, HAI_ExpSub_ID = ? WHERE HAI_ID = ?", 
                                            (exp_id, expsub_id, hai_id))
                                logger.debug(f"Set category {exp_id},{expsub_id} for HAI_ID={hai_id}")
                            except ValueError:
                                logger.error(f"Invalid category format in Action {action_id}: {value}")
                        elif acto_id == 10:  # Set Description to
                            cur.execute("UPDATE HA_Import SET HAI_Desc = ? WHERE HAI_ID = ?", (value, hai_id))
                            hai_desc = value
                            logger.debug(f"Set description to '{value}' for HAI_ID={hai_id}")
                        elif acto_id == 11:  # Append Description with
                            cur.execute("UPDATE HA_Import SET HAI_Desc = HAI_Desc || ? WHERE HAI_ID = ?", (value, hai_id))
                            hai_desc += value
                            logger.debug(f"Appended '{value}' to description for HAI_ID={hai_id}")
                        elif acto_id == 12:  # Prepend Description with
                            cur.execute("UPDATE HA_Import SET HAI_Desc = ? || HAI_Desc WHERE HAI_ID = ?", (value, hai_id))
                            hai_desc = value + hai_desc
                            logger.debug(f"Prepended '{value}' to description for HAI_ID={hai_id}")
                        elif acto_id == 13:  # Set Destination Account to
                            try:
                                acc_to = int(value)
                                cur.execute("UPDATE HA_Import SET HAI_Acc_To = ? WHERE HAI_ID = ?", (acc_to, hai_id))
                                hai_acc_to = acc_to
                                logger.debug(f"Set destination account to {acc_to} for HAI_ID={hai_id}")
                            except ValueError:
                                logger.error(f"Invalid account ID in Action {action_id}: {value}")
                        elif acto_id == 14:  # Set Source Account to
                            try:
                                acc_from = int(value)
                                cur.execute("UPDATE HA_Import SET HAI_Acc_From = ? WHERE HAI_ID = ?", (acc_from, hai_id))
                                hai_acc_from = acc_from
                                logger.debug(f"Set source account to {acc_from} for HAI_ID={hai_id}")
                            except ValueError:
                                logger.error(f"Invalid account ID in Action {action_id}: {value}")
                    
                    if no_actions == False:
                        # Prepare to Update Trans_Rules history
                        cur.execute("SELECT Rule_Name FROM Rules WHERE Rule_ID = ?", (rule_id,))
                        rule_name = cur.fetchone()[0] 
                        rule_desc = "Rule Name >> " + rule_name + ": " + rule_desc
                        logger.debug(f"rule_desc={rule_desc}")
                    
                        # UPDATE as Trans_Rules record should already exist. Append new data to end of existing to keep history
                        try:
                            cur.execute("SELECT Rule_Desc FROM Trans_Rules WHERE HAI_ID = ?", (hai_id,))
                            curr_desc = cur.fetchone()[0] 
                            logger.debug(f"Fetched curr_desc = {curr_desc}")
                            rule_desc = curr_desc + "\n" + rule_desc
                            logger.debug(f"Fetched Rule_Desc, added actions = {rule_desc}")
                            cur.execute("UPDATE Trans_Rules SET Rule_Desc = ? WHERE HAI_ID = ?", (rule_desc, hai_id))
                            logger.debug(f"Updated Trans_Rules: HAI_ID={hai_id}, Rule_Desc={rule_desc}")
                        except sqlite3.Error as e:
                            logger.error(f"Error updating Trans_Rules: {e}\n{traceback.format_exc()}")
                            raise            

                    group_proceed = rule_proceed
                
                conn.commit()
        
        logger.debug(f"Completed rule application for HAI_ID={hai_id}")
        return deleted
    
    except sqlite3.Error as e:
        logger.error(f"Database error in apply_rules: {e}")
        raise
    except Exception as e:
        logger.error(f"Error in apply_rules: {e}")
        raise

def cleanup_ha_import(conn):
    """
    Delete HA_Import and Rule_Tags records older than specified days.
    """
    days = CONFIG['IMPORT_DAYS_TO_KEEP']
    logger.debug(f"Cleaning up HA_Import and Rule_Tags records older than {days} days")
    try:
        cur = conn.cursor()
        cutoff_date = datetime.now() - timedelta(days=days)
        cutoff_timestamp = int(cutoff_date.timestamp())
        
        cur.execute("DELETE FROM Rule_Tags WHERE Created_At < ?", (cutoff_timestamp,))
        deleted_tags = cur.rowcount
        cur.execute("""
            DELETE FROM HA_Import
            WHERE HAI_Year * 10000 + HAI_Month * 100 + HAI_Day < ?
        """, (cutoff_date.year * 10000 + cutoff_date.month * 100 + cutoff_date.day,))
        deleted_imports = cur.rowcount
        conn.commit()
        logger.info(f"Deleted {deleted_tags} Rule_Tags and {deleted_imports} HA_Import records older than {days} days")
    except sqlite3.Error as e:
        logger.error(f"Error cleaning up HA_Import/Rule_Tags: {e}")
        raise

def test_rules(json_path, conn, acc_id):
    """
    Test rules on a JSON file and return results for GUI display.
    """
    logger.debug(f"Testing rules on {json_path} for account {acc_id}")
    
    results = []
    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    for account_data in data:
        for trans_type in ["booked", "pending"]:
            for trans in account_data.get("transactions", {}).get(trans_type, []):
                amount = float(trans["transactionAmount"]["amount"])
                date = trans.get("bookingDate", "") or trans.get("valueDate", "") or datetime.now().strftime("%Y-%m-%d")
                description = trans.get("remittanceInformationUnstructured", "")
                category = "Uncategorized"
                cur = conn.cursor()
                cur.execute("SELECT Pattern, Category FROM Match_Rules WHERE MRule_Type = 'description'")
                for pattern, cat in cur.fetchall():
                    if re.search(pattern, description, re.IGNORECASE):
                        category = cat
                        logger.debug(f"Matched pattern {pattern} to category {cat} for transaction {trans.get('transactionId', '')}")
                        break
                results.append({
                    "transaction_id": trans.get("transactionId", ""),
                    "status": trans_type,
                    "booking_date": date,
                    "amount": amount,
                    "currency": trans["transactionAmount"]["currency"],
                    "description": description,
                    "category": category
                })
    logger.debug(f"Tested {len(results)} transactions")                
    return results


