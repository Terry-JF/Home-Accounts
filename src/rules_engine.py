# src/rules_engine.py
# Transaction processing and rule application

import json
import logging
import re
import sqlite3
from datetime import datetime, timedelta
from config import get_config
from db import open_db

# Get logger
logger = logging.getLogger('HA.rules_engine')

def process_transactions(json_path, conn, acc_id):
    """
    Process transactions from a JSON file into HA_Import and then Trans table.
    """
    logger.debug(f"Processing transactions from {json_path} for account {acc_id}")
    
    try:
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

        for date, trans_type, trans, account_id in transactions:
            # Create HAI_UID
            transaction_id = trans.get("transactionId", "")
            hai_uid = f"{acc_id}-{transaction_id}"
            logger.debug(f"Processing transaction {hai_uid}")
            
            # Map JSON to HA_Import
            amount = float(trans["transactionAmount"]["amount"])
            hai_type = 1 if amount > 0 else 2
            hai_stat = 3 if trans_type == "booked" else 2
            hai_amount = abs(amount)
            date_parts = date.split("-")
            if len(date_parts) != 3:
                logger.warning(f"Invalid date format for transaction {transaction_id}: {date}")
                continue
            hai_year, hai_month, hai_day = map(int, date_parts)
            hai_desc = trans.get("remittanceInformationUnstructured", "")
            hai_acc_from = acc_id if amount < 0 else 0
            hai_acc_to = acc_id if amount > 0 else 0
            
            # Insert into HA_Import
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
            
            # Apply rules to set HAI_Type, HAI_Desc, etc.
            apply_rules(conn, hai_id, acc_id)
            
            # Fetch updated HA_Import record
            cur.execute("""
                SELECT HAI_Type, HAI_Day, HAI_Month, HAI_Year, HAI_Amount, 
                    HAI_Acc_From, HAI_Acc_To, HAI_Stat, HAI_Desc
                FROM HA_Import WHERE HAI_ID = ?
            """, (hai_id,))
            result = cur.fetchone()
            if not result:
                logger.error(f"HA_Import record {hai_id} not found after apply_rules")
                continue
            hai_type, hai_day, hai_month, hai_year, hai_amount, hai_acc_from, hai_acc_to, hai_stat, hai_desc = result
            
            # Matching logic
            tr_id = match_transaction(cur, hai_type, hai_day, hai_month, hai_year, hai_amount, hai_acc_from, hai_acc_to, hai_stat)
            
            if tr_id:
                # Update existing Trans record
                dow = datetime(hai_year, hai_month, hai_day).weekday() + 1
                cur.execute("""
                    UPDATE Trans
                    SET Tr_DOW = ?, Tr_Day = ?, Tr_Month = ?, Tr_Year = ?, Tr_Stat = ?, 
                        Tr_Amount = ?, Tr_FF_Journal_ID = ?, Tr_Desc = ?
                    WHERE Tr_ID = ?
                """, (dow, hai_day, hai_month, hai_year, hai_stat, hai_amount, hai_id, hai_desc, tr_id))
                cur.execute("UPDATE HA_Import SET HAI_Disp = 'updated', HAI_Tr_ID = ? WHERE HAI_ID = ?", (tr_id, hai_id))
                logger.debug(f"Updated Trans record: Tr_ID={tr_id}, HAI_ID={hai_id}")
            else:
                # Create new Trans record
                dow = datetime(hai_year, hai_month, hai_day).weekday() + 1
                cur.execute("""
                    INSERT INTO Trans (
                        Tr_Type, Tr_Reg_ID, Tr_DOW, Tr_Day, Tr_Month, Tr_Year, Tr_Stat, 
                        Tr_Query_Flag, Tr_Amount, Tr_Desc, Tr_Exp_ID, Tr_ExpSub_ID, 
                        Tr_Acc_From, Tr_Acc_To, Tr_FF_Journal_ID
                    ) VALUES (?, 0, ?, ?, ?, ?, ?, 0, ?, ?, 99, 99, ?, ?, ?)
                """, (
                    hai_type, dow, hai_day, hai_month, hai_year, hai_stat,
                    hai_amount, hai_desc, hai_acc_from, hai_acc_to, hai_id
                ))
                tr_id = cur.lastrowid
                cur.execute("UPDATE HA_Import SET HAI_Disp = 'created', HAI_Tr_ID = ? WHERE HAI_ID = ?", (tr_id, hai_id))
                logger.debug(f"No match found for {hai_id}, new transaction created as {tr_id}")
                
            conn.commit()
    
    except Exception as e:
        logger.error(f"Error processing transactions: {str(e)}")
        raise

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
        AND Tr_Amount = ? AND Tr_Type = ? AND Tr_Stat < 3 AND Tr_FF_Journal_ID IS NULL
        AND (Tr_Acc_From = ? OR Tr_Acc_To = ?)
    """
    cur.execute(query, (hai_day, hai_month, hai_year, hai_amount, hai_type, hai_acc_from, hai_acc_to))
    result = cur.fetchall()
    if len(result) == 1:
        logger.debug(f"Match found at Step 1 - Date={hai_day}-{hai_month}-{hai_year}, £={hai_amount}, Type={hai_type}, From={hai_acc_from}, To={hai_acc_to}, Tr_id={result[0][0]}")
        return result[0][0]
    if len(result) > 1:
        logger.debug("Multiple matches found at Step 1")
        return None
    logger.debug("Step 1: No match")
    
    # Steps 2–11: Booked transaction matching pending (up to 10 days back)
    if hai_stat == 3:  # Booked
        for i in range(1, 11):
            check_date = datetime(hai_year, hai_month, hai_day) - timedelta(days=i)
            check_day, check_month, check_year = check_date.day, check_date.month, check_date.year
            cur.execute("""
                SELECT Tr_ID FROM Trans
                WHERE Tr_Day = ? AND Tr_Month = ? AND Tr_Year = ? 
                AND Tr_Amount = ? AND Tr_Type = ? AND Tr_Stat = 2 AND Tr_FF_Journal_ID IS NULL
                AND (Tr_Acc_From = ? OR Tr_Acc_To = ?)
            """, (check_day, check_month, check_year, hai_amount, hai_type, hai_acc_from, hai_acc_to))
            result = cur.fetchall()
            if len(result) == 1:
                logger.debug(f"Match found at Step {i+1} - Date={check_day}-{check_month}-{check_year}, £={hai_amount}, Type={hai_type}, From={hai_acc_from}, To={hai_acc_to}, Tr_id={result[0][0]}")
                return result[0][0]
            if len(result) > 1:
                logger.debug(f"Multiple matches found at Step {i+1}")
                return None
            logger.debug(f"Step {i+1}: No match")
    
    # Steps 12–21: Check ±3 days, then -4 to -7 days for forecast/pending
    for i, offset in enumerate([-1, 1, -2, 2, -3, 3, -4, -5, -6, -7], 12):
        check_date = datetime(hai_year, hai_month, hai_day) + timedelta(days=offset)
        check_day, check_month, check_year = check_date.day, check_date.month, check_date.year
        cur.execute("""
            SELECT Tr_ID FROM Trans
            WHERE Tr_Day = ? AND Tr_Month = ? AND Tr_Year = ? 
            AND Tr_Amount = ? AND Tr_Type = ? AND Tr_Stat < 3 AND Tr_FF_Journal_ID IS NULL
            AND (Tr_Acc_From = ? OR Tr_Acc_To = ?)
        """, (check_day, check_month, check_year, hai_amount, hai_type, hai_acc_from, hai_acc_to))
        result = cur.fetchall()
        if len(result) == 1:
            logger.debug(f"Match found at Step {i} - Date={check_day}-{check_month}-{check_year}, £={hai_amount}, Type={hai_type}, From={hai_acc_from}, To={hai_acc_to}, Tr_id={result[0][0]}")
            return result[0][0]
        if len(result) > 1:
            logger.debug(f"Multiple matches found at Step {i}")
            return None
        logger.debug(f"Step {i}: No match")
            
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
                    AND Tr_Type = 2 AND Tr_Stat = 2 AND Tr_FF_Journal_ID IS NULL
                    AND Tr_Acc_From = ? AND Tr_Desc LIKE ?
                """, (check_day, check_month, check_year, hai_acc_from, f"%{pattern}%"))
                result = cur.fetchall()
                if len(result) == 1:
                    logger.debug(f"Match found at Pattern match Description -{i+21} - Date={check_day}-{check_month}-{check_year}, From={hai_acc_from}, Pattern={pattern}, Tr_id={result[0][0]}")
                    return result[0][0]
                if len(result) > 1:
                    logger.debug(f"Multiple matches found at Pattern match Description -{i+21}, Pattern={pattern}")
                    return None
                logger.debug(f"Step {i+21}: No match with pattern {pattern}")
    
    # Step 29: Forecast with zero amount
    cur.execute("""
        SELECT Tr_ID FROM Trans
        WHERE Tr_Day = ? AND Tr_Month = ? AND Tr_Year = ? 
        AND Tr_Amount = 0 AND Tr_Type = ? AND Tr_Stat = 1 AND Tr_FF_Journal_ID IS NULL
        AND (Tr_Acc_From = ? OR Tr_Acc_To = ?)
    """, (hai_day, hai_month, hai_year, hai_type, hai_acc_from, hai_acc_to))
    result = cur.fetchall()
    if len(result) == 1:
        logger.debug(f"Match found at Step 29 - Date={hai_day}-{hai_month}-{hai_year}, Type={hai_type}, From={hai_acc_from}, To={hai_acc_to}, Tr_id={result[0][0]}")
        return result[0][0]
    if len(result) > 1:
        logger.debug("Multiple matches found at Step 29 - Forecast with zero amount")
        return None
    logger.debug("Step 29: No match")
    
    # Step 30: Forecast with amount mismatch
    tolerance = 1 if hai_amount < 10 else hai_amount * 0.1
    cur.execute("""
        SELECT Tr_ID FROM Trans
        WHERE Tr_Day = ? AND Tr_Month = ? AND Tr_Year = ? 
        AND Tr_Amount BETWEEN ? AND ? AND Tr_Type = ? AND Tr_Stat = 1 AND Tr_FF_Journal_ID IS NULL
        AND (Tr_Acc_From = ? OR Tr_Acc_To = ?)
    """, (hai_day, hai_month, hai_year, hai_amount - tolerance, hai_amount + tolerance, hai_type, hai_acc_from, hai_acc_to))
    result = cur.fetchall()
    if len(result) == 1:
        logger.debug(f"Match found at Step 30 - Date={hai_day}-{hai_month}-{hai_year}, £ from {hai_amount - tolerance} to {hai_amount + tolerance}, Type={hai_type}, From={hai_acc_from}, To={hai_acc_to}, Tr_id={result[0][0]}")
        return result[0][0]
    if len(result) > 1:
        logger.debug("Step 30: Multiple matches found")
        return None
    logger.debug("Step 30: No match")    
    
    # Step 31: No match
    logger.debug("At Step 31 - No Match Found, creating new record")
    return None

def apply_rules(conn, hai_id, acc_id):
    """
    Apply rules to categorize and modify HA_Import transactions.
    """
    logger.debug(f"Applying rules to HA_Import HAI_ID={hai_id}, Acc_ID={acc_id}")
    
    try:
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
            return
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
            logger.debug(f"Processing Rule Group {group_id} (Sequence {group_sequence})")
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
                logger.debug(f"Processing Rule {rule_id} (Sequence {rule_sequence})")
                
                # Fetch triggers
                cur.execute("""
                    SELECT Trigger_ID, TrigO_ID, Value, Trigger_Sequence
                    FROM Triggers
                    WHERE Rule_ID = ?
                    ORDER BY Trigger_Sequence
                """, (rule_id,))
                triggers = cur.fetchall()
                
                all_triggers_matched = True
                any_trigger_matched = False
                for trigger_id, trigo_id, value, trigger_sequence in triggers:
                    logger.debug(f"Evaluating Trigger {trigger_id} (TrigO_ID={trigo_id}, Value={value})")
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
                        cur.execute("SELECT Tag_Text FROM Rule_Tags WHERE Rule_ID = ?", (rule_id,))
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
                    
                    # Description triggers
                    elif trigo_id == 12:  # Description is exactly
                        trigger_matched = hai_desc == value
                    elif trigo_id == 13:  # Description starts with
                        trigger_matched = hai_desc.startswith(value)
                    elif trigo_id == 14:  # Description ends with
                        trigger_matched = hai_desc.endswith(value)
                    elif trigo_id == 15:  # Description contains
                        trigger_matched = value in hai_desc
                    
                    # Account triggers
                    elif trigo_id == 16:  # Destination Account Name is
                        trigger_matched = int(hai_acc_to) == int(value) if hai_acc_to else False
                    elif trigo_id == 17:  # Destination Account Name is not
                        trigger_matched = int(hai_acc_to) != int(value) if hai_acc_to else True
                    elif trigo_id == 18:  # Source Account Name is
                        trigger_matched = int(hai_acc_from) == int(value) if hai_acc_from else False
                    elif trigo_id == 19:  # Source Account Name is not
                        trigger_matched = int(hai_acc_from) != int(value) if hai_acc_from else True
                    
                    # Date triggers
                    elif trigo_id in (20, 21, 22):  # Booked Date is on/before/after
                        try:
                            trans_date = datetime(hai_year, hai_month, hai_day)
                            value_date = datetime.strptime(value, "%Y-%m-%d")
                            if trigo_id == 20:  # On
                                trigger_matched = trans_date.date() == value_date.date()
                            elif trigo_id == 21:  # Before
                                trigger_matched = trans_date.date() < value_date.date()
                            elif trigo_id == 22:  # After
                                trigger_matched = trans_date.date() > value_date.date()
                        except ValueError as e:
                            logger.error(f"Invalid date format in Trigger {trigger_id}: {value} ({e})")
                            trigger_matched = False
                    
                    # Transaction type triggers
                    elif trigo_id == 23:  # Withdrawal
                        trigger_matched = hai_type == 2
                    elif trigo_id == 24:  # Deposit
                        trigger_matched = hai_type == 1
                    elif trigo_id == 25:  # Transfer
                        trigger_matched = hai_type == 3
                    
                    # Status triggers
                    elif trigo_id == 26:  # Pending
                        trigger_matched = hai_stat == 2
                    elif trigo_id == 27:  # Booked
                        trigger_matched = hai_stat == 3
                    
                    logger.debug(f"Trigger {trigger_id} {'matched' if trigger_matched else 'did not match'}")
                    
                    if trigger_mode == 'Any' and trigger_matched:
                        any_trigger_matched = True
                        break
                    elif trigger_mode == 'ALL' and not trigger_matched:
                        all_triggers_matched = False
                        break
                
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
                    
                    for action_id, acto_id, value, action_sequence in actions:
                        logger.debug(f"Executing Action {action_id} (ActO_ID={acto_id}, Value={value})")
                        
                        if acto_id == 1:  # Add Tag
                            cur.execute("SELECT Tag_ID FROM Rule_Tags WHERE Rule_ID = ? AND Tag_Text = ?", (rule_id, value))
                            if not cur.fetchone():
                                cur.execute("INSERT INTO Rule_Tags (Rule_ID, Tag_Text) VALUES (?, ?)", (rule_id, value))
                                logger.debug(f"Added tag '{value}' for Rule {rule_id}")
                        
                        elif acto_id == 2:  # Remove Tag
                            cur.execute("DELETE FROM Rule_Tags WHERE Rule_ID = ? AND Tag_Text = ?", (rule_id, value))
                            logger.debug(f"Removed tag '{value}' for Rule {rule_id}")
                        
                        elif acto_id == 3:  # Remove all Tags
                            cur.execute("DELETE FROM Rule_Tags WHERE Rule_ID = ?", (rule_id,))
                            logger.debug(f"Removed all tags for Rule {rule_id}")
                        
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
                            return  # Skip further processing
                        
                        elif acto_id == 8:  # Set Category
                            try:
                                exp_id, expsub_id = map(int, value.split(','))
                                cur.execute("UPDATE HA_Import SET HAI_Exp_ID = ?, HAI_ExpSub_ID = ? WHERE HAI_ID = ?", 
                                            (exp_id, expsub_id, hai_id))
                                logger.debug(f"Set category {exp_id},{expsub_id} for HAI_ID={hai_id}")
                            except ValueError:
                                logger.error(f"Invalid category format in Action {action_id}: {value}")
                        
                        elif acto_id == 10:  # Set Description
                            cur.execute("UPDATE HA_Import SET HAI_Desc = ? WHERE HAI_ID = ?", (value, hai_id))
                            hai_desc = value
                            logger.debug(f"Set description to '{value}' for HAI_ID={hai_id}")
                        
                        elif acto_id == 11:  # Append Description
                            cur.execute("UPDATE HA_Import SET HAI_Desc = HAI_Desc || ? WHERE HAI_ID = ?", (value, hai_id))
                            hai_desc += value
                            logger.debug(f"Appended '{value}' to description for HAI_ID={hai_id}")
                        
                        elif acto_id == 12:  # Prepend Description
                            cur.execute("UPDATE HA_Import SET HAI_Desc = ? || HAI_Desc WHERE HAI_ID = ?", (value, hai_id))
                            hai_desc = value + hai_desc
                            logger.debug(f"Prepended '{value}' to description for HAI_ID={hai_id}")
                        
                        elif acto_id == 13:  # Set Destination Account
                            try:
                                acc_to = int(value)
                                cur.execute("UPDATE HA_Import SET HAI_Acc_To = ? WHERE HAI_ID = ?", (acc_to, hai_id))
                                hai_acc_to = acc_to
                                logger.debug(f"Set destination account to {acc_to} for HAI_ID={hai_id}")
                            except ValueError:
                                logger.error(f"Invalid account ID in Action {action_id}: {value}")
                        
                        elif acto_id == 14:  # Set Source Account
                            try:
                                acc_from = int(value)
                                cur.execute("UPDATE HA_Import SET HAI_Acc_From = ? WHERE HAI_ID = ?", (acc_from, hai_id))
                                hai_acc_from = acc_from
                                logger.debug(f"Set source account to {acc_from} for HAI_ID={hai_id}")
                            except ValueError:
                                logger.error(f"Invalid account ID in Action {action_id}: {value}")
                    
                    conn.commit()
                    group_proceed = rule_proceed
                
                conn.commit()
    
        logger.debug(f"Completed rule application for HAI_ID={hai_id}")
    
    except sqlite3.Error as e:
        logger.error(f"Database error in apply_rules: {e}")
        raise
    except Exception as e:
        logger.error(f"Error in apply_rules: {e}")
        raise

def cleanup_ha_import(conn, days=30):
    """
    Delete HA_Import records older than specified days.
    """
    logger.debug(f"Cleaning up HA_Import records older than {days} days")
    try:
        cur = conn.cursor()
        cutoff_date = datetime.now() - timedelta(days=days)
        cutoff_timestamp = cutoff_date.timestamp()
        
        cur.execute("""
            DELETE FROM HA_Import
            WHERE HAI_Year * 10000 + HAI_Month * 100 + HAI_Day < ?
        """, (cutoff_date.year * 10000 + cutoff_date.month * 100 + cutoff_date.day,))
        deleted = cur.rowcount
        conn.commit()
        logger.info(f"Deleted {deleted} HA_Import records older than {days} days")
    except sqlite3.Error as e:
        logger.error(f"Error cleaning up HA_Import: {e}")
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


