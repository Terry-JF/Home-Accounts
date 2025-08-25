import json
from datetime import datetime, timedelta
import sqlite3
import re
import logging


def process_transactions(json_path, conn, acc_id):
    """
    Process transactions from a JSON file into HA_Import and Trans tables.
    """
    logger = logging.getLogger('HA.transactions')
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
            
            # Matching logic
            tr_id = match_transaction(cur, hai_type, hai_day, hai_month, hai_year, hai_amount, hai_acc_from, hai_acc_to, hai_stat)
            
            if tr_id:
                # Update existing Trans record
                dow = datetime(hai_year, hai_month, hai_day).weekday() + 1
                cur.execute("""
                    UPDATE Trans
                    SET Tr_DOW = ?, Tr_Day = ?, Tr_Month = ?, Tr_Year = ?, Tr_Stat = ?, 
                        Tr_Amount = ?, Tr_FF_Journal_ID = ?
                    WHERE Tr_ID = ?
                """, (dow, hai_day, hai_month, hai_year, hai_stat, hai_amount, hai_id, tr_id))
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
            
            # Pass to rules engine (placeholder)
            apply_rules(conn, tr_id, hai_desc)
            
    except Exception as e:
        logger.error(f"Error processing transactions: {str(e)}")

def match_transaction(cur, hai_type, hai_day, hai_month, hai_year, hai_amount, hai_acc_from, hai_acc_to, hai_stat):
    """
    Match a transaction against Trans table using 31 steps.
    Returns Tr_ID if matched, None if no match.
    """
    logger = logging.getLogger('HA.transactions')
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

def apply_rules(conn, tr_id, description):      # just a placeholder
    """
    Apply rules to categorize transactions.
    """
    logger = logging.getLogger('HA.transactions')
    logger.debug(f"Applying rules to Tr_ID={tr_id}, Description={description}")

    cur = conn.cursor()
    # Default category
    cur.execute("UPDATE Trans SET Tr_Exp_ID = 99, Tr_ExpSub_ID = 99 WHERE Tr_ID = ?", (tr_id,))
    conn.commit()
    logger.debug(f"Assigned default category 99 to Tr_ID={tr_id}")

def test_rules(json_path, conn, acc_id):        # just a placeholder
    """
    Test rules on a JSON file and return results for GUI display.
    """
    logger = logging.getLogger('HA.transactions')
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











