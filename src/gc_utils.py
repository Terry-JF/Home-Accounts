import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import time
import sqlite3
import requests
import webbrowser
import subprocess
import re
import json
import time
import os
import logging
from datetime import datetime, timedelta
from uuid import uuid4
from ui_utils import (resource_path, open_form_with_position, close_form_with_position, center_window, sc)
from rules_engine import (process_transactions, test_rules)
from config import CONFIG, COLORS, get_config
import config

# Set up logger
logger = logging.getLogger('HA.gc_utils')

### GoCardless Maintenance Form and related functions ###

def validate_time_format(time_str):
    """Validate time input in HH:MM AM/PM format."""
    try:
        if not re.match(r'^(1[0-2]|0?[1-9]):[0-5][0-9] (AM|PM)$', time_str, re.IGNORECASE):
            return False
        return True
    except:
        return False

def convert_to_24hr(time_str):
    """Convert HH:MM AM/PM to 24-hour HH:MM format for schtasks."""
    try:
        # Split time and AM/PM
        time_part, period = time_str.split()
        hours, minutes = map(int, time_part.split(':'))
        if period.upper() == 'PM' and hours != 12:
            hours += 12
        elif period.upper() == 'AM' and hours == 12:
            hours = 0
        return f'{hours:02d}:{minutes:02d}'
    except:
        return None

def create_or_update_task(task_name, time1, time2=None, enabled=True):
    """Create or update Windows Scheduler task for fetch_bank_trans.py."""
    python_exe = resource_path(".venv/Scripts/python.exe")                          # Adjust if not using .venv
    script_path = resource_path("fetch_bank_trans.py")                              # Full path to script
    
    # Convert times to 24-hour format
    time1_24hr = convert_to_24hr(time1)
    if not time1_24hr:
        logger.error(f"Invalid time1 format: {time1}")
        return False
    time2_24hr = convert_to_24hr(time2) if time2 else None
    
    # Build schtasks command
    cmd = f'schtasks /CREATE /SC DAILY /TN "{task_name}" /TR "\"{python_exe}\" \"{script_path}\"" /ST {time1_24hr}'
    if time2_24hr:
        cmd += ' /RI 1440 /DU 24:00'  # Repeat daily with second trigger
    if not enabled:
        cmd = f'schtasks /CHANGE /TN "{task_name}" /DISABLE'
    
    logger.debug(f"Executing schtasks command: {cmd}")
    try:
        result = subprocess.run(cmd, shell=True, check=True, capture_output=True, text=True)
        logger.debug(f"schtasks output: {result.stdout}")
        return True
    except subprocess.CalledProcessError as e:
        logger.error(f"schtasks failed: {e.stderr}")
        return False

def delete_task(task_name):
    """Delete the specified Windows Scheduler task."""
    cmd = f'schtasks /DELETE /TN "{task_name}" /F'
    logger.debug(f"Executing schtasks command: {cmd}")
    try:
        result = subprocess.run(cmd, shell=True, check=True, capture_output=True, text=True)
        logger.debug(f"schtasks output: {result.stdout}")
        return True
    except subprocess.CalledProcessError as e:
        logger.error(f"schtasks failed: {e.stderr}")
        return False

def get_task_status(task_name):
    """Check if the task exists and is enabled."""
    cmd = f'schtasks /QUERY /TN "{task_name}" /FO CSV'
    logger.debug(f"Executing schtasks command: {cmd}")
    try:
        result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
        if task_name in result.stdout:
            return "Enabled" if "Running" in result.stdout or "Ready" in result.stdout else "Disabled"
        return "Not Found"
    except subprocess.CalledProcessError as e:
        logger.error(f"schtasks query failed: {e.stderr}")
        return "Not Found"

# Entry point for the form
def create_gocardless_maint_form(parent, conn, cursor):                 # Win_ID = 25
    """Create the GoCardless maintenance form with tabbed navigation."""
    logger.debug("Creating GC Maint Form")
    form = tk.Toplevel(parent)
    form.config(bg=config.master_bg)
    form.update_form = lambda: init_form()  # Allow setup to refresh Treeview
    win_id = 25
    form.geometry(f"{sc(800)}x{sc(700)}")
    form.transient(parent)
    form.attributes("-topmost", True)
    open_form_with_position(form, conn, cursor, win_id, "Manage GoCardless Configuration")
    
    # Get current HA year setting
    year = int(parent.year_var.get())  # Use global year_var
    
    # Create notebook for tabs
    notebook = ttk.Notebook(form)
    notebook.pack(pady=10, expand=True, fill="both")
    
    tabs = ttk.Style()
    tabs.configure('TFrame', background=config.master_bg, font=config.ha_normal)
    
    # Tab 1: Maintain GoCardless
    tab1 = ttk.Frame(notebook, style='TFrame')
    notebook.add(tab1, text="GoCardless Access")

    # Tab 2: Maintain Windows Scheduler
    tab2 = ttk.Frame(notebook, style='TFrame')
    notebook.add(tab2, text="Windows Scheduler")  

    # Tab 3: Testing GC Import (placeholder)
    tab3 = ttk.Frame(notebook, style='TFrame')
    notebook.add(tab3, text="Test Rules / Match")

    # Tab 4: Other Settings
    tab4 = ttk.Frame(notebook, style='TFrame')
    notebook.add(tab4, text="  Other Settings   ")    
    
    # Treeview
    gctree = ttk.Treeview(tab1, columns=("Acc_ID", "Acc_Name", "Active", "Days", "Bank", "Status", "Setup"), show="headings", height=12)
    gctree.place(x=sc(20), y=sc(20), width=sc(760))
    gctree.heading("Acc_ID", text="Account ID")
    gctree.heading("Acc_Name", text="Account Name")
    gctree.heading("Active", text="Active")
    gctree.heading("Days", text="Days to Fetch")
    gctree.heading("Bank", text="Bank Name")
    gctree.heading("Status", text="Link Status")
    gctree.heading("Setup", text="Setup")
    gctree.column("Acc_ID", width=sc(70), anchor="center")
    gctree.column("Acc_Name", width=sc(110))
    gctree.column("Active", width=sc(70), anchor="center")
    gctree.column("Days", width=sc(100), anchor="center")
    gctree.column("Bank", width=sc(150))
    gctree.column("Status", width=sc(100), anchor="center")
    gctree.column("Setup", width=sc(50), anchor="center")
    
    # Treeview styling
    style = ttk.Style()
    style.configure("Treeview", font=(config.ha_normal))
    style.configure("Treeview.Heading", font=(config.ha_head11))
    gctree.tag_configure("expiring", foreground="orange")
    gctree.tag_configure("expired", foreground="red")
    gctree.tag_configure("ready", background=COLORS["pale_green"])
    
    # Mappings
    active_map = {"Yes": 1, "No": 0}
    active_map_rev = {0: "No", 1: "Yes"}
    days_values = list(range(1, 91))
    
    # Form variables
    acc_id_var = tk.StringVar()
    acc_name_var = tk.StringVar()
    active_var = tk.StringVar()
    days_var = tk.StringVar()
    bank_var = tk.StringVar()
    status_var = tk.StringVar()  # Hidden var for Link Status
    filename_var = tk.StringVar()
    accounts_var = tk.StringVar()   # list of accounts
    files_var = tk.StringVar()      # list of files
    
    # Routines for Tab 1
    # Load accounts
    def init_form():
        gctree.delete(*gctree.get_children())
        now = time.time()
        cursor.execute("""
            SELECT a.Acc_ID, a.Acc_Name, COALESCE(gc.Active, 1), COALESCE(gc.Fetch_Days, 5), gc.Bank_Name, gc.Link_Status, gc.Expiry_Date
            FROM Account a
            LEFT JOIN GC_Account gc ON a.Acc_ID = gc.Acc_ID
            WHERE a.Acc_Year = ? AND a.Acc_ID < 13 ORDER BY a.Acc_ID
        """, (datetime.now().year,))
        for row in cursor.fetchall():
            acc_id, acc_name, active, days, bank, status, expiry = row
            # Update status if expired
            if status == "Linked" and expiry and now > expiry:
                cursor.execute("UPDATE GC_Account SET Link_Status = 'Expired' WHERE Acc_ID = ?", (acc_id,))
                conn.commit()
                status = "Expired"
            tags = ()
            if status == "Linked" and expiry and now > expiry - get_config('EXPIRY_WARNING_DAYS') * 86400:
                tags = ("expiring",) if now < expiry else ("expired",)
            if tags == () and status == "Linked":
                tags = ("ready")
            gctree.insert("", "end", values=(
                acc_id, acc_name, active_map_rev.get(active, "No"), days, bank or "", status or "Not Set", "Setup"
            ), tags=tags)
    
    # Handle Treeview click
    def on_gctree_select(event):
        selection = gctree.selection()
        if selection:
            values = gctree.item(selection[0], "values")
            acc_id_var.set(values[0])
            acc_name_var.set(values[1])
            active_var.set(values[2])
            days_var.set(values[3])
            bank_var.set(values[4])
            status_var.set(values[5])  # Set hidden status
            # Enable Fetch Now if conditions met
            if active_var.get() == "Yes" and status_var.get() in ["Linked", "Expiring"]:
                fetch_button.config(state="normal")
            else:
                fetch_button.config(state="disabled")
        else:
            fetch_button.config(state="disabled")
    
    # Handle Setup button click
    def on_setup_click(event):
        item = gctree.identify_row(event.y)
        if item and gctree.identify_column(event.x) == "#7":  # Setup column
            acc_id = int(gctree.item(item, "values")[0])
            setup_gc_requisition(conn, acc_id, form, form)
    
    def fetch_now():
        logger = logging.getLogger('HA.fetch_now')
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        if not acc_id_var.get():
            acc_id = None
        else:
            acc_id = int(acc_id_var.get())
        log_file = os.path.join(get_config('LOG_DIR'), f"{timestamp}_acc{acc_id}_fetch_trans.log")
        file_handler = logger.FileHandler(log_file)
        file_handler.setFormatter(logger.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        logger.addHandler(file_handler)
        logger.setLevel(logger.DEBUG)        
        
        try:
            if not acc_id_var.get():
                logger.error("No account selected for fetch")
                messagebox.showerror("Error", "No account selected.", parent=form)
                return
            acc_id = int(acc_id_var.get())
            fetch_days = int(days_var.get())
            # Get Requisition_ID from DB (since not in edit fields)
            cursor.execute("SELECT Requisition_ID FROM GC_Account WHERE Acc_ID = ?", (acc_id,))
            row = cursor.fetchone()
            if not row or not row[0]:
                logger.error(f"No Requisition ID found for account {acc_id}")
                messagebox.showerror("Error", "No Requisition ID found for this account. Setup first.", parent=form)
                return
            requisition_id = row[0]
            
            access_token = get_access_token(conn)
            if not access_token:
                logger.error("Failed to get access token")
                return
            
            # Temporary json path
            json_path = os.path.join(get_config('BANK_DIR'), f"{timestamp}_test_transactions_{acc_id}.json")
            logger.debug(f"Fetching transactions for account {acc_id} to {json_path}")        
                
            # Fetch and save to json using fetch_bank_transactions.py
            fetch_transactions(access_token, requisition_id, json_path, fetch_days, conn, acc_id)
            process_transactions(json_path, conn, acc_id)  # Process immediately
            display_transactions_form(form, json_path)
            logger.debug(f"Completed fetch and processing for account {acc_id}")
        except Exception as e:
            logger.error(f"Failed to fetch transactions: {str(e)}")
            messagebox.showerror("Error", f"Failed to fetch transactions: {str(e)}", parent=form)
        finally:
            logger.removeHandler(file_handler)
            file_handler.close()
                        
    def display_transactions_form(parent_form, json_path):
        logger = logging.getLogger()
        display_form = tk.Toplevel(parent_form)
        display_form.title(f"Transactions in json file: {json_path}")
        dx = sc(210)
        dy = sc(200)
        display_form.geometry(f"{sc(1600)}x{sc(500)}+{dx}+{dy}")
        display_form.transient(parent_form)
        display_form.attributes("-topmost", True)
        
        columns = ["transaction_id", "status", "booking_date", "amount", "currency", "description"]
        tree = ttk.Treeview(display_form, columns=columns, show="headings", height=20)
        tree.place(x=sc(20), y=sc(20), width=sc(1560))
        tree.heading("transaction_id", text="Transaction ID")
        tree.heading("status", text="Status")
        tree.heading("booking_date", text="Booking Date")
        tree.heading("amount", text="Amount")
        tree.heading("currency", text="Currency")
        tree.heading("description", text="Description")
        tree.column("transaction_id", width=sc(680), anchor="w")
        tree.column("status", width=sc(100), anchor="center")
        tree.column("booking_date", width=sc(100), anchor="center")
        tree.column("amount", width=sc(100), anchor="center")
        tree.column("currency", width=sc(100), anchor="center")
        tree.column("description", width=sc(450), anchor="w")
        
        scrollbar = tk.Scrollbar(display_form, orient="vertical", command=tree.yview)
        scrollbar.pack(side="right", fill="y")
        tree.config(yscrollcommand=scrollbar.set)
                
        try:
            with open(json_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            for account_data in data:
                for trans_type in ["booked", "pending"]:
                    for trans in account_data.get("transactions", {}).get(trans_type, []):
                        tree.insert("", "end", values=(
                            trans.get("transactionId", ""),
                            trans_type,
                            trans.get("bookingDate", ""),
                            trans["transactionAmount"]["amount"],
                            trans["transactionAmount"]["currency"],
                            trans.get("remittanceInformationUnstructured", "")
                        ))
                logger.debug(f"Displayed transactions from {json_path}")        
        except Exception as e:
            logger.error(f"Error loading JSON: {str(e)}")
            tree.insert("", "end", values=[f"Error loading JSON: {str(e)}"] + [""] * (len(columns) - 1))
        
        tk.Button(display_form, text="Close", width=15, command=display_form.destroy).place(x=sc(650), y=sc(450))
        
    def display_test_results(parent_form, results):
        logger = logging.getLogger()
        display_form = tk.Toplevel(parent_form)
        display_form.title("Test Rules Results")
        dx = sc(210)
        dy = sc(200)
        display_form.geometry(f"{sc(1600)}x{sc(500)}+{dx}+{dy}")
        display_form.transient(parent_form)
        display_form.attributes("-topmost", True)
        
        columns = ["transaction_id", "status", "booking_date", "amount", "currency", "description", "category"]
        tree = ttk.Treeview(display_form, columns=columns, show="headings", height=20)
        tree.place(x=sc(20), y=sc(20), width=sc(1560))
        tree.heading("transaction_id", text="Transaction ID")
        tree.heading("status", text="Status")
        tree.heading("booking_date", text="Booking Date")
        tree.heading("amount", text="Amount")
        tree.heading("currency", text="Currency")
        tree.heading("description", text="Description")
        tree.heading("category", text="Category")
        tree.column("transaction_id", width=sc(580), anchor="w")
        tree.column("status", width=sc(100), anchor="center")
        tree.column("booking_date", width=sc(100), anchor="center")
        tree.column("amount", width=sc(100), anchor="center")
        tree.column("currency", width=sc(100), anchor="center")
        tree.column("description", width=sc(450), anchor="w")
        tree.column("category", width=sc(100), anchor="center")
        
        scrollbar = tk.Scrollbar(display_form, orient="vertical", command=tree.yview)
        scrollbar.pack(side="right", fill="y")
        tree.config(yscrollcommand=scrollbar.set)
        
        for result in results:
            tree.insert("", "end", values=(
                result["transaction_id"], result["status"], result["booking_date"],
                result["amount"], result["currency"], result["description"], result["category"]
            ))
        logger.debug("Displayed test rules results")
        
        tk.Button(display_form, text="Close", width=15, command=display_form.destroy).place(x=sc(650), y=sc(450))
        
    def clear_fields():
        acc_id_var.set("")
        acc_name_var.set("")
        active_var.set("")
        days_var.set("")
        bank_var.set("")
        status_var.set("")
        filename_var.set("")
        gctree.selection_remove(gctree.selection())
        fetch_button.config(state="disabled")
    
    def save_account():
        logger = logging.getLogger()
        try:
            if not acc_id_var.get():
                logger.error("No account selected for save")
                messagebox.showerror("Error", "No account selected.", parent=form)
                return
            acc_id = int(acc_id_var.get())
            active = active_map.get(active_var.get(), 0)
            try:
                days = int(days_var.get())
                if not (1 <= days <= 30):
                    raise ValueError
            except ValueError:
                logger.error("Invalid Days to Fetch value")
                messagebox.showerror("Error", "Days to Fetch must be a number between 1 and 30.", parent=form)
                return
            
            cursor.execute("SELECT COUNT(*) FROM GC_Account WHERE Acc_ID = ?", (acc_id,))
            exists = cursor.fetchone()[0] > 0
            if exists:
                cursor.execute("""
                    UPDATE GC_Account 
                    SET Active = ?, Fetch_Days = ?
                    WHERE Acc_ID = ?
                """, (active, days, acc_id))
            else:
                cursor.execute("""
                    INSERT INTO GC_Account (Acc_ID, Active, Fetch_Days, Country_Code)
                    VALUES (?, ?, ?, 'GB')
                """, (acc_id, active, days))
            
            conn.commit()
            logger.info(f"Settings saved for account {acc_name_var.get()}")
            messagebox.showinfo("Success", f"Settings for {acc_name_var.get()} saved.", parent=form)
            init_form()
            clear_fields()
        except Exception as e:
            logger.error(f"Failed to save account settings: {str(e)}")
            messagebox.showerror("Error", f"Failed to save: {str(e)}", parent=form)
    
    def toggle_es():
        logger.debug(f"is_enabled before toggle: {is_enabled.get()}")
        if is_enabled.get() :
            is_enabled.set(False)   # Clear checkbox
            es_box.config(image=unchecked_img)
        else:
            is_enabled.set(True)    # Set checkbox
            es_box.config(image=checked_img)
        logger.debug(f"is_enabled after toggle: {is_enabled.get()}")
        
        test=resource_path("test")
        logger.debug(f"resource_path: {test}")
    
    # Save and Delete buttons - Tab 2
    def save_schedule():                        # Tab 2
        time1 = time1_entry.get().strip()
        time2 = time2_entry.get().strip()
        if not validate_time_format(time1):
            messagebox.showerror("Error", "Invalid Time 1 format. Use HH:MM AM/PM.")
            return
        if time2 and not validate_time_format(time2):
            messagebox.showerror("Error", "Invalid Time 2 format. Use HH:MM AM/PM.")
            return
        logger.debug(f"time1={time1}, time2={time2}")
        if create_or_update_task(task_name, time1, time2, is_enabled.get()):
            messagebox.showinfo("Success", "Task scheduled successfully.")
        else:
            messagebox.showerror("Error", "Failed to schedule task.")

    def delete_schedule():                      # Tab 2
        if delete_task(task_name):
            is_enabled.set(False)
            messagebox.showinfo("Success", "Task deleted successfully.")
        else:
            messagebox.showerror("Error", "Failed to delete task.")

    # Functions for Tab 3

# Tab 3 Functions
    def populate_accounts_combo(cursor, accounts_var, accounts_combo):
        """Populate accounts combo with Acc_ID, Acc_Name, Bank_Name, Link_Status."""
        cursor.execute("""
            SELECT a.Acc_ID, a.Acc_Name, gc.Bank_Name, gc.Link_Status
            FROM Account a
            LEFT JOIN GC_Account gc ON a.Acc_ID = gc.Acc_ID
            WHERE a.Acc_Year = ? AND a.Acc_ID < 13
            ORDER BY a.Acc_ID
        """, (datetime.now().year,))
        accounts = [
            f"{row[0]}: {row[1]} ({row[2] or 'No Bank'}, {row[3] or 'Not Set'})"
            for row in cursor.fetchall()
        ]
        accounts_var.set("")
        accounts_combo['values'] = accounts
        if accounts:
            accounts_var.set(accounts[0])
        #logger.debug(f"Populated accounts combo: {accounts}")

    def count_transaction_ids(data):
        """Recursively count objects with 'transactionId' key in JSON data."""
        count = 0
        if isinstance(data, dict):
            if "transactionId" in data or "internalTransactionId" in data:
                count += 1
            for value in data.values():
                count += count_transaction_ids(value)
        elif isinstance(data, list):
            for item in data:
                count += count_transaction_ids(item)
        return count
    
    def populate_files_combo(cursor, accounts_var, files_var):
        """Populate files combo with JSON files for the selected account."""
        if not accounts_var.get():
            files_var.set("")
            files_combo['values'] = []
            return
        acc_id = accounts_var.get().split(":")[0].strip()  # Extract Acc_ID
        bank_dir = get_config('BANK_DIR')
        files = []
        try:
            for filename in os.listdir(bank_dir):
                if filename.endswith(f"_transactions_{acc_id}.json"):
                    filepath = os.path.join(bank_dir, filename)
                    try:
                        # Extract date/time from filename (assuming YYYYMMDD_HHMMSS_AccID.json)
                        date_str = filename[:15]  # e.g., "20250827_120000"
                        import_date = datetime.strptime(date_str, "%Y%m%d_%H%M%S")
                        # Count records in JSON
                        with open(filepath, 'r') as f:
                            data = json.load(f)
                            record_count = count_transaction_ids(data)
                        files.append(f"{filename} ({import_date.strftime('%Y-%m-%d %H:%M:%S')}, {record_count} records)")
                    except (ValueError, json.JSONDecodeError) as e:
                        logger.error(f"Error processing {filename}: {e}")
                        continue
            files_var.set("")
            files_combo['values'] = sorted(files, reverse=True)  # Newest first
            if files:
                files_var.set(files[0])
            #logger.debug(f"Populated files combo for Acc_ID {acc_id}: {files}")
        except OSError as e:
            logger.error(f"Error accessing {bank_dir}: {e}")
            files_combo['values'] = []
            files_var.set("")

    def show_now():
        """Display the selected JSON file's contents in a new window."""
        if not files_var.get():
            messagebox.showerror("Error", "Please select a JSON file.")
            return
        filename = files_var.get().split(" (")[0]  # Extract filename
        filepath = os.path.join(get_config('BANK_DIR'), filename)
        try:
            with open(filepath, 'r') as f:
                data = json.load(f)
            # Create a new window to display JSON
            json_window = tk.Toplevel(form)
            json_window.title(f"JSON Content: {filename}")
            json_window.geometry(f"{sc(1000)}x{sc(605)}+{sc(400)}+{sc(200)}")
            json_window.transient(form)  # Set as transient to parent form
            json_window.attributes("-topmost", True)  # Ensure on top
            json_window.focus_set()  # Set focus to window
            text_area = tk.Text(json_window, wrap="none", font=("Courier", 10))
            text_area.insert("1.0", json.dumps(data, indent=2))
            text_area.config(state="disabled")
            text_area.place(x=sc(10), y=sc(10), height=sc(550), width=sc(960))
            scrollbar = tk.Scrollbar(json_window, orient="vertical", command=text_area.yview)
            scrollbar.place(x=sc(970), y=sc(10), height=sc(550))
            text_area.configure(yscrollcommand=scrollbar.set)
            tk.Button(json_window, text="Close", width=15, command=json_window.destroy).place(x=sc(450), y=sc(570))
            logger.debug(f"Displayed JSON file: {filepath}")
        except (OSError, json.JSONDecodeError) as e:
            messagebox.showerror("Error", f"Failed to load JSON file: {e}")
            logger.error(f"Failed to load {filepath}: {e}")

    def test_rules_now(root):
        """Import the selected JSON file into the database."""
        if not files_var.get():
            messagebox.showerror("Error", "Please select a JSON file.")
            return
        filename = files_var.get().split(" (")[0]  # Extract filename
        filepath = os.path.join(get_config('BANK_DIR'), filename)
        acc_id = accounts_var.get().split(":")[0].strip()  # Extract Acc_ID
        try:
            db_path = get_config('DB_PATH_TEST' if get_config('APP_ENV') == 'test' else 'DB_PATH')
            process_transactions(filepath, conn, acc_id)  # process JSON file
            messagebox.showinfo("Success", f"Imported {filename} successfully. \n See LOGS/app.log for details", parent=form)
            logger.debug(f"Imported JSON file: {filepath} to {db_path}")
        except (OSError, json.JSONDecodeError, sqlite3.Error) as e:
            messagebox.showerror("Error", f"Failed to import JSON file: {e}")
            logger.error(f"Failed to import {filepath}: {e}")

    # Save Settings button - Tab 4
    def save_settings():                        # Tab 4
        try:
            expiry_days = int(expiry_entry.get())
            log_days = int(log_days_entry.get())
            token_days = int(token_margin_entry.get())
            req_days = int(req_valid_entry.get())
            import_days =int(import_days_entry.get())
            
            if expiry_days < 0 or log_days < 0 or token_days < 0 or req_days < 0 or import_days < 0:
                messagebox.showerror("Error", "Values must be non-negative.")
                return
            cursor.execute("INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)", ('EXPIRY_WARNING_DAYS', expiry_days))
            cursor.execute("INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)", ('LOG_DAYS_TO_KEEP', log_days))
            cursor.execute("INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)", ('TOKEN_MARGIN', token_days))
            cursor.execute("INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)", ('REQUISITION_VALIDITY_DAYS', req_days))
            cursor.execute("INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)", ('IMPORT_DAYS_TO_KEEP', import_days))
            cursor.commit()
            CONFIG['EXPIRY_WARNING_DAYS'] = expiry_days
            CONFIG['LOG_DAYS_TO_KEEP'] = log_days
            CONFIG['TOKEN_MARGIN'] = token_days
            CONFIG['REQUISITION_VALIDITY_DAYS'] = req_days
            CONFIG['IMPORT_DAYS_TO_KEEP'] = import_days
            messagebox.showinfo("Success", "Settings updated successfully.")
        except ValueError:
            messagebox.showerror("Error", "Invalid number for settings.")
        except cursor.Error as e:
            messagebox.showerror("Error", f"Failed to update settings: {e}")

    ###### TAB LAYOUTS
    
    # Tab 1 Layout - Manage GC settings for each Account
    ###################################################################
    gctree.bind("<<TreeviewSelect>>", on_gctree_select)
    gctree.bind("<Double-1>", on_setup_click)
    
    tk.Label(tab1, text="(Click on a row to edit that record)", anchor="w", bg=config.master_bg, font=(config.ha_note), width=40).place(x=sc(20), y=sc(260))
    tk.Label(tab1, text="(To link HA account to bank account - Double-click on 'Setup')", anchor="e", bg=config.master_bg, font=(config.ha_note), width=50).place(x=sc(430), y=sc(260))
    
    tk.Label(tab1, text="Account ID:", anchor="e", bg=config.master_bg, font=(config.ha_normal), width=20).place(x=sc(150), y=sc(300))
    tk.Label(tab1, textvariable=acc_id_var, anchor="w", bg=config.master_bg, font=(config.ha_normal), width=20).place(x=sc(350), y=sc(300))
    
    tk.Label(tab1, text="Account Name:", anchor="e", bg=config.master_bg, font=(config.ha_normal), width=20).place(x=sc(150), y=sc(340))
    tk.Label(tab1, textvariable=acc_name_var, anchor="w", bg=config.master_bg, font=(config.ha_normal), width=20).place(x=sc(350), y=sc(340))
    
    tk.Label(tab1, text="Account Active:", anchor="e", bg=config.master_bg, font=(config.ha_normal), width=20).place(x=sc(150), y=sc(380))
    active_combo = ttk.Combobox(tab1, textvariable=active_var, values=["Yes", "No"], font=(config.ha_normal), width=10, state="readonly")
    active_combo.place(x=sc(320), y=sc(380))
    
    tk.Label(tab1, text="Days to Fetch:", anchor="e", bg=config.master_bg, font=(config.ha_normal), width=20).place(x=sc(150), y=sc(420))
    days_combo = ttk.Combobox(tab1, textvariable=days_var, values=days_values, font=(config.ha_normal), width=10, state="readonly")
    days_combo.place(x=sc(320), y=sc(420))
    tk.Label(tab1, text="(max 30 to save)", anchor="w", font=("Arial", 9, "italic"), bg=config.master_bg, width=20).place(x=sc(440), y=sc(420))
    
    tk.Label(tab1, text="Bank Name:", anchor="e", bg=config.master_bg, font=(config.ha_normal), width=20).place(x=sc(150), y=sc(460))
    tk.Label(tab1, textvariable=bank_var, anchor="w", bg=config.master_bg, font=(config.ha_normal), width=20).place(x=sc(350), y=sc(460))
    
    # Fetch Now Button (initially disabled)
    fetch_button = tk.Button(tab1, text="Fetch Now", font=(config.ha_button), width=20, state="disabled", command=lambda: fetch_now())
    fetch_button.place(x=sc(550), y=sc(380))

    # Save and Close Buttons
    save_btn = tk.Button(tab1, text="Save", font=(config.ha_button), width=15, command=lambda: save_account())
    save_btn.place(x=sc(150), y=sc(600))
    close_btn = tk.Button(tab1, text="Close", font=(config.ha_button), bg=COLORS["exit_but_bg"], width=15, command=lambda: close_gocardless_form(form, conn, cursor, win_id, parent))
    close_btn.place(x=sc(500), y=sc(600))
    
    ###################################################################
    # Tab 2 Layout - Task Scheduler settings
    task_name = "FetchBankTransTask"
    tk.Label(tab2, text="Windows Scheduler can be set to automatically fetch recent", bg=config.master_bg, font=(config.ha_button)).place(x=sc(200), y=sc(30))
    tk.Label(tab2, text="    transactions from your banks, up to twice per day.    ", bg=config.master_bg, font=(config.ha_button)).place(x=sc(200), y=sc(50))

    # Enable/Disable checkbox
    is_enabled = tk.BooleanVar(value=get_task_status(task_name) != "Not Found")
    #tk.Checkbutton(tab2, text="Enable Scheduler", variable=is_enabled, font=("Arial", 10)).place(x=sc(300), y=sc(100))

    unchecked_img = tk.PhotoImage(file=resource_path("icons/unchecked_16.png")).zoom(sc(1))
    checked_img = tk.PhotoImage(file=resource_path("icons/checked_16.png")).zoom(sc(1))

    tk.Label(tab2, text="Enable Scheduler", anchor=tk.W, width=sc(14), font=("Arial", 11), 
            bg=config.master_bg, fg=COLORS["black"]).place(x=sc(330), y=sc(100))
    if is_enabled:
        image=checked_img
    else:
        image=unchecked_img
    es_box=tk.Button(tab2, image=image, bg=config.master_bg, command=toggle_es)
    es_box.place(x=sc(300), y=sc(100))
    
    # Time 1 input
    tk.Label(tab2, text="Run Time 1 (HH:MM AM/PM):", bg=config.master_bg, font=("Arial", 10)).place(x=sc(200), y=sc(160))
    time1_entry = tk.Entry(tab2, width=10, font=("Arial", 10))
    time1_entry.insert(0, "06:30 AM")
    time1_entry.place(x=sc(400), y=sc(160))

    # Time 2 input (optional)
    tk.Label(tab2, text="Run Time 2 (HH:MM AM/PM):", bg=config.master_bg, font=("Arial", 10)).place(x=sc(200), y=sc(220))
    time2_entry = tk.Entry(tab2, width=10, font=("Arial", 10))
    time2_entry.insert(0, "")
    time2_entry.place(x=sc(400), y=sc(220))
    tk.Label(tab2, text="(optional)", bg=config.master_bg, font=("Arial", 10)).place(x=sc(500), y=sc(220))
    
    tk.Button(tab2, text="Save Schedule", width=20, font=(config.ha_button), command=save_schedule).place(x=sc(200), y=sc(300))
    tk.Button(tab2, text="Delete Task", width=20, bg=COLORS["del_but_bg"], font=(config.ha_button), command=delete_schedule).place(x=sc(500), y=sc(300))
    
    tk.Button(tab2, text="Close - without making changes", bg=COLORS["exit_but_bg"], width=30, font=(config.ha_button),
            command=lambda: close_gocardless_form(form, conn, cursor, win_id, parent)).place(x=sc(500), y=sc(500))
    
    # Tab 3 Layout - Test Rules and Matching
    ###################################################################
    # Populate accounts combo
    tk.Label(tab3, text="Choose Account:", anchor="e", bg=config.master_bg, font=(config.ha_normal), width=15).place(x=sc(50), y=sc(100))
    accounts_combo = ttk.Combobox(tab3, textvariable=accounts_var, width=50, state="readonly")
    accounts_combo.place(x=sc(200), y=sc(100))
    
    populate_accounts_combo(cursor, accounts_var, accounts_combo)
    
    # Populate files combo when account changes
    def on_account_select(event):
        populate_files_combo(cursor, accounts_var, files_var)
    accounts_combo.bind("<<ComboboxSelected>>", on_account_select)

    tk.Label(tab3, text="Choose File Name:", bg=config.master_bg, font=(config.ha_normal), anchor="e", width=15).place(x=sc(50), y=sc(175))
    files_combo = ttk.Combobox(tab3, textvariable=files_var, width=70, state="readonly")
    files_combo.place(x=sc(200), y=sc(175))

    show_button = tk.Button(tab3, text="Show JSON file", font=config.ha_button, width=25, command=lambda: show_now())
    show_button.place(x=sc(300), y=sc(250))

    test_import_button = tk.Button(tab3, text="Import JSON file", font=config.ha_button, width=25, command=lambda: test_rules_now(parent))
    test_import_button.place(x=sc(300), y=sc(300))

    tk.Button(tab3, text="Close", bg=COLORS["exit_but_bg"], font=config.ha_button, width=15, 
            command=lambda: close_gocardless_form(form, conn, cursor, win_id, parent)).place(x=sc(530), y=sc(530))
    
    # Manage 'pending' match string 
    test_rules_button = tk.Button(tab3, text="Manage 'pending' Match Strings", font=(config.ha_button), width=35, command=lambda: create_mrules_form(form, conn, cursor))
    test_rules_button.place(x=sc(80), y=sc(530))
    
    ###################################################################
    # Tab 4 Layout - 
    label_x = 100
    entry_x = 280
    notes_x = 350
    row_y   = 100
    
    # Expiry Warning Days
    tk.Label(tab4, text="Expiry Warning Days:", anchor="e", bg=config.master_bg, font=(config.ha_normal)).place(x=sc(label_x), y=sc(row_y))
    expiry_entry = tk.Entry(tab4, font=(config.ha_normal), width=5)
    expiry_entry.insert(0, str(get_config('EXPIRY_WARNING_DAYS')))
    expiry_entry.place(x=sc(entry_x), y=sc(row_y))
    tk.Label(tab4, text="Warn me before bank account access expires (default 10)", bg=config.master_bg, font=(config.ha_note)).place(x=sc(notes_x), y=sc(row_y))

    # Log Days to Keep
    tk.Label(tab4, text="Log Days to Keep:", anchor="e", bg=config.master_bg, font=(config.ha_normal)).place(x=sc(label_x), y=sc(row_y + 50))
    log_days_entry = tk.Entry(tab4, font=(config.ha_normal), width=5)
    log_days_entry.insert(0, str(get_config('LOG_DAYS_TO_KEEP')))
    log_days_entry.place(x=sc(entry_x), y=sc(row_y + 50))
    tk.Label(tab4, text="Delete log files older than this (default 10)", bg=config.master_bg, font=(config.ha_note)).place(x=sc(notes_x), y=sc(row_y + 50))

    # Log Days to Keep
    tk.Label(tab4, text="Token Margin:", anchor="e", bg=config.master_bg, font=(config.ha_normal)).place(x=sc(label_x), y=sc(row_y + 100))
    token_margin_entry = tk.Entry(tab4, font=(config.ha_normal), width=5)
    token_margin_entry.insert(0, str(get_config('TOKEN_MARGIN')))
    token_margin_entry.place(x=sc(entry_x), y=sc(row_y + 100))
    tk.Label(tab4, text="DO NOT CHANGE - seconds before new bank token is used (60)", bg=config.master_bg, font=(config.ha_note)).place(x=sc(notes_x), y=sc(row_y + 100))

    # Log Days to Keep
    tk.Label(tab4, text="Requisition Validity Days:", anchor="e", bg=config.master_bg, font=(config.ha_normal)).place(x=sc(label_x), y=sc(row_y + 150))
    req_valid_entry = tk.Entry(tab4, font=(config.ha_normal), width=5)
    req_valid_entry.insert(0, str(get_config('REQUISITION_VALIDITY_DAYS')))
    req_valid_entry.place(x=sc(entry_x), y=sc(row_y + 150))
    tk.Label(tab4, text="Max 90 - Request period for bank account access (default 90)", bg=config.master_bg, font=(config.ha_note)).place(x=sc(notes_x), y=sc(row_y + 150))

    # Log Days to Keep
    tk.Label(tab4, text="Import Days to Keep:", anchor="e", bg=config.master_bg, font=(config.ha_normal)).place(x=sc(label_x), y=sc(row_y + 200))
    import_days_entry = tk.Entry(tab4, font=(config.ha_normal), width=5)
    import_days_entry.insert(0, str(get_config('IMPORT_DAYS_TO_KEEP')))
    import_days_entry.place(x=sc(entry_x), y=sc(row_y + 200))
    tk.Label(tab4, text="Delete older import bank transaction log files (default 10)", bg=config.master_bg, font=(config.ha_note)).place(x=sc(notes_x), y=sc(row_y + 200))


    tk.Button(tab4, text="Save Settings", font=(config.ha_button), width=25, command=save_settings).place(x=sc(300), y=sc(400))
    
    init_form()
    
    form.protocol("WM_DELETE_WINDOW", lambda: close_gocardless_form(form, conn, cursor, win_id, parent))
    form.wait_window(form)

def create_mrules_form(parent, conn, cursor):
    form = tk.Toplevel(parent)
    form.title("Manage Description Rules")
    form.geometry(f"{sc(600)}x{sc(400)}")
    form.transient(parent)
    form.attributes("-topmost", True)
    
    tree = ttk.Treeview(form, columns=("Pattern", "Category"), show="headings", height=10)
    tree.place(x=sc(20), y=sc(20), width=sc(560))
    tree.heading("Pattern", text="Pattern")
    tree.heading("Category", text="Category")
    tree.column("Pattern", width=sc(300))
    tree.column("Category", width=sc(200))
    
    style = ttk.Style()
    style.configure("Treeview", font=("Arial", 10))
    style.configure("Treeview.Heading", font=("Arial", 11, "bold"))
    
    pattern_var = tk.StringVar()
    category_var = tk.StringVar()
    
    tk.Label(form, text="Pattern:", anchor="e", width=15).place(x=sc(20), y=sc(260))
    tk.Entry(form, textvariable=pattern_var, width=30).place(x=sc(150), y=sc(260))
    
    tk.Label(form, text="Category:", anchor="e", width=15).place(x=sc(20), y=sc(290))
    tk.Entry(form, textvariable=category_var, width=30).place(x=sc(150), y=sc(290))
    
    def load_rules():
        tree.delete(*tree.get_children())
        cursor.execute("SELECT Pattern, Category FROM Match_Rules WHERE MRule_Type = 'description'")
        for row in cursor.fetchall():
            tree.insert("", "end", values=row)
    
    def add_rule():
        pattern = pattern_var.get().strip()
        category = category_var.get().strip()
        if not pattern or not category:
            messagebox.showerror("Error", "Pattern and Category are required.", parent=form)
            return
        cursor.execute("INSERT INTO Match_Rules (MRule_Type, Pattern, Category) VALUES ('description', ?, ?)", (pattern, category))
        conn.commit()
        load_rules()
        pattern_var.set("")
        category_var.set("")
    
    def delete_rule():
        selection = tree.selection()
        if not selection:
            messagebox.showerror("Error", "No rule selected.", parent=form)
            return
        pattern = tree.item(selection[0], "values")[0]
        cursor.execute("DELETE FROM Match_Rules WHERE MRule_Type = 'description' AND Pattern = ?", (pattern,))
        conn.commit()
        load_rules()
    
    tk.Button(form, text="Add Rule", width=15, command=add_rule).place(x=sc(150), y=sc(320))
    tk.Button(form, text="Delete Rule", width=15, command=delete_rule).place(x=sc(300), y=sc(320))
    tk.Button(form, text="Close", width=15, command=form.destroy).place(x=sc(450), y=sc(320))
    
    load_rules()
    form.wait_window()

def check_requisition_status(conn, requisition_id, parent, progress_dialog=None):
    try:
        access_token = get_access_token(conn)
        if not access_token:
            if progress_dialog:
                progress_dialog.destroy()
            return None
        headers = {"Authorization": f"Bearer {access_token}"}
        response = requests.get(f"{get_config('API_BASE_URL')}/requisitions/{requisition_id}/", headers=headers)
        if response.status_code != 200:
            if progress_dialog:
                progress_dialog.destroy()
            messagebox.showerror("Error", f"Failed to check status: {response.status_code} {response.text}", parent=parent)
            return None
        return response.json()['status']
    except Exception as e:
        if progress_dialog:
            progress_dialog.destroy()
        messagebox.showerror("Error", f"Failed to check status: {str(e)}", parent=parent)
        return None

def setup_gc_requisition(conn, acc_id, parent, form):
    logger = logging.getLogger('HA.setup_gc_requisition')

    cur = conn.cursor()
    cur.execute("""
        SELECT a.Acc_Name, gc.Country_Code
        FROM Account a
        LEFT JOIN GC_Account gc ON a.Acc_ID = gc.Acc_ID
        WHERE a.Acc_ID = ? AND a.Acc_Year = ?
    """, (acc_id, datetime.now().year))
    row = cur.fetchone()
    if not row:
        logger.error(f"Account ID {acc_id} not found")
        messagebox.showerror("Error", f"Account ID {acc_id} not found.", parent=parent)
        return
    
    institution_name, country_code = row
    access_token = get_access_token(conn)
    if not access_token:
        logger.error("Failed to get access token for requisition setup")
        return
    
    headers = {"Authorization": f"Bearer {access_token}"}
    response = requests.get(f"{get_config('API_BASE_URL')}/institutions/?country={country_code}", headers=headers)
    if response.status_code != 200:
        messagebox.showerror("Error", f"Failed to fetch institutions: {response.status_code} {response.text}", parent=parent)
        return
    institutions = response.json()
    institution_id = next((inst['id'] for inst in institutions if inst['name'] == institution_name), None)
    if not institution_id:
        inst_names = [inst['name'] for inst in institutions]
        selected = tk.StringVar(parent, value=inst_names[0] if inst_names else "")
        dialog = tk.Toplevel(parent, padx=20, pady=20)
        dialog.title("Select Bank")
        dialog.transient(parent)
        dialog.attributes("-topmost", True)
        center_window(dialog, sc(400), sc(150))
        tk.Label(dialog, text="Select GoCardless Bank Name:", font=("Arial", 10)).pack(pady=5)
        combo = ttk.Combobox(dialog, textvariable=selected, font=("Arial", 10), values=inst_names, width=50)
        combo.pack(pady=5)
        tk.Button(dialog, text="OK", font=("Arial", 10), command=dialog.destroy).pack(pady=5)
        dialog.grab_set()
        dialog.wait_window()
        institution_name = selected.get()
        institution_id = next((inst['id'] for inst in institutions if inst['name'] == institution_name), None)
        if not institution_id:
            logger.error(f"Selected bank '{institution_name}' not found")
            messagebox.showerror("Error", f"Selected bank '{institution_name}' not found.", parent=parent)
            return
    
    payload = {
        "institution_id": institution_id,
        "redirect": "https://gocardless.com",
        "reference": str(uuid4())
    }
    response = requests.post(f"{get_config('API_BASE_URL')}/requisitions/", json=payload, headers=headers)
    if response.status_code != 201:
        logger.error(f"Failed to create requisition: {response.status_code} {response.text}")
        messagebox.showerror("Error", f"Failed to create requisition: {response.status_code} {response.text}", parent=parent)
        return
    data = response.json()
    
    auth_dialog = tk.Toplevel(parent, padx=20, pady=20)
    auth_dialog.title("Authorize Bank Account")
    auth_dialog.transient(parent)
    auth_dialog.attributes("-topmost", True)
    center_window(auth_dialog, sc(400), sc(200))
    tk.Label(auth_dialog, text=f"Please authorize {institution_name} in your browser:", font=("Arial", 10), wraplength=400).pack(pady=5)
    link_entry = tk.Entry(auth_dialog, font=("Arial", 10), width=50)
    link_entry.insert(0, data['link'])
    link_entry.config(state="readonly")
    link_entry.pack(pady=5)
    tk.Button(auth_dialog, text="Copy Link", font=("Arial", 10), command=lambda: parent.clipboard_clear() or parent.clipboard_append(data['link'])).pack(pady=5)
    tk.Button(auth_dialog, text="Open in Browser", font=("Arial", 10), command=lambda: webbrowser.open(data['link'])).pack(pady=5)
    tk.Label(auth_dialog, text="Waiting for authorization...", font=("Arial", 10)).pack(pady=5)
    progress = ttk.Progressbar(auth_dialog, mode="indeterminate", length=sc(300))
    progress.pack(pady=5)
    progress.start(10)
    
    def poll_status(attempts_left=60):
        if attempts_left <= 0:
            auth_dialog.destroy()
            cur.execute("UPDATE GC_Account SET Link_Status = 'Not Set' WHERE Acc_ID = ?", (acc_id,))
            conn.commit()
            form.update_form()
            logger.warning(f"Authorization timeout for Account {acc_id}")
            messagebox.showwarning("Timeout", f"Authorization for Account {acc_id} not completed. Try again.", parent=parent)
            return
        status = check_requisition_status(conn, data['id'], parent, auth_dialog)
        if status == "LN":
            auth_dialog.destroy()
            expiry = time.time() + get_config('REQUISITION_VALIDITY_DAYS') * 86400
            cur.execute("SELECT COUNT(*) FROM GC_Account WHERE Acc_ID = ?", (acc_id,))
            exists = cur.fetchone()[0] > 0
            if exists:
                cur.execute("""
                    UPDATE GC_Account 
                    SET Requisition_ID = ?, Bank_Name = ?, Country_Code = ?, Link_Status = ?, Expiry_Date = ?
                    WHERE Acc_ID = ?
                """, (data['id'], institution_name, country_code, "Linked", expiry, acc_id))
            else:
                cur.execute("""
                    INSERT INTO GC_Account (Acc_ID, Active, Fetch_Days, Requisition_ID, Country_Code, Bank_Name, Link_Status, Expiry_Date)
                    VALUES (?, 1, 5, ?, ?, ?, ?, ?)
                """, (acc_id, data['id'], country_code, institution_name, "Linked", expiry))
            conn.commit()
            form.update_form()
            logger.info(f"Account {acc_id} linked successfully")
            messagebox.showinfo("Success", f"Account {acc_id} linked successfully.", parent=parent)
        else:
            form.after(5000, lambda: poll_status(attempts_left - 1))
    
    auth_dialog.grab_set()
    form.after(0, lambda: poll_status(60))

def get_access_token(conn):
    """
    Retrieve or refresh a GoCardless access token from the GC_Admin table.
    If the refresh token is invalid (401), clear it and generate a new token.
    
    Args:
        conn: SQLite connection object
    Returns:
        str: Valid access token
    Raises:
        ValueError: If no secrets are found or token generation fails
    """
    logger.debug("Getting access token")
    
    cur = conn.cursor()
    cur.execute("SELECT Secret_ID, Secret_Key, Access_Token, Refresh_Token, Access_Expires, Refresh_Expires FROM GC_Admin")
    row = cur.fetchone()
    if not row:
        logger.error("No GoCardless secrets found in GC_Admin table")
        raise ValueError("No GoCardless secrets found in GC_Admin table")
    
    admin = dict(zip(['secret_id', 'secret_key', 'access', 'refresh', 'access_expires', 'refresh_expires'], row))
    now = time.time()
    
    # Return existing access token if valid
    if admin['access'] and admin['access_expires'] and now < admin['access_expires']:
        logger.debug("Returning valid access token")
        return admin['access']
    
    # Attempt to refresh or generate new token
    new_data = None
    if admin['refresh'] and admin['refresh_expires'] and now < admin['refresh_expires']:
        logger.debug("Refreshing access token...")
        response = requests.post(f"{get_config('API_BASE_URL')}/token/refresh/", json={"refresh": admin['refresh']})
        
        if response.status_code == 401:
            logger.debug("Refresh token invalid (401), clearing token data")
            cur.execute("""
                UPDATE GC_Admin 
                SET Access_Token = '', Refresh_Token = '', Access_Expires = '', Refresh_Expires = ''
            """)
            conn.commit()
            logger.debug("Cleared invalid token data from GC_Admin")
        elif response.status_code == 200:
            try:
                token_data = response.json()
                new_data = {
                    'secret_id': admin['secret_id'],
                    'secret_key': admin['secret_key'],
                    'access': token_data['access'],
                    'refresh': token_data.get('refresh', admin['refresh']),
                    'access_expires': now + token_data['access_expires'] - get_config('TOKEN_MARGIN'),
                    'refresh_expires': now + token_data.get('refresh_expires', admin['refresh_expires']) - get_config('TOKEN_MARGIN')
                }
                logger.debug("Successfully refreshed access token")
            except ValueError as e:
                logger.error(f"Failed to parse refresh response: {response.text} ({e})")
                raise ValueError(f"Failed to parse refresh response: {response.text}")
        else:
            logger.error(f"Failed to refresh token: {response.status_code} {response.text}")
            raise ValueError(f"Failed to refresh token: {response.status_code} {response.text}")
    
    # If refresh failed (e.g., 401) or no valid refresh token, generate new token
    if not new_data:
        if not admin['secret_id'] or not admin['secret_key']:
            logger.error("Invalid or missing Secret_ID/Secret_Key in GC_Admin")
            raise ValueError("Invalid or missing Secret_ID/Secret_Key in GC_Admin")
        
        logger.debug("Generating new access token...")
        response = requests.post(
            f"{get_config('API_BASE_URL')}/token/new/",
            json={"secret_id": admin['secret_id'], "secret_key": admin['secret_key']}
        )
        if response.status_code != 200:
            logger.error(f"Failed to generate token: {response.status_code} {response.text}")
            raise ValueError(f"Failed to generate token: {response.status_code} {response.text}")
        
        try:
            token_data = response.json()
            new_data = {
                'secret_id': admin['secret_id'],
                'secret_key': admin['secret_key'],
                'access': token_data['access'],
                'refresh': token_data['refresh'],
                'access_expires': now + token_data['access_expires'] - get_config('TOKEN_MARGIN'),
                'refresh_expires': now + token_data['refresh_expires'] - get_config('TOKEN_MARGIN')
            }
            logger.debug("Successfully generated new access token")
        except ValueError as e:
            logger.error(f"Failed to parse new token response: {response.text} ({e})")
            raise ValueError(f"Failed to parse new token response: {response.text}")
    
    # Update or insert token data into GC_Admin
    cur.execute("SELECT COUNT(*) FROM GC_Admin")
    exists = cur.fetchone()[0] > 0
    if exists:
        cur.execute("""
            UPDATE GC_Admin 
            SET Access_Token = ?, Refresh_Token = ?, Access_Expires = ?, Refresh_Expires = ?
        """, (new_data['access'], new_data['refresh'], new_data['access_expires'], new_data['refresh_expires']))
        logger.debug("Updated GC_Admin record with new token data")
    else:
        cur.execute("""
            INSERT INTO GC_Admin (Secret_ID, Secret_Key, Access_Token, Refresh_Token, Access_Expires, Refresh_Expires)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (new_data['secret_id'], new_data['secret_key'], new_data['access'], new_data['refresh'],
            new_data['access_expires'], new_data['refresh_expires']))
        logger.debug("Inserted new GC_Admin record with token data")
    
    conn.commit()
    return new_data['access']


def fetch_transactions(access_token, requisition_id, output_file, fetch_days, conn, acc_id):
    logger = logging.getLogger('HA.fetch_transactions')
    logger.debug("Fetching transactions from GC")
    
    headers = {"Authorization": f"Bearer {access_token}"}
    
    # Calculate date range (omit date_to to include all pending)
    date_from = (datetime.now() - timedelta(days=fetch_days)).strftime("%Y-%m-%d")
    
    # Get requisition data
    response = requests.get(f"{get_config('API_BASE_URL')}/requisitions/{requisition_id}/", headers=headers)
    if response.status_code != 200:
        logger.error(f"Error fetching requisition {requisition_id}: {response.status_code} {response.text}")
        if response.status_code in (403, 404):  # Invalid or expired requisition
            cur = conn.cursor()
            cur.execute("UPDATE GC_Account SET Link_Status = 'Expired' WHERE Requisition_ID = ?", (requisition_id,))
            conn.commit()
        return
    
    requisition_data = response.json()
    logger.debug(f"Requisition data for {requisition_id}: {json.dumps(requisition_data, indent=2)}")
    
    if requisition_data["status"] != "LN":
        logger.debug(f"Requisition {requisition_id} not linked. Status: {requisition_data['status']}")
        if requisition_data["status"] in ("EX", "SU"):  # Expired or suspended
            cur = conn.cursor()
            cur.execute("UPDATE GC_Account SET Link_Status = 'Expired' WHERE Requisition_ID = ?", (requisition_id,))
            conn.commit()
        return
    
    # Process each account in the requisition
    all_transactions = []
    for account_id in requisition_data["accounts"]:
        logger.debug(f"Fetching transactions for account {account_id}")
        try:
            response = requests.get(
                f"{get_config('API_BASE_URL')}/accounts/{account_id}/transactions/",
                headers=headers,
                params={"date_from": date_from}
            )
            if response.status_code != 200:
                logger.error(f"Error fetching transactions for account {account_id}: {response.status_code} {response.text}")
                continue
            trans_data = response.json()
            #logger.debug(f"Full API response for transactions: {json.dumps(trans_data, indent=2)}")
            
            # Save raw JSON with account_id and requisition_id metadata
            transactions = {
                "account_id": account_id,
                "requisition_id": requisition_id,
                "transactions": trans_data.get("transactions", {})
            }
            all_transactions.append(transactions)
        except Exception as e:
            logger.error(f"Error fetching transactions for account {account_id}: {e}")
    
    if all_transactions:
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(all_transactions, f, indent=2)
        logger.debug(f"Transactions saved to {output_file}")
        # Record import metadata
        cur = conn.cursor()
        cur.execute("""
            INSERT INTO GC_Import (Acc_ID, Import_Time, JSON_File, Processed)
            VALUES (?, ?, ?, 0)
        """, (acc_id, time.time(), output_file))
        conn.commit()
    else:
        logger.debug(f"No transactions found for requisition {requisition_id}")

def close_gocardless_form(form, conn, cursor, win_id, root):
    root.refresh_home()  # Call go_to_today() to refresh Home Form
    logger.debug("Triggered Home Form refresh after import")
    close_form_with_position(form, conn, cursor, win_id)





