# All Maintenance Forms
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import logging

from db import (fetch_years, fetch_accounts_by_year, insert_account, update_account, fetch_month_rows, save_window_position,
                fetch_lookup_values, copy_accounts_from_previous_year, bring_forward_opening_balances)

from ui_utils import (refresh_grid, resource_path, open_form_with_position, close_form_with_position)

# Setup logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

# Module-level month_names
month_names = {
    1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
    7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec"
}

# Handle DPI scaling for high-resolution displays
try:
    from ctypes import windll
    windll.shcore.SetProcessDpiAwareness(1)
except Exception as e:
    logging.warning(f"Failed to set DPI awareness: {e}")

# Configuration
REQUISITION_VALIDITY_DAYS = 90                      # GoCardless requisition validity
EXPIRY_WARNING_DAYS = 7                             # Warn if expiry within 7 days    

# Form Function to Win_ID Mappings
# Below forms are in this file
#  1: create_account_maint_form - "Manage Accounts"                                 300
#  2: create_category_maint_form - "Manage Income/Expense Categories"                   10
#  3: create_regular_transactions_maint_form - "Manage Regular Transactions"        575
#  4: Summary of Annual Performance - "Annual Accounts Summary"                     
#  5: create_account_years_maint_form - "Manage Account Years"                          10
#  6: create_annual_budget_maint_form - "Manage Annual Budget"                          10
#  7: create_transaction_options_maint_form - "Manage Transaction Options"              10
#  8: create_form_positions_maint_form - "Manage Form Window Positions"             80
#  9: create_ff_mappings_maint_form - "Manage FF Category Mappings"                     10
# 10: create_export_transactions_form - "Export Transactions to .CSV file"              10
# 11: create_summary_form - "Annual Income & Expenditure Summary"                   430
################################
# These next 6 forms are in gui.py
# 12: create_budget_form - "Annual Budget Performance"
# 13: create_compare_form - "Compare Current Year to Previous Year"
# 14: show_maint_toolbox - "List Maintenance Forms"
# 15: create_home_screen - "Home screen of HA"
# 16: create_account_list_form - "List all Transactions for a single Account"
# 17: create_edit_form - "New/Edit Transaction"
################################
# 18: create_regular_transaction_form - "Regular Transaction Entry/Edit"
# 19: show_drill_down - "Transactions making up this row of the Summary"
# 20: create_monthly_focus_form - "Focus on Actual / Budget for given Month"        180
# 21: create_colour_scheme_maint_form - "Manage Colour Scheme"                          10
# 22: create_rules_form - "Manage Rule"
# 23: create_new_group_popup - "New Rule Group"                                     40          move to gui_maint_rules
# 24: edit_rule_form - "Edit Rule"
# 25: create_gocardless_maint_form

# Next free ID: 26

# Maintenence forms

def create_account_maint_form(parent, conn, cursor, root=None):         # Win_ID = 1                                288
    form = tk.Toplevel(parent)
    win_id = 1
    open_form_with_position(form, conn, cursor, win_id, "Manage Accounts")
    scaling_factor = parent.winfo_fpixels('1i') / 96  # Get scaling factor (e.g., 2.0 for 200% scaling)
    form.geometry(f"{int(1100 * scaling_factor)}x{int(800 * scaling_factor)}")  # Adjust size
    #form.geometry("1100x800") # Hardcoded size   
    form.transient(parent)
    form.attributes("-topmost", True)

    #current_month = parent.get_current_month() if hasattr(parent, 'get_current_month') else 3
    #current_year = int(parent.year_var.get()) if hasattr(parent, 'year_var') else 2025

    # Year Combobox
    tk.Label(form, text="Select Year:", font=("Arial", 12)).place(x=int(20 * scaling_factor), y=int(20 * scaling_factor))
    year_var = tk.StringVar()
    years = fetch_years(cursor)
    year_combo = ttk.Combobox(form, textvariable=year_var, values=years, width=10, state="readonly", font=("Arial", 12))
    year_combo.place(x=int(120 * scaling_factor), y=int(20 * scaling_factor))
    year_combo.set(str(datetime.now().year))

    # Treeview
    tree = ttk.Treeview(form, columns=("SEQ", "Type", "Name", "Short", "Last4", "Credit", "Colour", "Open", "Stmt", "Prev"), show="headings", height=16)
    tree.place(x=int(20 * scaling_factor), y=int(60 * scaling_factor), width=int(1060 * scaling_factor))
    tree.heading("SEQ", text="SEQ")
    tree.heading("Type", text="Account Type")
    tree.heading("Name", text="Account Name")
    tree.heading("Short", text="Short Name")
    tree.heading("Last4", text="Last 4")
    tree.heading("Credit", text="Credit Limit")
    tree.heading("Colour", text="Colour")
    tree.heading("Open", text="Opening Balance")
    tree.heading("Stmt", text="St. Date")
    tree.heading("Prev", text="Prev. Month")
    tree.column("SEQ", width=50)
    tree.column("Type", width=100)
    tree.column("Name", width=200)
    tree.column("Short", width=100)
    tree.column("Last4", width=80, anchor="center")
    tree.column("Credit", width=100, anchor="e")
    tree.column("Colour", width=80, anchor="center")
    tree.column("Open", width=120, anchor="e")
    tree.column("Stmt", width=80, anchor="center")
    tree.column("Prev", width=80, anchor="center")

    # Mappings
    acc_types = fetch_lookup_values(cursor, 5)
    type_map = {i + 1: desc for i, desc in enumerate(acc_types)}
    reverse_type_map = {desc: i + 1 for i, desc in enumerate(acc_types)}
    colour_map = {0: "Grey", 1: "Purple", 2: "Blue", 3: "Pink", 4: "Green"}
    reverse_colour_map = {v: k for k, v in colour_map.items()}

    def populate_tree(year):
        for item in tree.get_children():
            tree.delete(item)
        accounts = fetch_accounts_by_year(cursor, year)
        if not accounts:
            form.focus_force()
            response = messagebox.askyesno(
                "No Accounts Found",
                f"No accounts exist for {year}. Do you want to create new accounts based on last year's records?",
                parent=form
            )
            if response:
                rows_copied = copy_accounts_from_previous_year(cursor, conn, year)
                if rows_copied > 0:
                    accounts = fetch_accounts_by_year(cursor, year)
                else:
                    messagebox.showinfo("No Data", f"No accounts found for {year - 1} to copy.", parent=form)
                    form.destroy()
                    return
            else:
                form.destroy()
                parent.focus_force()
                return
        for acc in accounts:
            acc_id, acc_type, name, short, last4, credit, colour, open_bal, stmt_date, prev = acc
            tree.insert("", "end", values=(
                acc_id,
                type_map.get(acc_type, str(acc_type)),
                name,
                short,
                last4,
                f"{int(credit):,}" if credit else "0",
                colour_map.get(colour, str(colour)),
                f"{open_bal:.2f}" if open_bal else "0.00",
                stmt_date if stmt_date else "",
                "Yes" if prev else ""
            ))

    def init_form():
        if year_var.get():
            populate_tree(int(year_var.get()))

    year_var.trace("w", lambda *args: populate_tree(int(year_var.get())) if year_var.get() else None)
    init_form()

    # Buttons above fields
    tk.Button(form, text="Amend Selected Account", width=28, command=lambda: amend_account()).place(x=int(250 * scaling_factor), y=int(440 * scaling_factor))
    tk.Button(form, text="Bring Forward Opening Balances", width=28, 
            command=lambda: bring_forward_handler()).place(x=int(650 * scaling_factor), y=int(440 * scaling_factor))

    # Left Side Fields
    tk.Label(form, text="Account Type:", anchor="e", width=20).place(x=int(80 * scaling_factor), y=int(500 * scaling_factor))
    type_var = tk.StringVar()
    type_combo = ttk.Combobox(form, textvariable=type_var, values=list(type_map.values()), width=15)
    type_combo.place(x=int(250 * scaling_factor), y=int(500 * scaling_factor))

    tk.Label(form, text="Account Name:", anchor="e", width=20).place(x=int(80 * scaling_factor), y=int(540 * scaling_factor))
    name_entry = tk.Entry(form, width=30)
    name_entry.place(x=int(250 * scaling_factor), y=int(540 * scaling_factor))

    tk.Label(form, text="Account Short Name:", anchor="e", width=20).place(x=int(80 * scaling_factor), y=int(580 * scaling_factor))
    short_entry = tk.Entry(form, width=15)
    short_entry.place(x=int(250 * scaling_factor), y=int(580 * scaling_factor))

    tk.Label(form, text="Account Last 4 Numbers:", anchor="e", width=20).place(x=int(80 * scaling_factor), y=int(620 * scaling_factor))
    last4_entry = tk.Entry(form, width=8)
    last4_entry.place(x=int(250 * scaling_factor), y=int(620 * scaling_factor))

    # Right Side Fields
    tk.Label(form, text="Account Credit Limit:", anchor="e", width=20).place(x=int(560 * scaling_factor), y=int(500 * scaling_factor))
    credit_entry = tk.Entry(form, width=15)
    credit_entry.place(x=int(730 * scaling_factor), y=int(500 * scaling_factor))

    tk.Label(form, text="Account Display Colour:", anchor="e", width=20).place(x=int(560 * scaling_factor), y=int(540 * scaling_factor))
    colour_var = tk.StringVar()
    colour_combo = ttk.Combobox(form, textvariable=colour_var, values=list(colour_map.values()), width=15)
    colour_combo.place(x=int(730 * scaling_factor), y=int(540 * scaling_factor))

    tk.Label(form, text="Account Opening Balance:", anchor="e", width=20).place(x=int(560 * scaling_factor), y=int(580 * scaling_factor))
    open_entry = tk.Entry(form, width=15)
    open_entry.place(x=int(730 * scaling_factor), y=int(580 * scaling_factor))

    tk.Label(form, text="Statement Date:", anchor="e", width=20).place(x=int(560 * scaling_factor), y=int(620 * scaling_factor))
    stmt_var = tk.StringVar()
    stmt_combo = ttk.Combobox(form, textvariable=stmt_var, values=[str(i) for i in range(1, 32)], width=5)
    stmt_combo.place(x=int(730 * scaling_factor), y=int(620 * scaling_factor))

    prev_var = tk.IntVar()
    #tk.Checkbutton(form, variable=prev_var, font=("Arial", 9)).place(x=int(800 * scaling_factor), y=int(620 * scaling_factor))
    #tk.Label(form, text="Statement Date is in previous month", font=("Arial", 9)).place(x=int(820 * scaling_factor), y=int(620 * scaling_factor))

    def toggle_prev():
        print(f"prev_var before toggle: {prev_var.get()}")
        if prev_var.get() == 0:
            prev_var.set(1)     # Set checkbox
            prev_box.config(image=checked_img)
        else:
            prev_var.set(0)    # clear checkbox
            prev_box.config(image=unchecked_img)
        print(f"prev_var after toggle: {prev_var.get()}")

    unchecked_img = tk.PhotoImage(file=resource_path("icons/unchecked_16.png")).zoom(int(scaling_factor))
    checked_img = tk.PhotoImage(file=resource_path("icons/checked_16.png")).zoom(int(scaling_factor))

    tk.Label(form, text="Statement Date is in previous month", anchor=tk.W, width=int(14 * scaling_factor), 
             font=("Arial", 9)).place(x=int(825 * scaling_factor), y=int(623 * scaling_factor))
    prev_box=tk.Button(form, image=unchecked_img, command=toggle_prev)
    prev_box.place(x=int(800 * scaling_factor), y=int(620 * scaling_factor))

    # Bottom Buttons
    tk.Button(form, text="Cancel", width=15, command=lambda: clear_fields()).place(x=int(20 * scaling_factor), y=int(750 * scaling_factor))
    tk.Button(form, text="Save", width=15, command=lambda: save_account()).place(x=int(450 * scaling_factor), y=int(750 * scaling_factor))
    tk.Button(form, text="Close", width=15, 
              command=lambda: close_form_with_position(form, conn, cursor, win_id)).place(x=int(900 * scaling_factor), y=int(750 * scaling_factor))

    def amend_account():
        selection = tree.selection()
        if selection:
            values = tree.item(selection[0], "values")
            type_var.set(values[1])
            name_entry.delete(0, tk.END)
            name_entry.insert(0, values[2])
            short_entry.delete(0, tk.END)
            short_entry.insert(0, values[3])
            last4_entry.delete(0, tk.END)
            last4_entry.insert(0, values[4])
            credit_entry.delete(0, tk.END)
            credit_entry.insert(0, values[5].replace(",", ""))
            colour_var.set(values[6])
            open_entry.delete(0, tk.END)
            open_entry.insert(0, values[7])
            stmt_var.set(values[8])
            prev_var.set(1 if values[9] == "Yes" else 0)
        if prev_var.get() == 1:
            prev_box.config(image=checked_img)
        else:
            prev_box.config(image=unchecked_img)

    def clear_fields():
        type_var.set("")
        name_entry.delete(0, tk.END)
        short_entry.delete(0, tk.END)
        last4_entry.delete(0, tk.END)
        credit_entry.delete(0, tk.END)
        colour_var.set("")
        open_entry.delete(0, tk.END)
        stmt_var.set("")
        prev_var.set(0)
        tree.selection_remove(tree.selection())
        prev_box.config(image=unchecked_img)

    def save_account():
        try:
            year = int(year_var.get())
            acc_type = reverse_type_map.get(type_var.get(), 0)
            name = name_entry.get().strip()
            short = short_entry.get().strip()
            last4 = last4_entry.get().strip() or ""
            credit = float(credit_entry.get().replace(",", "") or "0")
            colour = reverse_colour_map.get(colour_var.get(), 0)
            open_bal = float(open_entry.get() or "0.00")
            stmt_date = int(stmt_var.get()) if stmt_var.get() else 0
            prev = prev_var.get()

            if not (name and short):
                raise ValueError("Name and Short Name must be filled")
            if not year:
                raise ValueError("Year must be specified")
            
            selection = tree.selection()
            if selection:
                acc_id = int(tree.item(selection[0], "values")[0])
                update_account(cursor, conn, acc_id, year, acc_type, name, short, last4, credit, colour, open_bal, stmt_date, prev)
            else:
                insert_account(cursor, conn, year, acc_type, name, short, last4, credit, colour, open_bal, stmt_date, prev)
            populate_tree(year)
            clear_fields()

            # Refresh home screen
            if hasattr(parent, 'year_var') and hasattr(parent, 'tree'):
                current_year = int(parent.year_var.get())
                current_month = parent.get_current_month()

                conn.commit()
                cursor.execute("SELECT Acc_ID, Acc_Short_Name, Acc_Open, Acc_Jan, Acc_Feb, Acc_Mar, Acc_Apr, Acc_May, Acc_Jun, "
                            "Acc_Jul, Acc_Aug, Acc_Sep, Acc_Oct, Acc_Nov, Acc_Dec FROM Account WHERE Acc_Year = ? ORDER BY Acc_ID",
                            (current_year,))
                parent.account_data = cursor.fetchall()
                parent.accounts = [row[1] for row in parent.account_data]
                global accounts
                accounts = parent.accounts
                parent.rows_container[0] = fetch_month_rows(cursor, current_month, current_year, parent.accounts, parent.account_data)
                if root:  # Check if root is provided
                    root.update_ahb(root, current_month, current_year)  # Use root instead of parent
                refresh_grid(parent.tree, parent.rows_container[0], parent.marked_rows)

                # Update column headers
                for col in range(14):
                    parent.tree.heading(f"col{col+6}", text=parent.accounts[col] if col < len(parent.accounts) else f"Col{col+1}")
                    parent.button_grid[0][col].config(text=parent.accounts[col] if col < len(parent.accounts) else f"Col{col+1}")

        except ValueError as e:
            messagebox.showerror("Error", str(e), parent=form)

    def bring_forward_handler():
        year = int(year_var.get())
        global accounts
        accounts = fetch_accounts_by_year(cursor, year)
        if not accounts:
            messagebox.showwarning("No Accounts", f"No accounts exist for {year} to update.", parent=form)
            return
        updated = bring_forward_opening_balances(cursor, conn, year)
        if updated > 0:
            messagebox.showinfo("Success", f"Opening balances updated for {updated} accounts in {year}.", parent=form)
            populate_tree(year)
            if hasattr(parent, 'year_var') and hasattr(parent, 'tree'):
                current_year = int(parent.year_var.get())
                current_month = parent.get_current_month()
                cursor.execute( "SELECT Acc_ID, Acc_Short_Name, Acc_Open, Acc_Jan, Acc_Feb, Acc_Mar, Acc_Apr, Acc_May, Acc_Jun, "
                                "Acc_Jul, Acc_Aug, Acc_Sep, Acc_Oct, Acc_Nov, Acc_Dec FROM Account WHERE Acc_Year = ? ORDER BY Acc_ID",
                                (current_year,))
                parent.account_data = cursor.fetchall()
                parent.accounts = [row[1] for row in parent.account_data]
                accounts = parent.accounts
                parent.rows_container[0] = fetch_month_rows(cursor, current_month, current_year, parent.accounts, parent.account_data)
                # print(f"Rows fetched for Treeview (Bring Forward): {parent.rows_container[0]}")  # Debug print
                parent.update_ahb(current_month, current_year)
                # print(f"AHB Row 1 after bring forward: {parent.button_grid[1][0]['text']}")  # Debug AHB
                refresh_grid(parent.tree, parent.rows_container[0], parent.marked_rows)
                for col in range(14):
                    parent.tree.heading(f"col{col+5}", text=parent.accounts[col] if col < len(parent.accounts) else f"Col{col+1}")
                    parent.button_grid[0][col].config(text=parent.accounts[col] if col < len(parent.accounts) else f"Col{col+1}")
        else:
            messagebox.showwarning("No Data", f"No prior year data found for {year - 1}.", parent=form)

    form.wait_window()

def create_category_maint_form(parent, conn, cursor):                   # Win_ID = 2                                10
    form = tk.Toplevel(parent)
    win_id = 2
    open_form_with_position(form, conn, cursor, win_id, "Manage Income/Expense Categories")
    form.geometry("800x600")
    tk.Label(form, text="Manage Income/Expense Categories - Under Construction", font=("Arial", 12)).pack(pady=20)
    tk.Button(form, text="Close", width=15, 
            command=lambda: close_form_with_position(form, conn, cursor, win_id)).pack(pady=10)
    form.wait_window()

def create_colour_scheme_maint_form(parent, conn, cursor):              # Win_ID = 21                               10
    form = tk.Toplevel(parent)
    win_id = 4
    open_form_with_position(form, conn, cursor, win_id, "Manage Colour Scheme")
    form.geometry("800x600")
    tk.Label(form, text="Manage Colour Scheme - Under Construction", font=("Arial", 12)).pack(pady=20)
    tk.Button(form, text="Close", width=15, 
            command=lambda: close_form_with_position(form, conn, cursor, win_id)).pack(pady=10)
    form.wait_window()

def create_account_years_maint_form(parent, conn, cursor):              # Win_ID = 5                                10
    form = tk.Toplevel(parent)
    win_id = 5
    open_form_with_position(form, conn, cursor, win_id, "Manage Account Years")
    form.geometry("800x600")
    tk.Label(form, text="Manage Account Years - Under Construction", font=("Arial", 12)).pack(pady=20)
    tk.Button(form, text="Close", width=15, 
            command=lambda: close_form_with_position(form, conn, cursor, win_id)).pack(pady=10)
    form.wait_window()

def create_annual_budget_maint_form(parent, conn, cursor):              # Win_ID = 6                                10
    form = tk.Toplevel(parent)
    win_id = 6
    open_form_with_position(form, conn, cursor, win_id, "Manage Annual Budget")
    form.geometry("800x600")
    tk.Label(form, text="Manage Annual Budget - Under Construction", font=("Arial", 12)).pack(pady=20)
    tk.Button(form, text="Close", width=15, 
            command=lambda: close_form_with_position(form, conn, cursor, win_id)).pack(pady=10)
    form.wait_window()

def create_transaction_options_maint_form(parent, conn, cursor):        # Win_ID = 7                                10
    form = tk.Toplevel(parent)
    win_id = 7
    open_form_with_position(form, conn, cursor, win_id, "Manage Transaction Options")
    form.geometry("800x600")
    tk.Label(form, text="Manage Transaction Options - Under Construction", font=("Arial", 12)).pack(pady=20)
    tk.Button(form, text="Close", width=15, 
            command=lambda: close_form_with_position(form, conn, cursor, win_id)).pack(pady=10)
    form.wait_window()

def create_form_positions_maint_form(parent, conn, cursor):             # Win_ID = 8                                80
    form = tk.Toplevel(parent)
    win_id = 8
    open_form_with_position(form, conn, cursor, win_id, "Manage Form Window Positions")
    scaling_factor = parent.winfo_fpixels('1i') / 96  # Get scaling factor (e.g., 2.0 for 200% scaling)
    form.geometry(f"{int(800 * scaling_factor)}x{int(600 * scaling_factor)}")  # Adjust size
    #form.geometry("800x600")  # Hardcoded size

    # Treeview
    tree = ttk.Treeview(form, columns=("ID", "Name", "Left", "Top"), show="headings", height=10)
    tree.place(x=int(20 * scaling_factor), y=int(20 * scaling_factor), width=int(760 * scaling_factor))
    tree.heading("ID", text="Win ID")
    tree.heading("Name", text="Form Name")
    tree.heading("Left", text="Left")
    tree.heading("Top", text="Top")
    tree.column("ID", width=50)
    tree.column("Name", width=350)
    tree.column("Left", width=150)
    tree.column("Top", width=150)

    def populate_tree():
        for item in tree.get_children():
            tree.delete(item)
        cursor.execute("SELECT Win_ID, Win_Name, Win_Left, Win_Top FROM Windows ORDER BY Win_ID")
        for row in cursor.fetchall():
            tree.insert("", "end", values=row)

    populate_tree()

    # Fields
    tk.Label(form, text="Form Name:", anchor="e", width=15).place(x=int(20 * scaling_factor), y=int(300 * scaling_factor))
    name_entry = tk.Entry(form, width=40)
    name_entry.place(x=int(150 * scaling_factor), y=int(300 * scaling_factor))

    tk.Label(form, text="Left:", anchor="e", width=15).place(x=int(20 * scaling_factor), y=int(340 * scaling_factor))
    left_entry = tk.Entry(form, width=10)
    left_entry.place(x=int(150 * scaling_factor), y=int(340 * scaling_factor))

    tk.Label(form, text="Top:", anchor="e", width=15).place(x=int(20 * scaling_factor), y=int(380 * scaling_factor))
    top_entry = tk.Entry(form, width=10)
    top_entry.place(x=int(150 * scaling_factor), y=int(380 * scaling_factor))

    def amend_position():
        selection = tree.selection()
        if selection:
            values = tree.item(selection[0], "values")
            name_entry.delete(0, tk.END)
            name_entry.insert(0, values[1])
            left_entry.delete(0, tk.END)
            left_entry.insert(0, values[2])
            top_entry.delete(0, tk.END)
            top_entry.insert(0, values[3])

    def save_position():
        try:
            win_id = int(tree.item(tree.selection()[0], "values")[0]) if tree.selection() else None
            name = name_entry.get().strip()
            left = int(left_entry.get() or 0)
            top = int(top_entry.get() or 0)
            if not (win_id and name):
                raise ValueError("Select a form and enter a name")
            save_window_position(cursor, conn, win_id, name, left, top)
            populate_tree()
            clear_fields()
        except (IndexError, ValueError) as e:
            messagebox.showerror("Error", f"Select a form and enter valid coordinates - {e}", parent=form)

    def clear_fields():
        name_entry.delete(0, tk.END)
        left_entry.delete(0, tk.END)
        top_entry.delete(0, tk.END)
        tree.selection_remove(tree.selection())

    # Buttons
    tk.Button(form, text="Amend Selected", width=15, command=amend_position).place(x=int(500 * scaling_factor), y=int(340 * scaling_factor))
    tk.Button(form, text="Save", width=15, command=save_position).place(x=int(650 * scaling_factor), y=int(340 * scaling_factor))
    tk.Button(  form, text="Close", width=15, 
                command=lambda: close_form_with_position(form, conn, cursor, win_id)).place(x=int(650 * scaling_factor), y=int(500 * scaling_factor))

    form.wait_window()

def create_ff_mappings_maint_form(parent, conn, cursor):                # Win_ID = 9                                10
    form = tk.Toplevel(parent)
    win_id = 9
    open_form_with_position(form, conn, cursor, win_id, "Manage FF Category Mappings")
    form.geometry("800x600")
    tk.Label(form, text="Manage FF Category Mappings - Under Construction", font=("Arial", 12)).pack(pady=20)
    tk.Button(form, text="Close", width=15, 
            command=lambda: close_form_with_position(form, conn, cursor, win_id)).pack(pady=10)
    form.wait_window()

def create_export_transactions_form(parent, conn, cursor):              # Win_ID = 10                               10
    form = tk.Toplevel(parent)
    win_id = 10
    open_form_with_position(form, conn, cursor, win_id, "Export Transactions to .CSV file")
    form.geometry("800x600")
    tk.Label(form, text="Export Transactions to .CSV file - Under Construction", font=("Arial", 12)).pack(pady=20)
    tk.Button(form, text="Close", width=15, 
            command=lambda: close_form_with_position(form, conn, cursor, win_id)).pack(pady=10)
    form.wait_window()

#                                                                                                                   



