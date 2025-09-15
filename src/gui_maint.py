# gui_maint.py -  All Maintenance Forms

import tkinter as tk
from tkinter import ttk, messagebox, colorchooser
from datetime import datetime
import logging
from db import (fetch_years, fetch_accounts_by_year, insert_account, update_account, fetch_month_rows, save_window_position,
                fetch_lookup_values, copy_accounts_from_previous_year, bring_forward_opening_balances, update_lookup_color)
from ui_utils import (refresh_grid, resource_path, open_form_with_position, close_form_with_position, sc)
from config import DEFAULT_COLORS, COLORS
import config

# Set up logging
logger = logging.getLogger('HA.gui_maint')

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
    logger.warning(f"Failed to set DPI awareness: {e}")

###### Form Function to Win_ID Mappings
# Below forms are in this file
#  1: create_account_maint_form - "Manage Accounts"                               
#  2: create_category_maint_form - "Manage Income/Expense Categories"             
#  3: create_regular_transactions_maint_form - "Manage Regular Transactions"    
#  4: Summary of Annual Performance - "Annual Accounts Summary"                     
#  5: create_account_years_maint_form - "Manage Account Years"                  
#  6: create_annual_budget_maint_form - "Manage Annual Budget"                   
#  7: create_transaction_options_maint_form - "Manage Transaction Options"      
#  8: create_form_positions_maint_form - "Manage Form Window Positions" 
#  9: create_ff_mappings_maint_form - "Manage FF Category Mappings"  
# 10: create_export_transactions_form - "Export Transactions to .CSV file"
# 11: create_summary_form - "Annual Income & Expenditure Summary"
################################
# These next 6 forms are in gui.py
# 12: create_budget_form - "Annual Budget Performance"
# 13: create_compare_form - "Compare Current Year to Previous Year"
# 14: show_maint_toolbox - "List Maintenance Forms"
# 15: create_home_screen - "Home screen of HA"
# 16: create_account_list_form - "List all Transactions for a single Account" is in main.py
# 17: create_edit_form - "New/Edit Transaction"
################################
# 18: create_regular_transaction_form - "Regular Transaction Entry/Edit"
# 19: show_drill_down - "Transactions making up this row of the Summary"
# 20: create_monthly_focus_form - "Focus on Actual / Budget for given Month"    
# 21: create_colour_scheme_maint_form - "Manage Colour Scheme"                  
# 22: create_rules_form - "Manage Rule"
# 23: create_new_group_popup - "New Rule Group" popup in gui_maint_rules.py                           
# 24: edit_rule_form - "Edit Rule" in gui_maint_rules.py  
# 25: create_gocardless_maint_form in gc_utils.py  
# 26: edit_group_name - Edit Rule Group name popup in gui_maint_rules.py
# 27: delete_rule_group - Delete Rule Group confirmation popup window in gui_maint_rules.py
# 28: show_history - Transaction History form on Edit Transaction


# Next free ID: 29

# Maintenence forms

def create_account_maint_form(parent, conn, cursor, root=None):         # Win_ID = 1
    form = tk.Toplevel(parent, bg=config.master_bg)
    win_id = 1
    open_form_with_position(form, conn, cursor, win_id, "Manage Accounts")
    form.geometry(f"{sc(1100)}x{sc(800)}")
    form.transient(parent)
    form.attributes("-topmost", True)
    form.option_add("*TCombobox*Listbox*Font", config.ha_normal)  # set font size for Combobox dropdown list

    # Prepare to use custom checkbox
    prev_checkbox_var = tk.IntVar(value=0)
    unchecked_img = tk.PhotoImage(file=resource_path("icons/unchecked_16.png")).zoom(int(sc(1)))
    checked_img = tk.PhotoImage(file=resource_path("icons/checked_16.png")).zoom(int(sc(1)))
    
    # Year Combobox
    tk.Label(form, text="Select Year:", font=(config.ha_head12), bg=config.master_bg).place(x=sc(20), y=sc(20))
    year_var = tk.StringVar()
    years = fetch_years(cursor)
    year_combo = ttk.Combobox(form, textvariable=year_var, values=years, width=8, state="readonly", font=(config.ha_normal))
    year_combo.place(x=sc(120), y=sc(20))
    year_combo.set(str(datetime.now().year))

    # Treeview
    tree = ttk.Treeview(form, columns=("SEQ", "Type", "Name", "Short", "Last4", "Credit", "Colour", "Open", "Stmt", "Prev"),
                        show="headings", height=14, style="Treeview")
    tree.place(x=sc(20), y=sc(60), width=sc(1060))
    tree.tag_configure("normal", font=(config.ha_normal))
    tree.heading("SEQ", text="SEQ")
    tree.heading("Type", text="Account Type")
    tree.heading("Name", text="Account Name")
    tree.heading("Short", text="Short Name")
    tree.heading("Last4", text="Last 4")
    tree.heading("Credit", text="Credit Limit")
    tree.heading("Colour", text="Colour")
    tree.heading("Open", text="Open Bal")
    tree.heading("Stmt", text="Stat. Date")
    tree.heading("Prev", text="Prev. Month")
    tree.column("SEQ", width=30, anchor="center")
    tree.column("Type", width=120)
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
    logger.debug(f"mappings: acc_types: {acc_types}")
    type_map = {i: desc for i, desc in acc_types}
    #logger.debug(f"type_map: {type_map}")
    reverse_type_map = {desc: i for i, desc in acc_types}
    colour_map = {0: "Grey", 1: "Purple", 2: "Blue", 3: "Pink", 4: "Green"}
    reverse_colour_map = {v: k for k, v in colour_map.items()}

    def populate_tree(year):
        tags = []
        tags.append("normal")
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
            #logger.debug(f"acc_type: {acc_type}")
            tree.insert("", "end", values=(
                acc_id,
                type_map.get(acc_type, str(acc_type)),
                "   " + name,
                "   " + short,
                last4,
                f"{int(credit):,}   " if credit else "0   ",
                colour_map.get(colour, str(colour)),
                f"{open_bal:.2f}   " if open_bal else "0.00   ",
                stmt_date if stmt_date else "",
                "Yes" if prev else ""
            ), tags=tags)

    def init_form():
        if year_var.get():
            populate_tree(int(year_var.get()))

    year_var.trace("w", lambda *args: populate_tree(int(year_var.get())) if year_var.get() else None)
    init_form()

    # Buttons above fields
    tk.Button(form, text="Amend Selected Account", width=28, font=(config.ha_button), command=lambda: amend_account()).place(x=sc(250), y=sc(400))
    tk.Button(form, text="Bring Forward Opening Balances", width=28, font=(config.ha_button), 
            command=lambda: bring_forward_handler()).place(x=sc(650), y=sc(400))

    # Left Side Fields
    tk.Label(form, text="Account Type:", anchor="e", width=20, font=(config.ha_normal), bg=config.master_bg).place(x=sc(80), y=sc(500))
    type_var = tk.StringVar()
    type_combo = ttk.Combobox(form, textvariable=type_var, values=list(type_map.values()), state="readonly", width=15, font=(config.ha_normal))
    type_combo.place(x=sc(250), y=sc(500))

    tk.Label(form, text="Account Name:", anchor="e", width=20, font=(config.ha_normal), bg=config.master_bg).place(x=sc(80), y=sc(540))
    name_entry = tk.Entry(form, width=30, font=(config.ha_normal))
    name_entry.place(x=sc(250), y=sc(540))

    tk.Label(form, text="Account Short Name:", anchor="e", width=20, font=(config.ha_normal), bg=config.master_bg).place(x=sc(80), y=sc(580))
    short_entry = tk.Entry(form, width=15, font=(config.ha_normal))
    short_entry.place(x=sc(250), y=sc(580))

    tk.Label(form, text="Account Last 4 Numbers:", anchor="e", width=20, font=(config.ha_normal), bg=config.master_bg).place(x=sc(80), y=sc(620))
    last4_entry = tk.Entry(form, width=8, font=(config.ha_normal))
    last4_entry.place(x=sc(250), y=sc(620))

    # Right Side Fields
    tk.Label(form, text="Account Credit Limit:", anchor="e", width=20, font=(config.ha_normal), bg=config.master_bg).place(x=sc(560), y=sc(500))
    credit_entry = tk.Entry(form, width=15, font=(config.ha_normal))
    credit_entry.place(x=sc(730), y=sc(500))

    tk.Label(form, text="Account Display Colour:", anchor="e", width=20, font=(config.ha_normal), bg=config.master_bg).place(x=sc(560), y=sc(540))
    colour_var = tk.StringVar()
    colour_combo = ttk.Combobox(form, textvariable=colour_var, values=list(colour_map.values()), width=15, font=(config.ha_normal))
    colour_combo.place(x=sc(730), y=sc(540))

    tk.Label(form, text="Account Opening Balance:", anchor="e", width=20, font=(config.ha_normal), bg=config.master_bg).place(x=sc(560), y=sc(580))
    open_entry = tk.Entry(form, width=15, font=(config.ha_normal))
    open_entry.place(x=sc(730), y=sc(580))

    tk.Label(form, text="Statement Date:", anchor="e", width=20, font=(config.ha_normal), bg=config.master_bg).place(x=sc(560), y=sc(620))
    stmt_var = tk.StringVar()
    stmt_combo = ttk.Combobox(form, textvariable=stmt_var, values=[str(i) for i in range(1, 32)], width=5, font=(config.ha_normal))
    stmt_combo.place(x=sc(730), y=sc(620))

    tk.Label(form, text="Tick if statement Date is in previous\n month from payment date", anchor=tk.W, width=30, 
            font=(config.ha_note), bg=config.master_bg).place(x=sc(825), y=sc(618))
    prev_box = tk.Button(form, image=unchecked_img, command=lambda: toggle_prev())
    prev_box.place(x=sc(800), y=sc(623))

    def toggle_prev():
        #logger.debug(f"prev_checkbox_var before toggle: {prev_var.get()}")
        if prev_checkbox_var.get() == 1:
            prev_checkbox_var.set(0)
            prev_box.config(image=unchecked_img)
        else:
            prev_checkbox_var.set(1)
            prev_box.config(image=checked_img)
        #logger.debug(f"prev_checkbox_var after toggle: {prev_var.get()}")

    # Bottom Buttons
    tk.Button(form, text="Cancel", width=20, font=(config.ha_button), command=lambda: clear_fields()).place(x=sc(100), y=sc(740))
    tk.Button(form, text="Save", width=20, font=(config.ha_button), command=lambda: save_account()).place(x=sc(450), y=sc(740))
    tk.Button(form, text="Close", width=20, font=(config.ha_button), bg=COLORS["exit_but_bg"], 
            command=lambda: close_form_with_position(form, conn, cursor, win_id)).place(x=sc(820), y=sc(740))

    def amend_account():
        selection = tree.selection()
        if selection:
            values = tree.item(selection[0], "values")
            type_var.set(values[1])
            name_entry.delete(0, tk.END)
            name_entry.insert(0, values[2].lstrip())
            short_entry.delete(0, tk.END)
            short_entry.insert(0, values[3].lstrip())
            last4_entry.delete(0, tk.END)
            last4_entry.insert(0, values[4])
            credit_entry.delete(0, tk.END)
            credit_entry.insert(0, values[5].replace(",", ""))
            colour_var.set(values[6])
            open_entry.delete(0, tk.END)
            open_entry.insert(0, values[7])
            stmt_var.set(values[8])
            prev_checkbox_var.set(1 if values[9] == "Yes" else 0)
            prev_box.config(image=checked_img if prev_checkbox_var.get() else unchecked_img)

    def clear_fields():
        type_var.set("")
        name_entry.delete(0, tk.END)
        short_entry.delete(0, tk.END)
        last4_entry.delete(0, tk.END)
        credit_entry.delete(0, tk.END)
        colour_var.set("")
        open_entry.delete(0, tk.END)
        stmt_var.set("")
        prev_checkbox_var.set(0)
        prev_box.config(image=unchecked_img)
        tree.selection_remove(tree.selection())

    def save_account():
        try:
            year = int(year_var.get())
            acc_type = reverse_type_map.get(type_var.get(), 0)
            logger.debug(f"type_var: {type_var.get()}, acc_type: {acc_type}")
            name = name_entry.get().strip()
            short = short_entry.get().strip()
            last4 = last4_entry.get().strip() or ""
            credit = float(credit_entry.get().replace(",", "") or "0")
            colour = reverse_colour_map.get(colour_var.get(), 0)
            open_bal = float(open_entry.get() or "0.00")
            stmt_date = int(stmt_var.get()) if stmt_var.get() else 0
            prev = prev_checkbox_var.get()

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
                if root:
                    root.update_ahb(root, current_month, current_year)
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
                cursor.execute("SELECT Acc_ID, Acc_Short_Name, Acc_Open, Acc_Jan, Acc_Feb, Acc_Mar, Acc_Apr, Acc_May, Acc_Jun, "
                            "Acc_Jul, Acc_Aug, Acc_Sep, Acc_Oct, Acc_Nov, Acc_Dec FROM Account WHERE Acc_Year = ? ORDER BY Acc_ID",
                            (current_year,))
                parent.account_data = cursor.fetchall()
                parent.accounts = [row[1] for row in parent.account_data]
                accounts = parent.accounts
                parent.rows_container[0] = fetch_month_rows(cursor, current_month, current_year, parent.accounts, parent.account_data)
                parent.update_ahb(current_month, current_year)
                refresh_grid(parent.tree, parent.rows_container[0], parent.marked_rows)
                for col in range(14):
                    parent.tree.heading(f"col{col+5}", text=parent.accounts[col] if col < len(parent.accounts) else f"Col{col+1}")
                    parent.button_grid[0][col].config(text=parent.accounts[col] if col < len(parent.accounts) else f"Col{col+1}")
        else:
            messagebox.showwarning("No Data", f"No prior year data found for {year - 1}.", parent=form)

    form.wait_window()

def create_category_maint_form(parent, conn, cursor):                   # Win_ID = 2
    form = tk.Toplevel(parent)
    win_id = 2
    open_form_with_position(form, conn, cursor, win_id, "Manage Income/Expense Categories")
    form.geometry("800x600")
    tk.Label(form, text="Manage Income/Expense Categories - Under Construction", font=("Arial", 12)).pack(pady=20)
    tk.Button(form, text="Close", width=15, 
            command=lambda: close_form_with_position(form, conn, cursor, win_id)).pack(pady=10)
    form.wait_window()

def create_colour_scheme_maint_form(parent, conn, cursor, refresh_callback=None):              # Win_ID = 21
    form = tk.Toplevel(parent, bg=config.master_bg)
    win_id = 21
    open_form_with_position(form, conn, cursor, win_id, "Colours Maintenance")
    form.geometry(f"{sc(1000)}x{sc(900)}")
    form.resizable(False, False)
    form.attributes("-topmost", True)
    form.configure(bg=config.master_bg)
    form.grab_set()

    logger.debug(f"Colours Maintenance form size: width={sc(1000)}, height={sc(900)}")
    logger.debug(f"Tkinter version: {tk.TkVersion}")
    logger.debug(f"Current theme: {ttk.Style().theme_use()}")

    # Labels for rows
    labels = [
        "Main Form Background", "Main Form Background (Test Mode)", "Flag Yellow Background", "Flag Green Background",
        "Flag Blue Background", "Drill Down Marker Background", "Reconciliation Marker Background",
        "Complete Transaction - Weekday", "Pending Transaction - Weekday", "Forecast Transaction - Weekday",
        "Complete Transaction - Weekend", "Pending Transaction - Weekend", "Forecast Transaction - Weekend",
        "Daily Totals Row", "Over-Limit Daily Totals Row", "Not Selected Tab", "Selected Tab",
        "Active Button", "Active Button Highlighted", "Last Statement Balance Row",
        "Title1 - EOM Summary Header Row", "Title2 - Transaction List Header Row", "Title3 Row", "Exit/Close Buttons", "Delete Buttons"
    ]

    color_map = {
        "home_bg": (2, 1), "home_test_bg": (2, 2), "flag_y_bg": (2, 3), "flag_g_bg": (2, 4), "flag_b_bg": (2, 5),
        "flag_dd_bg": (2, 6), "flag_mk_bg": (2, 7), "tran_wk_bg": (2, 8), "tran_we_bg": (2, 9), "dtot_bg": (2, 10),
        "dtot_ol_bg": (2, 11), "tab_bg": (2, 12), "tab_act_bg": (2, 13), "act_but_bg": (2, 14), "act_but_hi_bg": (2, 15),
        "last_stat_bg": (2, 16), "title1_bg": (2, 17), "title2_bg": (2, 18), "title3_bg": (2, 19), "exit_but_bg": (2, 20),
        "del_but_bg": (2, 21), "normal_tx": (3, 1), "disabled_tx": (3, 2), "complete_tx": (3, 3), "pending_tx": (3, 4),
        "forecast_tx": (3, 5), "dtot_tx": (3, 6), "dtot_ol_tx": (3, 7), "tab_tx": (3, 8), "tab_act_tx": (3, 9),
        "act_but_tx": (3, 10), "act_but_hi_tx": (3, 11), "last_stat_tx": (3, 12), "title1_tx": (3, 13), "title2_tx": (3, 14),
        "title3_tx": (3, 15), "exit_but_tx": (3, 16), "del_but_tx": (3, 17)
    }

    # Load current colors from COLORS dict, allows user to experiment with current settings
    colors = {}
    lup_type_id = 2
    for i in range(21):
        key = [(k, v) for k, v in color_map.items() if v == (lup_type_id, i+1)]
        if key:
            colors[(lup_type_id, i+1)] = COLORS[key[0][0]]
        else:
            logger.warning(f"No color map entry for Lup_LupT_ID={lup_type_id}, Lup_Seq={i+1}")
    lup_type_id = 3
    for i in range(17):
        key = [(k, v) for k, v in color_map.items() if v == (lup_type_id, i+1)]
        if key:
            colors[(lup_type_id, i+1)] = COLORS[key[0][0]]
        else:
            logger.warning(f"No color map entry for Lup_LupT_ID={lup_type_id}, Lup_Seq={i+1}")
    logger.debug(f"colors dict: {colors}")
    logger.debug("COLORS loaded into colors dict")

    # Create UI elements using grid
    entries = []
    bg_buttons = []
    text_buttons = []

    # Column headers
    tk.Label(form, text="Background", font=(config.ha_head11), bg=config.master_bg).grid(row=0, column=2, padx=sc(20), pady=10, sticky="ew")
    tk.Label(form, text="Text", font=(config.ha_head11), bg=config.master_bg).grid(row=0, column=3, padx=sc(20), pady=10, sticky="ew")
    logger.debug("Created column headers")

    # Create rows
    for i, label_text in enumerate(labels):
        # Label
        tk.Label(form, text=label_text, font=(config.ha_normal), bg=config.master_bg, anchor="e", width=30).grid(row=i+1, column=0, padx=sc(20), pady=2, sticky="e")
        
        # Entry
        entry = tk.Entry(form, font=(config.ha_normal), bg=config.master_bg, width=30, justify="center")
        entry.insert(0, " TEST TEXT 123.45 678.90 ")
        entry.config(state="readonly")
        entry.grid(row=i+1, column=1, padx=sc(20), pady=2, sticky="w")
        entries.append(entry)
        logger.debug(f"Created entry for row {i}: row={i+1}, column=1")

        # Background button
        if i <= 6 or i >= 13 or i == 8 or i == 11:  # Rows 0-6, 13-24, weekday (8), weekend (11)
            bg_btn = tk.Button(form, text="Select Colour", font=(config.ha_button), width=17, command=lambda idx=i: select_color(idx, 0))
            bg_btn.grid(row=i+1, column=2, padx=sc(20), pady=2, sticky="ew")
            bg_buttons.append(bg_btn)
            logger.debug(f"Created background button for row {i}")
        else:
            bg_buttons.append(None)

        # Text button
        if i in (7, 8, 9) or i >= 13:  # Rows 7, 8, 9 and 13-24
            text_btn = tk.Button(form, text="Select Colour", font=(config.ha_button), width=17, command=lambda idx=i: select_color(idx, 1))
            text_btn.grid(row=i+1, column=3, padx=sc(20), pady=2, sticky="ew")
            text_buttons.append(text_btn)
            logger.debug(f"Created text button for row {i}")
        else:
            text_buttons.append(None)

    # Bottom buttons
    save_btn = tk.Button(form, text="Save", font=(config.ha_button), width=15, command=lambda: save_colors())
    save_btn.place(x=sc(100), y=sc(830), width=sc(150), height=sc(35))
    ### Under construction ###
    apply_btn = tk.Button(form, text="Apply", font=(config.ha_button), width=15, command=lambda: apply_colors(refresh_callback))
    apply_btn.place(x=sc(300), y=sc(830), width=sc(150), height=sc(35))
    
    reset_btn = tk.Button(form, text="Reset All to Defaults", font=(config.ha_button), width=30, command=lambda: reset_colors())
    reset_btn.place(x=sc(500), y=sc(830), width=sc(200), height=sc(35))
    close_btn = tk.Button(form, text="Close", font=(config.ha_button), bg=COLORS["exit_but_bg"], width=15, command=lambda: close_form_with_position(form, conn, cursor, win_id))
    close_btn.place(x=sc(750), y=sc(830), width=sc(150), height=sc(35))
    logger.debug("Created bottom buttons")

    def select_color(row_idx, bk_txt):
        if bk_txt == 0:                     # background colour
            lup_type_id = 2 
            if row_idx <= 6:
                lup_seq = row_idx + 1
            elif row_idx == 8:              # weekday background
                lup_seq = row_idx
            elif row_idx == 11:             # weekday background
                lup_seq = row_idx - 2
            elif row_idx >= 13:
                lup_seq = row_idx - 3
        else:
            lup_type_id = 3                 # text colour
            if row_idx in (7, 8, 9):
                lup_seq = row_idx - 4
            elif row_idx >= 13:
                lup_seq = row_idx - 7
        
        current_color = colors[lup_type_id, lup_seq]
        color = colorchooser.askcolor(color=current_color, title=f"Select {'Background' if bk_txt == 0 else 'Text'} Color", parent=form)
        if color[1]:
            colors[(lup_type_id, lup_seq)] = color[1]
        logger.debug(f"Color selected: lup_type_id={lup_type_id}, lup_seq={lup_seq}, type={'background' if bk_txt == 0 else 'text'}, color={color[1]}")
        update_entries()    

    def reset_colors():
        # Copies the default colours to the colors dict, these will then show on the form - not saved anywhere
        #COLORS.update(DEFAULT_COLORS)
        for lup_type_id, lup_seq in colors:
            key = [(k, v) for k, v in color_map.items() if v == (lup_type_id, lup_seq)]
            if key:
                colors[(lup_type_id, lup_seq)] = DEFAULT_COLORS[key[0][0]]
            else:
                logger.warning(f"No color map entry for Lup_LupT_ID={lup_type_id}, Lup_Seq={lup_seq}")
        update_entries()
        logger.debug("Colors reset to defaults (locally - not saved)")

    def save_colors():
        # Saves the colors dict to the database Lookups, these will then load on next restart of app
        try:
            for (lup_type_id, lup_seq), value in colors.items():
                logger.debug(f"lup_type_id= {lup_type_id}, lup_seq= {lup_seq}, value= {value}")
                update_lookup_color(cursor, conn, lup_seq, lup_type_id, value)
            conn.commit()
            messagebox.showinfo("Success", "Colors saved successfully", parent=form)
            logger.debug("Colors saved to Lookups table")
        except Exception as e:
            logger.error(f"Failed to save colors: {e}")
            messagebox.showerror("Error", f"Failed to save colors: {e}", parent=form)
            
    def apply_colors(refresh_callback):
        """Saves the colors dict back to the COLORS dict, applying to the current app without saving to DB."""
        for lup_type_id, lup_seq in colors:
            # Find the COLORS key that maps to (lup_type_id, lup_seq) in color_map
            for color_key, color_tuple in color_map.items():
                if color_tuple == (lup_type_id, lup_seq):
                    COLORS[color_key] = colors[(lup_type_id, lup_seq)]
                    logger.debug(f"Updated COLORS[{color_key}] = {colors[(lup_type_id, lup_seq)]}")
                    break
            else:
                logger.warning(f"No color_map entry for Lup_LupT_ID={lup_type_id}, Lup_Seq={lup_seq}")
        
        logger.debug("COLORS updated with colors dict")
        logger.debug(f"COLORS: {COLORS}")
        logger.debug("apply_colors called")
        
        # Trigger the home form's refresh
        if refresh_callback:
            refresh_callback()
            logger.debug("Refresh callback triggered")        
        
    def update_entries():       # uses COLORS dictionary to set entries[] background and foreground colours
        lup_type_id = 2
        for i in range(25):     # loads all background colours to entry fields
            if i <= 6:
                entries[i].config(readonlybackground=colors[lup_type_id, i+1])
                temp=colors[(lup_type_id, i+1)]
                logger.debug(f"update_entries bg - entries[{i}]: colors: {temp}")
            elif i in (7, 8, 9):
                entries[i].config(readonlybackground=colors[lup_type_id, 8])
                temp=colors[(lup_type_id, 8)]
                logger.debug(f"update_entries bg - entries[{i}]: colors:{temp}")
            elif i in (10, 11, 12):
                entries[i].config(readonlybackground=colors[lup_type_id, 9])
                temp=colors[(lup_type_id, 9)]
                logger.debug(f"update_entries bg - entries[{i}]: colors:{temp}")
            elif i >= 13:
                entries[i].config(readonlybackground=colors[lup_type_id, i-3])
                temp=colors[(lup_type_id, i-3)]
                logger.debug(f"update_entries bg - entries[{i}]: colors:{temp}")
        lup_type_id = 3
        for i in range(17):     # load all foreground colours into entry fields
            if i in (0, 1):
                logger.debug(f"first 2 text colors [{i}]: skipped")
            if i in (2, 3, 4):
                entries[i+5].config(fg=colors[lup_type_id, i+1])
                entries[i+8].config(fg=colors[lup_type_id, i+1])
                temp=colors[(lup_type_id, i+1)]
                logger.debug(f"update_entries fg - entries[{i+5}] + [{i+8}]: colors:{temp}")
            if i >= 5:
                entries[i+8].config(fg=colors[(lup_type_id, i+1)])
                temp=colors[(lup_type_id, i+1)]
                logger.debug(f"update_entries fg - entries[{i+8}]: colors:{temp}")
        entries[i].update()
        logger.debug("Entries updated with current colors")
        #logger.debug(f"Colors dictionary: {colors}")

    update_entries()
    form.wait_window()

def create_account_years_maint_form(parent, conn, cursor):              # Win_ID = 5
    form = tk.Toplevel(parent)
    win_id = 5
    open_form_with_position(form, conn, cursor, win_id, "Manage Account Years")
    form.geometry("800x600")
    tk.Label(form, text="Manage Account Years - Under Construction", font=("Arial", 12)).pack(pady=20)
    tk.Button(form, text="Close", width=15, 
            command=lambda: close_form_with_position(form, conn, cursor, win_id)).pack(pady=10)
    form.wait_window()

def create_annual_budget_maint_form(parent, conn, cursor):              # Win_ID = 6
    form = tk.Toplevel(parent)
    win_id = 6
    open_form_with_position(form, conn, cursor, win_id, "Manage Annual Budget")
    form.geometry("800x600")
    tk.Label(form, text="Manage Annual Budget - Under Construction", font=("Arial", 12)).pack(pady=20)
    tk.Button(form, text="Close", width=15, 
            command=lambda: close_form_with_position(form, conn, cursor, win_id)).pack(pady=10)
    form.wait_window()

def create_transaction_options_maint_form(parent, conn, cursor):        # Win_ID = 7
    form = tk.Toplevel(parent)
    win_id = 7
    open_form_with_position(form, conn, cursor, win_id, "Manage Transaction Options")
    form.geometry("800x600")
    tk.Label(form, text="Manage Transaction Options - Under Construction", font=("Arial", 12)).pack(pady=20)
    tk.Button(form, text="Close", width=15, 
            command=lambda: close_form_with_position(form, conn, cursor, win_id)).pack(pady=10)
    form.wait_window()

def create_form_positions_maint_form(parent, conn, cursor):             # Win_ID = 8
    form = tk.Toplevel(parent)
    win_id = 8
    open_form_with_position(form, conn, cursor, win_id, "Manage Form Window Positions")
    form.geometry(f"{sc(800)}x{sc(600)}")
    form.config(bg=config.master_bg)

    # Treeview styling
    style = ttk.Style()
    style.configure("Treeview", font=(config.ha_normal))

    # Treeview
    tree = ttk.Treeview(form, columns=("ID", "Name", "Left", "Top"), show="headings", height=20)
    tree.place(x=sc(20), y=sc(20), width=sc(760))
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
    tk.Label(form, text="Form Name:", font=(config.ha_normal), bg=config.master_bg, anchor="e", width=15).place(x=sc(20), y=sc(420))
    name_entry = tk.Entry(form, width=40, font=(config.ha_normal))
    name_entry.place(x=sc(150), y=sc(420))

    tk.Label(form, text="Left:", font=(config.ha_normal), bg=config.master_bg, anchor="e", width=15).place(x=sc(20), y=sc(460))
    left_entry = tk.Entry(form, width=10, font=(config.ha_normal))
    left_entry.place(x=sc(150), y=sc(460))

    tk.Label(form, text="Top:", font=(config.ha_normal), bg=config.master_bg, anchor="e", width=15).place(x=sc(20), y=sc(500))
    top_entry = tk.Entry(form, width=10, font=(config.ha_normal))
    top_entry.place(x=sc(150), y=sc(500))

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
    tk.Button(form, text="Amend Selected", font=(config.ha_button), width=15, command=amend_position).place(x=sc(500), y=sc(420))
    tk.Button(form, text="Save", font=(config.ha_button), width=15, command=save_position).place(x=sc(100), y=sc(550))
    tk.Button(form, text="Close", font=(config.ha_button), bg=COLORS["exit_but_bg"], width=15, 
            command=lambda: close_form_with_position(form, conn, cursor, win_id)).place(x=sc(600), y=sc(550))

    form.wait_window()

def create_ff_mappings_maint_form(parent, conn, cursor):                # Win_ID = 9
    form = tk.Toplevel(parent)
    win_id = 9
    open_form_with_position(form, conn, cursor, win_id, "Manage FF Category Mappings")
    form.geometry("800x600")
    tk.Label(form, text="Manage FF Category Mappings - Under Construction", font=("Arial", 12)).pack(pady=20)
    tk.Button(form, text="Close", width=15, 
            command=lambda: close_form_with_position(form, conn, cursor, win_id)).pack(pady=10)
    form.wait_window()

def create_export_transactions_form(parent, conn, cursor):              # Win_ID = 10
    form = tk.Toplevel(parent)
    win_id = 10
    open_form_with_position(form, conn, cursor, win_id, "Export Transactions to .CSV file")
    form.geometry("800x600")
    tk.Label(form, text="Export Transactions to .CSV file - Under Construction", font=("Arial", 12)).pack(pady=20)
    tk.Button(form, text="Close", width=15, 
            command=lambda: close_form_with_position(form, conn, cursor, win_id)).pack(pady=10)
    form.wait_window()
    
    
    
    