# gui_maint.py -  All Maintenance Forms

import tkinter as tk
from tkinter import ttk, messagebox, colorchooser
from datetime import datetime
import logging
import sqlite3
from db import (fetch_accounts_by_year, insert_account, update_account, fetch_month_rows, save_window_position, copy_categories_from_previous_year,
                fetch_lookup_values, copy_accounts_from_previous_year, bring_forward_opening_balances, update_lookup_color)
from ui_utils import (refresh_grid, resource_path, open_form_with_position, close_form_with_position, sc)
from config import DEFAULT_COLORS, COLORS, ICON_CACHE
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
# 29: create_category_maint_form - sub-forms - New/Edit Category

# Next free ID: 30

# Maintenence forms

def create_account_maint_form(parent, conn, cursor, tree_refresh_callback, root=None):         # Win_ID = 1
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
    
    # Year Label
    root_year = root.year_var.get()
    tk.Label(form, text=f"Year: {root_year}", font=(config.ha_head12), bg=config.master_bg).place(x=sc(20), y=sc(20))

    # Treeview
    tree = ttk.Treeview(form, columns=("SEQ", "Type", "Name", "Short", "Last4", "Credit", "Colour", "Open", "Stmt", "Prev"),
                        show="headings", height=14, style="Treeview")
    tree.place(x=sc(20), y=sc(60), width=sc(1060))
    tk.Label(form, text="Select a row to amend", font=(config.ha_normal), bg=config.master_bg).place(x=sc(40), y=sc(360))
    tk.Label(form, text="^ ^ ^", font=(config.ha_vlarge), bg=config.master_bg).place(x=sc(820), y=sc(365))
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
    type_map = {i: desc for i, desc in acc_types}
    reverse_type_map = {desc: i for i, desc in acc_types}
    colour_map = {0: "Grey", 1: "Purple", 2: "Blue", 3: "Pink", 4: "Green"}
    reverse_colour_map = {v: k for k, v in colour_map.items()}

    def populate_tree(root_year):
        tags = []
        tags.append("normal")
        for item in tree.get_children():
            tree.delete(item)
        accounts = fetch_accounts_by_year(cursor, root_year)
        if not accounts:
            form.focus_force()
            response = messagebox.askyesno(
                "No Accounts Found",
                f"No accounts exist for {root_year}. Do you want to create new accounts based on last year's records?",
                parent=form
            )
            if response:
                rows_copied = copy_accounts_from_previous_year(cursor, conn, root_year)
                if rows_copied > 0:
                    accounts = fetch_accounts_by_year(cursor, root_year)
                else:
                    messagebox.showinfo("No Data", f"No accounts found for {root_year - 1} to copy.", parent=form)
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
                "   " + name,
                "   " + short,
                last4,
                f"{int(credit):,}   " if credit else "0   ",
                colour_map.get(colour, str(colour)),
                f"{open_bal:.2f}   " if open_bal else "0.00   ",
                stmt_date if stmt_date else "",
                "Yes" if prev else ""
            ), tags=tags)

    populate_tree(root_year)

    # Buttons above fields
    amend_btn = tk.Button(form, text="Amend Selected Account", width=28, font=(config.ha_button), command=lambda: amend_account())
    amend_btn.place(x=sc(120), y=sc(390))
    bringfwd_btn = tk.Button(form, text="Bring Forward Opening Balances", width=28, font=(config.ha_button), command=lambda: bring_forward_handler())
    bringfwd_btn.place(x=sc(720), y=sc(380))
    tk.Label(form, text="from end of previous year", font=(config.ha_note), bg=config.master_bg).place(x=sc(770), y=sc(411))

    # Frame for edit fields
    edit_frame = tk.LabelFrame(form, text="  Account Fields  ", font=(config.ha_button), bg=config.master_bg)
    edit_frame.place(x=sc(20), y=sc(430), width=sc(1060), height=sc(280))
    
    # Left Side Fields
    tk.Label(edit_frame, text="Account Type:", anchor="e", width=20, font=(config.ha_normal), bg=config.master_bg).place(x=sc(60), y=sc(20))
    type_var = tk.StringVar()
    type_combo = ttk.Combobox(edit_frame, textvariable=type_var, values=list(type_map.values()), state="disabled", width=15, font=(config.ha_normal))
    type_combo.place(x=sc(230), y=sc(20))

    tk.Label(edit_frame, text="Account Name:", anchor="e", width=20, font=(config.ha_normal), bg=config.master_bg).place(x=sc(60), y=sc(60))
    name_entry = tk.Entry(edit_frame, width=30, font=(config.ha_normal), state="disabled")
    name_entry.place(x=sc(230), y=sc(60))

    tk.Label(edit_frame, text="Account Short Name:", anchor="e", width=20, font=(config.ha_normal), bg=config.master_bg).place(x=sc(60), y=sc(100))
    short_entry = tk.Entry(edit_frame, width=15, font=(config.ha_normal), state="disabled")
    short_entry.place(x=sc(230), y=sc(100))

    tk.Label(edit_frame, text="Account Last 4 Numbers:", anchor="e", width=20, font=(config.ha_normal), bg=config.master_bg).place(x=sc(60), y=sc(140))
    last4_entry = tk.Entry(edit_frame, width=8, font=(config.ha_normal), state="disabled")
    last4_entry.place(x=sc(230), y=sc(140))

    # Right Side Fields
    tk.Label(edit_frame, text="Account Credit Limit:", anchor="e", width=20, font=(config.ha_normal), bg=config.master_bg).place(x=sc(540), y=sc(20))
    credit_entry = tk.Entry(edit_frame, width=15, font=(config.ha_normal), state="disabled")
    credit_entry.place(x=sc(710), y=sc(20))

    tk.Label(edit_frame, text="Account Display Colour:", anchor="e", width=20, font=(config.ha_normal), bg=config.master_bg).place(x=sc(540), y=sc(60))
    colour_var = tk.StringVar()
    colour_combo = ttk.Combobox(edit_frame, textvariable=colour_var, values=list(colour_map.values()), width=15, font=(config.ha_normal), state="disabled")
    colour_combo.place(x=sc(710), y=sc(60))

    tk.Label(edit_frame, text="Account Opening Balance:", anchor="e", width=20, font=(config.ha_normal), bg=config.master_bg).place(x=sc(540), y=sc(100))
    open_entry = tk.Entry(edit_frame, width=15, font=(config.ha_normal), state="disabled")
    open_entry.place(x=sc(710), y=sc(100))

    tk.Label(edit_frame, text="CC Statement Date:", anchor="e", width=20, font=(config.ha_normal), bg=config.master_bg).place(x=sc(540), y=sc(140))
    stmt_var = tk.StringVar()
    stmt_combo = ttk.Combobox(edit_frame, textvariable=stmt_var, values=[str(i) for i in range(1, 32)], width=5, font=(config.ha_normal), state="disabled")
    stmt_combo.place(x=sc(710), y=sc(140))

    tk.Label(edit_frame, text="Tick if statement Date is in previous\n month from payment date", anchor=tk.W, width=30, 
            font=(config.ha_note), bg=config.master_bg).place(x=sc(805), y=sc(138))
    prev_box = tk.Button(edit_frame, state="disabled", image=unchecked_img, command=lambda: toggle_prev())
    prev_box.place(x=sc(780), y=sc(138))
    tk.Label(edit_frame, text="(Enter a date to make the 'Last Statement Balance' field visible\n for this account on the main form)", font=(config.ha_note), bg=config.master_bg).place(x=sc(580), y=sc(175))

    # Cancel & Save buttons inside frame
    cancel_btn = tk.Button(edit_frame, text="Cancel", width=20, font=(config.ha_button), state="disabled", command=lambda: clear_fields())
    cancel_btn.place(x=sc(80), y=sc(200))
    save_btn = tk.Button(edit_frame, text="Save", width=20, font=(config.ha_button), state="disabled", command=lambda: save_account())
    save_btn.place(x=sc(380), y=sc(200))

    def toggle_prev():
        #logger.debug(f"prev_checkbox_var before toggle: {prev_var.get()}")
        if prev_checkbox_var.get() == 1:
            prev_checkbox_var.set(0)
            prev_box.config(image=unchecked_img)
        else:
            prev_checkbox_var.set(1)
            prev_box.config(image=checked_img)
        #logger.debug(f"prev_checkbox_var after toggle: {prev_var.get()}")

    # Close Button
    tk.Button(form, text="Close", width=20, font=(config.ha_button), bg=COLORS["exit_but_bg"], 
            command=lambda: [close_form_with_position(form, conn, cursor, win_id), tree_refresh_callback()]).place(x=sc(820), y=sc(740))

    def amend_account():
        selection = tree.selection()
        if selection:
            values = tree.item(selection[0], "values")
            type_var.set(values[1])
            type_combo.config(state="normal")
            name_entry.config(state="normal")
            name_entry.delete(0, tk.END)
            name_entry.insert(0, values[2].lstrip())
            short_entry.config(state="normal")
            short_entry.delete(0, tk.END)
            short_entry.insert(0, values[3].lstrip())
            last4_entry.config(state="normal")
            last4_entry.delete(0, tk.END)
            last4_entry.insert(0, values[4])
            credit_entry.config(state="normal")
            credit_entry.delete(0, tk.END)
            credit_entry.insert(0, values[5].replace(",", ""))
            colour_var.set(values[6])
            colour_combo.config(state="normal")
            open_entry.config(state="normal")
            open_entry.delete(0, tk.END)
            open_entry.insert(0, values[7])
            stmt_var.set(values[8])
            stmt_combo.config(state="normal")
            prev_checkbox_var.set(1 if values[9] == "Yes" else 0)
            prev_box.config(state="normal")
            prev_box.config(image=checked_img if prev_checkbox_var.get() else unchecked_img)
            cancel_btn.config(state="normal")
            save_btn.config(state="normal")

    def clear_fields():
        type_var.set("")
        type_combo.config(state="disabled")
        name_entry.delete(0, tk.END)
        name_entry.config(state="disabled")
        short_entry.delete(0, tk.END)
        short_entry.config(state="disabled")
        last4_entry.delete(0, tk.END)
        last4_entry.config(state="disabled")
        credit_entry.delete(0, tk.END)
        credit_entry.config(state="disabled")
        colour_var.set("")
        colour_combo.config(state="disabled")
        open_entry.delete(0, tk.END)
        open_entry.config(state="disabled")
        stmt_var.set("")
        stmt_combo.config(state="disabled")
        prev_checkbox_var.set(0)
        prev_box.config(image=unchecked_img)
        prev_box.config(state="disabled")
        tree.selection_remove(tree.selection())
        cancel_btn.config(state="disabled")
        save_btn.config(state="disabled")

    def save_account():
        try:
            acc_type = reverse_type_map.get(type_var.get(), 0)
            #logger.debug(f"type_var: {type_var.get()}, acc_type: {acc_type}")
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
            if not root_year:
                raise ValueError("Year must be specified")
            
            selection = tree.selection()
            if selection:
                acc_id = int(tree.item(selection[0], "values")[0])
                update_account(cursor, conn, acc_id, root_year, acc_type, name, short, last4, credit, colour, open_bal, stmt_date, prev)
            else:
                insert_account(cursor, conn, root_year, acc_type, name, short, last4, credit, colour, open_bal, stmt_date, prev)
            populate_tree(root_year)
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
                current = datetime.now()
                is_current_month = True if current_month == current.month and current_year == current.year else False
                refresh_grid(parent.tree, parent.rows_container[0], is_current_month, parent.marked_rows)
                
                # Update column headers
                for col in range(14):
                    parent.tree.heading(f"col{col+6}", text=parent.accounts[col] if col < len(parent.accounts) else f"Col{col+1}")
                    parent.button_grid[0][col].config(text=parent.accounts[col] if col < len(parent.accounts) else f"Col{col+1}")

        except ValueError as e:
            messagebox.showerror("Error", str(e), parent=form)

    def bring_forward_handler():
        global accounts
        accounts = fetch_accounts_by_year(cursor, root_year)
        if not accounts:
            messagebox.showwarning("No Accounts", f"No accounts exist for {root_year} to update.", parent=form)
            return
        updated = bring_forward_opening_balances(cursor, conn, root_year)
        if updated > 0:
            messagebox.showinfo("Success", f"Opening balances updated for {updated} accounts in {root_year}.", parent=form)
            populate_tree(root_year)
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
                current = datetime.now()
                is_current_month = True if current_month == current.month and current_year == current.year else False
                refresh_grid(parent.tree, parent.rows_container[0], is_current_month, parent.marked_rows)
                for col in range(14):
                    parent.tree.heading(f"col{col+5}", text=parent.accounts[col] if col < len(parent.accounts) else f"Col{col+1}")
                    parent.button_grid[0][col].config(text=parent.accounts[col] if col < len(parent.accounts) else f"Col{col+1}")
        else:
            messagebox.showwarning("No Data", f"No prior year data found for {root_year - 1}.", parent=form)

    form.wait_window()

def create_category_maint_form(parent, conn, cursor, tree_refresh_callback=None):
    """
    Create a form to manage Income/Expense categories from IE_Cata table.
    
    Args:
        parent: Parent Tkinter widget
        conn: SQLite connection
        cursor: SQLite cursor
    """
    year = int(parent.year_var.get())  # Current HA year
    #logger.debug(f"Opening Manage Income/Expense Categories form for year {year}")

    # Main form
    form = tk.Toplevel(parent)
    form.title("Income & Expenditure Category Maintenance")
    form.transient(parent)
    form.grab_set()
    form.configure(bg=config.master_bg)
    win_id = 2  # Unique Win_ID
    open_form_with_position(form, conn, cursor, win_id, "Income & Expenditure Category Maintenance")
    form.geometry(f"{sc(420)}x{sc(1000)}")

    # Treeview for categories
    tree = ttk.Treeview(form, show="tree")
    tree.place(x=sc(25), y=sc(65), width=sc(350), height=sc(825))
    tree.tag_configure("parent", background=COLORS["act_but_bg"], font=(config.ha_normal_bold))
    tree.tag_configure("child_odd", background="#F0F0F0", font=(config.ha_normal))
    tree.tag_configure("child_even", background="#FFFFFF", font=(config.ha_normal))
    tree.insert("", "end", text=" Income & Expense Categories ", tags=("parent",))
    scrollbar = tk.Scrollbar(form, orient="vertical", command=tree.yview)
    scrollbar.place(x=sc(375), y=sc(65), height=sc(825))
    tree.configure(yscrollcommand=scrollbar.set)
    
    parent_items = {}
    form.image_refs = []
    
    # Fetch icons
    up_img = ICON_CACHE.get("up_l")
    form.image_refs.append(up_img)
    down_img = ICON_CACHE.get("down_l")
    form.image_refs.append(down_img)
    
    # Expand all treeview rows
    is_expanded = True

    # Top Buttons
    fold_btn = tk.Button(form, text="Collapse All", font=config.ha_normal,
                        command=lambda: toggle_expand())
    fold_btn.place(x=sc(25), y=sc(15), width=sc(100), height=sc(35))
    
    tk.Label(form, text=f"{year}", font=config.ha_head12, bg=config.master_bg).place(x=sc(150), y=sc(15), width=sc(80), height=sc(35))
    
    up_btn = tk.Button(form, image=up_img, font=config.ha_normal, state="disabled",
                        command=lambda: move_up())
    up_btn.place(x=sc(250), y=sc(10), width=sc(40), height=sc(40))

    down_btn = tk.Button(form, image=down_img, font=config.ha_normal, state="disabled",
                        command=lambda: move_down())
    down_btn.place(x=sc(320), y=sc(10), width=sc(40), height=sc(40))
    
    # 4 Bottom Buttons
    new_btn = tk.Button(form, text="New", font=config.ha_normal,
                        command=lambda: [create_new_category_form(form, conn, cursor, year, tree), expand_all()])
    new_btn.place(x=sc(25), y=sc(905), width=sc(150), height=sc(35))
    
    update_btn = tk.Button(form, text="Edit", font=config.ha_normal, state="disabled",
                        command=lambda: [create_edit_category_form(form, conn, cursor, year, tree), expand_all()])
    update_btn.place(x=sc(225), y=sc(905), width=sc(150), height=sc(35))
    
    delete_btn = tk.Button(form, text="Delete", font=config.ha_normal, bg=COLORS["del_but_bg"], state="disabled",
                        command=lambda: [create_delete_category_form(form, conn, cursor, year, tree), expand_all()])
    delete_btn.place(x=sc(25), y=sc(955), width=sc(150), height=sc(35))
    
    close_btn = tk.Button(form, text="Close", font=config.ha_normal, bg=COLORS["exit_but_bg"],
                        command=lambda: [close_form_with_position(form, conn, cursor, win_id), tree_refresh_callback()])
    close_btn.place(x=sc(225), y=sc(955), width=sc(150), height=sc(35))
    
    def refresh_treeview():
        """Refresh Treeview with IE_Cata data for the given year."""
        for item in tree.get_children():
            tree.delete(item)
        title_iid = tree.insert("", "end", text=" Income & Expense Categories ", tags=("parent",))
        #logger.debug(f"Inserted title: iid={title_iid}")

        try:
            # Fetch parent categories, sorted by IE_Seq
            cursor.execute("""
                SELECT IE_PID, IE_Desc
                FROM IE_Cata
                WHERE IE_Year = ? AND IE_CID = 0 AND IE_PID > 0
                ORDER BY IE_Seq
            """, (year,))
            parents = cursor.fetchall()
            #logger.debug(f"Parent categories: {parents}")

            # Fetch child categories, sorted by IE_PID, IE_Seq
            cursor.execute("""
                SELECT IE_PID, IE_CID, IE_Desc
                FROM IE_Cata
                WHERE IE_Year = ? AND IE_CID > 0
                ORDER BY IE_PID, IE_Seq
            """, (year,))
            children = cursor.fetchall()
            #logger.debug(f"Child categories: {children}")

            # Check that we have records - else ask to copy previous year
            if not parents:
                form.focus_force()
                response = messagebox.askyesno(
                    "No Category Records Found",
                    f"No Category records exist for {year}. Do you want to create Categories based on last year's records?",
                    parent=form
                )
                if response:
                    rows_copied = copy_categories_from_previous_year(cursor, conn, year)
                    if rows_copied > 0:
                        # Fetch parent categories, sorted by IE_Seq
                        cursor.execute("""
                            SELECT IE_PID, IE_Desc
                            FROM IE_Cata
                            WHERE IE_Year = ? AND IE_CID = 0 AND IE_PID > 0
                            ORDER BY IE_Seq
                        """, (year,))
                        parents = cursor.fetchall()
                        #logger.debug(f"Parent categories: {parents}")

                        # Fetch child categories, sorted by IE_PID, IE_Seq
                        cursor.execute("""
                            SELECT IE_PID, IE_CID, IE_Desc
                            FROM IE_Cata
                            WHERE IE_Year = ? AND IE_CID > 0
                            ORDER BY IE_PID, IE_Seq
                        """, (year,))
                        children = cursor.fetchall()
                        #logger.debug(f"Child categories: {children}")
                    else:
                        messagebox.showinfo("No Data", f"No Categories found for {year - 1} to copy.", parent=form)
                        form.destroy()
                        return
                else:
                    form.destroy()
                    parent.focus_force()
                    return

            # Build Treeview
            parent_items.clear()
            row_count = 0
            for pid, desc in parents:
                parent_iid = tree.insert("", "end", text=desc, values=(pid, 0), tags=("parent",))
                parent_items[pid] = parent_iid
                #logger.debug(f"Inserted parent: IE_PID={pid}, text={desc}, iid={parent_iid}")
            
            for pid, cid, desc in children:
                if pid in parent_items:
                    tag = "child_odd" if row_count % 2 else "child_even"
                    child_iid = tree.insert(parent_items[pid], "end", text=desc, values=(pid, cid), tags=(tag,))
                    row_count += 1
                    #logger.debug(f"Inserted child: IE_PID={pid}, IE_CID={cid}, text={desc}, iid={child_iid}")
                else:
                    logger.warning(f"Child category IE_PID={pid}, IE_CID={cid} has no parent")
            
            # Log final Treeview contents
            tree_contents = []
            for item in tree.get_children():
                tree_contents.append(tree.item(item))
                for child in tree.get_children(item):
                    tree_contents.append(tree.item(child))
            #logger.debug(f"Final Treeview contents: {tree_contents}")

            # Force GUI update
            form.update_idletasks()

        except sqlite3.Error as e:
            logger.error(f"Error fetching IE_Cata: {e}")
            messagebox.showerror("Database Error", f"Failed to fetch categories: {e}", parent=form)
            return

    def on_tree_select(event):
        """Enable/disable buttons based on selection and position."""
        selected = tree.selection()
        if not selected:
            update_btn.config(state="disabled")
            delete_btn.config(state="disabled")
            up_btn.config(state="disabled")
            down_btn.config(state="disabled")
            return

        selected_item = tree.item(selected[0])
        values = selected_item["values"]
        selected_text = selected_item["text"]
        try:
            pid, cid = values
            if cid > 0:  # Child category (IE_CID > 0)
                cursor.execute("""
                    SELECT IE_CID, IE_Seq
                    FROM IE_Cata
                    WHERE IE_Year = ? AND IE_PID = ? AND IE_CID > 0
                    ORDER BY IE_Seq
                """, (year, pid))
                children = cursor.fetchall()
                child_seqs = [(c_cid, seq) for c_cid, seq in children]
                current_idx = next(i for i, (c_cid, _) in enumerate(child_seqs) if c_cid == cid)
                update_btn.config(state="normal")
                delete_btn.config(state="normal")
                up_btn.config(state="normal" if current_idx > 0 else "disabled")
                down_btn.config(state="normal" if current_idx < len(child_seqs) - 1 else "disabled")
            else:  # Parent category (IE_CID = 0)
                # Find IE_PID by matching text with parent_items
                #selected_text = selected_item["text"]
                #pid = next((p_pid for p_pid, p_iid in parent_items.items() if tree.item(p_iid)["text"] == selected_text), None)
                if pid is None:
                    logger.warning(f"No IE_PID found for selected parent: {selected_text}")
                    update_btn.config(state="disabled")
                    delete_btn.config(state="disabled")
                    up_btn.config(state="disabled")
                    down_btn.config(state="disabled")
                    return
                cursor.execute("""
                    SELECT IE_PID, IE_Seq
                    FROM IE_Cata
                    WHERE IE_Year = ? AND IE_CID = 0 AND IE_PID > 0
                    ORDER BY IE_Seq
                """, (year,))
                parents = cursor.fetchall()
                parent_seqs = [(p_pid, seq) for p_pid, seq in parents]
                #logger.debug(f"pid={pid}, parent_seqs={parent_seqs}")
                current_idx = next(i for i, (p_pid, _) in enumerate(parent_seqs) if p_pid == pid)
                update_btn.config(state="normal" if pid != 1 else "disabled")  # Disable for Income (IE_PID=1)
                delete_btn.config(state="normal" if pid != 1 else "disabled")
                up_btn.config(state="normal" if current_idx > 1 else "disabled")
                down_btn.config(state="normal" if current_idx < len(parent_seqs) - 1 and current_idx > 0 else "disabled")
        except sqlite3.Error as e:
            logger.error(f"Error checking position: {e}")
            messagebox.showerror("Database Error", f"Failed to check position: {e}", parent=form)

    def expand_all():
        """Expand all parent nodes."""
        for pid, iid in parent_items.items():
            tree.item(iid, open=True)
            #logger.debug(f"Expanded parent: IE_PID={pid}, iid={iid}")

    def collapse_all():
        """Collapse all parent nodes."""
        for pid, iid in parent_items.items():
            tree.item(iid, open=False)
            #logger.debug(f"Collapsed parent: IE_PID={pid}, iid={iid}")
        
    def toggle_expand():
        """Toggle Treeview expansion state."""
        nonlocal is_expanded
        if is_expanded:
            collapse_all()
            is_expanded = False
            fold_btn.config(text="Expand All")
        else:
            expand_all()
            is_expanded = True
            fold_btn.config(text="Collapse All")

    def move_up():
        """Move the selected category up by swapping IE_Seq with the previous item."""
        selected = tree.selection()
        if not selected:
            return
        selected_item = tree.item(selected[0])
        values = selected_item["values"]
        
        try:
            if values:  # Child category
                pid, cid = values
                cursor.execute("""
                    SELECT IE_CID, IE_Seq
                    FROM IE_Cata
                    WHERE IE_Year = ? AND IE_PID = ? AND IE_CID > 0
                    ORDER BY IE_Seq
                """, (year, pid))
                children = cursor.fetchall()
                child_seqs = [(c_cid, seq) for c_cid, seq in children]
                current_idx = next(i for i, (c_cid, _) in enumerate(child_seqs) if c_cid == cid)
                if current_idx == 0:
                    #logger.debug(f"Child IE_PID={pid}, IE_CID={cid} is already at the top")
                    return
                prev_cid, prev_seq = child_seqs[current_idx - 1]
                current_seq = child_seqs[current_idx][1]
                # Swap IE_Seq values
                with conn:
                    cursor.execute("""
                        UPDATE IE_Cata
                        SET IE_Seq = ?
                        WHERE IE_Year = ? AND IE_PID = ? AND IE_CID = ?
                    """, (prev_seq, year, pid, cid))
                    cursor.execute("""
                        UPDATE IE_Cata
                        SET IE_Seq = ?
                        WHERE IE_Year = ? AND IE_PID = ? AND IE_CID = ?
                    """, (current_seq, year, pid, prev_cid))
                #logger.debug(f"Swapped child IE_PID={pid}, IE_CID={cid} (IE_Seq={current_seq}) with IE_CID={prev_cid} (IE_Seq={prev_seq})")
            else:  # Parent category
                selected_text = selected_item["text"]
                pid = next((p_pid for p_pid, p_iid in parent_items.items() if tree.item(p_iid)["text"] == selected_text), None)
                if pid is None:
                    logger.warning(f"No IE_PID found for selected parent: {selected_text}")
                    return
                cursor.execute("""
                    SELECT IE_PID, IE_Seq
                    FROM IE_Cata
                    WHERE IE_Year = ? AND IE_CID = 0 AND IE_PID > 0
                    ORDER BY IE_Seq
                """, (year,))
                parents = cursor.fetchall()
                parent_seqs = [(p_pid, seq) for p_pid, seq in parents]
                current_idx = next(i for i, (p_pid, _) in enumerate(parent_seqs) if p_pid == pid)
                if current_idx == 0:
                    #logger.debug(f"Parent IE_PID={pid} is already at the top")
                    return
                prev_pid, prev_seq = parent_seqs[current_idx - 1]
                current_seq = parent_seqs[current_idx][1]
                # Swap IE_Seq values
                with conn:
                    cursor.execute("""
                        UPDATE IE_Cata
                        SET IE_Seq = ?
                        WHERE IE_Year = ? AND IE_PID = ? AND IE_CID = 0
                    """, (prev_seq, year, pid))
                    cursor.execute("""
                        UPDATE IE_Cata
                        SET IE_Seq = ?
                        WHERE IE_Year = ? AND IE_PID = ? AND IE_CID = 0
                    """, (current_seq, year, prev_pid))
                #logger.debug(f"Swapped parent IE_PID={pid} (IE_Seq={current_seq}) with IE_PID={prev_pid} (IE_Seq={prev_seq})")
            form.refresh_treeview()
            form.expand_all()
            # Reselect the moved item
            for item in tree.get_children():
                if values and tree.item(item)["values"] == values:
                    tree.selection_set(item)
                    break
                for child in tree.get_children(item):
                    if values and tree.item(child)["values"] == values:
                        tree.selection_set(child)
                        break
                if not values and tree.item(item)["text"] == selected_text:
                    tree.selection_set(item)
                    break
        except sqlite3.Error as e:
            logger.error(f"Error moving category up: {e}")
            messagebox.showerror("Database Error", f"Failed to move category: {e}", parent=form)

    def move_down():
        """Move the selected category down by swapping IE_Seq with the next item."""
        selected = tree.selection()
        if not selected:
            return
        selected_item = tree.item(selected[0])
        values = selected_item["values"]
        
        try:
            if values:  # Child category
                pid, cid = values
                cursor.execute("""
                    SELECT IE_CID, IE_Seq
                    FROM IE_Cata
                    WHERE IE_Year = ? AND IE_PID = ? AND IE_CID > 0
                    ORDER BY IE_Seq
                """, (year, pid))
                children = cursor.fetchall()
                child_seqs = [(c_cid, seq) for c_cid, seq in children]
                current_idx = next(i for i, (c_cid, _) in enumerate(child_seqs) if c_cid == cid)
                if current_idx == len(child_seqs) - 1:
                    #logger.debug(f"Child IE_PID={pid}, IE_CID={cid} is already at the bottom")
                    return
                next_cid, next_seq = child_seqs[current_idx + 1]
                current_seq = child_seqs[current_idx][1]
                # Swap IE_Seq values
                with conn:
                    cursor.execute("""
                        UPDATE IE_Cata
                        SET IE_Seq = ?
                        WHERE IE_Year = ? AND IE_PID = ? AND IE_CID = ?
                    """, (next_seq, year, pid, cid))
                    cursor.execute("""
                        UPDATE IE_Cata
                        SET IE_Seq = ?
                        WHERE IE_Year = ? AND IE_PID = ? AND IE_CID = ?
                    """, (current_seq, year, pid, next_cid))
                #logger.debug(f"Swapped child IE_PID={pid}, IE_CID={cid} (IE_Seq={current_seq}) with IE_CID={next_cid} (IE_Seq={next_seq})")
            else:  # Parent category
                selected_text = selected_item["text"]
                pid = next((p_pid for p_pid, p_iid in parent_items.items() if tree.item(p_iid)["text"] == selected_text), None)
                if pid is None:
                    logger.warning(f"No IE_PID found for selected parent: {selected_text}")
                    return
                cursor.execute("""
                    SELECT IE_PID, IE_Seq
                    FROM IE_Cata
                    WHERE IE_Year = ? AND IE_CID = 0 AND IE_PID > 0
                    ORDER BY IE_Seq
                """, (year,))
                parents = cursor.fetchall()
                parent_seqs = [(p_pid, seq) for p_pid, seq in parents]
                current_idx = next(i for i, (p_pid, _) in enumerate(parent_seqs) if p_pid == pid)
                if current_idx == len(parent_seqs) - 1:
                    #logger.debug(f"Parent IE_PID={pid} is already at the bottom")
                    return
                next_pid, next_seq = parent_seqs[current_idx + 1]
                current_seq = parent_seqs[current_idx][1]
                # Swap IE_Seq values
                with conn:
                    cursor.execute("""
                        UPDATE IE_Cata
                        SET IE_Seq = ?
                        WHERE IE_Year = ? AND IE_PID = ? AND IE_CID = 0
                    """, (next_seq, year, pid))
                    cursor.execute("""
                        UPDATE IE_Cata
                        SET IE_Seq = ?
                        WHERE IE_Year = ? AND IE_PID = ? AND IE_CID = 0
                    """, (current_seq, year, next_pid))
                #logger.debug(f"Swapped parent IE_PID={pid} (IE_Seq={current_seq}) with IE_PID={next_pid} (IE_Seq={next_seq})")
            form.refresh_treeview()
            form.expand_all()
            # Reselect the moved item
            for item in tree.get_children():
                if values and tree.item(item)["values"] == values:
                    tree.selection_set(item)
                    break
                for child in tree.get_children(item):
                    if values and tree.item(child)["values"] == values:
                        tree.selection_set(child)
                        break
                if not values and tree.item(item)["text"] == selected_text:
                    tree.selection_set(item)
                    break
        except sqlite3.Error as e:
            logger.error(f"Error moving category down: {e}")
            messagebox.showerror("Database Error", f"Failed to move category: {e}", parent=form)

    tree.bind("<<TreeviewSelect>>", on_tree_select)
    # Bind refresh_treeview and expand_all to form for access in subforms
    form.refresh_treeview = refresh_treeview    
    form.expand_all = expand_all  
    refresh_treeview()
    expand_all()
    form.wait_window(form)

# Subforms used by create_category_maint_form
def create_new_category_form(parent, conn, cursor, year, tree):
    """Create a subform for adding a new parent or child category."""
    subform = tk.Toplevel(parent)
    subform.title("New Category Record")
    subform.transient(parent)
    subform.grab_set()
    subform.configure(bg=config.master_bg)
    win_id = 29
    open_form_with_position(subform, conn, cursor, win_id, "New Category Record")
    subform.geometry(f"{sc(400)}x{sc(400)}")

    is_parent = tk.BooleanVar(value=False)
    
    tk.Label(subform, text=f"Year: {year}", font=config.ha_vlarge, bg=config.master_bg).place(x=sc(130), y=sc(25), width=sc(130), height=sc(40))
    
    new_parent_btn = tk.Button(subform, text="New Parent Category", font=config.ha_normal,
                            command=lambda: show_parent_fields())
    new_parent_btn.place(x=sc(25), y=sc(300), width=sc(150), height=sc(35))
    
    new_child_btn = tk.Button(subform, text="New Child Category", font=config.ha_normal,
                            command=lambda: show_child_fields())
    new_child_btn.place(x=sc(225), y=sc(300), width=sc(150), height=sc(35))
    
    cancel_btn = tk.Button(subform, text="Cancel", font=config.ha_normal, bg=COLORS["exit_but_bg"],
                        command=lambda: close_form_with_position(subform, conn, cursor, win_id))
    cancel_btn.place(x=sc(225), y=sc(350), width=sc(150), height=sc(35))
    
    save_btn = tk.Button(subform, text="Save", font=config.ha_normal, state="disabled",
                        command=lambda: save_new_category())
    save_btn.place(x=sc(25), y=sc(350), width=sc(150), height=sc(35))
    
    parent_label = tk.Label(subform, text="Category Group Heading:", font=config.ha_normal, bg=config.master_bg, state="disabled")
    parent_label.place(x=sc(50), y=sc(75), width=sc(300), height=sc(35))
    
    parent_input = tk.Entry(subform, font=config.ha_normal, bg="white", state="disabled")
    parent_input.place(x=sc(50), y=sc(155), width=sc(300), height=sc(30))
    
    parent_combo = ttk.Combobox(subform, font=config.ha_normal, background="white", state="disabled")
    parent_combo.place(x=sc(50), y=sc(115), width=sc(300), height=sc(30))
    
    child_label = tk.Label(subform, text="Category Title:", font=config.ha_normal, bg=config.master_bg, state="disabled")
    child_label.place(x=sc(50), y=sc(205), width=sc(300), height=sc(35))
    
    child_input = tk.Entry(subform, font=config.ha_normal, bg="white", state="disabled")
    child_input.place(x=sc(50), y=sc(245), width=sc(300), height=sc(30))

    def fetch_parent_categories():
        """Fetch parent category names for the combo box."""
        try:
            cursor.execute("SELECT IE_Desc FROM IE_Cata WHERE IE_Year = ? AND IE_CID = 0 ORDER BY IE_Seq", (year,))
            return [row[0] for row in cursor.fetchall()]
        except sqlite3.Error as e:
            logger.error(f"Error fetching parent categories: {e}")
            return []

    def show_parent_fields():
        """Show fields for creating a parent category."""
        is_parent.set(True)
        parent_label.config(state="normal")
        parent_input.config(state="normal")
        parent_combo.config(state="disabled")
        child_label.config(state="disabled")
        child_input.config(state="disabled")
        new_parent_btn.config(state="disabled")
        new_child_btn.config(state="disabled")
        cancel_btn.config(state="normal")
        save_btn.config(state="normal")

    def show_child_fields():
        """Show fields for creating a child category."""
        is_parent.set(False)
        parent_label.config(state="normal")
        parent_input.config(state="disabled")
        parent_combo.config(state="readonly")
        parent_combo["values"] = fetch_parent_categories()
        child_label.config(state="normal")
        child_input.config(state="normal")
        new_parent_btn.config(state="disabled")
        new_child_btn.config(state="disabled")
        cancel_btn.config(state="normal")
        save_btn.config(state="normal")

    def save_new_category():
        """Save new parent or child category to IE_Cata."""
        if is_parent.get():
            desc = parent_input.get().strip()
            if len(desc) < 3:
                messagebox.showerror("Error", "The description must be at least 3 characters", parent=subform)
                return
            try:
                cursor.execute("SELECT MAX(IE_PID) FROM IE_Cata WHERE IE_Year = ?", (year,))
                max_pid = cursor.fetchone()[0] or 0
                pid = max_pid + 1
                cursor.execute("INSERT INTO IE_Cata (IE_PID, IE_CID, IE_Year, IE_Desc, IE_Cat_ID, IE_Seq) VALUES (?, ?, ?, ?, ?, ?)",
                            (pid, 0, year, desc, 0, pid)) # set IE_Seq to place at bottom of list
                conn.commit()
                #logger.debug(f"Inserted parent category: IE_PID={pid}, IE_Desc={desc}")
                tree.yview_moveto(1)
            except sqlite3.Error as e:
                logger.error(f"Error inserting parent category: {e}")
                messagebox.showerror("Database Error", f"Failed to save category: {e}", parent=subform)
                return
        else:
            parent_desc = parent_combo.get().strip()
            child_desc = child_input.get().strip()
            if len(child_desc) < 3:
                messagebox.showerror("Error", "The description must be at least 3 characters", parent=subform)
                return
            try:
                cursor.execute("SELECT IE_PID FROM IE_Cata WHERE IE_Year = ? AND IE_CID = 0 AND IE_Desc = ?",
                            (year, parent_desc))
                pid = cursor.fetchone()
                if not pid:
                    messagebox.showerror("Error", "Selected parent category not found", parent=subform)
                    return
                pid = pid[0]
                cursor.execute("SELECT MAX(IE_CID) FROM IE_Cata WHERE IE_Year = ? AND IE_PID = ?", (year, pid))
                max_cid = cursor.fetchone()[0] or 0
                if max_cid >= 19:
                    messagebox.showerror("Error", "A maximum of 19 subcategories are allowed per category", parent=subform)
                    return
                cid = max_cid + 1
                cursor.execute("INSERT INTO IE_Cata (IE_PID, IE_CID, IE_Year, IE_Desc, IE_Cat_ID, IE_Seq) VALUES (?, ?, ?, ?, ?, ?)",
                            (pid, cid, year, child_desc, 0, cid))  # set IE_Seq to place at bottom of list
                conn.commit()
                #logger.debug(f"Inserted child category: IE_PID={pid}, IE_CID={cid}, IE_Desc={child_desc}")
            except sqlite3.Error as e:
                logger.error(f"Error inserting child category: {e}")
                messagebox.showerror("Database Error", f"Failed to save category: {e}", parent=subform)
                return
        close_form_with_position(subform, conn, cursor, win_id)
        parent.refresh_treeview()

    subform.wait_window(subform)

def create_edit_category_form(parent, conn, cursor, year, tree):
    """Create a subform for editing an existing category."""
    selected = tree.selection()
    if not selected:
        return
    selected_item = tree.item(selected[0])
    values = selected_item["values"]
    
    if values:  # Child category
        pid, cid = tree.item(selected[0])["values"] or (0, 0)

        try:
            cursor.execute("SELECT IE_Desc FROM IE_Cata WHERE IE_Year = ? AND IE_PID = ? AND IE_CID = 0", (year, pid))
            parent_desc = cursor.fetchone()[0]
            child_desc = ""
            if cid > 0:
                cursor.execute("SELECT IE_Desc FROM IE_Cata WHERE IE_Year = ? AND IE_PID = ? AND IE_CID = ?",
                            (year, pid, cid))
                child_desc = cursor.fetchone()[0]
                
        except sqlite3.Error as e:
            logger.error(f"Error fetching category: {e}")
            messagebox.showerror("Database Error", f"Failed to fetch category: {e}", parent=parent)
            return

    else:  # Parent category
        parent_desc = selected_item["text"]
        #pid = next((p_pid for p_pid, p_iid in parent_items.items() if tree.item(p_iid)["text"] == selected_text), None)
        if pid is None:
            logger.warning(f"No IE_PID found for selected parent: {parent_desc}")
            return
        if pid == 1 and cid == 0:
            messagebox.showinfo("Not Allowed", "The 'Income' category heading is reserved and may not be changed", parent=parent)
            return
            

    subform = tk.Toplevel(parent)
    subform.title("Edit Category Record")
    subform.transient(parent)
    subform.grab_set()
    subform.configure(bg=config.master_bg)
    win_id = 29
    open_form_with_position(subform, conn, cursor, win_id, "Edit Category Record")
    subform.geometry(f"{sc(400)}x{sc(400)}")

    tk.Label(subform, text=f"Year: {year}", font=config.ha_vlarge, bg=config.master_bg, anchor="center").place(x=sc(20), y=sc(25), width=sc(360), height=sc(40))
    
    tk.Label(subform, text="Category Group Heading:", font=config.ha_normal, bg=config.master_bg, anchor="center").place(x=sc(20), y=sc(75), width=sc(360), height=sc(35))
    
    if cid == 0:
        input_field = tk.Entry(subform, font=config.ha_normal)
        input_field.insert(0, parent_desc)
        input_field.place(x=sc(50), y=sc(115), width=sc(300), height=sc(35))
    else:
        tk.Label(subform, text=parent_desc, font=config.ha_normal, bg=config.master_bg, anchor="center").place(x=sc(20), y=sc(115), width=sc(360), height=sc(35))
        tk.Label(subform, text="Category Title:", font=config.ha_normal, bg=config.master_bg, anchor="w").place(x=sc(50), y=sc(165), width=sc(200), height=sc(35))
        input_field = tk.Entry(subform, font=config.ha_normal)
        input_field.insert(0, child_desc)
        input_field.place(x=sc(50), y=sc(205), width=sc(300), height=sc(30))

    cancel_btn = tk.Button(subform, text="Cancel", font=config.ha_button, bg=COLORS["exit_but_bg"],
                        command=lambda: close_form_with_position(subform, conn, cursor, win_id))
    cancel_btn.place(x=sc(225), y=sc(350), width=sc(150), height=sc(35))
    
    save_btn = tk.Button(subform, text="Save", font=config.ha_button, command=lambda: save_edit_category())
    save_btn.place(x=sc(25), y=sc(350), width=sc(150), height=sc(35))

    def save_edit_category():
        """Save updated category description."""
        desc = input_field.get().strip().replace("'", "''")
        if len(desc) < 3:
            messagebox.showerror("Error", "The description must be at least 3 characters", parent=subform)
            return
        try:
            cursor.execute("UPDATE IE_Cata SET IE_Desc = ? WHERE IE_Year = ? AND IE_PID = ? AND IE_CID = ?",
                        (desc, year, pid, cid))
            conn.commit()
            #logger.debug(f"Updated category: IE_PID={pid}, IE_CID={cid}, IE_Desc={desc}")
            close_form_with_position(subform, conn, cursor, win_id)
            parent.refresh_treeview()
        except sqlite3.Error as e:
            logger.error(f"Error updating category: {e}")
            messagebox.showerror("Database Error", f"Failed to update category: {e}", parent=subform)

    subform.wait_window(subform)

def create_delete_category_form(parent, conn, cursor, year, tree):
    """Create a subform for deleting a category."""
    selected = tree.selection()
    if not selected or not tree.item(selected[0])["values"]:
        return
    pid, cid = tree.item(selected[0])["values"] or (0, 0)
    if pid == 1 and cid == 0:
        messagebox.showinfo("Not Allowed", "The 'Income' category heading is reserved and may not be changed", parent=parent)
        return

    try:
        cursor.execute("SELECT IE_Desc FROM IE_Cata WHERE IE_Year = ? AND IE_PID = ? AND IE_CID = 0", (year, pid))
        parent_desc = cursor.fetchone()[0]
        child_desc = ""
        if cid > 0:
            cursor.execute("SELECT IE_Desc FROM IE_Cata WHERE IE_Year = ? AND IE_PID = ? AND IE_CID = ?",
                        (year, pid, cid))
            child_desc = cursor.fetchone()[0]
        
        # Check for transactions using this category
        if cid == 0:
            cursor.execute("SELECT COUNT(*) FROM Trans WHERE Tr_Exp_ID = ? LIMIT 10", (pid,))
        else:
            cursor.execute("SELECT COUNT(*) FROM Trans WHERE Tr_Exp_ID = ? AND Tr_ExpSub_ID = ? LIMIT 10", (pid, cid))
        trans_count = cursor.fetchone()[0]
        if trans_count > 0:
            messagebox.showerror("Error", "Cannot delete this category while it is used by transactions", parent=parent)
            return
    except sqlite3.Error as e:
        logger.error(f"Error checking category: {e}")
        messagebox.showerror("Database Error", f"Failed to check category: {e}", parent=parent)
        return

    subform = tk.Toplevel(parent)
    subform.title("Delete Category Record")
    subform.transient(parent)
    subform.grab_set()
    subform.configure(bg=config.master_bg)
    win_id = 29
    open_form_with_position(subform, conn, cursor, win_id, "Delete Category Record")
    subform.geometry(f"{sc(400)}x{sc(400)}")

    tk.Label(subform, text=f"Year: {year}", font=config.ha_vlarge, bg=config.master_bg, anchor="center").place(x=sc(20), y=sc(25), width=sc(360), height=sc(40))
    tk.Label(subform, text="Category Group Heading:", font=config.ha_normal, bg=config.master_bg, anchor="center").place(x=sc(20), y=sc(75), width=sc(360), height=sc(35))
    tk.Label(subform, text=parent_desc, font=config.ha_normal, bg=config.master_bg, anchor="center").place(x=sc(20), y=sc(115), width=sc(360), height=sc(35))
    
    if cid > 0:
        tk.Label(subform, text="Category Title to Delete:", font=config.ha_normal, bg=config.master_bg, anchor="center").place(x=sc(25), y=sc(165), width=sc(360), height=sc(35))
        tk.Label(subform, text=child_desc, font=config.ha_normal, bg=config.master_bg, anchor="center").place(x=sc(25), y=sc(205), width=sc(360), height=sc(35))

    cancel_btn = tk.Button(subform, text="Cancel", font=config.ha_button, bg=COLORS["exit_but_bg"],
                        command=lambda: close_form_with_position(subform, conn, cursor, win_id))
    cancel_btn.place(x=sc(225), y=sc(350), width=sc(150), height=sc(35))
    
    delete_btn = tk.Button(subform, text="Confirm Delete", font=config.ha_button, bg="red",
                        command=lambda: confirm_delete())
    delete_btn.place(x=sc(25), y=sc(350), width=sc(150), height=sc(35))

    def confirm_delete():
        """Delete the category from IE_Cata."""
        try:
            cursor.execute("DELETE FROM IE_Cata WHERE IE_Year = ? AND IE_PID = ? AND IE_CID = ?", (year, pid, cid))
            conn.commit()
            #logger.debug(f"Deleted category: IE_PID={pid}, IE_CID={cid}")
            close_form_with_position(subform, conn, cursor, win_id)
            parent.refresh_treeview()
        except sqlite3.Error as e:
            logger.error(f"Error deleting category: {e}")
            messagebox.showerror("Database Error", f"Failed to delete category: {e}", parent=subform)

    subform.wait_window(subform)

def create_colour_scheme_maint_form(parent, conn, cursor, colours_refresh_callback=None):              # Win_ID = 21
    form = tk.Toplevel(parent, bg=config.master_bg)
    win_id = 21
    open_form_with_position(form, conn, cursor, win_id, "Colours Maintenance")
    form.geometry(f"{sc(1000)}x{sc(900)}")
    form.resizable(False, False)
    form.attributes("-topmost", True)
    form.configure(bg=config.master_bg)
    form.grab_set()

    #logger.debug(f"Colours Maintenance form size: width={sc(1000)}, height={sc(900)}")
    #logger.debug(f"Tkinter version: {tk.TkVersion}")
    #logger.debug(f"Current theme: {ttk.Style().theme_use()}")

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
    #logger.debug(f"colors dict: {colors}")
    #logger.debug("COLORS loaded into colors dict")

    # Create UI elements using grid
    entries = []
    bg_buttons = []
    text_buttons = []

    # Column headers
    tk.Label(form, text="Background", font=(config.ha_head11), bg=config.master_bg).grid(row=0, column=2, padx=sc(20), pady=10, sticky="ew")
    tk.Label(form, text="Text", font=(config.ha_head11), bg=config.master_bg).grid(row=0, column=3, padx=sc(20), pady=10, sticky="ew")
    #logger.debug("Created column headers")

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
        #logger.debug(f"Created entry for row {i}: row={i+1}, column=1")

        # Background button
        if i <= 6 or i >= 13 or i == 8 or i == 11:  # Rows 0-6, 13-24, weekday (8), weekend (11)
            bg_btn = tk.Button(form, text="Select Colour", font=(config.ha_button), width=17, command=lambda idx=i: select_color(idx, 0))
            bg_btn.grid(row=i+1, column=2, padx=sc(20), pady=2, sticky="ew")
            bg_buttons.append(bg_btn)
            #logger.debug(f"Created background button for row {i}")
        else:
            bg_buttons.append(None)

        # Text button
        if i in (7, 8, 9) or i >= 13:  # Rows 7, 8, 9 and 13-24
            text_btn = tk.Button(form, text="Select Colour", font=(config.ha_button), width=17, command=lambda idx=i: select_color(idx, 1))
            text_btn.grid(row=i+1, column=3, padx=sc(20), pady=2, sticky="ew")
            text_buttons.append(text_btn)
            #logger.debug(f"Created text button for row {i}")
        else:
            text_buttons.append(None)

    # Bottom buttons
    save_btn = tk.Button(form, text="Save", font=(config.ha_button), width=15, command=lambda: save_colors())
    save_btn.place(x=sc(100), y=sc(830), width=sc(150), height=sc(35))
    ### Under construction ###
    apply_btn = tk.Button(form, text="Apply", font=(config.ha_button), width=15, command=lambda: apply_colors(colours_refresh_callback))
    apply_btn.place(x=sc(300), y=sc(830), width=sc(150), height=sc(35))
    
    reset_btn = tk.Button(form, text="Reset All to Defaults", font=(config.ha_button), width=30, command=lambda: reset_colors())
    reset_btn.place(x=sc(500), y=sc(830), width=sc(200), height=sc(35))
    close_btn = tk.Button(form, text="Close", font=(config.ha_button), bg=COLORS["exit_but_bg"], width=15, command=lambda: close_form_with_position(form, conn, cursor, win_id))
    close_btn.place(x=sc(750), y=sc(830), width=sc(150), height=sc(35))
    #logger.debug("Created bottom buttons")

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
        #logger.debug(f"Color selected: lup_type_id={lup_type_id}, lup_seq={lup_seq}, type={'background' if bk_txt == 0 else 'text'}, color={color[1]}")
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
        #logger.debug("Colors reset to defaults (locally - not saved)")

    def save_colors():
        # Saves the colors dict to the database Lookups, these will then load on next restart of app
        try:
            for (lup_type_id, lup_seq), value in colors.items():
                #logger.debug(f"lup_type_id= {lup_type_id}, lup_seq= {lup_seq}, value= {value}")
                update_lookup_color(cursor, conn, lup_seq, lup_type_id, value)
            conn.commit()
            messagebox.showinfo("Success", "Colors saved successfully", parent=form)
            #logger.debug("Colors saved to Lookups table")
        except Exception as e:
            logger.error(f"Failed to save colors: {e}")
            messagebox.showerror("Error", f"Failed to save colors: {e}", parent=form)
            
    def apply_colors(colours_refresh_callback):
        """Saves the colors dict back to the COLORS dict, applying to the current app without saving to DB."""
        for lup_type_id, lup_seq in colors:
            # Find the COLORS key that maps to (lup_type_id, lup_seq) in color_map
            for color_key, color_tuple in color_map.items():
                if color_tuple == (lup_type_id, lup_seq):
                    COLORS[color_key] = colors[(lup_type_id, lup_seq)]
                    #logger.debug(f"Updated COLORS[{color_key}] = {colors[(lup_type_id, lup_seq)]}")
                    break
            else:
                logger.warning(f"No color_map entry for Lup_LupT_ID={lup_type_id}, Lup_Seq={lup_seq}")
        
        # Trigger the home form's refresh
        if colours_refresh_callback:
            colours_refresh_callback()
            #logger.debug("Refresh callback triggered")        
        
    def update_entries():       # uses COLORS dictionary to set entries[] background and foreground colours
        lup_type_id = 2
        for i in range(25):     # loads all background colours to entry fields
            if i <= 6:
                entries[i].config(readonlybackground=colors[lup_type_id, i+1])
                temp=colors[(lup_type_id, i+1)]
                #logger.debug(f"update_entries bg - entries[{i}]: colors: {temp}")
            elif i in (7, 8, 9):
                entries[i].config(readonlybackground=colors[lup_type_id, 8])
                temp=colors[(lup_type_id, 8)]
                #logger.debug(f"update_entries bg - entries[{i}]: colors:{temp}")
            elif i in (10, 11, 12):
                entries[i].config(readonlybackground=colors[lup_type_id, 9])
                temp=colors[(lup_type_id, 9)]
                #logger.debug(f"update_entries bg - entries[{i}]: colors:{temp}")
            elif i >= 13:
                entries[i].config(readonlybackground=colors[lup_type_id, i-3])
                temp=colors[(lup_type_id, i-3)]
                #logger.debug(f"update_entries bg - entries[{i}]: colors:{temp}")
        lup_type_id = 3
        for i in range(17):     # load all foreground colours into entry fields
            #if i in (0, 1):
                #logger.debug(f"first 2 text colors [{i}]: skipped")
            if i in (2, 3, 4):
                entries[i+5].config(fg=colors[lup_type_id, i+1])
                entries[i+8].config(fg=colors[lup_type_id, i+1])
                temp=colors[(lup_type_id, i+1)]
                #logger.debug(f"update_entries fg - entries[{i+5}] + [{i+8}]: colors:{temp}")
            if i >= 5:
                entries[i+8].config(fg=colors[(lup_type_id, i+1)])
                temp=colors[(lup_type_id, i+1)]
                #logger.debug(f"update_entries fg - entries[{i+8}]: colors:{temp}")
        entries[i].update()
        #logger.debug("Entries updated with current colors")
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

def create_form_positions_maint_form(parent, conn, cursor, focus_refresh_callback=None):             # Win_ID = 8
    form = tk.Toplevel(parent)
    win_id = 8
    open_form_with_position(form, conn, cursor, win_id, "Manage Form Window Positions")
    form.geometry(f"{sc(800)}x{sc(600)}")
    form.config(bg=config.master_bg)
    
    scroll_row_var = tk.StringVar(value=str(config.scroll_row))
    save_flag = 0   # state of edit/save button
    scroll_rows = []
    for i in range(1, 36):
        scroll_rows.append(i)
            
    # Treeview styling
    style = ttk.Style()
    style.configure("Treeview", font=(config.ha_normal))

    # Treeview
    tree = ttk.Treeview(form, columns=("ID", "Name", "Left", "Top"), show="headings", height=18)
    tree.place(x=sc(20), y=sc(20), width=sc(760))
    tree.heading("ID", text="Win ID")
    tree.heading("Name", text="Form Name")
    tree.heading("Left", text="Left")
    tree.heading("Top", text="Top")
    tree.column("ID", width=50)
    tree.column("Name", width=450)
    tree.column("Left", width=100)
    tree.column("Top", width=100)

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

    tk.Label(form, text="Top:", font=(config.ha_normal), bg=config.master_bg, anchor="e", width=15).place(x=sc(220), y=sc(460))
    top_entry = tk.Entry(form, width=10, font=(config.ha_normal))
    top_entry.place(x=sc(350), y=sc(460))
    
    focus_row_frame = tk.LabelFrame(form, text=" Focus Row of Main Form ", font=(config.ha_normal), bg=config.master_bg)
    focus_row_frame.place(x=sc(520), y=sc(420), width= sc(260), height = sc(100))
    fr1 = tk.Label(focus_row_frame, text="(adjusts scroll markers)", font=(config.ha_note), bg=config.master_bg)
    fr1.place(x=sc(10), y=sc(1), width= sc(150), height = sc(20))
    scroll_combo = ttk.Combobox(focus_row_frame, textvariable=scroll_row_var, values=[str(y) for y in scroll_rows], font=(config.ha_button), width=6, state="readonly")
    scroll_combo.place(x=sc(100), y=sc(30))
    fr2 = tk.Label(focus_row_frame, text="(Range 1-35, default 12)", font=(config.ha_note), bg=config.master_bg)
    fr2.place(x=sc(30), y=sc(55), width= sc(200), height = sc(20))
    
    scroll_combo.set(str(config.scroll_row))
    
    def on_scroll_change(event):
        new_scroll = str(scroll_row_var.get())
        cursor.execute("UPDATE settings SET value = ? WHERE key = ?", (new_scroll, 'SCROLL_ROW'))
        conn.commit()
        config.scroll_row = int(new_scroll)
        if focus_refresh_callback:
            focus_refresh_callback()
        
    scroll_combo.bind("<<ComboboxSelected>>", on_scroll_change)
    
    def amend_position():
        nonlocal save_flag
        if save_flag == 0:
            selection = tree.selection()
            if selection:
                values = tree.item(selection[0], "values")
                name_entry.delete(0, tk.END)
                name_entry.insert(0, values[1])
                left_entry.delete(0, tk.END)
                left_entry.insert(0, values[2])
                top_entry.delete(0, tk.END)
                top_entry.insert(0, values[3])
                amend_save_btn.config(text="Save Changes")
                save_flag = 1
        elif save_flag == 1:
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
        nonlocal save_flag
        name_entry.delete(0, tk.END)
        left_entry.delete(0, tk.END)
        top_entry.delete(0, tk.END)
        tree.selection_remove(tree.selection())
        amend_save_btn.config(text="Amend Selected Row")
        save_flag = 0

    # Buttons
    amend_save_btn = tk.Button(form, text="Edit Selected Row", font=(config.ha_button), width=25, command=amend_position)
    amend_save_btn.place(x=sc(150), y=sc(500))
    tk.Button(form, text="Close", font=(config.ha_button), bg=COLORS["exit_but_bg"], width=15, 
            command=lambda: close_form_with_position(form, conn, cursor, win_id)).place(x=sc(580), y=sc(540))

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
    
    
    
    