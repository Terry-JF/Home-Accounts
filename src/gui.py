# Creates the New/Edit form and Toolbox form

import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import Calendar
from datetime import datetime
import logging
from config import COLORS
from db import (insert_transaction, update_transaction, delete_transaction, fetch_notes, fetch_categories, fetch_subcategories,
                fetch_account_full_names, fetch_statement_balances, update_account_month_transaction_total)
from ui_utils import (refresh_grid, open_form_with_position, close_form_with_position, sc)
from gui_maint import (create_account_maint_form, create_category_maint_form, create_export_transactions_form,
                    create_colour_scheme_maint_form, create_account_years_maint_form, create_annual_budget_maint_form, 
                    create_transaction_options_maint_form, create_form_positions_maint_form, create_ff_mappings_maint_form)
from gc_utils import (create_gocardless_maint_form, resource_path)
from gui_maint_rules import (create_rules_form)
from m_reg_trans import (create_regular_transactions_maint_form)
import config

# Set up logging
logger = logging.getLogger('HA.gui')

# Handle DPI scaling for high-resolution displays
try:
    from ctypes import windll
    windll.shcore.SetProcessDpiAwareness(1)
except Exception as e:
    logger.warning(f"Failed to set DPI awareness: {e}")

def create_edit_form(parent, rows, tree, fetch_month_rows, selected_row=None, conn=None, cursor=None, month=None, year=None, local_accounts=None, account_data=None):
    if conn is None or cursor is None or month is None or year is None or local_accounts is None or account_data is None:
        raise ValueError("Database connection, month/year, and account data required")
    current_time = datetime.now()
    day = current_time.day
        
    form = tk.Toplevel(parent)
    win_id = 17
    if selected_row is None:
        open_form_with_position(form, conn, cursor, win_id, "New Transaction")
    else:
        open_form_with_position(form, conn, cursor, win_id, "Edit Transaction")
        
    form.geometry(f"{sc(850)}x{sc(500)}")  # Adjust size
    form.resizable(False, False)
    form.grab_set()
    
    ############################################################################
    # Prepare to use custom checkbox
    ignore_checkbox_var = tk.IntVar(value=0)
    unchecked_img = tk.PhotoImage(file=resource_path("icons/unchecked_16.png")).zoom(sc(1))
    checked_img = tk.PhotoImage(file=resource_path("icons/checked_16.png")).zoom(sc(1))
    
    # Use 32x32 base icons for better scaling on high-DPI displays
    radio_unchecked_img = tk.PhotoImage(file=resource_path("icons/radio_0_32.png")).zoom(int(sc(1) * 0.5))
    radio_checked_img = tk.PhotoImage(file=resource_path("icons/radio_1_32.png")).zoom(int(sc(1) * 0.5))
    
    form.option_add("*TCombobox*Listbox*Font", (config.ha_normal))  # normal HA font size for Combobox dropdown list
    
    # Initialize image references to prevent garbage collection
    image_refs = []
    image_refs.extend([unchecked_img, checked_img, radio_unchecked_img, radio_checked_img])
    ############################################################################
    
    # Define before using Ignore checkbox control to avoid scope error
    def toggle_ignore():
        if ignore_checkbox_var.get() == 1:  # Checkbox is ticked
            ignore_box.config(image=unchecked_img)
            ignore_checkbox_var.set(0)
        else:
            ignore_box.config(image=checked_img)
            ignore_checkbox_var.set(1)
        update_fields("ignore", "", "")
            
    # Define all widgets first to ensure they are in scope for functions
    # Left Side
    tk.Label(form, text="Transaction Date:", font=(config.ha_normal), 
            padx=sc(5), pady=sc(2)).place(x=sc(90), y=sc(20))
    cal = Calendar(form, selectmode="day", year=year, month=month, day=day)
    cal.place(x=sc(40), y=sc(50))

    # Transaction Type Control
    tt_frame = tk.Frame(form, relief="solid", borderwidth=1, height=sc(100), width=sc(270))
    tt_frame.place(x=sc(40), y=sc(220))
    
    tk.Label(tt_frame, text="Transaction Type:", font=(config.ha_normal), width=14, anchor=tk.E, 
            fg=COLORS["black"]).place(x=sc(20), y=sc(40))
    radio_var = tk.StringVar(value="Expense")
    # Store transaction type buttons and their values
    trans_type_buttons = []
    
    # Income button
    tk.Label(tt_frame, text="Income", anchor=tk.W, width=10,
            font=(config.ha_normal), fg=COLORS["black"]).place(
            x=sc(180), y=sc(20))
    ttype_box1 = tk.Button(tt_frame, image=radio_unchecked_img, bd=0)
    ttype_box1.place(x=sc(150), y=sc(20))
    trans_type_buttons.append((ttype_box1, 1))

    # Expense button
    tk.Label(tt_frame, text="Expense", anchor=tk.W, width=10,
            font=(config.ha_normal), fg=COLORS["black"]).place(
            x=sc(180), y=sc(40))
    ttype_box2 = tk.Button(tt_frame, image=radio_unchecked_img, bd=0)
    ttype_box2.place(x=sc(150), y=sc(40))
    trans_type_buttons.append((ttype_box2, 2))

    # Transfer button
    tk.Label(tt_frame, text="Transfer", anchor=tk.W, width=10,
            font=(config.ha_normal), fg=COLORS["black"]).place(
            x=sc(180), y=sc(60))
    ttype_box3 = tk.Button(tt_frame, image=radio_unchecked_img, bd=0)
    ttype_box3.place(x=sc(150), y=sc(60))
    trans_type_buttons.append((ttype_box3, 3))

    # Status Control
    ts_frame = tk.Frame(form, relief="solid", borderwidth=1, height=sc(100), width=sc(270))
    ts_frame.place(x=sc(40), y=sc(340))
    
    tk.Label(ts_frame, text="Status:", font=(config.ha_normal), width=14, anchor=tk.E,
            fg=COLORS["black"]).place(x=sc(20), y=sc(40))
    status_var = tk.StringVar(value="Forecast")
    
    # Store status buttons and their values
    status_buttons = []
    
    # Complete button
    tk.Label(ts_frame, text="Complete", anchor=tk.W, width=10, 
            font=(config.ha_normal), fg=COLORS["black"]).place(
            x=sc(180), y=sc(20))
    status_box1 = tk.Button(ts_frame, image=radio_unchecked_img, bd=0)
    status_box1.place(x=sc(150), y=sc(20))
    status_buttons.append((status_box1, 3))
    
    # Processing button
    tk.Label(ts_frame, text="Processing", anchor=tk.W, width=10, 
            font=(config.ha_normal), fg=COLORS["black"]).place(
            x=sc(180), y=sc(40))
    status_box2 = tk.Button(ts_frame, image=radio_unchecked_img, bd=0)
    status_box2.place(x=sc(150), y=sc(40))
    status_buttons.append((status_box2, 2))
    
    # Forecast button
    tk.Label(ts_frame, text="Forecast", anchor=tk.W, width=10, 
            font=(config.ha_normal), fg=COLORS["black"]).place(
            x=sc(180), y=sc(60))
    status_box3 = tk.Button(ts_frame, image=radio_unchecked_img, bd=0)
    status_box3.place(x=sc(150), y=sc(60))
    status_buttons.append((status_box3, 1))

    # Right Side
    flag_btn_frame = tk.LabelFrame(form, text="Set/Clear Flag:", font=(config.ha_normal), 
                                padx=sc(10), pady=sc(20))
    flag_btn_frame.place(x=sc(400), y=sc(20))
    tk.Button(flag_btn_frame, text="Set", font=(config.ha_normal), width=6, bg=COLORS["flag_y_bg"], 
            command=lambda: set_flag_edit(1)).pack(side="left", padx=sc(10))
    tk.Button(flag_btn_frame, text="Set", font=(config.ha_normal), width=6, bg=COLORS["flag_g_bg"], 
            command=lambda: set_flag_edit(2)).pack(side="left", padx=sc(10))
    tk.Button(flag_btn_frame, text="Set", font=(config.ha_normal), width=6, bg=COLORS["flag_b_bg"], 
            command=lambda: set_flag_edit(4)).pack(side="left", padx=sc(10))

    tk.Label(form, text="Amount:   £", font=(config.ha_normal)).place(
        x=sc(400), y=sc(130))
    amount_entry = tk.Entry(form, width=12, font=(config.ha_normal), justify="left")
    amount_entry.place(x=sc(530), y=sc(130))

    cc_pay_btn = tk.Button(form, text="", anchor="w", font=(config.ha_normal), state="disabled", bg=COLORS["pale_green"])
    cc_pay_btn.place(x=sc(650), y=sc(125), width=sc(150))

    tk.Label(form, text="Description:", font=(config.ha_normal)).place(
        x=sc(400), y=sc(170))
    desc_entry = tk.Entry(form, width=36, font=(config.ha_normal), justify="left")
    desc_entry.place(x=sc(530), y=sc(170))

    # Fetch full account names for the selected year
    full_accounts = fetch_account_full_names(cursor, year)
    source_label = tk.Label(form, text="Source Account:", font=(config.ha_normal))
    source_label.place(x=sc(400), y=sc(210))
    source_var = tk.StringVar(value=full_accounts[0] if full_accounts else "Cash")
    source_menu = ttk.Combobox(form, state="readonly", textvariable=source_var, values=full_accounts, font=(config.ha_normal))
    source_menu.place(x=sc(530), y=sc(205), 
                    width=sc(180), height=sc(30))
    source_menu.set(full_accounts[0] if full_accounts else "Cash")    

    dest_label = tk.Label(form, text="Destination Account:", font=(config.ha_normal))
    dest_label.place(x=sc(400), y=sc(250))
    dest_var = tk.StringVar(value=full_accounts[0] if full_accounts else "Cash")
    dest_menu = ttk.Combobox(form, textvariable=dest_var, values=full_accounts, state="readonly", font=(config.ha_normal))
    dest_menu.place(x=sc(530), y=sc(245), 
                    width=sc(180), height=sc(30))
    dest_menu.set(full_accounts[0] if full_accounts else "Cash")    

    cat_label = tk.Label(form, text="Category:", font=(config.ha_normal))
    cat_label.place(x=sc(400), y=sc(290))
    cat_var = tk.StringVar(value="")
    categories = fetch_categories(cursor, year, is_income=False)
    cat_options = {row[1]: row[0] for row in categories}
    cat_menu = ttk.Combobox(form, textvariable=cat_var, values=list(cat_options.keys()), state="readonly", font=(config.ha_normal))
    cat_menu.place(x=sc(530), y=sc(285), 
                width=sc(180), height=sc(30))

    # Display Ignore Control and Initialise
    ignore_label = tk.Label(form, text="Ignore", anchor=tk.W, width=10, font=(config.ha_button))
    ignore_label.place(x=sc(745), y=sc(290))
    ignore_box = tk.Button(form, image=unchecked_img, command = lambda: toggle_ignore())
    ignore_box.place(x=sc(720), y=sc(290))
    
    scat_label = tk.Label(form, text="Sub-Category:", font=(config.ha_normal))
    scat_label.place(x=sc(400), y=sc(330))
    scat_var = tk.StringVar(value="")
    scat_menu = ttk.Combobox(form, textvariable=scat_var, values=[], state="readonly", font=(config.ha_normal))
    scat_menu.place(x=sc(530), y=sc(325), 
                    width=sc(180), height=sc(30))

    tk.Label(form, text="Notes:", width=12, font=(config.ha_normal), anchor=tk.W).place(
        x=sc(400), y=sc(370))
    notes_label = tk.Label(form, text=fetch_notes(cursor), width=35, height=4, anchor=tk.W, justify="left")
    notes_label.place(x=sc(530), y=sc(370))

    save_btn = tk.Button(form, text="Save Transaction", font=(config.ha_normal))
    save_btn.place(x=sc(650), y=sc(450), width=sc(150))
    if selected_row is not None:
        delete_btn = tk.Button(form, text="Delete Transaction", font=(config.ha_normal), 
                            bg=COLORS["del_but_bg"])
        delete_btn.place(x=sc(50), y=sc(450), width=sc(150))
        cancel_btn = tk.Button(form, text="Cancel", font=(config.ha_normal), bg=COLORS["exit_but_bg"])
        cancel_btn.place(x=sc(375), y=sc(450), width=sc(100))
    else:
        cancel_btn = tk.Button(form, text="Cancel", font=(config.ha_normal), bg=COLORS["exit_but_bg"])
        cancel_btn.place(x=sc(375), y=sc(450), width=sc(100))

    # Define functions after widgets to ensure all widgets are in scope
    def set_flag_edit(flag_new_value):
        flag_map_num = {1: COLORS["flag_y_bg"], 2: COLORS["flag_g_bg"], 4: COLORS["flag_b_bg"]}
        current_colour = desc_entry.cget("bg")
        new_colour = flag_map_num.get(flag_new_value, "SystemWindow")
        desc_entry.configure(bg="SystemWindow" if current_colour == new_colour else new_colour)

    # Forward declaration for update_fields to allow set_trans_type and set_status to reference it
    def update_fields(*args):
        pass

    def set_trans_type(ttype_flag, buttons, radio_var, checked_img, unchecked_img):
        """
        Update transaction type radio buttons and set radio_var.
        
        Args:
            ttype_flag (int): Transaction type (1=Income, 2=Expense, 3=Transfer)
            buttons (list): List of (button, value) tuples for radio buttons
            radio_var (tk.StringVar): Variable tracking selected transaction type
            checked_img (tk.PhotoImage): Image for checked radio button
            unchecked_img (tk.PhotoImage): Image for unchecked radio button
        """
        #logger.debug(f"Setting transaction type: {ttype_flag}")
        value_map = {1: "Income", 2: "Expense", 3: "Transfer"}
        if ttype_flag not in value_map:
            logger.error(f"Invalid transaction type flag: {ttype_flag}")
            return
        
        # Update button images
        for btn, value in buttons:
            btn.config(image=checked_img if value == ttype_flag else unchecked_img)
        
        # Set radio_var and update fields
        radio_var.set(value_map[ttype_flag])
        update_fields("", "", "")  # Placeholders, as args are not used

    def set_status(status_flag, buttons, status_var, checked_img, unchecked_img):
        """
        Update status radio buttons and set status_var.
        
        Args:
            status_flag (int): Status type (1=Forecast, 2=Processing, 3=Complete)
            buttons (list): List of (button, value) tuples for radio buttons
            status_var (tk.StringVar): Variable tracking selected status
            checked_img (tk.PhotoImage): Image for checked radio button
            unchecked_img (tk.PhotoImage): Image for unchecked radio button
        """
        #logger.debug(f"Setting status: {status_flag}")
        value_map = {1: "Forecast", 2: "Processing", 3: "Complete"}
        if status_flag not in value_map:
            logger.error(f"Invalid status flag: {status_flag}")
            return
        
        # Update button images
        for btn, value in buttons:
            btn.config(image=checked_img if value == status_flag else unchecked_img)
        
        # Set status_var and update fields
        status_var.set(value_map[status_flag])
        update_fields("", "", "")  # Placeholders, as args are not used

    def update_fields(ignore, a, b):
        """
        Update form fields based on selected transaction type.
        
        Args:
            *args: Ignored trace callback arguments (name, index, mode)
        """
        if ignore == "ignore" and ignore_checkbox_var.get() == 1:      # we came here from toggle_ignore() and Checkbox is ticked
            cat_menu.configure(state="disabled")
            cat_label.configure(foreground="grey")
            scat_menu.configure(state="disabled")
            scat_label.configure(foreground="grey")
            cat_var.set("Ignore")
            scat_var.set("")
        #    else:
        #        cat_menu.configure(state="readonly")
        #        cat_label.configure(foreground="black")
        #        categories = fetch_categories(cursor, year, is_income=True)
        #        cat_options.clear()
        #        cat_options.update({row[1]: row[0] for row in categories})
        #        cat_menu.configure(values=list(cat_options.keys()))
        #        cat_var.set("Income" if cat_options else "")
                #logger.debug(f"Income categories: {list(cat_options.keys())}")
        #        scat_menu.configure(state="readonly")
        #        scat_label.configure(foreground="black")
        else:
            radio = radio_var.get()
            #logger.debug(f"Updating fields based on radio_var: {radio}")
            if radio == "Income":
                source_menu.configure(state="disabled")
                source_label.configure(foreground=COLORS["grey"])
                dest_menu.configure(state="readonly")
                dest_label.configure(foreground=COLORS["black"])
                source_var.set("")
                cat_menu.configure(state="readonly")
                cat_label.configure(foreground=COLORS["black"])
                categories = fetch_categories(cursor, year, is_income=True)
                cat_options.clear()
                cat_options.update({row[1]: row[0] for row in categories})
                cat_menu.configure(values=list(cat_options.keys()))
                cat_var.set("Income" if cat_options else "")
                #logger.debug(f"Income categories: {list(cat_options.keys())}")
                scat_menu.configure(state="readonly")
                scat_label.configure(foreground=COLORS["black"])
                cc_pay_btn.config(text="", state="disabled")
                cc_pay_btn.place_forget()
                ignore_box.configure(state="normal", foreground=COLORS["black"])
                if ignore_checkbox_var.get() == 1:
                    toggle_ignore()
                ignore_label.configure(foreground=COLORS["black"])

            elif radio == "Expense":
                source_menu.configure(state="readonly")
                source_label.configure(foreground=COLORS["black"])
                dest_menu.configure(state="disabled")
                dest_label.configure(foreground=COLORS["grey"])
                dest_var.set("")
                cat_menu.configure(state="readonly")
                cat_label.configure(foreground=COLORS["black"])
                categories = fetch_categories(cursor, year, is_income=False)
                cat_options.clear()
                cat_options.update({row[1]: row[0] for row in categories})
                cat_menu.configure(values=list(cat_options.keys()))
                cat_var.set("")
                #logger.debug(f"Expense categories: {list(cat_options.keys())}")
                if cat_var.get():
                    scat_menu.configure(state="readonly")
                    scat_label.configure(foreground=COLORS["black"])
                else:
                    scat_menu.configure(state="disabled")
                    scat_label.configure(foreground=COLORS["grey"])
                    scat_var.set("")
                cc_pay_btn.config(text="", state="disabled")
                cc_pay_btn.place_forget()
                ignore_box.configure(state="normal", foreground=COLORS["black"])
                if ignore_checkbox_var.get() == 1:
                    toggle_ignore()
                ignore_label.configure(foreground=COLORS["black"])


            else:  # Transfer
                source_menu.configure(state="readonly")
                source_label.configure(foreground=COLORS["black"])
                dest_menu.configure(state="readonly")
                dest_label.configure(foreground=COLORS["black"])
                cat_menu.configure(state="disabled")
                cat_label.configure(foreground=COLORS["grey"])
                scat_menu.configure(state="disabled")
                scat_label.configure(foreground=COLORS["grey"])
                cat_var.set("")
                scat_var.set("")
                if selected_row is None and not desc_entry.get().strip():
                    desc_entry.delete(0, tk.END)
                    desc_entry.insert(0, "Transfer")
                dest_idx = full_accounts.index(dest_var.get()) + 1 if dest_var.get() in full_accounts else 0 
                if 6 < dest_idx < 13:  # Transfer to a credit account
                    # Fetch Last Statement Balance
                    stmt_balances = fetch_statement_balances(cursor, month, year, local_accounts)
                    pay_amt = stmt_balances[dest_idx - 1]
                    if pay_amt is not None and pay_amt < 0:
                        pay_amt = 0 - pay_amt
                        cc_pay_btn.config(text="<<   {:,.2f}".format(pay_amt), bg=COLORS["act_but_bg"], highlightbackground=COLORS["act_but_hi_bg"], state="active")
                        cc_pay_btn.place(x=sc(650), y=sc(125), width=sc(150))
                        cc_pay_btn.update()  # Force button redraw
                        #logger.debug(f"cc_pay_btn shown: text='<<   {pay_amt:,.2f}', bg={COLORS['pale_green']}")
                    else:
                        cc_pay_btn.config(bg=COLORS["grey"], highlightbackground=COLORS["grey"], state="disabled")
                        cc_pay_btn.place_forget()
                        #logger.debug(f"cc_pay_btn hidden: bg={COLORS['pale_grey']}, state=disabled")
                else:
                    cc_pay_btn.config(bg=COLORS["grey"], highlightbackground=COLORS["grey"], state="disabled")
                    cc_pay_btn.place_forget()
                    #logger.debug(f"cc_pay_btn hidden: bg={COLORS['pale_grey']}, state=disabled")
                ignore_box.config(image=unchecked_img)
                ignore_box.configure(state="disabled", foreground=COLORS["grey"])
                ignore_checkbox_var.set(0)
                ignore_label.configure(foreground=COLORS["grey"])
        form.update_idletasks()  # Force redraw

    def fill_amount():
        amount_entry.delete(0, tk.END)
        pay_amt = float(cc_pay_btn.cget("text")[5:].replace(",", ""))
        amount_entry.insert(0, f"{pay_amt:.2f}")

    def update_subcategories(*args):
        scat_menu.configure(state="readonly")
        scat_label.configure(foreground=COLORS["black"])
        selected_cat = cat_var.get()
        if selected_cat in cat_options and selected_cat != "Ignore":
            pid = cat_options[selected_cat]
            subcategories = fetch_subcategories(cursor, pid, year)
            scat_options = {row[1]: row[0] for row in subcategories}
            scat_menu.configure(values=list(scat_options.keys()))
            scat_var.set(subcategories[0][1] if subcategories else "")
            #logger.debug(f"Subcategories for {selected_cat}: {list(scat_options.keys())}")
        else:
            scat_menu.configure(values=[])
            scat_var.set("")
            scat_menu.configure(state="disabled")
            scat_label.configure(foreground=COLORS["grey"])
        form.update_idletasks()  # Force redraw

    def update_paycc(*args):
        dest_idx = full_accounts.index(dest_var.get()) + 1 if dest_var.get() in full_accounts else 0 
        if radio_var.get() == "Transfer" and 6 < dest_idx < 13:  # Transfer to a credit account
            # Fetch Last Statement Balance
            stmt_balances = fetch_statement_balances(cursor, month, year, local_accounts)
            pay_amt = stmt_balances[dest_idx - 1]
            if pay_amt is not None and pay_amt < 0:
                pay_amt = 0 - pay_amt
                cc_pay_btn.config(text="<<   {:,.2f}".format(pay_amt), bg=COLORS["pale_green"], highlightbackground=COLORS["pale_green"], state="active")
                cc_pay_btn.place(x=sc(650), y=sc(125), width=sc(150))
                cc_pay_btn.update()  # Force button redraw
                #logger.debug(f"cc_pay_btn shown: text='<<   {pay_amt:,.2f}', bg={COLORS['pale_green']}")
            else:
                cc_pay_btn.config(bg=COLORS["pale_grey"], highlightbackground=COLORS["pale_grey"], state="disabled")
                cc_pay_btn.place_forget()
                #logger.debug(f"cc_pay_btn hidden: bg={COLORS['pale_grey']}, state=disabled")
        else:
            cc_pay_btn.config(bg=COLORS["pale_grey"], highlightbackground=COLORS["pale_grey"], state="disabled")
            cc_pay_btn.place_forget()
            #logger.debug(f"cc_pay_btn hidden: bg={COLORS['pale_grey']}, state=disabled")
        form.update_idletasks()  # Force redraw

    def save_transaction():
        nonlocal tr_id
        try:
            amount_str = amount_entry.get().strip()
            if not amount_str:
                raise ValueError("Amount cannot be empty")
            try:
                amount = float(amount_str)
                if amount < 0:
                    raise ValueError("Amount must be zero or positive")
            except ValueError as e:
                if str(e) != "Amount must be zero or positive":
                    raise ValueError("Amount must be numeric")
                raise
            desc = desc_entry.get().strip()
            if len(desc.replace(" ", "")) < 3:
                raise ValueError("Description must be at least 3 non-space characters")
            radio = radio_var.get()
            if radio == "Income":
                if not dest_var.get().strip():
                    raise ValueError("Destination Account cannot be empty for Income")
            elif radio == "Expense":
                if not source_var.get().strip():
                    raise ValueError("Source Account cannot be empty for Expense")
            elif radio == "Transfer":
                if not source_var.get().strip():
                    raise ValueError("Source Account cannot be empty for Transfer")
                if not dest_var.get().strip():
                    raise ValueError("Destination Account cannot be empty for Transfer")
            if radio in ["Income", "Expense"]:
                if not cat_var.get().strip():
                    raise ValueError("Category cannot be empty for Income or Expense")
                if cat_var.get() != "Ignore" and not scat_var.get().strip():
                    raise ValueError("Sub-Category cannot be empty when Category is selected")
            status_map = {"Unknown": 0, "Forecast": 1, "Processing": 2, "Complete": 3}
            type_map = {"Income": 1, "Expense": 2, "Transfer": 3}
            status = status_map.get(status_var.get(), 3)
            flag_map_txt = {COLORS["flag_y_bg"]: 1, COLORS["flag_g_bg"]: 2, COLORS["flag_b_bg"]: 4}
            current_colour = desc_entry.cget("bg")
            flag = flag_map_txt.get(current_colour, 0)
            date_obj = cal.get_date()
            date = datetime.strptime(date_obj, "%m/%d/%y")
            day = date.day
            month = date.month
            year = date.year
            source_idx = full_accounts.index(source_var.get()) + 1 if source_var.get() in full_accounts else 0
            dest_idx = full_accounts.index(dest_var.get()) + 1 if dest_var.get() in full_accounts else 0
            if ignore_checkbox_var.get() == 1:  # Checkbox is ticked
                cat_pid = 99
                subcat_cid = 0
            else:
                cat_pid = cat_options.get(cat_var.get()) if radio != "Transfer" else 0
                subcat_options = {row[1]: row[0] for row in fetch_subcategories(cursor, cat_pid or 0, year)} if radio != "Transfer" else {}
                subcat_cid = subcat_options.get(scat_var.get()) if radio != "Transfer" else 0
            if radio == "Income":
                source_idx = 0
            elif radio == "Expense":
                dest_idx = 0
            if selected_row is not None and tr_id is not None:
                update_transaction(cursor, conn, tr_id, type_map[radio], day, month, year, status, flag, amount, desc, source_idx, dest_idx, cat_pid, subcat_cid)
                focus_day = 0
            else:
                tr_id = insert_transaction(cursor, conn, type_map[radio], day, month, year, status, flag, amount, desc, source_idx, dest_idx, cat_pid, subcat_cid)
                focus_day = str(day)
            update_account_month_transaction_total(cursor, conn, month, year, local_accounts)
            form.destroy()
            parent.selected_row.set(-1)
            tree.selection_remove(tree.selection())
            new_rows = fetch_month_rows(cursor, month, year, local_accounts, account_data)
            rows[:] = new_rows
            if hasattr(parent, 'year_var') and hasattr(parent, 'tree'):
                current_year = int(parent.year_var.get())
                current_month = parent.get_current_month()
            if parent:  # Check if root is provided
                parent.update_ahb(parent, current_month, current_year)
            if focus_day == 0:
                refresh_grid(tree, rows, parent.marked_rows, selected_row, 0)
            else:
                refresh_grid(tree, rows, parent.marked_rows, 0, focus_day)
        except ValueError as e:
            logger.error(f"Validation failed: {e}")
            messagebox.showerror("Validation Error", str(e))
        except Exception as e:
            logger.error(f"Save failed: {e}")
            messagebox.showerror("Error", f"Failed to save transaction: {e}")

    def delete_transaction_handler():
        if selected_row is not None and tr_id is not None:
            delete_transaction(cursor, conn, tr_id)
        date_obj = cal.get_date()
        date = datetime.strptime(date_obj, "%m/%d/%y")
        day = date.day
        focus_day = str(day)
        form.destroy()
        parent.selected_row.set(-1)
        tree.selection_remove(tree.selection())
        new_rows = fetch_month_rows(cursor, month, year, local_accounts, account_data)
        rows[:] = new_rows
        refresh_grid(tree, rows, parent.marked_rows, focus_day=focus_day)

    def cancel_transaction():
        close_form_with_position(form, conn, cursor, win_id)

    # Assign commands to buttons after function definitions
    ttype_box1.config(command=lambda: set_trans_type(1, trans_type_buttons, radio_var, radio_checked_img, radio_unchecked_img))
    ttype_box2.config(command=lambda: set_trans_type(2, trans_type_buttons, radio_var, radio_checked_img, radio_unchecked_img))
    ttype_box3.config(command=lambda: set_trans_type(3, trans_type_buttons, radio_var, radio_checked_img, radio_unchecked_img))
    status_box1.config(command=lambda: set_status(3, status_buttons, status_var, radio_checked_img, radio_unchecked_img))
    status_box2.config(command=lambda: set_status(2, status_buttons, status_var, radio_checked_img, radio_unchecked_img))
    status_box3.config(command=lambda: set_status(1, status_buttons, status_var, radio_checked_img, radio_unchecked_img))
    cc_pay_btn.config(command=lambda: fill_amount())
    save_btn.config(command=save_transaction)
    if selected_row is not None:
        delete_btn.config(command=delete_transaction_handler)
        cancel_btn.config(command=cancel_transaction)
    else:
        cancel_btn.config(command=cancel_transaction)

    # Set up traces
    radio_var.trace("w", lambda *args: update_fields(*args))
    status_var.trace("w", lambda *args: update_fields(*args))
    cat_var.trace("w", update_subcategories)
    dest_var.trace("w", update_paycc)

    # Initialize buttons based on initial values
    initial_trans_flag = {"Income": 1, "Expense": 2, "Transfer": 3}.get(radio_var.get(), 2)
    set_trans_type(initial_trans_flag, trans_type_buttons, radio_var, radio_checked_img, radio_unchecked_img)
    initial_status_flag = {"Forecast": 1, "Processing": 2, "Complete": 3}.get(status_var.get(), 1)
    set_status(initial_status_flag, status_buttons, status_var, radio_checked_img, radio_unchecked_img)

    # Handle existing transaction
    tr_id = None
    if selected_row is not None and 0 <= selected_row < len(rows) and rows[selected_row]["status"] != "Total":
        tr_id = rows[selected_row].get("tr_id")
        if tr_id:
            cursor.execute("SELECT Tr_Type, Tr_Day, Tr_Month, Tr_Year, Tr_Stat, Tr_Query_Flag, Tr_Amount, Tr_Desc, Tr_Acc_From, Tr_Acc_To, Tr_Exp_ID, Tr_ExpSub_ID "
                        "FROM Trans WHERE Tr_ID=?", (tr_id,))
            result = cursor.fetchone()
            if result:
                tr_type, tr_day, tr_month, tr_year, tr_stat, tr_query_flag, tr_amount, tr_desc, tr_acc_from, tr_acc_to, tr_exp_id, tr_expsub_id = result
                type_map = {1: "Income", 2: "Expense", 3: "Transfer"}
                radio_var.set(type_map.get(tr_type, "Expense"))
                set_trans_type({"Income": 1, "Expense": 2, "Transfer": 3}.get(radio_var.get(), 2), 
                            trans_type_buttons, radio_var, radio_checked_img, radio_unchecked_img)
                cal.selection_set(f"{tr_month:02d}/{tr_day:02d}/{str(tr_year)[-2:]}")
                amount_entry.insert(0, f"{abs(tr_amount):.2f}")
                desc_entry.insert(0, tr_desc or "")
                status_map = {0: "Unknown", 1: "Forecast", 2: "Processing", 3: "Complete"}
                status_var.set(status_map.get(tr_stat, "Forecast"))
                set_status({"Forecast": 1, "Processing": 2, "Complete": 3}.get(status_var.get(), 1), 
                        status_buttons, status_var, radio_checked_img, radio_unchecked_img)
                flag = (tr_query_flag if tr_query_flag in [0, 1, 2, 4] else 0)
                set_flag_edit(flag)
                source_account = full_accounts[tr_acc_from - 1] if tr_acc_from and 0 < tr_acc_from <= len(full_accounts) else ""
                dest_account = full_accounts[tr_acc_to - 1] if tr_acc_to and 0 < tr_acc_to <= len(full_accounts) else ""
                if radio_var.get() == "Income":
                    dest_var.set(dest_account)
                elif radio_var.get() == "Expense":
                    source_var.set(source_account)
                elif radio_var.get() == "Transfer":
                    source_var.set(source_account)
                    dest_var.set(dest_account)
                if radio_var.get() == "Transfer" and 6 < tr_acc_to < 13:
                    stmt_balances = fetch_statement_balances(cursor, month, year, local_accounts)
                    pay_amt = stmt_balances[tr_acc_to - 1]
                    if pay_amt is not None and pay_amt < 0:
                        pay_amt = 0 - pay_amt
                        cc_pay_btn.config(text="<<   {:,.2f}".format(pay_amt), bg=COLORS["act_but_bg"], highlightbackground=COLORS["act_but_hi_bg"], state="active")
                        cc_pay_btn.place(x=sc(650), y=sc(125), width=sc(150))
                        cc_pay_btn.update()  # Force button redraw
                        #logger.debug(f"cc_pay_btn shown (existing transaction): text='<<   {pay_amt:,.2f}', bg={COLORS['pale_green']}")
                    else:
                        cc_pay_btn.config(bg=COLORS["grey"], highlightbackground=COLORS["grey"], state="disabled")
                        cc_pay_btn.place_forget()
                        #logger.debug(f"cc_pay_btn hidden (existing transaction): bg={COLORS['pale_grey']}, state=disabled")
                update_fields("", "", "")
                if tr_exp_id and tr_type != 3:
                    if tr_exp_id == 99:
                        ignore_box.configure(state="normal", foreground=COLORS["black"])
                        ignore_checkbox_var.set(1)
                        ignore_label.configure(foreground=COLORS["black"])
                        cat_var.set("Ignore")
                        scat_var.set("")
                        cat_menu.configure(state="disabled")
                        cat_label.configure(foreground=COLORS["grey"])
                        scat_menu.configure(state="disabled")
                        scat_label.configure(foreground=COLORS["grey"])
                    else:
                        for desc, pid in cat_options.items():
                            if pid == tr_exp_id:
                                cat_var.set(desc)
                                update_subcategories()
                                break
                if tr_expsub_id and tr_type != 3:
                    subcategories = fetch_subcategories(cursor, tr_exp_id, year)
                    scat_options = {row[1]: row[0] for row in subcategories}
                    for desc, cid in scat_options.items():
                        if cid == tr_expsub_id:
                            scat_var.set(desc)
                            break
    else:
        update_fields("", "", "")

    # Keep image references alive
    form.image_refs = image_refs

    form.wait_window()

def show_maint_toolbox(parent, conn, cursor):
    toolbox = tk.Toplevel(parent)
    win_id = 14
    open_form_with_position(toolbox, conn, cursor, win_id, "Maintenance Forms")
    toolbox.geometry(f"{sc(280)}x{sc(450)}")  # Adjust size
    toolbox.resizable(False, False)
    toolbox.grab_set()

    tk.Label(toolbox, text="Maintenance Forms", font=(config.ha_head12)).pack(pady=5)
    tk.Button(toolbox, text="Manage Accounts", width=36,
            command=lambda: [close_form_with_position(toolbox, conn, cursor, win_id), 
                            create_account_maint_form(parent, conn, cursor, root=parent)]).pack(pady=2)
    tk.Button(toolbox, text="Manage Income/Expense Categories", width=36, bg=COLORS["home_test_bg"],
            command=lambda: [close_form_with_position(toolbox, conn, cursor, win_id),
                            create_category_maint_form(parent, conn, cursor)]).pack(pady=2)
    tk.Button(toolbox, text="Manage Regular Transactions", width=36,
            command=lambda: [close_form_with_position(toolbox, conn, cursor, win_id),
                            create_regular_transactions_maint_form(parent, conn, cursor)]).pack(pady=2)
    tk.Button(toolbox, text="Manage Colour Scheme", width=36, 
            command=lambda: [close_form_with_position(toolbox, conn, cursor, win_id),
                            create_colour_scheme_maint_form(parent, conn, cursor)]).pack(pady=2)
    tk.Button(toolbox, text="Manage Account Years", width=36, bg=COLORS["home_test_bg"], 
            command=lambda: [close_form_with_position(toolbox, conn, cursor, win_id),
                            create_account_years_maint_form(parent, conn, cursor)]).pack(pady=2)
    tk.Button(toolbox, text="Manage Annual Budget", width=36, bg=COLORS["home_test_bg"],
            command=lambda: [close_form_with_position(toolbox, conn, cursor, win_id),
                            create_annual_budget_maint_form(parent, conn, cursor)]).pack(pady=2)
    tk.Button(toolbox, text="Manage Transaction Options", width=36, bg=COLORS["home_test_bg"],
            command=lambda: [close_form_with_position(toolbox, conn, cursor, win_id),
                            create_transaction_options_maint_form(parent, conn, cursor)]).pack(pady=2)
    tk.Button(toolbox, text="Manage Form Window Positions", width=36, 
            command=lambda: [close_form_with_position(toolbox, conn, cursor, win_id),
                            create_form_positions_maint_form(parent, conn, cursor)]).pack(pady=2)
    tk.Button(toolbox, text="Manage FF Category Mappings", width=36, bg=COLORS["home_test_bg"], 
            command=lambda: [close_form_with_position(toolbox, conn, cursor, win_id),
                            create_ff_mappings_maint_form(parent, conn, cursor)]).pack(pady=2)
    tk.Button(toolbox, text="Export Transactions to .CSV file", width=36, bg=COLORS["home_test_bg"], 
            command=lambda: [close_form_with_position(toolbox, conn, cursor, win_id),
                            create_export_transactions_form(parent, conn, cursor)]).pack(pady=2)
    tk.Button(toolbox, text="Manage GoCardless Configuration", width=36, 
            command=lambda: [close_form_with_position(toolbox, conn, cursor, win_id),
                            create_gocardless_maint_form(parent, conn, cursor)]).pack(pady=2)
    tk.Button(toolbox, text="Data Import Rules", width=36, 
            command=lambda: [close_form_with_position(toolbox, conn, cursor, win_id),
                            create_rules_form(parent, conn, cursor)]).pack(pady=2)
    tk.Button(toolbox, text="Close", width=36, bg=COLORS["exit_but_bg"],
            command=lambda: close_form_with_position(toolbox, conn, cursor, win_id)).pack(pady=5)