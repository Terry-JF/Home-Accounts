# Creates main form for the HA2 program

import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import Calendar
from datetime import datetime
import logging

from db import  (insert_transaction, update_transaction, delete_transaction, fetch_notes, fetch_categories, fetch_subcategories,
                fetch_account_full_names, fetch_statement_balances, update_account_month_transaction_total)
from ui_utils import (COLORS, refresh_grid, open_form_with_position, close_form_with_position)
from gui_maint import   (create_account_maint_form, create_category_maint_form, create_export_transactions_form,
                        create_colour_scheme_maint_form, create_account_years_maint_form, create_annual_budget_maint_form, 
                        create_transaction_options_maint_form, create_form_positions_maint_form, create_ff_mappings_maint_form)
from gc_utils import (create_gocardless_maint_form)
from gui_maint_rules import (create_rules_form)
from m_reg_trans import (create_regular_transactions_maint_form)

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
    
    ignore_checkbox_var = tk.IntVar(value=0)
    
    form = tk.Toplevel(parent)
    win_id = 17
    if selected_row is None:
        open_form_with_position(form, conn, cursor, win_id, "New Transaction")
    else:
        open_form_with_position(form, conn, cursor, win_id, "Edit Transaction")
        
    scaling_factor = form.winfo_fpixels('1i') / 96  # Get scaling factor (e.g., 2.0 for 200% scaling)
    form.geometry(f"{int(850 * scaling_factor)}x{int(500 * scaling_factor)}")  # Adjust size
    #form.geometry("850x500") 
    form.resizable(False, False)
    form.grab_set()

    def set_flag_edit(flag_new_value):
        flag_map_num = {1: COLORS["yellow"], 2: COLORS["green"], 4: COLORS["cyan"]}
        current_colour = desc_entry.cget("bg")
        new_colour = flag_map_num.get(flag_new_value, "SystemWindow")
        desc_entry.configure(bg="SystemWindow" if current_colour == new_colour else new_colour)

    # Left Side
    tk.Label(form, text="Transaction Date:", font=("Arial", 11), padx=int(5 * scaling_factor), pady=int(2 * scaling_factor)).place(x=int(90 * scaling_factor), y=int(20 * scaling_factor))
    cal = Calendar(form, selectmode="day", year=year, month=month, day=day)
    cal.place(x=int(40 * scaling_factor), y=int(50 * scaling_factor))

    tk.Label(form, text="Transaction Type:", font=("Arial", 11)).place(x=int(50 * scaling_factor), y=int(280 * scaling_factor))
    radio_var = tk.StringVar(value="Expense")
    tk.Radiobutton(form, text="Income", variable=radio_var, value="Income", font=("Arial", 11)).place(x=int(180 * scaling_factor), y=int(260 * scaling_factor))
    tk.Radiobutton(form, text="Expense", variable=radio_var, value="Expense", font=("Arial", 11)).place(x=int(180 * scaling_factor), y=int(280 * scaling_factor))
    tk.Radiobutton(form, text="Transfer", variable=radio_var, value="Transfer", font=("Arial", 11)).place(x=int(180 * scaling_factor), y=int(300 * scaling_factor))

    tk.Label(form, text="Status:", font=("Arial", 11)).place(x=int(120 * scaling_factor), y=int(360 * scaling_factor))
    status_var = tk.StringVar(value="Forecast")
    tk.Radiobutton(form, text="Complete", variable=status_var, value="Complete", font=("Arial", 11)).place(x=int(180 * scaling_factor), y=int(340 * scaling_factor))
    tk.Radiobutton(form, text="Processing", variable=status_var, value="Processing", font=("Arial", 11)).place(x=int(180 * scaling_factor), y=int(360 * scaling_factor))
    tk.Radiobutton(form, text="Forecast", variable=status_var, value="Forecast", font=("Arial", 11)).place(x=int(180 * scaling_factor), y=int(380 * scaling_factor))

    # Right Side
    flag_btn_frame = tk.LabelFrame(form, text="Set/Clear Flag:", font=("Arial", 11), padx=int(10 * scaling_factor), pady=int(20 * scaling_factor))
    # flag_var = tk.IntVar(value=0)
    flag_btn_frame.place(x=int(400 * scaling_factor), y=int(20 * scaling_factor))
    tk.Button(flag_btn_frame, text="Set", font=("Arial", 10), width=6, bg=COLORS["yellow"], command=lambda: set_flag_edit(1)).pack(side="left", padx=int(10 * scaling_factor))
    tk.Button(flag_btn_frame, text="Set", font=("Arial", 10), width=6, bg=COLORS["green"], command=lambda: set_flag_edit(2)).pack(side="left", padx=int(10 * scaling_factor))
    tk.Button(flag_btn_frame, text="Set", font=("Arial", 10), width=6, bg=COLORS["cyan"], command=lambda: set_flag_edit(4)).pack(side="left", padx=int(10 * scaling_factor))

    tk.Label(form, text="Amount:   £", font=("Arial", 10)).place(x=int(400 * scaling_factor), y=int(130 * scaling_factor))
    amount_entry = tk.Entry(form, width=12, font=("Arial", 10), justify="right")
    amount_entry.place(x=int(530 * scaling_factor), y=int(130 * scaling_factor))

    cc_pay_btn = tk.Button(form, text="", anchor="w", font=("Arial", 10), state="disabled", command=lambda: fill_amount())
    cc_pay_btn.place(x=int(650 * scaling_factor), y=int(125 * scaling_factor), width=int(150 * scaling_factor))

    tk.Label(form, text="Description:", font=("Arial", 10)).place(x=int(400 * scaling_factor), y=int(170 * scaling_factor))
    desc_entry = tk.Entry(form, width=40, font=("Arial", 10), justify="left")
    desc_entry.place(x=int(530 * scaling_factor), y=int(170 * scaling_factor))

    # Fetch full account names for the selected year
    full_accounts = fetch_account_full_names(cursor, year)
    source_label = tk.Label(form, text="Source Account:", font=("Arial", 10))
    source_label.place(x=int(400 * scaling_factor), y=int(210 * scaling_factor))
    source_var = tk.StringVar(value=full_accounts[0] if full_accounts else "Cash")
    source_menu = ttk.OptionMenu(form, source_var, full_accounts[0] if full_accounts else "Cash", *full_accounts)
    source_menu.place(x=int(530 * scaling_factor), y=int(205 * scaling_factor), width=int(180 * scaling_factor), height=int(30 * scaling_factor))

    dest_label = tk.Label(form, text="Destination Account:", font=("Arial", 10))
    dest_label.place(x=int(400 * scaling_factor), y=int(250 * scaling_factor))
    dest_var = tk.StringVar(value=full_accounts[0] if full_accounts else "Cash")
    dest_menu = ttk.OptionMenu(form, dest_var, full_accounts[0] if full_accounts else "Cash", *full_accounts)
    dest_menu.place(x=int(530 * scaling_factor), y=int(245 * scaling_factor), width=int(180 * scaling_factor), height=int(30 * scaling_factor))

    cat_label = tk.Label(form, text="Category:", font=("Arial", 10))
    cat_label.place(x=int(400 * scaling_factor), y=int(290 * scaling_factor))
    cat_var = tk.StringVar(value="")
    categories = fetch_categories(cursor, year, is_income=False)
    cat_options = {row[1]: row[0] for row in categories}
    cat_menu = ttk.OptionMenu(form, cat_var, "", *cat_options.keys(), command=lambda _: update_subcategories())
    cat_menu.place(x=int(530 * scaling_factor), y=int(285 * scaling_factor), width=int(180 * scaling_factor), height=int(30 * scaling_factor))


    ignore_box=tk.Checkbutton(form, text="Ignore", font=("Arial", 12), command=lambda: toggle_ignore(), variable=ignore_checkbox_var)
    ignore_box.place(x=int(720 * scaling_factor), y=int(285 * scaling_factor))
    

    scat_label = tk.Label(form, text="Sub-Category:", font=("Arial", 10))
    scat_label.place(x=int(400 * scaling_factor), y=int(330 * scaling_factor))
    scat_var = tk.StringVar(value="")
    scat_menu = ttk.OptionMenu(form, scat_var, "")
    scat_menu.place(x=int(530 * scaling_factor), y=int(325 * scaling_factor), width=int(180 * scaling_factor), height=int(30 * scaling_factor))

    def fill_amount():
        amount_entry.delete(0,20) 
        pay_amt = float(cc_pay_btn.cget("text")[5:20].replace(",", ""))
        amount_entry.insert(0, pay_amt)

    def toggle_ignore():
        if ignore_checkbox_var.get() == 1:  # checkbox is ticked
            cat_menu.configure(state="disabled")
            cat_label.configure(foreground="grey")
            scat_menu.configure(state="disabled")
            scat_label.configure(foreground="grey")
            cat_var.set("Ignore")
            scat_var.set("")
        else:
            update_fields()

    def update_subcategories(*args):
        scat_menu.configure(state="normal")
        scat_label.configure(foreground="black")
        selected_cat = cat_var.get()
        if selected_cat in cat_options:
            pid = cat_options[selected_cat]
            subcategories = fetch_subcategories(cursor, pid, year)
            scat_options = {row[1]: row[0] for row in subcategories}
            scat_menu.set_menu(subcategories[0][1] if subcategories else "", *scat_options.keys())
            scat_var.set(subcategories[0][1] if subcategories else "")
        else:
            scat_menu.set_menu("", "")
            scat_menu.configure(state="disabled")
            scat_label.configure(foreground="grey")

    tk.Label(form, text="Notes:", width=12, font=("Arial", 10), anchor=tk.W).place(x=int(400 * scaling_factor), y=int(370 * scaling_factor))
    notes_label = tk.Label(form, text=fetch_notes(cursor), width=35, height=4, anchor=tk.W, justify="left")
    notes_label.place(x=int(525 * scaling_factor), y=int(370 * scaling_factor))

    def update_fields(*args):
        radio = radio_var.get()
        if radio == "Income":
            source_menu.configure(state="disabled")
            source_label.configure(foreground="grey")
            dest_menu.configure(state="normal")
            dest_label.configure(foreground="black")
            source_var.set("")
            cat_menu.configure(state="normal")
            cat_label.configure(foreground="black")
            categories = fetch_categories(cursor, year, is_income=True)
            cat_options.clear()
            cat_options.update({row[1]: row[0] for row in categories})
            cat_menu.set_menu("Income", *cat_options.keys())
            scat_menu.configure(state="normal")
            scat_label.configure(foreground="black")

            cc_pay_btn.config(text="", state="disabled")
            ignore_box.configure(state="normal", foreground="black")

        elif radio == "Expense":
            source_menu.configure(state="normal")
            source_label.configure(foreground="black")
            dest_menu.configure(state="disabled")
            dest_label.configure(foreground="grey")
            dest_var.set("")
            cat_menu.configure(state="normal")
            cat_label.configure(foreground="black")
            categories = fetch_categories(cursor, year, is_income=False)
            cat_options.clear()
            cat_options.update({row[1]: row[0] for row in categories})
            cat_menu.set_menu("", *cat_options.keys())
            cat_var.set("")
            if cat_var.get():
                scat_menu.configure(state="normal")
                scat_label.configure(foreground="black")
            else:
                scat_menu.configure(state="disabled")
                scat_label.configure(foreground="grey")
                scat_var.set("")

            cc_pay_btn.config(text="", state="disabled")
            ignore_box.configure(state="normal", foreground="black")

        else:  # Transfer
            source_menu.configure(state="normal")
            source_label.configure(foreground="black")
            dest_menu.configure(state="normal")
            dest_label.configure(foreground="black")
            cat_menu.configure(state="disabled")
            cat_label.configure(foreground="grey")
            scat_menu.configure(state="disabled")
            scat_label.configure(foreground="grey")
            cat_var.set("")
            scat_var.set("")
            if selected_row is None and not desc_entry.get().strip():
                desc_entry.delete(0, tk.END)
                desc_entry.insert(0, "Transfer")

            dest_idx = full_accounts.index(dest_var.get()) + 1 if dest_var.get() in full_accounts else 0 
            if 6 < dest_idx < 13:  # Transfer to a credit account
                # Fetch Last Statement Balance
                stmt_balances = fetch_statement_balances(cursor, month, year, local_accounts)
                pay_amt= stmt_balances[tr_acc_to - 1]
                if pay_amt is not None and pay_amt < 0:
                    pay_amt = 0 - pay_amt
                    cc_pay_btn.config(text="<<   {:,.2f}".format(pay_amt), state="active")
                else:
                    cc_pay_btn.config(text="", state="disabled")

            ignore_box.configure(state="disabled", foreground="grey")
            ignore_checkbox_var.set(0)


    def update_paycc(*args):
        dest_idx = full_accounts.index(dest_var.get()) + 1 if dest_var.get() in full_accounts else 0 
        if radio_var.get() == "Transfer" and 6 < dest_idx < 13:  # Transfer to a credit account
            # Fetch Last Statement Balance
            stmt_balances = fetch_statement_balances(cursor, month, year, local_accounts)
            pay_amt=stmt_balances[dest_idx - 1]
            if pay_amt is not None and pay_amt < 0:
                pay_amt = 0 - pay_amt
                cc_pay_btn.config(text="<<   {:,.2f}".format(pay_amt), state="active")
            else:
                cc_pay_btn.config(text="", state="disabled")
        else:
            cc_pay_btn.config(text="", state="disabled")

    radio_var.trace("w", update_fields)
    cat_var.trace("w", update_subcategories)
    dest_var.trace("w", update_paycc)

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
                cal.selection_set(f"{tr_month:02d}/{tr_day:02d}/{str(tr_year)[-2:]}")
                amount_entry.insert(0, f"{abs(tr_amount):.2f}")
                desc_entry.insert(0, tr_desc or "")
                status_map = {0: "Unknown", 1: "Forecast", 2: "Processing", 3: "Complete"}
                status_var.set(status_map.get(tr_stat, "Unknown"))

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
                if radio_var.get() == "Transfer" and 6 < tr_acc_to < 13:  # Transfer to a credit account
                    # Fetch Last Statement Balance
                    stmt_balances = fetch_statement_balances(cursor, month, year, local_accounts)
                    pay_amt=stmt_balances[tr_acc_to - 1]
                    if pay_amt is not None and pay_amt < 0:
                        pay_amt = 0 - pay_amt
                        cc_pay_btn.config(text="<<   {:,.2f}".format(pay_amt), state="active")
                    else:
                        cc_pay_btn.config(text="", state="disabled")

                update_fields()
                if tr_exp_id and tr_type != 3:
                    if tr_exp_id == 99:
                        ignore_box.configure(state="normal", foreground="black")
                        ignore_checkbox_var.set(1)
                        cat_var.set("Ignore")
                        scat_var.set("")
                        cat_menu.configure(state="disabled")
                        cat_label.configure(foreground="grey")
                        scat_menu.configure(state="disabled")
                        scat_label.configure(foreground="grey")
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
        update_fields()

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

            flag_map_txt = {COLORS["yellow"]: 1, COLORS["green"]: 2, COLORS["cyan"]: 4}
            current_colour = desc_entry.cget("bg")
            flag = flag_map_txt.get(current_colour, 0)

            date_obj = cal.get_date()
            date = datetime.strptime(date_obj, "%m/%d/%y")
            day = date.day
            month = date.month
            year = date.year
            source_idx = full_accounts.index(source_var.get()) + 1 if source_var.get() in full_accounts else 0
            dest_idx = full_accounts.index(dest_var.get()) + 1 if dest_var.get() in full_accounts else 0
            if ignore_checkbox_var.get() == 1:  # checkbox is ticked
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

    save_btn = tk.Button(form, text="Save Transaction", font=("Arial", 10), command=save_transaction)
    save_btn.place(x=int(650 * scaling_factor), y=int(450 * scaling_factor), width=int(150 * scaling_factor))
    if selected_row is not None:
        delete_btn = tk.Button(form, text="Delete Transaction", font=("Arial", 10), bg=COLORS["red"], command=delete_transaction_handler)
        delete_btn.place(x=int(50 * scaling_factor), y=int(450 * scaling_factor), width=int(150 * scaling_factor))
        cancel_btn = tk.Button(form, text="Cancel", font=("Arial", 10), command=cancel_transaction)
        cancel_btn.place(x=int(375 * scaling_factor), y=int(450 * scaling_factor), width=int(100 * scaling_factor))
    else:
        cancel_btn = tk.Button(form, text="Cancel", font=("Arial", 10), command=cancel_transaction)
        cancel_btn.place(x=int(375 * scaling_factor), y=int(450 * scaling_factor), width=int(100 * scaling_factor))

    form.wait_window()

def show_maint_toolbox(parent, conn, cursor):
    toolbox = tk.Toplevel(parent)
    win_id = 14
    open_form_with_position(toolbox, conn, cursor, win_id, "Maintenance Forms")
    scaling_factor = toolbox.winfo_fpixels('1i') / 96  # Get scaling factor (e.g., 2.0 for 200% scaling)
    toolbox.geometry(f"{int(280 * scaling_factor)}x{int(450 * scaling_factor)}")  # Adjust size
    #toolbox.geometry("280x450")  # Increased height for 12 items/buttons
    toolbox.resizable(False, False)
    toolbox.grab_set()

    tk.Label(toolbox, text="Maintenance Forms", font=("Arial", 12, "bold")).pack(pady=5)
    tk.Button(toolbox, text="Manage Accounts", width=int(18 * scaling_factor), bg=COLORS["pale_green"],
            command=lambda: [close_form_with_position(toolbox, conn, cursor, win_id), 
                            create_account_maint_form(parent, conn, cursor, root=parent)]).pack(pady=2)
    tk.Button(toolbox, text="Manage Income/Expense Categories", width=int(18 * scaling_factor), bg=COLORS["pink"],
            command=lambda: [close_form_with_position(toolbox, conn, cursor, win_id),
                            create_category_maint_form(parent, conn, cursor)]).pack(pady=2)
    tk.Button(toolbox, text="Manage Regular Transactions", width=int(18 * scaling_factor), bg=COLORS["pale_green"],
            command=lambda: [close_form_with_position(toolbox, conn, cursor, win_id),
                            create_regular_transactions_maint_form(parent, conn, cursor)]).pack(pady=2)
    tk.Button(toolbox, text="Manage Colour Scheme", width=int(18 * scaling_factor), bg=COLORS["pink"],
            command=lambda: [close_form_with_position(toolbox, conn, cursor, win_id),
                            create_colour_scheme_maint_form(parent, conn, cursor)]).pack(pady=2)
    tk.Button(toolbox, text="Manage Account Years", width=int(18 * scaling_factor), bg=COLORS["pink"], 
            command=lambda: [close_form_with_position(toolbox, conn, cursor, win_id),
                            create_account_years_maint_form(parent, conn, cursor)]).pack(pady=2)
    tk.Button(toolbox, text="Manage Annual Budget", width=int(18 * scaling_factor), bg=COLORS["pink"],
            command=lambda: [close_form_with_position(toolbox, conn, cursor, win_id),
                            create_annual_budget_maint_form(parent, conn, cursor)]).pack(pady=2)
    tk.Button(toolbox, text="Manage Transaction Options", width=int(18 * scaling_factor), bg=COLORS["pink"],
            command=lambda: [close_form_with_position(toolbox, conn, cursor, win_id),
                            create_transaction_options_maint_form(parent, conn, cursor)]).pack(pady=2)
    tk.Button(toolbox, text="Manage Form Window Positions", width=int(18 * scaling_factor), bg=COLORS["pale_green"],
            command=lambda: [close_form_with_position(toolbox, conn, cursor, win_id),
                            create_form_positions_maint_form(parent, conn, cursor)]).pack(pady=2)
    tk.Button(toolbox, text="Manage FF Category Mappings", width=int(18 * scaling_factor), bg=COLORS["pink"], 
            command=lambda: [close_form_with_position(toolbox, conn, cursor, win_id),
                            create_ff_mappings_maint_form(parent, conn, cursor)]).pack(pady=2)
    tk.Button(toolbox, text="Export Transactions to .CSV file", width=int(18 * scaling_factor), bg=COLORS["pink"], 
            command=lambda: [close_form_with_position(toolbox, conn, cursor, win_id),
                            create_export_transactions_form(parent, conn, cursor)]).pack(pady=2)
    tk.Button(toolbox, text="Manage GoCardless Configuration", width=int(18 * scaling_factor), bg=COLORS["pale_green"], 
            command=lambda: [close_form_with_position(toolbox, conn, cursor, win_id),
                            create_gocardless_maint_form(parent, conn, cursor)]).pack(pady=2)
    tk.Button(toolbox, text="Data Import Rules", width=int(18 * scaling_factor), bg=COLORS["pale_green"],
            command=lambda: [close_form_with_position(toolbox, conn, cursor, win_id),
                            create_rules_form(parent, conn, cursor)]).pack(pady=2)
    tk.Button(toolbox, text="Close", width=int(18 * scaling_factor), command=lambda: close_form_with_position(toolbox, conn, cursor, win_id)).pack(pady=5)


