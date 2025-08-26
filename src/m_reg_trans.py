
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
from tkcalendar import Calendar
import logging
import ast

from db import  (fetch_years, fetch_subcategories, fetch_regular_for_year, fetch_category_name, fetch_subcategory_name, fetch_account_full_name,
                delete_transaction, fetch_regular_by_id, fetch_account_full_names, fetch_inc_categories, fetch_exp_categories, fetch_category_id,
                fetch_subcategory_id, fetch_account_id_by_name)
from ui_utils import (COLORS, TEXT_COLORS, open_form_with_position, close_form_with_position)

# Set up logging
logger = logging.getLogger('HA.m_reg_trans')

# Handle DPI scaling for high-resolution displays
try:
    from ctypes import windll
    windll.shcore.SetProcessDpiAwareness(1)
except Exception as e:
    logger.warning(f"Failed to set DPI awareness: {e}")

# Module-level month_names
month_names = {
    1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
    7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec"
}

def create_regular_transactions_maint_form(parent, conn, cursor):       # Win_ID = 3    
    form = tk.Toplevel(parent)
    win_id = 3
    open_form_with_position(form, conn, cursor, win_id, "Regular Transaction Maintenance")
    scaling_factor = parent.winfo_fpixels('1i') / 96  # Get scaling factor (e.g., 2.0 for 200% scaling)
    form.geometry(f"{int(1600 * scaling_factor)}x{int(800 * scaling_factor)}")  # Adjust size
    #form.geometry("1600x800")
    form.resizable(False, False)
    form.configure(bg=COLORS["very_pale_blue"])
    form.grab_set()

    # Variables
    selected_year = tk.StringVar(value=str(datetime.today().year))
    refresh_needed = tk.BooleanVar(value=True)
    check_all_var = tk.BooleanVar(value=True)  # For Select/Clear All
    reg_records = []  # List to hold regular transaction data
    checked_items = {}  # Dict to track checked state: {tree_item: bool}
    tree_items = []

    # Select/Clear All Checkbox
    chk_all = tk.Button(form, text="Select/Clear ALL Rows", font=("Arial", 11), bg=COLORS["very_pale_blue"], 
                        command=lambda: toggle_all_checkboxes()).place(x=int(25 * scaling_factor), y=int(25 * scaling_factor), width=int(180 * scaling_factor))

    # Year Selection
    tk.Label(form, text="Select Year:", font=("Arial", 11), bg=COLORS["very_pale_blue"]).place(x=int(300 * scaling_factor), y=int(30 * scaling_factor))
    year_combo = ttk.Combobox(  form, textvariable=selected_year, values=fetch_years(cursor), 
                                font=("Arial", 11), width=6, state="readonly")
    year_combo.place(x=int(390 * scaling_factor), y=int(25 * scaling_factor))
    year_combo.bind("<<ComboboxSelected>>", lambda e: refresh_needed.set(True))

    # ListView (Treeview)
    columns = ("Sel", "Frequency", "T Day", "T Month", "Trans Type", "Amount", "Description", 
            "Start Date", "Stop Date", "Exp Category", "Exp Sub-Category", "Account From", "Account To", "Flag")
    tree = ttk.Treeview(form, columns=columns, show="headings", height=30, selectmode="browse")
    tree.place(x=int(20 * scaling_factor), y=int(60 * scaling_factor), width=int(1560 * scaling_factor), height=int(675 * scaling_factor))
    scrollbar = tk.Scrollbar(form, orient="vertical", command=tree.yview)
    scrollbar.place(x=int(1580 * scaling_factor), y=int(60 * scaling_factor), height=int(675 * scaling_factor))
    tree.configure(yscrollcommand=scrollbar.set)

    # Column widths and headings
    w40 = int(40 * scaling_factor)
    w50 = int(50 * scaling_factor)
    w60 = int(60 * scaling_factor)
    w80 = int(80 * scaling_factor)
    w100 = int(100 * scaling_factor)
    w150 = int(150 * scaling_factor)
    w250 = int(250 * scaling_factor)
    widths = [w50, w80, w40, w60, w100, w80, w250, w80, w80, w150, w150, w100, w100, w80]
    anchors = ["center", "center", "center", "w", "w", "e", "w", "w", "w", "w", "w", "w", "w", "center"]
    for col, width, anchor in zip(columns, widths, anchors):
        tree.heading(col, text=col, anchor="center")
        tree.column(col, width=width, anchor=anchor)

    # Style for checked rows
    #style = ttk.Style()
    tree.tag_configure("checked", background=COLORS["pale_blue"])  # Use tag_configure for Treeview

    # Buttons
    tk.Button(  form, text="Copy 'Selected Rows ONLY' into Accounts", font=("Arial", 10), width=40,
                command=lambda: copy_selected_to_trans()).place(x=int(50 * scaling_factor), y=int(750 * scaling_factor))
    tk.Button(  form, text="Amend Highlighted Row", font=("Arial", 10), width=25,
                command=lambda: amend_regular()).place(x=int(450 * scaling_factor), y=int(750 * scaling_factor))
    tk.Button(  form, text="New Regular Transaction", font=("Arial", 10), width=25,
                command=lambda: new_regular()).place(x=int(750 * scaling_factor), y=int(750 * scaling_factor))
    tk.Button(  form, text="Close", font=("Arial", 10), width=25,
                command=lambda: close_form_with_position(form, conn, cursor, win_id)).place(x=int(1300 * scaling_factor), y=int(750 * scaling_factor))

    def toggle_all_checkboxes():
        state = check_all_var.get()
        for item in tree_items:
            checked_items[item] = state
            values = list(tree.item(item, "values"))
            values[0] = "☑" if state else "☐"
            tree.item(item, values=values, tags=("checked" if state else "",))
        check_all_var.set(False) if state else check_all_var.set(True)

    def refresh_list():
        nonlocal reg_records
        if not refresh_needed.get():
            return
        year = int(selected_year.get())
        reg_records = fetch_regular_for_year(cursor, year)
        
        if not reg_records:
            if messagebox.askyesno("No Regular Records", 
                                "No Regular records exist for the selected year. Copy from last year?"):
                prev_year = year - 1
                prev_records = fetch_regular_for_year(cursor, prev_year)
                if not prev_records:
                    messagebox.showerror("No Records", "No Regular records exist for last year. Set up manually.")
                else:
                    for rec in prev_records:
                        rec["Reg_Year"] = year
                        rec["Reg_Start"] = 0
                        if rec["Reg_Stop"] == 0:
                            cursor.execute(
                                "INSERT INTO Regular (Reg_Year, Reg_Frequency, Reg_Day, Reg_Month, Reg_Type, "
                                "Reg_Amount, Reg_Desc, Reg_Start, Reg_Stop, Reg_Exp_ID, Reg_ExpSub_ID, "
                                "Reg_Acc_From, Reg_Acc_To, Reg_Query_Flag) "
                                "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                                (year, rec["Reg_Frequency"], rec["Reg_Day"], rec["Reg_Month"], rec["Reg_Type"],
                                rec["Reg_Amount"], rec["Reg_Desc"], 0, 0, rec["Reg_Exp_ID"], rec["Reg_ExpSub_ID"],
                                rec["Reg_Acc_From"], rec["Reg_Acc_To"], rec["Reg_Query_Flag"]))
                    conn.commit()
                    reg_records = fetch_regular_for_year(cursor, year)

        tree.delete(*tree.get_children())
        tree_items.clear()
        checked_items.clear()
        for rec in reg_records:
            try:
                start_date = "None" if rec["Reg_Start"] == 0 else datetime.fromordinal(int(rec["Reg_Start"]) - 1721425).strftime("%d/%m/%Y")
                stop_date = "None" if rec["Reg_Stop"] == 0 else datetime.fromordinal(int(rec["Reg_Stop"]) - 1721425).strftime("%d/%m/%Y")
            except (ValueError, TypeError) as e:
                logger.error(f"Error processing dates for Reg_ID {rec['Reg_ID']}: Start={rec['Reg_Start']}, Stop={rec['Reg_Stop']}, Error={e}")
                start_date = "Invalid"
                stop_date = "Invalid"

            values = (
                "☐",
                {1: "Monthly", 2: "Weekly", 3: "Yearly", 4: "2-Weekly", 5: "4-Weekly"}.get(rec["Reg_Frequency"], ""),
                rec["Reg_Day"],
                "" if rec["Reg_Month"] == 0 else datetime(2023, rec["Reg_Month"], 1).strftime("%B"),
                {1: "Income", 2: "Expenditure", 3: "Transfer"}.get(rec["Reg_Type"], ""),
                f"{rec['Reg_Amount']:.2f}   ",
                rec["Reg_Desc"],
                start_date,
                stop_date,
                fetch_category_name(cursor, rec["Reg_Exp_ID"], year),
                fetch_subcategory_name(cursor, rec["Reg_Exp_ID"], rec["Reg_ExpSub_ID"], year),
                fetch_account_full_name(cursor, rec["Reg_Acc_From"], year),
                fetch_account_full_name(cursor, rec["Reg_Acc_To"], year),
                "Set" if rec["Reg_Query_Flag"] else "-"
            )
            item = tree.insert("", "end", text=str((rec["Reg_ID"], rec["Reg_Acc_From"], rec["Reg_Acc_To"])), values=values, tags=("",))
            tree_items.append(item)
            checked_items[item] = False
        refresh_needed.set(False)

    def copy_selected_to_trans():
        selected = [tree.item(item)["values"] for item in tree_items if checked_items.get(item, False)]
        if not selected:
            messagebox.showerror("Error", "No rows selected for copying.")
            return
        today = datetime.today().toordinal() + 1721425
        year = int(selected_year.get())
        # Parse text tuple correctly
        selected_data = [ast.literal_eval(tree.item(item, "text")) for item in tree_items if checked_items.get(item, False)]
        selected_ids = [data[0] for data in selected_data]  # Reg_IDs
        new_recs = generate_transactions(selected, selected_data, year, today)
        if year < datetime.today().year:
            messagebox.showerror("Error", "Cannot apply Regular transactions to a previous year.")
            return
        
        cursor.execute("SELECT Tr_ID, Tr_Reg_ID, Tr_Day, Tr_Month, Tr_Year FROM Trans WHERE Tr_Year = ? AND Tr_Reg_ID != 0", (year,))
        trans_recs = cursor.fetchall()
        
        ins_count, del_count = 0, 0
        for tr in trans_recs:
            tr_date = datetime(tr[4], tr[3], tr[2]).toordinal() + 1721425
            if tr_date >= today and tr[1] in selected_ids:
                delete_transaction(cursor, conn, tr[0])
                del_count += 1
        
        for rec in new_recs:
            tr_date = datetime(rec["Tr_Year"], rec["Tr_Month"], rec["Tr_Day"]).toordinal() + 1721425
            if tr_date >= today:
                cursor.execute("""
                    INSERT INTO Trans (Tr_Type, Tr_Day, Tr_Month, Tr_Year, Tr_Stat, Tr_Query_Flag, Tr_Amount, 
                                Tr_Desc, Tr_Acc_From, Tr_Acc_To, Tr_Exp_ID, Tr_ExpSub_ID, Tr_Reg_ID)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (rec["Tr_Type"], rec["Tr_Day"], rec["Tr_Month"], rec["Tr_Year"], 1, rec["Tr_Query_Flag"],
                    rec["Tr_Amount"], rec["Tr_Desc"], rec["Tr_Acc_From"], rec["Tr_Acc_To"], 
                    rec["Tr_Exp_ID"], rec["Tr_ExpSub_ID"], rec["Tr_Reg_ID"]))
                ins_count += 1
        conn.commit()
        messagebox.showinfo("Regular Records Processed", f"A total of {ins_count} records created\nand {del_count} old records deleted")
        refresh_needed.set(True)

    def amend_regular():
        selected = tree.selection()
        if not selected:
            messagebox.showerror("Error", "Please select a row to amend.")
            return
        reg_id = ast.literal_eval(tree.item(selected[0], "text"))[0]
        edit_form = tk.Toplevel(form)       # create child form so we can see when it closes
        create_regular_transaction_form(form, edit_form, conn, cursor, reg_id, is_new=False)
        form.wait_window(edit_form)         # pause until child form closes
        refresh_needed.set(True)
        refresh_list()

    def new_regular():
        edit_form = tk.Toplevel(form)       # create child form so we can see when it closes
        create_regular_transaction_form(form, edit_form, conn, cursor, is_new=True)
        form.wait_window(edit_form)         # pause until child form closes
        refresh_needed.set(True)
        refresh_list()

    def on_tree_click(event):
        item = tree.identify_row(event.y)
        if not item:
            return
        col = tree.identify_column(event.x)
        if col == "#1":  # Click on RecID column
            current_state = checked_items.get(item, False)
            new_state = not current_state
            checked_items[item] = new_state
            values = list(tree.item(item, "values"))
            values[0] = "☑" if new_state else "☐"
            tree.item(item, values=values, tags=("checked" if new_state else "",))
        else:  # Highlight for amend
            tree.selection_set(item)

    tree.bind("<Button-1>", on_tree_click)

    # Main loop
    refresh_list()
    while form.winfo_exists():
        if refresh_needed.get():
            refresh_list()
        form.update_idletasks()
        form.update()

def create_regular_transaction_form(parent, form, conn, cursor, curr_rec_id=0, is_new=False):   # Win_ID = 18
    # form = tk.Toplevel(parent)
    win_id = 18
    open_form_with_position(form, conn, cursor, win_id, "Regular Transaction Entry/Edit")
    scaling_factor = parent.winfo_fpixels('1i') / 96  # Get scaling factor (e.g., 2.0 for 200% scaling)
    form.geometry(f"{int(1400 * scaling_factor)}x{int(300 * scaling_factor)}")  # Adjust size
    #form.geometry("1400x300")
    form.resizable(False, False)
    form.configure(bg=COLORS["very_pale_blue"])
    form.grab_set()

    year = int(parent.children["!combobox"].get())  # Get year from parent form
    reg_rec = {"Reg_ID": 0} if is_new else fetch_regular_by_id(cursor, curr_rec_id)
    if not is_new and not reg_rec:
        messagebox.showerror("Error", "A valid record was not found")
        form.destroy()
        return

    # Variables
    frequency_var = tk.StringVar(value="Monthly" if is_new else {1: "Monthly", 2: "Weekly", 3: "Yearly", 4: "2-Weekly", 5: "4-Weekly"}.get(reg_rec["Reg_Frequency"], ""))
    day_var = tk.StringVar(value="1" if is_new else str(reg_rec["Reg_Day"]))
    month_var = tk.StringVar(value="" if is_new else datetime(2023, reg_rec["Reg_Month"], 1).strftime("%B") if reg_rec["Reg_Month"] else "")
    type_var = tk.StringVar(value="Expenditure" if is_new else {1: "Income", 2: "Expenditure", 3: "Transfer"}.get(reg_rec["Reg_Type"], ""))
    amount_var = tk.StringVar(value="" if is_new else f"{reg_rec['Reg_Amount']:.2f}")
    desc_var = tk.StringVar(value="" if is_new else reg_rec["Reg_Desc"])
    start_var = tk.StringVar(value="None" if is_new or reg_rec["Reg_Start"] == 0 else datetime.fromordinal(int(reg_rec["Reg_Start"]) - 1721425).strftime("%d/%m/%Y"))
    stop_var = tk.StringVar(value="None" if is_new or reg_rec["Reg_Stop"] == 0 else datetime.fromordinal(int(reg_rec["Reg_Stop"]) - 1721425).strftime("%d/%m/%Y"))
    exp_cat_var = tk.StringVar(value="" if is_new else fetch_category_name(cursor, reg_rec["Reg_Exp_ID"], year))
    exp_sub_var = tk.StringVar(value="" if is_new else fetch_subcategory_name(cursor, reg_rec["Reg_Exp_ID"], reg_rec["Reg_ExpSub_ID"], year))
    acc_from_var = tk.StringVar(value="" if is_new else fetch_account_full_name(cursor, reg_rec["Reg_Acc_From"], year))
    acc_to_var = tk.StringVar(value="" if is_new else fetch_account_full_name(cursor, reg_rec["Reg_Acc_To"], year))
    flag_var = tk.BooleanVar(value=False if is_new else bool(reg_rec["Reg_Query_Flag"]))

    # Widgets
    tk.Label(form, text="Frequency:", font=("Arial", 10), bg=COLORS["very_pale_blue"]).place(x=int(50 * scaling_factor), y=int(50 * scaling_factor))
    freq_combo = ttk.Combobox(form, textvariable=frequency_var, values=["Monthly", "Weekly", "Yearly", "2-Weekly", "4-Weekly"], state="readonly", font=("Arial", 10))
    freq_combo.place(x=int(180 * scaling_factor), y=int(45 * scaling_factor), width=int(150 * scaling_factor))

    day_label = tk.Label(form, text="Day of Month:", font=("Arial", 10), bg=COLORS["very_pale_blue"])
    day_label.place(x=int(50 * scaling_factor), y=int(100 * scaling_factor))
    day_combo = ttk.Combobox(form, textvariable=day_var, state="readonly", font=("Arial", 10))
    day_combo.place(x=int(180 * scaling_factor), y=int(95 * scaling_factor), width=int(150 * scaling_factor))

    tk.Label(form, text="Month of Year:", font=("Arial", 10), bg=COLORS["very_pale_blue"]).place(x=int(50 * scaling_factor), y=int(150 * scaling_factor))
    month_combo = ttk.Combobox(form, textvariable=month_var, values=[""] + [datetime(2023, m, 1).strftime("%B") for m in range(1, 13)], state="readonly", font=("Arial", 10))
    month_combo.place(x=int(180 * scaling_factor), y=int(145 * scaling_factor), width=int(150 * scaling_factor))

    tk.Label(form, text="Transaction Type:", font=("Arial", 10), bg=COLORS["very_pale_blue"]).place(x=int(50 * scaling_factor), y=int(200 * scaling_factor))
    type_combo = ttk.Combobox(form, textvariable=type_var, values=["Income", "Expenditure", "Transfer"], state="readonly", font=("Arial", 10))
    type_combo.place(x=int(180 * scaling_factor), y=int(195 * scaling_factor), width=int(150 * scaling_factor))

    tk.Label(form, text="Amount:", font=("Arial", 10), bg=COLORS["very_pale_blue"]).place(x=int(400 * scaling_factor), y=int(50 * scaling_factor))
    amount_entry = tk.Entry(form, textvariable=amount_var, font=("Arial", 10))
    amount_entry.place(x=int(500 * scaling_factor), y=int(45 * scaling_factor), width=int(150 * scaling_factor))

    tk.Label(form, text="Description:", font=("Arial", 10), bg=COLORS["very_pale_blue"]).place(x=int(400 * scaling_factor), y=int(100 * scaling_factor))
    desc_entry = tk.Entry(form, textvariable=desc_var, font=("Arial", 10))
    desc_entry.place(x=int(500 * scaling_factor), y=int(95 * scaling_factor), width=int(200 * scaling_factor))

    tk.Label(form, text="Start Date:", font=("Arial", 10), bg=COLORS["very_pale_blue"]).place(x=int(400 * scaling_factor), y=int(150 * scaling_factor))
    start_entry = tk.Entry(form, textvariable=start_var, font=("Arial", 10), state="readonly")
    start_entry.place(x=int(500 * scaling_factor), y=int(145 * scaling_factor), width=int(150 * scaling_factor))
    tk.Button(form, text="Cal", font=("Arial", 10), command=lambda: pick_date(start_entry, form)).place(x=int(660 * scaling_factor), y=int(145 * scaling_factor), 
                                                                                                        width=int(30 * scaling_factor))

    tk.Label(form, text="Stop Date:", font=("Arial", 10), bg=COLORS["very_pale_blue"]).place(x=int(400 * scaling_factor), y=int(200 * scaling_factor))
    stop_entry = tk.Entry(form, textvariable=stop_var, font=("Arial", 10), state="readonly")
    stop_entry.place(x=int(500 * scaling_factor), y=int(195 * scaling_factor), width=int(150 * scaling_factor))
    tk.Button(form, text="Cal", font=("Arial", 10), command=lambda: pick_date(stop_entry, form)).place(x=int(660 * scaling_factor), y=int(195 * scaling_factor), 
                                                                                                       width=int(30 * scaling_factor))

    tk.Label(form, text="Exp Category:", font=("Arial", 10), bg=COLORS["very_pale_blue"]).place(x=int(800 * scaling_factor), y=int(50 * scaling_factor))
    exp_cat_combo = ttk.Combobox(form, textvariable=exp_cat_var, state="readonly", font=("Arial", 10))
    exp_cat_combo.place(x=int(960 * scaling_factor), y=int(45 * scaling_factor), width=int(200 * scaling_factor))

    tk.Label(form, text="Exp Sub-Category:", font=("Arial", 10), bg=COLORS["very_pale_blue"]).place(x=int(800 * scaling_factor), y=int(100 * scaling_factor))
    exp_sub_combo = ttk.Combobox(form, textvariable=exp_sub_var, state="readonly", font=("Arial", 10))
    exp_sub_combo.place(x=int(960 * scaling_factor), y=int(95 * scaling_factor), width=int(200 * scaling_factor))

    tk.Label(form, text="Account From:", font=("Arial", 10), bg=COLORS["very_pale_blue"]).place(x=int(800 * scaling_factor), y=int(150 * scaling_factor))
    acc_from_combo = ttk.Combobox(form, textvariable=acc_from_var, values=fetch_account_full_names(cursor, year), state="readonly", font=("Arial", 10))
    acc_from_combo.place(x=int(960 * scaling_factor), y=int(145 * scaling_factor), width=int(200 * scaling_factor))

    tk.Label(form, text="Account To:", font=("Arial", 10), bg=COLORS["very_pale_blue"]).place(x=int(800 * scaling_factor), y=int(200 * scaling_factor))
    acc_to_combo = ttk.Combobox(form, textvariable=acc_to_var, values=fetch_account_full_names(cursor, year), state="readonly", font=("Arial", 10))
    acc_to_combo.place(x=int(960 * scaling_factor), y=int(195 * scaling_factor), width=int(200 * scaling_factor))

    flag_btn = tk.Button(form, text="Flag OFF", font=("Arial", 10), bg=COLORS["grey"],
                            command=lambda: toggle_flag_btn())
    flag_btn.place(x=int(1250 * scaling_factor), y=int(50 * scaling_factor), width=int(100 * scaling_factor), height=int(35 * scaling_factor))
    if not is_new and reg_rec["Reg_Query_Flag"]:
        flag_var.set(True)
        flag_btn.config(text="Flag ON", bg=COLORS["yellow"])
    else:
        flag_var.set(False)
        flag_btn.config(text="Flag OFF", bg=COLORS["grey"])

    tk.Button(form, text="Save", font=("Arial", 10), command=lambda: save_record()).place(x=int(250 * scaling_factor), y=int(250 * scaling_factor), width=int(150 * scaling_factor))
    tk.Button(form, text="Cancel", font=("Arial", 10), command=lambda: close_form()).place(x=int(850 * scaling_factor), y=int(250 * scaling_factor), width=int(150 * scaling_factor))
    if not is_new:
        tk.Button(form, text="DELETE", font=("Arial", 10), bg="red", command=lambda: delete_record()).place(x=int(1250 * scaling_factor), y=int(200 * scaling_factor), 
                                                                                                            width=int(100 * scaling_factor))
    def toggle_flag_btn():
        if flag_var.get():
            flag_var.set(False)
            flag_btn.config(text="Flag OFF", bg=COLORS["grey"])
        else:
            flag_var.set(True)
            flag_btn.config(text="Flag ON", bg=COLORS["yellow"])



    # Dynamic Behavior
    def update_frequency(*args):
        freq = frequency_var.get()
        if freq in ["Weekly", "2-Weekly", "4-Weekly"]:
            day_label.config(text="Day of Week:")
            day_combo.config(values=["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"])
            day_var.set("Monday" if day_var.get().isdigit() else day_var.get())
            month_combo.config(state="disabled")
            month_var.set("")
            start_entry.config(state="normal")
            stop_entry.config(state="normal")
        else:
            day_label.config(text="Day of Month:")
            day_combo.config(values=[str(i) for i in range(1, 29)])
            day_var.set("1" if not day_var.get().isdigit() else day_var.get())
            month_combo.config(state="normal" if freq == "Yearly" else "disabled")
            month_var.set("January" if freq == "Yearly" and not month_var.get() else "" if freq != "Yearly" else month_var.get())
            start_entry.config(state="disabled" if freq == "Yearly" else "normal")
            stop_entry.config(state="disabled" if freq == "Yearly" else "normal")
            if freq == "Yearly":
                start_var.set("None")
                stop_var.set("None")
        update_type()

    def update_type(*args):
        ttype = type_var.get()
        if ttype == "Income":
            acc_from_combo.config(state="disabled")
            acc_from_var.set("")
            acc_to_combo.config(state="normal")
            exp_cat_combo.config(state="normal", values=fetch_inc_categories(cursor, year))
            exp_sub_combo.config(state="normal")
        elif ttype == "Expenditure":
            acc_from_combo.config(state="normal")
            acc_to_combo.config(state="disabled")
            acc_to_var.set("")
            exp_cat_combo.config(state="normal", values=fetch_exp_categories(cursor, year))
            exp_sub_combo.config(state="normal")
        else:  # Transfer
            acc_from_combo.config(state="normal")
            acc_to_combo.config(state="normal")
            exp_cat_combo.config(state="disabled")
            exp_sub_combo.config(state="disabled")
            exp_cat_var.set("")
            exp_sub_var.set("")
        update_exp_cat()

    def update_exp_cat(*args):
        cat = exp_cat_var.get()
        if cat and cat != "Ignore":
            exp_id = fetch_category_id(cursor, cat, year)
            exp_sub_combo.config(values=fetch_subcategories(cursor, exp_id, year), state="normal")
            if not exp_sub_var.get():
                exp_sub_var.set(fetch_subcategories(cursor, exp_id, year)[0] if fetch_subcategories(cursor, exp_id, year) else "")
        else:
            exp_sub_combo.config(state="disabled", values=[""])
            exp_sub_var.set("")

    def pick_date(entry, parent_form):
        cal = tk.Toplevel(parent_form)
        cal.title("Select Date")
        cal.grab_set()
        cal.attributes("-topmost", True)
        cal.withdraw()  # Hide the window initially

        # Get parent_form's position and size
        parent_x = parent_form.winfo_x()
        parent_y = parent_form.winfo_y()
        parent_width = parent_form.winfo_width()
        parent_height = parent_form.winfo_height()

        # Set initial date, handling invalid formats
        try:
            current = datetime.strptime(entry.get(), "%d/%m/%Y") if entry.get() != "None" else datetime.today()
        except ValueError:
            current = datetime.today()  # Fallback to today if date is invalid
            if entry.get() != "None":
                messagebox.showwarning("Warning", f"Invalid date format: {entry.get()}. Using today's date. Please use DD/MM/YYYY (e.g., 01/04/2025).")
                entry.delete(0, tk.END)
                entry.insert(0, current.strftime("%d/%m/%Y"))

        # Create the calendar widget
        current = datetime.strptime(entry.get(), "%d/%m/%Y") if entry.get() != "None" else datetime.today()
        calendar = Calendar(cal, selectmode="day", year=current.year, month=current.month, day=current.day, date_pattern="dd/mm/yyyy")
        
        # Pack the calendar and buttons
        calendar.pack(pady=10)
        tk.Button(cal, text="   OK   ", command=lambda: [entry.delete(0, tk.END), entry.insert(0, calendar.get_date()), 
                                                         cal.destroy()]).pack(padx=int(20 * scaling_factor), pady=int(10 * scaling_factor))
        tk.Button(cal, text=" Clear ", command=lambda: [entry.delete(0, tk.END), entry.insert(0, "None"), 
                                                        cal.destroy()]).pack(padx=int(20 * scaling_factor), pady=int(10 * scaling_factor))

        # Center the calendar
        cal.update_idletasks()
        cal_width = max(cal.winfo_width(), int(300 * scaling_factor))  # Minimum width for calendar
        cal_height = max(cal.winfo_height(), int(320 * scaling_factor))  # Minimum height to fit calendar + buttons
        center_x = parent_x + (parent_width - cal_width) // 2
        center_y = parent_y + (parent_height - cal_height) // 2
        cal.geometry(f"{cal_width}x{cal_height}+{center_x}+{center_y}")

        cal.deiconify()  # Show the window after positioning

    def save_record():
        reg = {}
        reg["Reg_Year"] = year
        freq_map = {"Monthly": 1, "Weekly": 2, "Yearly": 3, "2-Weekly": 4, "4-Weekly": 5}
        reg["Reg_Frequency"] = freq_map.get(frequency_var.get(), 0)
        reg["Reg_Day"] = int(day_var.get()) if day_var.get().isdigit() else {"Sunday": 1, "Monday": 2, "Tuesday": 3, "Wednesday": 4, "Thursday": 5, "Friday": 6, "Saturday": 7}.get(day_var.get(), 0)
        reg["Reg_Month"] = 0 if not month_var.get() else [datetime.strptime(m, "%B").month for m in [month_var.get()]][0]
        type_map = {"Income": 1, "Expenditure": 2, "Transfer": 3}
        reg["Reg_Type"] = type_map.get(type_var.get(), 0)
        try:
            reg["Reg_Amount"] = float(amount_var.get())
            if not (0 <= reg["Reg_Amount"] <= 99999.99):
                raise ValueError
        except ValueError:
            messagebox.showerror("Error", "Amount must be a number between 0 and 99999.99")
            return
        
        reg["Reg_Desc"] = desc_var.get()
        try:
            reg["Reg_Start"] = 0 if start_var.get() == "None" else datetime.strptime(start_var.get(), "%d/%m/%Y").toordinal() + 1721425
        except ValueError as e:
            messagebox.showerror("Error", f"Invalid Start Date format: {start_var.get()}. Use DD/MM/YYYY (e.g., 01/04/2025)")
            return
        try:
            reg["Reg_Stop"] = 0 if stop_var.get() == "None" else datetime.strptime(stop_var.get(), "%d/%m/%Y").toordinal() + 1721425
        except ValueError as e:
            messagebox.showerror("Error", f"Invalid Stop Date format: {stop_var.get()}. Use DD/MM/YYYY (e.g., 01/04/2025)")
            return
        reg["Reg_Exp_ID"] = fetch_category_id(cursor, exp_cat_var.get(), year) if exp_cat_var.get() else 0
        reg["Reg_ExpSub_ID"] = fetch_subcategory_id(cursor, reg["Reg_Exp_ID"], exp_sub_var.get(), year) if exp_sub_var.get() else 0
        reg["Reg_Acc_From"] = fetch_account_id_by_name(cursor, acc_from_var.get(), year) if acc_from_var.get() else 0
        reg["Reg_Acc_To"] = fetch_account_id_by_name(cursor, acc_to_var.get(), year) if acc_to_var.get() else 0
        reg["Reg_Query_Flag"] = 1 if flag_var.get() else 0

        # Validation
        if reg["Reg_Frequency"] == 0 or reg["Reg_Day"] == 0 or reg["Reg_Type"] == 0:
            messagebox.showerror("Error", "Please select valid Frequency, Day, and Type")
            return
        if reg["Reg_Frequency"] in [2, 4, 5] and reg["Reg_Start"] == 0:
            messagebox.showerror("Error", "Weekly, 2-Weekly, and 4-Weekly require a Start Date")
            return
        if reg["Reg_Type"] in [1, 2] and not exp_cat_var.get():
            messagebox.showerror("Error", "Please choose a Category")
            return
        if reg["Reg_Type"] in [2, 3] and not acc_from_var.get():
            messagebox.showerror("Error", "Please choose an Account From")
            return
        if reg["Reg_Type"] in [1, 3] and not acc_to_var.get():
            messagebox.showerror("Error", "Please choose an Account To")
            return

        if is_new:
            cursor.execute("""
                INSERT INTO Regular (Reg_Year, Reg_Frequency, Reg_Day, Reg_Month, Reg_Type, Reg_Amount, Reg_Desc, 
                                    Reg_Start, Reg_Stop, Reg_Exp_ID, Reg_ExpSub_ID, Reg_Acc_From, Reg_Acc_To, Reg_Query_Flag)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (reg["Reg_Year"], reg["Reg_Frequency"], reg["Reg_Day"], reg["Reg_Month"], reg["Reg_Type"], reg["Reg_Amount"],
                reg["Reg_Desc"], reg["Reg_Start"], reg["Reg_Stop"], reg["Reg_Exp_ID"], reg["Reg_ExpSub_ID"],
                reg["Reg_Acc_From"], reg["Reg_Acc_To"], reg["Reg_Query_Flag"]))
        else:
            cursor.execute("""
                UPDATE Regular SET Reg_Frequency=?, Reg_Day=?, Reg_Month=?, Reg_Type=?, Reg_Amount=?, Reg_Desc=?,
                            Reg_Start=?, Reg_Stop=?, Reg_Exp_ID=?, Reg_ExpSub_ID=?, Reg_Acc_From=?, Reg_Acc_To=?, Reg_Query_Flag=?
                WHERE Reg_ID=? AND Reg_Year=?
            """, (reg["Reg_Frequency"], reg["Reg_Day"], reg["Reg_Month"], reg["Reg_Type"], reg["Reg_Amount"], reg["Reg_Desc"],
                reg["Reg_Start"], reg["Reg_Stop"], reg["Reg_Exp_ID"], reg["Reg_ExpSub_ID"], reg["Reg_Acc_From"],
                reg["Reg_Acc_To"], reg["Reg_Query_Flag"], curr_rec_id, year))
        conn.commit()
        close_form_with_position(form, conn, cursor, win_id)
        return 1

    def delete_record():
        response = messagebox.askyesnocancel("Delete Template & Transactions?",
                                            "YES = Delete both Template and Transactions,\nNO = Delete Template ONLY\n\nCompleted transactions will not be deleted")
        if response is not None:
            if response:  # Yes
                cursor.execute("DELETE FROM Trans WHERE Tr_Reg_ID=? AND Tr_Stat=1", (curr_rec_id,))
            cursor.execute("DELETE FROM Regular WHERE Reg_ID=?", (curr_rec_id,))
            conn.commit()
            close_form_with_position(form, conn, cursor, win_id)
            return 1

    def close_form():
        close_form_with_position(form, conn, cursor, win_id)
        return 0

    # Bindings
    frequency_var.trace("w", update_frequency)
    type_var.trace("w", update_type)
    exp_cat_var.trace("w", update_exp_cat)
    update_frequency()  # Initial setup

def generate_transactions(selected, selected_data, year, today):
    new_recs = []
    freq_map = {"Monthly": 1, "Weekly": 2, "Yearly": 3, "2-Weekly": 4, "4-Weekly": 5}
    type_map = {"Income": 1, "Expenditure": 2, "Transfer": 3}
    month_map = {v: k for k, v in enumerate([""] + [datetime(2023, m, 1).strftime("%B") for m in range(1, 13)], 1)}

    for rec, (reg_id, acc_from, acc_to) in zip(selected, selected_data):
        freq = freq_map[rec[1]]
        day = int(rec[2])
        month = month_map.get(rec[3], 0)
        tr_type = type_map[rec[4]]
        amount = float(rec[5].strip())
        desc = rec[6]
        start = 0 if rec[7] == "None" else datetime.strptime(rec[7], "%d/%m/%Y").toordinal() + 1721425
        stop = 0 if rec[8] == "None" else datetime.strptime(rec[8], "%d/%m/%Y").toordinal() + 1721425
        exp_id = int(rec[9]) if rec[9].isdigit() else 0
        exp_sub_id = int(rec[10]) if rec[10].isdigit() else 0
        acc_to = 0 if tr_type == 2 else acc_to  # Force 0 for expenditure
        acc_from = 0 if tr_type == 1 else acc_from  # Force 0 for income
        flag = 1 if rec[13] == "Set" else 0

        if freq == 1:  # Monthly
            start_m = 1 if start == 0 else max(1, datetime.fromordinal(start - 1721425).month)
            stop_m = 12 if stop == 0 else min(12, datetime.fromordinal(stop - 1721425).month)
            for m in range(start_m, stop_m + 1):
                new_recs.append({"Tr_Type": tr_type, "Tr_Day": day, "Tr_Month": m, "Tr_Year": year,
                                "Tr_Amount": amount, "Tr_Desc": desc, "Tr_Exp_ID": exp_id,
                                "Tr_ExpSub_ID": exp_sub_id, "Tr_Acc_From": acc_from, "Tr_Acc_To": acc_to,
                                "Tr_Query_Flag": flag, "Tr_Reg_ID": reg_id})
        elif freq in [2, 4, 5]:  # Weekly, 2-Weekly, 4-Weekly
            step = {2: 7, 4: 14, 5: 28}[freq]
            start_d = start if start else today
            stop_d = stop if stop else datetime(year, 12, 31).toordinal() + 1721425
            d = start_d
            while d <= stop_d:
                dt = datetime.fromordinal(d - 1721425)
                new_recs.append({"Tr_Type": tr_type, "Tr_Day": dt.day, "Tr_Month": dt.month, "Tr_Year": year,
                                "Tr_Amount": amount, "Tr_Desc": desc, "Tr_Exp_ID": exp_id,
                                "Tr_ExpSub_ID": exp_sub_id, "Tr_Acc_From": acc_from, "Tr_Acc_To": acc_to,
                                "Tr_Query_Flag": flag, "Tr_Reg_ID": reg_id})
                d += step
        elif freq == 3:  # Yearly
            new_recs.append({"Tr_Type": tr_type, "Tr_Day": day, "Tr_Month": month, "Tr_Year": year,
                            "Tr_Amount": amount, "Tr_Desc": desc, "Tr_Exp_ID": exp_id,
                            "Tr_ExpSub_ID": exp_sub_id, "Tr_Acc_From": acc_from, "Tr_Acc_To": acc_to,
                            "Tr_Query_Flag": flag, "Tr_Reg_ID": reg_id})
    return new_recs




