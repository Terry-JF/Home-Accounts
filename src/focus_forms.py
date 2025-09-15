
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import logging
from db import  (fetch_years, fetch_categories, fetch_subcategories, fetch_actuals)
from ui_utils import (open_form_with_position, close_form_with_position, sc)
from config import COLORS
import config

# Set up logging
logger = logging.getLogger('HA.focus_forms')

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

def create_summary_form(parent, conn, cursor):                          # Win_ID = 4 
    form = tk.Toplevel(parent)
    win_id = 4
    open_form_with_position(form, conn, cursor, win_id, "Summary")
    form.geometry(f"{sc(1675)}x{sc(1020)}")  # Adjust size
    #form.geometry("1675x1020")
    form.resizable(False, False)
    form.configure(bg=config.master_bg)
    form.grab_set()

    # Variables
    selected_year = tk.StringVar(value=str(datetime.today().year))
    refresh_needed = tk.BooleanVar(value=True)
    selected_row = tk.StringVar()
    is_expanded = tk.BooleanVar(value=False)  # Track expansion state

    # Year Selection
    tk.Label(form, text="Year:", font=(config.ha_button), bg=config.master_bg).place(x=sc(20), y=sc(20))
    year_combo = ttk.Combobox(  form, textvariable=selected_year, values=fetch_years(cursor), 
                                font=(config.ha_button), width=6, state="readonly")
    year_combo.place(x=sc(60), y=sc(20))
    year_combo.bind("<<ComboboxSelected>>", lambda e: refresh_needed.set(True))

    # Treeview
    columns = ("Category", "Bfwd", "JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC", "TOTALS", "AVG/Month")
    tree = ttk.Treeview(form, columns=columns, show="headings", height=sc(40), selectmode="browse")
    tree.place(x=sc(20), y=sc(100), width=sc(1635), height=sc(890))
    tree.tag_configure("parent", background=COLORS["act_but_bg"], font=(config.ha_button))  # Arial 11 for parent rows
    tree.tag_configure("child", font=(config.ha_normal))  # Arial 10 for child rows

    widths = [sc(200)] + [sc(85)] * 13 + [sc(100), sc(100)]
    for col, width in zip(columns, widths):
        tree.heading(col, text=col, anchor="center")
        tree.column(col, width=width, anchor="e")
    tree.column("Category", anchor="w")
    tree.column("#0", width=sc(5), anchor="center")

    # Buttons
    expand_button = tk.Button(  form, text="Expand All", width=20, font=(config.ha_button), command=lambda: toggle_treeview())
    expand_button.place(x=sc(20), y=sc(60))
    
    month_buttons_x = sc(325)
    for i, month in enumerate(["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]):
        tk.Button(form, text=month, font=(config.ha_normal), bg=COLORS["act_but_bg"],
                    command=lambda m=i+1: show_monthly_focus(m)).place(x=month_buttons_x + i * sc(93), y=sc(60), width=sc(85))

    drill_down_button = tk.Button(  form, text="Drill Down", width=20, font=(config.ha_button), command=lambda: show_drill_down(), state="disabled")
    drill_down_button.place(x=sc(1480), y=sc(20))

    tk.Button(form, text="Close", width=20, font=(config.ha_button), bg=COLORS["exit_but_bg"],
                command=lambda: close_form_with_position(form, conn, cursor, win_id)).place(x=sc(1480), y=sc(60))

    # User guide
    user_guide=tk.Label(form, text="(double-click a parent row to open/close that category)", font=(config.ha_note), background=config.master_bg)
    user_guide.place(x=sc(20), y=sc(995))

    def toggle_treeview():
        if is_expanded.get():
            # Fold: Close all parent rows to hide children
            for item in tree.get_children():
                tree.item(item, open=False)
            expand_button.config(text="Expand All")
            is_expanded.set(False)
        else:
            # Expand: Open all parent and child rows
            for item in tree.get_children():
                tree.item(item, open=True)
                for child in tree.get_children(item):
                    tree.item(child, open=True)
            expand_button.config(text="Fold All")
            is_expanded.set(True)

    def fetch_excluded_actuals(pid, cid, tr_type, month, year):
        # Fetch sum of Tr_Amount for excluded/uncoded transactions, excluding transfers
        query = """
            SELECT COALESCE(SUM(Tr_Amount), 0.0)
            FROM Trans
            WHERE Tr_Year=? AND Tr_Month=? AND Tr_Exp_ID=? AND Tr_ExpSub_ID=? AND Tr_Type=?
        """
        cursor.execute(query, (year, month, pid, cid, tr_type))
        return cursor.fetchone()[0] or 0.0

    def fetch_account_balance(acc_ids, month, year):
        # Fetch EOM balance for given Acc_IDs up to specified month
        acc_ids = [aid for aid in acc_ids if cursor.execute("SELECT 1 FROM Account WHERE Acc_ID=? AND Acc_Year=?", (aid, year)).fetchone()]
        if not acc_ids:
            return 0.0
        month_fields = ["Acc_Jan", "Acc_Feb", "Acc_Mar", "Acc_Apr", "Acc_May", "Acc_Jun",
                        "Acc_Jul", "Acc_Aug", "Acc_Sep", "Acc_Oct", "Acc_Nov", "Acc_Dec"]
        fields = ["Acc_Open"] + month_fields[:month]
        query = """
            SELECT COALESCE(SUM({}), 0.0)
            FROM Account
            WHERE Acc_Year=? AND Acc_ID IN ({})
        """.format(" + ".join(fields), ",".join("?" * len(acc_ids)))
        params = [year] + list(acc_ids)
        cursor.execute(query, params)
        result = cursor.fetchone()[0] or 0.0
        return result

    def fetch_bfwd_balance(acc_ids, year):
        # Fetch BFWD balance (Acc_Open for Jan 1 of current year)
        acc_ids = [aid for aid in acc_ids if cursor.execute("SELECT 1 FROM Account WHERE Acc_ID=? AND Acc_Year=?", (aid, year)).fetchone()]
        if not acc_ids:
            return 0.0
        query = """
            SELECT COALESCE(SUM(Acc_Open), 0.0)
            FROM Account
            WHERE Acc_Year=? AND Acc_ID IN ({})
        """.format(",".join("?" * len(acc_ids)))
        params = [year] + list(acc_ids)
        cursor.execute(query, params)
        result = cursor.fetchone()[0] or 0.0
        return result

    def fetch_transactions(year, pid, cid):
        # Fetch transactions for the given year, PID, and CID, sorted by date
        query = """
            SELECT Tr_Day, Tr_Month, Tr_Desc, Tr_Amount, Tr_Query_Flag, Tr_ID, Tr_Stat
            FROM Trans
            WHERE Tr_Year=? AND Tr_Exp_ID=? AND Tr_ExpSub_ID=? AND Tr_Type IN (1, 2)
            ORDER BY Tr_Month, Tr_Day
        """
        cursor.execute(query, (year, pid, cid))
        return cursor.fetchall()

    def show_drill_down():
        selected = tree.selection()
        if not selected:
            return
        item = selected[0]
        if not item.startswith("child_"):
            return
        try:
            _, pid, cid = item.split("_")
            pid, cid = int(pid), int(cid)
        except ValueError:
            return
        year = int(selected_year.get())

        dd_form = tk.Toplevel(form)
        open_form_with_position(dd_form, conn, cursor, 19, "Transactions making up this row of the Summary")
        dd_form.geometry(f"{sc(600)}x{sc(935)}")
        dd_form.resizable(False, False)
        dd_form.configure(bg=config.master_bg)
        dd_form.grab_set()
        dd_form.lift()
        dd_form.focus_set()

        dd_columns = ("Day", "Month", "Description", "Amount")
        dd_tree = ttk.Treeview(dd_form, columns=dd_columns, show="headings", height=40)
        dd_tree.place(x=sc(20), y=sc(10), width=sc(550), height=sc(860))
        scrollbar = ttk.Scrollbar(dd_form, orient="vertical", command=dd_tree.yview)
        scrollbar.place(x=sc(570), y=sc(10), height=sc(860))
        dd_tree.configure(yscrollcommand=scrollbar.set)
        
        dd_tree.tag_configure("total", background=COLORS["act_but_bg"], font=(config.ha_normal))
        dd_tree.tag_configure("marked", background=COLORS["flag_mk_bg"], font=(config.ha_normal))
        dd_tree.tag_configure("forecast", foreground=COLORS["forecast_tx"], font=(config.ha_normal))
        dd_tree.tag_configure("processing", foreground=COLORS["pending_tx"], font=(config.ha_normal))
        dd_tree.tag_configure("complete", foreground=COLORS["complete_tx"], font=(config.ha_normal))

        dd_widths = [20, 20, 350, 70]
        dd_anchors = ["center", "center", "w", "e"]
        for col, width, anchor in zip(dd_columns, dd_widths, dd_anchors):
            dd_tree.heading(col, text=col, anchor="center")
            dd_tree.column(col, width=width, anchor=anchor) 

        transactions = fetch_transactions(year, pid, cid)
        current_month = None
        month_total = 0.0
        month_items = []

        for tr_day, tr_month, tr_desc, tr_amount, tr_query_flag, tr_id, tr_stat in transactions:
            if current_month is None:
                current_month = tr_month
            if tr_month != current_month:
                if month_items:
                    dd_tree.insert("", "end", values=("", "", f"{month_names[current_month]} Total", f"{month_total:,.2f}  "),
                        tags=("total",))
                current_month = tr_month
                month_total = 0.0
                month_items = []
            month_total += tr_amount
            tags = []
            if tr_query_flag & 8:
                tags.append("marked")
            if tr_stat == 1:
                tags.append("forecast")
            elif tr_stat == 2:
                tags.append("processing")
            elif tr_stat == 3:
                tags.append("complete")
            dd_tree.insert("", "end", values=(tr_day, month_names[tr_month], tr_desc, f"{tr_amount:,.2f}  "),
                tags=tags, iid=f"tr_{tr_id}")
            month_items.append((tr_day, tr_month, tr_desc, tr_amount, tr_query_flag, tr_id, tr_stat))

        if month_items:
            dd_tree.insert("", "end", values=("", "", f"{month_names[current_month]} Total", f"{month_total:,.2f}  "),
                tags=("total",))

        tk.Button(  dd_form, text="Close", font=(config.ha_button), bg=COLORS["exit_but_bg"],
                    command=lambda: close_form_with_position(dd_form, conn, cursor, 19)).place(x=sc(450), y=sc(890), width=sc(100))
        tk.Button(  dd_form, text="Mark (space)", font=(config.ha_button),
                    command=lambda: toggle_mark()).place(x=sc(50), y=sc(890), width=sc(100))

        def toggle_mark():
            selected = dd_tree.selection()
            if not selected:
                messagebox.showwarning("Warning", "Please select a transaction to mark.")
                return
            item = selected[0]
            if not item.startswith("tr_"):
                messagebox.showwarning("Warning", "Please select a transaction, not a totals row.")
                return
            tr_id = int(item.replace("tr_", ""))
            cursor.execute("SELECT Tr_Query_Flag, Tr_Stat FROM Trans WHERE Tr_ID=?", (tr_id,))
            current_flag, tr_stat = cursor.fetchone()
            new_flag = current_flag | 8 if not (current_flag & 8) else current_flag & ~8
            cursor.execute("UPDATE Trans SET Tr_Query_Flag=? WHERE Tr_ID=?", (new_flag, tr_id))
            conn.commit()
            tr_data = cursor.execute("SELECT Tr_Day, Tr_Month, Tr_Desc, Tr_Amount FROM Trans WHERE Tr_ID=?", (tr_id,)).fetchone()
            tags = []
            if new_flag & 8:
                tags.append("marked")
            if tr_stat == 1:
                tags.append("forecast")
            elif tr_stat == 2:
                tags.append("processing")
            elif tr_stat == 3:
                tags.append("complete")
            dd_tree.item(item, values=(tr_data[0], month_names[tr_data[1]], tr_data[2], f"{tr_data[3]:,.2f}  "), tags=tags)

        def on_spacebar(event):
            toggle_mark()

        dd_tree.bind("<space>", on_spacebar)

    def refresh_data():
        if not refresh_needed.get():
            return
        year = int(selected_year.get())
        #prev_year = year - 1
        tree.delete(*tree.get_children())

        def populate_category(cat_pid, cat_desc, is_income):
            subcats = fetch_subcategories(cursor, cat_pid, year)

            # Parent row: Sum actuals across subcategories
            monthly_actuals = [0.0] * 12
            if subcats:
                for cid, _ in subcats:
                    for m in range(1, 13):
                        actual = fetch_actuals(cursor, cat_pid, cid, m, year)
                        monthly_actuals[m-1] += actual
            else:
                for m in range(1, 13):
                    actual = fetch_actuals(cursor, cat_pid, 0, m, year)
                    monthly_actuals[m-1] = actual

            total = sum(monthly_actuals)
            avg_month = total / 12 if total else 0.0

            values = [cat_desc] + [""] + ["" if m == 0.0 else f"{m:,.2f}  " for m in monthly_actuals] + \
                    ["" if total == 0.0 else f"{total:,.2f}  ", "" if avg_month == 0.0 else f"{avg_month:,.2f}  "]
            parent = tree.insert("", "end", text="", values=values, open=False, tags=("parent",))

            # Child rows with description in PID/CID column
            for cid, sub_desc in subcats:
                sub_monthly = [fetch_actuals(cursor, cat_pid, cid, m, year) for m in range(1, 13)]
                sub_total = sum(sub_monthly)
                sub_avg = sub_total / 12 if sub_total else 0.0
                sub_values = [f"   {sub_desc}"] + [""] + ["" if m == 0.0 else f"{m:,.2f}  " for m in sub_monthly] + \
                            ["" if sub_total == 0.0 else f"{sub_total:,.2f}  ", "" if sub_avg == 0.0 else f"{sub_avg:,.2f}  "]
                tree.insert(parent, "end", text="", values=sub_values, iid=f"child_{cat_pid}_{cid}", tags=("child",))

            return parent

        # Populate Income and Expense categories
        income_cats = fetch_categories(cursor, year, is_income=True)
        expense_cats = fetch_categories(cursor, year, is_income=False)

        income_total = [0.0] * 12
        for pid, desc in income_cats:
            parent = populate_category(pid, desc, True)
            for m in range(1, 13):
                income_total[m-1] += sum(fetch_actuals(cursor, pid, cid, m, year) 
                    for cid, _ in fetch_subcategories(cursor, pid, year) or [(0, "")])

        last_pid = max(pid for pid, _ in expense_cats if pid != 99) if expense_cats else None
        misc_parent = None
        for pid, desc in expense_cats:
            if pid != 99:
                parent = populate_category(pid, desc, False)
            if pid == last_pid:
                misc_parent = parent

        # Total Expenditure
        if misc_parent:
            monthly_expenditure = [0.0] * 12
            for pid, _ in expense_cats:
                if pid != 99:
                    subcats = fetch_subcategories(cursor, pid, year)
                    for cid, _ in subcats:
                        for m in range(1, 13):
                            monthly_expenditure[m-1] += fetch_actuals(cursor, pid, cid, m, year)

            exp_total = sum(monthly_expenditure)
            exp_avg = exp_total / 12 if exp_total else 0.0
            tree.insert("", "end", text="", values=["Total Expenditure"] + [""] + 
                        ["" if m == 0.0 else f"{m:,.2f}  " for m in monthly_expenditure] + 
                        ["" if exp_total == 0.0 else f"{exp_total:,.2f}  ", "" if exp_avg == 0.0 else f"{exp_avg:,.2f}  "],
                        open=False, tags=("parent",))

        # Blank separator row
        tree.insert("", "end", text="", values=[""] * 16)

        # Excluded Transactions parent row
        excl_parent = tree.insert("", "end", text="", values=["Excluded Transactions"] + [""] * 15, 
                        open=False, tags=("parent",))
        # Excluded Transactions child rows
        for desc, pid, tr_type in [
            ("Ignored Income", 99, 1), ("Ignored Expenditure", 99, 2),
            ("Uncoded Income", 0, 1), ("Uncoded Expenditure", 0, 2)
        ]:
            monthly = [fetch_excluded_actuals(pid, 0, tr_type, m, year) for m in range(1, 13)]
            total = sum(monthly)
            values = [f"   {desc}"] + [""] + ["" if m == 0.0 else f"{m:,.2f}  " for m in monthly] + \
                    ["" if total == 0.0 else f"{total:,.2f}  ", ""]
            tree.insert(excl_parent, "end", text="", values=values, tags=("child",))

        # Blank separator row
        tree.insert("", "end", text="", values=[""] * 16)

        # End of Month Analysis parent row
        eom_parent = tree.insert("", "end", text="", values=["End of Month Analysis"] + [""] * 15, 
                        open=False, tags=("parent",))
        # End of Month Analysis child rows
        monthly_inc_exp = [income_total[m-1] - monthly_expenditure[m-1] for m in range(1, 13)]
        inc_exp_total = sum(monthly_inc_exp)

        cash_balances = [fetch_account_balance(range(1, 7), m, year) for m in range(1, 13)]
        cash_bfwd = fetch_bfwd_balance(range(1, 7), year)
        cash_changes = [cash_balances[0] - cash_bfwd] + [cash_balances[m] - cash_balances[m-1] for m in range(1, 12)]
        
        credit_balances = [fetch_account_balance(range(7, 13), m, year) for m in range(1, 13)]
        credit_bfwd = fetch_bfwd_balance(range(7, 13), year)
        credit_changes = [credit_balances[0] - credit_bfwd] + [credit_balances[m] - credit_balances[m-1] for m in range(1, 12)]
        
        invest_balances = [fetch_account_balance([13], m, year) for m in range(1, 13)]
        invest_bfwd = fetch_bfwd_balance([13], year)
        invest_changes = [invest_balances[0] - invest_bfwd] + [invest_balances[m] - invest_balances[m-1] for m in range(1, 12)]
        
        loan_balances = [fetch_account_balance([14], m, year) for m in range(1, 13)]
        loan_bfwd = fetch_bfwd_balance([14], year)
        loan_changes = [loan_balances[0] - loan_bfwd] + [loan_balances[m] - loan_balances[m-1] for m in range(1, 12)]
        
        net_positions = [cash_balances[m] + invest_balances[m] + credit_balances[m] + loan_balances[m] for m in range(12)]
        net_bfwd = cash_bfwd + credit_bfwd + invest_bfwd + loan_bfwd
        net_changes = [net_positions[0] - net_bfwd] + [net_positions[m] - net_positions[m-1] for m in range(1, 12)]

        eom_rows = [
            ("Monthly Inc-Exp", False, monthly_inc_exp, True, inc_exp_total, False),
            ("", False, [], False, 0.0, False),
            ("EOM Cash Balance", True, cash_balances, False, 0.0, False),
            ("EOM Cash Change", False, cash_changes, True, sum(cash_changes), False),
            ("", False, [], False, 0.0, False),
            ("EOM Credit Balance", True, credit_balances, False, 0.0, False),
            ("EOM Credit Change", False, credit_changes, True, sum(credit_changes), False),
            ("", False, [], False, 0.0, False),
            ("EOM Investments", True, invest_balances, False, 0.0, False),
            ("EOM Invest Change", False, invest_changes, True, sum(invest_changes), False),
            ("", False, [], False, 0.0, False),
            ("EOM Loans Balance", True, loan_balances, False, 0.0, False),
            ("EOM Loans Change", False, loan_changes, True, sum(loan_changes), False),
            ("", False, [], False, 0.0, False),
            ("Net Position", True, net_positions, False, 0.0, False),
            ("Net Position Change", False, net_changes, True, sum(net_changes), False)
        ]

        bfwd_map = {
            "EOM Cash Balance": cash_bfwd,
            "EOM Credit Balance": credit_bfwd,
            "EOM Investments": invest_bfwd,
            "EOM Loans Balance": loan_bfwd,
            "Net Position": net_bfwd
        }

        for desc, use_bfwd, monthly, use_totals, total, use_avg in eom_rows:
            bfwd = ""
            if use_bfwd and desc in bfwd_map:
                bfwd = "" if bfwd_map[desc] == 0.0 else f"{bfwd_map[desc]:,.2f}  "

            values = [f"   {desc}" if desc else ""] + [bfwd] + \
                    ["" if m == 0.0 else f"{m:,.2f}  " for m in monthly] + \
                    ["" if total == 0.0 else f"{total:,.2f}  " if use_totals else ""] + \
                    ["" if use_avg else ""]
            if not desc:  # Blank rows
                values = [""] * 16
            tree.insert(eom_parent, "end", text="", values=values, tags=("child",))

        refresh_needed.set(False)

    def show_monthly_focus(month):
        create_monthly_focus_form(form, conn, cursor, int(selected_year.get()), month)

    def on_tree_select(event):
        selected = tree.selection()
        if selected:
            item = selected[0]
            parent_item = tree.parent(item)
            # Enable Drill Down only for child rows under Income/Expense categories
            is_child = item.startswith("child_") and parent_item and \
                tree.item(parent_item, "values")[0] not in ["Excluded Transactions", "End of Month Analysis"]
            drill_down_button.config(state="normal" if is_child else "disabled")
            selected_row.set(tree.item(item, "values")[0].strip() if is_child else "")
        else:
            drill_down_button.config(state="disabled")
            selected_row.set("")

    tree.bind("<<TreeviewSelect>>", on_tree_select)

    refresh_data()

    while form.winfo_exists():
        if refresh_needed.get():
            refresh_data()
        form.update_idletasks()
        form.update()

def create_monthly_focus_form(parent, conn, cursor, year, month):       # Win_ID = 20                               180
    form = tk.Toplevel(parent)
    win_id = 20
    open_form_with_position(form, conn, cursor, win_id, f"Monthly Focus - {month_names[month]}")
    form.geometry(f"{sc(800)}x{sc(1025)}")  # Adjust size
    #form.geometry("800x1045")
    form.resizable(False, False)
    form.configure(bg=config.master_bg)
    form.grab_set()

    # Treeview
    columns = ("Description", "Completed", "Forecast", "Total", "Budget", "Difference", "WARN")
    tree = ttk.Treeview(form, columns=columns, show="tree headings", height=46)
    tree.place(x=sc(10), y=sc(10), width=sc(780), height=sc(960))
    tree.tag_configure("parent", background=config.master_bg, font=("Arial", 10))
    tree.tag_configure("child", font=("Arial", 10))

    # Configure header style
    style = ttk.Style()
    style.configure("Treeview.Heading", font=("Arial", 10), rowheight=25)

    colw_90 = sc(90)
    widths = [sc(230), colw_90, colw_90, colw_90, colw_90, colw_90, sc(70)]
    for col, width in zip(columns, widths):
        tree.heading(col, text=col, anchor="center")
        tree.column(col, width=width, anchor="e" if col != "Description" else "w")
    tree.column("WARN", anchor="center")
    tree.column("#0", width=0)

    # Close Button
    tk.Button(form, text="Close", font=("Arial", 10), bg="white",
            command=lambda: close_form_with_position(form, conn, cursor, win_id)).place(x=sc(350), y=sc(985), width=sc(100))

    # Fetch Data
    def fetch_transactions_sum(pid, cid, tr_stat, month, year):
        query = """
            SELECT COALESCE(SUM(Tr_Amount), 0.0)
            FROM Trans
            WHERE Tr_Year=? AND Tr_Month=? AND Tr_Exp_ID=? AND Tr_ExpSub_ID=? 
                AND Tr_Stat IN ({}) AND Tr_Type IN (1, 2)
        """
        if isinstance(tr_stat, tuple):
            query = query.format(",".join("?" * len(tr_stat)))
            params = [year, month, pid, cid] + list(tr_stat)
        else:
            query = query.format("?")
            params = [year, month, pid, cid, tr_stat]
        cursor.execute(query, params)
        return cursor.fetchone()[0] or 0.0

    def fetch_budget(pid, cid, month, year):
        month_field = f"Bud_M{month}"
        query = """
            SELECT COALESCE({}, 0.0)
            FROM Budget
            WHERE Bud_Year=? AND Bud_PID=? AND Bud_CID=?
        """.format(month_field)
        cursor.execute(query, (year, pid, cid))
        result = cursor.fetchone()
        return result[0] if result else 0.0

    # Populate Treeview
    income_cats = fetch_categories(cursor, year, is_income=True)
    expense_cats = fetch_categories(cursor, year, is_income=False)

    income_total = 0.0
    expenditure_total = 0.0
    income_completed = 0.0
    income_forecast = 0.0
    expenditure_completed = 0.0
    expenditure_forecast = 0.0
    income_budget = 0.0
    expenditure_budget = 0.0

    def populate_category(pid, cat_desc, is_income):
        nonlocal income_total, expenditure_total, income_completed, income_forecast, \
                expenditure_completed, expenditure_forecast, income_budget, expenditure_budget
        subcats = fetch_subcategories(cursor, pid, year)

        # Parent row
        completed = 0.0
        forecast = 0.0
        if subcats:
            for cid, _ in subcats:
                completed += fetch_transactions_sum(pid, cid, (2, 3), month, year)
                forecast += fetch_transactions_sum(pid, cid, 1, month, year)
        else:
            completed = fetch_transactions_sum(pid, 0, (2, 3), month, year)
            forecast = fetch_transactions_sum(pid, 0, 1, month, year)

        month_total = completed + forecast
        budget_total = sum(fetch_budget(pid, cid, month, year) for cid, _ in subcats) if subcats else fetch_budget(pid, 0, month, year)
        budget_remaining = (month_total - budget_total) if is_income else (budget_total - month_total)

        if is_income:
            income_total += month_total
            income_completed += completed
            income_forecast += forecast
            income_budget += budget_total
        else:
            expenditure_total += month_total
            expenditure_completed += completed
            expenditure_forecast += forecast
            expenditure_budget += budget_total

        values = [cat_desc,
                    "" if completed == 0.0 else f"{completed:,.2f}  ",
                    "" if forecast == 0.0 else f"{forecast:,.2f}  ",
                    "" if month_total == 0.0 else f"{month_total:,.2f}  ",
                    "" if budget_total == 0.0 else f"{budget_total:,.2f}  ",
                    "" if budget_remaining == 0.0 else f"{budget_remaining:,.2f}  ",
                    ""]
        parent = tree.insert("", "end", text="", values=values, open=True, tags=("parent",))

        # Child rows
        for cid, sub_desc in subcats:
            completed = fetch_transactions_sum(pid, cid, (2, 3), month, year)
            forecast = fetch_transactions_sum(pid, cid, 1, month, year)
            month_total = completed + forecast
            budget_total = fetch_budget(pid, cid, month, year)
            budget_remaining = (month_total - budget_total) if is_income else (budget_total - month_total)

            sub_values = [f"   {sub_desc}",
                            "" if completed == 0.0 else f"{completed:,.2f}  ",
                            "" if forecast == 0.0 else f"{forecast:,.2f}  ",
                            "" if month_total == 0.0 else f"{month_total:,.2f}  ",
                            "" if budget_total == 0.0 else f"{budget_total:,.2f}  ",
                            "" if budget_remaining == 0.0 else f"{budget_remaining:,.2f}  ",
                            ""]
            tree.insert(parent, "end", text="", values=sub_values, open=True, tags=("child",))

        return parent

    # Populate Income and Expense categories
    for pid, desc in income_cats:
        populate_category(pid, desc, True)

    last_pid = max(pid for pid, _ in expense_cats if pid != 99) if expense_cats else None
    misc_parent = None
    for pid, desc in expense_cats:
        if pid != 99:
            parent = populate_category(pid, desc, False)
        if pid == last_pid:
            misc_parent = parent

    # Total Expenditure
    if misc_parent:
        month_total = expenditure_completed + expenditure_forecast
        budget_remaining = expenditure_budget - month_total
        values = ["Total Expenditure",
                    "" if expenditure_completed == 0.0 else f"{expenditure_completed:,.2f}  ",
                    "" if expenditure_forecast == 0.0 else f"{expenditure_forecast:,.2f}  ",
                    "" if month_total == 0.0 else f"{month_total:,.2f}  ",
                    "" if expenditure_budget == 0.0 else f"{expenditure_budget:,.2f}  ",
                    "" if budget_remaining == 0.0 else f"{budget_remaining:,.2f}  ",
                    ""]
        tree.insert("", "end", text="", values=values, open=True, tags=("parent",))

    # Blank row
    tree.insert("", "end", text="", values=[""] * 7)

    # Total Income Less Outgoings
    #inc_less_out = income_total - expenditure_total
    inc_less_out_completed = income_completed - expenditure_completed
    inc_less_out_forecast = income_forecast - expenditure_forecast
    inc_less_out_month_total = inc_less_out_completed + inc_less_out_forecast
    inc_less_out_budget = income_budget - expenditure_budget
    inc_less_out_remaining = inc_less_out_month_total - inc_less_out_budget
    values = ["Total Income Less Outgoings",
                "" if inc_less_out_completed == 0.0 else f"{inc_less_out_completed:,.2f}  ",
                "" if inc_less_out_forecast == 0.0 else f"{inc_less_out_forecast:,.2f}  ",
                "" if inc_less_out_month_total == 0.0 else f"{inc_less_out_month_total:,.2f}  ",
                "" if inc_less_out_budget == 0.0 else f"{inc_less_out_budget:,.2f}  ",
                "" if inc_less_out_remaining == 0.0 else f"{inc_less_out_remaining:,.2f}  ",
                ""]
    tree.insert("", "end", text="", values=values, open=True, tags=("parent",))


