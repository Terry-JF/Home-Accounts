# Creates main form for the HA2 program

import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import os
import sys
import webbrowser
import logging
from gui import (show_maint_toolbox, create_edit_form)
from db import  (open_db, close_db, fetch_month_rows, fetch_transaction_sums, fetch_statement_balances, fetch_years,
                update_account_year_transactions)
from ui_utils import (COLORS, TEXT_COLORS, refresh_grid, open_form_with_position, close_form_with_position, resource_path)
from focus_forms import (create_summary_form)
from config import (init_config, get_config)

# Setup logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

# Handle DPI scaling for high-resolution displays
try:
    from ctypes import windll
    windll.shcore.SetProcessDpiAwareness(1)
except Exception as e:
    print(f"Failed to set DPI awareness: {e}")


accounts = []  #Global variable

def main():
    # Configure root logger for console output
    logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[logging.StreamHandler()]
    )
    logger = logging.getLogger()
    logger.debug("Starting HA application")
    
    # Initialize configuration
    init_config() 
    
    # Open database connection
    conn, cursor = open_db()
    
    # Start the GUI, passing the connection and cursor
    create_home_screen(conn, cursor)
    
    # Close database connection when GUI exits
    close_db(conn)


def create_home_screen(conn, cursor):
    global accounts  # Keep global for now, could refactor later
    global year_var
    root = tk.Tk()
    win_id = 15
    open_form_with_position(root, conn, cursor, win_id, "HA2 Home Screen")
    scaling_factor = root.winfo_fpixels('1i') / 96  # Get scaling factor (e.g., 2.0 for 200% scaling)
    root.geometry(f"{int(1920 * scaling_factor)}x{int(1045 * scaling_factor)}")  # Adjust size
    #root.geometry("1920x1045")
    root.configure(bg=COLORS["very_pale_blue"])
    root.protocol("WM_DELETE_WINDOW", lambda: None)
#    root.state('normal')  # Not minimized
#    root.lift()           # Bring to front
#    root.focus_force()    # Force focus 

    root.selected_row = tk.IntVar(value=-1)
    root.marked_rows = set()
    year_var = tk.StringVar(value="2025")
    root.rows_container = None # Set later
    root.account_data = None  # Set later
    root.accounts = None  # Set later

    current = datetime.today()

    tc_forecast_var = tk.StringVar(value="")
    tc_processing_var = tk.StringVar(value="")
    tc_complete_var = tk.StringVar(value="")
    tc_total_var = tk.StringVar(value="")

    ff_checkbox_var = tk.IntVar(value=0)

    # Initial account data fetch
    cursor.execute( "SELECT Acc_ID, Acc_Short_Name, Acc_Open, Acc_Jan, Acc_Feb, Acc_Mar, Acc_Apr, Acc_May, Acc_Jun, "
                    "Acc_Jul, Acc_Aug, Acc_Sep, Acc_Oct, Acc_Nov, Acc_Dec, Acc_Credit_Limit "
                    "FROM Account WHERE Acc_Year = ? ORDER BY Acc_ID", (2025,))
    root.account_data = cursor.fetchall()
    root.accounts = [row[1] for row in root.account_data]
    root.credit_limits = [float(row[15] or 0.0) for row in root.account_data]  # Index 15 = Acc_Credit_Limit
    accounts = root.accounts
    account_data = root.account_data

    # Initial transaction fetch data
    root.rows_container = [fetch_month_rows(cursor, current.month, current.year, accounts, root.account_data)]
    year_label = ttk.Label(root, text="Select Year:", font=("Arial", 14), background=COLORS["very_pale_blue"])
    year_label.place(x=int(200 * scaling_factor), y=int(10 * scaling_factor))
    root.year_var = tk.StringVar(value=str(current.year))
    year_rows = fetch_years(cursor)
    year_combo = ttk.Combobox(  root, textvariable=root.year_var, values=[str(y) for y in year_rows], 
                                font=("Arial", 14, "bold"), width=6, state="readonly")
    year_combo.place(x=int(320 * scaling_factor), y=int(10 * scaling_factor))
    year_combo.set(str(current.year))
    year_label.lift()
    year_combo.lift()

    today_date = datetime.today().strftime("%d/%m/%Y")
    current = datetime.today()

    def go_to_today():
        root.year_var.set(str(current.year))
        focus_day=str(current.day)
        notebook.select(tabs[month_map_num[current.month]])
        rows = fetch_month_rows(cursor, current.month, current.year, accounts, account_data)
        refresh_grid(root.tree, rows, root.marked_rows, None, focus_day) 

    today_button = tk.Button(   root, text=f"Today: {today_date}", font=("Arial", 12, "bold"), bg=COLORS["pale_green"],
                                width=int(200 * scaling_factor), height=int(35 * scaling_factor), command=go_to_today)
    today_button.place(x=int(690 * scaling_factor), y=int(10 * scaling_factor), width=int(200 * scaling_factor), height=int(35 * scaling_factor))

    gear_btn = tk.Button(   root, text="⚙", font=("Arial", 16), width=int(2 * scaling_factor), height=1, relief="raised", bg=COLORS["pale_green"],
                            command=lambda: show_maint_toolbox(root, conn, cursor))
    gear_btn.place(x=int(10 * scaling_factor), y=int(10 * scaling_factor))

    help_btn = tk.Button(   root, text="?", font=("Arial", 16), width=int(2 * scaling_factor), height=1, relief="raised", bg=COLORS["pale_green"],
                            command=lambda: show_user_guide())
    help_btn.place(x=int(1850 * scaling_factor), y=int(10 * scaling_factor))

    def exit_program():
        close_form_with_position(root, conn, cursor, win_id)

    exit_button = tk.Button(root, text="EXIT PROGRAM", font=("Arial", 12, "bold"), 
                            bg=COLORS["pale_blue"], width=int(200 * scaling_factor), height=int(35 * scaling_factor), command=exit_program)
    exit_button.place(x=int(1180 * scaling_factor), y=int(10 * scaling_factor), width=int(210 * scaling_factor), height=int(35 * scaling_factor))

    cursor.execute( "SELECT Acc_ID, Acc_Short_Name, Acc_Open, Acc_Jan, Acc_Feb, Acc_Mar, Acc_Apr, Acc_May, Acc_Jun, "
                    "Acc_Jul, Acc_Aug, Acc_Sep, Acc_Oct, Acc_Nov, Acc_Dec FROM Account WHERE Acc_Year = ? ORDER BY Acc_ID", 
                    (int(root.year_var.get()),))
    root.account_data = cursor.fetchall()
    root.accounts = [row[1] for row in account_data] if account_data else ["Cash", "Bank 1", "Bank 2", "Bank 3", "Bank 4", "Bank 5", "CC 1", "CC 2", "CC 3", "CC 4", "CC 5", "Savings", "Investments", "Loans"]
    
    grid_labels = root.accounts
    headers = ["Day", "Date", "R", "Income", "Expense", "Description / Payee"] + grid_labels

    # Account Header Block - AHB
    root.button_grid = []
    btn_width = int(85 * scaling_factor)
    btn_height = int(23 * scaling_factor)
    for row in range(9):
        row_buttons = []
        for col in range(14):
            x_pos = int(200 * scaling_factor) + col * btn_width
            y_pos = int(50 * scaling_factor) + row * btn_height
            if row == 0:
                btn = tk.Button(root, text=grid_labels[col] if col < len(grid_labels) else f"Col{col+1}", 
                                font=("Arial", 10, "bold"), bg=COLORS["dark_grey"], fg=COLORS["white"],
                                command=lambda c=col: on_ahb_acc_button(c))
            elif row == 3:
                btn = tk.Button(root, text=f"R{row+1}C{col+1}", font=("Arial", 10), 
                                bg=COLORS["pale_green"], command=lambda c=col: on_ahb_today_button(c))
            else:
                btn = tk.Button(root, text="", font=("Arial", 10), 
                                bg=COLORS["white"], state="disabled", disabledforeground=COLORS["black"])
                if row == 4:  # Processing Transactions
                    btn.config(fg=TEXT_COLORS["Processing"])
                elif row == 6:  # Forecast Transactions
                    btn.config(fg=TEXT_COLORS["Forecast"])
                elif row == 8:  # Last Statement Balance row
                    btn.config(bg=COLORS["grey"])
            btn.place(x=x_pos, y=y_pos, width=btn_width, height=btn_height)
            row_buttons.append(btn)
        root.button_grid.append(row_buttons)

    grid_label_texts = [
        "Start of Month Balance:", "Completed Transactions:", "Account Balance Today:", 
        "Processing Transactions:", "Processing Balance Today:", "Forecast Transactions:", 
        "Forecast EOM Balance:", "Last Statement Balance:"
    ]

    for i, text in enumerate(grid_label_texts):
        label = tk.Label(root, text=text, font=("Arial", 10), bg=COLORS["very_pale_blue"], anchor="e")
        if i == 3:  # Processing Transactions
            label.config(fg=TEXT_COLORS["Processing"])
        elif i == 5:  # Forecast Transactions
            label.config(fg=TEXT_COLORS["Forecast"])
        label.place(x=int(10 * scaling_factor), y=int(50 * scaling_factor) + (i + 1) * btn_height, width=int(185 * scaling_factor), height=btn_height)

    #  End of Month Summary Block - EOMSB
    eomsbx = int(1400 * scaling_factor)
    eomsby = int(50 * scaling_factor)
    eom_summary_label = ttk.Label(root, text="End of Month Summary", font=("Arial", 14), background=COLORS["very_pale_blue"])
    eom_summary_label.place(x = eomsbx + int(190 * scaling_factor), y = eomsby)

    grid_labels2 = "Cash", "Credit", "Invest", "Loans", "Net"
    root.button_grid2 = []
    btn_width2 = int(86 * scaling_factor)
    btn_height2 = int(35 * scaling_factor)
    for row2 in range(4):
        row_buttons2 = []
        for col2 in range(5):
            x_pos = eomsbx + int(75 * scaling_factor) + col2 * btn_width2
            y_pos = eomsby + int(30 * scaling_factor) + row2 * btn_height2
            if row2 == 0:
                btn = tk.Button(root, text=grid_labels2[col2] if col2 < len(grid_labels2) else f"Col{col+1}", font=("Arial", 11, "bold"), 
                        bg=COLORS["dark_grey"], state="disabled", disabledforeground=COLORS["white"])
            else:
                btn = tk.Button(root, text=f"R{row2+1}C{col2+1}", font=("Arial", 11), 
                        bg=COLORS["white"], state="disabled", disabledforeground=COLORS["black"])
            btn.place(x=x_pos, y=y_pos, width=btn_width2, height=btn_height2)
            row_buttons2.append(btn)
        root.button_grid2.append(row_buttons2)

    grid_label_texts2 = [
        "SOM Bal:", "EOM Bal:", "Change:"
    ]

    for i, text in enumerate(grid_label_texts2):
        label = tk.Label(root, text=text, font=("Arial", 10), bg=COLORS["very_pale_blue"], anchor="e")
        label.place(x = eomsbx + int(10 * scaling_factor), y = eomsby + int(30 * scaling_factor) + (i + 1) * btn_height2, width=int(60 * scaling_factor), height=btn_height2)

    style = ttk.Style()
    style.theme_use("winnative")  # Force a theme 
    style.configure("TNotebook", background=COLORS["very_pale_blue"])
    style.configure("TNotebook.Tab", font=("Arial", 12, "bold"), padding=[int(26 * scaling_factor), int(2 * scaling_factor)], background=COLORS["white"])

    # Map active/inactive states
    style.map("TNotebook.Tab",
            background=[("selected", COLORS["dark_brown"]), ("!selected", COLORS["white"])],
            foreground=[("selected", COLORS["white"]), ("!selected", COLORS["black"])])

    # Frame to constrain notebook width
    notebook_frame = tk.Frame(root, background=COLORS["very_pale_blue"])
    notebook_frame.pack(pady=(int(260 * scaling_factor), 0), padx=int(10 * scaling_factor), fill="x")  # Keep fill="x" for now

    notebook = ttk.Notebook(notebook_frame)
    notebook.pack(side="left", fill="none")  # No fill, let tabs define width

    tabs = {
        "SUMMARY": ttk.Frame(notebook), "BUDGET": ttk.Frame(notebook), "PREV": ttk.Frame(notebook),
        "JAN": ttk.Frame(notebook), "FEB": ttk.Frame(notebook), "MAR": ttk.Frame(notebook),
        "APR": ttk.Frame(notebook), "MAY": ttk.Frame(notebook), "JUN": ttk.Frame(notebook),
        "JUL": ttk.Frame(notebook), "AUG": ttk.Frame(notebook), "SEP": ttk.Frame(notebook),
        "OCT": ttk.Frame(notebook), "NOV": ttk.Frame(notebook), "DEC": ttk.Frame(notebook),
        "NEXT": ttk.Frame(notebook), "COMPARE": ttk.Frame(notebook)
    }
    tab_names = ["SUMMARY", "BUDGET", "PREV", "JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC", "NEXT", "COMPARE"]
    for tab_name in tab_names:
        notebook.add(tabs[tab_name], text=tab_name)

    root.tab_change_handled = False  # Add this flag

    # Main Transaction List - Treeview
    tree_frame = tk.Frame(root)
    tree_frame.pack(padx=(int(10 * scaling_factor), 0), pady=0, anchor="nw") 

    # Define Treeview styles
    style = ttk.Style()
    style.configure("Treeview.Heading", font=("Arial", 11, "bold"))  # Can set font for header row
    style.configure("Treeview", rowheight=int(18 * scaling_factor), background="white")
    root.tree = ttk.Treeview(tree_frame, columns=[f"col{i}" for i in range(20)], show="headings", height=40, style="Treeview")
    scrollbar = tk.Scrollbar(tree_frame, orient="vertical", command=root.tree.yview)
    root.tree.configure(yscrollcommand=scrollbar.set)
    root.tree.pack(side="left", fill="y")
    scrollbar.pack(side="right", fill="y")

    def on_tree_select(event):
        selected = root.tree.selection()
        if selected:
            item = selected[0]
            idx = int(item)
            root.selected_row.set(idx)
        else:
            root.selected_row.set(-1)

    root.tree.bind("<<TreeviewSelect>>", on_tree_select)

    w1 = int(83 * scaling_factor)
    widths = [int(40 * scaling_factor), int(40 * scaling_factor), int(20 * scaling_factor), int(80 * scaling_factor), int(80 * scaling_factor), int(250 * scaling_factor),
            w1, w1, w1, w1, w1, w1, w1, w1, w1, w1, w1, w1, w1, w1]
    anchors = ["w", "w", "w", "e", "e", "w", "e", "e", "e", "e", "e", "e", "e", "e", "e", "e", "e", "e", "e", "e"]
    for i, (header, width, anchor) in enumerate(zip(headers, widths, anchors)):
        root.tree.heading(f"col{i}", text=header, anchor="center")
        root.tree.column(f"col{i}", width=width, anchor=anchor, stretch=False)

    # Define Total row tags
    root.tree.tag_configure("overdrawn", background="#FFC1CC", foreground="#FF0000", font=("Arial", 11, "bold"))
    root.tree.tag_configure("daily_total", background="#555555", foreground="#FFFFFF", font=("Arial", 10))

    # Define non-Total row tags
    root.tree.tag_configure("weekday", background="white")
    root.tree.tag_configure("weekend", background=COLORS["oldlace"])
    root.tree.tag_configure("marked", background=COLORS["marker"])
    root.tree.tag_configure("yellow", background=COLORS["yellow"])
    root.tree.tag_configure("green", background=COLORS["green"])
    root.tree.tag_configure("cyan", background=COLORS["cyan"])
    root.tree.tag_configure("orange", background=COLORS["orange"])
    root.tree.tag_configure("forecast", foreground="brown", font=("Arial", 10))
    root.tree.tag_configure("processing", foreground="blue", font=("Arial", 10))
    root.tree.tag_configure("complete", foreground="black", font=("Arial", 10))

    # Focal point markers at row 12
    marker_x = int(280 * scaling_factor)
    marker_y = int(18 * scaling_factor)
    left_canvas = tk.Canvas(root, bg=COLORS["very_pale_blue"], highlightthickness=0, width=int(10 * scaling_factor), height=int(40 * scaling_factor))
    left_canvas.place(x=0, y=int(250 * scaling_factor) + marker_x)
    left_canvas.create_text(5, marker_y, text=">", font=("Arial", 14, "bold"), fill="red", anchor="center")
    right_canvas = tk.Canvas(root, bg=COLORS["very_pale_blue"], highlightthickness=0, width=int(20 * scaling_factor), height=int(40 * scaling_factor))
    right_canvas.place(x=int(1700 * scaling_factor), y=int(250 * scaling_factor) + marker_x)
    right_canvas.create_text(5, marker_y, text="<", font=("Arial", 14, "bold"), fill="red", anchor="center")

    # Initialize rows
    month_map_num = {1: "JAN", 2: "FEB", 3: "MAR", 4: "APR", 5: "MAY", 6: "JUN", 7: "JUL", 8: "AUG", 9: "SEP", 10: "OCT", 11: "NOV", 12: "DEC"}
    root.rows_container[0] = fetch_month_rows(cursor, current.month, current.year, root.accounts, root.account_data)
    notebook.select(tabs[month_map_num[current.month]])
    month_map_txt = {"JAN": 1, "FEB": 2, "MAR": 3, "APR": 4, "MAY": 5, "JUN": 6, "JUL": 7, "AUG": 8, "SEP": 9, "OCT": 10, "NOV": 11, "DEC": 12}

    def get_current_month():
        selected_tab = notebook.tab(notebook.select(), "text")
        return month_map_txt.get(selected_tab, 3)
    root.get_current_month = get_current_month  # Attach to root    

    def create_budget_form(parent, conn, cursor):  # PLACEHOLDER ***************
        form = tk.Toplevel(parent)
        win_id = 12
        open_form_with_position(form, conn, cursor, win_id, "Annual Budget Performance")
        form.geometry(f"{int(1800 * scaling_factor)}x{int(1040 * scaling_factor)}")  # Adjust size 
        #form.geometry("1800x1170")
        tk.Label(form, text="Annual Budget Performance - Under Construction", font=("Arial", 12)).pack(pady=20)
        tk.Button(  form, text="Close", width=15, 
                    command=lambda: close_form_with_position(form, conn, cursor, win_id)).pack(pady=10)
        form.wait_window()

    def create_compare_form(parent, conn, cursor):  # PLACEHOLDER ***************
        form = tk.Toplevel(parent)
        win_id = 13
        open_form_with_position(form, conn, cursor, win_id, "Compare Current Year to Previous Year")
        form.geometry(f"{int(800 * scaling_factor)}x{int(600 * scaling_factor)}")  # Adjust size 
        #form.geometry("800x860")
        tk.Label(form, text="Compare Current Year to Previous Year - Under Construction", font=("Arial", 12)).pack(pady=20)
        tk.Button(  form, text="Close", width=15, 
                    command=lambda: close_form_with_position(form, conn, cursor, win_id)).pack(pady=10)
        form.wait_window()

    def create_account_list_form(parent, conn, cursor, acc):  # PLACEHOLDER ***************
        form = tk.Toplevel(parent)
        win_id = 16
        open_form_with_position(form, conn, cursor, win_id, "List all Transactions for a single Account")
        form.geometry(f"{int(1020 * scaling_factor)}x{int(1010 * scaling_factor)}")  # Adjust size 
        #form.geometry("1020x1010")
        tk.Label(form, text="List all Transactions for a single Account - Under Construction", font=("Arial", 12)).pack(pady=20)
        tk.Label(form, text=grid_labels[acc], font=("Arial", 12)).pack(pady=20)
        tk.Button(  form, text="Close", width=15, 
                    command=lambda: close_form_with_position(form, conn, cursor, win_id)).pack(pady=10)
        form.wait_window()

    def update_ahb(root, month, year):      # Now also updates EOMSB
        # Fetch fresh account data with credit limits
        cursor.execute( "SELECT Acc_ID, Acc_Short_Name, Acc_Open, Acc_Jan, Acc_Feb, Acc_Mar, Acc_Apr, Acc_May, Acc_Jun, "
                        "Acc_Jul, Acc_Aug, Acc_Sep, Acc_Oct, Acc_Nov, Acc_Dec, Acc_Credit_Limit "
                        "FROM Account WHERE Acc_Year = ? ORDER BY Acc_ID", (year,))
        root.account_data = cursor.fetchall()
        root.accounts = [row[1] for row in root.account_data]
        root.credit_limits = [float(row[15] or 0.0) for row in root.account_data]
        global accounts
        accounts = root.accounts
        cash_som = 0
        credit_som = 0
        invest_som = 0
        loans_som = 0
        cash_eom = 0
        credit_eom = 0
        invest_eom = 0
        loans_eom = 0

        # Start of Month Balance - also EOMSB
        for col, account in enumerate(root.account_data):
            acc_open = float(account[2] or 0.0)
            balance = acc_open
            for m in range(month - 1):
                monthly_change = float(account[3 + m] or 0.0)
                balance += monthly_change
            root.button_grid[1][col].config(text="{:,.2f}".format(balance))
            if col < 6:
                cash_som += balance
            elif col < 12:
                credit_som += balance
            elif col == 12:
                invest_som = balance
            else:
                loans_som = balance

        # Transaction sums
        completed, processing, forecast, tc_complete, tc_processing, tc_forecast, tc_total = fetch_transaction_sums(cursor, month, year, root.accounts)

        # Update Transaction Counts
        tc_forecast_var.set(str(tc_forecast))
        tc_processing_var.set(str(tc_processing))
        tc_complete_var.set(str(tc_complete))
        tc_total_var.set(str(tc_total))

#        print(f"Forecast:{tc_forecast}  Processing:{tc_processing}  Complete:{tc_complete}  Total:{tc_total}")
#        print(f"Forecast:{tc_forecast_var.get()}  Processing:{tc_processing_var.get()}  Complete:{tc_complete_var.get()}  Total:{tc_total_var.get()}")

        # Completed Transactions
        for col, total in enumerate(completed):
            root.button_grid[2][col].config(text="{:,.2f}".format(total))

        # Account Balance Today
        for col in range(len(root.accounts)):
            som = float(root.button_grid[1][col]["text"].replace(",", ""))
            comp = completed[col]
            root.button_grid[3][col].config(text="{:,.2f}".format(som + comp))

        # Processing Transactions
        for col, total in enumerate(processing):
            root.button_grid[4][col].config(text="{:,.2f}".format(total))

        # Processing Balance Today
        for col in range(len(root.accounts)):
            today = float(root.button_grid[3][col]["text"].replace(",", ""))
            proc = processing[col]
            root.button_grid[5][col].config(text="{:,.2f}".format(today + proc))

        # Forecast Transactions
        for col, total in enumerate(forecast):
            root.button_grid[6][col].config(text="{:,.2f}".format(total))

        # Forecast EOM Balance - also EOMSB
        for col in range(len(root.accounts)):
            proc_today = float(root.button_grid[5][col]["text"].replace(",", ""))
            fore = forecast[col]
            fore += proc_today
            root.button_grid[7][col].config(text="{:,.2f}".format(fore))
            if col < 6:
                cash_eom += fore
            elif col < 12:
                credit_eom += fore
            elif col == 12:
                invest_eom = fore
            else:
                loans_eom = fore

        # Last Statement Balance
        stmt_balances = fetch_statement_balances(cursor, month, year, root.accounts)
        for col, balance in enumerate(stmt_balances):
            if balance is None:
                root.button_grid[8][col].place_forget()
            else:
                root.button_grid[8][col].place(x=int((200 + col * 85) * scaling_factor), y=int((50 + 8 * 23) * scaling_factor), width=int(85 * scaling_factor), height=int(23 * scaling_factor))
                root.button_grid[8][col].config(text="{:,.2f}".format(balance))

        net_som = cash_som + invest_som + credit_som + loans_som
        net_eom = cash_eom + invest_eom + credit_eom + loans_eom
        cash_change = cash_eom - cash_som
        credit_change = credit_eom - credit_som
        invest_change = invest_eom - invest_som
        loans_change = loans_eom - loans_som
        net_change = net_eom - net_som

        root.button_grid2[1][0].config(text="{:,.2f}".format(cash_som))
        root.button_grid2[1][1].config(text="{:,.2f}".format(credit_som))
        root.button_grid2[1][2].config(text="{:,.2f}".format(invest_som))
        root.button_grid2[1][3].config(text="{:,.2f}".format(loans_som))
        root.button_grid2[1][4].config(text="{:,.2f}".format(net_som))
        root.button_grid2[2][0].config(text="{:,.2f}".format(cash_eom))
        root.button_grid2[2][1].config(text="{:,.2f}".format(credit_eom))
        root.button_grid2[2][2].config(text="{:,.2f}".format(invest_eom))
        root.button_grid2[2][3].config(text="{:,.2f}".format(loans_eom))
        root.button_grid2[2][4].config(text="{:,.2f}".format(net_eom))
        root.button_grid2[3][0].config(text="{:,.2f}".format(cash_change))
        root.button_grid2[3][1].config(text="{:,.2f}".format(credit_change))
        root.button_grid2[3][2].config(text="{:,.2f}".format(invest_change))
        root.button_grid2[3][3].config(text="{:,.2f}".format(loans_change))
        root.button_grid2[3][4].config(text="{:,.2f}".format(net_change))

    # Attach to root
    root.update_ahb = update_ahb

    # Initial call
    update_ahb(root, current.month, current.year)

    def on_ahb_today_button(col):
        if 0 <= col < len(root.button_grid[0]):
            if root.button_grid[3][col].cget("bg") == COLORS["green"]:
                root.button_grid[3][col].config(bg=COLORS["pale_green"])
            else:
                root.button_grid[3][col].config(bg=COLORS["green"])

    def on_ahb_acc_button(col):
        if 0 <= col < len(root.button_grid[0]):
            create_account_list_form(root, conn, cursor, col)

    def show_user_guide():
        if getattr(sys, 'frozen', False):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.dirname(os.path.dirname(__file__))
        pdf_path = os.path.join(base_path, "docs", "HA2_User_Guide.pdf")
        webbrowser.open(pdf_path)

    def on_tab_change(event):
        selected_tab = notebook.tab(notebook.select(), "text")

        if selected_tab in month_map_txt:
            root.previous_tab = selected_tab
            month_idx = month_map_txt[selected_tab] # get tab number, Jan=1, Dec=12
            root.rows_container[0] = fetch_month_rows(cursor, month_idx, int(root.year_var.get()), root.accounts, root.account_data)
            # If it is current month/year then set focus_day to today, else 1
            if month_idx == int(current.month) and int(root.year_var.get()) == int(current.year):
                focus_day = str(current.day) 
            else:
                focus_day = 1
            refresh_grid(root.tree, root.rows_container[0], root.marked_rows, focus_day=focus_day)
            root.update_ahb(root, month_idx, int(root.year_var.get()))  # Pass root explicitly
        elif selected_tab == "SUMMARY":
            create_summary_form(root, conn, cursor)
            notebook.select(tabs[root.previous_tab])  # Return to previous
            notebook.event_generate("<<NotebookTabChanged>>")  # Force event
        elif selected_tab == "BUDGET":
            create_budget_form(root, conn, cursor)
            notebook.select(tabs[root.previous_tab])  # Return to previous
            notebook.event_generate("<<NotebookTabChanged>>")  # Force event

        elif selected_tab == "PREV":
            current_year = int(root.year_var.get())
            year_rows = fetch_years(cursor) # get valid year range from Lookup table
            if str(current_year - 1) in [str(y) for y in year_rows]: # we have a previous year
                root.year_var.set(str(current_year - 1))
                notebook.select(tabs["DEC"])
                update_account_year_transactions(cursor, conn, current_year-1, accounts)
                root.rows_container[0] = fetch_month_rows(cursor, 12, current_year - 1, accounts, account_data)
                refresh_grid(root.tree, root.rows_container[0], root.marked_rows)
                for col, account in enumerate(account_data):
                    acc_open = float(account[2]) if account[2] is not None else 0.0
                    balance = acc_open + sum(float(account[3 + m]) if account[3 + m] is not None else 0.0 for m in range(11))
                    root.button_grid[1][col].config(text="{:,.2f}".format(balance))
            else:  # No PREV year exists in system
                messagebox.showinfo("Not a valid Year", "No PREV year exists in system")
                notebook.select(tabs[root.previous_tab])  # Return to previous
                notebook.event_generate("<<NotebookTabChanged>>")  # Force event

        elif selected_tab == "NEXT":
            current_year = int(root.year_var.get())
            year_rows = fetch_years(cursor) # get valid year range from Lookup table
            if str(current_year + 1) in [str(y) for y in year_rows]: # we have a next year
                root.year_var.set(str(current_year + 1))
                notebook.select(tabs["JAN"])
                update_account_year_transactions(cursor, conn, current_year+1, accounts)
                root.rows_container[0] = fetch_month_rows(cursor, 1, current_year + 1, accounts, account_data)
                refresh_grid(root.tree, root.rows_container[0], root.marked_rows)
                for col, account in enumerate(account_data):
                    acc_open = float(account[2]) if account[2] is not None else 0.0
                    root.button_grid[1][col].config(text="{:,.2f}".format(acc_open))
            else:  # No NEXT year exists in system
                messagebox.showinfo("Not a valid Year", "No NEXT year exists in system")
                notebook.select(tabs[root.previous_tab])  # Return to previous
                notebook.event_generate("<<NotebookTabChanged>>")  # Force event
        elif selected_tab == "COMPARE":
            create_compare_form(root, conn, cursor)
            notebook.select(tabs[root.previous_tab])  # Return to previous
            notebook.event_generate("<<NotebookTabChanged>>")  # Force event
        else:
            for item in root.tree.get_children():
                root.tree.delete(item)

        root.tab_change_handled = True  # Set after first run

    notebook.bind("<<NotebookTabChanged>>", on_tab_change)

    def on_year_change(event):
        selected_tab = notebook.tab(notebook.select(), "text")
        new_year = root.year_var.get()
        if selected_tab in month_map_txt:
            month_idx = month_map_txt[selected_tab]
            cursor.execute("SELECT Acc_ID, Acc_Short_Name, Acc_Open, Acc_Jan, Acc_Feb, Acc_Mar, Acc_Apr, Acc_May, Acc_Jun, "
                        "Acc_Jul, Acc_Aug, Acc_Sep, Acc_Oct, Acc_Nov, Acc_Dec, Acc_Credit_Limit "
                        "FROM Account WHERE Acc_Year = ? ORDER BY Acc_ID", (int(root.year_var.get()),))
            root.account_data = cursor.fetchall()
            root.accounts = [row[1] for row in root.account_data]
            root.credit_limits = [float(row[15] or 0.0) for row in root.account_data]  # Update credit limits too
            # Sync globals
            global accounts, account_data
            accounts = root.accounts
            account_data = root.account_data
            # Fetch and refresh grid
            root.rows_container[0] = fetch_month_rows(cursor, month_idx, int(new_year), root.accounts, root.account_data)
            current = datetime.now()
            focus_day = str(current.day) if month_idx == current.month and int(new_year) == current.year else "1"
            refresh_grid(root.tree, root.rows_container[0], root.marked_rows, focus_day=focus_day)
            root.update_ahb(root, month_idx, int(new_year))
            # Update headers
            for col in range(14):
                root.button_grid[0][col].config(text=root.accounts[col] if col < len(root.accounts) else f"Col{col+1}")
                root.tree.heading(f"col{col+5}", text=root.accounts[col] if col < len(root.accounts) else f"Col{col+1}")

    year_combo.bind("<<ComboboxSelected>>", on_year_change)

    # Top and Bottom buttons for treeview scrolling
    top_button = tk.Button(root, text="T", font=("Arial", 14, "bold"), width=2, height=1, command=lambda: root.tree.yview_moveto(0))
    top_button.place(x=int(1705 * scaling_factor), y=int(293 * scaling_factor))  # Top aligns with tree_frame top

    bottom_button = tk.Button(root, text="B", font=("Arial", 14, "bold"), width=2, height=1, command=lambda: root.tree.yview_moveto(1))
    bottom_button.place(x=int(1705 * scaling_factor), y=int(1003 * scaling_factor))  # Bottom aligns with treeview bottom (42 rows * ~25px) 

    # Configue right hand side buttons for New / Edit / Status change / Flags / Row Marker 
    def set_flag(flag_value):
        selected = root.selected_row.get()
        if selected >= 0 and selected < len(root.rows_container[0]) and root.rows_container[0][selected]["status"] != "Total":
            tr_id = root.rows_container[0][selected]["tr_id"]
            cursor.execute("UPDATE Trans SET Tr_Query_flag = ? WHERE Tr_ID = ?", (flag_value, tr_id))
            conn.commit()
            root.rows_container[0] = fetch_month_rows(cursor, get_current_month(), int(root.year_var.get()), accounts, account_data)
#            root.rows_container[0][selected]["flag"] = flag_value
            refresh_grid(root.tree, root.rows_container[0], root.marked_rows, focus_idx=selected)
            update_ahb(root, get_current_month(), int(root.year_var.get()))  # Refresh AHB

    def clear_flag():
        selected = root.selected_row.get()
        if selected >= 0 and selected < len(root.rows_container[0]) and root.rows_container[0][selected]["status"] != "Total":
            tr_id = root.rows_container[0][selected]["tr_id"]
            cursor.execute("SELECT Tr_Query_Flag FROM Trans WHERE Tr_ID=?", (tr_id,))
            current_flag = cursor.fetchone()[0]
            new_flag = 0 if not (current_flag & 8) else current_flag & ~8
            cursor.execute("UPDATE Trans SET Tr_Query_Flag=? WHERE Tr_ID=?", (new_flag, tr_id))
            conn.commit()
            root.rows_container[0] = fetch_month_rows(cursor, get_current_month(), int(root.year_var.get()), accounts, account_data)
            refresh_grid(root.tree, root.rows_container[0], root.marked_rows, focus_idx=selected)
            update_ahb(root, get_current_month(), int(root.year_var.get()))  # Refresh AHB

    def set_marker():
        selected = root.selected_row.get()
        if selected >= 0 and selected < len(root.rows_container[0]) and root.rows_container[0][selected]["status"] != "Total":
            if selected in root.marked_rows:
                root.marked_rows.remove(selected)
            else:
                root.marked_rows.add(selected)
            refresh_grid(root.tree, root.rows_container[0], root.marked_rows, focus_idx=selected)

    def set_status(status_value):
#        print("def set status ", status_value)  # Debug
        selected = root.selected_row.get()
#        print("set status ", selected, len(root.rows_container[0]))  # Debug
        if selected >= 0 and selected < len(root.rows_container[0]) and root.rows_container[0][selected]["status"] != "Total":
            tr_id = root.rows_container[0][selected]["tr_id"]
            status_map = {"Complete": 3, "Processing": 2, "Forecast": 1}
#            print("set status - status_map[status_value], tr_id ", status_map[status_value], tr_id, )  # Debug
            cursor.execute("UPDATE Trans SET Tr_Stat = ? WHERE Tr_ID = ?", (status_map[status_value], tr_id))
            conn.commit()
            root.rows_container[0] = fetch_month_rows(cursor, get_current_month(), int(root.year_var.get()), accounts, account_data)
            refresh_grid(root.tree, root.rows_container[0], root.marked_rows, focus_idx=selected)
            update_ahb(root, get_current_month(), int(root.year_var.get()))  # Refresh AHB

    def clear_all_markers():
        selected = root.selected_row.get()
        current_scroll = root.tree.yview()[0]  # Fraction from 0.0 to 1.0
        root.marked_rows.clear()
        if selected >= 0 and selected < len(root.rows_container[0]):
            refresh_grid(root.tree, root.rows_container[0], root.marked_rows, focus_idx=selected)
        else:
            refresh_grid(root.tree, root.rows_container[0], root.marked_rows)
            root.tree.yview_moveto(current_scroll)  # Restore scroll

    def toggle_ffid():
        print(f"ff_checkbox_var before toggle: {ff_checkbox_var.get()}")
        if ff_checkbox_var.get() == 0:
            ff_checkbox_var.set(1)     # Set checkbox
            ff_box.config(image=checked_img)
            cursor.execute("UPDATE Lookups SET Lup_Seq = 1 WHERE Lup_LupT_ID=8")
            conn.commit()
            print("Lup_Seq = 1")
        else:
            ff_checkbox_var.set(0)    # clear checkbox
            ff_box.config(image=unchecked_img)
            cursor.execute("UPDATE Lookups SET Lup_Seq = 0 WHERE Lup_LupT_ID=8")
            conn.commit()
            print("Lup_Seq = 0")
        print(f"ff_checkbox_var after toggle: {ff_checkbox_var.get()}")

        # Fetch fresh rows with updated ff_flag
        month_idx = root.get_current_month()        # gets currently selected tab
        root.rows_container[0] = fetch_month_rows(cursor, month_idx, int(root.year_var.get()), root.accounts, root.account_data)
        # Refresh grid with current focus
        current = datetime.now()
        focus_day = str(current.day) if month_idx == current.month and int(root.year_var.get()) == current.year else "1"
        refresh_grid(root.tree, root.rows_container[0], root.marked_rows, focus_day=focus_day)

    # NEW & EDIT Buttons
    tk.Button(  root, text="NEW", font=("Arial", 14), width=7, height=3, bg=COLORS["pale_blue"],
                command=lambda: [root.selected_row.set(-1), create_edit_form(root, root.rows_container[0], root.tree, fetch_month_rows, None, conn, 
                    cursor, get_current_month(), int(root.year_var.get()), local_accounts=accounts, account_data=account_data), 
                    update_ahb(root, get_current_month(), int(root.year_var.get()))]).place(x=int(1820 * scaling_factor), y=int(350 * scaling_factor))
    tk.Button(  root, text="EDIT", font=("Arial", 14), width=7, height=3,
                command=lambda: create_edit_form(root, root.rows_container[0], root.tree, fetch_month_rows, 
                    root.selected_row.get(), conn, cursor, root.get_current_month(), 
                    int(root.year_var.get()), local_accounts=root.accounts, 
                    account_data=root.account_data) if root.selected_row.get() >= 0 and root.selected_row.get() < len(root.rows_container[0]) else 
                    messagebox.showinfo("HA2", "Please select a row to edit!")).place(x=int(1715 * scaling_factor), y=int(350 * scaling_factor))

    # Update Transaction Status Buttons
    btn_frame2 = tk.LabelFrame(root, text="Update Status", font=("Arial", 11), bg=COLORS["very_pale_blue"], padx=int(23 * scaling_factor), pady=int(10 * scaling_factor))
    btn_frame2.place(x=int(1715 * scaling_factor), y=int(445 * scaling_factor))
    tk.Button(  btn_frame2, text="Complete", font=("Arial", 11), fg=TEXT_COLORS["Complete"], width=14, height=1,
                command=lambda: set_status("Complete")).pack(side="top", padx=int(5 * scaling_factor), pady=int(3 * scaling_factor))
    tk.Button(  btn_frame2, text="Processing", font=("Arial", 11), fg=TEXT_COLORS["Processing"], width=14, height=1,
                command=lambda: set_status("Processing")).pack(side="top", padx=int(5 * scaling_factor), pady=int(3 * scaling_factor))
    tk.Button(  btn_frame2, text="Forecast", font=("Arial", 11), fg=TEXT_COLORS["Forecast"], width=14, height=1,
                command=lambda: set_status("Forecast")).pack(side="top", padx=int(5 * scaling_factor), pady=int(3 * scaling_factor))
    
    # Set/Clear Flag Buttons
    flag_frame = tk.LabelFrame(root, text="Set/Clear Flag", font=("Arial", 11), bg=COLORS["very_pale_blue"], padx=int(5 * scaling_factor), pady=int(10 * scaling_factor))
    flag_frame.place(x=int(1715 * scaling_factor), y=int(600 * scaling_factor))
    tk.Button(flag_frame, text="Set", width=6, bg=COLORS["yellow"], command=lambda: set_flag(1)).grid(row=1, column=0, padx=int(4 * scaling_factor))
    tk.Button(flag_frame, text="Set", width=6, bg=COLORS["green"], command=lambda: set_flag(2)).grid(row=1, column=1, padx=int(4 * scaling_factor))
    tk.Button(flag_frame, text="Set", width=6, bg=COLORS["cyan"], command=lambda: set_flag(4)).grid(row=1, column=2, padx=int(4 * scaling_factor))
    tk.Button(flag_frame, text="Clear Flag", width=23, command=clear_flag).grid(row=2, column=0, columnspan=3, padx=int(4 * scaling_factor), pady=int(5 * scaling_factor))

    # Row Marker Buttons
    btn_frame3 = tk.LabelFrame(root, text="Set/Clear Row Marker", font=("Arial", 11), bg=COLORS["very_pale_blue"], padx=int(5 * scaling_factor), pady=int(5 * scaling_factor))
    btn_frame3.place(x=int(1715 * scaling_factor), y=int(710 * scaling_factor))
    tk.Button(btn_frame3, text="Set/Clear Row Marker", width=23, bg=COLORS["marker"], command=set_marker).pack(side="top", padx=int(5 * scaling_factor), pady=int(5 * scaling_factor))
    tk.Button(btn_frame3, text="Clear ALL Row Markers", width=23, command=clear_all_markers).pack(side="top", padx=int(5 * scaling_factor), pady=int(10 * scaling_factor))

    # Transaction Count Frame
    tc_frame = tk.LabelFrame(root, text="Transaction Count", font=("Arial", 11), bg=COLORS["very_pale_blue"], padx=int(5 * scaling_factor), pady=int(5 * scaling_factor))
    tc_frame.place(x=int(1715 * scaling_factor), y=int(830 * scaling_factor), width=int(192 * scaling_factor), height=int(130 * scaling_factor))
    tk.Label(tc_frame, text="   Forecast:", anchor=tk.W, width=16, font=("Arial", 11), 
             bg=COLORS["very_pale_blue"], fg=TEXT_COLORS["Forecast"]).place(x=int(10 * scaling_factor), y=0)
    tk.Label(tc_frame, text=str(tc_forecast_var.get()), anchor=tk.E, width=4, font=("Arial", 11), 
             bg=COLORS["very_pale_blue"], fg=TEXT_COLORS["Forecast"]).place(x=int(120 * scaling_factor), y=0)
    tk.Label(tc_frame, text="Processing:", anchor=tk.W, width=16, font=("Arial", 11), 
             bg=COLORS["very_pale_blue"], fg=TEXT_COLORS["Processing"]).place(x=int(10 * scaling_factor), y=int(25 * scaling_factor))
    tk.Label(tc_frame, text=str(tc_processing_var.get()), anchor=tk.E, width=4, font=("Arial", 11), 
             bg=COLORS["very_pale_blue"], fg=TEXT_COLORS["Processing"]).place(x=int(120 * scaling_factor), y=int(25 * scaling_factor))
    tk.Label(tc_frame, text="   Complete:", anchor=tk.W, width=16, font=("Arial", 11), 
             bg=COLORS["very_pale_blue"], fg=TEXT_COLORS["Complete"]).place(x=int(10 * scaling_factor), y=int(50 * scaling_factor))
    tk.Label(tc_frame, text=str(tc_complete_var.get()), anchor=tk.E, width=4, font=("Arial", 11), 
             bg=COLORS["very_pale_blue"], fg=TEXT_COLORS["Complete"]).place(x=int(120 * scaling_factor), y=int(50 * scaling_factor))
    tk.Label(tc_frame, text="          Total:", anchor=tk.W, width=16, font=("Arial", 11), 
             bg=COLORS["very_pale_blue"], fg=COLORS["black"]).place(x=int(10 * scaling_factor), y=int(75 * scaling_factor))
    tk.Label(tc_frame, text=str(tc_total_var.get()), anchor=tk.E, width=4, font=("Arial", 11), 
             bg=COLORS["very_pale_blue"], fg=COLORS["black"]).place(x=int(120 * scaling_factor), y=int(75 * scaling_factor))

    # Display Firefly ID# Control and Initialise
    #ff_box=tk.Checkbutton(  root, text="Show FF Journal_ID in list", font=("Arial", 11), bg=COLORS["very_pale_blue"], 
    #                        command=toggle_ffid, variable=ff_checkbox_var)
    #ff_box.place(x=int(1710 * scaling_factor), y=int(970 * scaling_factor))
    
    unchecked_img = tk.PhotoImage(file=resource_path("icons/unchecked_16.png")).zoom(int(scaling_factor))
    checked_img = tk.PhotoImage(file=resource_path("icons/checked_16.png")).zoom(int(scaling_factor))

    tk.Label(root, text="Show GC ID in list", anchor=tk.W, width=int(14 * scaling_factor), font=("Arial", 11), 
             bg=COLORS["very_pale_blue"], fg=COLORS["black"]).place(x=int(1750 * scaling_factor), y=int(970 * scaling_factor))
    ff_box=tk.Button(root, image=unchecked_img, bg=COLORS["very_pale_blue"], command=toggle_ffid)
    ff_box.place(x=int(1720 * scaling_factor), y=int(970 * scaling_factor))
    
    # Initialise GC ID 'checkbox' image
    cursor.execute("SELECT Lup_seq FROM Lookups WHERE Lup_LupT_ID = 8")
    ff_rows = cursor.fetchone()
    ff_flag = ff_rows[0]
    ff_checkbox_var.set(int(ff_flag))
    if ff_flag == 1:
        ff_box.config(image=checked_img)
    else:
        ff_box.config(image=unchecked_img)

#    print(f"ff checkbox initialised to: {ff_checkbox_var.get()}")

    def on_mouse_wheel(event):
        root.tree.yview_scroll(-1 * (event.delta // 120), "units")

    root.bind("<MouseWheel>", on_mouse_wheel)

    root.mainloop()


if __name__ == "__main__":
    main()












