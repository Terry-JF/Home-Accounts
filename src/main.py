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
from ui_utils import (refresh_grid, open_form_with_position, close_form_with_position, resource_path, set_sc, sc)
from focus_forms import (create_summary_form)
from config import (COLORS, init_config, load_colors_from_db)
import config

# Initialize configuration - sets up logger and loads DB settings
# Sets HA for either Test or Production
init_config() 

# Setup logging
logger = logging.getLogger('HA.main')
logger.debug("Starting HA application")

# Handle DPI scaling for high-resolution displays
try:
    from ctypes import windll
    windll.shcore.SetProcessDpiAwareness(1)
except Exception as e:
    logger.error(f"Failed to set DPI awareness: {e}")

accounts = []  #Global variable

def main():
    # Open database connection
    logger.debug("Calling open_db()")
    conn, cursor, db_path = open_db()
    logger.debug(f"Database opened at {db_path}")
    
    # Load the colour dictionary
    load_colors_from_db(db_path)
    
    # Initialise TK then call set_sc IMMEDIATELY after
    root = tk.Tk()
    set_sc(root)
        
    # Start the GUI, passing the connection and cursor
    logger.debug("Calling create_home_screen()")
    create_home_screen(root, conn, cursor)
    
    # Close database connection when GUI exits
    logger.debug("Calling close_db()")
    close_db(conn)

def create_home_screen(root, conn, cursor):
    global accounts  # Keep global for now, could refactor later
    global year_var
    
    win_id = 15
    open_form_with_position(root, conn, cursor, win_id, "HA Home Screen")
    root.geometry(f"{sc(1920)}x{sc(1045)}")  # Adjust size
    root.configure(bg=config.master_bg)
    root.protocol("WM_DELETE_WINDOW", lambda: None)
    logger.debug("Created HA Home Screen")

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

    # load and anchor images
    logger.debug(f"scaling_factor={sc(1)}")
    unchecked_img = tk.PhotoImage(file=resource_path("icons/unchecked_16.png")).zoom(sc(1))
    root.unchecked_img = unchecked_img
    checked_img = tk.PhotoImage(file=resource_path("icons/checked_16.png")).zoom(sc(1))
    root.checked_img = checked_img
    radio_unchecked_img = tk.PhotoImage(file=resource_path("icons/radio_0_32.png")).zoom(int(sc(1) * 0.5))
    root.radio_unchecked_img = radio_unchecked_img
    radio_checked_img = tk.PhotoImage(file=resource_path("icons/radio_1_32.png")).zoom(int(sc(1) * 0.5))
    root.radio_checked_img = radio_checked_img
    
    ############################################################################
    # Prepare to use custom checkbox
    ff_checkbox_var = tk.IntVar(value=0)        # <<< Replace with local vars
    
    root.option_add("*TCombobox*Listbox*Font", (config.ha_normal))  # normal HA font size for Combobox dropdown list
    
    # Initialize image references to prevent garbage collection
    #image_refs = []
    #image_refs.extend([unchecked_img, checked_img, radio_unchecked_img, radio_checked_img])
    ############################################################################


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
    year_label = ttk.Label(root, text="Select Year:", font=(config.ha_head14), background=config.master_bg)
    year_label.place(x=sc(200), y=sc(10))
    root.year_var = tk.StringVar(value=str(current.year))
    year_rows = fetch_years(cursor)
    year_combo = ttk.Combobox(  root, textvariable=root.year_var, values=[str(y) for y in year_rows], 
                                font=(config.ha_head14), width=6, state="readonly")
    year_combo.place(x=sc(320), y=sc(10))
    year_combo.set(str(current.year))
    year_label.lift()
    year_combo.lift()
    
    # Define go_to_today to refresh the Home Form grid
    def go_to_today():
        root.year_var.set(str(current.year))
        focus_day = str(current.day)
        notebook.select(tabs[month_map_num[current.month]])
        rows = fetch_month_rows(cursor, current.month, current.year, accounts, account_data)
        refresh_grid(root.tree, rows, root.marked_rows, None, focus_day)
        logger.debug("Home Form grid refreshed")

    # Store go_to_today in root for external access
    root.refresh_home = go_to_today

    today_date = datetime.today().strftime("%d/%m/%Y")
    current = datetime.today()

    today_button = tk.Button(   root, text=f"Today: {today_date}", font=(config.ha_head12), bg=COLORS["pale_green"],
                                width=sc(200), height=sc(35), command=go_to_today)
    today_button.place(x=sc(690), y=sc(10), width=sc(200), height=sc(35))

    gear_btn = tk.Button(   root, text="⚙", font=(config.ha_help), width=sc(2), height=1, relief="raised", bg=COLORS["pale_green"],
                            command=lambda: show_maint_toolbox(root, conn, cursor))
    gear_btn.place(x=sc(10), y=sc(10))

    help_btn = tk.Button(   root, text="?", font=(config.ha_help), width=sc(2), height=1, relief="raised", bg=COLORS["pale_green"],
                            command=lambda: show_user_guide())
    help_btn.place(x=sc(1850), y=sc(10))

    def exit_program():
        close_form_with_position(root, conn, cursor, win_id)

    exit_button = tk.Button(root, text="EXIT PROGRAM", font=(config.ha_head14), 
                            bg=COLORS["pale_blue"], width=sc(200), height=sc(35), command=exit_program)
    exit_button.place(x=sc(1180), y=sc(10), width=sc(210), height=sc(35))

    cursor.execute( "SELECT Acc_ID, Acc_Short_Name, Acc_Open, Acc_Jan, Acc_Feb, Acc_Mar, Acc_Apr, Acc_May, Acc_Jun, "
                    "Acc_Jul, Acc_Aug, Acc_Sep, Acc_Oct, Acc_Nov, Acc_Dec FROM Account WHERE Acc_Year = ? ORDER BY Acc_ID", 
                    (int(root.year_var.get()),))
    root.account_data = cursor.fetchall()
    root.accounts = [row[1] for row in account_data] if account_data else ["Cash", "Bank 1", "Bank 2", "Bank 3", "Bank 4", "Bank 5", "CC 1", "CC 2", "CC 3", "CC 4", "CC 5", "Savings", "Investments", "Loans"]
    
    grid_labels = root.accounts
    headers = ["Day", "Date", "R", "Income", "Expense", "Description / Payee"] + grid_labels

    # Account Header Block - AHB
    root.button_grid = []
    btn_width = sc(85)
    btn_height = sc(23)
    for row in range(9):
        row_buttons = []
        for col in range(14):
            x_pos = sc(200) + col * btn_width
            y_pos = sc(50) + row * btn_height
            if row == 0:
                btn = tk.Button(root, text=grid_labels[col] if col < len(grid_labels) else f"Col{col+1}", 
                                font=(config.ha_normal_bold), bg=COLORS["title1_bg"], fg=COLORS["title1_tx"],
                                command=lambda c=col: on_ahb_acc_button(c))
            elif row == 3:
                btn = tk.Button(root, text=f"R{row+1}C{col+1}", font=(config.ha_normal), 
                                bg=COLORS["act_but_bg"], command=lambda c=col: on_ahb_today_button(c))
            else:
                btn = tk.Button(root, text="", font=(config.ha_normal), 
                                bg=COLORS["white"], state="disabled", disabledforeground=COLORS["black"])
                if row == 4:  # Processing Transactions
                    btn.config(fg=COLORS["pending_tx"])
                elif row == 6:  # Forecast Transactions
                    btn.config(fg=COLORS["forecast_tx"])
                elif row == 8:  # Last Statement Balance row
                    btn.config(bg=COLORS["last_stat_bg"])
            btn.place(x=x_pos, y=y_pos, width=btn_width, height=btn_height)
            row_buttons.append(btn)
        root.button_grid.append(row_buttons)

    grid_label_texts = [
        "Start of Month Balance:", "Completed Transactions:", "Account Balance Today:", 
        "Processing Transactions:", "Processing Balance Today:", "Forecast Transactions:", 
        "Forecast EOM Balance:", "Last Statement Balance:"
    ]

    for i, text in enumerate(grid_label_texts):
        label = tk.Label(root, text=text, font=(config.ha_normal), bg=config.master_bg, anchor="e")
        if i == 3:  # Processing Transactions
            label.config(fg=COLORS["pending_tx"])
        elif i == 5:  # Forecast Transactions
            label.config(fg=COLORS["forecast_tx"])
        label.place(x=sc(10), y=sc(50) + (i + 1) * btn_height, width=sc(185), height=btn_height)

    #  End of Month Summary Block - EOMSB
    eomsbx = sc(1400)
    eomsby = sc(50)
    eom_summary_label = ttk.Label(root, text="End of Month Summary", font=(config.ha_head14), background=config.master_bg)
    eom_summary_label.place(x = eomsbx + sc(190), y = eomsby)

    grid_labels2 = "Cash", "Credit", "Invest", "Loans", "Net"
    root.button_grid2 = []
    btn_width2 = sc(86)
    btn_height2 = sc(35)
    for row2 in range(4):
        row_buttons2 = []
        for col2 in range(5):
            x_pos = eomsbx + sc(75) + col2 * btn_width2
            y_pos = eomsby + sc(30) + row2 * btn_height2
            if row2 == 0:
                btn = tk.Button(root, text=grid_labels2[col2] if col2 < len(grid_labels2) else f"Col{col+1}", font=(config.ha_head12), 
                        bg=COLORS["title1_bg"], state="disabled", disabledforeground=COLORS["title1_tx"])
            else:
                btn = tk.Button(root, text=f"R{row2+1}C{col2+1}", font=(config.ha_button), 
                        bg=COLORS["white"], state="disabled", disabledforeground=COLORS["black"])
            btn.place(x=x_pos, y=y_pos, width=btn_width2, height=btn_height2)
            row_buttons2.append(btn)
        root.button_grid2.append(row_buttons2)

    grid_label_texts2 = [
        "SOM Bal:", "EOM Bal:", "Change:"
    ]

    for i, text in enumerate(grid_label_texts2):
        label = tk.Label(root, text=text, font=(config.ha_normal), bg=config.master_bg, anchor="e")
        label.place(x = eomsbx + sc(10), y = eomsby + sc(30) + (i + 1) * btn_height2, width=sc(60), height=btn_height2)

    style = ttk.Style()
    style.theme_use("winnative")  # Force a theme 
    style.configure("TNotebook", background=config.master_bg)
    style.configure("TNotebook.Tab", font=(config.ha_head12), padding=[sc(26), sc(2)], background=COLORS["tab_bg"])

    # Map active/inactive states
    style.map("TNotebook.Tab",
            background=[("selected", COLORS["tab_act_bg"]), ("!selected", COLORS["tab_bg"])],
            foreground=[("selected", COLORS["tab_act_tx"]), ("!selected", COLORS["tab_tx"])])

    # Frame to constrain notebook width
    notebook_frame = tk.Frame(root, background=config.master_bg)
    notebook_frame.pack(pady=(sc(260), 0), padx=sc(10), fill="x")  # Keep fill="x" for now

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
    tree_frame.pack(padx=(sc(10), 0), pady=0, anchor="nw") 

    # Define Treeview styles
    style = ttk.Style()
    style.configure("Treeview.Heading", font=(config.ha_head11))  # Can set font for header row
    style.configure("Treeview", rowheight=sc(18), background="title2_bg")
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

    w1 = sc(83)
    widths = [sc(40), sc(40), sc(20), sc(80), sc(80), sc(250),
            w1, w1, w1, w1, w1, w1, w1, w1, w1, w1, w1, w1, w1, w1]
    anchors = ["w", "w", "w", "e", "e", "w", "e", "e", "e", "e", "e", "e", "e", "e", "e", "e", "e", "e", "e", "e"]
    for i, (header, width, anchor) in enumerate(zip(headers, widths, anchors)):
        root.tree.heading(f"col{i}", text=header, anchor="center")
        root.tree.column(f"col{i}", width=width, anchor=anchor, stretch=False)

    # Define Total row tags
    root.tree.tag_configure("overdrawn", background=COLORS["dtot_ol_bg"], foreground=COLORS["dtot_ol_tx"], font=(config.ha_head11))
    root.tree.tag_configure("daily_total", background=COLORS["dtot_bg"], foreground=COLORS["dtot_tx"], font=(config.ha_normal))

    # Define non-Total row tags
    root.tree.tag_configure("weekday", background=COLORS["tran_wk_bg"])
    root.tree.tag_configure("weekend", background=COLORS["tran_we_bg"])
    root.tree.tag_configure("marked", background=COLORS["flag_mk_bg"])
    root.tree.tag_configure("yellow", background=COLORS["flag_y_bg"])
    root.tree.tag_configure("green", background=COLORS["flag_g_bg"])
    root.tree.tag_configure("cyan", background=COLORS["flag_b_bg"])
    root.tree.tag_configure("orange", background=COLORS["flag_dd_bg"])
    root.tree.tag_configure("forecast", foreground=COLORS["forecast_tx"], font=(config.ha_normal))
    root.tree.tag_configure("processing", foreground=COLORS["pending_tx"], font=(config.ha_normal))
    root.tree.tag_configure("complete", foreground=COLORS["complete_tx"], font=(config.ha_normal))

    # Focal point markers at row 12
    marker_x = sc(280)
    marker_y = sc(18)
    left_canvas = tk.Canvas(root, bg=config.master_bg, highlightthickness=0, width=sc(10), height=sc(40))
    left_canvas.place(x=0, y=sc(218) + marker_x)
    left_canvas.create_text(5, marker_y, text=">", font=(config.ha_head14), fill="red", anchor="center")
    right_canvas = tk.Canvas(root, bg=config.master_bg, highlightthickness=0, width=sc(20), height=sc(40))
    right_canvas.place(x=sc(1702), y=sc(218) + marker_x)
    right_canvas.create_text(5, marker_y, text="<", font=(config.ha_head14), fill="red", anchor="center")

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
        form.geometry(f"{sc(1800)}x{sc(1040)}")  # Adjust size 
        #form.geometry("1800x1170")
        tk.Label(form, text="Annual Budget Performance - Under Construction", font=(config.ha_large)).pack(pady=20)
        tk.Button(  form, text="Close", width=15, background=COLORS["exit_but_bg"],
                    command=lambda: close_form_with_position(form, conn, cursor, win_id)).pack(pady=10)
        form.wait_window()

    def create_compare_form(parent, conn, cursor):  # PLACEHOLDER ***************
        form = tk.Toplevel(parent)
        win_id = 13
        open_form_with_position(form, conn, cursor, win_id, "Compare Current Year to Previous Year")
        form.geometry(f"{sc(800)}x{sc(600)}")  # Adjust size 
        #form.geometry("800x860")
        tk.Label(form, text="Compare Current Year to Previous Year - Under Construction", font=(config.ha_large)).pack(pady=20)
        tk.Button(  form, text="Close", width=15, background=COLORS["exit_but_bg"],
                    command=lambda: close_form_with_position(form, conn, cursor, win_id)).pack(pady=10)
        form.wait_window()

    def create_account_list_form(parent, conn, cursor, acc):  # PLACEHOLDER ***************
        form = tk.Toplevel(parent)
        win_id = 16
        open_form_with_position(form, conn, cursor, win_id, "List all Transactions for a single Account")
        form.geometry(f"{sc(1020)}x{sc(1010)}")  # Adjust size 
        #form.geometry("1020x1010")
        tk.Label(form, text="List all Transactions for a single Account - Under Construction", font=(config.ha_large)).pack(pady=20)
        tk.Label(form, text=grid_labels[acc], font=(config.ha_large)).pack(pady=20)
        tk.Button(  form, text="Close", width=15, background=COLORS["exit_but_bg"],
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
        #logger.debug(f"Updated transaction sum, tc_forecast={tc_forecast}")

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
                root.button_grid[8][col].place(x=sc((200 + col * 85)), y=sc((50 + 8 * 23)), width=sc(85), height=sc(23))
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
            if root.button_grid[3][col].cget("bg") == COLORS["act_but_hi_bg"]:
                root.button_grid[3][col].config(bg=COLORS["act_but_bg"])
            else:
                root.button_grid[3][col].config(bg=COLORS["act_but_hi_bg"])

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
            if month_idx == sc(current.month) and sc(root.year_var.get()) == sc(current.year):
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
    top_button = tk.Button(root, text="T", font=(config.ha_head14), width=2, height=1, command=lambda: root.tree.yview_moveto(0))
    top_button.place(x=sc(1705), y=sc(293))  # Top aligns with tree_frame top

    bottom_button = tk.Button(root, text="B", font=(config.ha_head14), width=2, height=1, command=lambda: root.tree.yview_moveto(1))
    bottom_button.place(x=sc(1705), y=sc(1003))  # Bottom aligns with treeview bottom (42 rows * ~25px) 

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
        selected = root.selected_row.get()
        if selected >= 0 and selected < len(root.rows_container[0]) and root.rows_container[0][selected]["status"] != "Total":
            tr_id = root.rows_container[0][selected]["tr_id"]
            status_map = {"Complete": 3, "Processing": 2, "Forecast": 1}
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
        logger.debug(f"ff_checkbox_var before toggle: {ff_checkbox_var.get()}")
        if ff_checkbox_var.get() == 0:
            ff_checkbox_var.set(1)     # Set checkbox
            ff_box.config(image=root.checked_img)
            cursor.execute("UPDATE Lookups SET Lup_Seq = 1 WHERE Lup_LupT_ID=8")
            conn.commit()
            logger.debug("Lup_Seq = 1")
        else:
            ff_checkbox_var.set(0)    # clear checkbox
            ff_box.config(image=root.unchecked_img)
            cursor.execute("UPDATE Lookups SET Lup_Seq = 0 WHERE Lup_LupT_ID=8")
            conn.commit()
            logger.debug("Lup_Seq = 0")
        logger.debug(f"ff_checkbox_var after toggle: {ff_checkbox_var.get()}")

        # Fetch fresh rows with updated ff_flag
        month_idx = root.get_current_month()        # gets currently selected tab
        root.rows_container[0] = fetch_month_rows(cursor, month_idx, int(root.year_var.get()), root.accounts, root.account_data)
        # Refresh grid with current focus
        current = datetime.now()
        focus_day = str(current.day) if month_idx == current.month and int(root.year_var.get()) == current.year else "1"
        refresh_grid(root.tree, root.rows_container[0], root.marked_rows, focus_day=focus_day)

    # NEW & EDIT Buttons
    tk.Button(  root, text="NEW", font=(config.ha_vlarge), width=7, height=3,
                command=lambda: [root.selected_row.set(-1), create_edit_form(root, root.rows_container[0], root.tree, fetch_month_rows, None, conn, 
                    cursor, get_current_month(), int(root.year_var.get()), local_accounts=accounts, account_data=account_data), 
                    update_ahb(root, get_current_month(), int(root.year_var.get()))]).place(x=sc(1820), y=sc(350))
    tk.Button(  root, text="EDIT", font=(config.ha_vlarge), width=7, height=3,
                command=lambda: create_edit_form(root, root.rows_container[0], root.tree, fetch_month_rows, 
                    root.selected_row.get(), conn, cursor, root.get_current_month(), 
                    int(root.year_var.get()), local_accounts=root.accounts, 
                    account_data=root.account_data) if root.selected_row.get() >= 0 and root.selected_row.get() < len(root.rows_container[0]) else 
                    messagebox.showinfo("HA2", "Please select a row to edit!")).place(x=sc(1715), y=sc(350))

    # Update Transaction Status Buttons
    btn_frame2 = tk.LabelFrame(root, text="Update Status", font=(config.ha_button), bg=config.master_bg, padx=sc(23), pady=sc(10))
    btn_frame2.place(x=sc(1715), y=sc(445))
    tk.Button(  btn_frame2, text="Complete", font=(config.ha_button), fg=COLORS["complete_tx"], width=14, height=1,
                command=lambda: set_status("Complete")).pack(side="top", padx=sc(5), pady=sc(3))
    tk.Button(  btn_frame2, text="Processing", font=(config.ha_button), fg=COLORS["pending_tx"], width=14, height=1,
                command=lambda: set_status("Processing")).pack(side="top", padx=sc(5), pady=sc(3))
    tk.Button(  btn_frame2, text="Forecast", font=(config.ha_button), fg=COLORS["forecast_tx"], width=14, height=1,
                command=lambda: set_status("Forecast")).pack(side="top", padx=sc(5), pady=sc(3))
    
    # Set/Clear Flag Buttons
    flag_frame = tk.LabelFrame(root, text="Set/Clear Flag", font=(config.ha_button), bg=config.master_bg, padx=sc(5), pady=sc(10))
    flag_frame.place(x=sc(1715), y=sc(600))
    tk.Button(flag_frame, text="Set", width=6, bg=COLORS["flag_y_bg"], command=lambda: set_flag(1)).grid(row=1, column=0, padx=sc(4))
    tk.Button(flag_frame, text="Set", width=6, bg=COLORS["flag_g_bg"], command=lambda: set_flag(2)).grid(row=1, column=1, padx=sc(4))
    tk.Button(flag_frame, text="Set", width=6, bg=COLORS["flag_b_bg"], command=lambda: set_flag(4)).grid(row=1, column=2, padx=sc(4))
    tk.Button(flag_frame, text="Clear Flag", width=23, command=clear_flag).grid(row=2, column=0, columnspan=3, padx=sc(4), pady=sc(5))

    # Row Marker Buttons
    btn_frame3 = tk.LabelFrame(root, text="Set/Clear Row Marker", font=(config.ha_button), bg=config.master_bg, padx=sc(5), pady=sc(5))
    btn_frame3.place(x=sc(1715), y=sc(710))
    tk.Button(btn_frame3, text="Set/Clear Row Marker", width=23, bg=COLORS["flag_mk_bg"], command=set_marker).pack(side="top", padx=sc(5), pady=sc(5))
    tk.Button(btn_frame3, text="Clear ALL Row Markers", width=23, command=clear_all_markers).pack(side="top", padx=sc(5), pady=sc(10))

    # Transaction Count Frame
    tc_frame = tk.LabelFrame(root, text="Transaction Count", font=(config.ha_button), bg=config.master_bg, padx=sc(5), pady=sc(5))
    tc_frame.place(x=sc(1715), y=sc(830), width=sc(192), height=sc(130))
    tk.Label(tc_frame, text="   Forecast:", anchor=tk.W, width=16, font=(config.ha_button), 
                bg=config.master_bg, fg=COLORS["forecast_tx"]).place(x=sc(10), y=0)
    tk.Label(tc_frame, textvariable=tc_forecast_var, anchor=tk.E, width=4, font=(config.ha_button), 
                bg=config.master_bg, fg=COLORS["forecast_tx"]).place(x=sc(120), y=0)
    tk.Label(tc_frame, text="Processing:", anchor=tk.W, width=16, font=(config.ha_button), 
                bg=config.master_bg, fg=COLORS["pending_tx"]).place(x=sc(10), y=sc(25))
    tk.Label(tc_frame, textvariable=tc_processing_var, anchor=tk.E, width=4, font=(config.ha_button), 
                bg=config.master_bg, fg=COLORS["pending_tx"]).place(x=sc(120), y=sc(25))
    tk.Label(tc_frame, text="   Complete:", anchor=tk.W, width=16, font=(config.ha_button), 
                bg=config.master_bg, fg=COLORS["complete_tx"]).place(x=sc(10), y=sc(50))
    tk.Label(tc_frame, textvariable=tc_complete_var, anchor=tk.E, width=4, font=(config.ha_button), 
                bg=config.master_bg, fg=COLORS["complete_tx"]).place(x=sc(120), y=sc(50))
    tk.Label(tc_frame, text="          Total:", anchor=tk.W, width=16, font=(config.ha_button), 
                bg=config.master_bg, fg=COLORS["black"]).place(x=sc(10), y=sc(75))
    tk.Label(tc_frame, textvariable=tc_total_var, anchor=tk.E, width=4, font=(config.ha_button), 
                bg=config.master_bg, fg=COLORS["black"]).place(x=sc(120), y=sc(75))

    # Display GC ID# Control and Initialise
    tk.Label(root, text="Show GC ID in list", anchor=tk.W, width=sc(14), font=(config.ha_button), 
            bg=config.master_bg, fg=COLORS["black"]).place(x=sc(1750), y=sc(970))
    ff_box=tk.Button(root, image=root.unchecked_img, bg=config.master_bg, command=toggle_ffid)
    ff_box.place(x=sc(1720), y=sc(970))
    
    # Initialise GC ID 'checkbox' image
    cursor.execute("SELECT Lup_seq FROM Lookups WHERE Lup_LupT_ID = 8")
    ff_rows = cursor.fetchone()
    ff_flag = ff_rows[0]
    ff_checkbox_var.set(int(ff_flag))
    if ff_flag == 1:
        ff_box.config(image=root.checked_img)
    else:
        ff_box.config(image=root.unchecked_img)

#    logger.debug(f"ff checkbox initialised to: {ff_checkbox_var.get()}")

    def on_mouse_wheel(event):
        root.tree.yview_scroll(-1 * (event.delta // 120), "units")

    root.bind("<MouseWheel>", on_mouse_wheel)

    root.mainloop()

if __name__ == "__main__":
    main()













