# Creates main form for the HA2 program

import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import os
import sys
from datetime import timedelta
import ctypes
from ctypes import wintypes
import webbrowser
import logging
from gui import (show_maint_toolbox, create_edit_form)
from db import  (open_db, close_db, fetch_month_rows, fetch_transaction_sums, fetch_statement_balances, fetch_years,
                update_account_year_transactions, fetch_account_transactions)
from ui_utils import (refresh_grid, open_form_with_position, close_form_with_position, resource_path, set_sc, sc, Tooltip)
from focus_forms import (create_summary_form)
from config import (COLORS, CONFIG, init_config, load_colors_from_db, update_master_bg)
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
    screen_height = sc(1045) - 2    # 1045 excludes the top title bar, so this is actually 1080 - reduce height slightly to allow task bar popup to trigger
    root.geometry(f"{sc(1920)}x{screen_height}")  # Adjusted height
    root.configure(bg=config.master_bg)
    root.resizable(False,False)
    root.protocol("WM_DELETE_WINDOW", lambda: None)
    
    # Initialize fullscreen state
    is_fullscreen = tk.BooleanVar(value=False)  # Start in windowed mode
    original_overrideredirect = False
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

    def get_current_monitor_bounds(root):
        """Get the bounds (left, top, width, height) of the monitor containing the window's center."""
        logger.debug("get_current_monitor_bounds")
        # Define Windows structures
        class RECT(ctypes.Structure):
            _fields_ = [('left', ctypes.c_long),
                        ('top', ctypes.c_long),
                        ('right', ctypes.c_long),
                        ('bottom', ctypes.c_long)]

        class MONITORINFO(ctypes.Structure):
            _fields_ = [('cbSize', wintypes.DWORD),
                        ('rcMonitor', RECT),
                        ('rcWork', RECT),
                        ('dwFlags', wintypes.DWORD)]

        user32 = ctypes.windll.user32
        #scaling_factor = root.winfo_fpixels('1i') / 96.0  # DPI scaling (2.0 from logs)
        #logger.debug(f"DPI scaling factor: {scaling_factor}")

        # Get window center position (unscaled for Windows API)
        win_x = root.winfo_x()
        win_y = root.winfo_y()
        win_width = root.winfo_width()
        win_height = root.winfo_height()
        win_center_x = win_x + (win_width // 2)
        win_center_y = win_y + (win_height // 2)
        logger.debug(f"Window: ({win_x}, {win_y}), size: {win_width}x{win_height}")
        logger.debug(f"Window center: ({win_center_x}, {win_center_y}), size: {win_width}x{win_height}")

        monitors = []

        # Callback for EnumDisplayMonitors
        def monitor_enum_proc(hMonitor, hdcMonitor, lprcMonitor, dwData):
            mi = MONITORINFO()
            mi.cbSize = ctypes.sizeof(MONITORINFO)
            if user32.GetMonitorInfoW(hMonitor, ctypes.byref(mi)):
                left = mi.rcMonitor.left
                top = mi.rcMonitor.top
                width = mi.rcMonitor.right - left
                height = mi.rcMonitor.bottom - top
                # Check if window center is within this monitor
                if (left <= win_center_x < left + width and
                    top <= win_center_y < top + height):
                    monitors.append((left, top, width, height))
                logger.debug(f"Monitor found: ({left}, {top}, {width}x{height})")
            return True

        # Enumerate monitors
        MonitorEnumProc = ctypes.WINFUNCTYPE(ctypes.c_int, wintypes.HMONITOR, wintypes.HDC, ctypes.POINTER(RECT), wintypes.LPARAM)
        enum_proc = MonitorEnumProc(monitor_enum_proc)
        user32.EnumDisplayMonitors(None, None, enum_proc, 0)

        if monitors:
            # Prefer the monitor with the window's center
            for left, top, width, height in monitors:
                if left < 100:  # Prioritize secondary monitor (e.g., -1920,0 unscaled)
                    return (left, top, width, height)
            # Fallback to first monitor if no negative coordinates
            left, top, width, height = monitors[0]
            return (left, top, width, height)
        else:
            # Fallback to primary monitor
            primary_width = root.winfo_screenwidth()
            primary_height = root.winfo_screenheight()
            logger.warning(f"No monitor matched window center ({win_center_x}, {win_center_y}), using primary: (0, 0, {primary_width}x{primary_height})")
            return (0, 0, primary_width, primary_height)

    # Add tooltip handling in root
    ttk.Style().configure("Toolbutton.TButton", font=(config.ha_head14), background=COLORS["act_but_bg"])
    
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
        root.rows_container[0] = fetch_month_rows(cursor, current.month, current.year, root.accounts, root.account_data)
        refresh_grid(root.tree, root.rows_container[0], root.marked_rows, None, focus_day)
        root.update_ahb(root, current.month, current.year)  # Pass root explicitly
        logger.debug("Home Form grid refreshed")

    # Routine to swap between fullscreen and windowed modes

    def toggle_fullscreen():
        nonlocal is_fullscreen, original_overrideredirect
        
        if is_fullscreen.get():
            # Exit fullscreen: Restore original state
            logger.debug(f"  >> Current geometry before toggle to windowed: {root.geometry()}")    
            root.overrideredirect(False)
            root.update_idletasks()  # Ensure border is restored
            
            # Get current monitor bounds (scaled)
            left, top, width, height = get_current_monitor_bounds(root)
            logger.debug(f"Current monitor bounds: l={left}, t={top}, w={width}, h={height}")  
            
            # Allow for tk titlebar at top (35px unscaled) and taskbar popup detection at bottom (1px unscaled)
            height = height - sc(36)
            # Adjust for Tkinter's left shift at x=0
            left = left - sc(6)
            
            root.geometry(f"{width}x{height}+{left}+{top}")
            root.update_idletasks()  # Ensure geometry is applied
            is_fullscreen.set(False)
            fullscreen_btn.config(text="⛶")  # Expand icon
            logger.debug(f"Exited fullscreen, restored geometry: {width}x{height}+{left}+{top}")
            logger.debug(f"  >> Current geometry after toggle to windowed: {root.geometry()}")
            
            # Move bottom button up 35px and contract treeview likewise
            bottom_button_y = bottom_button.winfo_y()
            bottom_button.place(y=bottom_button_y - sc(35))
            tree_frame.place(height=sc(750))
            
        else:
            # Enter fullscreen: Save state and cover current monitor
            #original_geometry = root.geometry()
            logger.debug(f"  >> Current geometry before toggle to fullscreen: {root.geometry()}")    
                    
            # Get current monitor bounds (scaled)
            left, top, width, height = get_current_monitor_bounds(root)
            logger.debug(f"Current monitor bounds: l={left}, t={top}, w={width}, h={height}")    
            
            # Apply borderless fullscreen on current monitor
            root.overrideredirect(True)
            root.geometry(f"{width}x{height}+{left}+{top}")
            root.update_idletasks()
            is_fullscreen.set(True)
            fullscreen_btn.config(text="⧉")  # Windowed icon
            logger.debug(f"Entered fullscreen on monitor at {left},{top} ({width}x{height})")
            logger.debug(f"  >> Current geometry after toggle to fullscreen: {root.geometry()}")
            
            # Move bottom button down 35px and expand treeview likewise
            bottom_button_y = bottom_button.winfo_y()
            bottom_button.place(y=bottom_button_y + sc(35))
            tree_frame.place(height=sc(785))
            
    # Store go_to_today in root for external access
    root.refresh_home = go_to_today
    
    # Today button
    today_date = datetime.today().strftime("%d/%m/%Y")
    current = datetime.today()
    today_button = tk.Button(root, text=f"Today: {today_date}", font=(config.ha_head12), bg=COLORS["act_but_bg"],
                                width=sc(200), height=sc(35), command=lambda: go_to_today())
    today_button.place(x=sc(690), y=sc(10), width=sc(200), height=sc(35))
    Tooltip(today_button, "Go to today's tab")

    # Settings button
    gear_btn = tk.Button(root, text="⚙", font=(config.ha_help), width=sc(2), height=1, relief="raised", bg=COLORS["act_but_bg"],
                            command=lambda: show_maint_toolbox(root, conn, cursor, refresh_colors))
    gear_btn.place(x=sc(10), y=sc(10))
    Tooltip(gear_btn, "Open Settings")
    
    # Fullscreen button
    fullscreen_btn = tk.Button(root, text="⛶", font=(config.ha_help), width=sc(2), height=1, relief="raised", bg=COLORS["act_but_bg"],
                            command=lambda: toggle_fullscreen())
    fullscreen_btn.place(x=sc(1780), y=sc(10))
    Tooltip(fullscreen_btn, "Toggle full-screen/windowed mode")    
    
    # Help Manual button
    help_btn = tk.Button(root, text="?", font=(config.ha_help), width=sc(2), height=1, relief="raised", bg=COLORS["act_but_bg"],
                            command=lambda: show_user_guide())
    help_btn.place(x=sc(1850), y=sc(10))
    Tooltip(help_btn, "Open User Guide")

    def exit_program():
        close_form_with_position(root, conn, cursor, win_id)

    exit_button = tk.Button(root, text="EXIT PROGRAM", font=(config.ha_head14), 
                            bg=COLORS["exit_but_bg"], width=sc(200), height=sc(35), command=exit_program)
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
                                font=(config.ha_normal_bold), bg=COLORS["act_but_bg"], fg=COLORS["act_but_tx"],
                                command=lambda c=col: on_ahb_acc_button(c))
                Tooltip(btn, "Show Account Transactions")
            elif row == 3:
                btn = tk.Button(root, text=f"R{row+1}C{col+1}", font=(config.ha_normal), 
                                bg=COLORS["act_but_bg"], command=lambda c=col: on_ahb_today_button(c))
            else:
                btn = tk.Button(root, text="", font=(config.ha_normal), 
                                bg=COLORS["white"], state="disabled", disabledforeground=COLORS["black"])
                if row == 4:  # Processing Transactions
                    btn.config(disabledforeground=COLORS["pending_tx"])
                elif row == 6:  # Forecast Transactions
                    btn.config(disabledforeground=COLORS["forecast_tx"])
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
    grid_labels = []

    for i, text in enumerate(grid_label_texts):
        label = tk.Label(root, text=text, font=(config.ha_normal), bg=config.master_bg, anchor="e")
        if i == 3:  # Processing Transactions
            label.config(fg=COLORS["pending_tx"])
        elif i == 5:  # Forecast Transactions
            label.config(fg=COLORS["forecast_tx"])
        label.place(x=sc(10), y=sc(50) + (i + 1) * btn_height, width=sc(185), height=btn_height)
        grid_labels.append(label)

    #  End of Month Summary Block - EOMSB
    eomsbx = sc(1400)
    eomsby = sc(65)
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
    grid_labels2 = []
    
    for i, text in enumerate(grid_label_texts2):
        label = tk.Label(root, text=text, font=(config.ha_normal), bg=config.master_bg, anchor="e")
        label.place(x = eomsbx + sc(10), y = eomsby + sc(30) + (i + 1) * btn_height2, width=sc(60), height=btn_height2)
        grid_labels2.append(label)

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
    notebook_frame.place(x=sc(10), y=sc(260), width=sc(1690) ) 
    #notebook_frame.pack(pady=(sc(260), 0), padx=sc(10), fill="x")  # Keep fill="x" for now

    notebook = ttk.Notebook(notebook_frame)
    notebook.pack(side="left", fill="none")  # No fill, let tabs define width

    tabs = {
        "SUMMARY": ttk.Frame(notebook), " BUDGET ": ttk.Frame(notebook), "PREV": ttk.Frame(notebook),
        " JAN": ttk.Frame(notebook), " FEB": ttk.Frame(notebook), " MAR": ttk.Frame(notebook),
        " APR": ttk.Frame(notebook), " MAY": ttk.Frame(notebook), " JUN": ttk.Frame(notebook),
        " JUL": ttk.Frame(notebook), " AUG": ttk.Frame(notebook), " SEP": ttk.Frame(notebook),
        " OCT": ttk.Frame(notebook), " NOV": ttk.Frame(notebook), " DEC": ttk.Frame(notebook),
        "NEXT": ttk.Frame(notebook), " COMPARE ": ttk.Frame(notebook)
    }
    tab_names = ["SUMMARY", " BUDGET ", "PREV", " JAN", " FEB", " MAR", " APR", " MAY", " JUN", " JUL", " AUG", " SEP", " OCT", " NOV", " DEC", "NEXT", " COMPARE "]
    for tab_name in tab_names:
        notebook.add(tabs[tab_name], text=tab_name)

    root.tab_change_handled = False  # Add this flag

    # Main Transaction List - Treeview
    tree_frame = tk.Frame(root)
    tree_frame.place(x=sc(10), y=sc(288), width=sc(1690), height=sc(750))
    #tree_frame.pack(padx=(sc(10), 0), pady=0, anchor="nw") 

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
    root.tree.tag_configure("overlimit", background=COLORS["dtot_ol_bg"], foreground=COLORS["dtot_ol_tx"], font=(config.ha_head11))
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
    month_map_num = {1: " JAN", 2: " FEB", 3: " MAR", 4: " APR", 5: " MAY", 6: " JUN", 7: " JUL", 8: " AUG", 9: " SEP", 10: " OCT", 11: " NOV", 12: " DEC"}
    root.rows_container[0] = fetch_month_rows(cursor, current.month, current.year, root.accounts, root.account_data)
    notebook.select(tabs[month_map_num[current.month]])
    month_map_txt = {" JAN": 1, " FEB": 2, " MAR": 3, " APR": 4, " MAY": 5, " JUN": 6, " JUL": 7, " AUG": 8, " SEP": 9, " OCT": 10, " NOV": 11, " DEC": 12}

    def get_current_month():
        selected_tab = notebook.tab(notebook.select(), "text")
        return month_map_txt.get(selected_tab, 3)
    
    root.get_current_month = get_current_month  # Attach to root    

    def create_budget_form(parent, conn, cursor):  # PLACEHOLDER ***************
        logger.debug("Create Budget Form")
        form = tk.Toplevel(parent)
        win_id = 12
        open_form_with_position(form, conn, cursor, win_id, "Annual Budget Performance")
        form.geometry(f"{sc(1800)}x{sc(1040)}")  # Adjust size 
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

    def create_account_list_form(parent, conn, cursor, acc):
        """Create a form to list all transactions for a single account for the current year, in date order by month."""
        logger.debug(f"Creating account transaction list form for account index {acc}")
        form = tk.Toplevel(parent)
        win_id = 16
        open_form_with_position(form, conn, cursor, win_id, "List all Transactions for a Single Account")
        form.geometry(f"{sc(1020)}x{sc(1010)}")
        form.configure(bg=config.master_bg) 
        form.resizable(False, False)
        
        # Get account name and ID
        account_name = parent.accounts[acc]
        acc_id = parent.account_data[acc][0]
        acc_limit = parent.account_data[acc][15]
        logger.debug(f"Account: {account_name}, Acc_ID: {acc_id}, Acc_Credit_Limit: {acc_limit}")

        # Initialize selected row
        form.selected_row = tk.IntVar(value=-1)
        focus_row = 0
        
        # Account Name label
        acc_name_label = tk.Label(form, text=account_name, font=(config.ha_head14), 
                                bg=config.master_bg, anchor="center")
        acc_name_label.place(x=sc(800), y=sc(10), width=sc(200), height=sc(30))
        
        # Transaction Count label
        trans_count_var = tk.StringVar(value="Transactions: 0")
        trans_count_label = tk.Label(form, textvariable=trans_count_var, bg=config.master_bg, font=(config.ha_large))     
        trans_count_label.place(x=sc(790), y=sc(50), width=sc(200), height=sc(30))
        
        # Treeview for transactions
        tree_frame = tk.Frame(form, bg=config.master_bg)
        tree_frame.place(x=sc(20), y=sc(10), width=sc(750), height=sc(990))
        columns = ("day", "date", "desc", "income", "expense", "balance")
        tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=49, style="Treeview")
        scrollbar = tk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        tree.pack(side="left", fill="y")
        scrollbar.pack(side="right", fill="y")
        
        # Configure Treeview styles
        style = ttk.Style()
        style.configure("Treeview.Heading", font=(config.ha_head11))
        style.configure("Treeview", rowheight=sc(20), font=(config.ha_normal), background=COLORS["white"])
        
        # Column headings and widths
        headers = ["Day", "Date", "Description / Payee", "Income", "Expense", "Balance"]
        widths = [sc(40), sc(90), sc(300), sc(100), sc(100), sc(100)]
        anchors = ["w", "w", "w", "e", "e", "e"]
        for col, header, width, anchor in zip(columns, headers, widths, anchors):
            tree.heading(col, text=header, anchor="center")
            tree.column(col, width=width, anchor=anchor, stretch=False)
        
        # Define Treeview tags
        tree.tag_configure("weekday", background=COLORS["tran_wk_bg"])
        tree.tag_configure("weekend", background=COLORS["tran_we_bg"])
        tree.tag_configure("forecast", foreground=COLORS["forecast_tx"])
        tree.tag_configure("processing", foreground=COLORS["pending_tx"])
        tree.tag_configure("complete", foreground=COLORS["complete_tx"])
        tree.tag_configure("daily_total", background=COLORS["dtot_bg"], foreground=COLORS["dtot_tx"])
        tree.tag_configure("overlimit", background=COLORS["dtot_ol_bg"], foreground=COLORS["dtot_ol_tx"])
        tree.tag_configure("flag1", background=COLORS["flag_y_bg"])
        tree.tag_configure("flag2", background=COLORS["flag_g_bg"])
        tree.tag_configure("flag3", background=COLORS["flag_b_bg"])
        
        # Fetch FF checkbox setting
        cursor.execute("SELECT Lup_seq FROM Lookups WHERE Lup_LupT_ID = 8")
        ff_flag = cursor.fetchone()[0]
        
        # Fetch transactions
        year = int(parent.year_var.get())
        trans_rows = fetch_account_transactions(cursor, acc_id, year, parent.account_data[acc], ff_flag, acc_limit)
        
        # Populate Treeview
        def refresh_tree():
            nonlocal trans_rows, focus_row
            trans_rows[:] = fetch_account_transactions(cursor, acc_id, year, parent.account_data[acc], ff_flag, acc_limit)
            tree.delete(*tree.get_children())
            trans_count = sum(1 for row in trans_rows if row["status"] != "Total")
            trans_count_var.set(f"Transactions: {trans_count}")
            for i, row in enumerate(trans_rows):
                tags = []
                if row["status"] == "Total":
                    tags.append("daily_total" if not row["overlimit"] else "overlimit")
                else:
                    tags.append("weekend" if row["is_weekend"] else "weekday")
                    if row["status"] == "Forecast":
                        tags.append("forecast")
                    elif row["status"] == "Processing":
                        tags.append("processing")
                    elif row["status"] == "Complete":
                        tags.append("complete")
                    if row["flag"] == 1:
                        tags.append("flag1")
                    elif row["flag"] == 2:
                        tags.append("flag2")
                    elif row["flag"] == 4:
                        tags.append("flag3")
                income_str = ""
                expense_str = ""
                if row["income"] or row["expense"]:
                    if row["exp_id"] == 99:
                        income_str = f"{row['income']:,.2f} ??" if row["income"] else ""
                        expense_str = f"{-row['expense']:,.2f} ??" if row["expense"] else ""
                    else:
                        income_str = f"{row['income']:,.2f}" if row["income"] else ""
                        expense_str = f"{-row['expense']:,.2f}" if row["expense"] else ""
                values = (
                    row["day"],
                    row["date"],
                    row["desc"],
                    income_str,
                    expense_str,
                    "{:,.2f}".format(row["balance"]) if row["status"] == "Total" and row["balance"] else ""
                )
                # Use unique iid: index for non-totals, total_index for totals
                iid = str(i) if row["status"] != "Total" else f"total_{i}"
                tree.insert("", "end", iid=iid, values=values, tags=tags)
            # Scroll to current day or last transaction
            if focus_row > 0:
                if focus_row > 14:
                    i = focus_row-15
                else:
                    i = 0
                tree.yview_moveto((i) / len(trans_rows) if len(trans_rows) > 0 else 0)
                tree.selection_set(str(focus_row))
                form.selected_row.set(focus_row)
            else:
                current = datetime.now()
                if year == current.year:
                    for i, row in enumerate(trans_rows):
                        if row["month"] == current.month and row["num_day"] and row["num_day"] >= current.day:
                            tree.yview_moveto((i-15) / len(trans_rows) if len(trans_rows) > 0 else 0)
                            tree.selection_set(str(i))
                            form.selected_row.set(i)
                            break
        
        # Bind refresh_tree to form for access in create_edit_form
        form.refresh_tree = refresh_tree
        refresh_tree()
        
        # Treeview selection
        def on_tree_select(event):
            selected = tree.selection()
            if selected:
                item = selected[0]
                idx = int(item) if not item.startswith("total_") else -1
                form.selected_row.set(idx)
            else:
                form.selected_row.set(-1)
        
        tree.bind("<<TreeviewSelect>>", on_tree_select)
        
        # Determine the default month (use current month or parent form's selected month)
        default_month = parent.get_current_month() if hasattr(parent, 'get_current_month') else datetime.now().month
        
        def set_focus_row():
            nonlocal focus_row
            focus_row = form.selected_row.get()

        # Edit button
        edit_btn = tk.Button(form, text="EDIT", font=(config.ha_vlarge))
        edit_btn.place(x=sc(795), y=sc(100), width=sc(80), height=sc(80))
        edit_btn.config(command=lambda: [set_focus_row(), create_edit_form(
            form, trans_rows, tree,
            lambda *args: fetch_account_transactions(cursor, acc_id, year, parent.account_data[acc], ff_flag, acc_limit),
            form.selected_row.get(),
            conn, cursor,
            trans_rows[form.selected_row.get()]["month"] if 0 <= form.selected_row.get() < len(trans_rows) and trans_rows[form.selected_row.get()]["status"] != "Total" else default_month,
            year,
            local_accounts=parent.accounts,
            account_data=[parent.account_data[acc]],
            single_acc=False,
            default_acc_id=acc_id,
            home_root=parent
        )])

        # New button 
        new_btn = tk.Button(form, text="NEW", font=(config.ha_vlarge))
        new_btn.place(x=sc(895), y=sc(100), width=sc(80), height=sc(80))
        new_btn.config(command=lambda: [set_focus_row(), create_edit_form(
            form, trans_rows, tree,
            lambda *args: fetch_account_transactions(cursor, acc_id, year, parent.account_data[acc], ff_flag, acc_limit),
            None,
            conn, cursor,
            default_month,
            year,
            local_accounts=parent.accounts,
            account_data=[parent.account_data[acc]],
            single_acc=False,
            default_acc_id=acc_id,
            home_root=parent
        )])
        
        close_btn = tk.Button(form, text="Close", font=(config.ha_head14), bg=COLORS["exit_but_bg"], width=15,
                            command=lambda: close_form_with_position(form, conn, cursor, win_id))
        close_btn.place(x=sc(795), y=sc(800), width=sc(200), height=sc(35))
        
        # Update Status Group
        status_frame = tk.LabelFrame(form, text="     Update Status    ", font=(config.ha_button), bg=config.master_bg)
        status_frame.place(x=sc(795), y=sc(205), width=sc(200), height=sc(140))
        
        def set_status(status_value):
            selected = form.selected_row.get()
            if selected >= 0 and selected < len(trans_rows) and trans_rows[selected]["status"] != "Total":
                tr_id = trans_rows[selected]["tr_id"]
                status_map = {"Complete": 3, "Processing": 2, "Forecast": 1}
                tr_date = datetime.strptime(trans_rows[selected]["date"], "%d/%m/%Y")
                today = datetime.now()
                if status_value == "Complete" and tr_date > today:
                    messagebox.showerror("Error", "Transaction date cannot be in the future if the transaction is Completed", parent=form)
                    return
                if status_value == "Processing" and (tr_date > today or tr_date < today - timedelta(days=10)):
                    messagebox.showerror("Error", "Transaction date must be within last 10 days if the transaction is Processing", parent=form)
                    return
                if status_value == "Forecast" and tr_date < today:
                    messagebox.showerror("Error", "Transaction date must be today or later if the transaction is Forecast", parent=form)
                    return
                cursor.execute("UPDATE Trans SET Tr_Stat = ?, Tr_Query_Flag = CASE WHEN ? IN (2, 3) THEN 0 ELSE Tr_Query_Flag END WHERE Tr_ID = ?", 
                            (status_map[status_value], status_map[status_value], tr_id))
                conn.commit()
                trans_rows[:] = fetch_account_transactions(cursor, acc_id, year, parent.account_data[acc], ff_flag, acc_limit)
                refresh_tree()
                if hasattr(parent, 'update_ahb'):
                    parent.update_ahb(parent, parent.get_current_month(), year)
        
        stat_comp_btn = tk.Button(status_frame, text="Complete", font=(config.ha_button), fg=COLORS["complete_tx"], width=15,
                                command=lambda: set_status("Complete"))
        stat_comp_btn.place(x=sc(25), y=sc(15), width=sc(150), height=sc(25))
        
        stat_pend_btn = tk.Button(status_frame, text="Pending", font=(config.ha_button), fg=COLORS["pending_tx"], width=15,
                                command=lambda: set_status("Processing"))
        stat_pend_btn.place(x=sc(25), y=sc(50), width=sc(150), height=sc(25))
        
        stat_fore_btn = tk.Button(status_frame, text="Forecast", font=(config.ha_button), fg=COLORS["forecast_tx"], width=15,
                                command=lambda: set_status("Forecast"))
        stat_fore_btn.place(x=sc(25), y=sc(85), width=sc(150), height=sc(25))
        
        # Set/Clear Flag Group
        flag_frame = tk.LabelFrame(form, text="    Set/Clear Flag    ", font=(config.ha_button), bg=config.master_bg)
        flag_frame.place(x=sc(795), y=sc(365), width=sc(200), height=sc(100))
        
        def set_flag(flag_value):
            selected = form.selected_row.get()
            if selected >= 0 and selected < len(trans_rows) and trans_rows[selected]["status"] != "Total":
                tr_id = trans_rows[selected]["tr_id"]
                if trans_rows[selected]["status"] == "Complete":
                    messagebox.showerror("Error", "Action Flag cannot be set on a Complete Transaction", parent=form)
                    return
                cursor.execute("UPDATE Trans SET Tr_Query_Flag = ? WHERE Tr_ID = ?", (flag_value, tr_id))
                conn.commit()
                trans_rows[:] = fetch_account_transactions(cursor, acc_id, year, parent.account_data[acc], ff_flag, acc_limit)
                refresh_tree()
                if hasattr(parent, 'update_ahb'):
                    parent.update_ahb(parent, parent.get_current_month(), year)
        
        flag1_btn = tk.Button(flag_frame, text="Set", font=(config.ha_normal), bg=COLORS["flag_y_bg"], width=4,
                            command=lambda: set_flag(1))
        flag1_btn.place(x=sc(25), y=sc(15), width=sc(40), height=sc(25))
        
        flag2_btn = tk.Button(flag_frame, text="Set", font=(config.ha_button), bg=COLORS["flag_g_bg"], width=4,
                            command=lambda: set_flag(2))
        flag2_btn.place(x=sc(80), y=sc(15), width=sc(40), height=sc(25))
        
        flag3_btn = tk.Button(flag_frame, text="Set", font=(config.ha_button), bg=COLORS["flag_b_bg"], width=4,
                            command=lambda: set_flag(4))
        flag3_btn.place(x=sc(135), y=sc(15), width=sc(40), height=sc(25))
        
        clear_flag_btn = tk.Button(flag_frame, text="Clear Flag", font=(config.ha_button), width=15,
                                command=lambda: set_flag(0) if form.selected_row.get() >= 0 and trans_rows[form.selected_row.get()]["status"] != "Total" 
                                else messagebox.showerror("Error", "Cannot update a Monthly Balance row", parent=form))
        clear_flag_btn.place(x=sc(25), y=sc(50), width=sc(150), height=sc(25))
        
        # Scroll Buttons
        top_btn = tk.Button(form, text="T", font=(config.ha_head14), width=2, height=1,
                            command=lambda: tree.yview_moveto(0))
        top_btn.place(x=sc(770), y=sc(10), width=sc(30), height=sc(40))
        
        bottom_btn = tk.Button(form, text="B", font=(config.ha_head14), width=2, height=1,
                            command=lambda: tree.yview_moveto(1))
        bottom_btn.place(x=sc(770), y=sc(960), width=sc(30), height=sc(40))
        
        # Row Markers
        left_marker = tk.Canvas(form, bg=config.master_bg, highlightthickness=0, width=sc(10), height=sc(20))
        left_marker.place(x=sc(5), y=sc(328))
        left_marker.create_text(sc(5), sc(10), text=">", font=(config.ha_head14), fill="red", anchor="center")
        
        right_marker = tk.Canvas(form, bg=config.master_bg, highlightthickness=0, width=sc(10), height=sc(20))
        right_marker.place(x=sc(770), y=sc(328))
        right_marker.create_text(sc(5), sc(10), text="<", font=(config.ha_head14), fill="red", anchor="center")
        
        # Mouse wheel scrolling
        def on_mouse_wheel(event):
            tree.yview_scroll(-1 * (event.delta // 120), "units")
        
        form.bind("<MouseWheel>", on_mouse_wheel)
        
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
            logger.debug(f"month_ix={month_idx}, ")
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
        elif selected_tab == " BUDGET ":
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
        elif selected_tab == " COMPARE ":
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
    new_btn = tk.Button(root, text="NEW", font=(config.ha_vlarge), width=7, height=3,
                command=lambda: [root.selected_row.set(-1), create_edit_form(root, root.rows_container[0], root.tree, fetch_month_rows, None, conn, 
                    cursor, get_current_month(), int(root.year_var.get()), local_accounts=accounts, account_data=account_data), 
                    update_ahb(root, get_current_month(), int(root.year_var.get()))])
    new_btn.place(x=sc(1820), y=sc(350))
    edit_btn = tk.Button(root, text="EDIT", font=(config.ha_vlarge), width=7, height=3,
                command=lambda: create_edit_form(root, root.rows_container[0], root.tree, fetch_month_rows, 
                    root.selected_row.get(), conn, cursor, root.get_current_month(), 
                    int(root.year_var.get()), local_accounts=root.accounts, 
                    account_data=root.account_data) if root.selected_row.get() >= 0 and root.selected_row.get() < len(root.rows_container[0]) else 
                    messagebox.showinfo("HA2", "Please select a row to edit!"))
    edit_btn.place(x=sc(1715), y=sc(350))

    # Update Transaction Status Buttons
    btn_frame2 = tk.LabelFrame(root, text="Update Status", font=(config.ha_button), bg=config.master_bg, padx=sc(23), pady=sc(10))
    btn_frame2.place(x=sc(1715), y=sc(445))
    stat_comp_btn=tk.Button(  btn_frame2, text="Complete", font=(config.ha_button), fg=COLORS["complete_tx"], width=14, height=1,
                command=lambda: set_status("Complete"))
    stat_comp_btn.pack(side="top", padx=sc(5), pady=sc(3))
    stat_pend_btn=tk.Button(  btn_frame2, text="Processing", font=(config.ha_button), fg=COLORS["pending_tx"], width=14, height=1,
                command=lambda: set_status("Processing"))
    stat_pend_btn.pack(side="top", padx=sc(5), pady=sc(3))
    stat_fore_btn=tk.Button(  btn_frame2, text="Forecast", font=(config.ha_button), fg=COLORS["forecast_tx"], width=14, height=1,
                command=lambda: set_status("Forecast"))
    stat_fore_btn.pack(side="top", padx=sc(5), pady=sc(3))
    
    # Set/Clear Flag Buttons
    flag_frame = tk.LabelFrame(root, text="Set/Clear Flag", font=(config.ha_button), bg=config.master_bg, padx=sc(5), pady=sc(10))
    flag_frame.place(x=sc(1715), y=sc(600))
    flag1_btn=tk.Button(flag_frame, text="Set", width=6, font=config.ha_normal, bg=COLORS["flag_y_bg"], command=lambda: set_flag(1))
    flag1_btn.grid(row=1, column=0, padx=sc(4))
    flag2_btn=tk.Button(flag_frame, text="Set", width=6, font=config.ha_normal, bg=COLORS["flag_g_bg"], command=lambda: set_flag(2))
    flag2_btn.grid(row=1, column=1, padx=sc(4))
    flag3_btn=tk.Button(flag_frame, text="Set", width=6, font=config.ha_normal, bg=COLORS["flag_b_bg"], command=lambda: set_flag(4))
    flag3_btn.grid(row=1, column=2, padx=sc(4))
    tk.Button(flag_frame, text="Clear Flag", width=23, command=clear_flag).grid(row=2, column=0, columnspan=3, padx=sc(4), pady=sc(5))

    # Row Marker Buttons
    btn_frame3 = tk.LabelFrame(root, text="Set/Clear Row Marker", font=(config.ha_button), bg=config.master_bg, padx=sc(5), pady=sc(5))
    btn_frame3.place(x=sc(1715), y=sc(710))
    rowmarker_btn=tk.Button(btn_frame3, text="Set/Clear Row Marker", width=23, bg=COLORS["flag_mk_bg"], command=set_marker)
    rowmarker_btn.pack(side="top", padx=sc(5), pady=sc(5))
    tk.Button(btn_frame3, text="Clear ALL Row Markers", width=23, command=clear_all_markers).pack(side="top", padx=sc(5), pady=sc(10))

    # Transaction Count Frame
    tc_frame = tk.LabelFrame(root, text="Transaction Count", font=(config.ha_button), bg=config.master_bg, padx=sc(5), pady=sc(5))
    tc_frame.place(x=sc(1715), y=sc(830), width=sc(192), height=sc(130))
    tc_fore_lbl=tk.Label(tc_frame, text="   Forecast:", anchor=tk.W, width=16, font=(config.ha_button), 
                bg=config.master_bg, fg=COLORS["forecast_tx"])
    tc_fore_lbl.place(x=sc(10), y=0)
    tc_fore_num=tk.Label(tc_frame, textvariable=tc_forecast_var, anchor=tk.E, width=4, font=(config.ha_button), 
                bg=config.master_bg, fg=COLORS["forecast_tx"])
    tc_fore_num.place(x=sc(120), y=0)
    tc_pend_lbl=tk.Label(tc_frame, text="Processing:", anchor=tk.W, width=16, font=(config.ha_button), 
                bg=config.master_bg, fg=COLORS["pending_tx"])
    tc_pend_lbl.place(x=sc(10), y=sc(25))
    tc_pend_num=tk.Label(tc_frame, textvariable=tc_processing_var, anchor=tk.E, width=4, font=(config.ha_button), 
                bg=config.master_bg, fg=COLORS["pending_tx"])
    tc_pend_num.place(x=sc(120), y=sc(25))
    tc_comp_lbl=tk.Label(tc_frame, text="   Complete:", anchor=tk.W, width=16, font=(config.ha_button), 
                bg=config.master_bg, fg=COLORS["complete_tx"])
    tc_comp_lbl.place(x=sc(10), y=sc(50))
    tc_comp_num=tk.Label(tc_frame, textvariable=tc_complete_var, anchor=tk.E, width=4, font=(config.ha_button), 
                bg=config.master_bg, fg=COLORS["complete_tx"])
    tc_comp_num.place(x=sc(120), y=sc(50))
    tc_tot_lbl=tk.Label(tc_frame, text="          Total:", anchor=tk.W, width=16, font=(config.ha_button), 
                bg=config.master_bg, fg=COLORS["black"])
    tc_tot_lbl.place(x=sc(10), y=sc(75))
    tc_tot_num=tk.Label(tc_frame, textvariable=tc_total_var, anchor=tk.E, width=4, font=(config.ha_button), 
                bg=config.master_bg, fg=COLORS["black"])
    tc_tot_num.place(x=sc(120), y=sc(75))

    # Display GC ID# Control and Initialise
    gcid_lbl=tk.Label(root, text="Show GC ID in list", anchor=tk.W, width=sc(14), font=(config.ha_button), 
            bg=config.master_bg, fg=COLORS["black"])
    gcid_lbl.place(x=sc(1750), y=sc(970))
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

    def refresh_colors():       # callback for change of colour scheme so it updates home form
        # Check widget initialization
        logger.debug(f"Widgets initialized: {', '.join(f'{name}={getattr(root, name, None)}' for name in ['today_button', 'gear_btn', 'fullscreen_btn', 'help_btn', 'exit_button', 'new_btn', 'edit_btn', 'rowmarker_btn', 'flag1_btn', 'flag2_btn', 'flag3_btn', 'clear_flag_btn', 'clear_all_markers_btn', 'tc_fore_lbl', 'tc_fore_num', 'tc_pend_lbl', 'tc_pend_num', 'tc_comp_lbl', 'tc_comp_num', 'stat_fore_btn', 'stat_pend_btn', 'stat_comp_btn', 'total_lbl', 'total_num', 'show_gc_id_lbl', 'ff_box'])}")
        
        """Update button grid colors based on COLORS dictionary."""
        if CONFIG['APP_ENV'] == 'test':
            colour = COLORS["home_test_bg"]
        else:
            colour = COLORS["home_bg"]
        update_master_bg(colour)
        
        for row in range(9):
            for col in range(14):
                btn = root.button_grid[row][col]
                if row == 0:
                    btn.config(bg=COLORS["act_but_bg"], fg=COLORS["act_but_tx"])
                elif row == 3:
                    btn.config(bg=COLORS["act_but_bg"])
                elif row == 4:
                    btn.config(bg=COLORS["white"], disabledforeground=COLORS["pending_tx"])
                elif row == 6:
                    btn.config(bg=COLORS["white"], disabledforeground=COLORS["forecast_tx"])
                elif row == 8:
                    btn.config(bg=COLORS["last_stat_bg"], disabledforeground=COLORS["black"])
                else:
                    btn.config(bg=COLORS["white"], disabledforeground=COLORS["black"])

        # update EOMSB title row
        for col in range(5):
            btn = root.button_grid2[0][col]
            btn.config(bg=COLORS["title1_bg"], fg=COLORS["title1_tx"])
            
        def config_widget(widget_name, widget, **kwargs):
            if widget is not None:
                widget.config(**kwargs)
            else:
                logger.warning(f"Widget {widget_name} is None: {kwargs}")    
        
        for i in (0, 1, 2, 3, 4, 5, 6, 7):
            grid_labels[i].config(bg=config.master_bg)
            
        for i in (0, 1, 2):
            grid_labels2[i].config(bg=config.master_bg)

        root.configure(bg=config.master_bg)
        config_widget("today_button", today_button, bg=COLORS["act_but_bg"], fg=COLORS["act_but_tx"])
        config_widget("gear_btn", gear_btn, bg=COLORS["act_but_bg"], fg=COLORS["act_but_tx"])
        config_widget("fullscreen_btn", fullscreen_btn, bg=COLORS["act_but_bg"], fg=COLORS["act_but_tx"])
        config_widget("help_btn", help_btn, bg=COLORS["act_but_bg"], fg=COLORS["act_but_tx"])
        config_widget("exit_button", exit_button, bg=COLORS["exit_but_bg"], fg=COLORS["exit_but_tx"])
        config_widget("rowmarker_btn", rowmarker_btn, bg=COLORS["flag_mk_bg"])
        config_widget("flag1_btn", flag1_btn, bg=COLORS["flag_y_bg"])
        config_widget("flag2_btn", flag2_btn, bg=COLORS["flag_g_bg"])
        config_widget("flag3_btn", flag3_btn, bg=COLORS["flag_b_bg"])
        config_widget("tc_frame", tc_frame, bg=config.master_bg)
        config_widget("tc_fore_lbl", tc_fore_lbl, fg=COLORS["forecast_tx"], bg=config.master_bg)
        config_widget("tc_fore_num", tc_fore_num, fg=COLORS["forecast_tx"], bg=config.master_bg)
        config_widget("tc_pend_lbl", tc_pend_lbl, fg=COLORS["pending_tx"], bg=config.master_bg)
        config_widget("tc_pend_num", tc_pend_num, fg=COLORS["pending_tx"], bg=config.master_bg)
        config_widget("tc_comp_lbl", tc_comp_lbl, fg=COLORS["complete_tx"], bg=config.master_bg)
        config_widget("tc_comp_num", tc_comp_num, fg=COLORS["complete_tx"], bg=config.master_bg)
        config_widget("tc_tot_lbl", tc_tot_lbl, fg=COLORS["black"], bg=config.master_bg)
        config_widget("tc_tot_num", tc_tot_num, fg=COLORS["black"], bg=config.master_bg)
        config_widget("stat_fore_btn", stat_fore_btn, fg=COLORS["forecast_tx"])
        config_widget("stat_pend_btn", stat_pend_btn, fg=COLORS["pending_tx"])
        config_widget("stat_comp_btn", stat_comp_btn, fg=COLORS["complete_tx"])
        config_widget("gcid_lbl", gcid_lbl, bg=config.master_bg)
        config_widget("ff_box", ff_box, bg=config.master_bg)
        config_widget("btn_frame2", btn_frame2, bg=config.master_bg)
        config_widget("btn_frame3", btn_frame3, bg=config.master_bg)
        config_widget("flag_frame", flag_frame, bg=config.master_bg)
        config_widget("ff_box", ff_box, bg=config.master_bg)
        config_widget("year_label", year_label, background=config.master_bg)
        config_widget("eom_summary_label", eom_summary_label, background=config.master_bg)
        config_widget("left_canvas", left_canvas, bg=config.master_bg)
        config_widget("right_canvas", right_canvas, bg=config.master_bg)

        # Re-Map active/inactive states
        style.map("TNotebook.Tab",
            background=[("selected", COLORS["tab_act_bg"]), ("!selected", COLORS["tab_bg"])],
            foreground=[("selected", COLORS["tab_act_tx"]), ("!selected", COLORS["tab_tx"])])
        
        
        # add other widgets here
        
        logger.debug("Home form colors refreshed")
        
    def on_mouse_wheel(event):
        root.tree.yview_scroll(-1 * (event.delta // 120), "units")

    root.bind("<MouseWheel>", on_mouse_wheel)

    root.mainloop()

if __name__ == "__main__":
    main()






































