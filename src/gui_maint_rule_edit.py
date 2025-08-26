# gui_maint_rule_edit.py
# Functions for editing rules, triggers, and actions in the HA2 project

import tkinter as tk
from tkinter import ttk, messagebox
import logging
import traceback
import re
from db import (fetch_rule_group_names, fetch_trigger_options, fetch_action_options, fetch_account_c_names,
                fetch_account_names, fetch_subcategories, fetch_all_categories, fetch_trigger_option, fetch_action_option)
from ui_utils import (VerticalScrolledFrame, resource_path, open_form_with_position, close_form_with_position, open_calendar, center_window, timed_message)

# Set up logging
logger = logging.getLogger('HA.maint_rule_edit')

def edit_rule_form(rule_id, group_id, conn, cursor, form, scrolled_frame):                  # Win_ID = 24
    # Create the Edit Rule form as a top-level window
    year = 2025             ############################################## fix later
    
    edit_form = tk.Toplevel(form)
    edit_form.title("Edit Rule")
    edit_form.transient(form)
    edit_form.grab_set()
    edit_form.configure(bg="lightgray")
    win_id = 24
    open_form_with_position(edit_form, conn, cursor, win_id, "Edit Rule")
    scaling_factor = edit_form.winfo_fpixels('1i') / 96
    edit_form.geometry(f"{int(930 * scaling_factor)}x{int(980 * scaling_factor)}")

    # Configure grid layout for the form
    edit_form.grid_rowconfigure(0, weight=0)
    edit_form.grid_rowconfigure(1, weight=1)
    edit_form.grid_rowconfigure(2, weight=0)
    edit_form.grid_columnconfigure(0, weight=1)

    # Fixed frame for rule details
    fixed_frame = tk.Frame(edit_form, bg="lightgray")
    fixed_frame.grid(row=0, column=0, sticky="ew", padx=50 * scaling_factor)
    
    # Initialize form attributes for actions
    edit_form.action_combos = {}
    #edit_form.action_rows = []
    edit_form.action_value_widgets = {}
    edit_form.next_action_row = 1

    # Fetch rule details from the database
    cursor.execute("SELECT Rule_Name, Group_ID, Rule_Active, Rule_Trigger_Mode, Rule_Proceed FROM Rules WHERE Rule_ID = ?", (rule_id,))
    rule = cursor.fetchone()
    if not rule:
        messagebox.showerror("Error", f"Rule ID {rule_id} not found", parent=edit_form)
        edit_form.destroy()
        return
    rule_name, group_id, rule_active, rule_trigger_mode, rule_proceed = rule

    # Fetch group name for the rule
    cursor.execute("SELECT Group_Name FROM RuleGroups WHERE Group_ID = ?", (group_id,))
    group_name_result = cursor.fetchone()
    group_name = group_name_result[0] if group_name_result else ""

    # Fetch available options for comboboxes
    group_names = fetch_rule_group_names(cursor)
    trigger_options = fetch_trigger_options(cursor)
    action_options = fetch_action_options(cursor)
    account_names = fetch_account_names(cursor, year)
    parent_categories = fetch_all_categories(cursor, year)
    #logger.debug(f"account_names = {account_names}")
    #logger.debug(f"parent_categories = {parent_categories}")

    # Rule Name field
    tk.Label(fixed_frame, text="Rule Name:", font=("Arial", 11), bg="lightgray").grid(row=0, column=0, sticky="e", padx=20 * scaling_factor, pady=5 * scaling_factor)
    rule_name_entry = tk.Entry(fixed_frame, width=42, font=("Arial", 10))
    rule_name_entry.grid(row=0, column=1, columnspan=2, sticky="w", padx=10 * scaling_factor)
    rule_name_entry.insert(0, rule_name)

    # Rule Group Name combobox
    tk.Label(fixed_frame, text="Rule Group Name:", font=("Arial", 11), bg="lightgray").grid(row=1, column=0, sticky="e", padx=20 * scaling_factor, pady=5 * scaling_factor)
    rule_group_combo = ttk.Combobox(fixed_frame, values=group_names, width=40, state="readonly", font=("Arial", 10))
    rule_group_combo.grid(row=1, column=1, columnspan=2, sticky="w", padx=10 * scaling_factor)
    rule_group_combo.set(group_name)

    # Trigger condition combobox (e.g., when a transaction is created/updated)
    tk.Label(fixed_frame, text="Trigger:", font=("Arial", 11), bg="lightgray").grid(row=2, column=0, sticky="e", padx=20 * scaling_factor, pady=5 * scaling_factor)
    trigger_options_list = ["when a transaction is created", "when a transaction is updated"]
    trigger_combo = ttk.Combobox(fixed_frame, values=trigger_options_list, width=40, state="readonly", font=("Arial", 10))
    trigger_combo.grid(row=2, column=1, columnspan=2, sticky="w", padx=10 * scaling_factor)
    trigger_combo.set("when a transaction is created")

    # Active status combobox
    tk.Label(fixed_frame, text="Active:", font=("Arial", 11), bg="lightgray").grid(row=3, column=0, sticky="e", padx=20 * scaling_factor, pady=5 * scaling_factor)
    active_options = ["Active", "Inactive"]
    active_combo = ttk.Combobox(fixed_frame, values=active_options, width=10, state="readonly", font=("Arial", 10))
    active_combo.grid(row=3, column=1, sticky="w", padx=10 * scaling_factor)
    active_combo.set("Active" if rule_active else "Inactive")
    tk.Label(fixed_frame, text="Inactive rules will never fire.", font=("Arial", 10), bg="lightgray").grid(row=3, column=2, sticky="w", padx=10 * scaling_factor)

    # Stop processing combobox
    tk.Label(fixed_frame, text="Stop processing:", font=("Arial", 11), bg="lightgray").grid(row=4, column=0, sticky="e", padx=20 * scaling_factor, pady=5 * scaling_factor)
    stop_options = ["Stop", "Proceed"]
    stop_combo = ttk.Combobox(fixed_frame, values=stop_options, width=10, state="readonly", font=("Arial", 10))
    stop_combo.grid(row=4, column=1, sticky="w", padx=10 * scaling_factor)
    stop_combo.set("Stop" if not rule_proceed else "Proceed")
    tk.Label(fixed_frame, text="When you select Stop, if triggered, later rules in the group will not be executed.", font=("Arial", 10), bg="lightgray").grid(row=4, column=2, sticky="w", padx=10 * scaling_factor)

    # Strict mode combobox
    tk.Label(fixed_frame, text="Strict mode:", font=("Arial", 11), bg="lightgray").grid(row=5, column=0, rowspan=2, sticky="e", padx=20 * scaling_factor, pady=5 * scaling_factor)
    strict_options = ["ALL", "Any"]
    strict_combo = ttk.Combobox(fixed_frame, values=strict_options, width=10, state="readonly", font=("Arial", 10))
    strict_combo.grid(row=5, column=1, rowspan=2, sticky="w", padx=10 * scaling_factor)
    strict_combo.set("ALL" if rule_trigger_mode == "All" else "Any")
    tk.Label(fixed_frame, text="In strict mode ALL triggers must fire for the action(s) to be executed.", font=("Arial", 10), bg="lightgray").grid(row=5, column=2, sticky="w", padx=10 * scaling_factor)
    tk.Label(fixed_frame, text="In non-strict mode, ANY trigger is enough for the action(s) to be executed.", font=("Arial", 10), bg="lightgray").grid(row=6, column=2, sticky="w", padx=10 * scaling_factor)

    logger.debug("Edit Rule form open with top 6 fields created")
    
    # Scrolled frame for triggers and actions
    edit_scrolled_frame = VerticalScrolledFrame(edit_form)
    edit_scrolled_frame.grid(row=1, column=0, sticky="nsew", pady=10 * scaling_factor)
    edit_scrolled_frame.canvas.focus_set()

    # Fetch existing triggers for the rule
    cursor.execute("SELECT Trigger_ID, TrigO_ID, Value, Trigger_Sequence FROM Triggers WHERE Rule_ID = ? ORDER BY Trigger_Sequence", (rule_id,))
    triggers = cursor.fetchall()
    #logger.debug(f"Rule_ID={rule_id}, Triggers={triggers}")

    # Fetch existing actions for the rule
    cursor.execute("SELECT Action_ID, ActO_ID, Value, Action_Sequence FROM Actions WHERE Rule_ID = ? ORDER BY Action_Sequence", (rule_id,))
    actions = cursor.fetchall()
    #logger.debug(f"Rule_ID={rule_id}, Actions={actions}")
        
    # Load delete icon for trigger/action rows
    delete_img = tk.PhotoImage(file=resource_path("icons/3_trash-48.png"))
    #image_refs = [delete_img]

    # Validation functions
    def validate_6d2_number(text):
        if not text:
            return True
        if len(text) > 9:  # Max length: 6 digits + '.' + 2 digits
            return False
        return bool(re.match(r'^\d{0,6}(\.\d{0,2})?$', text))

    def validate_date(text):
        if not text:
            return False
        if len(text) > 10:  # Max length: DD/MM/YYYY
            return False
        return bool(re.match(r'^(\d{0,2}(/\d{0,2}(/\d{0,4})?)?)?$', text))

    vcmd_6d2 = (edit_form.register(validate_6d2_number), '%P')
    vcmd_date = (edit_form.register(validate_date), '%P')

    # Triggers frame
    triggers_frame = ttk.LabelFrame(edit_scrolled_frame.interior, text="Rule triggers when", width=960 * scaling_factor)
    triggers_frame.grid(row=0, column=0, sticky="ew", padx=20 * scaling_factor, pady=10 * scaling_factor)
    triggers_frame.grid_columnconfigure(0, minsize=20 * scaling_factor)
    triggers_frame.grid_columnconfigure(1, weight=0)
    triggers_frame.grid_columnconfigure(2, weight=0)
    triggers_frame.grid_columnconfigure(3, weight=1)  # Extra column for calendar button

    tk.Label(triggers_frame, text="Delete", font=("Arial", 10, "bold")).grid(row=0, column=0, sticky="w", padx=10 * scaling_factor, pady=5 * scaling_factor)
    tk.Label(triggers_frame, text="Trigger", font=("Arial", 10, "bold")).grid(row=0, column=1, sticky="w", padx=10 * scaling_factor, pady=5 * scaling_factor)
    tk.Label(triggers_frame, text="Trigger on value", font=("Arial", 10, "bold")).grid(row=0, column=2, sticky="w", padx=10 * scaling_factor, pady=5 * scaling_factor)
    #tk.Label(triggers_frame, text="", font=("Arial", 10)).grid(row=0, column=3, sticky="w", padx=10 * scaling_factor, pady=5 * scaling_factor)  ### added
    
    logger.debug("Trigger frame and headings created")

    trigger_rows = []
    for i, trigger in enumerate(triggers, start=1):
        trigger_id, trigo_id, trigger_value, trigger_sequence = trigger
        trig_option = fetch_trigger_option(cursor, trigo_id)
        # Handle cases where trig_option is a tuple
        #logger.debug(f"Trigger ID {trigger_id}: TrigO_ID={trigo_id}, trig_option={trig_option}, trigger_value={trigger_value}, trigger_sequence={trigger_sequence}")
        trigger_desc = trig_option[0] if trig_option and len(trig_option) > 0 else f"Unknown Trigger (ID: {trigo_id})"
        if not trig_option:
            logger.warning(f"No trigger option found for TrigO_ID {trigo_id}")

        # Create delete button for the trigger row
        delete_btn = tk.Button(triggers_frame, image=delete_img, bg="red", command=lambda tid=trigger_id: delete_trigger_row(tid, conn, cursor, triggers_frame, trigger_rows))
        delete_btn.grid(row=i, column=0, sticky="w", padx=20 * scaling_factor)

        # Create combobox for trigger condition
        combo = ttk.Combobox(triggers_frame, values=[desc for _, desc in trigger_options], width=40, state="readonly", font=("Arial", 10))
        combo.set(trigger_desc)
        combo.grid(row=i, column=1, sticky="w", padx=10 * scaling_factor)
        combo.bind("<MouseWheel>", lambda e: scroll_frame(edit_scrolled_frame, e))  # Redirect mouse wheel scrolling to frame

        # Create dynamic value field based on TrigO_ID
        value_widget = None
        extra_widget = None
        cspan = 2
        if 1 <= trigo_id <= 3:                              # Amount - Numeric field (6d2 format)
            value_widget = tk.Entry(triggers_frame, width=10, font=("Arial", 10), validate="key", validatecommand=vcmd_6d2)
            value_widget.insert(0, trigger_value or "")
            #logger.debug(f"Numeric 6d2 edit created for trigo_id= {trigo_id}")
        elif 4 <= trigo_id <= 15:                           # Tag or Description -Text field (40 chars)
            value_widget = tk.Entry(triggers_frame, width=58, font=("Arial", 10))
            value_widget.insert(0, trigger_value or "")
            if trigger_value and len(trigger_value) > 40:
                value_widget.delete(40, tk.END)
            #logger.debug(f"Edit 40 created for trigo_id= {trigo_id}")
        elif 16 <= trigo_id <= 19:                          # Account Name - Combobox with account names
            accid = int(trigger_value)
            value_widget = ttk.Combobox(triggers_frame, values=[name for _, name in account_names], width=25, state="readonly", font=("Arial", 10))
            value_widget.set(next((name for a, name in account_names if a == accid), ""))
            value_widget.bind("<MouseWheel>", lambda e: scroll_frame(edit_scrolled_frame, e))  # Redirect mouse wheel scrolling to frame
            #logger.debug(f"Combobox created for trigo_id= {trigo_id}, trigger_value={trigger_value}")
        elif 20 <= trigo_id <= 22:                          # Date - Date field with calendar button
            value_widget = tk.Entry(triggers_frame, width=10, font=("Arial", 10), validate="key", validatecommand=vcmd_date)
            value_widget.insert(0, trigger_value or "")
            extra_widget = tk.Button(triggers_frame, text="Cal", font=("Arial", 8), command=lambda ve=value_widget: open_calendar(edit_form, ve))
            extra_widget.grid(row=i, column=3, sticky="w", padx=5 * scaling_factor)
            cspan = 1
            #logger.debug(f"Date/Cal created for trigo_id= {trigo_id}")
        elif 23 <= trigo_id <= 27:                          # No field required
            value_widget = tk.Label(triggers_frame, state="disabled")
            #value_widget.insert(0, "")
            #logger.debug(f"Disabled Edit created for trigo_id= {trigo_id}")

        if value_widget:
            value_widget.grid(row=i, column=2, columnspan=cspan, sticky="w", padx=10 * scaling_factor)

        trigger_rows.append((trigger_id, delete_btn, combo, value_widget, trigger_sequence, extra_widget, trigo_id))  # Add trigo_id for reference

        def make_trigger_handler(row=i, combo=combo, value_widget=value_widget, extra_widget=extra_widget):
            return lambda e: update_trigger_value_field(e, row, combo, value_widget, extra_widget)
        combo.bind("<<ComboboxSelected>>", make_trigger_handler())

        #logger.debug(f"Trigger row {i} created - ")

    def update_trigger_value_field(event, row, combo, value_widget, extra_widget):
        
        #logger.debug(f"update_trigger_value_field - row={row}, combo={combo}")
        
        selected_desc = combo.get()
        cursor.execute("SELECT TrigO_ID FROM Trig_Options WHERE TrigO_Description = ?", (selected_desc,))
        trigo_id_result = cursor.fetchone()
        if not trigo_id_result:
            return
        trigo_id = trigo_id_result[0]

        if value_widget:
            value_widget.destroy()
        if extra_widget:
            extra_widget.destroy()
            extra_widget = None
        cspan = 2

        if 1 <= trigo_id <= 3:                              # Amount - Numeric field (6d2 format)
            value_widget = tk.Entry(triggers_frame, width=10, font=("Arial", 10), validate="key", validatecommand=vcmd_6d2)
        elif 4 <= trigo_id <= 15:                           # Tag or Description -Text field (40 chars)
            value_widget = tk.Entry(triggers_frame, width=58, font=("Arial", 10))
        elif 16 <= trigo_id <= 19:                          # Account Name - Combobox with account names
            value_widget = ttk.Combobox(triggers_frame, values=[name for _, name in account_names], width=25, state="readonly", font=("Arial", 10))
            value_widget.bind("<MouseWheel>", lambda e: scroll_frame(edit_scrolled_frame, e))  # Redirect mouse wheel scrolling to frame
        elif 20 <= trigo_id <= 22:                          # Date - Date field with calendar button
            value_widget = tk.Entry(triggers_frame, width=10, font=("Arial", 10), validate="key", validatecommand=vcmd_date)
            extra_widget = tk.Button(triggers_frame, text="Cal", font=("Arial", 8), command=lambda ve=value_widget: open_calendar(edit_form, ve))
            extra_widget.grid(row=row, column=3, sticky="w", padx=5 * scaling_factor)
            cspan = 1
        elif 23 <= trigo_id <= 27:                          # No field required
            value_widget = tk.Label(triggers_frame, state="disabled")
        value_widget.grid(row=row, column=2, columnspan=cspan, sticky="w", padx=10 * scaling_factor)
        
        #logger.debug(f"value_widget.grid - row={row}")
        
        trigger_rows[row - 1] = (trigger_rows[row - 1][0], trigger_rows[row - 1][1], combo, value_widget, trigger_rows[row - 1][4], extra_widget, trigo_id)

        edit_form.update_idletasks()
            
    def scroll_frame(form, event):
        # Forward mouse wheel event to the triggers_frame's parent canvas
        if edit_scrolled_frame.canvas.winfo_exists():
            if event.delta > 0:
                form.canvas.yview_scroll(-1, "units")
                #logger.debug("Mouse wheel up")
            elif event.delta < 0:
                form.canvas.yview_scroll(1, "units")
                #logger.debug("Mouse wheel down")
        #logger.debug(f"Mouse wheel event: delta={event.delta}, num={getattr(event, 'num', None)}")
        return "break"
    
    def add_trigger_row():
        # Add a new trigger row to the UI
        nonlocal trigger_rows
        row = len(trigger_rows) + 1
        delete_btn = tk.Button(triggers_frame, image=delete_img, bg="red", command=lambda: delete_new_trigger_row(row, triggers_frame, trigger_rows))
        delete_btn.grid(row=row, column=0, sticky="w", padx=20 * scaling_factor)

        combo = ttk.Combobox(triggers_frame, values=[desc for _, desc in trigger_options], width=40, state="readonly", font=("Arial", 10))
        combo.grid(row=row, column=1, sticky="w", padx=10 * scaling_factor)
        combo.bind("<MouseWheel>", lambda e: scroll_frame(edit_scrolled_frame, e))  # Redirect mouse wheel scrolling to frame

        # Initialize with a default value field (updated on combobox selection)
        value_widget = tk.Entry(triggers_frame, width=58, font=("Arial", 10))
        value_widget.grid(row=row, column=2, columnspan=2, sticky="w", padx=10 * scaling_factor)
        extra_widget = None

        def make_trigger_handler(row=row, combo=combo, value_widget=value_widget, extra_widget=extra_widget):
            return lambda e: update_trigger_value_field(e, row, combo, value_widget, extra_widget)
        combo.bind("<<ComboboxSelected>>", make_trigger_handler())

        trigger_rows.append((None, delete_btn, combo, value_widget, row, extra_widget, None))
        add_trigger_btn.grid(row=row + 1, column=0, columnspan=4)
        edit_scrolled_frame.canvas.configure(scrollregion=edit_scrolled_frame.canvas.bbox("all"))
        edit_form.update_idletasks()
        #logger.debug(f"Added new trigger row at row {row}")

    def delete_trigger_row(trigger_id, conn, cursor, frame, rows):
        # Delete an existing trigger from the database and UI
        try:
            cursor.execute("DELETE FROM Triggers WHERE Trigger_ID = ?", (trigger_id,))
            conn.commit()
            for i, (tid, dbtn, cb, vw, seq, ew, tid_id) in enumerate(rows):
                if tid == trigger_id:
                    dbtn.destroy()
                    cb.destroy()
                    if vw:
                        vw.destroy()
                    if ew:
                        ew.destroy()
                    rows.pop(i)
                    break
            for j, (tid, dbtn, cb, vw, seq, ew, tid_id) in enumerate(rows):
                dbtn.grid(row=j + 1, column=0)
                cb.grid(row=j + 1, column=1)
                if vw:
                    vw.grid(row=j + 1, column=2)
                if ew:
                    ew.grid(row=j + 1, column=3)
                rows[j] = (tid, dbtn, cb, vw, j + 1, ew, tid_id)
            add_trigger_btn.grid(row=len(rows) + 1, column=0, columnspan=4)
            edit_scrolled_frame.canvas.configure(scrollregion=edit_scrolled_frame.canvas.bbox("all"))
            edit_form.update_idletasks()
            #logger.debug(f"Deleted trigger ID {trigger_id}")
        except Exception as e:
            logger.error(f"Error deleting trigger ID {trigger_id}: {e}\n{traceback.format_exc()}")
            messagebox.showerror("Error", f"Failed to delete trigger: {e}", parent=edit_form)

    def delete_new_trigger_row(row, frame, rows):
        # Delete a new (unsaved) trigger row from the UI
        for i, (tid, dbtn, cb, vw, seq, ew, tid_id) in enumerate(rows):
            if seq == row:
                dbtn.destroy()
                cb.destroy()
                if vw:
                    vw.destroy()
                if ew:
                    ew.destroy()
                rows.pop(i)
                break
        for j, (tid, dbtn, cb, vw, seq, ew, tid_id) in enumerate(rows):
            dbtn.grid(row=j + 1, column=0)
            cb.grid(row=j + 1, column=1)
            if vw:
                vw.grid(row=j + 1, column=2)
            if ew:
                ew.grid(row=j + 1, column=3)
            rows[j] = (tid, dbtn, cb, vw, j + 1, ew, tid_id)
        add_trigger_btn.grid(row=len(rows) + 1, column=0, columnspan=4)
        edit_scrolled_frame.canvas.configure(scrollregion=edit_scrolled_frame.canvas.bbox("all"))
        edit_form.update_idletasks()
        #logger.debug(f"Deleted new trigger row at row {row}")

    # Create button to add new trigger
    add_trigger_btn = tk.Button(triggers_frame, text="Add new trigger", font=("Arial", 10), width=20, command=add_trigger_row)
    add_trigger_btn.grid(row=len(triggers)+1, column=0, columnspan=4, sticky="w", padx=10 * scaling_factor, pady=10 * scaling_factor)

    ##### Actions frame starts here ###################################
    actions_frame = ttk.LabelFrame(edit_scrolled_frame.interior, text="Rule will", width=960 * scaling_factor)
    actions_frame.grid(row=1, column=0, sticky="ew", padx=20 * scaling_factor, pady=10 * scaling_factor)
    actions_frame.grid_columnconfigure(0, minsize=20 * scaling_factor)
    actions_frame.grid_columnconfigure(1, weight=0)
    actions_frame.grid_columnconfigure(2, weight=0)
    actions_frame.grid_columnconfigure(3, weight=1)

    tk.Label(actions_frame, text="Delete", font=("Arial", 10, "bold")).grid(row=0, column=0, sticky="w", padx=10 * scaling_factor, pady=5 * scaling_factor)
    tk.Label(actions_frame, text="Action", font=("Arial", 10, "bold")).grid(row=0, column=1, sticky="w", padx=10 * scaling_factor, pady=5 * scaling_factor)
    tk.Label(actions_frame, text="Action Value", font=("Arial", 10, "bold")).grid(row=0, column=2, sticky="w", padx=10 * scaling_factor, pady=5 * scaling_factor)

    logger.debug("  ### Action frame and headings created")
    
    action_rows = []
    for i, action in enumerate(actions, start=1):
        action_id, acto_id, action_value, action_sequence = action
        act_option = fetch_action_option(cursor, acto_id)
        # Handle cases where act_option is a tuple
        #logger.debug(f"Action ID {action_id}: ActO_ID={acto_id}, act_option={act_option}")
        action_desc = act_option[0] if act_option and len(act_option) > 0 else f"Unknown Action (ID: {acto_id})"
        if not act_option:
            logger.warning(f"No action option found for ActO_ID {acto_id}")

        # Create delete button for the action row
        delete_btn = tk.Button(actions_frame, image=delete_img, bg="red", command=lambda aid=action_id: delete_action_row(aid, conn, cursor, actions_frame, action_rows))
        delete_btn.grid(row=i, column=0, sticky="w", padx=20 * scaling_factor)

        # Create combobox for action condition
        combo = ttk.Combobox(actions_frame, values=[desc for _, desc in action_options], width=40, state="readonly", font=("Arial", 10))
        combo.set(action_desc)
        combo.grid(row=i, column=1, sticky="w", padx=10 * scaling_factor)
        combo.bind("<MouseWheel>", lambda e: scroll_frame(edit_scrolled_frame, e))  # Redirect mouse wheel scrolling to frame
        
        # Create dynamic value field based on ActO_ID
        value_widget = None
        extra_widget = None
        cspan = 2
        if acto_id in (1, 2, 10, 11, 12):                       # Text field (40 chars)
            value_widget = tk.Entry(actions_frame, width=58, font=("Arial", 10))
            value_widget.insert(0, action_value or "")
            if action_value and len(action_value) > 40:
                value_widget.delete(40, tk.END)
            #logger.debug(f"Edit 40 created for acto_id= {acto_id}")
        elif acto_id in (3, 4, 5, 6, 7):                        # No field required
            value_widget = tk.Label(actions_frame, state="disabled")
            #value_widget.insert(0, "")
            #logger.debug(f"Disabled Edit created for acto_id= {acto_id}")
        elif acto_id == 8:                                      # Two comboboxes for parent and sub-category
            try:
                pid, cid = map(int, action_value.split(',')) if action_value and ',' in action_value else (0, 0)
            except ValueError:
                pid, cid = 0, 0
            value_widget = ttk.Combobox(actions_frame, values=[name for _, name in parent_categories], width=25, state="readonly", font=("Arial", 10))
            value_widget.set(next((name for p, name in parent_categories if p == pid), ""))
            value_widget.bind("<MouseWheel>", lambda e: scroll_frame(edit_scrolled_frame, e))  # Redirect mouse wheel scrolling to frame
            value_widget.grid(row=i, column=2, sticky="w", padx=10 * scaling_factor)
            sub_categories = fetch_subcategories(cursor, pid, year) if pid else []
            extra_widget = ttk.Combobox(actions_frame, values=[name for _, name in sub_categories], width=25, state="readonly", font=("Arial", 10))
            extra_widget.set(next((name for c, name in sub_categories if c == cid), ""))
            extra_widget.bind("<MouseWheel>", lambda e: scroll_frame(edit_scrolled_frame, e))  # Redirect mouse wheel scrolling to frame
            extra_widget.grid(row=i, column=3, sticky="w", padx=10 * scaling_factor)
            cspan = 1
            #logger.debug(f"Dual Comboboxes created for acto_id= {acto_id}")
        elif acto_id in (13, 14):                               # Combobox with account names
            accid = int(action_value)
            value_widget = ttk.Combobox(actions_frame, values=[name for _, name in account_names], width=25, state="readonly", font=("Arial", 10))
            value_widget.set(next((name for a, name in account_names if a == accid), ""))
            value_widget.bind("<MouseWheel>", lambda e: scroll_frame(edit_scrolled_frame, e))  # Redirect mouse wheel scrolling to frame
            #logger.debug(f"Accounts Combobox created for acto_id = {acto_id}, accid = {accid}")
        if value_widget and acto_id != 8:
            value_widget.grid(row=i, column=2, columnspan=cspan, sticky="w", padx=10 * scaling_factor)

        action_rows.append((action_id, delete_btn, combo, value_widget, action_sequence, extra_widget, acto_id))

        def make_action_handler(row=i, combo=combo, value_widget=value_widget, extra_widget=extra_widget):
            return lambda e: update_action_value_field(edit_form, e)
        combo.bind("<<ComboboxSelected>>", make_action_handler())
        if value_widget and acto_id == 8:
            value_widget.bind("<<ComboboxSelected>>", make_action_handler())

        #logger.debug(f"Action row {i} created")
        
        edit_form.next_action_row = i+1
        edit_form.action_combos[i] = combo
        edit_form.action_value_widgets[i] = value_widget

    def update_action_value_field(edit_form, event):
        combo = event.widget
        row = combo.grid_info()['row']
        if row in edit_form.action_value_widgets:
            edit_form.action_value_widgets[row].destroy()
        
        selected_desc = combo.get()
        cursor.execute("SELECT ActO_ID FROM Act_Options WHERE ActO_Description = ?", (selected_desc,))
        acto_id_result = cursor.fetchone()
        if not acto_id_result:
            return
        acto_id = acto_id_result[0]
        value_widget, extra_widget = create_action_value_widget(edit_form, acto_id, "")
        
        if acto_id == 8:
            value_widget.grid(row=row, column=2, sticky="w", padx=10 * scaling_factor)
            extra_widget.grid(row=row, column=3, sticky="w", padx=10 * scaling_factor)
        else:
            value_widget.grid(row=row, column=2, columnspan=2, sticky="w", padx=10 * scaling_factor)
            
        action_rows[row - 1] = (action_rows[row - 1][0], action_rows[row - 1][1], combo, value_widget, action_rows[row - 1][4], extra_widget, acto_id)
        edit_form.update_idletasks()

    def create_action_value_widget(edit_form, acto_id, preload):
        extra_widget = None
        
        if acto_id in (1, 2, 10, 11, 12):       # Tag text or Description text
            value_widget = tk.Entry(actions_frame, width=58, font=("Arial", 10))
        elif acto_id in (3, 4, 5, 6, 7):        # none required
            value_widget = tk.Label(actions_frame, state="disabled")
        elif acto_id == 8:                      # pid, cid values - numeric pair
            value_widget = ttk.Combobox(actions_frame, values=[name for _, name in parent_categories], width=25, state="readonly", font=("Arial", 10))
            value_widget.bind("<MouseWheel>", lambda e: scroll_frame(edit_scrolled_frame, e))  # Redirect mouse wheel scrolling to frame
            extra_widget = ttk.Combobox(actions_frame, values=[], width=25, state="readonly", font=("Arial", 10))
            extra_widget.bind("<MouseWheel>", lambda e: scroll_frame(edit_scrolled_frame, e))  # Redirect mouse wheel scrolling to frame

            def update_sub_categories(event):
                parent_name = value_widget.get()
                cursor.execute("SELECT IE_PID FROM IE_Cata WHERE IE_Desc = ? AND IE_CID = 0 AND IE_Year = ?", (parent_name, year))
                pid_result = cursor.fetchone()
                pid = pid_result[0] if pid_result else 0
                #logger.debug(f"from combo 1 select - pid = {pid} ")
                sub_categories = fetch_subcategories(cursor, pid, year) if pid else []
                #logger.debug(f"subcategories = {sub_categories} ")
                extra_widget['values'] = [name for _, name in sub_categories]
                extra_widget.set("")
                edit_form.update_idletasks()

            value_widget.bind("<<ComboboxSelected>>", update_sub_categories)
        elif acto_id in (13, 14):               # acc_id value - numeric 1-12
            value_widget = ttk.Combobox(actions_frame, values=[name for _, name in account_names], width=25, state="readonly", font=("Arial", 10))
            value_widget.bind("<MouseWheel>", lambda e: scroll_frame(edit_scrolled_frame, e))  # Redirect mouse wheel scrolling to frame

        return value_widget, extra_widget

    def get_available_action_options(cursor, row, selected_actions):
        # Fetch all action options from Act_Options table
        #action_options = fetch_action_options(cursor)
        cursor.execute("SELECT ActO_ID, ActO_Description FROM Act_Options ORDER BY ActO_Seq")
        action_options = cursor.fetchall()
        
        #logger.debug(f"selected_actions={selected_actions}")
        #logger.debug(f"action_options={action_options}")
        
        # Define restriction sets
        unique_actions = {3, 8, 10, 13, 14}  # Can appear only once
        group_actions = {4, 5, 6}  # Mutually exclusive group
        delete_action = {7}  # Only in first row
        
        # Start with all options
        available_options = []
        for acto_id, acto_desc in action_options:
            # Skip ActO_ID 7 if not the first row
            if acto_id in delete_action and row != 1:
                continue
            # Skip unique actions already selected elsewhere
            if acto_id in unique_actions and any(acto_id == sel_id for sel_row, sel_desc, sel_id in selected_actions if sel_row != row):
                continue
            # Skip group actions if another in the group is selected elsewhere
            if acto_id in group_actions:
                selected_group_ids = {sel_id for sel_row, sel_desc, sel_id in selected_actions if sel_id in group_actions and sel_row != row}
                #logger.debug(f"selected_group_ids={selected_group_ids}")
                if selected_group_ids:
                    continue
            available_options.append((acto_id, acto_desc))
        
        #logger.debug(f"available_options={available_options}")
        return available_options
    
    def add_action_row(edit_form, cursor, preload=None):
        # Add a new action row to the UI
        nonlocal action_rows
        row = len(action_rows) + 1
        
        #logger.debug(f"#######  add_action_row called - preload={preload}")
        
        # Set default or preloaded value
        #if preload:
        #    act_option = fetch_action_option(cursor, preload['acto_id'])
        #    combo.set(act_option if act_option in [name for _, name in available_options] else "")
        #else:
        #   combo.set("")
        
        #edit_form.next_action_row = getattr(edit_form, 'next_action_row', 1)
        #row = edit_form.next_action_row
        #edit_form.action_combos = getattr(edit_form, 'action_combos', {})
        #edit_form.action_rows = getattr(edit_form, 'action_rows', [])
        #edit_form.action_value_widgets = getattr(edit_form, 'action_value_widgets', {}) 
        
        #row = edit_form.next_action_row 
        
        delete_btn = tk.Button(actions_frame, image=delete_img, bg="red", command=lambda: delete_new_action_row(row, actions_frame, action_rows))
        delete_btn.grid(row=row, column=0, sticky="w", padx=20 * scaling_factor)

        combo = ttk.Combobox(actions_frame, width=40, state="readonly", font=("Arial", 10))
        combo.grid(row=row, column=1, sticky="w", padx=10 * scaling_factor)
        
        # Get selected actions from other rows
        selected_actions = [(r, edit_form.action_combos[r].get(), next((action[6] for action in action_rows if action[2] is edit_form.action_combos[r]), None))
                            for r in edit_form.action_combos if edit_form.action_combos[r].get()]
        #logger.debug(f"selected_actions = {selected_actions}")
        
        # Set combobox values based on restrictions
        available_options = get_available_action_options(cursor, row, selected_actions)
        combo['values'] = [name for _, name in available_options]
        
        # Set default or preloaded value
        #if preload:
        #    act_option = fetch_action_option(cursor, preload['acto_id'])
        #    combo.set(act_option if act_option in [name for _, name in available_options] else "")
        #else:
        #    combo.set("")
            
        combo.bind("<MouseWheel>", lambda e: scroll_frame(edit_scrolled_frame, e))  # Redirect mouse wheel scrolling to frame

        # Initialize with a default value field (updated on combobox selection)
        value_widget = tk.Entry(actions_frame, width=58, state="readonly", font=("Arial", 10))
        value_widget.grid(row=row, column=2, columnspan=2, sticky="w", padx=10 * scaling_factor)
        extra_widget = None
        edit_form.action_value_widgets[row] = value_widget
        
        def make_action_handler(row=row, combo=combo, value_widget=value_widget, extra_widget=extra_widget):
            return lambda e: update_action_value_field(edit_form, e)
        combo.bind("<<ComboboxSelected>>", make_action_handler())
        
        edit_form.action_combos[row] = combo
        action_rows.append((None, delete_btn, combo, value_widget, row, extra_widget, None))
        add_action_btn.grid(row=row + 1, column=0, columnspan=4)  # Adjust columnspan for extra column
        #edit_form.next_action_row += 1
        edit_scrolled_frame.canvas.configure(scrollregion=edit_scrolled_frame.canvas.bbox("all"))
        edit_form.update_idletasks()
        #logger.debug(f"Added new action row at row {row}")

    def delete_action_row(action_id, conn, cursor, frame, rows):
        # Delete an existing action from the database and UI
        try:
            cursor.execute("DELETE FROM Actions WHERE Action_ID = ?", (action_id,))
            conn.commit()
            for i, (aid, dbtn, cb, vw, seq, ew, aid_id) in enumerate(rows):
                if aid == action_id:
                    dbtn.destroy()
                    cb.destroy()
                    if vw:
                        vw.destroy()
                    if ew:
                        ew.destroy()
                    rows.pop(i)
                    break
            for j, (aid, dbtn, cb, vw, seq, ew, aid_id) in enumerate(rows):
                dbtn.grid(row=j + 1, column=0)
                cb.grid(row=j + 1, column=1)
                if vw:
                    vw.grid(row=j + 1, column=2)
                if ew:
                    ew.grid(row=j + 1, column=3)
                rows[j] = (aid, dbtn, cb, vw, j + 1, ew, aid_id)
            add_action_btn.grid(row=len(rows) + 1, column=0, columnspan=4)  # Adjust columnspan for extra column
            edit_scrolled_frame.canvas.configure(scrollregion=edit_scrolled_frame.canvas.bbox("all"))
            edit_form.update_idletasks()
            #logger.debug(f"Deleted action ID {action_id}")
        except Exception as e:
            logger.error(f"Error deleting action ID {action_id}: {e}\n{traceback.format_exc()}")
            messagebox.showerror("Error", f"Failed to delete action: {e}", parent=edit_form)

    def delete_new_action_row(row, frame, rows):
        # Delete a new (unsaved) action row from the UI
        for i, (aid, dbtn, cb, vw, seq, ew, aid_id) in enumerate(rows):
            if seq == row:
                dbtn.destroy()
                cb.destroy()
                if vw:
                    vw.destroy()
                if ew:
                    ew.destroy()
                rows.pop(i)
                break
        for j, (aid, dbtn, cb, vw, seq, ew, aid_id) in enumerate(rows):
            dbtn.grid(row=j + 1, column=0)
            cb.grid(row=j + 1, column=1)
            if vw:
                vw.grid(row=j + 1, column=2)
            if ew:
                ew.grid(row=j + 1, column=3)
            rows[j] = (aid, dbtn, cb, vw, j + 1, ew, aid_id)
        add_action_btn.grid(row=len(rows) + 1, column=0, columnspan=4)  # Adjust columnspan for extra column
        edit_scrolled_frame.canvas.configure(scrollregion=edit_scrolled_frame.canvas.bbox("all"))
        edit_form.update_idletasks()
        #logger.debug(f"Deleted new action row at row {row}")

    # Create button to add new action
    add_action_btn = tk.Button(actions_frame, text="Add new action", font=("Arial", 10), width=20, command=lambda:add_action_row(edit_form, cursor, ""))
    add_action_btn.grid(row=len(actions)+1, column=0, columnspan=4, sticky="w", padx=10 * scaling_factor, pady=10 * scaling_factor)

    def save_rule_changes():
        try:
            updated_rule_name = rule_name_entry.get().strip()
            updated_group_name = rule_group_combo.get()
            updated_trigger = trigger_combo.get()
            updated_active = 1 if active_combo.get() == "Active" else 0
            updated_stop = 0 if stop_combo.get() == "Proceed" else 1
            updated_strict = "All" if strict_combo.get() == "ALL" else "Any"

            # Validate rule name
            if not updated_rule_name:
                messagebox.showerror("Error", "Rule Name cannot be empty", parent=edit_form)
                return

            # Validate group name
            cursor.execute("SELECT Group_ID FROM RuleGroups WHERE Group_Name = ?", (updated_group_name,))
            new_group_id_result = cursor.fetchone()
            if not new_group_id_result:
                messagebox.showerror("Error", "Selected Rule Group not found", parent=edit_form)
                return
            new_group_id = new_group_id_result[0]

            # Update rule details
            cursor.execute("UPDATE Rules SET Rule_Name = ?, Group_ID = ?, Rule_Active = ?, Rule_Trigger_Mode = ?, Rule_Proceed = ? WHERE Rule_ID = ?", 
                        (updated_rule_name, new_group_id, updated_active, updated_strict, not updated_stop, rule_id))

            # Handle triggers: delete removed ones, update existing, insert new
            #logger.debug(f"trigger_rows={trigger_rows}")
            existing_trigger_ids = set(tid for tid, _, _, _, _, _, _ in trigger_rows if tid is not None)
            cursor.execute("SELECT Trigger_ID FROM Triggers WHERE Rule_ID = ?", (rule_id,))
            db_trigger_ids = set(tid[0] for tid in cursor.fetchall())
            triggers_to_delete = db_trigger_ids - existing_trigger_ids
            for tid in triggers_to_delete:
                cursor.execute("DELETE FROM Triggers WHERE Trigger_ID = ?", (tid,))
            
            for i, (tid, _, combo, value_widget, seq, extra_widget, _) in enumerate(trigger_rows):
                trig_desc = combo.get()
                if trig_desc == "":                 # skip any blank triggers
                    continue
                cursor.execute("SELECT TrigO_ID FROM Trig_Options WHERE TrigO_Description = ?", (trig_desc,))
                trigo_id_result = cursor.fetchone()
                if not trigo_id_result:
                    messagebox.showerror("Error", f"Invalid trigger condition: {trig_desc}", parent=edit_form)
                    return
                trigo_id = trigo_id_result[0]
                trig_value = ""
                
                if trigo_id in (1, 2, 3):                                       # Amount
                    trig_value = value_widget.get().strip()
                    #logger.debug(f"1 - Saved changes for tid={tid}, trig_value={trig_value}")
                elif trigo_id in (4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15):    # Tag text or Description text
                    trig_value = value_widget.get().strip()
                    #logger.debug(f"2 - Saved changes for tid={tid}, trig_value={trig_value}")
                elif trigo_id in (16, 17, 18, 19):                              # Account Name
                    account_name = value_widget.get()
                    cursor.execute("SELECT Acc_ID FROM Account WHERE Acc_Name = ? AND Acc_Year = ?", (account_name, year))
                    acc_result = cursor.fetchone()
                    trig_value = acc_result[0] if acc_result else 0
                    #logger.debug(f"3 - Saved changes for tid={tid}, trig_value={trig_value}")
                elif trigo_id in (20, 21, 22):                                  # Date
                    trig_value = value_widget.get().strip()
                    #logger.debug(f"4 - Saved changes for tid={tid}, trig_value={trig_value}")
                elif trigo_id in (23, 24, 25, 26, 27):                          # None - (Transaction Type or Status)
                    trig_value = "none"
                    #logger.debug(f"5 - Saved changes for tid={tid}, trig_value={trig_value}")
                if tid is None:
                    cursor.execute("INSERT INTO Triggers (Rule_ID, TrigO_ID, Value, Trigger_Sequence) VALUES (?, ?, ?, ?)", 
                                (rule_id, trigo_id, trig_value, i + 1))
                else:
                    cursor.execute("UPDATE Triggers SET TrigO_ID = ?, Value = ?, Trigger_Sequence = ? WHERE Trigger_ID = ?", 
                                (trigo_id, trig_value, i + 1, tid))

            # Handle actions: delete removed ones, update existing, insert new
            #logger.debug(f"action_rows={action_rows}")
            existing_action_ids = set(aid for aid, _, _, _, _, _, _ in action_rows if aid is not None)
            cursor.execute("SELECT Action_ID FROM Actions WHERE Rule_ID = ?", (rule_id,))
            db_action_ids = set(aid[0] for aid in cursor.fetchall())
            actions_to_delete = db_action_ids - existing_action_ids
            for aid in actions_to_delete:
                cursor.execute("DELETE FROM Actions WHERE Action_ID = ?", (aid,))
            
            for i, (aid, _, combo, value_widget, seq, extra_widget, _) in enumerate(action_rows):
                act_desc = combo.get()
                if act_desc == "":                 # skip any blank actions
                    continue
                cursor.execute("SELECT ActO_ID FROM Act_Options WHERE ActO_Description = ?", (act_desc,))
                acto_id_result = cursor.fetchone()
                if not acto_id_result:
                    messagebox.showerror("Error", f"Invalid action: {act_desc}", parent=edit_form)
                    return
                acto_id = acto_id_result[0]
                #logger.debug(f"Saving changes - acto_id={acto_id}")
                act_value = ""
                if acto_id in (1, 2, 10, 11, 12):       # Tag text or Description text
                    act_value = value_widget.get()
                    #logger.debug(f"1 - Saved changes for aid={aid}, act_value={act_value}")
                elif acto_id in (3, 4, 5, 6, 7):        # none required
                    act_value = "none"
                    #logger.debug(f"2 - Saved changes for aid={aid}, act_value={act_value}")
                elif acto_id == 8:                      # pid, cid values - numeric pair
                    parent_name = value_widget.get()
                    sub_name = extra_widget.get() if extra_widget else ""
                    cursor.execute("SELECT IE_PID FROM IE_Cata WHERE IE_Desc = ? AND IE_CID = 0 AND IE_Year = ?", (parent_name, year))
                    pid_result = cursor.fetchone()
                    pid = pid_result[0] if pid_result else 0
                    #logger.debug(f"3 - Saved changes for aid={aid}, pid={pid}")
                    cursor.execute("SELECT IE_CID FROM IE_Cata WHERE IE_Desc = ? AND IE_PID = ? AND IE_Year = ?", (sub_name, pid, year))
                    cid_result = cursor.fetchone()
                    cid = cid_result[0] if cid_result else 0
                    act_value = f"{pid},{cid}"
                    #logger.debug(f"4 - Saved changes for aid={aid}, act_value={act_value}")
                elif acto_id in (13, 14):               # acc_id value - numeric 1-12
                    account_name = value_widget.get()
                    cursor.execute("SELECT Acc_ID FROM Account WHERE Acc_Name = ? AND Acc_Year = ?", (account_name, year))
                    acc_result = cursor.fetchone()
                    act_value = acc_result[0] if acc_result else 0
                    #logger.debug(f"5 - Saved changes for aid={aid}, act_value={act_value}")
                if aid is None:
                    cursor.execute("INSERT INTO Actions (Rule_ID, ActO_ID, Value, Action_Sequence) VALUES (?, ?, ?, ?)", 
                                (rule_id, acto_id, act_value, i + 1))
                else:
                    cursor.execute("UPDATE Actions SET ActO_ID = ?, Value = ?, Action_Sequence = ? WHERE Action_ID = ?", 
                                (acto_id, act_value, i + 1, aid))

            conn.commit()
            timed_message(edit_form, "Success", "Rule updated successfully", 3000)
            close_edit_rule_form(edit_form)
            form.event_generate("<<RefreshRules>>")
            #logger.debug(f"Saved changes for Rule ID {rule_id}")
        except Exception as e:
            logger.error(f"Error saving rule changes for Rule ID {rule_id}: {e}\n{traceback.format_exc()}")
            messagebox.showerror("Error", f"Failed to save rule changes: {e}", parent=edit_form)

    # Bottom frame for save and close buttons
    buttons_frame = tk.Frame(edit_form, bg="lightgray")
    buttons_frame.grid(row=2, column=0, sticky="ew")

    # Create save and close buttons
    edit_save_button = tk.Button(buttons_frame, text="Save changes", command=save_rule_changes, bg="white", font=("Arial", 10))
    edit_save_button.pack(side="left", padx=50 * scaling_factor, pady=5 * scaling_factor)

    edit_close_button = tk.Button(buttons_frame, text="Close - do not save", command=lambda: close_edit_rule_form(edit_form), bg="white", font=("Arial", 10))
    edit_close_button.pack(side="right", padx=50 * scaling_factor, pady=5 * scaling_factor)
    
    #logger.debug(f"#######  edit_form.action_row: {action_rows}")
    #logger.debug(f"#######  edit_form.next_action_row: {edit_form.next_action_row}")
    
    def close_edit_rule_form(edit_form):
        # Close the edit form
        scrolled_frame.canvas.focus_set()
        scrolled_frame.canvas.bind_all("<MouseWheel>", scrolled_frame._on_mousewheel)
        close_form_with_position(edit_form, conn, cursor, win_id)

    edit_form.update_idletasks()
    logger.debug(f"Opened Edit Rule form for Rule ID {rule_id}")

    form.wait_window(edit_form)



