# gui_maint_rules.py
# Functions for managing rules forms and groups in the HA project

import tkinter as tk
from tkinter import ttk, messagebox
from tkinter import *
import logging
import traceback
import sqlite3
from db import (fetch_trigger_option, fetch_action_option, create_rule_group, fetch_account_full_name, fetch_category_name, fetch_subcategory_name)
from ui_utils import (VerticalScrolledFrame, resource_path, open_form_with_position, close_form_with_position, sc)
from gui_maint_rule_edit import edit_rule_form

# Set up logging
logger = logging.getLogger('HA.gui_maint_rules')

# Handle DPI scaling for high-resolution displays
try:
    from ctypes import windll
    windll.shcore.SetProcessDpiAwareness(1)
except Exception as e:
    logger.warning(f"Failed to set DPI awareness: {e}")
    
###  Rules Maintenance Form and Rules Engine functions  ###

def create_rules_form(parent, conn, cursor):                            # Win_ID = 22
    
    # Get current HA year setting
    year = int(parent.year_var.get())  # Use global year_var
    
    form = tk.Toplevel(parent)
    form.title("Manage Rules")
    form.transient(parent)
    form.grab_set()
    form.configure(bg="lightgray")
    win_id = 22
    open_form_with_position(form, conn, cursor, win_id, "Manage Rules")
    scaling_factor = form.winfo_fpixels('1i') / 96
    form.geometry(f"{int(1600 * scaling_factor)}x{int(1000 * scaling_factor)}")
    logger.info(f"Opening Manage Rules form with scaling factor: {scaling_factor}")

    image_refs = []
    button_frames = {}
    rule_data = {}
    new_rule_buttons = {}
    group_states = {}
    group_heights = {}
    group_rules = {}
    form.ui_elements = {
        'group_frames': [],
        'close_button': None,
        'bottom_new_rule_group_button': None,
        'global_up_btns': {},
        'global_down_btns': {}
    }
    
    top_padding = int(50 * scaling_factor)
    
    group_frame_padding = int(95 * scaling_factor)
    group_frame_collapsed = int(50 * scaling_factor)
    group_frame_gap = int(10 * scaling_factor)
    group_frame_bottom_padding = int(50 * scaling_factor)
    group_frame_width = int(1560 * scaling_factor)
    group_frame_x = int(20 * scaling_factor)
    
    tree_height_padding = int(25 * scaling_factor)
    tree_x = int(285 * scaling_factor)
    tree_y = int(31 * scaling_factor)
    tree_width = int(1260 * scaling_factor)
    
    max_y = int(120 * scaling_factor)
    max_x = int(1400 * scaling_factor)
    
    new_rule_group_button_offset = int(10 * scaling_factor)
    new_rule_group_button_x = int(20 * scaling_factor)
    new_rule_group_button_width = int(150 * scaling_factor)
    new_rule_group_button_height = int(28 * scaling_factor)
    
    close_button_offset = int(10 * scaling_factor)
    close_button_x = int(1400 * scaling_factor)
    close_button_width = int(165 * scaling_factor)
    close_button_height = int(28 * scaling_factor)
    close_button_nogroups = int(110 * scaling_factor)
    
    new_rule_button_offset = int(35 * scaling_factor)
    new_rule_button_width = int(165 * scaling_factor)
    new_rule_button_height = int(28 * scaling_factor)
    
    bottom_buttons_padding = int(10 * scaling_factor)
    
    up_down_button_offset = int(40 * scaling_factor)
    up_down_button_height = int(24 * scaling_factor)
    up_button_x = int(350 * scaling_factor)
    down_button_x = int(400 * scaling_factor)
    
    icon_size = int(48 * scaling_factor)
    bf_y_offset = int(50 * scaling_factor)
    bf_width = int(5 * 53 * scaling_factor)
    bf_x = int(12 * scaling_factor)
    bbox_adjustment = int(-18 * scaling_factor)
    
    expand_button_x = int(1502 * scaling_factor)
    expandl_button_x = int(10 * scaling_factor)
    menu_button_x = int(1527 * scaling_factor)
    expand_menu_button_size = int(20 * scaling_factor)
    
    scrolled_frame_bottom_padding = int(60 * scaling_factor)
    
    tree_colw_300 = int(300 * scaling_factor)
    tree_colw_65 = int(65 * scaling_factor)
    tree_colw_381 = int(381 * scaling_factor)

    def count_visible_rows(tree, item=""):
        """Count all visible rows in the Treeview, including expanded children."""
        #logger.debug("   ** count_visible_rows() ** ")
        if item == "":
            return sum(count_visible_rows(tree, child) for child in tree.get_children())
        else:
            count = 1
            if tree.item(item, "open"):
                for child in tree.get_children(item):
                    count += count_visible_rows(tree, child)
            return count

    def toggle_group_frame(group_id, group_frame, tree, new_rule_btn, rules, form, scrolled_frame):
        #logger.debug("   ** toggle_group_frame() ** ")
        #logger.debug(f"Toggling Group ID {group_id}, current state: {group_states.get(group_id, False)}")
        if group_states.get(group_id, False):
            #logger.debug(f"Collapsing Group ID {group_id}")
            group_states[group_id] = False
            group_frame.configure(height=group_frame_collapsed)
        else:
            #logger.debug(f"Expanding Group ID {group_id}")
            group_states[group_id] = True
            total_rows = count_visible_rows(tree)
            tree_height = total_rows * icon_size + tree_height_padding
            tree.place(x=tree_x, y=tree_y, width=tree_width, height=tree_height)
            
            if new_rule_btn:
                new_rule_btn.place(x=bf_x, y=tree_height + new_rule_button_offset, width=new_rule_button_width, height=new_rule_button_height)
                #logger.debug(f"Placed New Rule button at x={bf_x}, y={tree_height + new_rule_button_offset}")
            if group_id in form.ui_elements.get('global_up_btns', {}):
                form.ui_elements['global_up_btns'][group_id].place(x=up_button_x, y=tree_height + up_down_button_offset, width=icon_size, height=up_down_button_height)
                #logger.debug(f"Placed Up button at x={up_button_x}, y={tree_height + up_down_button_offset}")
            if group_id in form.ui_elements.get('global_down_btns', {}):
                form.ui_elements['global_down_btns'][group_id].place(x=down_button_x, y=tree_height + up_down_button_offset, width=icon_size, height=up_down_button_height)
                #logger.debug(f"Placed Down button at x={down_button_x}, y={tree_height + up_down_button_offset}")
            group_frame.configure(height=tree_height + group_frame_padding)
            
            # Force full update to ensure Treeview is rendered
            tree.update_idletasks()  # Ensure Treeview is rendered
            form.update_idletasks()  # Ensure form is updated
            #logger.debug(f"Tree mapped: {tree.winfo_ismapped()}, New Rule btn mapped: {new_rule_btn.winfo_ismapped() if new_rule_btn else None}")
        if form.ui_elements['close_button'] is not None:
            update_all_group_positions(form, scrolled_frame)            

    def update_all_group_positions(form, scrolled_frame):
        #logger.debug("   ** update_all_group_positions() ** ")
        y_offset = top_padding
        maximum_y = 0
        for gid, gf, tree in form.ui_elements['group_frames']:
            if group_states.get(gid, False):
                total_rows = count_visible_rows(tree)
                height = total_rows * icon_size + tree_height_padding + group_frame_padding
                #logger.debug(f"total_rows = {total_rows}, height = {height}")
            else:
                height = group_frame_collapsed
            gf.place(y=y_offset, height=height)
            maximum_y = y_offset + height
            y_offset += height + group_frame_gap
            #logger.debug(f"maximum_y = {maximum_y}")

        form.ui_elements['close_button'].place(y=maximum_y + close_button_offset)
        form.ui_elements['bottom_new_rule_group_button'].place(y=maximum_y + new_rule_group_button_offset)
        #logger.debug(f"bottom new rule button placed at: y= {maximum_y + new_rule_group_button_offset}")
        form.update_idletasks()
        total_height = maximum_y + scrolled_frame_bottom_padding
        scrolled_frame.interior.configure(height=total_height)
        scrolled_frame.canvas.config(scrollregion=(0, 0, 1400 * scaling_factor, total_height))

    def update_button_positions(group_id, tree, button_frames, new_rule_btn, rules, group_frames, scrolled_frame, close_button, bottom_new_rule_group_button, global_up_btns, global_down_btns):
        #logger.debug("   ** update_button_positions() ** ")
        #logger.debug(f"Button positioning for Group ID {group_id}")
        total_rows = count_visible_rows(tree)
        tree_height = total_rows * icon_size + tree_height_padding
        children = tree.get_children()
        
        if children:
            for rule_id, rule_name, rule_active, rule_trigger_mode, rule_proceed in rules:
                #logger.debug(f"inside update_button_positions() - rule_id = {rule_id}")
                if rule_id in button_frames:
                    bf = button_frames[rule_id]
                    try:
                        item_index = [r[0] for r in rules].index(rule_id)
                        item_id = tree.get_children()[item_index]
                        tree.update_idletasks()
                        bbox = tree.bbox(item_id, 0) if item_id else None
                        tree_y = tree.winfo_y()
                        #logger.debug(f"item_id = {item_id}, bbox = {bbox}, tree_y = {tree_y}")
                        if bbox and bbox != "":
                            #logger.debug("inside - if bbox:")
                            adjusted_y = tree_y + bbox[1] + bbox_adjustment
                            #logger.debug(f"1- tree_y={tree_y}, bbox[1]={bbox[1]}, bbox_adjustment={bbox_adjustment}, bf_y_offset={bf_y_offset}")
                            bf.place(x=bf_x, y=adjusted_y, width=bf_width, height=icon_size)
                            #logger.debug(f"Placed button frame for Rule ID {rule_id} at x={bf_x}, y={adjusted_y}")
                        else:
                            bf.place_forget()
                            logger.warning(f"bbox is empty or None for item_id {item_id}, hiding button frame for Rule ID {rule_id}")
                    except (IndexError, ValueError):
                        logger.warning(f"Rule ID {rule_id} not found in Treeview children, hiding button frame")
                        bf.place_forget()
                        
            if group_states.get(group_id, False):
                if group_id in global_up_btns:
                    global_up_btns[group_id].place(x=up_button_x, y=tree_height + up_down_button_offset, width=icon_size, height=up_down_button_height)
                    #logger.debug(f"Placed Up button at x={up_button_x}, y={tree_height + up_down_button_offset}")
                if group_id in global_down_btns:
                    global_down_btns[group_id].place(x=down_button_x, y=tree_height + up_down_button_offset, width=icon_size, height=up_down_button_height)
                    #logger.debug(f"Placed Down button at x={down_button_x}, y={tree_height + up_down_button_offset}")
                
        else:
            logger.debug(f"No children in Treeview for Group ID {group_id}, skipping icon button positioning")

        if group_states.get(group_id, False):
            if new_rule_btn:
                new_rule_btn.place(x=bf_x, y=tree_height + new_rule_button_offset, width=new_rule_button_width, height=new_rule_button_height)
                #logger.debug(f"Placed New Rule button at x={bf_x}, y={tree_height + new_rule_button_offset}")
                
        group_frame = next((gf for g, gf, _ in group_frames if g == group_id), None)
        if group_frame:
            max_y = max((w.winfo_y() + w.winfo_height() for w in group_frame.winfo_children() if w.winfo_ismapped()), default=0)
            group_frame.configure(height=max_y + int(10 * scaling_factor))
            group_heights[group_id] = max_y + int(10 * scaling_factor)
            #logger.debug(f"Configured group frame height for Group ID {group_id} to {max_y + int(10 * scaling_factor)}")

    def on_tree_click(event, tree, tree_height, scrolled_frame, group_frames, close_button, bottom_new_rule_group_button, rules, group_id, global_up_btns, global_down_btns):
        #logger.debug("   ** on_tree_click() ** ")
        #logger.debug(f"Click at x={event.x}, y={event.y}, event={event.type}, region={tree.identify_region(event.x, event.y)}, item={tree.identify('item', event.x, event.y)}, selection={tree.selection()}, focus={tree.focus()}")
        #logger.debug(f"Treeview bounds: x=0, y=0, width={tree.winfo_width()}, height={tree.winfo_height()}")
        item = tree.identify('item', event.x, event.y)
        if not item:
            #logger.debug("No item identified for click")
            return
        column = tree.identify_column(event.x)
        #logger.debug(f"Click on item {item}, column {column}")
        if column not in ("#5", "#6"):  # Triggers, Actions
            return
        tags = tree.item(item, "tags")
        #logger.debug(f"Item tags: {tags}")
        try:
            rule_id = int(tags[1])
        except IndexError:
            logger.error(f"No rule_id tag found for item {item}, tags: {tags}")
            return
        if rule_id not in rule_data:
            logger.error(f"Rule ID {rule_id} not in rule_data")
            return
        values = tree.item(item, "values")
        triggers = rule_data[rule_id]["triggers"]
        actions = rule_data[rule_id]["actions"]
        current_triggers = values[4]
        current_actions = values[5]
        original_index = tree.index(item)
        trigger_expanded = current_triggers != "   Show Triggers"
        action_expanded = current_actions != "   Show Actions"
        
        # Map TrigO_ID to TrigO_Description for triggers
        trigger_lines = []
        for t in triggers:
            trigo_id, trig_value, trig_sequence = t                     # get trigger option type (id) and related value (if any) for this trigger row
            trig_option = fetch_trigger_option(cursor, trigo_id)        # fetch text description of trigger option
            if trig_option:
                trig_desc = trig_option[0] if trig_option and len(trig_option) > 0 and len(trig_option[0]) > 0 else f"Unknown Trigger (ID: {trigo_id})"
                trig_desc = trig_desc[:-2]                              # remove the 2 dots at the end of description string
                trig_desc_sp = trig_desc.ljust(35)                      # set to be exact size for better alignment
                if 1 <= trigo_id <= 3:                                  # Amount - Numeric field (6d2 format)
                    trigger_lines.append(f'   {trig_desc_sp} £{trig_value}')
                elif 4 <= trigo_id <= 15:                               # Tag or Description
                    trigger_lines.append(f'   {trig_desc_sp} "{trig_value}"')
                elif 16 <= trigo_id <= 19:                              # Account Name
                    acc_id = int(trig_value)
                    acc_name = fetch_account_full_name(cursor, acc_id, year)
                    trigger_lines.append(f'   {trig_desc}   {acc_name}')
                elif 20 <= trigo_id <= 22:                              # Date
                    trigger_lines.append(f'   {trig_desc_sp} {trig_value}')
                elif 23 <= trigo_id <= 27:                              # No field required
                    trigger_lines.append(f'   {trig_desc}')
            else:
                logger.warning(f"No description found for TrigO_ID {trigo_id}")
                trigger_lines.append(f'   Unknown Trigger "{trig_value}"')
        
        # Map ActO_ID to ActO_Description for actions
        action_lines = []
        for a in actions:
            acto_id, act_value, act_sequence = a
            act_option = fetch_action_option(cursor, acto_id)
            if act_option:
                act_desc = act_option[0] if act_option and len(act_option) > 0 and len(act_option[0]) > 0 else f"Unknown Action (ID: {acto_id})"
                act_desc_sp = act_desc.ljust(30)                         # set to be exactly 30 chars
                if acto_id in (1, 2, 10, 11, 12):                       # Text field (40 chars)
                    action_lines.append(f'   {act_desc_sp} "{act_value}"')
                elif acto_id in (3, 4, 5, 6, 7):                        # No field required
                    action_lines.append(f'   {act_desc_sp}')
                elif acto_id == 8:                                      # Two strings for parent and sub-category
                    try:
                        pid, cid = map(int, act_value.split(',')) if act_value and ',' in act_value else (0, 0)
                    except ValueError:
                        pid, cid = 0, 0
                    parent_cat = fetch_category_name(cursor, pid, year)
                    child_cat = fetch_subcategory_name(cursor, pid, cid, year)
                    action_lines.append(f'   {act_desc} - {parent_cat} - {child_cat}')
                elif acto_id in (13, 14):                               # Account name
                    acc_id = int(act_value)
                    acc_name = fetch_account_full_name(cursor, acc_id, year)
                    action_lines.append(f'   {act_desc_sp} {acc_name}')
            else:
                logger.warning(f"No description found for ActO_ID {acto_id}")
                action_lines.append(f'   Unknown Action "{act_value or ""}"')
        
        # Toggle state
        if column == "#5":
            trigger_expanded = not trigger_expanded
            if trigger_expanded:
                tree.column("Triggers", anchor='nw')
            else:
                tree.column("Triggers", anchor='w')
                
        elif column == "#6":
            action_expanded = not action_expanded
            if action_expanded:
                tree.column("Actions", anchor='nw')
            else:
                tree.column("Actions", anchor='w')
            
        # Prepare rule row
        cell_triggers = "\n".join(trigger_lines[:3]) if trigger_expanded and trigger_lines else "   Show Triggers"
        cell_actions = "\n".join(action_lines[:3]) if action_expanded and action_lines else "   Show Actions"
        new_values = (values[0], values[1], values[2], values[3], cell_triggers, cell_actions)
        # Store existing child row values
        existing_children = tree.get_children(item)
        #logger.debug(f"Existing children before delete: {existing_children}")
        child_values = {i: tree.item(child, "values") for i, child in enumerate(existing_children) if tree.exists(child)}
        # Delete parent row
        tree.delete(item)
        new_iid = tree.insert("", original_index, values=new_values, tags=("rule", str(rule_id)))
        # Refresh existing children (re-fetch after delete)
        existing_children = tree.get_children(new_iid)
        #logger.debug(f"Existing children after insert: {existing_children}")
        extra_rows = []
        # Calculate required child rows
        trigger_chunks = [trigger_lines[i:i+3] for i in range(3, len(trigger_lines), 3)] if trigger_expanded else []
        action_chunks = [action_lines[i:i+3] for i in range(3, len(action_lines), 3)] if action_expanded else []
        max_chunks = max(len(trigger_chunks), len(action_chunks))
        # Preserve existing trigger/action content
        for i in range(max_chunks):
            trigger_text = "\n".join(trigger_chunks[i]) if i < len(trigger_chunks) else ""
            action_text = "\n".join(action_chunks[i]) if i < len(action_chunks) else ""
            # Use existing values if available and not toggling off
            if i < len(child_values):
                existing_trigger = child_values.get(i, ("", "", "", "", "", ""))[4]
                existing_action = child_values.get(i, ("", "", "", "", "", ""))[5]
                if trigger_expanded and existing_trigger and not trigger_text:
                    trigger_text = existing_trigger
                if action_expanded and existing_action and not action_text:
                    action_text = existing_action
            if i < len(existing_children) and tree.exists(existing_children[i]):
                # Update existing child row
                tree.item(existing_children[i], values=("", "", "", "", trigger_text, action_text))
                extra_rows.append(existing_children[i])
            else:
                # Create new child row
                extra_rows.append(tree.insert(new_iid, "end", values=("", "", "", "", trigger_text, action_text), tags=("child",)))
            #logger.debug(f"Child row {i+1}: trigger_text='{trigger_text}', action_text='{action_text}', row_id={extra_rows[-1]}")
        # Remove extra child rows
        for child in existing_children[max_chunks:]:
            if tree.exists(child):
                tree.delete(child)
                #logger.debug(f"Deleted extra child row: {child}")
        if new_iid:
            tree.item(new_iid, open=True if extra_rows else False)
        # Update Treeview placement and delegate height adjustments to update_button_positions
        total_rows = count_visible_rows(tree)
        tree_height = total_rows * icon_size + tree_height_padding
        tree.place(x=tree_x, y=tree_y, width=tree_width, height=tree_height)
        tree.update()
        tree.update_idletasks()
        tree.lift()
        # Update button positions and group frame heights
        update_button_positions(group_id, tree, button_frames, new_rule_buttons.get(group_id), rules, group_frames, scrolled_frame, close_button, bottom_new_rule_group_button, global_up_btns, global_down_btns)
        #logger.debug(f"Updated Treeview with total_rows={total_rows}, tree_height={tree_height}")
        update_all_group_positions(form, scrolled_frame)
    
    def refresh_rules(event=None):
        #logger.debug("   ** Refreshing Rules form() ** ")
        try:
            preserved_states = group_states.copy()
            #logger.debug(f"Preserved states before refresh: {preserved_states}")
            for widget in form.winfo_children():
                widget.destroy()
            button_frames.clear()
            rule_data.clear()
            new_rule_buttons.clear()
            group_heights.clear()
            group_rules.clear()
            form.ui_elements['group_frames'].clear()
            form.ui_elements['close_button'] = None
            form.ui_elements['bottom_new_rule_group_button'] = None
            #logger.debug("bottom new rule button placed at: None")

            global scrolled_frame
            scrolled_frame = VerticalScrolledFrame(form)
            scrolled_frame.pack(fill=BOTH, expand=TRUE)

            tk.Button(scrolled_frame.interior, text="New Rule Group", command=lambda: create_new_group_popup(form, conn, cursor), bg="white", 
                    font=("Arial", 10)).place(x=new_rule_group_button_x, y=new_rule_group_button_offset, width=new_rule_group_button_width)
            tk.Button(scrolled_frame.interior, text="Close", command=lambda: close_form_with_position(form, conn, cursor, win_id), bg="white", 
                    font=("Arial", 10)).place(x=close_button_x, y=new_rule_group_button_offset, width=close_button_width)
            cursor.execute("SELECT Group_ID, Group_Name FROM RuleGroups WHERE Group_Enabled = 1 ORDER BY Group_Sequence")
            groups = cursor.fetchall()
            #logger.debug(f"groups = {groups}")
            if not groups:
                tk.Label(scrolled_frame.interior, text="No Rule Groups found. Click 'New Rule Group' to add one.", bg="white", 
                         font=("Arial", 10)).place(x=new_rule_group_button_x, y=60 * scaling_factor)
                tk.Button(scrolled_frame.interior, text="Close", command=lambda: close_form_with_position(form, conn, cursor, win_id), bg="white", width=int(150 / 10 * scaling_factor), 
                        font=("Arial", 10)).place(x=close_button_x, y=close_button_nogroups, width=close_button_width)
                form.ui_elements['close_button'] = tk.Button(scrolled_frame.interior, text="Close", command=lambda: close_form_with_position(form, conn, cursor, win_id), bg="white", width=int(150 / 10 * scaling_factor), 
                                                            font=("Arial", 10))
                form.ui_elements['bottom_new_rule_group_button'] = tk.Button(scrolled_frame.interior, text="New Rule Group", command=lambda: create_new_group_popup(form, conn, cursor), 
                                                                            bg="white", width=int(150 / 10 * scaling_factor), font=("Arial", 10))
                form.update_idletasks()
                scrolled_frame.interior.configure(width=max_x, height=max_y)
                scrolled_frame.canvas.config(scrollregion=(0, 0, max_x, max_y))
                return

            y_offset = top_padding
            for group_id, group_name in groups:
                group_states[group_id] = preserved_states.get(group_id, False)
                #logger.debug(f"Set group_states[{group_id}] = {group_states[group_id]}")
                group_frame = ttk.LabelFrame(scrolled_frame.interior, text=(f"  {group_name}  "), width=group_frame_width)
                group_frame.configure(style="Debug.TLabelframe")
                style = ttk.Style()
                style.configure("Debug.TLabelframe", background="white")
                style.configure("Debug.TLabelframe.Label", font=("Arial", 12), background="white")
                style.configure("Rules.Treeview", rowheight=icon_size, font=("Arial", 10))
                style.configure("Rules.Treeview.Heading", font=("Arial", 10, "bold"))
                group_frame.place(x=group_frame_x, y=y_offset, width=group_frame_width)

                cursor.execute("SELECT Rule_ID, Rule_Name, Rule_Active, Rule_Trigger_Mode, Rule_Proceed FROM Rules WHERE Group_ID = ? ORDER BY Rule_Sequence", (group_id,))
                rules = cursor.fetchall()
                group_rules[group_id] = rules

                tree_height = len(rules) * icon_size + tree_height_padding
                group_heights[group_id] = tree_height + group_frame_padding

                tree = ttk.Treeview(group_frame, columns=("Rule", "Active", "Mode", "Stop", "Triggers", "Actions"), show="headings", style="Rules.Treeview")
                
                expand_btn = tk.Button(group_frame, text="+", width=2, height=1, font=("Arial", 10), command=lambda gid=group_id, gf=group_frame, t=tree, rs=rules: toggle_group_frame(gid, gf, t, new_rule_buttons.get(gid), rs, form, scrolled_frame))
                expand_btn.place(x=expand_button_x, y=0, width=expand_menu_button_size, height=expand_menu_button_size)
                
                expandl_btn = tk.Button(group_frame, text="+", width=2, height=1, font=("Arial", 10), command=lambda gid=group_id, gf=group_frame, t=tree, rs=rules: toggle_group_frame(gid, gf, t, new_rule_buttons.get(gid), rs, form, scrolled_frame))
                expandl_btn.place(x=expandl_button_x, y=5, width=expand_menu_button_size, height=expand_menu_button_size)
                
                menu_btn = tk.Button(group_frame, text="...", width=2, height=1, font=("Arial", 10), command=lambda gid=group_id, gn=group_name: show_group_menu(gid, gn))
                menu_btn.place(x=menu_button_x, y=0, width=expand_menu_button_size, height=expand_menu_button_size)

                tree.heading("Rule", text="Rule Name")
                tree.heading("Active", text="Active")
                tree.heading("Mode", text="All/Any")
                tree.heading("Stop", text="Stop")
                tree.heading("Triggers", text="Rule triggers when")
                tree.heading("Actions", text="Rule will")
                tree.column("Rule", width=tree_colw_300, minwidth=tree_colw_300, stretch=0)
                tree.column("Active", width=tree_colw_65, minwidth=tree_colw_65, stretch=0, anchor='center')
                tree.column("Mode", width=tree_colw_65, minwidth=tree_colw_65, stretch=0, anchor='center')
                tree.column("Stop", width=tree_colw_65, minwidth=tree_colw_65, stretch=0, anchor='center')
                tree.column("Triggers", width=tree_colw_381, minwidth=tree_colw_381, stretch=0, anchor='w')
                tree.column("Actions", width=tree_colw_381, minwidth=tree_colw_381, stretch=0, anchor='w')

                tree.bind("<Button-1>", lambda e, t=tree, gid=group_id: on_tree_click(e, t, tree_height, scrolled_frame, form.ui_elements['group_frames'], form.ui_elements['close_button'], form.ui_elements['bottom_new_rule_group_button'], group_rules[gid], gid, form.ui_elements['global_up_btns'], form.ui_elements['global_down_btns']))
                tree.bind("<<TreeviewSelect>>", lambda e: logger.debug(f"Selected rule: {tree.selection()}"))

                button_y = bf_y_offset
                for rule_id, rule_name, rule_active, rule_trigger_mode, rule_proceed in rules:
                    cursor.execute("SELECT TrigO_ID, Value, Trigger_Sequence FROM Triggers WHERE Rule_ID = ? ORDER BY Trigger_Sequence", (rule_id,))
                    triggers = cursor.fetchall()
                    cursor.execute("SELECT ActO_ID, Value, Action_Sequence FROM Actions WHERE Rule_ID = ? ORDER BY Action_Sequence", (rule_id,))
                    actions = cursor.fetchall()
                    
                    rule_data[rule_id] = {"triggers": triggers, "actions": actions}
                    buttons_frame = ttk.Frame(group_frame, width=bf_width, height=icon_size, style="Debug.TFrame")
                    buttons_frame.place(x=bf_x, y=button_y, width=bf_width, height=icon_size)
                    button_frames[rule_id] = buttons_frame
                    form.ui_elements.setdefault('button_frames', {})[rule_id] = buttons_frame

                    icons = ["2_edit-48.png", "3_trash-48.png", "4_see_match-48.png", "5_apply_rule-48.png", "6_duplicate-48.png"]
                    for i, icon in enumerate(icons):
                        try:
                            img = tk.PhotoImage(file=resource_path(f"icons/{icon}")).zoom(int(scaling_factor))
                            image_refs.append(img)
                            bg_col = "red" if i == 1 else "lightgray"
                            btn = tk.Button(buttons_frame, image=img, bg=bg_col, width=icon_size, height=icon_size, 
                                            command=lambda r=rule_id, a=i, g=group_id: button_action(r, a, g, conn, cursor, form))
                            btn.image = img
                            btn.place(x=i * 53 * scaling_factor, y=0)
                        except Exception as e:
                            logger.error(f"Error loading icon {icon}: {e}")
                            btn = tk.Button(buttons_frame, text=f"Btn{i+1}", width=6, height=2, font=("Arial", 10), 
                                            command=lambda r=rule_id, a=i, g=group_id: button_action(r, a, g, conn, cursor, form))
                            btn.place(x=i * 53 * scaling_factor, y=0)

                    tree.insert("", "end", iid=f"I{rule_id}", values=(rule_name, "Active" if rule_active else "Inactive", rule_trigger_mode, "Proceed" if rule_proceed else "Stop", 
                                                                    "   Show Triggers", "   Show Actions"), tags=("rule", str(rule_id)))
                    tree.item(f"I{rule_id}", tags=("rule", str(rule_id)))
                    button_y += icon_size

                # New Rule button implementation
                def add_new_rule(gid=group_id):
                    try:
                        # Find the correct group_frame and tree for the given group_id
                        group_info = next((g, gf, t) for g, gf, t in form.ui_elements['group_frames'] if g == gid)
                        if not group_info:
                            logger.error(f"Group ID {gid} not found in group_frames")
                            return
                        _, group_frame, tree = group_info

                        # Determine new Rule_Sequence
                        cursor.execute("SELECT MAX(Rule_Sequence) FROM Rules WHERE Group_ID = ?", (gid,))
                        max_sequence = cursor.fetchone()[0]
                        new_sequence = (max_sequence or 0) + 1
                        
                        # Ensure unique Rule_Name
                        base_name = "New Rule"
                        new_name = base_name
                        counter = 1
                        while True:
                            cursor.execute("SELECT COUNT(*) FROM Rules WHERE Rule_Name = ?", (new_name,))
                            if cursor.fetchone()[0] == 0:
                                break
                            counter += 1
                            new_name = f"{base_name} ({counter})"
                            #logger.debug(f"new name = {new_name}")
                        
                        # Insert new rule into Rules table
                        cursor.execute("""
                            INSERT INTO Rules (Group_ID, Rule_Name, Rule_Sequence, Rule_Enabled, Rule_Active, Rule_Trigger_Mode, Rule_Proceed)
                            VALUES (?, ?, ?, 1, 1, 'ALL', 0)
                        """, (gid, new_name, new_sequence))
                        new_rule_id = cursor.lastrowid
                        conn.commit()
                        #logger.debug(f"new_rule_id = {new_rule_id}")
                        
                        # Add new rule to Treeview
                        new_iid = f"I{new_rule_id}"
                        tree.insert("", "end", iid=new_iid, values=(new_name, "Active", "ALL", "Proceed", "   Show Triggers", "   Show Actions"), tags=("rule", str(new_rule_id)))
                        
                        # Calculate button_y using Treeview bbox for precise alignment
                        tree.update_idletasks()
                        bbox = tree.bbox(new_iid, 0) if tree.exists(new_iid) else None
                        button_y = tree_y + bbox[1] if bbox and bbox != "" else bf_y_offset + len(group_rules[gid]) * icon_size
                        #logger.debug(f"2- button_y={button_y}, tree_y={tree_y}, bbox[1]={bbox[1]}, bbox_adjustment={bbox_adjustment}, bf_y_offset={bf_y_offset}")
                        
                        # Create button frame for new rule
                        #logger.debug(f"group_frame = {group_frame}, button_y = {button_y}")
                        buttons_frame = ttk.Frame(group_frame, width=bf_width, height=icon_size, style="Debug.TFrame")
                        buttons_frame.place(x=bf_x, y=button_y, width=bf_width, height=icon_size)
                        button_frames[new_rule_id] = buttons_frame
                        form.ui_elements.setdefault('button_frames', {})[new_rule_id] = buttons_frame
                        
                        # Add action buttons to new rule
                        for i, icon in enumerate(icons):
                            try:
                                img = tk.PhotoImage(file=resource_path(f"icons/{icon}")).zoom(int(scaling_factor))
                                image_refs.append(img)
                                bg_col = "red" if i == 1 else "lightgray"
                                btn = tk.Button(buttons_frame, image=img, bg=bg_col, width=icon_size, height=icon_size, 
                                                command=lambda r=new_rule_id, a=i, g=gid: button_action(r, a, g, conn, cursor, form))
                                btn.image = img
                                btn.place(x=i * 53 * scaling_factor, y=0)
                            except Exception as e:
                                logger.error(f"Error loading icon {icon} for new rule: {e}")
                                btn = tk.Button(buttons_frame, text=f"Btn{i+1}", width=6, height=2, font=("Arial", 10), 
                                                command=lambda r=new_rule_id, a=i, g=gid: button_action(r, a, g, conn, cursor, form))
                                btn.place(x=i * 53 * scaling_factor, y=0)
                        
                        # Update group_rules and rule_data
                        group_rules[gid].append((new_rule_id, new_name, 1, "ALL", 1))
                        rule_data[new_rule_id] = {"triggers": [], "actions": []}
                        
                        # Update UI
                        total_rows = count_visible_rows(tree)
                        tree_height = total_rows * icon_size + tree_height_padding
                        group_heights[gid] = tree_height + group_frame_padding
                        if group_states.get(gid, False):
                            tree.place(x=tree_x, y=tree_y, width=tree_width, height=tree_height)
                            new_rule_buttons[gid].place(x=bf_x, y=tree_height + new_rule_button_offset, width=new_rule_button_width, height=new_rule_button_height)
                            if gid in form.ui_elements['global_up_btns']:
                                form.ui_elements['global_up_btns'][gid].place(x=up_button_x, y=tree_height + up_down_button_offset, width=icon_size, height=up_down_button_height)
                            if gid in form.ui_elements['global_down_btns']:
                                form.ui_elements['global_down_btns'][gid].place(x=down_button_x, y=tree_height + up_down_button_offset, width=icon_size, height=up_down_button_height)
                            group_frame.configure(height=tree_height + group_frame_padding)
                        
                        update_all_group_positions(form, scrolled_frame)
                        tree.update_idletasks()
                        buttons_frame.update_idletasks()
                        form.update_idletasks()
                        
                        # Open edit form for the new rule, indicating it's a new rule
                        edit_rule_form(new_rule_id, gid, conn, cursor, form, scrolled_frame, parent, is_new_rule=True)
                        
                    except Exception as e:
                        logger.error(f"Error adding new rule to Group ID {gid}: {e}\n{traceback.format_exc()}")
                        messagebox.showerror("Error", f"Failed to add new rule: {e}", parent=form)

                new_rule_btn = tk.Button(group_frame, text="New Rule", width=int(100 / 10 * scaling_factor), font=("Arial", 10), command=add_new_rule, bg="lightgrey")
                new_rule_buttons[group_id] = new_rule_btn

                up_img = tk.PhotoImage(file=resource_path("icons/7_up-24.png")).zoom(int(scaling_factor))
                down_img = tk.PhotoImage(file=resource_path("icons/8_down-24.png")).zoom(int(scaling_factor))
                image_refs.extend([up_img, down_img])
                
                form.ui_elements['global_up_btns'][group_id] = tk.Button(group_frame, image=up_img, bg="lightgray", width=icon_size, height=24 * scaling_factor, 
                                                                        command=lambda g=group_id: move_rule_up_selected(g, conn, cursor, form))
                form.ui_elements['global_down_btns'][group_id] = tk.Button(group_frame, image=down_img, bg="lightgray", width=icon_size, height=24 * scaling_factor, 
                                                                        command=lambda g=group_id: move_rule_down_selected(g, conn, cursor, form))

                form.ui_elements['group_frames'].append((group_id, group_frame, tree))
                y_offset += group_heights[group_id] if group_states[group_id] else group_frame_collapsed + group_frame_gap

                if group_states[group_id]:
                    tree.update_idletasks()  # Ensure Treeview is ready
                    group_states[group_id] = False
                    toggle_group_frame(group_id, group_frame, tree, new_rule_btn, rules, form, scrolled_frame)

            form.ui_elements['close_button'] = tk.Button(scrolled_frame.interior, text="Close", command=lambda: close_form_with_position(form, conn, cursor, win_id), bg="white", 
                                                        width=int(150 / 10 * scaling_factor), font=("Arial", 10))
            form.ui_elements['close_button'].place(x=close_button_x, y=y_offset, width=close_button_width, height=close_button_height)

            form.ui_elements['bottom_new_rule_group_button'] = tk.Button(scrolled_frame.interior, text="New Rule Group", command=lambda: create_new_group_popup(form, conn, cursor), 
                                                                        bg="white", width=int(150 / 10 * scaling_factor), font=("Arial", 10))
            form.ui_elements['bottom_new_rule_group_button'].place(x=new_rule_group_button_x, y=y_offset, width=new_rule_group_button_width, 
                                                                height=new_rule_group_button_height)
            #logger.debug(f"bottom new rule button placed at: x= {new_rule_group_button_x}, y= {y_offset}")

            form.update_idletasks()
            update_all_group_positions(form, scrolled_frame)
            form.update()
            #logger.debug(f"After refresh: group_states = {group_states}")

        except Exception as e:
            logger.error(f"Error in refresh_rules: {e}\n{traceback.format_exc()}")
            tk.Label(scrolled_frame.interior, text=f"Error loading rules: {e}", fg="red", bg="white", font=("Arial", 10)).place(x=20 * scaling_factor, y=60 * scaling_factor)
            tk.Button(scrolled_frame.interior, text="Close", command=lambda: close_form_with_position(form, conn, cursor, win_id), bg="white", width=int(150 / 10 * scaling_factor), 
                    font=("Arial", 10)).place(x=close_button_x, y=close_button_nogroups, width=close_button_width)
            form.update_idletasks()
            scrolled_frame.interior.configure(width=max_x, height=max_y)
            scrolled_frame.canvas.config(scrollregion=(0, 0, max_x, max_y))

    def move_rule_up_selected(group_id, conn, cursor, form):
        #logger.debug("   ** move_rule_up_selected() ** ")
        for gid, group_frame, _ in form.ui_elements['group_frames']:
            if gid == group_id:
                tree = [w for w in group_frame.winfo_children() if isinstance(w, ttk.Treeview)][0]
                break
        selected = tree.selection()
        if not selected:
            return
        item_id = selected[0]
        index = tree.index(item_id)
        if index == 0:
            return
        tree.move(item_id, '', index - 1)
        tree.selection_set(item_id)
        #logger.debug(f"item_id = {item_id} ")
        rule_id = int(item_id[1:])  # Strip 'I' and convert to integer
        cursor.execute("SELECT Rule_ID, Rule_Sequence FROM Rules WHERE Group_ID = ? AND Rule_ID = ?", (group_id, rule_id))
        rec_1 = cursor.fetchall()
        #logger.debug(f"up rec_1 = {rec_1} ")
        cursor.execute("SELECT Rule_ID, Rule_Sequence FROM Rules WHERE Group_ID = ? AND Rule_Sequence = ?", (group_id, rec_1[0][1] - 1))
        rec_2 = cursor.fetchall()
        #logger.debug(f"up rec_2 = {rec_2} ")
        cursor.execute("UPDATE Rules SET Rule_Sequence = ? WHERE Rule_ID = ?", (rec_1[0][1], rec_2[0][0]))
        cursor.execute("UPDATE Rules SET Rule_Sequence = ? WHERE Rule_ID = ?", (rec_2[0][1], rec_1[0][0]))
        conn.commit()

    def move_rule_down_selected(group_id, conn, cursor, form):
        #logger.debug("   ** move_rule_down_selected() ** ")
        for gid, group_frame, _ in form.ui_elements['group_frames']:
            if gid == group_id:
                tree = [w for w in group_frame.winfo_children() if isinstance(w, ttk.Treeview)][0]
                break
        selected = tree.selection()
        if not selected:
            return
        item_id = selected[0]
        index = tree.index(item_id)
        if index == len(tree.get_children()) - 1:
            return
        tree.move(item_id, '', index + 1)
        tree.selection_set(item_id)
        #logger.debug(f"item_id = {item_id} ")
        rule_id = int(item_id[1:])  # Strip 'I' and convert to integer
        cursor.execute("SELECT Rule_ID, Rule_Sequence FROM Rules WHERE Group_ID = ? AND Rule_ID = ?", (group_id, rule_id))
        rec_1 = cursor.fetchall()
        #logger.debug(f"down rec_1 = {rec_1} ")
        cursor.execute("SELECT Rule_ID, Rule_Sequence FROM Rules WHERE Group_ID = ? AND Rule_Sequence = ?", (group_id, rec_1[0][1] + 1))
        rec_2 = cursor.fetchall()
        #logger.debug(f"down rec_2 = {rec_2} ")
        cursor.execute("UPDATE Rules SET Rule_Sequence = ? WHERE Rule_ID = ?", (rec_1[0][1], rec_2[0][0]))
        cursor.execute("UPDATE Rules SET Rule_Sequence = ? WHERE Rule_ID = ?", (rec_2[0][1], rec_1[0][0]))
        conn.commit()

    def button_action(rule_id, action_index, group_id, conn, cursor, form):
        #logger.debug("   ** button_action() ** ")
        actions = {
            0: lambda r, g: edit_rule_form(r, g, conn, cursor, form, scrolled_frame, parent, is_new_rule=False),
            1: lambda r, g: delete_rule(r, g, conn, cursor, form),
            2: lambda r, g: messagebox.showinfo("See Matching", f"Show matching transactions for rule {r}", parent=form),
            3: lambda r, g: messagebox.showinfo("Apply Rule", f"Apply rule {r} to selected transactions", parent=form),
            4: lambda r, g: duplicate_rule(r, g, conn, cursor, form)
        }
        action = actions.get(action_index)
        if action:
            action(rule_id, group_id)

    def delete_rule(rule_id, group_id, conn, cursor, form):
        #logger.debug("   ** delete_rule() ** ")
        for gid, group_frame, _ in form.ui_elements['group_frames']:
            if gid == group_id:
                tree = [w for w in group_frame.winfo_children() if isinstance(w, ttk.Treeview)][0]
                break
        item_id = "I" + str(rule_id)
        if item_id not in tree.get_children():
            return
        tree.selection_set(item_id)
        tree.focus_set()
        form.update_idletasks()
        if messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete rule '{tree.item(item_id, 'values')[0]}' (ID: {rule_id})?", parent=form):
            try:
                cursor.execute("DELETE FROM Triggers WHERE Rule_ID = ?", (rule_id,))
                cursor.execute("DELETE FROM Actions WHERE Rule_ID = ?", (rule_id,))
                cursor.execute("DELETE FROM Rules WHERE Rule_ID = ?", (rule_id,))
                conn.commit()
                renumber_sequences(group_id, conn, cursor)
                form.event_generate("<<RefreshRules>>")
            except Exception as e:
                logger.error(f"Error deleting Rule ID {rule_id}: {e}\n{traceback.format_exc()}")
        else:
            tree.selection_remove(item_id)
            form.update_idletasks()
            form.event_generate("<<RefreshRules>>")

    def duplicate_rule(rule_id, group_id, conn, cursor, form):
        #logger.debug("   ** duplicate_rule() ** ")
        try:
            cursor.execute("SELECT Rule_Name, Rule_Sequence, Rule_Enabled, Rule_Active, Rule_Trigger_Mode, Rule_Proceed FROM Rules WHERE Rule_ID = ?", (rule_id,))
            original_rule = cursor.fetchone()
            if not original_rule:
                logger.error(f"Rule ID {rule_id} not found")
                return
            base_name = f"Copy of {original_rule[0]}"
            new_name = base_name
            counter = 1
            while True:
                cursor.execute("SELECT COUNT(*) FROM Rules WHERE Rule_Name = ?", (new_name,))
                if cursor.fetchone()[0] == 0:
                    break
                counter += 1
                new_name = f"{base_name} ({counter})"
            new_sequence = original_rule[1] + 1
            cursor.execute("UPDATE Rules SET Rule_Sequence = Rule_Sequence + 1 WHERE Group_ID = ? AND Rule_Sequence >= ?", (group_id, new_sequence))
            cursor.execute("INSERT INTO Rules (Group_ID, Rule_Name, Rule_Sequence, Rule_Enabled, Rule_Active, Rule_Trigger_Mode, Rule_Proceed) VALUES (?, ?, ?, ?, ?, ?, ?)", 
                        (group_id, new_name, new_sequence, original_rule[2], original_rule[3], original_rule[4], original_rule[5]))
            new_rule_id = cursor.lastrowid
            cursor.execute("SELECT TrigO_ID, Value, Trigger_Sequence FROM Triggers WHERE Rule_ID = ? ORDER BY Trigger_Sequence", (rule_id,))
            for trigger in cursor.fetchall():
                cursor.execute("INSERT INTO Triggers (Rule_ID, TrigO_ID, Value, Trigger_Sequence) VALUES (?, ?, ?, ?)", (new_rule_id, trigger[0], trigger[1], trigger[2]))
            cursor.execute("SELECT ActO_ID, Value, Action_Sequence FROM Actions WHERE Rule_ID = ? ORDER BY Action_Sequence", (rule_id,))
            for action in cursor.fetchall():
                cursor.execute("INSERT INTO Actions (Rule_ID, ActO_ID, Value, Action_Sequence) VALUES (?, ?, ?, ?)", (new_rule_id, action[0], action[1], action[2]))
            conn.commit()
            form.event_generate("<<RefreshRules>>")
        except Exception as e:
            logger.error(f"Error duplicating Rule ID {rule_id}: {e}\n{traceback.format_exc()}")

    def renumber_sequences(group_id, conn, cursor):
        #logger.debug("   ** renumber_sequences() ** ")
        try:
            cursor.execute("SELECT Rule_ID FROM Rules WHERE Group_ID = ? ORDER BY Rule_Sequence", (group_id,))
            rules = cursor.fetchall()
            for index, rule in enumerate(rules, start=1):
                cursor.execute("UPDATE Rules SET Rule_Sequence = ? WHERE Rule_ID = ?", (index, rule[0]))
            conn.commit()
        except Exception as e:
            logger.error(f"Error renumbering sequences for Group ID {group_id}: {e}\n{traceback.format_exc()}")

    def show_group_menu(group_id, group_name):
        # Placeholder for group menu
        #logger.debug(f"Showing menu for Group ID {group_id}, Name: {group_name}")
        messagebox.showinfo("Group Menu", f"Menu for Group {group_name} (ID: {group_id})")

    def create_new_group_popup(parent, conn, cursor):                       # Win_ID = 23
        popup = tk.Toplevel(parent)
        popup.title("New Rule Group")
        popup.transient(parent)
        popup.grab_set()
        win_id = 23
        open_form_with_position(popup, conn, cursor, win_id, "New Rule Group")
        scaling_factor = popup.winfo_fpixels('1i') / 96
        popup.geometry(f"{int(250 * scaling_factor)}x{int(150 * scaling_factor)}")
        tk.Label(popup, text="Group Name:").place(x=20 * scaling_factor, y=20 * scaling_factor)
        name_entry = tk.Entry(popup, width=30)
        name_entry.place(x=20 * scaling_factor, y=50 * scaling_factor)
        name_entry.focus_set()

        def save_group():
            name = name_entry.get().strip()
            if not name:
                messagebox.showerror("Error", "Group name cannot be empty", parent=popup)
                return
            try:
                cursor.execute("SELECT COUNT(*) FROM RuleGroups WHERE Group_Name = ?", (name,))
                if cursor.fetchone()[0] > 0:
                    messagebox.showerror("Error", f"Group name '{name}' already exists", parent=popup)
                    return
                create_rule_group(cursor, conn, name)
                logger.info(f"Created Rule Group: {name}")
                close_form_with_position(popup, conn, cursor, win_id)
                parent.event_generate("<<RefreshRules>>")
                parent.lift()
            except sqlite3.IntegrityError as e:
                logger.error(f"IntegrityError saving Rule Group: {e}")
                messagebox.showerror("Error", f"Group name '{name}' already exists", parent=popup)
            except Exception as e:
                logger.error(f"Error saving Rule Group: {e}")
                messagebox.showerror("Error", f"Failed to save Rule Group: {e}", parent=popup)
                parent.lift()

        tk.Button(popup, text="Save", command=save_group).place(x=80 * scaling_factor, y=100 * scaling_factor)
        tk.Button(popup, text="Cancel", command=lambda: [close_form_with_position(popup, conn, cursor, win_id), parent.lift()]).place(x=160 * scaling_factor, y=100 * scaling_factor)

    form.bind("<<RefreshRules>>", lambda e: refresh_rules())
    try:
        refresh_rules()
        form.update_idletasks()
    except Exception as e:
        logger.error(f"Initial refresh_rules failed: {e}")
        tk.Label(form, text=f"Error initializing rules: {e}", fg="red", bg="white", font=("Arial", 10)).place(x=20 * scaling_factor, y=100 * scaling_factor)
        form.update()

    return form




