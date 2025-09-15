# gui_maint_rules.py
# Functions for managing rules forms and groups in the HA project

import tkinter as tk
from tkinter import ttk, messagebox
from tkinter import *
import logging
import traceback
import sqlite3
import time
from db import (fetch_trigger_option, fetch_action_option, create_rule_group, fetch_account_full_name, fetch_category_name, fetch_subcategory_name)
from ui_utils import (VerticalScrolledFrame, resource_path, open_form_with_position, close_form_with_position, sc, Tooltip)
from gui_maint_rule_edit import edit_rule_form
from config import COLORS
import config

# Set up logging
logger = logging.getLogger('HA.gui_maint_rules')

# Handle DPI scaling for high-resolution displays
try:
    from ctypes import windll
    windll.shcore.SetProcessDpiAwareness(1)
except Exception as e:
    logger.warning(f"Failed to set DPI awareness: {e}")
    
# Global image cache to reuse PhotoImage across form instances
_image_cache = {}
    
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
    form.geometry(f"{sc(1600)}x{sc(1000)}")
    logger.info(f"Opening Manage Rules form with scaling factor: {sc(1)}")

    # Add header frame for fixed buttons
    style = ttk.Style()
    style.configure("Debug.TFrame", background=config.master_bg, font=(config.ha_normal))
    header_height = sc(50)
    header_frame = ttk.Frame(form, height=header_height, style="Debug.TFrame")
    header_frame.pack(fill=tk.X, side=tk.TOP)
    header_frame.grid_propagate(False)  # Prevent frame from resizing

    button_frames = {}
    rule_data = {}
    new_rule_buttons = {}
    group_states = {}
    group_heights = {}
    group_rules = {}
    form.ui_elements = {
        'group_frames': [],
        'close_button': None,
        'header_new_rule_group_button': None,
        'global_up_btns': {},
        'global_down_btns': {}
    }
    form.image_refs = []

    # Use cached bitmap images or create new ones
    global _image_cache
    icon_files = [
        "2_edit-48.png",
        "3_trash-48.png",
        "4_see_match-48.png",
        "5_apply_rule-48.png",
        "6_duplicate-48.png",
        "7_up-24.png",
        "8_down-24.png"
    ]
    for icon in icon_files:
        if icon not in _image_cache:
            img = tk.PhotoImage(file=resource_path(f"icons/{icon}"))
            img = img.zoom(sc(1))  # Apply scaling for icons
            _image_cache[icon] = img
            logger.debug(f"Created new PhotoImage for {icon}")
        form.image_refs.append(_image_cache[icon])
        logger.debug(f"Reusing cached PhotoImage for {icon}")

    # Assign cached images
    edit_img = _image_cache["2_edit-48.png"]
    delete_img = _image_cache["3_trash-48.png"]
    see_match_img = _image_cache["4_see_match-48.png"]
    apply_rule_img = _image_cache["5_apply_rule-48.png"]
    duplicate_img = _image_cache["6_duplicate-48.png"]
    up_img = _image_cache["7_up-24.png"]
    down_img = _image_cache["8_down-24.png"]
    
    top_padding = sc(10)  # Reduced to account for header
    group_frame_padding = sc(95)
    group_frame_collapsed = sc(50)
    group_frame_gap = sc(10)
    group_frame_width = sc(1560)
    group_frame_x = sc(20)
    
    tree_height_padding = sc(25)
    tree_x = sc(285)
    tree_y = sc(31)
    tree_width = sc(1260)
    
    max_y = sc(120)
    max_x = sc(1400)
    
    new_rule_group_button_x = sc(20)
    new_rule_group_button_width = sc(150)
    new_rule_group_button_height = sc(28)
    new_rule_group_button_y = sc(10)
    
    close_button_x = sc(1410)
    close_button_width = sc(165)
    close_button_height = sc(28)
    close_button_y = sc(10)
    
    new_rule_button_offset = sc(35)
    new_rule_button_width = sc(165)
    new_rule_button_height = sc(28)
    
    up_down_button_offset = sc(40)
    up_down_button_height = sc(24)
    up_button_x = sc(350)
    down_button_x = sc(400)
    
    icon_size = sc(48)
    bf_y_offset = sc(50)
    bf_width = sc(5 * 53)
    bf_x = sc(12)
    bbox_adjustment = sc(-18)
    
    expand_l_button_x = sc(10)
    expand_r_button_x = sc(1500)
    menu_button_x = sc(1527)
    rg_button_y = sc(5)
    rg_button_size = sc(20)
    
    scrolled_frame_bottom_padding = sc(20)
    
    tree_colw_300 = sc(300)
    tree_colw_65 = sc(65)
    tree_colw_381 = sc(381)
    
    pop_button_width = 48       # scaled inside function
    pop_button_height = 48      # scaled inside function

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
        update_all_group_positions(form, scrolled_frame)            

    def update_all_group_positions(form, scrolled_frame):
        #logger.debug("   ** update_all_group_positions() ** ")
        y_offset = top_padding  # Offset by header height
        maximum_y = 0
        for gid, gf, tree in form.ui_elements['group_frames']:
            if group_states.get(gid, False):
                total_rows = count_visible_rows(tree)
                height = total_rows * icon_size + tree_height_padding + group_frame_padding
                #logger.debug(f"total_rows = {total_rows}, height = {height}")
            else:
                height = group_frame_collapsed
            gf.place(x=group_frame_x, y=y_offset, height=height)
            maximum_y = y_offset + height
            y_offset += height + group_frame_gap
            #logger.debug(f"maximum_y = {maximum_y}")

        form.update_idletasks()
        total_height = maximum_y + scrolled_frame_bottom_padding
        scrolled_frame.interior.configure(height=total_height)
        width = sc(1400)
        scrolled_frame.canvas.config(scrollregion=(0, 0, width, total_height))

    def update_button_positions(group_id, tree, button_frames, new_rule_btn, rules, group_frames, scrolled_frame, close_button, global_up_btns, global_down_btns):
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
            group_frame.configure(height=max_y + sc(10))
            group_heights[group_id] = max_y + sc(10)
            #logger.debug(f"Configured group frame height for Group ID {group_id} to {max_y + sc(10)}")

    def on_tree_click(event, tree, tree_height, scrolled_frame, group_frames, close_button, global_up_btns, global_down_btns, rules, group_id):
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
            trigo_id, trig_value, trig_sequence = t
            trig_option = fetch_trigger_option(cursor, trigo_id)
            if trig_option:
                trig_desc = trig_option[0] if trig_option and len(trig_option) > 0 and len(trig_option[0]) > 0 else f"Unknown Trigger (ID: {trigo_id})"
                trig_desc = trig_desc[:-2]  # Remove trailing dots
                trig_desc_sp = trig_desc.ljust(30)  # Align to 30 chars
                if 1 <= trigo_id <= 3:  # Amount - Numeric field (6d2 format)
                    trigger_lines.append(f'   {trig_desc_sp} £{trig_value}')
                elif 4 <= trigo_id <= 15:  # Tag or Description
                    trigger_lines.append(f'   {trig_desc_sp} "{trig_value}"')
                elif 16 <= trigo_id <= 19:  # Account Name
                    acc_id = int(trig_value)
                    acc_name = fetch_account_full_name(cursor, acc_id, year)
                    trigger_lines.append(f'   {trig_desc}   {acc_name}')
                elif 20 <= trigo_id <= 22:  # Date
                    trigger_lines.append(f'   {trig_desc_sp} {trig_value}')
                elif 23 <= trigo_id <= 27:  # No field required
                    trigger_lines.append(f'   {trig_desc}')
            else:
                logger.warning(f"No description found for TrigO_ID {trigo_id}")
                trigger_lines.append(f'   Unknown Trigger "{trig_value}"')

        # Modify trigger_lines based on Rule_Trigger_Mode (column 3)
        rule_trigger_mode = values[2]  # Get Rule_Trigger_Mode from column 3
        if rule_trigger_mode == "Any":
            trigger_lines = [f"      {trigger_lines[0]}"] + [f"OR {line}" for line in trigger_lines[1:]]
        elif rule_trigger_mode == "ALL":
            trigger_lines = [f"      {trigger_lines[0]}"] + [f"AND{line}" for line in trigger_lines[1:]]
        else:
            logger.warning(f"Invalid Rule_Trigger_Mode: {rule_trigger_mode}")
            
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
        update_button_positions(group_id, tree, button_frames, new_rule_buttons.get(group_id), rules, group_frames, scrolled_frame, form.ui_elements['close_button'], form.ui_elements['global_up_btns'], form.ui_elements['global_down_btns'])
        #logger.debug(f"Updated Treeview with total_rows={total_rows}, tree_height={tree_height}")
        update_all_group_positions(form, scrolled_frame)
    
    def close_rules_form():
        # Clean up form and ensure no resource leaks
        for widget in form.winfo_children():
            widget.destroy()
        form.image_refs = []
        import gc
        gc.collect()
        logger.debug(f"Closed Manage Rules form, GC collected: {gc.get_count()}")
        close_form_with_position(form, conn, cursor, win_id)

    def refresh_rules(event=None):
        #logger.debug("   ** Refreshing Rules form() ** ")
        try:
            preserved_states = group_states.copy()
            #logger.debug(f"Preserved states before refresh: {preserved_states}")
            for widget in form.winfo_children():
                if widget != header_frame:
                    widget.destroy()
            button_frames.clear()
            rule_data.clear()
            new_rule_buttons.clear()
            group_heights.clear()
            group_rules.clear()
            form.ui_elements['group_frames'].clear()
            form.ui_elements['close_button'] = None
            form.ui_elements['header_new_rule_group_button'] = None

            global scrolled_frame
            scrolled_frame = VerticalScrolledFrame(form)
            scrolled_frame.canvas.config(background=config.master_bg)
            scrolled_frame.interior.config(background=config.master_bg)
            scrolled_frame.scrollbar.config(background=config.master_bg)
            scrolled_frame.pack(fill=tk.BOTH, expand=True, pady=(5, 0))

            form.ui_elements['header_new_rule_group_button'] = tk.Button(header_frame, text="New Rule Group", 
                                                                        command=lambda: create_new_group_popup(form, conn, cursor), 
                                                                        font=(config.ha_button))
            form.ui_elements['header_new_rule_group_button'].place(x=new_rule_group_button_x, y=new_rule_group_button_y, 
                                                                width=new_rule_group_button_width, height=new_rule_group_button_height)
            
            form.ui_elements['close_button'] = tk.Button(header_frame, text="Close", 
                                                        command=lambda: close_rules_form(), 
                                                        bg=COLORS["exit_but_bg"], font=(config.ha_button))
            form.ui_elements['close_button'].place(x=close_button_x, y=close_button_y, 
                                                width=close_button_width, height=close_button_height)

            cursor.execute("SELECT Group_ID, Group_Name FROM RuleGroups WHERE Group_Enabled = 1 ORDER BY Group_Sequence")
            groups = cursor.fetchall()
            #logger.debug(f"groups = {groups}")
            if not groups:
                tk.Label(scrolled_frame.interior, text="No Rule Groups found. Click 'New Rule Group' to add one.", 
                        bg="white", font=(config.ha_normal)).place(x=new_rule_group_button_x, y=60)
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
                style.configure("Debug.TLabelframe.Label", font=(config.ha_large), background="white")
                style.configure("Rules.Treeview", rowheight=icon_size, font=(config.ha_normal))
                style.configure("Rules.Treeview.Heading", font=(config.ha_normal_bold))
                group_frame.place(x=group_frame_x, y=0, width=group_frame_width)

                cursor.execute("SELECT Rule_ID, Rule_Name, Rule_Active, Rule_Trigger_Mode, Rule_Proceed FROM Rules WHERE Group_ID = ? ORDER BY Rule_Sequence", (group_id,))
                rules = cursor.fetchall()
                group_rules[group_id] = rules

                tree_height = len(rules) * icon_size + tree_height_padding
                group_heights[group_id] = tree_height + group_frame_padding

                tree = ttk.Treeview(group_frame, columns=("Rule", "Active", "Mode", "Stop", "Triggers", "Actions"), show="headings", style="Rules.Treeview")
                
                # Setup Rule Group buttons
                expand_l_btn = tk.Button(group_frame, text="+", width=2, height=1, font=(config.ha_normal), command=lambda gid=group_id, gf=group_frame, t=tree, rs=rules: 
                    toggle_group_frame(gid, gf, t, new_rule_buttons.get(gid), rs, form, scrolled_frame))
                expand_l_btn.place(x=expand_l_button_x, y=rg_button_y, width=rg_button_size, height=rg_button_size)
                
                expand_r_btn = tk.Button(group_frame, text="+", width=2, height=1, font=(config.ha_normal), command=lambda gid=group_id, gf=group_frame, t=tree, rs=rules: 
                    toggle_group_frame(gid, gf, t, new_rule_buttons.get(gid), rs, form, scrolled_frame))
                expand_r_btn.place(x=expand_r_button_x, y=rg_button_y, width=rg_button_size, height=rg_button_size)

                menu_btn = create_menu_button(group_frame, group_id, menu_button_x, rg_button_y, rg_button_size, conn, cursor, form, scrolled_frame)

                menu_btn = tk.Button(group_frame, text="...", width=2, height=1, font=(config.ha_normal), command=lambda gid=group_id, gf=group_frame, mb=menu_btn: 
                    show_popup_palette(gid, gf, mb, conn, cursor, form, scrolled_frame))
                menu_btn.place(x=menu_button_x, y=rg_button_y, width=rg_button_size, height=rg_button_size)

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

                tree.bind("<Button-1>", lambda e, t=tree, gid=group_id: on_tree_click(e, t, tree_height, scrolled_frame, form.ui_elements['group_frames'], form.ui_elements['close_button'], form.ui_elements['global_up_btns'], form.ui_elements['global_down_btns'], group_rules[gid], gid))
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

                    icons = [edit_img, delete_img, see_match_img, apply_rule_img, duplicate_img]
                    for i, icon in enumerate(icons):
                        try:
                            bg_col = "red" if i == 1 else "lightgray"
                            btn = tk.Button(buttons_frame, image=icon, bg=bg_col, width=icon_size, height=icon_size, 
                                            command=lambda r=rule_id, a=i, g=group_id: button_action(r, a, g, conn, cursor, form))
                            #btn.image = icon
                            btn.place(x=i * sc(53), y=0)
                        except Exception as e:
                            logger.error(f"Error loading icon {icon}: {e}")
                            btn = tk.Button(buttons_frame, text=f"Btn{i+1}", width=6, height=2, font=(config.ha_normal), 
                                            command=lambda r=rule_id, a=i, g=group_id: button_action(r, a, g, conn, cursor, form))
                            btn.place(x=i * sc(53), y=0)

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
                                bg_col = "red" if i == 1 else "lightgray"
                                btn = tk.Button(buttons_frame, image=icon, bg=bg_col, width=icon_size, height=icon_size, 
                                                command=lambda r=new_rule_id, a=i, g=gid: button_action(r, a, g, conn, cursor, form))
                                #btn.image = icon
                                btn.place(x=i * sc(53), y=0)
                            except Exception as e:
                                logger.error(f"Error loading icon {icon} for new rule: {e}")
                                btn = tk.Button(buttons_frame, text=f"Btn{i+1}", width=6, height=2, font=(config.ha_normal), 
                                                command=lambda r=new_rule_id, a=i, g=gid: button_action(r, a, g, conn, cursor, form))
                                btn.place(x=i * sc(53), y=0)
                        
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
                        
                        update_button_positions(gid, tree, button_frames, new_rule_buttons.get(gid), group_rules[gid], form.ui_elements['group_frames'], scrolled_frame, form.ui_elements['close_button'], form.ui_elements['global_up_btns'], form.ui_elements['global_down_btns'])
                        update_all_group_positions(form, scrolled_frame)
                        tree.update_idletasks()
                        buttons_frame.update_idletasks()
                        form.update_idletasks()
                        
                        # Open edit form for the new rule, indicating it's a new rule
                        edit_rule_form(new_rule_id, gid, conn, cursor, form, scrolled_frame, parent, is_new_rule=True)
                        
                    except Exception as e:
                        logger.error(f"Error adding new rule to Group ID {gid}: {e}\n{traceback.format_exc()}")
                        messagebox.showerror("Error", f"Failed to add new rule: {e}", parent=form)

                new_rule_btn = tk.Button(group_frame, text="New Rule", width=sc(100 / 10), font=(config.ha_normal), command=add_new_rule, bg="lightgrey")
                new_rule_buttons[group_id] = new_rule_btn

                form.ui_elements['global_up_btns'][group_id] = tk.Button(group_frame, image=up_img, bg="lightgray", width=icon_size, height=sc(24), 
                                                                        command=lambda g=group_id: move_rule_up_selected(g, conn, cursor, form))
                form.ui_elements['global_down_btns'][group_id] = tk.Button(group_frame, image=down_img, bg="lightgray", width=icon_size, height=sc(24), 
                                                                        command=lambda g=group_id: move_rule_down_selected(g, conn, cursor, form))

                form.ui_elements['group_frames'].append((group_id, group_frame, tree))
                y_offset += group_heights[group_id] if group_states[group_id] else group_frame_collapsed + group_frame_gap

                if group_states[group_id]:
                    tree.update_idletasks()  # Ensure Treeview is ready
                    group_states[group_id] = False
                    toggle_group_frame(group_id, group_frame, tree, new_rule_btn, rules, form, scrolled_frame)

            form.update_idletasks()
            update_all_group_positions(form, scrolled_frame)
            form.update()
            #logger.debug(f"After refresh: group_states = {group_states}")

        except Exception as e:
            logger.error(f"Error in refresh_rules: {e}\n{traceback.format_exc()}")
            tk.Label(scrolled_frame.interior, text=f"Error loading rules: {e}", fg="red", bg="white", 
                    font=(config.ha_normal)).place(x=sc(20), y=60)
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
        try:
            tree.move(item_id, '', index - 1)
            tree.selection_set(item_id)
            logger.debug(f"item_id = {item_id} ")
            rule_id = int(item_id[1:])  # Strip 'I' and convert to integer
            cursor.execute("SELECT Rule_ID, Rule_Sequence FROM Rules WHERE Group_ID = ? AND Rule_ID = ?", (group_id, rule_id))
            rec_1 = cursor.fetchall()
            logger.debug(f"up rec_1 = {rec_1} ")
            cursor.execute("SELECT Rule_ID, Rule_Sequence FROM Rules WHERE Group_ID = ? AND Rule_Sequence = ?", (group_id, rec_1[0][1] - 1))
            rec_2 = cursor.fetchall()
            logger.debug(f"up rec_2 = {rec_2} ")
            cursor.execute("UPDATE Rules SET Rule_Sequence = ? WHERE Rule_ID = ?", (rec_1[0][1], rec_2[0][0]))
            logger.debug(f"updated Rule_ID = {rec_2[0][0]},  Rule_Sequence = {rec_1[0][1]}")
            cursor.execute("UPDATE Rules SET Rule_Sequence = ? WHERE Rule_ID = ?", (rec_2[0][1], rec_1[0][0]))
            logger.debug(f"updated Rule_ID = {rec_1[0][0]},  Rule_Sequence = {rec_2[0][1]}")
            conn.commit()

            # Update group_rules to reflect new order
            cursor.execute("SELECT Rule_ID, Rule_Name, Rule_Active, Rule_Trigger_Mode, Rule_Proceed FROM Rules WHERE Group_ID = ? ORDER BY Rule_Sequence", (group_id,))
            group_rules[group_id] = cursor.fetchall()
            logger.debug(f"Updated group_rules[{group_id}] = {group_rules[group_id]}")

            # Reposition buttons and update layout
            update_button_positions(group_id, tree, button_frames, new_rule_buttons.get(group_id), group_rules[group_id], form.ui_elements['group_frames'], scrolled_frame, form.ui_elements['close_button'], form.ui_elements['global_up_btns'], form.ui_elements['global_down_btns'])
            update_all_group_positions(form, scrolled_frame)
            tree.update_idletasks()
            form.update_idletasks()
        except Exception as e:
            logger.error(f"Error moving Rule ID {item_id}: {e}\n{traceback.format_exc()}")

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
        try:
            tree.move(item_id, '', index + 1)
            tree.selection_set(item_id)
            logger.debug(f"item_id = {item_id} ")
            rule_id = int(item_id[1:])  # Strip 'I' and convert to integer
            cursor.execute("SELECT Rule_ID, Rule_Sequence FROM Rules WHERE Group_ID = ? AND Rule_ID = ?", (group_id, rule_id))
            rec_1 = cursor.fetchall()
            logger.debug(f"down rec_1 = {rec_1} ")
            cursor.execute("SELECT Rule_ID, Rule_Sequence FROM Rules WHERE Group_ID = ? AND Rule_Sequence = ?", (group_id, rec_1[0][1] + 1))
            rec_2 = cursor.fetchall()
            logger.debug(f"down rec_2 = {rec_2} ")
            cursor.execute("UPDATE Rules SET Rule_Sequence = ? WHERE Rule_ID = ?", (rec_1[0][1], rec_2[0][0]))
            cursor.execute("UPDATE Rules SET Rule_Sequence = ? WHERE Rule_ID = ?", (rec_2[0][1], rec_1[0][0]))
            conn.commit()

            # Update group_rules to reflect new order
            cursor.execute("SELECT Rule_ID, Rule_Name, Rule_Active, Rule_Trigger_Mode, Rule_Proceed FROM Rules WHERE Group_ID = ? ORDER BY Rule_Sequence", (group_id,))
            group_rules[group_id] = cursor.fetchall()
            logger.debug(f"Updated group_rules[{group_id}] = {group_rules[group_id]}")

            # Reposition buttons and update layout
            update_button_positions(group_id, tree, button_frames, new_rule_buttons.get(group_id), group_rules[group_id], form.ui_elements['group_frames'], scrolled_frame, form.ui_elements['close_button'], form.ui_elements['global_up_btns'], form.ui_elements['global_down_btns'])
            update_all_group_positions(form, scrolled_frame)
            tree.update_idletasks()
            form.update_idletasks()
        except Exception as e:
            logger.error(f"Error moving Rule ID {item_id}: {e}\n{traceback.format_exc()}")

    def button_action(rule_id, action_index, group_id, conn, cursor, form):
        logger.debug("   ** button_action() ** ")
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

    def create_menu_button(group_frame, group_id, menu_button_x, rg_button_y, rg_button_size, conn, cursor, form, scrolled_frame):
        """Factory function to create menu button with proper lambda reference."""
        menu_btn = tk.Button(group_frame, text="...", width=2, height=1, font=(config.ha_normal),
                            command=lambda: show_popup_palette(group_id, group_frame, menu_btn, conn, cursor, form, scrolled_frame))
        menu_btn.place(x=menu_button_x, y=rg_button_y, width=rg_button_size, height=rg_button_size)
        return menu_btn

    # Functions for Rule Group menu button actions
    def edit_group_name(group_id, group_frame, conn, cursor, parent, scrolled_frame):
        """Create a popup to edit the Rule Group name."""
        logger.debug(f"Editing group name for group_id={group_id}")
        popup = tk.Toplevel(parent)
        popup.title("Edit Rule Group")
        popup.transient(parent)
        popup.grab_set()
        win_id = 26  # Unique Win_ID for Edit Rule Group name popup
        open_form_with_position(popup, conn, cursor, win_id, "Edit Rule Group Name")
        popup.geometry(f"{sc(300)}x{sc(150)}")
        popup.configure(bg=config.master_bg)

        # Get current group name
        cursor.execute("SELECT Group_Name FROM RuleGroups WHERE Group_ID = ?", (group_id,))
        current_name = cursor.fetchone()[0]

        # Create UI elements
        tk.Label(popup, text="Edit Rule Group Name:", bg=config.master_bg, font=config.ha_normal).place(x=sc(30), y=sc(20))
        name_entry = tk.Entry(popup, font=config.ha_normal, width=35)
        name_entry.insert(0, current_name)
        name_entry.place(x=sc(30), y=sc(50))
        name_entry.focus_set()

        def save_group_name():
            new_name = name_entry.get().strip()
            if not new_name:
                messagebox.showerror("Error", "Group name cannot be empty", parent=popup)
                return
            try:
                cursor.execute("SELECT COUNT(*) FROM RuleGroups WHERE Group_Name = ? AND Group_ID != ?", (new_name, group_id))
                if cursor.fetchone()[0] > 0:
                    messagebox.showerror("Error", f"Group name '{new_name}' already exists", parent=popup)
                    return
                cursor.execute("UPDATE RuleGroups SET Group_Name = ? WHERE Group_ID = ?", (new_name, group_id))
                conn.commit()
                #logger.info(f"Updated Rule Group name to '{new_name}' for Group_ID={group_id}")
                close_form_with_position(popup, conn, cursor, win_id)
                parent.event_generate("<<RefreshRules>>")
                #logger.debug("Triggered RefreshRules for edit_group_name")
                parent.lift()
            except sqlite3.IntegrityError as e:
                logger.error(f"IntegrityError updating Rule Group name: {e}")
                messagebox.showerror("Error", f"Group name '{new_name}' already exists", parent=popup)
            except Exception as e:
                logger.error(f"Error updating Rule Group name: {e}")
                messagebox.showerror("Error", f"Failed to update Rule Group: {e}", parent=popup)
                parent.lift()

        tk.Button(popup, text="Save", font=config.ha_button, width=10, command=save_group_name).place(x=sc(30), y=sc(100))
        tk.Button(popup, text="Cancel", font=config.ha_button, bg=COLORS["exit_but_bg"], width=10, 
                command=lambda: [close_form_with_position(popup, conn, cursor, win_id), parent.lift()]).place(x=sc(170), y=sc(100))

    def delete_rule_group(group_id, group_frame, conn, cursor, parent, scrolled_frame):
        """Create a confirmation popup to delete a Rule Group and its associated Rules, Triggers, and Actions."""
        logger.debug(f"Deleting Rule Group for group_id={group_id}")
        popup = tk.Toplevel(parent)
        popup.title("Delete Rule Group")
        popup.transient(parent)
        popup.grab_set()
        win_id = 27  # Unique Win_ID for Delete Rule Group confirmation popup window
        open_form_with_position(popup, conn, cursor, win_id, "Delete Rule Group")
        popup.geometry(f"{sc(300)}x{sc(150)}")
        popup.configure(bg=config.master_bg)

        # Get group name for display
        cursor.execute("SELECT Group_Name FROM RuleGroups WHERE Group_ID = ?", (group_id,))
        group_name = cursor.fetchone()[0]

        # Create UI elements
        tk.Label(popup, text=f"Delete Rule Group '{group_name}'?", bg=config.master_bg, font=config.ha_normal).place(x=sc(30), y=sc(20))
        tk.Label(popup, text="This will delete all associated rules.", bg=config.master_bg, font=config.ha_normal).place(x=sc(30), y=sc(50))

        def confirm_delete():
            try:
                # Delete associated Triggers and Actions via Rules
                cursor.execute("SELECT Rule_ID FROM Rules WHERE Group_ID = ?", (group_id,))
                rule_ids = [row[0] for row in cursor.fetchall()]
                for rule_id in rule_ids:
                    cursor.execute("DELETE FROM Triggers WHERE Rule_ID = ?", (rule_id,))
                    cursor.execute("DELETE FROM Actions WHERE Rule_ID = ?", (rule_id,))
                # Delete Rules
                cursor.execute("DELETE FROM Rules WHERE Group_ID = ?", (group_id,))
                # Delete Rule Group
                cursor.execute("DELETE FROM RuleGroups WHERE Group_ID = ?", (group_id,))
                # Renumber Group_Sequence for remaining groups
                renumber_group_sequences(conn, cursor)
                conn.commit()
                logger.info(f"Deleted Rule Group ID={group_id} and associated Rules, Triggers, Actions")
                close_form_with_position(popup, conn, cursor, win_id)
                parent.event_generate("<<RefreshRules>>")
                #logger.debug("Triggered RefreshRules for delete_rule_group")
                parent.lift()
            except Exception as e:
                logger.error(f"Error deleting Rule Group ID={group_id}: {e}\n{traceback.format_exc()}")
                messagebox.showerror("Error", f"Failed to delete Rule Group: {e}", parent=popup)
                parent.lift()

        tk.Button(popup, text="Yes", font=config.ha_button, width=10, command=confirm_delete).place(x=sc(30), y=sc(100))
        tk.Button(popup, text="No", font=config.ha_button, bg=COLORS["exit_but_bg"], width=10, 
                command=lambda: [close_form_with_position(popup, conn, cursor, win_id), parent.lift()]).place(x=sc(180), y=sc(100))

    def move_rg_up(group_id, group_frame, conn, cursor, parent, scrolled_frame):
        """Move the Rule Group up in the sequence by updating Group_Sequence."""
        logger.debug(f"Moving Rule Group ID={group_id} up")
        try:
            # Get current group's sequence
            cursor.execute("SELECT Group_Sequence FROM RuleGroups WHERE Group_ID = ?", (group_id,))
            current_seq = cursor.fetchone()
            if not current_seq:
                logger.error(f"Group ID {group_id} not found")
                return
            current_seq = current_seq[0]
            if current_seq == 1:
                logger.debug(f"Group ID {group_id} is already at the top")
                return

            # Find the group with the previous sequence
            cursor.execute("SELECT Group_ID, Group_Sequence FROM RuleGroups WHERE Group_Sequence = ? AND Group_Enabled = 1", (current_seq - 1,))
            prev_group = cursor.fetchone()
            if not prev_group:
                logger.error(f"No group found with sequence {current_seq - 1}")
                return

            # Swap sequences
            cursor.execute("UPDATE RuleGroups SET Group_Sequence = ? WHERE Group_ID = ?", (current_seq, prev_group[0]))
            cursor.execute("UPDATE RuleGroups SET Group_Sequence = ? WHERE Group_ID = ?", (prev_group[1], group_id))
            conn.commit()
            #logger.debug(f"Swapped Group_Sequence: Group_ID {group_id} to position {prev_group[1]}, Group_ID {prev_group[0]} to position {current_seq}")

            # Refresh the form
            parent.event_generate("<<RefreshRules>>")
            parent.update_idletasks()  # Force UI update
            #logger.debug("Triggered RefreshRules for move_rg_up")
        except Exception as e:
            logger.error(f"Error moving Rule Group ID {group_id} up: {e}\n{traceback.format_exc()}")

    def move_rg_down(group_id, group_frame, conn, cursor, parent, scrolled_frame):
        """Move the Rule Group down in the sequence by updating Group_Sequence."""
        logger.debug(f"Moving Rule Group ID={group_id} down")
        try:
            # Get current group's sequence
            cursor.execute("SELECT Group_Sequence FROM RuleGroups WHERE Group_ID = ?", (group_id,))
            current_seq = cursor.fetchone()
            if not current_seq:
                logger.error(f"Group ID {group_id} not found")
                return
            current_seq = current_seq[0]

            # Find the group with the next sequence
            cursor.execute("SELECT Group_ID, Group_Sequence FROM RuleGroups WHERE Group_Sequence = ? AND Group_Enabled = 1", (current_seq + 1,))
            next_group = cursor.fetchone()
            if not next_group:
                logger.debug(f"Group ID {group_id} is already at the bottom")
                return

            # Swap sequences
            cursor.execute("UPDATE RuleGroups SET Group_Sequence = ? WHERE Group_ID = ?", (current_seq, next_group[0]))
            cursor.execute("UPDATE RuleGroups SET Group_Sequence = ? WHERE Group_ID = ?", (next_group[1], group_id))
            conn.commit()
            #logger.debug(f"Swapped Group_Sequence: Group_ID {group_id} to position {next_group[1]}, Group_ID {next_group[0]} to position {current_seq}")

            # Refresh the form
            parent.event_generate("<<RefreshRules>>")
            parent.update_idletasks()  # Force UI update
            #logger.debug("Triggered RefreshRules for move_rg_down")            
        except Exception as e:
            logger.error(f"Error moving Rule Group ID {group_id} down: {e}\n{traceback.format_exc()}")

    def show_popup_palette(group_id, group_frame, menu_btn, conn, cursor, parent, scrolled_frame):
        """
        Create a popup button palette with 4 horizontal buttons and a close button.
        
        Args:
            group_id: Identifier for the group (e.g., for button actions).
            group_frame: The parent frame containing the calling button.
            menu_btn: The button widget that triggered the palette (for positioning).
        """        
        # Add tooltip handling in palette
        ttk.Style().configure("Palettebutton.TButton")
        
        # Create Toplevel window for the palette
        palette = tk.Toplevel(group_frame.winfo_toplevel())  # Use root as parent        
        palette.overrideredirect(True)  # Remove titlebar
        palette.attributes("-topmost", True)  # Keep on top
        palette.configure(bg=config.master_bg, borderwidth=1, relief='solid')
        
        # Calculate dimensions
        border = sc(20)  # 20px unscaled border, scaled
        button_width = sc(pop_button_width)
        button_height = sc(pop_button_height)
        total_width = 4 * button_width + 2 * border  # 4 buttons + borders
        total_height = button_height + 2 * border  # Buttons + borders
        close_button_size = sc(20)  # Close button 20x20 unscaled
        
        # Position palette: Align top-right corner with menu_btn's top-right
        button_x = menu_btn.winfo_rootx()  # Screen coordinates
        button_y = menu_btn.winfo_rooty()
        button_width_actual = menu_btn.winfo_width()
        palette_x = button_x + button_width_actual - total_width  # Align right edges
        palette_y = button_y  # Align top edges
        palette.geometry(f"{total_width}x{total_height}+{palette_x}+{palette_y}")
        palette.update_idletasks()  # Ensure position is applied
        #logger.debug(f"Palette created at {palette_x},{palette_y} ({total_width}x{total_height})")        
        
        pop_button_images = []
        icons = ["2_edit-48.png", "3_trash-48.png", "3_up-48.png", "4_down-48.png"]
        for i, icon in enumerate(icons):
            img = tk.PhotoImage(file=resource_path(f"icons/{icon}")).zoom(sc(1))
            pop_button_images.append(img)
        
        # Create frame to hold buttons
        frame = tk.Frame(palette, bg=config.master_bg)
        frame.place(x=border, y=border, width=4 * button_width, height=button_height)
        
        # Create 4 buttons with icons, no spacing
        button1 = tk.Button(frame, image=pop_button_images[0], bg="lightgray", command=lambda: [close_popup(palette), 
                                            edit_group_name(group_id, group_frame, conn, cursor, parent, scrolled_frame)])
        button1.place(x=0, y=0, width=button_width, height=button_height)
        Tooltip(button1, "Edit Rule Group name")
        
        button2 = tk.Button(frame, image=pop_button_images[1], bg="red", command=lambda: [close_popup(palette),
                                            delete_rule_group(group_id, group_frame, conn, cursor, parent, scrolled_frame)])
        button2.place(x=button_width, y=0, width=button_width, height=button_height)
        Tooltip(button2, "Delete Rule Group")
        
        button3 = tk.Button(frame, image=pop_button_images[2], bg="lightgray", command=lambda: [close_popup(palette),
                                            move_rg_up(group_id, group_frame, conn, cursor, parent, scrolled_frame)])
        button3.place(x=2 * button_width, y=0, width=button_width, height=button_height)
        Tooltip(button3, "Move Rule Group up in list")
        
        button4 = tk.Button(frame, image=pop_button_images[3], bg="lightgray", command=lambda: [close_popup(palette),
                                            move_rg_down(group_id, group_frame, conn, cursor, parent, scrolled_frame)])
        button4.place(x=3 * button_width, y=0, width=button_width, height=button_height)
        Tooltip(button4, "Move Rule Group down in list")
        
        # Schedule auto-close
        timer_id = palette.after(5000, lambda: close_popup(palette))
        
        def close_popup(window):
            """Helper function to clean up and close the popup."""
            palette.after_cancel(timer_id)  # Cancel the timer
            window.grab_release()  # Release the grab
            window.destroy()  # Destroy the window
        
        # Create close button ("X") in top-right corner
        close_button = tk.Button(palette, text="X", font=(config.ha_large), bg=config.master_bg, fg='black',
                                command=lambda: close_popup(palette))
        close_button.place(x=total_width - close_button_size, y=1,
                        width=close_button_size - 1, height=close_button_size - 1)
        
        #logger.debug("Popup palette created with 4 buttons and close button")
        
        # Keep references to images to prevent garbage collection
        palette.pop_button_images = pop_button_images
        
    def create_new_group_popup(parent, conn, cursor):                       # Win_ID = 23
        popup = tk.Toplevel(parent)
        popup.title("New Rule Group")
        popup.transient(parent)
        popup.grab_set()
        win_id = 23
        open_form_with_position(popup, conn, cursor, win_id, "New Rule Group")
        popup.geometry(f"{sc(300)}x{sc(150)}")
        popup.configure(bg=config.master_bg)
        tk.Label(popup, text="Enter Rule Group Name:", bg=config.master_bg, font=config.ha_normal).place(x=sc(30), y=sc(20))
        name_entry = tk.Entry(popup, width=35)
        name_entry.place(x=sc(30), y=sc(50))
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

        tk.Button(popup, text="Save", font=config.ha_button, width=10, command=save_group).place(x=sc(30), y=sc(100))
        tk.Button(popup, text="Cancel", font=config.ha_button, bg=COLORS["exit_but_bg"], width=10, 
                command=lambda: [close_form_with_position(popup, conn, cursor, win_id), parent.lift()]).place(x=sc(170), y=sc(100))

    def renumber_group_sequences(conn, cursor):
        """Renumber Group_Sequence in RuleGroups to ensure contiguous sequence."""
        logger.debug("Renumbering Group_Sequence in RuleGroups")
        try:
            cursor.execute("SELECT Group_ID, Group_Sequence FROM RuleGroups WHERE Group_Enabled = 1 ORDER BY Group_Sequence")
            groups = cursor.fetchall()
            for index, (group_id, _) in enumerate(groups, start=1):
                cursor.execute("UPDATE RuleGroups SET Group_Sequence = ? WHERE Group_ID = ?", (index, group_id))
            conn.commit()
            logger.debug(f"Renumbered {len(groups)} groups with contiguous Group_Sequence")
        except Exception as e:
            logger.error(f"Error renumbering Group_Sequence: {e}\n{traceback.format_exc()}")

    form.bind("<<RefreshRules>>", lambda e: refresh_rules())
    try:
        refresh_rules()
        form.update_idletasks()
    except Exception as e:
        logger.error(f"Initial refresh_rules failed: {e}")
        tk.Label(form, text=f"Error initializing rules: {e}", fg="red", bg="white", font=(config.ha_normal)).place(x=sc(20), y=100)
        form.update()

    return form






