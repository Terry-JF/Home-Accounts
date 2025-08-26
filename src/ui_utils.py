#from tkinter import ttk
#import traceback
import sys
import logging
import tkinter as tk
from tkcalendar import Calendar
import os
from db import (get_window_position, save_window_position)

# Colour palettes
COLORS = {
    "white": "#FFFFFF",             # Weekday BG, Active Tab Text, Daily Totals Text, Non-Active Tab BG
    "oldlace": "#DEE7DF",           # Weekend BG
    "yellow": "#FFFF80",            # Flag 1
    "green": "#80FF80",             # Flag 2
    "cyan": "#B9FFFF",              # Flag 4
    "marker": "#ED87ED",            # Row Marker
    "pale_blue": "#ADD8E6",         # was "#DFFFFF" - Exit Button
    "very_pale_blue": "#DFFFFF",    # Main Form BG - was #E6F0FA
    "red": "#DD0000",               # was "#FAA0A0" - Daily Totals OD Text
    "dark_grey": "#5D5D5D",         # Daily Totals BG
    "pink": "#FADBFD",              # Daily Totals OD BG
    "black": "#000000",             # Non-Active Tab Text
    "dark_brown": "#803624",        # Active Tab BG
    "orange": "#FFC993",            # Drill Down Marker
    "grey": "#E0E0E0",              # AHB row 8 background
    "pale_green": "#E0FFE0"         # AHB row 3 background
}

TEXT_COLORS = {                     # used for transaction status
    "Unknown": "gray",
    "Forecast": "#800040",          # Brown
    "Processing": "#0000FF",        # Blue
    "Complete": "#000000"           # Black 
}

# Set up logging
logger = logging.getLogger('HA.ui_utils')

def refresh_grid(tree, rows, marked_rows=None, focus_idx=None, focus_day=None):
    if marked_rows is None:
        marked_rows = set()
    trans_rows = [r for r in rows if r["status"] != "Total"]
    trans_rows.sort(key=lambda x: int(x["values"][1]) if x["values"][1].isdigit() else 0)
    sorted_rows = []
    trans_idx = 0
    for row in rows:
        if row["status"] == "Total":
            sorted_rows.append(row)
        else:
            if trans_idx < len(trans_rows):
                sorted_rows.append(trans_rows[trans_idx])
                trans_idx += 1
    tree.delete(*tree.get_children())

    root = tree.master.master

    for i, row_data in enumerate(sorted_rows):
        values = list(row_data["values"])  # Convert tuple to list for modification
        tags = []
        if row_data["status"] == "Total":
            tags.append("daily_total")
            overdrawn = False
            if hasattr(root, 'credit_limits'):
                for col_idx in range(6, min(17, 6 + len(root.accounts))):  # Cols 6-17
                    try:
                        balance = float(values[col_idx].replace(",", "") or 0.0)
                        acc_idx = col_idx - 6
                        if acc_idx < len(root.credit_limits):
                            if balance < root.credit_limits[acc_idx]:
                                overdrawn = True
                                values[col_idx] = f"\u25B6{balance:,.0f}\u25C0"
                    except ValueError:
                        pass
            if overdrawn:
                tags.append("overdrawn")
            tree.insert("", "end", iid=str(i), values=values, tags=tags)
        else:
            day = values[0].lower()
            is_weekend = day in ["sat", "sun"]
            flag = row_data["flag"]
            status = row_data["status"]
            bg_key = "weekend" if is_weekend else "weekday"
            if i in marked_rows:
                tags.append("marked")
            elif flag & 8:
                tags.append("orange")
            elif flag & 4:
                tags.append("cyan")
            elif flag & 2:
                tags.append("green")
            elif flag & 1:
                tags.append("yellow")
            else:
                tags.append(bg_key)
            if status == "Forecast":
                tags.append("forecast")
            elif status == "Processing":
                tags.append("processing")
            else:
                tags.append("complete")

            tree.insert("", "end", iid=str(i), values=values, tags=tags)

    if focus_day:
        focus_day_int = int(focus_day)
        closest_idx = None
        #closest_day = None
        for i, row in enumerate(sorted_rows):
            if row["status"] != "Total":
                row_day = int(row["values"][1])
                if row_day >= focus_day_int:
                    scroll_offset = max(0, i - 11)
                    tree.yview_moveto(scroll_offset / len(rows))
                    break
                closest_idx = i
                #closest_day = row_day
        else:
            if closest_idx is not None:
                scroll_offset = max(0, closest_idx - 11)
                tree.yview_moveto(scroll_offset / len(rows))
    elif focus_idx is not None and 0 <= focus_idx < len(sorted_rows):
        scroll_offset = max(0, focus_idx - 11)
        tree.yview_moveto(scroll_offset / len(rows))

class VerticalScrolledFrame(tk.Frame):
    def __init__(self, parent, *args, **kw):
        tk.Frame.__init__(self, parent, *args, **kw)
        self.canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0)
        self.scrollbar = tk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.interior = tk.Frame(self.canvas)
        self.interior_id = self.canvas.create_window(0, 0, window=self.interior, anchor="nw")
        
        self.interior.bind("<Configure>", self._configure_interior)
        self.canvas.bind("<Configure>", self._configure_canvas)
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind_all("<Button-4>", self._on_mousewheel)
        self.canvas.bind_all("<Button-5>", self._on_mousewheel)
        self.canvas.focus_set()

    def _on_mousewheel(self, event):
        if self.canvas.winfo_exists():
            if event.num == 4 or event.delta > 0:
                self.canvas.yview_scroll(-1, "units")
                #logger.debug("Mouse wheel up")
            elif event.num == 5 or event.delta < 0:
                self.canvas.yview_scroll(1, "units")
                #logger.debug("Mouse wheel down")
        #logger.debug(f"Mouse wheel event: delta={event.delta}, num={getattr(event, 'num', None)}")
        return "break"

    def _configure_interior(self, event):
        size = (self.interior.winfo_reqwidth(), self.interior.winfo_reqheight())
        self.canvas.config(scrollregion=(0, 0, size[0], size[1]))
        if self.interior.winfo_reqwidth() != self.canvas.winfo_width():
            self.canvas.config(width=self.interior.winfo_reqwidth())
        #logger.debug(f"Interior configured: size={size}, scrollregion={(0, 0, size[0], size[1])}")

    def _configure_canvas(self, event):
        if self.interior.winfo_reqwidth() != self.canvas.winfo_width():
            self.canvas.itemconfigure(self.interior_id, width=self.canvas.winfo_width())
        #logger.debug(f"Canvas configured: width={self.canvas.winfo_width()}, interior_id={self.interior_id}")
        
    def destroy(self):
        # Unbind the mouse wheel event before destruction
        if self.canvas.winfo_exists():
            self.canvas.unbind("<MouseWheel>")
            self.canvas.unbind("<Button-4>")
            self.canvas.unbind("<Button-5>")
        # Call the parent destroy method
        tk.Frame.destroy(self)        

def resource_path(relative_path):
    """Get absolute path to resource, works for dev and PyInstaller"""
    try:
        if getattr(sys, 'frozen', False):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(base_path, relative_path)
    except Exception as e:
        logger.error(f"Error resolving resource path: {e}")
        raise

# Window Management functions
def open_form_with_position(form, conn, cursor, win_id, default_title):
    name, left, top = get_window_position(cursor, win_id)           # get details from DB table
    form.title(name if name else default_title)
    if left is not None and top is not None:
        form.geometry(f"+{left}+{top}")
    else:
        form.geometry("+200+200")  # Default to primary monitor
    form.transient(form.master)
    # form.attributes("-topmost", True)
    # Save initial position if not in DB
    if name is None:
        save_window_position(cursor, conn, win_id, default_title, form.winfo_x(), form.winfo_y())

def close_form_with_position(form, conn, cursor, win_id):
    left = form.winfo_x()
    top = form.winfo_y()
    title = form.title()
    save_window_position(cursor, conn, win_id, title, left, top)
    form.destroy()        

# Calendar Control
def open_calendar(form, value_entry):
    # Open calendar for date selection
    def set_date():
        date = cal.get_date()
        value_entry.delete(0, tk.END)
        value_entry.insert(0, date)
        top.destroy()
    top = tk.Toplevel(form)
    top.transient(form)
    top.grab_set()
    scaling_factor = top.winfo_fpixels('1i') / 96  # Get scaling factor (e.g., 2.0 for 200% scaling)
    center_window(top, int(200 * scaling_factor), int(200 * scaling_factor))
    cal = Calendar(top, date_pattern="dd/mm/yyyy")
    cal.pack(padx=10, pady=10)
    tk.Button(top, text="Select", command=set_date).pack(pady=5)

# Center any Window on screen
def center_window(window, width, height):                               #                           7
    window.update_idletasks()  # Ensure window size is updated
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2
    window.geometry(f"{width}x{height}+{x}+{y}")

# Timed Popup message
def timed_message(parent, title, message, timer_ms):
    """
    Display a timed popup message that auto-closes after timer_ms milliseconds
    or when the Close button is clicked.
    
    Args:
        parent: Parent Tkinter window
        title: String for popup title
        message: String for popup message
        timer_ms: Integer, time in milliseconds before auto-close
    """
    # Create popup window
    popup = tk.Toplevel(parent)
    popup.title(title)
    popup.transient(parent)  # Keep popup on top of parent
    popup.grab_set()  # Make popup modal
    popup.configure(bg="lightgray")

    # Adjust size based on screen scaling
    scaling_factor = parent.winfo_fpixels('1i') / 96
    popup_width = int(200 * scaling_factor)
    popup_height = int(100 * scaling_factor)

    # Center the popup relative to parent
    parent_x = parent.winfo_x()
    parent_y = parent.winfo_y()
    parent_width = parent.winfo_width()
    parent_height = parent.winfo_height()
    x = parent_x + (parent_width - popup_width) // 2
    y = parent_y + (parent_height - popup_height) // 2
    popup.geometry(f"{popup_width}x{popup_height}+{x}+{y}")

    # Add message label
    tk.Label(
        popup,
        text=message,
        font=("Arial", 10),
        wraplength=int(180 * scaling_factor),
        bg="lightgray"
    ).pack(pady=int(10 * scaling_factor), padx=int(10 * scaling_factor))

    # Add Close button
    button = tk.Button(
        popup,
        text="Close",
        font=("Arial", 10),
        width=10,
        command=lambda: close_popup(popup)
    )
    button.pack(pady=int(10 * scaling_factor))

    # Schedule auto-close
    timer_id = popup.after(timer_ms, lambda: close_popup(popup))

    def close_popup(window):
        """Helper function to clean up and close the popup."""
        popup.after_cancel(timer_id)  # Cancel the timer
        window.grab_release()  # Release the grab
        window.destroy()  # Destroy the window

    # Ensure popup is focused
    button.focus_set()
    popup.protocol("WM_DELETE_WINDOW", lambda: close_popup(popup))  # Handle window close button
    popup.wait_window()  # Wait for window to close before returning
    
    
    
    
    
    
    