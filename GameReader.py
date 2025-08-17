###
###  I know.. This code is.. well not great, all made with AI, but it works. feel free to make any changes!
###
###  FIXED: Right and Left Ctrl keys now properly distinguished using scan code detection
###  The keyboard.is_pressed() function doesn't reliably distinguish left/right modifier keys
###  on Windows, so we now use scan codes (29 for Left Ctrl, 157 for Right Ctrl) for accurate detection.
###

# Standard library imports
import datetime
import io
import json
import os
import re
import sys
import tempfile
import threading
import time
import webbrowser
from functools import partial
from tkinter import filedialog, messagebox, simpledialog, ttk

# Third-party imports
import keyboard
import mouse
import pyttsx3
import pytesseract
import requests
import tkinter as tk
import win32api
import win32com.client
import win32con
import win32gui
import win32ui
import win32process
from PIL import Image, ImageEnhance, ImageFilter, ImageGrab, ImageTk
import ctypes
import winsound
import asyncio
import queue

# Controller support
try:
    import inputs
    CONTROLLER_AVAILABLE = True
    print("Controller support enabled - 'inputs' library loaded successfully")
except ImportError:
    CONTROLLER_AVAILABLE = False
    print("Warning: 'inputs' library not available. Controller support disabled.")
    print("To enable controller support, install with: pip install inputs")
try:
    # winsdk is the package name; modules generally import from winrt.*
    import importlib
    UWP_TTS_AVAILABLE = False
    _uwp_import_error = None
    try:
        from winsdk.windows.media.speechsynthesis import SpeechSynthesizer
        from winsdk.windows.storage.streams import DataReader
        UWP_TTS_AVAILABLE = True
    except Exception as e:
        _uwp_import_error = e
        try:
            # Attempt to import winsdk meta package and retry
            importlib.import_module('winsdk')
            from winsdk.windows.media.speechsynthesis import SpeechSynthesizer
            from winsdk.windows.storage.streams import DataReader
            UWP_TTS_AVAILABLE = True
            _uwp_import_error = None
        except Exception as e2:
            _uwp_import_error = e2
            # As a last attempt, try alternate import path (rare)
            try:
                from winsdk.windows.media.speechsynthesis import SpeechSynthesizer  # type: ignore
                from winsdk.windows.storage.streams import DataReader  # type: ignore
                UWP_TTS_AVAILABLE = True
                _uwp_import_error = None
            except Exception as e3:
                _uwp_import_error = e3
except Exception as _e_init:
    UWP_TTS_AVAILABLE = False
    _uwp_import_error = _e_init

def _ensure_uwp_available():
    global UWP_TTS_AVAILABLE
    if UWP_TTS_AVAILABLE:
        return True
    try:
        import importlib
        # Try both ways
        try:
            importlib.import_module('winsdk')
        except Exception:
            pass
        try:
            from winsdk.windows.media.speechsynthesis import SpeechSynthesizer as _SS  # noqa: F401
            from winsdk.windows.storage.streams import DataReader as _DR  # noqa: F401
            UWP_TTS_AVAILABLE = True
        except Exception:
            from winsdk.windows.media.speechsynthesis import SpeechSynthesizer as _SS  # type: ignore # noqa: F401
            from winsdk.windows.storage.streams import DataReader as _DR  # type: ignore # noqa: F401
            UWP_TTS_AVAILABLE = True
    except Exception as e:
        UWP_TTS_AVAILABLE = False
        try:
            print(f"UWP import error: {e}")
        except Exception:
            pass
    return UWP_TTS_AVAILABLE

# Simple stub functions to replace removed complex functions
def get_current_keyboard_layout():
    """Stub function - always returns None for simplicity"""
    return None

def normalize_key_name(key_name, scan_code=None):
    """Stub function - returns key_name as-is for simplicity"""
    return key_name

def is_special_character(key_name):
    """Check if a key name contains special characters that may cause issues"""
    if not key_name:
        return False
    
    # Check for Nordic/Special characters that commonly cause issues
    special_chars = ['å', 'ä', 'ö', '¨', '´', '`', '~', '^', '°', '§', '±', 'µ', '¶', '·', '¸', '¹', '²', '³']
    
    # Check for any special characters in the key name
    for char in special_chars:
        if char in key_name:
            return True
    
    # Check for other potentially problematic characters
    if any(ord(char) > 127 for char in key_name):  # Non-ASCII characters
        return True
    
    return False

def suggest_alternative_key(special_char):
    """Suggest alternative keys for special characters"""
    alternatives = {
        'å': 'a',
        'ä': 'a', 
        'ö': 'o',
        '¨': 'u',
        '´': "'",
        '`': "'",
        '~': '~',
        '^': '^',
        '°': 'o',
        '§': 's',
        '±': '=',
        'µ': 'u',
        '¶': 'p',
        '·': '.',
        '¸': ',',
        '¹': '1',
        '²': '2',
        '³': '3'
    }
    
    return alternatives.get(special_char, None)

def detect_ctrl_keys():
    """
    Detect which Ctrl keys are currently pressed using scan code detection.
    Returns a tuple of (left_ctrl_pressed, right_ctrl_pressed).
    This function provides more reliable left/right distinction than keyboard.is_pressed().
    """
    left_ctrl_pressed = False
    right_ctrl_pressed = False
    
    try:
        # Check if any Ctrl key is pressed first
        if keyboard.is_pressed('ctrl'):
            # Use scan code to determine which one
            for event in keyboard._listener.pressed_events:
                if hasattr(event, 'scan_code'):
                    if event.scan_code == 29:  # Left Ctrl
                        left_ctrl_pressed = True
                    elif event.scan_code == 157:  # Right Ctrl
                        right_ctrl_pressed = True
            
            # Fallback: if scan code detection fails, assume left
            if not left_ctrl_pressed and not right_ctrl_pressed:
                left_ctrl_pressed = True
    except Exception:
        # Fallback to basic detection
        if keyboard.is_pressed('ctrl'):
            left_ctrl_pressed = True
    
    return left_ctrl_pressed, right_ctrl_pressed



# Try to import tkinterdnd2 for drag and drop functionality
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    TKDND_AVAILABLE = True
except ImportError:
    TKDND_AVAILABLE = False
    print("Warning: tkinterdnd2 not available. Drag and drop functionality will be disabled.")

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)  # FIX DPI ON WINDOWS
except AttributeError:
    ctypes.windll.user32.SetProcessDPIAware()
except Exception as e:
    print(f"Warning: Could not set DPI awareness: {e}")

APP_VERSION = "0.8.6"

CHANGELOG = """
0.8.6:
New Features:
- Controller support added.
- More secure layout loading.
- Custom hotkey combinations (e.g., Shift + S).

UI Improvements:
- The Info/Help window now more clearly indicates whether OCR is installed.
- The Info/Help window now includes information on how to add more voices.
- Changed the area name dialog to be more user friendly.
- Added numbers to the voice selection dropdown.

Bug Fixes:
- Fixed a race condition that allowed multiple area selection windows to open at once.


Thanks to everyone who has sent in feedback and bug reports!

Thanks for using GameReader!
"""


# Create a StringIO buffer to capture print statements
log_buffer = io.StringIO()

# Redirect standard output to the StringIO buffer
sys.stdout = log_buffer

# --- Custom Hotkey Conflict Warning Dialog (No Symbol, Styled OK Button) ---
def show_thinkr_warning(game_reader, area_name):
    # Disable all hotkeys when dialog is shown
    try:
        keyboard.unhook_all()
        mouse.unhook_all()
    except Exception as e:
        print(f"Error disabling hotkeys for warning dialog: {e}")

    win = tk.Toplevel(game_reader.root)
    win.title("Hotkey Conflict Detected!")
    win.geometry("370x170")
    win.resizable(False, False)
    win.grab_set()
    win.transient(game_reader.root)
    
    # Set the window icon
    try:
        icon_path = os.path.join(os.path.dirname(__file__), 'Assets', 'icon.ico')
        if os.path.exists(icon_path):
            win.iconbitmap(icon_path)
    except Exception as e:
        print(f"Error setting warning dialog icon: {e}")

    # Center the dialog
    win.update_idletasks()
    x = game_reader.root.winfo_rootx() + game_reader.root.winfo_width() // 2 - 185
    y = game_reader.root.winfo_rooty() + game_reader.root.winfo_height() // 2 - 85
    win.geometry(f"370x170+{x}+{y}")

    # Remove the warning icon (if any)
    for child in win.winfo_children():
        if isinstance(child, tk.Label) and child.cget("image"):
            child.destroy()

    # Add a message label
    msg = tk.Label(win, text=f"This key is already used by area:\n'{area_name}'.\n\nPlease choose a different hotkey.", font=("Helvetica", 12), wraplength=340, justify="center")
    msg.pack(pady=(28, 6))

    # Add OK button
    btn = tk.Button(win, text="OK", width=12, height=1, font=("Helvetica", 11, "bold"), relief="raised", bd=2)
    btn.pack(pady=(6, 10))

    # Focus the button for keyboard users
    btn.focus_set()

    # Bind Enter key to OK
    win.bind("<Return>", lambda e: win.destroy())

    # Disable all hotkeys while the dialog is open
    try:
        keyboard.unhook_all()
        mouse.unhook_all()
    except Exception as e:
        print(f"Error disabling hotkeys: {e}")

    # Restore hotkeys when dialog is closed
    def on_close():
        try:
            game_reader.restore_all_hotkeys()
        except Exception as e:
            print(f"Error restoring hotkeys: {e}")
        win.destroy()

    win.protocol("WM_DELETE_WINDOW", on_close)
    # Also patch the OK button and <Return> binding to use on_close
    btn.config(command=on_close)
    win.bind("<Return>", lambda e: on_close())

class ConsoleWindow:
    def __init__(self, root, log_buffer, layout_file_var, latest_images, latest_area_name_var):
        self.window = tk.Toplevel(root)
        self.window.title("Debug Console")
        
        # Set the window icon
        try:
            icon_path = os.path.join(os.path.dirname(__file__), 'Assets', 'icon.ico')
            if os.path.exists(icon_path):
                self.window.iconbitmap(icon_path)
        except Exception as e:
            print(f"Error setting console window icon: {e}")
        
        self.latest_images = latest_images
        self.window.geometry("690x500")  # Initial size, will adjust based on image

        # Create a top frame for controls
        top_frame = tk.Frame(self.window)
        top_frame.pack(fill='x', padx=10, pady=5)

        # Add checkbox for image display
        self.show_image_var = tk.BooleanVar(value=True)
        self.image_checkbox = tk.Checkbutton(
            top_frame,
            text="Show last processed image",
            variable=self.show_image_var,
            command=self.update_image_display
        )
        self.image_checkbox.pack(side='left')
        
        # Add scale dropdown
        scale_frame = tk.Frame(top_frame)
        scale_frame.pack(side='left', padx=10)
        
        tk.Label(scale_frame, text="Scale:").pack(side='left')
        self.scale_var = tk.StringVar(value="100")
        scales = [str(i) for i in range(10, 101, 10)]  # Creates ["10", "20", ..., "100"]
        scale_menu = tk.OptionMenu(scale_frame, self.scale_var, *scales, command=self.update_image_display)
        scale_menu.pack(side='left')
        tk.Label(scale_frame, text="%").pack(side='left')

        # Add Save Log button
        save_log_button = tk.Button(top_frame, text="Save Log", command=self.save_log)
        save_log_button.pack(side='left', padx=(10, 0))

        # Add Clear Console button
        clear_console_button = tk.Button(top_frame, text="Clear Console", command=self.clear_console)
        clear_console_button.pack(side='left', padx=(10, 0))

        # Add Save Image button
        save_image_button = tk.Button(top_frame, text="Save Image", command=self.save_image)
        save_image_button.pack(side='left', padx=(10, 0))

        # Create a middle frame for image display
        image_frame = tk.Frame(self.window)
        image_frame.pack(fill='x', padx=10, pady=5)
        
        # Add image label to the middle frame
        self.image_label = tk.Label(image_frame)
        self.image_label.pack(fill='x')

        # Create a bottom frame for the log output
        log_frame = tk.Frame(self.window)
        log_frame.pack(fill='both', expand=True, padx=10, pady=5)

        # Add text widget for log output
        self.text_widget = tk.Text(log_frame)
        self.text_widget.pack(fill='both', expand=True)
        self.text_widget.config(state=tk.DISABLED)

        # Enable mouse wheel scrolling for the debug log
        def _on_mousewheel_debug(event):
            self.text_widget.yview_scroll(int(-1 * (event.delta / 120)), 'units')
            return "break"
        def _bind_mousewheel_debug(event):
            self.text_widget.bind_all('<MouseWheel>', _on_mousewheel_debug)
        def _unbind_mousewheel_debug(event):
            self.text_widget.unbind_all('<MouseWheel>')
        self.text_widget.bind('<Enter>', _bind_mousewheel_debug)
        self.text_widget.bind('<Leave>', _unbind_mousewheel_debug)

        # Add right-click context menu
        self.context_menu = tk.Menu(self.text_widget, tearoff=0)
        self.context_menu.add_command(label="Copy", command=self.copy_selection)
        self.context_menu.add_command(label="Select All", command=self.select_all)
        self.text_widget.bind("<Button-3>", self.show_context_menu)

        self.log_buffer = log_buffer
        self.layout_file_var = layout_file_var
        self.latest_area_name_var = latest_area_name_var
        self.photo = None  # Keep a reference to prevent garbage collection

        # Add line limit constant
        self.MAX_LINES = 250
        
        self.update_console()

    def show_context_menu(self, event):
        """Show the context menu at the mouse position."""
        try:
            self.context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.context_menu.grab_release()

    def copy_selection(self):
        """Copy selected text to clipboard."""
        try:
            selected_text = self.text_widget.get("sel.first", "sel.last")
            self.window.clipboard_clear()
            self.window.clipboard_append(selected_text)
        except tk.TclError:
            pass  # No text selected

    def select_all(self):
        """Select all text in the widget."""
        self.text_widget.tag_add("sel", "1.0", "end")

    def update_image_display(self, *args):
        if not self.window.winfo_exists():
            return
            
        area_name = self.latest_area_name_var.get()
        if self.show_image_var.get() and area_name in self.latest_images:
            image = self.latest_images[area_name]
            
            # Clean up previous photo if it exists
            if hasattr(self, 'photo'):
                del self.photo
            
            # Scale the image according to the selected percentage
            scale_factor = int(self.scale_var.get()) / 100
            if scale_factor != 1:
                new_width = int(image.width * scale_factor)
                new_height = int(image.height * scale_factor)
                image = image.resize((new_width, new_height), Image.Resampling.LANCZOS)
                
            # Calculate new window height based on scaled image height
            window_height = image.height + 300  # Add space for controls and log
            window_height = max(500, window_height)
            
            # Get current window position and width
            window_x = self.window.winfo_x()
            window_y = self.window.winfo_y()
            window_width = self.window.winfo_width()
            
            # Update window geometry
            self.window.geometry(f"{window_width}x{window_height}+{window_x}+{window_y}")
            
            self.photo = ImageTk.PhotoImage(image)
            if self.image_label.winfo_exists():
                self.image_label.config(image=self.photo)
        else:
            if self.image_label.winfo_exists():
                self.image_label.config(image='')
            if hasattr(self, 'photo'):
                del self.photo

    def update_console(self):
        if not hasattr(self, 'text_widget') or not self.text_widget.winfo_exists():
            return
            
        self.text_widget.config(state=tk.NORMAL)
        
        # Get all text and split into lines
        text = self.log_buffer.getvalue()
        lines = text.splitlines()
        
        # Keep only the last MAX_LINES
        if len(lines) > self.MAX_LINES:
            # Join the last MAX_LINES with newlines
            text = '\n'.join(lines[-self.MAX_LINES:]) + '\n'
            # Update the buffer with truncated text
            self.log_buffer.truncate(0)
            self.log_buffer.seek(0)
            self.log_buffer.write(text)
        
        # Update the text widget
        self.text_widget.delete(1.0, tk.END)
        self.text_widget.insert(tk.END, text)
        self.text_widget.config(state=tk.DISABLED)
        self.text_widget.see(tk.END)

    def write(self, message):
        """Write to the console window if it exists"""
        if not self.window.winfo_exists():
            return
            
        self.log_buffer.write(message)  # Write to the buffer
        self.update_console()  # Update the console window with line limit
        if self.show_image_var.get():  # Update image if checkbox is checked
            self.update_image_display()

    def flush(self):
        pass

    def save_log(self):
        # Get the current date and time
        current_time = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        # Get the name of the save file
        save_file_name = self.layout_file_var.get().split('/')[-1].split('.')[0]
        # Suggest a file name
        suggested_name = f"Log_{save_file_name}_{current_time}.txt"
        file_path = filedialog.asksaveasfilename(defaultextension=".txt", initialfile=suggested_name, filetypes=[("Text files", "*.txt")])
        if file_path:
            with open(file_path, 'w') as f:
                f.write(self.log_buffer.getvalue())
            print(f"Log saved to {file_path}\n--------------------------")
     
            
    def save_image(self):
        """Save the currently displayed image"""
        if not self.window.winfo_exists():
            return
            
        area_name = self.latest_area_name_var.get()
        latest_image = self.latest_images.get(area_name)  # Access the image for the current area
        if not isinstance(latest_image, Image.Image):
            messagebox.showerror("Error", "No image to save.")
            return

        current_time = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        suggested_name = f"{area_name}_{current_time}.png"
        file_path = filedialog.asksaveasfilename(
            defaultextension=".png",
            initialfile=suggested_name,
            filetypes=[("PNG files", "*.png")]
        )
        if file_path:
            latest_image.save(file_path, "PNG")
            print(f"Image saved to {file_path}\n--------------------------")

    def clear_console(self):
        """Clear the console text widget and log buffer"""
        if not self.window.winfo_exists():
            return
            
        # Clear the text widget
        self.text_widget.config(state=tk.NORMAL)
        self.text_widget.delete(1.0, tk.END)
        self.text_widget.config(state=tk.DISABLED)
        
        # Clear the log buffer
        self.log_buffer.seek(0)
        self.log_buffer.truncate(0)
        
        # Add a confirmation message
        print("Console cleared\n--------------------------")

class ImageProcessingWindow:
    def __init__(self, root, area_name, latest_images, settings, game_text_reader):
        self.window = tk.Toplevel(root)
        self.window.title(f"Image Processing for: {area_name}")
        
        # Set the window icon
        try:
            icon_path = os.path.join(os.path.dirname(__file__), 'Assets', 'icon.ico')
            if os.path.exists(icon_path):
                self.window.iconbitmap(icon_path)
        except Exception as e:
            print(f"Error setting image processing window icon: {e}")
        
        self.area_name = area_name
        self.latest_images = latest_images
        self.settings = settings
        self.game_text_reader = game_text_reader
        
        # Set up protocol to re-enable hotkeys when window closes
        self.window.protocol("WM_DELETE_WINDOW", self.on_close)

        # Check if there is an image for the area
        if area_name not in latest_images:
            messagebox.showerror("Error", "No image to process, generate an image by pressing the hotkey.")
            self.window.destroy()
            return

        self.image = latest_images[area_name]
        self.processed_image = self.image.copy()

        # Add note about hotkeys being disabled
        hotkey_note = ttk.Label(self.window, text="Note: Hotkeys (including controller hotkeys) are disabled while this window is open.", 
                               font=("Helvetica", 10, "bold"), foreground='#666666')
        hotkey_note.grid(row=0, column=0, columnspan=5, padx=10, pady=(10, 5), sticky='w')
        
        # Disable hotkeys when this window opens
        self.game_text_reader.disable_all_hotkeys()
        
        # Create a canvas to display the image
        self.image_frame = ttk.Frame(self.window)
        self.image_frame.grid(row=1, column=0, columnspan=5, padx=10, pady=5)
        self.canvas = tk.Canvas(self.image_frame, width=self.image.width, height=self.image.height)
        self.canvas.pack()

        # Display the image on the canvas
        self.photo_image = ImageTk.PhotoImage(self.image)
        self.image_on_canvas = self.canvas.create_image(0, 0, anchor=tk.NW, image=self.photo_image)
        
        # Add a label under the image with larger text - centered
        info_text = f"Showing previous image captured in area: {area_name}\n\nProcessing applies to unprocessed images; results may differ if the preview is already processed."
        info_label = ttk.Label(self.image_frame, text=info_text, font=("Helvetica", 12), justify='center')
        info_label.pack(pady=(10, 0), fill='x')

        # Create a frame for bottom controls
        control_frame = ttk.Frame(self.window)
        control_frame.grid(row=2, column=0, columnspan=5, pady=10)

        # Add scale dropdown
        scale_frame = ttk.Frame(control_frame)
        scale_frame.pack(side='left', padx=10)
        
        ttk.Label(scale_frame, text="Preview Scale:").pack(side='left')
        self.scale_var = tk.StringVar(value="100")
        scales = [str(i) for i in range(10, 101, 10)]
        scale_menu = tk.OptionMenu(scale_frame, self.scale_var, *scales, command=self.update_preview)
        scale_menu.pack(side='left')
        ttk.Label(scale_frame, text="%").pack(side='left')

        # Add buttons
        ttk.Button(control_frame, text="Apply img. processing", command=self.save_settings).pack(side='left', padx=10)
        ttk.Button(control_frame, text="Reset to default", command=self.reset_all).pack(side='left', padx=10)

        # Add sliders for image processing
        self.brightness_var = tk.DoubleVar(value=settings.get('brightness', 1.0))
        self.contrast_var = tk.DoubleVar(value=settings.get('contrast', 1.0))
        self.saturation_var = tk.DoubleVar(value=settings.get('saturation', 1.0))
        self.sharpness_var = tk.DoubleVar(value=settings.get('sharpness', 1.0))
        self.blur_var = tk.DoubleVar(value=settings.get('blur', 0.0))
        self.threshold_var = tk.IntVar(value=settings.get('threshold', 128))
        self.hue_var = tk.DoubleVar(value=settings.get('hue', 0.0))
        self.exposure_var = tk.DoubleVar(value=settings.get('exposure', 1.0))
        self.threshold_enabled_var = tk.BooleanVar(value=settings.get('threshold_enabled', False))

        self.create_slider("Brightness", self.brightness_var, 0.1, 2.0, 1.0, 3, 0)
        self.create_slider("Contrast", self.contrast_var, 0.1, 2.0, 1.0, 3, 1)
        self.create_slider("Saturation", self.saturation_var, 0.1, 2.0, 1.0, 3, 2)
        self.create_slider("Sharpness", self.sharpness_var, 0.1, 2.0, 1.0, 3, 3)
        self.create_slider("Blur", self.blur_var, 0.0, 10.0, 0.0, 3, 4)
        self.create_slider("Threshold", self.threshold_var, 0, 255, 128, 4, 0, self.threshold_enabled_var)
        self.create_slider("Hue", self.hue_var, -1.0, 1.0, 0.0, 4, 1)
        self.create_slider("Exposure", self.exposure_var, 0.1, 2.0, 1.0, 4, 2)

    def create_slider(self, label, variable, from_, to, initial, row, col, enabled_var=None):
        frame = ttk.Frame(self.window)
        frame.grid(row=row, column=col, padx=10, pady=5)

        # Use a label frame for consistent structure
        label_frame = ttk.LabelFrame(frame, text=label)
        label_frame.pack(fill='both', expand=True)
    
        ttk.Label(label_frame, text=label).pack()

        entry_var = tk.StringVar(value=f'{initial:.2f}')
        # Add trace to variable to update entry field
        variable.trace_add('write', lambda *args: entry_var.set(f'{variable.get():.2f}'))

        slider = ttk.Scale(label_frame, from_=from_, to=to, orient='horizontal', variable=variable, command=self.update_image)
        slider.set(initial)
        slider.pack()

        # Create entry with context menu
        entry = ttk.Entry(label_frame, textvariable=entry_var)
        entry.pack()
        
        # Add context menu for copy/paste
        entry_menu = tk.Menu(entry, tearoff=0)
        entry_menu.add_command(label="Cut", command=lambda: entry.event_generate('<<Cut>>'))
        entry_menu.add_command(label="Copy", command=lambda: entry.event_generate('<<Copy>>'))
        entry_menu.add_command(label="Paste", command=lambda: entry.event_generate('<<Paste>>'))
        entry_menu.add_separator()
        entry_menu.add_command(label="Select All", command=lambda: entry.selection_range(0, 'end'))
        
        def show_entry_menu(event):
            entry_menu.post(event.x_root, event.y_root)
            
        entry.bind('<Button-3>', show_entry_menu)
    
        ttk.Button(label_frame, text="Reset", command=lambda: self.reset_slider(slider, entry, initial, variable)).pack()

        # Create checkbox for threshold slider
        if label == "Threshold":
            checkbox_frame = ttk.Frame(label_frame)
            checkbox_frame.pack(anchor='w')
        
            checkbox = ttk.Checkbutton(checkbox_frame, variable=enabled_var, command=self.update_image)
            checkbox.pack(side=tk.LEFT)

            ttk.Label(checkbox_frame, text="Enabled").pack(side=tk.LEFT, padx=(5, 0))

        setattr(self, f"{label.lower()}_slider", frame)
        frame.slider, frame.entry = slider, entry
        
    
    def reset_slider(self, slider, entry, initial, variable):
        slider.set(initial)
        variable.set(initial)
        entry.delete(0, tk.END)
        entry.insert(0, str(round(float(initial), 2)))
        self.update_image()
        

    def reset_all(self):
        self.brightness_var.set(1.0)
        self.contrast_var.set(1.0)
        self.saturation_var.set(1.0)
        self.sharpness_var.set(1.0)
        self.blur_var.set(0.0)
        self.threshold_var.set(128)
        self.hue_var.set(0.0)
        self.exposure_var.set(1.0)
        self.threshold_enabled_var.set(False)
        self.update_image()


    def update_image(self, _=None):
        if self.image:
            # Clean up previous processed image if it exists
            if self.processed_image:
                self.processed_image.close()
            self.processed_image = self.image.copy()

            # Apply brightness
            enhancer = ImageEnhance.Brightness(self.processed_image)
            self.processed_image = enhancer.enhance(self.brightness_var.get())

            # Apply contrast
            enhancer = ImageEnhance.Contrast(self.processed_image)
            self.processed_image = enhancer.enhance(self.contrast_var.get())

            # Apply saturation
            enhancer = ImageEnhance.Color(self.processed_image)
            self.processed_image = enhancer.enhance(self.saturation_var.get())

            # Apply sharpness
            enhancer = ImageEnhance.Sharpness(self.processed_image)
            self.processed_image = enhancer.enhance(self.sharpness_var.get())

            # Apply blur
            if self.blur_var.get() > 0:
                self.processed_image = self.processed_image.filter(ImageFilter.GaussianBlur(self.blur_var.get()))

            # Apply threshold if enabled
            if self.threshold_enabled_var.get():
                self.processed_image = self.processed_image.point(lambda p: p > self.threshold_var.get() and 255)

            # Apply hue (simplified, for demonstration purposes)
            self.processed_image = self.processed_image.convert('HSV')
            channels = list(self.processed_image.split())
            channels[0] = channels[0].point(lambda p: (p + int(self.hue_var.get() * 255)) % 256)
            self.processed_image = Image.merge('HSV', channels).convert('RGB')

            # Apply exposure (simplified, for demonstration purposes)
            enhancer = ImageEnhance.Brightness(self.processed_image)
            self.processed_image = enhancer.enhance(self.exposure_var.get())

            # Clean up previous photo_image if it exists
            if self.photo_image:
                del self.photo_image
            self.photo_image = ImageTk.PhotoImage(self.processed_image)
            self.canvas.itemconfig(self.image_on_canvas, image=self.photo_image)

    def save_settings(self):
        # First, update all settings in the processing_settings dictionary
        self.settings['brightness'] = self.brightness_var.get()
        self.settings['contrast'] = self.contrast_var.get()
        self.settings['saturation'] = self.saturation_var.get()
        self.settings['sharpness'] = self.sharpness_var.get()
        self.settings['blur'] = self.blur_var.get()
        self.settings['hue'] = self.hue_var.get()
        self.settings['exposure'] = self.exposure_var.get()
        if self.threshold_enabled_var.get():
            self.settings['threshold'] = self.threshold_var.get()
        else:
            self.settings['threshold'] = None
        self.settings['threshold_enabled'] = self.threshold_enabled_var.get()

        # Ensure the settings are properly stored in the game_text_reader's processing_settings
        area_name = self.area_name
        self.game_text_reader.processing_settings[area_name] = self.settings.copy()

        # Save Auto Read settings to file immediately
        if area_name == "Auto Read":
            import json
            import os
            import tempfile
            
            # First save the processing settings
            self.game_text_reader.processing_settings[area_name] = self.settings.copy()
            
            # Call the existing save_auto_read_settings function to save all settings
            # This will include hotkey, checkboxes, and other settings
            if hasattr(self.game_text_reader, 'save_auto_read_settings'):
                # Get a reference to the save_auto_read_settings function
                save_func = None
                for area in self.game_text_reader.areas:
                    area_frame2, _, _, area_name_var2, _, _, _ = area
                    if area_name_var2.get() == "Auto Read":
                        # This is a bit of a hack - we're accessing the nested function through the frame's children
                        for child in area[0].winfo_children():
                            if hasattr(child, '_name') and child._name == 'save_auto_read_settings':
                                save_func = child
                                break
                        if save_func:
                            break
                
                if save_func:
                    # Call the save function
                    save_func()
                    
                    # Show feedback in status label if available
                    if hasattr(self.game_text_reader, 'status_label'):
                        self.game_text_reader.status_label.config(text="Auto Read settings saved")
                        if hasattr(self.game_text_reader, '_feedback_timer') and self.game_text_reader._feedback_timer:
                            self.game_text_reader.root.after_cancel(self.game_text_reader._feedback_timer)
                        self.game_text_reader._feedback_timer = self.game_text_reader.root.after(2000, 
                            lambda: self.game_text_reader.status_label.config(text=""))

        # Find and enable the preprocess checkbox for this area
        for area_frame, _, _, area_name_var, preprocess_var, _, _ in self.game_text_reader.areas:
            if area_name_var.get() == area_name:
                preprocess_var.set(True)  # Enable the checkbox
                break

        # Check if this is the Auto Read area or if there's a layout file
        is_auto_read = self.area_name == "Auto Read"
        has_layout_file = bool(self.game_text_reader.layout_file.get())
        
        if not has_layout_file and not is_auto_read:
            # Create custom dialog for non-Auto Read areas without a layout file
            dialog = tk.Toplevel(self.window)
            dialog.title("No Save File")
            dialog.geometry("400x150")
            
            # Set the window icon
            try:
                icon_path = os.path.join(os.path.dirname(__file__), 'Assets', 'icon.ico')
                if os.path.exists(icon_path):
                    dialog.iconbitmap(icon_path)
            except Exception as e:
                print(f"Error setting dialog icon: {e}")
            
            dialog.transient(self.window)  # Make dialog modal
            dialog.grab_set()  # Make dialog modal
            
            # Center the dialog on the screen
            dialog.geometry("+%d+%d" % (
                self.window.winfo_rootx() + self.window.winfo_width()/2 - 200,
                self.window.winfo_rooty() + self.window.winfo_height()/2 - 75))
            
            # Add message
            message = tk.Label(dialog, 
                text="No save file exists. You need to save the layout\nto preserve these settings.\n\nCreate save file now?",
                pady=20)
            message.pack()
            
            # Add buttons frame
            button_frame = tk.Frame(dialog)
            button_frame.pack(pady=10)
            
            # Create Yes button
            def on_yes():
                dialog.destroy()
                self.game_text_reader.save_layout()
                
            # Create No button
            def on_no():
                dialog.destroy()
                return
                
            yes_button = tk.Button(button_frame, text="Yes", command=on_yes, width=10)
            yes_button.pack(side='left', padx=10)
            
            no_button = tk.Button(button_frame, text="No", command=on_no, width=10)
            no_button.pack(side='left', padx=10)
            
            # Center the dialog on the screen
            dialog.update_idletasks()
            width = dialog.winfo_width()
            height = dialog.winfo_height()
            x = (dialog.winfo_screenwidth() // 2) - (width // 2)
            y = (dialog.winfo_screenheight() // 2) - (height // 2)
            dialog.geometry(f'{width}x{height}+{x}+{y}')
            
            # Make the dialog modal
            dialog.transient(self.window)
            dialog.grab_set()
            
            # Wait for dialog to close
            self.window.wait_window(dialog)
            
            # If we get here, the user closed the dialog without clicking a button
            return

        # Store a reference to game_text_reader before destroying window
        game_text_reader = self.game_text_reader

        # --- AUTO SAVE for Auto Read area ---
        if area_name == "Auto Read":
            import tempfile, os, json
            # Try to get the preprocess, voice, and speed settings for Auto Read area
            preprocess = None
            voice = None
            speed = None
            for area_frame, _, _, area_name_var, preprocess_var, voice_var, speed_var in game_text_reader.areas:
                if area_name_var.get() == area_name:
                    preprocess = preprocess_var.get() if hasattr(preprocess_var, 'get') else preprocess_var
                    voice = voice_var.get() if hasattr(voice_var, 'get') else voice_var
                    speed = speed_var.get() if hasattr(speed_var, 'get') else speed_var
                    break
            # Find the hotkey for the Auto Read area
            hotkey = None
            for area_frame2, hotkey_button2, _, area_name_var2, _, _, _ in game_text_reader.areas:
                if area_name_var2.get() == area_name:
                    hotkey = getattr(hotkey_button2, 'hotkey', None)
                    break
            # Save to temp file
            settings = {
                'preprocess': preprocess,
                'voice': voice,
                'speed': speed,
                'brightness': self.brightness_var.get(),
                'contrast': self.contrast_var.get(),
                'saturation': self.saturation_var.get(),
                'sharpness': self.sharpness_var.get(),
                'blur': self.blur_var.get(),
                'hue': self.hue_var.get(),
                'exposure': self.exposure_var.get(),
                'threshold': self.threshold_var.get() if self.threshold_enabled_var.get() else None,
                'threshold_enabled': self.threshold_enabled_var.get(),
                'hotkey': hotkey,
                'stop_read_on_select': getattr(game_text_reader, 'interrupt_on_new_scan_var', tk.BooleanVar(value=True)).get() if hasattr(game_text_reader, 'interrupt_on_new_scan_var') else True,
            }
            temp_path = os.path.join(tempfile.gettempdir(), 'auto_read_settings.json')
            with open(temp_path, 'w') as f:
                json.dump(settings, f)
            # Show status message if available
            if hasattr(game_text_reader, 'status_label'):
                game_text_reader.status_label.config(text="Auto Read area settings saved (auto)")
                if hasattr(game_text_reader, '_feedback_timer') and game_text_reader._feedback_timer:
                    game_text_reader.root.after_cancel(game_text_reader._feedback_timer)
                game_text_reader._feedback_timer = game_text_reader.root.after(2000, lambda: game_text_reader.status_label.config(text=""))
            # Destroy window (if not already destroyed)
            self.window.destroy()
            return

        # For all other areas, continue with manual/dialog save logic
        # Destroy window
        self.window.destroy()
        # Now that everything is properly synchronized, save the layout
        game_text_reader.save_layout()


    def update_preview(self, *args):
        """Update the preview with current settings and scale"""
        # Apply current processing settings
        self.processed_image = preprocess_image(
            self.image,
            brightness=self.brightness_var.get(),
            contrast=self.contrast_var.get(),
            saturation=self.saturation_var.get(),
            sharpness=self.sharpness_var.get(),
            blur=self.blur_var.get(),
            threshold=self.threshold_var.get() if self.threshold_enabled_var.get() else None,
            hue=self.hue_var.get(),
            exposure=self.exposure_var.get()
        )

        # Scale the image according to the selected percentage
        scale_factor = int(self.scale_var.get()) / 100
        if scale_factor != 1:
            new_width = int(self.processed_image.width * scale_factor)
            new_height = int(self.processed_image.height * scale_factor)
            display_image = self.processed_image.resize((new_width, new_height), Image.Resampling.LANCZOS)
        else:
            display_image = self.processed_image

        # Update the canvas size
        self.canvas.config(width=display_image.width, height=display_image.height)
        
        # Update the displayed image
        self.photo_image = ImageTk.PhotoImage(display_image)
        self.canvas.itemconfig(self.image_on_canvas, image=self.photo_image)
    
    def on_close(self):
        """Re-enable hotkeys when the window is closed"""
        self.game_text_reader.restore_all_hotkeys()
        self.window.destroy()


def preprocess_image(image, brightness=1.0, contrast=1.0, saturation=1.0, sharpness=1.0, blur=0.0, threshold=None, hue=0.0, exposure=1.0):
    print("Preprocessing image...")

    # Apply brightness
    enhancer = ImageEnhance.Brightness(image)
    image = enhancer.enhance(brightness)

    # Apply contrast
    enhancer = ImageEnhance.Contrast(image)
    image = enhancer.enhance(contrast)

    # Apply saturation
    enhancer = ImageEnhance.Color(image)
    image = enhancer.enhance(saturation)

    # Apply sharpness
    enhancer = ImageEnhance.Sharpness(image)
    image = enhancer.enhance(sharpness)

    # Apply blur
    if blur > 0:
        image = image.filter(ImageFilter.GaussianBlur(blur))

    # Apply threshold if not None
    if threshold is not None:
        image = image.point(lambda p: p > threshold and 255)

    # Apply hue (simplified, for demonstration purposes)
    image = image.convert('HSV')
    channels = list(image.split())
    channels[0] = channels[0].point(lambda p: (p + int(hue * 255)) % 256)
    image = Image.merge('HSV', channels).convert('RGB')

    # Apply exposure (simplified, for demonstration purposes)
    enhancer = ImageEnhance.Brightness(image)
    image = enhancer.enhance(exposure)

    return image

def extract_changelog_from_code(code):
    """Extracts the CHANGELOG string from the code."""
    match = re.search(r'CHANGELOG\s*=\s*([ru]?)(["\reedom\']{3})(.*?)\2', code, re.DOTALL)
    if match:
        return match.group(3).strip()
    return None

def check_for_update(local_version, force=False):  #for testing the updatewindow. false for release.
    """
    Fetch the remote GameReader.py from GitHub, extract version and changelog, compare to local_version.
    If remote version is newer or force=True, show a popup.
    """
    GITHUB_RAW_URL = "https://raw.githubusercontent.com/MertenNor/GameReader/main/GameReader.py"
    try:
        resp = requests.get(GITHUB_RAW_URL, timeout=5)
        if resp.status_code == 200:
            remote_content = resp.text
            remote_version = extract_version_from_code(remote_content)
            remote_changelog = extract_changelog_from_code(remote_content)
            if force or (remote_version and version_tuple(remote_version) > version_tuple(local_version)):
                # Create a custom popup window
                import tkinter as tk
                from tkinter import ttk
                
                popup = tk.Toplevel()
                popup.title("Update Available")
                popup.geometry("750x350")  # Set initial size
                popup.minsize(400, 150)    # Set minimum size
                
                # Set the window icon
                try:
                    icon_path = os.path.join(os.path.dirname(__file__), 'Assets', 'icon.ico')
                    if os.path.exists(icon_path):
                        popup.iconbitmap(icon_path)
                except Exception as e:
                    print(f"Error setting update popup icon: {e}")
                
                # Make window resizable
                popup.resizable(True, True)
                
                # Configure grid weights
                popup.grid_rowconfigure(0, weight=1)
                popup.grid_columnconfigure(0, weight=1)
                
                # Create main frame with padding
                main_frame = ttk.Frame(popup, padding="20")
                main_frame.grid(row=0, column=0, sticky='nsew')
                main_frame.grid_rowconfigure(1, weight=1)  # Make text area expandable
                main_frame.grid_columnconfigure(0, weight=1)
                
                # Version info
                version_frame = ttk.Frame(main_frame)
                version_frame.grid(row=0, column=0, sticky='w', pady=(0, 15))
                
                ttk.Label(version_frame, text="A new version of GameReader is available!", 
                         font=('Helvetica', 12, 'bold')).grid(row=0, column=0, columnspan=2, sticky='w', pady=(0, 10))
                
                ttk.Label(version_frame, text=f"Current version: {local_version}", font=('Helvetica', 10)).grid(row=1, column=0, sticky='w')
                ttk.Label(version_frame, text=f"Latest version: {remote_version}", font=('Helvetica', 10)).grid(row=2, column=0, sticky='w')
                
                # Changelog section
                ttk.Label(main_frame, text="What's new:", font=('Helvetica', 10, 'bold')).grid(row=1, column=0, sticky='nw', pady=(10, 5))
                
                # Create a frame for the text widget and scrollbar
                text_frame = ttk.Frame(main_frame)
                text_frame.grid(row=2, column=0, sticky='nsew', pady=(0, 15))
                text_frame.grid_rowconfigure(0, weight=1)
                text_frame.grid_columnconfigure(0, weight=1)
                
                # Add text widget with scrollbar
                text = tk.Text(text_frame, wrap=tk.WORD, width=60, height=10, 
                             padx=10, pady=10, relief='flat', bg='#f0f0f0')
                scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=text.yview)
                text.configure(yscrollcommand=scrollbar.set)
                
                text.grid(row=0, column=0, sticky='nsew')
                scrollbar.grid(row=0, column=1, sticky='ns')
                
                # Insert changelog text
                changelog = remote_changelog if remote_changelog else "No changelog available."
                text.insert('1.0', changelog)
                text.config(state='disabled')  # Make text read-only
                
                # Buttons frame
                button_frame = ttk.Frame(main_frame)
                button_frame.grid(row=3, column=0, sticky='e')
                
                def open_github():
                    import webbrowser
                    webbrowser.open('https://github.com/MertenNor/GameReader/releases')
                    popup.destroy()
                
                def close_popup():
                    popup.destroy()
                
                ttk.Button(button_frame, text="Later...", command=close_popup).pack(side='right', padx=5)
                ttk.Button(button_frame, text="Open download page", command=open_github).pack(side='right', padx=5)
                
                # Center the popup on screen
                popup.update_idletasks()
                width = popup.winfo_width()
                height = popup.winfo_height()
                x = (popup.winfo_screenwidth() // 2) - (width // 2)
                y = (popup.winfo_screenheight() // 2) - (height // 2)
                popup.geometry(f'{width}x{height}+{x}+{y}')
                
                # Make popup modal
                popup.transient(tk._default_root or popup)
                popup.grab_set()
                popup.wait_window()
    except Exception as e:
        # Fail silently if no internet or any error
        pass

def version_tuple(v):
    """Convert a version string like '0.6.1' to a tuple of ints: (0,6,1)"""
    return tuple(int(x) for x in v.split('.') if x.isdigit())

def extract_version_from_code(code):
    """Extracts the version string from APP_VERSION = "x.y" in GameReader.py."""
    match = re.search(r'APP_VERSION\s*=\s*"([\d.]+)"', code)
    if match:
        return match.group(1)
    return None

class ControllerHandler:
    """Handles controller input detection for hotkey assignment"""
    
    def __init__(self):
        self.running = False
        self.controller_thread = None
        self.controller_available = CONTROLLER_AVAILABLE
        self.last_button_press = None
        self.button_press_event = threading.Event()
        
    def start_monitoring(self):
        """Start monitoring controller inputs in a separate thread"""
        if not self.controller_available:
            return False
            
        try:
            self.running = True
            self.controller_thread = threading.Thread(target=self._monitor_controller)
            self.controller_thread.daemon = True
            self.controller_thread.start()
            return True
        except Exception as e:
            print(f"Error starting controller monitoring: {e}")
            return False
            
    def stop_monitoring(self):
        """Stop monitoring controller inputs"""
        self.running = False
        if self.controller_thread:
            self.controller_thread.join(timeout=1)
            
    def wait_for_button_press(self, timeout=10):
        """Wait for a controller button press and return the button name"""
        if not self.controller_available:
            return None
            
        self.last_button_press = None
        self.button_press_event.clear()
        
        # Start monitoring if not already running
        if not self.running:
            self.start_monitoring()
            
        # Wait for button press
        if self.button_press_event.wait(timeout):
            return self.last_button_press
        return None
        
    def _monitor_controller(self):
        """Monitor controller events in a loop"""
        while self.running:
            try:
                events = inputs.get_gamepad()
                for event in events:
                    button_name = None
                    
                    # Debug logging for D-Pad troubleshooting (uncomment if needed)
                    # if event.ev_type == 'Absolute' and ('HAT' in event.code or 'X' in event.code or 'Y' in event.code):
                    #     print(f"Debug - D-Pad event: {event.ev_type} {event.code} = {event.state}")
                    
                    # Handle Key events (regular buttons)
                    if event.ev_type == 'Key' and event.state == 1:  # Button press
                        button_name = self._get_button_name(event.code)
                    
                    # Handle Absolute events (D-Pad only - no analog sticks)
                    elif event.ev_type == 'Absolute':
                        button_name = self._get_absolute_button_name(event.code, event.state)
                    
                    # If we got a button name, process it
                    if button_name:
                        # print(f"Controller button detected: {button_name} (from {event.ev_type} {event.code})")  # Debug output disabled
                        self.last_button_press = button_name
                        self.button_press_event.set()
                        
                        # Trigger any active controller hotkeys
                        self._trigger_controller_hotkeys(button_name)
                        
                        # Notify the main class about the button press
                        if hasattr(self, 'game_reader') and self.game_reader:
                            self.game_reader._check_controller_hotkeys(button_name)
                        break
            except inputs.UnpluggedError:
                # Controller disconnected
                time.sleep(1)
            except Exception as e:
                print(f"Controller error: {e}")
                time.sleep(1)
                
    def _get_button_name(self, code):
        """Convert controller button code to readable name"""
        # Use generic button numbers that work for all controller types
        button_mapping = {
            # Face buttons - these are universal across controllers
            'BTN_SOUTH': 'Btn 1',      # A on Xbox, Cross on PlayStation, A on Nintendo
            'BTN_EAST': 'Btn 2',       # B on Xbox, Circle on PlayStation, B on Nintendo  
            'BTN_NORTH': 'Btn 3',      # Y on Xbox, Triangle on PlayStation, Y on Nintendo
            'BTN_WEST': 'Btn 4',       # X on Xbox, Square on PlayStation, X on Nintendo
            
            # Shoulder buttons
            'BTN_TL': 'Btn 5',         # LB on Xbox, L1 on PlayStation, L on Nintendo
            'BTN_TR': 'Btn 6',         # RB on Xbox, R1 on PlayStation, R on Nintendo
            
            # Stick buttons
            'BTN_THUMBL': 'Btn 7',     # LS on Xbox, L3 on PlayStation, Left Stick on Nintendo
            'BTN_THUMBR': 'Btn 8',     # RS on Xbox, R3 on PlayStation, Right Stick on Nintendo
            
            # Menu buttons
            'BTN_START': 'Btn 9',      # START on Xbox, OPTIONS on PlayStation, + on Nintendo
            'BTN_SELECT': 'Btn 10',    # SELECT on Xbox, SHARE on PlayStation, - on Nintendo
            'BTN_MODE': 'Btn 11',      # HOME on Xbox, PS Button on PlayStation, HOME on Nintendo
            
            # D-Pad buttons (digital only - no analog stick)
            'BTN_DPAD_UP': 'DPAD_UP',
            'BTN_DPAD_DOWN': 'DPAD_DOWN',
            'BTN_DPAD_LEFT': 'DPAD_LEFT',
            'BTN_DPAD_RIGHT': 'DPAD_RIGHT',
            
            # Additional D-Pad codes that some controllers use
            'BTN_DPAD_UP_ALT': 'DPAD_UP',
            'BTN_DPAD_DOWN_ALT': 'DPAD_DOWN',
            'BTN_DPAD_LEFT_ALT': 'DPAD_LEFT',
            'BTN_DPAD_RIGHT_ALT': 'DPAD_RIGHT',
            
            # Some controllers use different naming conventions
            'BTN_HAT_UP': 'DPAD_UP',
            'BTN_HAT_DOWN': 'DPAD_DOWN',
            'BTN_HAT_LEFT': 'DPAD_LEFT',
            'BTN_HAT_RIGHT': 'DPAD_RIGHT'
        }
        return button_mapping.get(code, f"Btn_{code}")
    
    def _get_absolute_button_name(self, code, state):
        """Convert absolute controller events (D-Pad only) to button names"""
        # Only handle true D-Pad events (HAT codes) - ignore analog stick movements
        # D-Pad events - many controllers use ABS_HAT0X and ABS_HAT0Y
        if code == 'ABS_HAT0X':
            if state == -1:  # Left
                return 'DPAD_LEFT'
            elif state == 1:  # Right
                return 'DPAD_RIGHT'
        elif code == 'ABS_HAT0Y':
            if state == -1:  # Up
                return 'DPAD_UP'
            elif state == 1:  # Down
                return 'DPAD_DOWN'
        
        # Alternative D-Pad codes that some controllers use
        elif code == 'ABS_HAT1X':
            if state == -1:  # Left
                return 'DPAD_LEFT'
            elif state == 1:  # Right
                return 'DPAD_RIGHT'
        elif code == 'ABS_HAT1Y':
            if state == -1:  # Up
                return 'DPAD_UP'
            elif state == 1:  # Down
                return 'DPAD_DOWN'
        
        # Return None for analog stick movements and other absolute events
        # This prevents accidental hotkey triggers from stick movements
        return None
    
    def _trigger_controller_hotkeys(self, button_name):
        """Trigger any active controller hotkeys for the given button"""
        try:
            # This method will be called by the main class to check for hotkeys
            # The actual hotkey checking is done in the main class
            pass
        except Exception as e:
            print(f"Error triggering controller hotkeys: {e}")

class GameTextReader:
    def __init__(self, root):
        self.root = root
        self.root.title(f"Game Reader v{APP_VERSION}")
        # --- Update check on startup ---
        local_version = APP_VERSION
        FORCE_UPDATE_CHECK = False  # Set to True to force update popup, False for normal behavior
        threading.Thread(target=lambda: check_for_update(local_version, force=FORCE_UPDATE_CHECK), daemon=True).start()
        # --- End update check ---
        
        # Don't set initial geometry here - let it be calculated after GUI setup
        # self.root.geometry("1115x260")  # Initial window size (height reduced for less vertical tallness)
        
        self.layout_file = tk.StringVar()
        self.latest_images = {}  # Use a dictionary to store images for each area
        self.latest_area_name = tk.StringVar()  # Ensure this is defined
        self.areas = []
        self.stop_hotkey = None  # Variable to store the STOP hotkey
        # Initialize text-to-speech engine with error handling
        self.engine = None
        self.engine_lock = threading.Lock()  # Lock for the text-to-speech engine
        try:
            self.engine = pyttsx3.init()
            # Test if engine is working by trying to get a property
            _ = self.engine.getProperty('rate')
        except Exception as e:
            print(f"Warning: Could not initialize text-to-speech engine: {e}")
            print("Text-to-speech functionality will be disabled.")
            self.engine = None
        self.bad_word_list = tk.StringVar()  # StringVar for the bad word list
        self.hotkeys = set()  # Track registered hotkeys
        self.is_speaking = False  # Flag to track if the engine is speaking
        self.processing_settings = {}  # Dictionary to store processing settings for each area
        self.processing_settings_widgets = {}  # Dictionary to store processing settings widgets for each area
        self.volume = tk.StringVar(value="100")  # Default volume 100%
        self.speaker = win32com.client.Dispatch("SAPI.SpVoice")
        self.speaker.Volume = int(self.volume.get())  # Set initial volume
        self.is_speaking = False
        
        # Wake up Online SAPI5 voices on program start to prevent first-call delays
        self._wake_up_online_voices()

        # Initialize all checkbox variables
        self.ignore_usernames_var = tk.BooleanVar(value=False)
        self.ignore_previous_var = tk.BooleanVar(value=False)
        self.ignore_gibberish_var = tk.BooleanVar(value=False)
        self.pause_at_punctuation_var = tk.BooleanVar(value=False)
        self.fullscreen_mode_var = tk.BooleanVar(value=False)

        # Hotkey management
        self.hotkey_scancodes = {}  # Dictionary to store scan codes for hotkeys
        self.setting_hotkey = False  # Flag to track if we're in hotkey setting mode
        self.unhook_timer = None  # Timer for hotkey unhooking
        self.keyboard_hooks = []  # List to track keyboard hooks
        self.mouse_hooks = []  # List to track mouse hooks
        self.info_window_open = False  # Flag to track if info window is open
        
        # Controller support
        self.controller_handler = ControllerHandler()
        self.controller_handler.game_reader = self  # Set reference to main class
        
        # Setup Tesseract command path if it's not in your PATH
        pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

        self.numpad_scan_codes = {
            82: '0',     # Numpad 0
            79: '1',     # Numpad 1
            80: '2',     # Numpad 2
            81: '3',     # Numpad 3
            75: '4',     # Numpad 4
            76: '5',     # Numpad 5
            77: '6',     # Numpad 6
            71: '7',     # Numpad 7
            72: '8',     # Numpad 8
            73: '9',     # Numpad 9
            55: 'multiply',  # Numpad * (changed from '*' to 'multiply')
            78: 'add',       # Numpad + (changed from '+' to 'add')
            74: 'subtract',  # Numpad - (changed from '-' to 'subtract')
            83: '.',         # Numpad .
            53: 'divide',    # Numpad / (changed from '/' to 'divide')
            28: 'enter'      # Numpad Enter
        }

        # VK codes for numpad keys, used for fullscreen fallback polling
        # Reference: https://learn.microsoft.com/windows/win32/inputdev/virtual-key-codes
        self.numpad_vk_codes = {
            '0': 0x60,  # VK_NUMPAD0
            '1': 0x61,  # VK_NUMPAD1
            '2': 0x62,  # VK_NUMPAD2
            '3': 0x63,  # VK_NUMPAD3
            '4': 0x64,  # VK_NUMPAD4
            '5': 0x65,  # VK_NUMPAD5
            '6': 0x66,  # VK_NUMPAD6
            '7': 0x67,  # VK_NUMPAD7
            '8': 0x68,  # VK_NUMPAD8
            '9': 0x69,  # VK_NUMPAD9
            '*': 0x6A,  # VK_MULTIPLY
            '+': 0x6B,  # VK_ADD
            '-': 0x6D,  # VK_SUBTRACT
            '.': 0x6E,  # VK_DECIMAL
            '/': 0x6F,  # VK_DIVIDE
            'enter': 0x0D  # VK_RETURN (cannot distinguish main vs numpad)
        }

        self.text_histories = {}  # Dictionary to store text history for each area
        self.ignore_previous_var = tk.BooleanVar(value=False)  # Variable for the checkbox
        self.ignore_gibberish_var = tk.BooleanVar(value=False)  # Variable for the gibberish checkbox
        self.pause_at_punctuation_var = tk.BooleanVar(value=False)  # Variable for punctuation pauses
        self.fullscreen_mode_var = tk.BooleanVar(value=False)  # Variable for fullscreen mode
        # Add variable for better measurement unit detection
        self.better_unit_detection_var = tk.BooleanVar(value=False)
        # Add variable for read game units
        self.read_game_units_var = tk.BooleanVar(value=False)
        
        # Add variable for interrupt on new scan
        self.interrupt_on_new_scan_var = tk.BooleanVar(value=True)

        # UWP TTS concurrency control
        self._uwp_lock = threading.Lock()
        self._uwp_player = None
        self._uwp_queue = queue.Queue()
        self._uwp_thread_stop = threading.Event()
        self._uwp_interrupt = threading.Event()
        self._uwp_thread = threading.Thread(target=self._uwp_worker, daemon=True)
        self._uwp_thread.start()

        # Numpad fallback polling is always enabled when a numpad hotkey is set

        # Load game units from JSON file
        self.game_units = self.load_game_units()

        self.setup_gui()
        # Get available voices using SAPI instead of pyttsx3
        try:
            # Research-based solution: Force Windows to load ALL installed voices
            all_voices = []
            # Disable heavy/side-effect discovery steps by default (no service restarts or PowerShell)
            enable_heavy_discovery = False
            
            # Method 1: Enumerate SAPI voices (quiet)
            try:
                # Create a new SAPI object specifically for enumeration
                enum_voice = win32com.client.Dispatch("SAPI.SpVoice")
                
                voices = enum_voice.GetVoices()
                for i, voice in enumerate(voices):
                    try:
                        all_voices.append(voice)
                    except Exception:
                        pass
                        
            except Exception as e1:
                print(f"Method 1 failed: {e1}")
            

            
            # Method 2: Try to force Windows to register all voices (disabled by default to avoid console popups)
            if enable_heavy_discovery:
                try:
                    import subprocess
                    try:
                        # Hide any console window
                        si = subprocess.STARTUPINFO()
                        si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                        creationflags = subprocess.CREATE_NO_WINDOW
                        subprocess.run(['net', 'stop', 'audiosrv'], capture_output=True, startupinfo=si, creationflags=creationflags)
                        subprocess.run(['net', 'start', 'audiosrv'], capture_output=True, startupinfo=si, creationflags=creationflags)
                    except Exception:
                        pass
                    # Re-enumerate voices
                    try:
                        enum_voice2 = win32com.client.Dispatch("SAPI.SpVoice")
                        voices2 = enum_voice2.GetVoices()
                        for i, voice in enumerate(voices2):
                            try:
                                if not any(v.GetDescription() == voice.GetDescription() for v in all_voices):
                                    all_voices.append(voice)
                            except Exception:
                                pass
                    except Exception:
                        pass
                except Exception:
                    pass
            
            # Method 3: Check Speech_OneCore registry locations (quiet)
            
            # Method 3.5: Try to force Windows to register OneCore voices (quiet)
            try:
                # Try to create a voice object and enumerate with different filters
                force_voice = win32com.client.Dispatch("SAPI.SpVoice")
                # Try to get voices with different enumeration methods
                try:
                    # Try to enumerate with a filter that might include OneCore voices
                    voices_force = force_voice.GetVoices("", "")
                    for i in range(voices_force.Count):
                        try:
                            voice = voices_force.Item(i)
                            if not any(v.GetDescription() == voice.GetDescription() for v in all_voices):
                                all_voices.append(voice)
                        except Exception:
                            pass
                except Exception as force_e:
                    print(f"  Force enumeration failed: {force_e}")
            except Exception as e3_5:
                print(f"Method 3.5 failed: {e3_5}")
            
            # Method 4: Try to force Windows to load OneCore voices by accessing them directly
            # Method 4: Try to force Windows to load OneCore voices directly (quiet mode)
            try:
                # Ensure OneCore registry locations are defined
                try:
                    import winreg
                    onecore_locations = [
                        r"SOFTWARE\\Microsoft\\Speech_OneCore\\Voices\\Tokens",
                        r"SOFTWARE\\WOW6432Node\\Microsoft\\Speech_OneCore\\Voices\\Tokens"
                    ]
                except Exception:
                    onecore_locations = []
                
                # Try to create voice objects for each OneCore token we found
                onecore_tokens = []
                for location in onecore_locations:
                    try:
                        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, location)
                        i = 0
                        while True:
                            try:
                                voice_token = winreg.EnumKey(key, i)
                                onecore_tokens.append(voice_token)
                                i += 1
                            except WindowsError:
                                break
                        winreg.CloseKey(key)
                    except Exception:
                        pass
                
                for token in onecore_tokens:
                    try:
                        # Try to create a voice object using the token directly
                        voice_obj = win32com.client.Dispatch("SAPI.SpVoice")
                        
                        # Try to set the voice using the token
                        try:
                            # Try to create voice using token as a filter
                            voices_enum = voice_obj.GetVoices()
                            for j in range(voices_enum.Count):
                                voice = voices_enum.Item(j)
                                desc = voice.GetDescription()
                                # Check if this voice matches our token
                                if (token in desc or 
                                    token.replace('MSTTS_V110_', '').replace('M', '') in desc or
                                    any(part in desc for part in token.split('_')[2:4])):
                                    if not any(v.GetDescription() == desc for v in all_voices):
                                        all_voices.append(voice)
                                    break
                        except Exception as enum_e:
                            pass
                        
                        # Try alternative method: Create voice using token category
                        try:
                            token_cat = win32com.client.Dispatch("SAPI.SpObjectTokenCategory")
                            token_cat.SetId("HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Speech_OneCore\\Voices", False)
                            tokens_enum = token_cat.EnumTokens()
                            for k in range(tokens_enum.Count):
                                token_obj = tokens_enum.Item(k)
                                if token in token_obj.GetId():
                                    # Try to create voice from this token
                                    try:
                                        new_voice = win32com.client.Dispatch("SAPI.SpVoice")
                                        new_voice.Voice = token_obj
                                        desc = new_voice.Voice.GetDescription()
                                        if not any(v.GetDescription() == desc for v in all_voices):
                                            all_voices.append(new_voice.Voice)
                                    except Exception as create_e:
                                        pass
                                    break
                        except Exception as token_e:
                            # Silently ignore token category access errors to reduce console noise
                            pass
                            
                    except Exception as token_voice_e:
                        print(f"    -> Error processing token {token}: {token_voice_e}")
                        
            except Exception as e4:
                print(f"Method 4 failed: {e4}")
            
            # Method 5: Try to force Windows to register OneCore voices by accessing Windows Speech settings
            # Method 5: Skipped opening Windows Speech settings to avoid UI interruptions
            
            # Method 6: Try to create working voice objects for OneCore voices
            # Method 6: Create working voice objects for OneCore voices (quiet mode)
            try:
                # For each OneCore token, try to create a working voice object
                for token in onecore_tokens:
                    try:
                        # Quiet: create working voice entries for UI selection
                        
                        # Try to create a voice object that can actually be used
                        class WorkingOneCoreVoice:
                            def __init__(self, token):
                                self._token = token
                                # Convert token to readable name
                                parts = token.split('_')
                                if len(parts) >= 4:
                                    lang = parts[2]
                                    name = parts[3].replace('M', '')
                                    self._desc = f"Microsoft {name} - {lang}"
                                else:
                                    self._desc = token
                                # Store the token for later use
                                self._voice_token = token
                            
                            def GetDescription(self):
                                return self._desc
                            
                            def GetId(self):
                                return self._token
                            
                            def GetToken(self):
                                return self._voice_token
                        
                        working_voice = WorkingOneCoreVoice(token)
                        if not any(v.GetDescription() == working_voice.GetDescription() for v in all_voices):
                            all_voices.append(working_voice)
                            # Quiet log
                        
                    except Exception as working_e:
                        print(f"    -> Error creating working voice for {token}: {working_e}")
                        
            except Exception as e6:
                print(f"Method 6 failed: {e6}")
            
            # Method 7: Try to force Windows to register OneCore voices by using Windows Speech API directly
            # Method 7: PowerShell forcing (disabled by default to avoid flashing a console window)
            if enable_heavy_discovery:
                try:
                    import subprocess
                    try:
                        ps_command = (
                            "Add-Type -AssemblyName System.Speech;"
                            "$synthesizer = New-Object System.Speech.Synthesis.SpeechSynthesizer;"
                            "$voices = $synthesizer.GetInstalledVoices();"
                            "$voices | ForEach-Object { $_.VoiceInfo.Name }"
                        )
                        si = subprocess.STARTUPINFO()
                        si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                        creationflags = subprocess.CREATE_NO_WINDOW
                        result = subprocess.run(['powershell', '-NoProfile', '-Command', ps_command],
                                                capture_output=True, text=True, startupinfo=si, creationflags=creationflags)
                        if result.returncode == 0:
                            try:
                                enum_voice4 = win32com.client.Dispatch("SAPI.SpVoice")
                                voices4 = enum_voice4.GetVoices()
                                for i in range(voices4.Count):
                                    try:
                                        voice = voices4.Item(i)
                                        if not any(v.GetDescription() == voice.GetDescription() for v in all_voices):
                                            all_voices.append(voice)
                                    except Exception:
                                        pass
                            except Exception:
                                pass
                    except Exception:
                        pass
                except Exception:
                    pass
            try:
                import winreg
                
                # Check Speech_OneCore registry locations
                onecore_locations = [
                    r"SOFTWARE\Microsoft\Speech_OneCore\Voices\Tokens",
                    r"SOFTWARE\WOW6432Node\Microsoft\Speech_OneCore\Voices\Tokens"
                ]
                
                for location in onecore_locations:
                    try:
                        # Quiet: skip registry location logging
                        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, location)
                        i = 0
                        while True:
                            try:
                                voice_token = winreg.EnumKey(key, i)
                                # Quiet: skip per-token logging
                                
                                # Try to create voice object from OneCore token
                                try:
                                    voice_obj = win32com.client.Dispatch("SAPI.SpVoice")
                                    # Try to enumerate and find this specific voice
                                    voices = voice_obj.GetVoices()
                                    for j in range(voices.Count):
                                        voice = voices.Item(j)
                                        desc = voice.GetDescription()
                                        # Try different matching strategies
                                        if (voice_token in desc or 
                                            voice_token.replace('MSTTS_V110_', '').replace('M', '') in desc or
                                            any(part in desc for part in voice_token.split('_')[2:4])):
                                            print(f"      -> Matched: {desc}")
                                            if not any(v.GetDescription() == desc for v in all_voices):
                                                all_voices.append(voice)
                                            break
                                    else:
                                        # If no match found, try to create a real SAPI voice object
                                        # Quiet
                                        try:
                                            # Try to create voice object directly using the token
                                            real_voice = win32com.client.Dispatch("SAPI.SpVoice")
                                            # Try to set the voice by token
                                            voices_enum = real_voice.GetVoices()
                                            for k in range(voices_enum.Count):
                                                voice_obj = voices_enum.Item(k)
                                                if voice_token in voice_obj.GetDescription():
                                                    print(f"        -> Found real voice: {voice_obj.GetDescription()}")
                                                    if not any(v.GetDescription() == voice_obj.GetDescription() for v in all_voices):
                                                        all_voices.append(voice_obj)
                                                    break
                                            else:
                                                # Try alternative method: Create voice object using token directly
                                                # Quiet
                                                try:
                                                    # Try to create voice using the token as a filter
                                                    voice_enum = win32com.client.Dispatch("SAPI.SpObjectTokenCategory")
                                                    voice_enum.SetId("HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Speech\\Voices", False)
                                                    tokens = voice_enum.EnumTokens()
                                                    for token_idx in range(tokens.Count):
                                                        token = tokens.Item(token_idx)
                                                        if voice_token in token.GetId():
                                                        # Quiet
                                                            # Try to create voice from this token
                                                            try:
                                                                voice_obj = win32com.client.Dispatch("SAPI.SpVoice")
                                                                voice_obj.Voice = token
                                                                desc = voice_obj.Voice.GetDescription()
                                                                # Quiet
                                                                if not any(v.GetDescription() == desc for v in all_voices):
                                                                    all_voices.append(voice_obj.Voice)
                                                                break
                                                            except Exception as token_e:
                                                                # Quiet
                                                                pass
                                                except Exception as alt_e:
                                                    # Quiet
                                                    pass
                                                
                                                # If still no match, create a mock voice
                                                if not any(v.GetDescription().startswith(f"Microsoft {voice_token.split('_')[3].replace('M', '')}") for v in all_voices):
                                                    print(f"        -> Creating mock voice for: {voice_token}")
                                                    class MockOneCoreVoice:
                                                        def __init__(self, token):
                                                            self._token = token
                                                            # Convert token to readable name
                                                            parts = token.split('_')
                                                            if len(parts) >= 4:
                                                                lang = parts[2]
                                                                name = parts[3].replace('M', '')
                                                                self._desc = f"Microsoft {name} - {lang}"
                                                            else:
                                                                self._desc = token
                                                        def GetDescription(self):
                                                            return self._desc
                                                    mock_voice = MockOneCoreVoice(voice_token)
                                                    if not any(v.GetDescription() == mock_voice.GetDescription() for v in all_voices):
                                                        all_voices.append(mock_voice)
                                        except Exception as real_voice_e:
                                            # Quiet
                                            # Fall back to mock voice
                                            class MockOneCoreVoice:
                                                def __init__(self, token):
                                                    self._token = token
                                                    # Convert token to readable name
                                                    parts = token.split('_')
                                                    if len(parts) >= 4:
                                                        lang = parts[2]
                                                        name = parts[3].replace('M', '')
                                                        self._desc = f"Microsoft {name} - {lang}"
                                                    else:
                                                        self._desc = token
                                                def GetDescription(self):
                                                    return self._desc
                                            mock_voice = MockOneCoreVoice(voice_token)
                                            if not any(v.GetDescription() == mock_voice.GetDescription() for v in all_voices):
                                                all_voices.append(mock_voice)
                                except Exception as voice_e:
                                    print(f"      -> Could not create voice: {voice_e}")
                                
                                i += 1
                            except WindowsError:
                                break
                        winreg.CloseKey(key)
                    except Exception as loc_e:
                        print(f"    Could not access {location}: {loc_e}")
                        
            except ImportError:
                print("winreg not available")
            
            # Use the combined list
            self.voices = all_voices
            print(f"\nFinal combined voice list: {len(self.voices)} voices")
                
        except Exception as e:
            print(f"Warning: Could not get SAPI voices: {e}")
            self.voices = []
        
        self.stop_keyboard_hook = None
        self.stop_mouse_hook = None
        self.setting_hotkey_mouse_hook = None
        self.unhook_timer = None
        
        # Add this line to handle window closing
        root.protocol("WM_DELETE_WINDOW", lambda: (self.cleanup(), root.destroy()))
        
        # Track if there are unsaved changes
        self._has_unsaved_changes = False
        
        # Enable drag and drop using TkinterDnD2 when available
        if hasattr(root, 'drop_target_register'):
            try:
                root.drop_target_register(DND_FILES)
                root.dnd_bind('<<Drop>>', self.on_drop)
                root.dnd_bind('<<DropEnter>>', lambda e: 'break')
                root.dnd_bind('<<DropPosition>>', lambda e: 'break')
            except Exception as dnd_error:
                print(f"Warning: drag-and-drop could not be initialized: {dnd_error}")
        else:
            print("Info: TkinterDnD not available; drag-and-drop is disabled.")

        # Controller support disabled - pygame removed to reduce Windows security flags
        self.controller = None

    def speak_text(self, text):
        """Speak text using win32com.client (SAPI.SpVoice)."""
        # Check if TTS is available; if not, try UWP fallback
        if not hasattr(self, 'speaker') or self.speaker is None:
            if _ensure_uwp_available():
                try:
                    loop = asyncio.new_event_loop()
                    asyncio.set_event_loop(loop)
                    loop.run_until_complete(self._speak_with_uwp(text))
                    loop.close()
                    return
                except Exception as _e:
                    pass
            print("Warning: Text-to-speech is not available. Please check your system's speech settings.")
            return
            
        # Always check and stop speech if interrupt is enabled
        if hasattr(self, 'interrupt_on_new_scan_var') and self.interrupt_on_new_scan_var.get():
            # Stop SAPI and also stop any ongoing UWP playback to prevent crashes when switching voices
            self.stop_speaking()
            if hasattr(self, '_uwp_lock'):
                try:
                    with self._uwp_lock:
                        if hasattr(self, '_uwp_player') and self._uwp_player is not None:
                            try:
                                self._uwp_player.pause()
                            except Exception:
                                pass
                            self._uwp_player = None
                except Exception:
                    pass
        elif self.is_speaking:
            print("Already speaking. Please stop the current speech first.")
            return
            
        self.is_speaking = True
        try:
            # Use a lower priority for speaking
            self.speaker.Speak(text, 1)  # 1 is SVSFlagsAsync
            print("Speech started.\n--------------------------")
        except Exception as e:
            print(f"Error during speech: {e}")
            self.is_speaking = False
            # Try UWP fallback if available
            if _ensure_uwp_available():
                try:
                    # Ensure previous UWP playback is stopped before starting new
                    if hasattr(self, '_uwp_lock'):
                        try:
                            with self._uwp_lock:
                                if hasattr(self, '_uwp_player') and self._uwp_player is not None:
                                    try:
                                        self._uwp_player.pause()
                                    except Exception:
                                        pass
                                    self._uwp_player = None
                        except Exception:
                            pass
                    loop = asyncio.new_event_loop()
                    asyncio.set_event_loop(loop)
                    loop.run_until_complete(self._speak_with_uwp(text))
                    loop.close()
                    return
                except Exception as _e:
                    pass

    def stop_speaking(self):
        """Stop the ongoing speech immediately."""
        # Stop both SAPI and UWP playback
        try:
            if hasattr(self, 'speaker') and self.speaker:
                try:
                    self.speaker.Speak("", 2)
                except Exception:
                    pass
            self.is_speaking = False
            # Signal UWP worker to stop current playback
            try:
                self._uwp_queue.put_nowait(("STOP", None, None))
            except Exception:
                pass
            # Reinitialize SAPI speaker
            try:
                self.speaker = win32com.client.Dispatch("SAPI.SpVoice")
                self.speaker.Volume = int(self.volume.get())
            except Exception:
                self.speaker = None
            print("Speech stopped.\n--------------------------")
        except Exception as e:
            print(f"Error stopping speech: {e}")
            self.is_speaking = False

    def controller_listener(self):
        """Controller support disabled - pygame removed to reduce Windows security flags"""
        pass
            


    def get_controller_button_name(self, button_number):
        """Controller support disabled - pygame removed to reduce Windows security flags"""
        return f"btn:{button_number}"
    
    def get_controller_hat_name(self, hat_number, hat_value):
        """Controller support disabled - pygame removed to reduce Windows security flags"""
        return f"hat{hat_number}_{hat_value[0]}_{hat_value[1]}"

    def detect_controllers(self):
        """Controller support disabled - pygame removed to reduce Windows security flags"""
        return []



    async def _speak_with_uwp(self, text: str, preferred_desc: str = None):
        """Speak using UWP Narrator (OneCore) via Windows.Media.SpeechSynthesis.
        This plays audio directly and does not integrate with SAPI voices. Used as a fallback
        to get OneCore/Narrator voices speaking when SAPI can't set them.
        """
        if not UWP_TTS_AVAILABLE:
            return
        # Import lazily to avoid hard dependency at import time
        try:
            from winsdk.windows.media.speechsynthesis import SpeechSynthesizer  # type: ignore
        except Exception:
            try:
                from winsdk.windows.media.speechsynthesis import SpeechSynthesizer  # type: ignore
            except Exception:
                return
        try:
            from winsdk.windows.media.playback import MediaPlayer  # type: ignore
            from winsdk.windows.media.core import MediaSource  # type: ignore
        except Exception:
            try:
                from winsdk.windows.media.playback import MediaPlayer  # type: ignore
                from winsdk.windows.media.core import MediaSource  # type: ignore
            except Exception:
                return

        synth = SpeechSynthesizer()
        # Try to map preferred voice to UWP voice list (match by name and normalized language)
        try:
            if preferred_desc:
                voices = list(SpeechSynthesizer.all_voices)
                name_part = preferred_desc
                lang_part = None
                if ' - ' in preferred_desc:
                    name_part, lang_part = [p.strip() for p in preferred_desc.split(' - ', 1)]
                # Remove vendor prefix
                name_key = name_part.replace('Microsoft', '').strip().lower()
                # Normalize language like enAU -> en-AU
                def norm_lang(code: str) -> str:
                    if not code:
                        return ''
                    code = code.strip()
                    if '-' in code:
                        return code
                    if len(code) == 4:
                        return f"{code[:2].lower()}-{code[2:].upper()}"
                    return code
                target_lang = norm_lang(lang_part) if lang_part else ''

                # First pass: match both name and language
                chosen = None
                for v in voices:
                    v_name = getattr(v, 'display_name', '')
                    v_lang = getattr(v, 'language', '')
                    if name_key and name_key in v_name.lower():
                        if not target_lang or v_lang.lower() == target_lang.lower():
                            chosen = v
                            break
                # Second pass: fuzzy language match (prefix)
                if not chosen and name_key:
                    for v in voices:
                        v_name = getattr(v, 'display_name', '')
                        v_lang = getattr(v, 'language', '')
                        if name_key in v_name.lower():
                            if not target_lang or v_lang.lower().startswith(target_lang.split('-')[0].lower()):
                                chosen = v
                                break
                # Third pass: fallback by language only
                if not chosen and target_lang:
                    for v in voices:
                        if getattr(v, 'language', '').lower() == target_lang.lower():
                            chosen = v
                            break
                if chosen is not None:
                    synth.voice = chosen
        except Exception as _e:
            pass
        stream = await synth.synthesize_text_to_stream_async(text)
        # Enqueue for worker playback to serialize and avoid crashes
        try:
            interrupt_flag = True
            try:
                if hasattr(self, 'interrupt_on_new_scan_var'):
                    interrupt_flag = bool(self.interrupt_on_new_scan_var.get())
            except Exception:
                pass
            # If not interrupting, queue the stream; if interrupting, signal to cut current
            if interrupt_flag:
                try:
                    self._uwp_interrupt.set()
                except Exception:
                    pass
            self._uwp_queue.put(("PLAY", stream, interrupt_flag))
        except Exception:
            pass

    def _uwp_worker(self):
        # Lazy imports inside worker
        try:
            try:
                from winsdk.windows.media.playback import MediaPlayer
                from winsdk.windows.media.core import MediaSource
            except Exception:
                from winsdk.windows.media.playback import MediaPlayer  # type: ignore
                from winsdk.windows.media.core import MediaSource  # type: ignore
        except Exception:
            MediaPlayer = None
            MediaSource = None
        player = None
        while not getattr(self, '_uwp_thread_stop', threading.Event()).is_set():
            try:
                cmd, payload, interrupt_flag = self._uwp_queue.get(timeout=0.1)
            except Exception:
                continue
            if cmd == "STOP":
                try:
                    if player is not None:
                        try:
                            player.pause()
                        except Exception:
                            pass
                        player = None
                except Exception:
                    pass
                continue
            if cmd == "PLAY" and MediaPlayer is not None and MediaSource is not None:
                stream = payload
                try:
                    # If not interrupting and player is active, wait for it to finish before playing next
                    if player is not None and not interrupt_flag:
                        try:
                            try:
                                from winsdk.windows.media.playback import MediaPlaybackState  # type: ignore
                            except Exception:
                                try:
                                    from winsdk.windows.media.playback import MediaPlaybackState  # type: ignore
                                except Exception:
                                    MediaPlaybackState = None
                            if MediaPlaybackState is not None:
                                while True:
                                    session = getattr(player, 'playback_session', None)
                                    current_state = None
                                    if session is not None:
                                        try:
                                            current_state = session.playback_state
                                        except Exception:
                                            pass
                                    # proceed when not actively playing or when interrupted
                                    if current_state is None or int(current_state) != int(MediaPlaybackState.PLAYING) or self._uwp_interrupt.is_set():
                                        break
                                    time.sleep(0.01)
                        except Exception:
                            # If we can't read state, just fall through without long waits
                            pass
                    # If interrupt flag set or interrupt event set, stop current
                    if player is not None and (interrupt_flag or self._uwp_interrupt.is_set()):
                        try:
                            player.pause()
                        except Exception:
                            pass
                        player = None
                        try:
                            self._uwp_interrupt.clear()
                        except Exception:
                            pass
                    # Start new
                    player = MediaPlayer()
                    try:
                        vol = float(self.volume.get()) if hasattr(self, 'volume') else 100.0
                        player.volume = max(0.0, min(1.0, vol / 100.0))
                    except Exception:
                        pass
                    player.source = MediaSource.create_from_stream(stream, 'audio/wav')
                    player.play()
                except Exception:
                    # swallow and continue
                    pass
            
    def check_tesseract_installed(self):
        """Check if Tesseract OCR is properly installed and accessible."""
        try:
            # Try to get Tesseract version
            version = pytesseract.get_tesseract_version()
            return True, f"Tesseract {version} - Installed"
        except Exception as e:
            # Check if the default path exists
            default_path = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
            if os.path.exists(default_path):
                return False, "Tesseract found but not working properly"
            else:
                return False, "Not found or not installed in default path: C:\Program Files\Tesseract-OCR"
    

    
    def restart_tesseract(self):
        """Forcefully stop the speech and reinitialize the system."""
        print("Forcing stop...")
        try:
            self.stop_speaking()  # Stop the speech
            print("System reinitialized. Audio stopped.")
        except Exception as e:
            print(f"Error during forced stop: {e}")


    
    def _ensure_speech_ready(self):
        """Ensure the speech engine is ready before speaking."""
        try:
            # Check if we need to prime the voice for this speech session
            if not hasattr(self, '_voice_primed') or not self._voice_primed:
                print("Priming voice for first speech call...")
                
                # Make a silent priming call to ensure the voice engine is ready
                self.speaker.Speak("", 1)  # Silent priming call
                time.sleep(0.1)  # Brief pause for engine initialization
                
                # Mark as primed for this session
                self._voice_primed = True
                print("Voice priming completed")
                
        except Exception as prime_error:
            print(f"Warning: Voice priming failed (non-critical): {prime_error}")
            # Don't fail speech if priming fails
    
    def _wake_up_online_voices(self):
        """Special initialization for Online SAPI5 voices that require network initialization."""
        try:
            print("Initializing Online voices...")
            
            # Get all voices and identify online ones
            voices = self.speaker.GetVoices()
            online_voices = []
            
            for i in range(voices.Count):
                try:
                    voice = voices.Item(i)
                    voice_desc = voice.GetDescription() if hasattr(voice, 'GetDescription') else ""
                    
                    # Check if this is an online voice (Microsoft Online voices typically contain "Online")
                    if "Online" in voice_desc and "Microsoft" in voice_desc:
                        online_voices.append((i, voice, voice_desc))
                except Exception as voice_error:
                    continue
            
            if not online_voices:
                print("No online voices found")
                return
            
            print(f"Found {len(online_voices)} online voices, initializing...")
            
            # Initialize each online voice with a longer warm-up
            for idx, voice, desc in online_voices[:2]:  # Limit to first 2 online voices
                try:
                    # Select this online voice
                    self.speaker.Voice = voice
                    
                    # Make a longer warm-up call for online voices
                    self.speaker.Speak("", 1)  # Use "Initializing" text
                    time.sleep(0.5)  # Longer wait for online voice initialization
                    
                except Exception as online_error:
                    print(f"Warning: Failed to initialize online voice: {online_error}")
                    continue
            
            # Restore the first voice as default
            if voices.Count > 0:
                try:
                    self.speaker.Voice = voices.Item(0)
                    print("Restored default voice selection after online voice initialization")
                except:
                    pass
            
            print("Online voice initialization completed")
            
        except Exception as e:
            print(f"Warning: Online voice initialization failed (non-critical): {e}")
            # Don't fail the program if online voice initialization fails
    
    def setup_gui(self):
        # Create main frames for organization
        top_frame = tk.Frame(self.root)
        top_frame.pack(fill='x', padx=10, pady=5)
        
        control_frame = tk.Frame(self.root)
        control_frame.pack(fill='x', padx=10, pady=5)
        
        options_frame = tk.Frame(self.root)
        options_frame.pack(fill='x', padx=10, pady=5)
        
        # Top frame contents - Title and buttons
        title_label = tk.Label(top_frame, text=f"GameReader v{APP_VERSION}", font=("Helvetica", 12, "bold"))
        title_label.pack(side='left', padx=(0, 20))
        
        # Volume control in top frame
        volume_frame = tk.Frame(top_frame)
        volume_frame.pack(side='left', padx=10)
        
        tk.Label(volume_frame, text="Volume %:").pack(side='left')
        vcmd = (self.root.register(lambda P: self.validate_numeric_input(P, is_speed=False)), '%P')
        volume_entry = tk.Entry(volume_frame, textvariable=self.volume, width=4, validate='all', validatecommand=vcmd)
        volume_entry.pack(side='left', padx=5)
        
        # Add Set Volume button
        set_volume_button = tk.Button(volume_frame, text="Set", command=lambda: self.set_volume())
        set_volume_button.pack(side='left', padx=5)
        
        # Right-aligned buttons in top frame
        buttons_frame = tk.Frame(top_frame)
        buttons_frame.pack(side='right')
        
        debug_button = tk.Button(buttons_frame, text="🔧 Debug Window", command=self.show_debug)
        debug_button.pack(side='left', padx=5)
        
        info_button = tk.Button(buttons_frame, text="ℹ️ Info/Help", command=self.show_info)
        info_button.pack(side='left', padx=5)
        
        # Remove stop_hotkey_button from here since we'll add it to add_area_frame
        
        # Control frame contents
        layout_frame = tk.Frame(control_frame)
        layout_frame.pack(side='left', fill='x', expand=True)
        
        tk.Label(layout_frame, text="Loaded Layout:").pack(side='left')
        # Show 'n/a' when no layout is loaded, without changing the underlying value used by logic
        self.layout_label = tk.Label(layout_frame, text="n/a", font=("Helvetica", 10, "bold"))
        self.layout_label.pack(side='left', padx=5)

        def _refresh_layout_label(*_):
            value = self.layout_file.get()
            self.layout_label.config(text=(os.path.basename(value) if value else "n/a"))

        # Update label whenever layout changes
        try:
            self.layout_file.trace_add('write', _refresh_layout_label)
        except Exception:
            # Fallback for older Tk versions
            self.layout_file.trace('w', _refresh_layout_label)
        _refresh_layout_label()
        
        # Layout control buttons
        layout_buttons_frame = tk.Frame(control_frame)
        layout_buttons_frame.pack(side='right')
        
        # Add Program Saves button
        program_saves_button = tk.Button(layout_buttons_frame, text="Program Saves...", 
                                       command=self.open_game_reader_folder)
        program_saves_button.pack(side='left', padx=5)
        
        save_button = tk.Button(layout_buttons_frame, text="💾 Save Layout", command=self.save_layout)
        save_button.pack(side='left', padx=5)
        
        load_button = tk.Button(layout_buttons_frame, text="📁 Load Layout..", command=self.load_layout)
        load_button.pack(side='left', padx=5)
        
        # Options frame contents
        # Word filtering frame
        filter_frame = tk.Frame(options_frame)
        filter_frame.pack(fill='x', pady=5)
        
        tk.Label(filter_frame, text="Ignored Word List:").pack(side='left')
        self.bad_word_entry = ttk.Entry(filter_frame, textvariable=self.bad_word_list)
        self.bad_word_entry.pack(side='left', fill='x', expand=True)
        
        # Add context menu for copy/paste
        self.bad_word_menu = tk.Menu(self.root, tearoff=0)
        self.bad_word_menu.add_command(label="Cut", command=lambda: self.bad_word_entry.event_generate('<<Cut>>'))
        self.bad_word_menu.add_command(label="Copy", command=lambda: self.bad_word_entry.event_generate('<<Copy>>'))
        self.bad_word_menu.add_command(label="Paste", command=lambda: self.bad_word_entry.event_generate('<<Paste>>'))
        self.bad_word_menu.add_separator()
        self.bad_word_menu.add_command(label="Select All", command=lambda: self.bad_word_entry.selection_range(0, 'end'))
        
        def show_bad_word_menu(event):
            self.bad_word_menu.post(event.x_root, event.y_root)
            
        self.bad_word_entry.bind('<Button-3>', show_bad_word_menu)
        
        # Single line of checkboxes
        checkbox_frame = tk.Frame(options_frame)
        checkbox_frame.pack(fill='x', pady=5)
        
        # Create all checkboxes in a single line
        self.create_checkbox(checkbox_frame, "Ignore usernames:", self.ignore_usernames_var, side='left', padx=5)
        self.create_checkbox(checkbox_frame, "Ignore previous spoken words:", self.ignore_previous_var, side='left', padx=5)
        self.create_checkbox(checkbox_frame, "Ignore gibberish:", self.ignore_gibberish_var, side='left', padx=5)
      #  self.create_checkbox(checkbox_frame, "Pause at punctuation:", self.pause_at_punctuation_var, side='left', padx=5) ### Removed no point.
        # Add the new checkbox for better unit detection before fullscreen
        self.create_checkbox(checkbox_frame, "Better unit detection:", self.better_unit_detection_var, side='left', padx=5)
        # Add the new checkbox for read game units
        self.create_checkbox(checkbox_frame, "Read gamer units:", self.read_game_units_var, side='left', padx=5)
        self.create_checkbox(checkbox_frame, "Fullscreen mode:", self.fullscreen_mode_var, side='left', padx=5)
        
        # Add checkbox to allow left/right mouse buttons as hotkeys (keep on same line)
        self.allow_mouse_buttons_var = tk.BooleanVar(value=False)
        self.create_checkbox(checkbox_frame, "Allow mouse left/right:", self.allow_mouse_buttons_var, side='left', padx=5)

        
        # Add read area button in a separate frame
        add_area_frame = tk.Frame(self.root)
        add_area_frame.pack(fill='x', padx=10, pady=5)
        
        # Left side - Add Read Area button
        add_area_button = tk.Button(add_area_frame, text="➕ Add Read Area", 
                                  command=self.add_read_area,
                                  font=("Helvetica", 10))
        add_area_button.pack(side='left')
        
        # Center - Status label with larger, classic font
        self.status_frame = tk.Frame(add_area_frame)
        self.status_frame.pack(side='left', fill='x', expand=True, padx=10)
        self.status_label = tk.Label(self.status_frame, text="", 
                                    font=("Helvetica", 10, ),  # Changed font and size
                                    fg="black")  # Optional: added color for better visibility
        self.status_label.pack(side='top')
        

        
        # Right side - Set Stop Hotkey button
        self.stop_hotkey_button = tk.Button(add_area_frame, text="Set Stop Hotkey", 
                                          command=self.set_stop_hotkey)
        self.stop_hotkey_button.pack(side='right')

        # Container for the Auto Read row (placed above the scroll area)
        self.auto_read_outer_frame = tk.Frame(self.root)
        self.auto_read_outer_frame.pack(fill='x', padx=10, pady=(4, 2))

        # Thin separator line under Auto Read and before the scroll area
        ttk.Separator(self.root, orient='horizontal').pack(fill='x', padx=10, pady=(2, 2))
        
        # Frame for the areas - now with scrollable canvas
        self.area_outer_frame = tk.Frame(self.root)
        self.area_outer_frame.pack(fill='both', expand=True, pady=5)

        self.area_canvas = tk.Canvas(self.area_outer_frame, highlightthickness=0)
        self.area_canvas.pack(side='left', fill='both', expand=True)
        self.area_scrollbar = tk.Scrollbar(self.area_outer_frame, orient='vertical', command=self.area_canvas.yview)
        self.area_scrollbar.pack(side='right', fill='y')

        # Enable mouse wheel scrolling for the canvas only when mouse is over it
        def _on_mousewheel(event):
            if self.area_canvas.bbox('all') and self.area_canvas.winfo_height() < (self.area_canvas.bbox('all')[3] - self.area_canvas.bbox('all')[1]):
                self.area_canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')
            return "break"
        def _bind_mousewheel(event):
            self.area_canvas.bind_all('<MouseWheel>', _on_mousewheel)
        def _unbind_mousewheel(event):
            self.area_canvas.unbind_all('<MouseWheel>')
        self.area_canvas.bind('<Enter>', _bind_mousewheel)
        self.area_canvas.bind('<Leave>', _unbind_mousewheel)
        # If you want to support Linux (Button-4/5), add similar binds for those events.
        
        # Create a frame inside the canvas for area frames
        self.area_frame = tk.Frame(self.area_canvas)
        self.area_window = self.area_canvas.create_window((0, 0), window=self.area_frame, anchor='nw')
        self.area_canvas.configure(yscrollcommand=self.area_scrollbar.set)
        
        # Bind resizing
        def on_frame_configure(event):
            self.area_canvas.configure(scrollregion=self.area_canvas.bbox('all'))
            # Center the inner frame by setting its width to the canvas width
            canvas_width = self.area_canvas.winfo_width()
            self.area_canvas.itemconfig(self.area_window, width=canvas_width)
        self.area_frame.bind('<Configure>', on_frame_configure)
        self.area_canvas.bind('<Configure>', on_frame_configure)
        
        # Only show scrollbar if needed (handled in resize_window)
        self.area_scrollbar.pack_forget()
        
        # Bind click event to root to remove focus from entry fields
        self.root.bind("<Button-1>", self.remove_focus)
        

        
        print("GUI setup complete.")
        
        # Check Tesseract installation and update status label if not installed
        tesseract_installed, tesseract_message = self.check_tesseract_installed()
        if not tesseract_installed:
            self.status_label.config(
                text="→ Tesseract OCR missing. click the [Info/Help] button for instructions. ←",
                fg="red",
                font=("Helvetica", 10, "bold")
            )
        


    def create_checkbox(self, parent, text, variable, side='top', padx=0, pady=2):
        """Helper method to create consistent checkboxes"""
        frame = tk.Frame(parent)
        frame.pack(side=side, padx=padx, pady=pady)
        
        checkbox = tk.Checkbutton(frame, variable=variable)
        checkbox.pack(side='right')
        
        label = tk.Label(frame, text=text)
        label.pack(side='right')

    def set_volume(self):
        """Helper method to set volume"""
        try:
            vol = int(self.volume.get())
            if 0 <= vol <= 100:
                self.speaker.Volume = vol
                print(f"Program volume set to {vol}%\n--------------------------")
            else:
                self.volume.set("100")
                self.speaker.Volume = 100
                print("Volume out of range, set to 100")
        except ValueError:
            self.volume.set("100")
            self.speaker.Volume = 100
            print("Invalid volume value, set to 100")

    def remove_focus(self, event):
        widget = event.widget
        if not isinstance(widget, tk.Entry):
            self.root.focus()
    
    def show_info(self):
        # Create Tkinter window with a modern look
        info_window = tk.Toplevel(self.root)
        info_window.title("GameReader - Information")
        info_window.geometry("900x600")  # Slightly taller for better spacing

        # --- Set flag to prevent hotkeys from interfering with info window ---
        self.info_window_open = True
        
        # On close, clear the flag
        def on_info_close():
            self.info_window_open = False
            info_window.destroy()

        info_window.protocol("WM_DELETE_WINDOW", on_info_close)
        info_window.bind('<Escape>', lambda e: on_info_close())
        
        # Set window icon if available
        try:
            icon_path = os.path.join(os.path.dirname(__file__), 'Assets', 'icon.ico')
            if os.path.exists(icon_path):
                info_window.iconbitmap(icon_path)
        except Exception as e:
            print(f"Error setting info window icon: {e}")
        
        # Main container with reduced padding
        main_frame = ttk.Frame(info_window, padding="15 10 15 5")
        main_frame.pack(fill='both', expand=True)
        
        # Title section
        title_frame = ttk.Frame(main_frame)
        title_frame.pack(fill='x', pady=(0, 10))
        
        title_label = ttk.Label(title_frame, 
                               text=f"GameReader v{APP_VERSION}", 
                               font=("Helvetica", 16, "bold"))
        title_label.pack(side='left')
        

        
        # Credits/Links Area replaced by clickable images
        credits_frame = ttk.Frame(main_frame)
        credits_frame.pack(fill='x', pady=(0, 10))

        images_row = ttk.Frame(credits_frame)
        images_row.pack(fill='x', pady=(0, 3))

        # Resolve Assets paths
        assets_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Assets')
        coffee_path = os.path.join(assets_dir, 'Coffe_info.png')
        google_form_path = os.path.join(assets_dir, 'Google_form_info.png')
        github_path = os.path.join(assets_dir, 'Github_info.png')

        # Load images and keep references on the window to avoid garbage collection
        try:
            coffee_img = Image.open(coffee_path)
            google_img = Image.open(google_form_path)
            github_img = Image.open(github_path)

            # Store PIL images
            info_window.coffee_pil = coffee_img
            info_window.google_pil = google_img
            info_window.github_pil = github_img

            # Determine available width and compute downscale factor so all three fit on one row
            info_window.update_idletasks()
            try:
                window_width = info_window.winfo_width()
            except Exception:
                window_width = 900
            horizontal_padding = 40  # main_frame left/right padding approx
            per_canvas_pad = 20      # padx=10 on each side per canvas
            total_pad = 3 * per_canvas_pad
            hover_scale = 1.08
            sum_orig_widths = coffee_img.size[0] + google_img.size[0] + github_img.size[0]
            available_width = max(200, window_width - horizontal_padding)
            needed_width = hover_scale * sum_orig_widths + total_pad
            base_scale = 1.0 if needed_width <= available_width else max(0.2, (available_width - total_pad) / (hover_scale * sum_orig_widths))

            # Create normal and hover-sized images with scaling applied
            def make_photos(pil_img):
                w, h = pil_img.size
                w_norm = max(1, int(w * base_scale))
                h_norm = max(1, int(h * base_scale))
                w_hover = max(1, int(w * base_scale * hover_scale))
                h_hover = max(1, int(h * base_scale * hover_scale))
                normal = ImageTk.PhotoImage(pil_img.resize((w_norm, h_norm), Image.LANCZOS))
                hover = ImageTk.PhotoImage(pil_img.resize((w_hover, h_hover), Image.LANCZOS))
                return normal, hover

            info_window.coffee_photo, info_window.coffee_photo_hover = make_photos(info_window.coffee_pil)
            info_window.google_photo, info_window.google_photo_hover = make_photos(info_window.google_pil)
            info_window.github_photo, info_window.github_photo_hover = make_photos(info_window.github_pil)

            # Smooth animations for hover effects
            def _cancel_anim(c):
                if hasattr(c, "_anim_job") and c._anim_job:
                    try:
                        c.after_cancel(c._anim_job)
                    except Exception:
                        pass
                    c._anim_job = None

            def animate_to_hover(canvas, image_id, pil_img):
                duration_ms = 100
                steps = 12
                _cancel_anim(canvas)
                frames = []

                def step(i):
                    t = i / steps
                    scale = base_scale + (base_scale * hover_scale - base_scale) * t
                    w = max(1, int(pil_img.size[0] * scale))
                    h = max(1, int(pil_img.size[1] * scale))
                    frame = ImageTk.PhotoImage(pil_img.resize((w, h), Image.LANCZOS))
                    frames.append(frame)
                    canvas.itemconfig(image_id, image=frame)
                    if i < steps:
                        canvas._anim_job = canvas.after(int(duration_ms / steps), lambda: step(i + 1))
                    else:
                        canvas._anim_job = None
                        canvas._anim_frames = frames  # keep refs

                step(0)

            def animate_to_normal(canvas, image_id, pil_img):
                duration_ms = 230
                steps = 15
                _cancel_anim(canvas)
                frames = []

                def step(i):
                    t = i / steps
                    scale = (base_scale * hover_scale) + (base_scale - base_scale * hover_scale) * t
                    w = max(1, int(pil_img.size[0] * scale))
                    h = max(1, int(pil_img.size[1] * scale))
                    frame = ImageTk.PhotoImage(pil_img.resize((w, h), Image.LANCZOS))
                    frames.append(frame)
                    canvas.itemconfig(image_id, image=frame)
                    if i < steps:
                        canvas._anim_job = canvas.after(int(duration_ms / steps), lambda: step(i + 1))
                    else:
                        canvas._anim_job = None
                        canvas._anim_frames = frames  # keep refs

                step(0)



            # Coffee image (fixed-size canvas, allows hover to be clipped by bounds)
            coffee_cw = info_window.coffee_photo_hover.width()
            coffee_ch = info_window.coffee_photo_hover.height()
            coffee_canvas = tk.Canvas(
                images_row,
                width=coffee_cw,
                height=coffee_ch,
                highlightthickness=0,
                bd=0,
                cursor='hand2',
                takefocus=1
            )
            coffee_canvas.pack(side='left', padx=10)
            coffee_img_id = coffee_canvas.create_image(coffee_cw // 2, coffee_ch // 2, image=info_window.coffee_photo)
            
            # Store original hover state for coffee canvas
            coffee_canvas._was_hovered = False
            
            def coffee_click_start(e):
                # Store current hover state and instantly shrink image
                coffee_canvas._was_hovered = getattr(coffee_canvas, '_is_hovered', False)
                # Cancel any ongoing animations and instantly show normal size
                _cancel_anim(coffee_canvas)
                w, h = info_window.coffee_pil.size
                w_norm = max(1, int(w * base_scale))
                h_norm = max(1, int(h * base_scale))
                normal_photo = ImageTk.PhotoImage(info_window.coffee_pil.resize((w_norm, h_norm), Image.LANCZOS))
                coffee_canvas.itemconfig(coffee_img_id, image=normal_photo)
                # Keep reference to prevent garbage collection
                coffee_canvas._click_photo = normal_photo
            
            def coffee_click_end(e):
                # Restore to hover state if it was hovered before
                if coffee_canvas._was_hovered:
                    animate_to_hover(coffee_canvas, coffee_img_id, info_window.coffee_pil)
                # Open URL when click is released
                print("Coffee image clicked!")
                open_url("https://buymeacoffee.com/mertennor")
            
            # Bind animation events - URL opening is handled in click_end
            coffee_canvas.bind("<ButtonPress-1>", coffee_click_start)
            coffee_canvas.bind("<ButtonRelease-1>", coffee_click_end)
            coffee_canvas.bind(
                "<Enter>",
                lambda e, c=coffee_canvas, iid=coffee_img_id: (setattr(c, '_is_hovered', True), animate_to_hover(c, iid, info_window.coffee_pil))
            )
            coffee_canvas.bind(
                "<Leave>",
                lambda e, c=coffee_canvas, iid=coffee_img_id: (setattr(c, '_is_hovered', False), animate_to_normal(c, iid, info_window.coffee_pil))
            )

            # Google Form image (fixed-size canvas)
            google_cw = info_window.google_photo_hover.width()
            google_ch = info_window.google_photo_hover.height()
            google_canvas = tk.Canvas(
                images_row,
                width=google_cw,
                height=google_ch,
                highlightthickness=0,
                bd=0,
                cursor='hand2',
                takefocus=1
            )
            google_canvas.pack(side='left', padx=10)
            google_img_id = google_canvas.create_image(google_cw // 2, google_ch // 2, image=info_window.google_photo)
            
            # Store original hover state for google canvas
            google_canvas._was_hovered = False
            
            def google_click_start(e):
                # Store current hover state and instantly shrink image
                google_canvas._was_hovered = getattr(google_canvas, '_is_hovered', False)
                # Cancel any ongoing animations and instantly show normal size
                _cancel_anim(google_canvas)
                w, h = info_window.google_pil.size
                w_norm = max(1, int(w * base_scale))
                h_norm = max(1, int(h * base_scale))
                normal_photo = ImageTk.PhotoImage(info_window.google_pil.resize((w_norm, h_norm), Image.LANCZOS))
                google_canvas.itemconfig(google_img_id, image=normal_photo)
                # Keep reference to prevent garbage collection
                google_canvas._click_photo = normal_photo
            
            def google_click_end(e):
                # Restore to hover state if it was hovered before
                if google_canvas._was_hovered:
                    animate_to_hover(google_canvas, google_img_id, info_window.google_pil)
                # Open URL when click is released
                print("Google Form image clicked!")
                open_url("https://forms.gle/8YBU8atkgwjyzdM79")
            
            # Bind animation events - URL opening is handled in click_end
            google_canvas.bind("<ButtonPress-1>", google_click_start)
            google_canvas.bind("<ButtonRelease-1>", google_click_end)
            google_canvas.bind(
                "<Enter>",
                lambda e, c=google_canvas, iid=google_img_id: (setattr(c, '_is_hovered', True), animate_to_hover(c, iid, info_window.google_pil))
            )
            google_canvas.bind(
                "<Leave>",
                lambda e, c=google_canvas, iid=google_img_id: (setattr(c, '_is_hovered', False), animate_to_normal(c, iid, info_window.google_pil))
            )

            # GitHub image (fixed-size canvas)
            github_cw = info_window.github_photo_hover.width()
            github_ch = info_window.github_photo_hover.height()
            github_canvas = tk.Canvas(
                images_row,
                width=github_cw,
                height=github_ch,
                highlightthickness=0,
                bd=0,
                cursor='hand2',
                takefocus=1
            )
            github_canvas.pack(side='left', padx=10)
            github_img_id = github_canvas.create_image(github_cw // 2, github_ch // 2, image=info_window.github_photo)
            
            # Store original hover state for github canvas
            github_canvas._was_hovered = False
            
            def github_click_start(e):
                # Store current hover state and instantly shrink image
                github_canvas._was_hovered = getattr(github_canvas, '_is_hovered', False)
                # Cancel any ongoing animations and instantly show normal size
                _cancel_anim(github_canvas)
                w, h = info_window.github_pil.size
                w_norm = max(1, int(w * base_scale))
                h_norm = max(1, int(h * base_scale))
                normal_photo = ImageTk.PhotoImage(info_window.github_pil.resize((w_norm, h_norm), Image.LANCZOS))
                github_canvas.itemconfig(github_img_id, image=normal_photo)
                # Keep reference to prevent garbage collection
                github_canvas._click_photo = normal_photo
            
            def github_click_end(e):
                # Restore to hover state if it was hovered before
                if github_canvas._was_hovered:
                    animate_to_hover(github_canvas, github_img_id, info_window.github_pil)
                # Open URL when click is released
                print("GitHub image clicked!")
                open_url("https://github.com/MertenNor/GameReader")
            
            # Bind animation events - URL opening is handled in click_end
            github_canvas.bind("<ButtonPress-1>", github_click_start)
            github_canvas.bind("<ButtonRelease-1>", github_click_end)
            github_canvas.bind(
                "<Enter>",
                lambda e, c=github_canvas, iid=github_img_id: (setattr(c, '_is_hovered', True), animate_to_hover(c, iid, info_window.github_pil))
            )
            github_canvas.bind(
                "<Leave>",
                lambda e, c=github_canvas, iid=github_img_id: (setattr(c, '_is_hovered', False), animate_to_normal(c, iid, info_window.github_pil))
            )
        except Exception as e:
            # Fallback text if images can't be displayed
            fallback = ttk.Label(credits_frame, text=f"Error displaying info images: {e}", foreground='red')
            fallback.pack(anchor='w')

        # Coffee note below the images
        coffee_note = ttk.Label(
            credits_frame,
            text="☕ Note: You don't have to fuel my caffeine addiction… but I wouldn't say no! Every coffee helps me argue with AI until the code finally works. All funds are shared between me and the few helping me bring this project to life.",
            font=("Helvetica", 9, "bold"),
            foreground='#666666',
            wraplength=800,
            justify='center'
        )
        coffee_note.pack(pady=(8, 10), anchor='center')

        
        # Tesseract Status Indicator
        tesseract_status_frame = ttk.Frame(credits_frame)
        tesseract_status_frame.pack(fill='x', pady=(5, 0))
        
        # Check Tesseract installation status
        tesseract_installed, tesseract_message = self.check_tesseract_installed()
        
        # Status label with appropriate color
        status_color = 'green' if tesseract_installed else 'red'
        status_text = "✓ " if tesseract_installed else "✗ "
        
        if tesseract_installed:
            # Simple status when installed - create a frame to hold both labels
            status_row = ttk.Frame(tesseract_status_frame)
            status_row.pack(anchor='w', pady=(0, 5))
            
            # Black text for main status
            main_status_label = ttk.Label(
                status_row,
                text="Tesseract OCR Status: ",
                font=("Helvetica", 11, "bold"),
                foreground='black'
            )
            main_status_label.pack(side='left')
            
            # Green checkmark
            checkmark_label = ttk.Label(
                status_row,
                text=status_text,
                font=("Helvetica", 11, "bold"),
                foreground=status_color
            )
            checkmark_label.pack(side='left')
            
            # Green text for (Installed)
            installed_label = ttk.Label(
                status_row,
                text="(Installed)",
                font=("Helvetica", 11, "bold"),
                foreground='green'
            )
            installed_label.pack(side='left')
        else:
            # Detailed status when not installed - create a frame to hold both labels
            status_row = ttk.Frame(tesseract_status_frame)
            status_row.pack(anchor='w', pady=(0, 5))
            
            # Black text for main status
            main_status_label = ttk.Label(
                status_row,
                text="Tesseract OCR Status: ",
                font=("Helvetica", 11, "bold"),
                foreground='black'
            )
            main_status_label.pack(side='left')
            
            # Red X
            x_label = ttk.Label(
                status_row,
                text=status_text,
                font=("Helvetica", 11, "bold"),
                foreground=status_color
            )
            x_label.pack(side='left')
            
            # Black text for (Required for GameReader to fully function)
            required_label = ttk.Label(
                status_row,
                text="(Requierd for GameReader to fully function)",
                font=("Helvetica", 11, "bold"),
                foreground='red'
            )
            required_label.pack(side='left')
            
            # Reason label
            reason_label = ttk.Label(
                tesseract_status_frame,
                text=f"Reason: {tesseract_message}",
                font=("Helvetica", 10),
                foreground='red'
            )
            reason_label.pack(anchor='w', pady=(0, 5))
        
        # Download instruction and clickable URLs on the same line (always visible)
        download_row = ttk.Frame(tesseract_status_frame)
        download_row.pack(anchor='w')
        download_label = ttk.Label(download_row,
                                   text="Download Tesseract OCR from here: ",
                                   font=("Helvetica", 10, "bold"),
                                   foreground='black')
        download_label.pack(side='left')
        
        # First link to Tesseract releases page
        tesseract_link = ttk.Label(download_row,
                                   text="https://github.com/tesseract-ocr/tesseract/releases",
                                   font=("Helvetica", 10),
                                   foreground='blue',
                                   cursor='hand2')
        tesseract_link.pack(side='left')
        tesseract_link.bind("<Button-1>", lambda e: open_url("https://github.com/tesseract-ocr/tesseract/releases"))
        tesseract_link.bind("<Enter>", lambda e: tesseract_link.configure(font=("Helvetica", 10, "underline")))
        tesseract_link.bind("<Leave>", lambda e: tesseract_link.configure(font=("Helvetica", 10)))
        
        # Add separator
        ttk.Label(download_row, text=" | ").pack(side='left')
        
        # Direct download link for Windows installer
        direct_link = ttk.Label(download_row,
                               text="Direct link to Tesseract Windows installer v.5.5.0",
                               font=("Helvetica", 10),
                               foreground='blue',
                               cursor='hand2')
        direct_link.pack(side='left')
        direct_link.bind("<Button-1>", lambda e: open_url("https://github.com/tesseract-ocr/tesseract/releases/download/5.5.0/tesseract-ocr-w64-setup-5.5.0.20241111.exe"))
        direct_link.bind("<Enter>", lambda e: direct_link.configure(font=("Helvetica", 10, "underline")))
        direct_link.bind("<Leave>", lambda e: direct_link.configure(font=("Helvetica", 10)))
        
        # Add NaturalVoiceSAPIAdapter information with reduced spacing
        
        # NaturalVoiceSAPIAdapter section
        natural_voice_frame = ttk.Frame(tesseract_status_frame)
        natural_voice_frame.pack(anchor='w', pady=(25, 0))
        
        natural_voice_label = ttk.Label(
            natural_voice_frame,
            text="For more and higher-quality voices:",
            font=("Helvetica", 12, "bold"),
            foreground='black'
        )
        natural_voice_label.pack(anchor='w')
        
        # Create a frame for the text and link to be on the same line
        text_link_frame = ttk.Frame(natural_voice_frame)
        text_link_frame.pack(anchor='w')
        
        natural_voice_desc = ttk.Label(
            text_link_frame,
            text="Download and install NaturalVoiceSAPIAdapter by gexgd0419 here: ",
            font=("Helvetica", 10, "bold"),
            foreground='black'
        )
        natural_voice_desc.pack(side='left')
        
        # Clickable link to NaturalVoiceSAPIAdapter
        natural_voice_link = ttk.Label(
            text_link_frame,
            text="https://github.com/gexgd0419/NaturalVoiceSAPIAdapter",
            font=("Helvetica", 10),
            foreground='blue',
            cursor='hand2'
        )
        natural_voice_link.pack(side='left')
        natural_voice_link.bind("<Button-1>", lambda e: open_url("https://github.com/gexgd0419/NaturalVoiceSAPIAdapter?tab=readme-ov-file#installation"))
        natural_voice_link.bind("<Enter>", lambda e: natural_voice_link.configure(font=("Helvetica", 10, "underline")))
        natural_voice_link.bind("<Leave>", lambda e: natural_voice_link.configure(font=("Helvetica", 10)))
        
        tesseract_frame = ttk.Frame(credits_frame)
        tesseract_frame.pack(fill='x', pady=(5, 0))
        
        # Create a frame with scrollbar for the main content
        content_frame = ttk.Frame(main_frame)
        content_frame.pack(fill='both', expand=True)
        
        # Add fullscreen hotkey warning above the text widget
        warning_label = ttk.Label(
            content_frame,
            text="Tip: If hotkeys don't work in fullscreen apps or games, run GameReader as Administrator.",
            font=("Helvetica", 10, "bold"),
            foreground='black'
        )
        warning_label.pack(anchor='w', pady=(3, 0))
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(content_frame)
        scrollbar.pack(side='right', fill='y')
        
        # Create text widget with custom styling - make it selectable
        text_widget = tk.Text(content_frame, 
                             wrap=tk.WORD, 
                             yscrollcommand=scrollbar.set,
                             font=("Helvetica", 10),
                             padx=8,
                             pady=6,
                             spacing1=2,  # Space between lines
                             spacing2=2,  # Space between paragraphs
                             background='#f5f5f5',  # Light gray background
                             border=1,
                             state='normal',  # Make it editable initially to insert text
                             cursor='xterm',  # Show text cursor
                             selectbackground='#0078d7',  # Blue selection color
                             selectforeground='white')  # White text on selection
        
        # Add right-click context menu for copy
        def show_context_menu(event):
            context_menu = tk.Menu(text_widget, tearoff=0)
            context_menu.add_command(label="Copy", command=lambda: text_widget.event_generate('<<Copy>>'))
            context_menu.add_command(label="Select All", command=lambda: text_widget.tag_add('sel', '1.0', 'end'))
            try:
                context_menu.tk.call('tk_popup', context_menu, event.x_root, event.y_root)
            finally:
                context_menu.grab_release()
        
        text_widget.bind("<Button-3>", show_context_menu)
        
        text_widget.pack(fill='both', expand=True)
        
        # Configure scrollbar and tags
        scrollbar.config(command=text_widget.yview)
        text_widget.tag_configure('bold', font=("Helvetica", 10, "bold"))
        
        # Info text with improved formatting - split into sections
        info_text = [

            ("How to Use the Program\n", 'bold'),
            ("═══════════════════════════════\n", None),
            ("• Click \"Set Area\": Left-click and drag to select the area you want the program to read. (Area name can be change with right-click)\n\n", None),
            ("• Click \"Set Hotkey\": Assign a hotkey for the selected area.\n\n", None),
            ("• Click \"Select Voice\": Choose a voice from the dropdown menu.\n\n", None),
            ("• Press the assigned area hotkey to make the program automatically read the text aloud.\n\n", None),
            ("• Use the stop hotkey (if set) to stop the current reading.\n\n", None),
            ("• Adjust the program volume by setting the volume percentage in the main window.\n\n", None),
            ("• The debug console displays the processed image of the last area read and its debug logs.\n\n", None),
            ("• Make sure to save your loadout once you are happy with your setup.\n\n\n", None),

                    
            ("BUTTONS AND FEATURES\n", 'bold'),
            ("═════════════════════\n\n", None),


            ("Auto Read\n", 'bold'),
            ("------------------------\n", None),
            ("When assigned a hotkey, the program will automatically read the text in the selected area.\n", None),
            ("The Save button here will save the settings for the AutoRead area only.\n", None),
            ("Note! This works best with applications in windowed borderless mode.\n", None),
            ("This save file can be found here: C:\\Users\\<username>\\AppData\\Local\\Temp\nFilename: auto_read_settings.json.\n", None),
            ("Alternatively, you can locate this save file by clicking the 'Program Saves...' button.\n", None),
            ("The checkbox 'Stop Read on Select' determines the behavior when scanning a new area while text is being read.\n", None),
            ("If checked, the ongoing text will stop immediately, and the newly scanned text will be read.\n", None),
            ("If unchecked, the newly scanned text will be added to a queue and read after the ongoing text finishes.\n\n", None),

            ("Add Read Area\n", 'bold'),
            ("------------------------\n", None),
            ("Creates a new area for text capture. You can define multiple areas on screen for different text sources.\n\n", None),
            
            ("Image Processing\n", 'bold'),
            ("------------------------------\n", None),
            ("Allows customization of image preprocessing before speaking. Useful for improving text recognition in difficult-to-read areas.\n\n", None),

            ("Debug window\n", 'bold'),
            ("---------------------------\n", None),
            ("Shows the captured text and processed images for troubleshooting.\n\n", None),

            ("Stop Hotkey\n", 'bold'),
            ("--------------------\n", None),
            ("Immediately stops any ongoing speech.\n\n", None),

            ("Ignored Word List\n", 'bold'),
            ("-------------------------\n", None),
            ("A list of words, phrases, or sentences (separated by commas) to ignore while reading text. Example: Chocolate, Apple, Banana, I love ice cream\n", None),
            ("These will then be ignored in all areas.\n\n", None),

            ("CHECKBOX OPTIONS\n", 'bold'),
            ("════════════════\n\n", None),

            ("Ignore usernames *EXPERIMENTAL*\n", 'bold'),
            ("--------------------------------\n", None),
            ("This option filters out usernames from the text before reading. It looks for patterns like \"Username:\" at the start of lines.\n\n", None),

            ("Ignore previous spoken words\n", 'bold'),
            ("-------------------------------------------------\n", None),
            ("This prevents the same text from being read multiple times. Useful for chat windows where messages might persist.\n\n", None),

            ("Ignore gibberish *EXPERIMENTAL*\n", 'bold'),
            ("-------------------------------------------------------\n", None),
            ("Filters out text that appears to be random characters or rendered artifacts. Helps prevent reading of non-meaningful text.\n\n", None),

            ("Pause at punctuation *EXPERIMENTAL*\n", 'bold'),
            ("------------------------------------\n", None),
            ("Adds natural pauses when encountering periods, commas, and other punctuation marks. Makes the speech sound more natural.\n\n", None),

            ("Fullscreen mode *EXPERIMENTAL*\n", 'bold'),
            ("--------------------------------------------------------\n", None),
            ("Feature for capturing text from fullscreen applications. May cause brief screen flicker during capture for the program to take an updated screenshot.\n\n", None),

            ("TIPS AND TRICKS\n", 'bold'),
            ("═════════════\n\n", None),

            ("• Use image processing for areas with difficult-to-read text\n\n", None),

            ("• Create two identical areas with different hotkeys: assign one a male voice and the other a female voice.\n", None),
            ("  This lets you easily switch between male and female voices for text, ideal for game dialogue.\n\n", None),

            ("• Experiment with different preprocessing settings for optimal text recognition in your specific use case.\n\n", None),

            ("• Want more Voices? Add them in Windows.\n", None),
            ("  For Windows 10: Go to Settings > Time & Language > Speech > Manage voices, then click 'Add voices' to install new ones.\n", None),
            ("  For Windows 11: Go to Settings > Accessibility > Speech > Manage voices, then click 'Add voices' to install new ones.\n", None),
            ("  You can also find other Narrator voices online that you can add to Windows.\n", None),
        ]
        
        # Insert text with tags
        for text, tag in info_text:
            text_widget.insert('end', text, tag)
        
        # Enable text selection and copying even when disabled
        def enable_text_selection(event=None):
            return 'break'
            
        text_widget.bind('<Key>', enable_text_selection)
        text_widget.bind('<Control-c>', lambda e: text_widget.event_generate('<<Copy>>') or 'break')
        text_widget.bind('<Control-a>', lambda e: (text_widget.tag_add('sel', '1.0', 'end'), 'break'))
        
        # Make text widget read-only but keep text selectable
        text_widget.config(state='disabled')
        
        # Add bottom frame for close button with padding
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(fill='x', pady=(20, 0))
        
        # Add close button
        close_button = ttk.Button(bottom_frame, 
                                 text="Close", 
                                 command=info_window.destroy,
                                 width=45)
        close_button.pack(side='right')
        
        # Center window on screen
        info_window.update_idletasks()
        width = info_window.winfo_width()
        height = info_window.winfo_height()
        x = (info_window.winfo_screenwidth() // 2) - (width // 2)
        y = (info_window.winfo_screenheight() // 2) - (height // 2)
        info_window.geometry(f'{width}x{height}+{x}+{y}')
        
        # Make window modal
        info_window.transient(self.root)
        info_window.grab_set()
        
    def test_hotkey_working(self, hotkey_str):
        """Test if a hotkey is working properly"""
        try:
            # Try to register the hotkey temporarily
            test_hook = keyboard.add_hotkey(hotkey_str, lambda: None, suppress=False)
            keyboard.remove_hotkey(hotkey_str)
            return True
        except Exception as e:
            print(f"Hotkey test failed for {hotkey_str}: {e}")
            return False
            
    def show_debug(self):
        if not hasattr(sys, 'stdout_original'):
            sys.stdout_original = sys.stdout
        
        if not hasattr(self, 'console_window') or not self.console_window.window.winfo_exists():
            self.console_window = ConsoleWindow(self.root, log_buffer, self.layout_file, self.latest_images, self.latest_area_name)
        else:
            self.console_window.update_console()
        sys.stdout = self.console_window
        
    def customize_processing(self, area_name_var):
        area_name = area_name_var.get()
        if area_name not in self.latest_images:
            messagebox.showerror("Error", "No image to process yet. Please generate an image by pressing the hotkey.")
            return

        if area_name not in self.processing_settings:
            self.processing_settings[area_name] = {}
        ImageProcessingWindow(self.root, area_name, self.latest_images, self.processing_settings[area_name], self)
        
    def set_stop_hotkey(self):
        # Clean up temporary hooks and disable all hotkeys
        try:
            self.disable_all_hotkeys()
        except Exception as e:
            print(f"Warning: Error cleaning up temporary hooks: {e}")

        self._hotkey_assignment_cancelled = False  # Guard flag to block late events
        self.setting_hotkey = True
        


        def finish_hotkey_assignment():
            # Restore all hotkeys after assignment is done
            try:
                self.stop_speaking()  # Stop the speech
                print("System reinitialized. Audio stopped.")
            except Exception as e:
                print(f"Error during forced stop: {e}")
            
            try:
                self.restore_all_hotkeys()
            except Exception as e:
                print(f"Warning: Error restoring hotkeys: {e}")
            # Cleanup any temp hooks and preview
            try:
                if hasattr(self, '_hotkey_preview_job') and self._hotkey_preview_job:
                    self.root.after_cancel(self._hotkey_preview_job)
                    self._hotkey_preview_job = None
            except Exception:
                pass
            try:
                if hasattr(self.stop_hotkey_button, 'keyboard_hook_temp'):
                    keyboard.unhook(self.stop_hotkey_button.keyboard_hook_temp)
                    delattr(self.stop_hotkey_button, 'keyboard_hook_temp')
            except Exception:
                try:
                    if hasattr(self.stop_hotkey_button, 'keyboard_hook_temp'):
                        delattr(self.stop_hotkey_button, 'keyboard_hook_temp')
                except Exception:
                    pass
            try:
                if hasattr(self.stop_hotkey_button, 'mouse_hook_temp'):
                    mouse.unhook(self.stop_hotkey_button.mouse_hook_temp)
                    delattr(self.stop_hotkey_button, 'mouse_hook_temp')
            except Exception:
                try:
                    if hasattr(self.stop_hotkey_button, 'mouse_hook_temp'):
                        delattr(self.stop_hotkey_button, 'mouse_hook_temp')
                except Exception:
                    pass
            try:
                if hasattr(self.stop_hotkey_button, 'shift_release_hooks'):
                    for h in getattr(self.stop_hotkey_button, 'shift_release_hooks', []) or []:
                        try:
                            keyboard.unhook(h)
                        except Exception:
                            pass
                    delattr(self.stop_hotkey_button, 'shift_release_hooks')
                if hasattr(self.stop_hotkey_button, 'ctrl_release_hooks'):
                    for h in getattr(self.stop_hotkey_button, 'ctrl_release_hooks', []) or []:
                        try:
                            keyboard.unhook(h)
                        except Exception:
                            pass
                    delattr(self.stop_hotkey_button, 'ctrl_release_hooks')
            except Exception:
                pass
        
        # Track whether a non-modifier was pressed
        combo_state = {'non_modifier_pressed': False}

        def _assign_stop_hotkey_and_register(hk_str):
            # Check duplicate against area hotkeys
            for area_frame, hotkey_button, _, area_name_var, _, _, _ in self.areas:
                if hasattr(hotkey_button, 'hotkey') and hotkey_button.hotkey == hk_str:
                    show_thinkr_warning(self, area_name_var.get())
                    self._hotkey_assignment_cancelled = True
                    self.setting_hotkey = False
                    self.stop_hotkey_button.config(text="Set Stop Hotkey")
                    finish_hotkey_assignment()
                    return False
            # Clean old stop hotkey hooks
            if hasattr(self, 'stop_hotkey'):
                try:
                    if hasattr(self.stop_hotkey_button, 'mock_button'):
                        self._cleanup_hooks(self.stop_hotkey_button.mock_button)
                except Exception as e:
                    print(f"Error cleaning up stop hotkey hooks: {e}")
            self.stop_hotkey = hk_str
            # Register
            mock_button = type('MockButton', (), {'hotkey': hk_str, 'is_stop_button': True})
            self.stop_hotkey_button.mock_button = mock_button
            self.setup_hotkey(self.stop_hotkey_button.mock_button, None)
            # Nicer display mapping for sided modifiers and numpad
            display_name = hk_str.replace('numpad ', 'NUMPAD ').replace('num_', 'num:') \
                                   .replace('ctrl','CTRL') \
                                   .replace('left alt','L-ALT').replace('right alt','R-ALT') \
                                   .replace('left shift','L-SHIFT').replace('right shift','R-SHIFT') \
                                   .replace('windows','WIN') \
                                   .replace('multiply', '*').replace('add', '+').replace('subtract', '-').replace('divide', '/')
            self.stop_hotkey_button.config(text=f"Stop Hotkey: [ {display_name.upper()} ]")
            print(f"Set Stop hotkey: {hk_str}\n--------------------------")
            self.setting_hotkey = False
            self._hotkey_assignment_cancelled = True
            finish_hotkey_assignment()
            return True

        def on_key_press(event):
            if self._hotkey_assignment_cancelled or not self.setting_hotkey:
                return
            # Ignore Escape
            if event.scan_code == 1:
                return
            
            # Use scan code and virtual key code for consistent behavior across keyboard layouts
            scan_code = getattr(event, 'scan_code', None)
            vk_code = getattr(event, 'vk_code', None)
            
            # Get current keyboard layout for consistency
            current_layout = get_current_keyboard_layout()
            
            # Determine key name based on scan code and virtual key code for consistency
            name = None
            side = None
            
            # Handle modifier keys consistently
            if scan_code == 29:  # Left Ctrl
                name = 'ctrl'
                side = 'left'
            elif scan_code == 157:  # Right Ctrl
                name = 'ctrl'
                side = 'right'
            elif scan_code == 42:  # Left Shift
                name = 'shift'
                side = 'left'
            elif scan_code == 54:  # Right Shift
                name = 'shift'
                side = 'right'
            elif scan_code == 56:  # Left Alt
                name = 'left alt'
                side = 'left'
            elif scan_code == 184:  # Right Alt
                name = 'right alt'
                side = 'right'
            elif scan_code == 91:  # Left Windows
                name = 'windows'
                side = 'left'
            elif scan_code == 92:  # Right Windows
                name = 'windows'
                side = 'right'
            else:
                # For non-modifier keys, use the event name but normalize it
                raw_name = (event.name or '').lower()
                name = normalize_key_name(raw_name)
                
                # Check if it's a numpad key
                if scan_code in self.numpad_scan_codes:
                    sym = self.numpad_scan_codes[scan_code]
                    name = f"num_{sym}"

            # Non-modifier pressed
            if name not in ('ctrl','alt','left alt','right alt','shift','windows'):
                combo_state['non_modifier_pressed'] = True
            # Bare modifier assignment path
            if name in ('ctrl','shift','alt','left alt','right alt','windows'):
                def _assign_bare_modifier():
                    if self._hotkey_assignment_cancelled or not self.setting_hotkey:
                        return
                    try:
                        held = []
                        # Use scan code detection for more reliable left/right distinction
                        left_ctrl_pressed, right_ctrl_pressed = detect_ctrl_keys()
                        
                        if left_ctrl_pressed or right_ctrl_pressed: held.append('ctrl')
                        if keyboard.is_pressed('shift'): held.append('shift')
                        if keyboard.is_pressed('left alt'): held.append('left alt')
                        if keyboard.is_pressed('right alt'): held.append('right alt')
                        if keyboard.is_pressed('left windows') or keyboard.is_pressed('right windows') or keyboard.is_pressed('windows'):
                            held.append('windows')
                        if len(held) == 1:
                            only = held[0]
                            # Determine base from name
                            base = None
                            if 'ctrl' in name: base = 'ctrl'
                            elif 'alt' in name: base = 'alt'
                            elif 'shift' in name: base = 'shift'
                            elif 'windows' in name: base = 'windows'
                            
                            if (base == 'ctrl' and only == 'ctrl') or \
                               (base == 'alt' and (only in ['left alt','right alt'])) or \
                               (base == 'shift' and only == 'shift') or \
                               (base == 'windows' and only == 'windows'):
                                key_name_local = only
                                _assign_stop_hotkey_and_register(key_name_local)
                                return
                    except Exception:
                        pass
                try:
                    print(f"Debug: Setting timer for _assign_bare_modifier with name: '{name}'")
                    self.root.after(200, _assign_bare_modifier)
                    print(f"Debug: Timer set successfully")
                except Exception as e:
                    print(f"Debug: Error setting timer: {e}")
                return

            # Build combo from held modifiers + base key
            try:
                mods = []
                # Use scan code detection for more reliable left/right distinction
                left_ctrl_pressed, right_ctrl_pressed = detect_ctrl_keys()
                
                if left_ctrl_pressed or right_ctrl_pressed: mods.append('ctrl')
                if keyboard.is_pressed('shift'): mods.append('shift')
                if keyboard.is_pressed('left alt'): mods.append('left alt')
                if keyboard.is_pressed('right alt'): mods.append('right alt')
                if keyboard.is_pressed('left windows') or keyboard.is_pressed('right windows') or keyboard.is_pressed('windows'):
                    mods.append('windows')
            except Exception:
                mods = []

            base_key = name
            if event.scan_code in self.numpad_scan_codes:
                sym = self.numpad_scan_codes[event.scan_code]
                base_key = f"num_{sym}"
            if base_key in ("ctrl", "shift", "alt", "windows", "left alt", "right alt"):
                combo_parts = (mods + [base_key]) if base_key not in mods else mods[:]
            else:
                combo_parts = mods + [base_key]
            key_name = "+".join(p for p in combo_parts if p)

            _assign_stop_hotkey_and_register(key_name)
            return
            
        def on_mouse_click(event):
            if (self._hotkey_assignment_cancelled or 
                not self.setting_hotkey or 
                not isinstance(event, mouse.ButtonEvent) or 
                event.event_type != mouse.DOWN):
                return
                
            # Only show warning for left (button1) and right (button2) mouse buttons when not allowed
            if event.button in [1, 2]:  # 1 = left button, 2 = right button
                if not hasattr(self, 'allow_mouse_buttons_var') or not self.allow_mouse_buttons_var.get():
                    messagebox.showwarning(
                        "Error", "Left and right mouse buttons cannot be used as hotkeys.\nCheck 'Allow mouse left/right:' to enable them.")
                    self._mouse_button_error_shown = True
                    self._hotkey_assignment_cancelled = True
                    self.setting_hotkey = False
                    self.stop_hotkey_button.config(text="Set Stop Hotkey")
                    finish_hotkey_assignment()
                    return
                
            key_name = f"button{event.button}"
            
            # Check if this mouse button is already used by any area
            for area_frame, hotkey_button, _, area_name_var, _, _, _ in self.areas:
                if hasattr(hotkey_button, 'hotkey') and hotkey_button.hotkey == key_name:
                    show_thinkr_warning(self, area_name_var.get())
                    self._hotkey_assignment_cancelled = True
                    self.setting_hotkey = False
                    self.stop_hotkey_button.config(text="Set Stop Hotkey")
                    finish_hotkey_assignment()
                    return
            
            # Remove existing stop hotkey if it exists
            if hasattr(self, 'stop_hotkey'):
                try:
                    if hasattr(self.stop_hotkey_button, 'mock_button'):
                        self._cleanup_hooks(self.stop_hotkey_button.mock_button)
                except Exception as e:
                    print(f"Error cleaning up stop hotkey hooks: {e}")
            
            self.stop_hotkey = key_name
            
            # Create a mock button object to use with setup_hotkey
            mock_button = type('MockButton', (), {'hotkey': key_name, 'is_stop_button': True})
            self.stop_hotkey_button.mock_button = mock_button  # Store reference to mock button
            
            # Setup the hotkey
            self.setup_hotkey(self.stop_hotkey_button.mock_button, None)  # Pass None as area_frame for stop hotkey
            
            display_name = f"Mouse Button {event.button}"
            self.stop_hotkey_button.config(text=f"Stop Hotkey: [ {display_name.upper()} ]")
            print(f"Set Stop hotkey: {key_name}\n--------------------------")
            
            self.setting_hotkey = False
            self._hotkey_assignment_cancelled = True
            finish_hotkey_assignment()
            return

        def on_controller_button():
            if self._hotkey_assignment_cancelled or not self.setting_hotkey:
                return
                
            # Wait for controller button press
            button_name = self.controller_handler.wait_for_button_press(timeout=10)
            if button_name:
                key_name = f"controller_{button_name}"
                
                # Check if this controller button is already used by any area
                for area_frame, hotkey_button, _, area_name_var, _, _, _ in self.areas:
                    if hasattr(hotkey_button, 'hotkey') and hotkey_button.hotkey == key_name:
                        show_thinkr_warning(self, area_name_var.get())
                        self._hotkey_assignment_cancelled = True
                        self.setting_hotkey = False
                        self.stop_hotkey_button.config(text="Set Stop Hotkey")
                        finish_hotkey_assignment()
                        return
                
                # Remove existing stop hotkey if it exists
                if hasattr(self, 'stop_hotkey'):
                    try:
                        if hasattr(self.stop_hotkey_button, 'mock_button'):
                            self._cleanup_hooks(self.stop_hotkey_button.mock_button)
                    except Exception as e:
                        print(f"Error cleaning up stop hotkey hooks: {e}")
                
                self.stop_hotkey = key_name
                
                # Create a mock button object to use with setup_hotkey
                mock_button = type('MockButton', (), {'hotkey': key_name, 'is_stop_button': True})
                self.stop_hotkey_button.mock_button = mock_button
                
                # Setup the hotkey
                self.setup_hotkey(self.stop_hotkey_button.mock_button, None)
                
                display_name = f"Controller {button_name}"
                self.stop_hotkey_button.config(text=f"Stop Hotkey: [ {display_name.upper()} ]")
                print(f"Set Stop hotkey: {key_name}\n--------------------------")
                
                self.setting_hotkey = False
                self._hotkey_assignment_cancelled = True
                finish_hotkey_assignment()
                return
            else:
                # Timeout or no button pressed
                self.stop_hotkey_button.config(text="Set Stop Hotkey")
                self.setting_hotkey = False
                finish_hotkey_assignment()

        # Set button to indicate we're waiting for input
        self.stop_hotkey_button.config(text="Press any key or combination...")
        
        # Set up temporary hooks for key and mouse input
        try:
            # Store the hooks as attributes of the button for cleanup
            self.stop_hotkey_button.keyboard_hook_temp = keyboard.on_press(on_key_press, suppress=True)
            self.stop_hotkey_button.mouse_hook_temp = mouse.hook(on_mouse_click)
            
            # Live preview of currently held modifiers while waiting for a non-modifier key
            def _update_hotkey_preview():
                if self._hotkey_assignment_cancelled or not self.setting_hotkey:
                    return
                try:
                    mods = []
                    # Use scan code detection for more reliable left/right distinction
                    left_ctrl_pressed, right_ctrl_pressed = detect_ctrl_keys()
                    
                    if left_ctrl_pressed or right_ctrl_pressed: mods.append('CTRL')
                    if keyboard.is_pressed('shift'): mods.append('SHIFT')
                    if keyboard.is_pressed('left alt'): mods.append('L-ALT')
                    if keyboard.is_pressed('right alt'): mods.append('R-ALT')
                    if keyboard.is_pressed('left windows') or keyboard.is_pressed('right windows') or keyboard.is_pressed('windows'):
                        mods.append('WIN')
                    preview = " + ".join(mods)
                    if preview:
                        self.stop_hotkey_button.config(text=f"Press any key or combination... [ {preview} + ]")
                    else:
                        self.stop_hotkey_button.config(text="Press any key or combination...")
                except Exception:
                    pass
                # Schedule next update
                try:
                    self._hotkey_preview_job = self.root.after(80, _update_hotkey_preview)
                except Exception:
                    pass
            
            # Start live preview polling
            try:
                self._hotkey_preview_job = self.root.after(80, _update_hotkey_preview)
            except Exception:
                pass
            
            # Start controller monitoring for stop hotkey assignment if controller support is available
            if CONTROLLER_AVAILABLE:
                self._start_controller_stop_hotkey_monitoring(finish_hotkey_assignment)
        except Exception as e:
            print(f"Error setting up hotkey hooks: {e}")
            self.stop_hotkey_button.config(text="Set Stop Hotkey")
            self.setting_hotkey = False
            finish_hotkey_assignment()
            return
        
        # Also listen for Shift key release to assign LSHIFT/RSHIFT reliably
        def on_shift_release(_e):
            if self._hotkey_assignment_cancelled or not self.setting_hotkey:
                return
            if combo_state.get('non_modifier_pressed'):
                return
            side_label = 'left'
            try:
                raw = (getattr(_e, 'name', '') or '').lower()
                if 'right' in raw or 'right shift' in raw:
                    side_label = 'right'
            except Exception:
                pass
            key_name_local = f"{side_label} shift"
            _assign_stop_hotkey_and_register(key_name_local)

        try:
            self.stop_hotkey_button.shift_release_hooks = [
                keyboard.on_release_key('left shift', on_shift_release),
                keyboard.on_release_key('right shift', on_shift_release),
            ]
        except Exception:
            self.stop_hotkey_button.shift_release_hooks = []
        
        # Also listen for Ctrl key release to allow assigning bare CTRL reliably for stop hotkey
        def on_ctrl_release_stop(_e):
            if self._hotkey_assignment_cancelled or not self.setting_hotkey:
                return
            if combo_state.get('non_modifier_pressed'):
                return
            # Determine which ctrl key was released using scan code for reliability
            side_label = 'left'
            try:
                scan_code = getattr(_e, 'scan_code', None)
                if scan_code == 157:  # Right Ctrl scan code
                    side_label = 'right'
                elif scan_code == 29:  # Left Ctrl scan code
                    side_label = 'left'
                else:
                    # Fallback to event name if scan code is not available
                    raw = (getattr(_e, 'name', '') or '').lower()
                    if 'right' in raw or 'right ctrl' in raw:
                        side_label = 'right'
            except Exception:
                pass
            # Assign bare CTRL (no longer sided)
            key_name_local = "ctrl"
            _assign_stop_hotkey_and_register(key_name_local)

        try:
            self.stop_hotkey_button.ctrl_release_hooks = [
                keyboard.on_release_key('ctrl', on_ctrl_release_stop),
            ]
        except Exception:
            self.stop_hotkey_button.ctrl_release_hooks = []

        # Set a timer to reset the button if no key is pressed
        def reset_button():
            if not hasattr(self, 'stop_hotkey') or not self.stop_hotkey:
                self.stop_hotkey_button.config(text="Set Stop Hotkey")
            self._hotkey_assignment_cancelled = True
            self.setting_hotkey = False
            finish_hotkey_assignment()
            
        self.unhook_timer = self.root.after(5000, reset_button)

    def set_controller_stop_hotkey(self):
        """Set stop hotkey using controller button"""
        if not CONTROLLER_AVAILABLE:
            messagebox.showwarning("Controller Not Available", 
                                 "Controller support is not available. Please install the 'inputs' library.")
            return
            
        # Clean up temporary hooks and disable all hotkeys
        try:
            self.disable_all_hotkeys()
        except Exception as e:
            print(f"Warning: Error cleaning up temporary hooks: {e}")

        self._hotkey_assignment_cancelled = False
        self.setting_hotkey = True
        
        def finish_hotkey_assignment():
            try:
                self.stop_speaking()
                print("System reinitialized. Audio stopped.")
            except Exception as e:
                print(f"Error during forced stop: {e}")
            
            try:
                self.restore_all_hotkeys()
            except Exception as e:
                print(f"Warning: Error restoring hotkeys: {e}")
            
            self.setting_hotkey = False
            self.controller_hotkey_button.config(text="Controller")
        
        # Set button to indicate we're waiting for controller input
        self.controller_hotkey_button.config(text="Press controller button...")
        
        # Start controller monitoring in a separate thread
        def monitor_controller():
            try:
                button_name = self.controller_handler.wait_for_button_press(timeout=15)
                if button_name and not self._hotkey_assignment_cancelled:
                    key_name = f"controller_{button_name}"
                    
                    # Check if this controller button is already used by any area
                    for area_frame, hotkey_button, _, area_name_var, _, _, _ in self.areas:
                        if hasattr(hotkey_button, 'hotkey') and hotkey_button.hotkey == key_name:
                            show_thinkr_warning(self, area_name_var.get())
                            self._hotkey_assignment_cancelled = True
                            finish_hotkey_assignment()
                            return
                    
                    # Remove existing stop hotkey if it exists
                    if hasattr(self, 'stop_hotkey'):
                        try:
                            if hasattr(self.stop_hotkey_button, 'mock_button'):
                                self._cleanup_hooks(self.stop_hotkey_button.mock_button)
                        except Exception as e:
                            print(f"Error cleaning up stop hotkey hooks: {e}")
                    
                    self.stop_hotkey = key_name
                    
                    # Create a mock button object to use with setup_hotkey
                    mock_button = type('MockButton', (), {'hotkey': key_name, 'is_stop_button': True})
                    self.stop_hotkey_button.mock_button = mock_button
                    
                    # Setup the hotkey
                    self.setup_hotkey(self.stop_hotkey_button.mock_button, None)
                    
                    display_name = f"Controller {button_name}"
                    self.stop_hotkey_button.config(text=f"Stop Hotkey: [ {display_name.upper()} ]")
                    print(f"Set Stop hotkey: {key_name}\n--------------------------")
                    
                    finish_hotkey_assignment()
                else:
                    # Timeout or cancelled
                    self.controller_hotkey_button.config(text="Controller")
                    finish_hotkey_assignment()
            except Exception as e:
                print(f"Error in controller monitoring: {e}")
                self.controller_hotkey_button.config(text="Controller")
                finish_hotkey_assignment()
        
        # Start controller monitoring in background
        threading.Thread(target=monitor_controller, daemon=True).start()
        
        # Set a timer to reset if no button is pressed
        def reset_button():
            if not hasattr(self, 'stop_hotkey') or not self.stop_hotkey:
                self.controller_hotkey_button.config(text="Controller")
            self._hotkey_assignment_cancelled = True
            finish_hotkey_assignment()
            
        self.unhook_timer = self.root.after(15000, reset_button)

    def validate_numeric_input(self, P, is_speed=False):
        """Validate input to only allow numbers with different limits for speed and volume"""
        if P == "":  # Allow empty field
            return True
        # Only allow digits, no other characters
        if not P.isdigit():
            return False
        value = int(P)
        if is_speed:  # No upper limit for speed
            return value >= 0  # Only check that it's not negative
        else:  # For volume, keep 0-100 limit
            return 0 <= value <= 100

    def add_read_area(self, removable=True, editable_name=True, area_name="Area Name"):
        # Decide parent: place 'Auto Read' above the scroll area
        parent_container = self.area_frame
        if not removable and area_name == "Auto Read" and hasattr(self, 'auto_read_outer_frame'):
            parent_container = self.auto_read_outer_frame

        area_frame = tk.Frame(parent_container)
        area_frame.pack(pady=(4, 0), anchor='center')
        area_name_var = tk.StringVar(value=area_name)
        area_name_label = tk.Label(area_frame, textvariable=area_name_var)
        area_name_label.pack(side="left")
        
        # For Auto Read, never allow editing or right-click
        if editable_name and not (not removable and area_name == "Auto Read"):
            def prompt_edit_area_name(event=None):
                try:
                    self.disable_all_hotkeys()
                    new_name = tk.simpledialog.askstring("Edit Area Name", "Enter new area name:", initialvalue=area_name_var.get())
                    if new_name and new_name.strip():
                        area_name_var.set(new_name.strip())
                finally:
                    try:
                        self.restore_all_hotkeys()
                    except Exception as e:
                        print(f"Error restoring hotkeys after rename: {e}")
                self.resize_window()
            area_name_label.bind('<Button-3>', prompt_edit_area_name)  # Right-click to edit

        # Initialize the button first
        if not removable and area_name == "Auto Read":
            # set_area_button = tk.Button(area_frame, text="Select Area")
            # set_area_button.pack(side="left")
            set_area_button = None
        else:
            set_area_button = tk.Button(area_frame, text="Set Area")
            set_area_button.pack(side="left")
        # Add separator
        tk.Label(area_frame, text=" ⏐ ").pack(side="left")
        # Configure the command separately
        # Custom set_area_button command for Auto Read: only open selection overlay, never trigger reading directly
        if set_area_button is not None:
            set_area_button.config(command=partial(self.set_area, area_frame, area_name_var, set_area_button))

        # Always add hotkey button for all areas, including Auto Read
        hotkey_button = tk.Button(area_frame, text="Set Hotkey")
        hotkey_button.config(command=lambda: self.set_hotkey(hotkey_button, area_frame))
        hotkey_button.pack(side="left")
        
        # Add separator
        tk.Label(area_frame, text=" ⏐ ").pack(side="left")

        # Add Img. Processing button with checkbox
        customize_button = tk.Button(area_frame, text="Img. Processing...", command=partial(self.customize_processing, area_name_var))
        customize_button.pack(side="left")
        tk.Label(area_frame, text=" Enable:").pack(side="left")  # Label for the checkbox
        preprocess_var = tk.BooleanVar()
        preprocess_checkbox = tk.Checkbutton(area_frame, variable=preprocess_var)
        preprocess_checkbox.pack(side="left")
        # Add separator
        tk.Label(area_frame, text=" ⏐ ").pack(side="left")

        voice_var = tk.StringVar(value="Select Voice")
        # Get voice descriptions for the dropdown menu and create display names
        voice_names = []
        voice_display_names = []
        voice_full_names = {}  # Map display names to full names
        
        if hasattr(self, 'voices') and self.voices:
            try:
                for i, voice in enumerate(self.voices, 1):
                    full_name = voice.GetDescription()
                    voice_names.append(full_name)
                    
                    # Create abbreviated display name with numbering
                    if "Microsoft" in full_name and " - " in full_name:
                        # Format: "Microsoft David - en-US" -> "1. David (en-US)"
                        parts = full_name.split(" - ")
                        if len(parts) == 2:
                            voice_part = parts[0].replace("Microsoft ", "")
                            lang_part = parts[1]
                            display_name = f"{i}. {voice_part} ({lang_part})"
                        else:
                            display_name = f"{i}. {full_name}"
                    elif " - " in full_name:
                        # Format: "David - en-US" -> "1. David (en-US)"
                        parts = full_name.split(" - ")
                        if len(parts) == 2:
                            display_name = f"{i}. {parts[0]} ({parts[1]})"
                        else:
                            display_name = f"{i}. {full_name}"
                    else:
                        display_name = f"{i}. {full_name}"
                    
                    
                    voice_display_names.append(display_name)
                    voice_full_names[display_name] = full_name
            except Exception as e:
                print(f"Warning: Could not get voice descriptions: {e}")
                voice_names = []
                voice_display_names = []
        
        # Voice selection setup
        
        # Function to update the actual voice when display name is selected
        def on_voice_selection(*args):
            selected_display = voice_var.get()
            if selected_display in voice_full_names:
                # Store the full name for actual speech
                voice_var._full_name = voice_full_names[selected_display]
            else:
                voice_var._full_name = selected_display
        

        
        # Create the OptionMenu with display names and command
        voice_menu = tk.OptionMenu(
            area_frame, 
            voice_var,
            "Select Voice",
            *voice_display_names,
            command=on_voice_selection
        )
        # Set a fixed width to prevent layout issues when voice names change
        # This ensures the dropdown doesn't change size and push other elements around
        voice_menu.config(width=40)  # Fixed width that can accommodate most voice names
        
        # Configure the OptionMenu to display text left-aligned instead of centered
        # This prevents long names from being cut off on the sides
        voice_menu.config(anchor="w")  # "w" = west (left-aligned)
        
        voice_menu.pack(side="left")
        


        

        

        
        # Add separator
        tk.Label(area_frame, text=" ⏐ ").pack(side="left")

        speed_var = tk.StringVar(value="100")
        tk.Label(area_frame, text="Reading Speed % :").pack(side="left")
        vcmd = (self.root.register(lambda P: self.validate_numeric_input(P, is_speed=True)), '%P')
        speed_entry = tk.Entry(area_frame, textvariable=speed_var, width=5, validate='all', validatecommand=vcmd)
        speed_entry.pack(side="left")
        # Add separator
        tk.Label(area_frame, text=" ⏐ ").pack(side="left")
        
        speed_entry.bind('<Control-v>', lambda e: 'break')
        speed_entry.bind('<Control-V>', lambda e: 'break')
        speed_entry.bind('<Key>', lambda e: self.validate_speed_key(e, speed_var))

        if removable:
            remove_area_button = tk.Button(area_frame, text="Remove Area", command=lambda: self.remove_area(area_frame, area_name_var.get()))
            remove_area_button.pack(side="left")
            # Add separator
            tk.Label(area_frame, text="").pack(side="left")  # No symbol for last separator; empty label
        else:
            # Save button for Auto Read area
            # Add 'Stop read on select' checkbox to the left of Save button
            self.interrupt_on_new_scan_var = tk.BooleanVar(value=True)
            stop_read_checkbox = tk.Checkbutton(area_frame, text="Stop read on select", variable=self.interrupt_on_new_scan_var)
            stop_read_checkbox.pack(side="left", padx=(10, 2))
            def save_auto_read_settings():
                import tempfile, os, json
                # Find the hotkey for the Auto Read area
                hotkey = None
                for area in self.areas:
                    area_frame2, hotkey_button2, _, area_name_var2, _, _, _ = area
                    if area_frame2 == area_frame and area_name_var2.get() == "Auto Read":
                        hotkey = getattr(hotkey_button2, 'hotkey', None)
                        break
                # Start with the basic settings
                # Save the full voice name, not the display name
                voice_to_save = getattr(voice_var, '_full_name', voice_var.get())
                settings = {
                    'preprocess': preprocess_var.get(),
                    'voice': voice_to_save,
                    'speed': speed_var.get(),
                    'hotkey': hotkey,
                    'stop_read_on_select': self.interrupt_on_new_scan_var.get(),
                }
                # Add image processing settings if they exist
                if 'Auto Read' in self.processing_settings:
                    if 'processing' not in settings:
                        settings['processing'] = {}
                    settings['processing'].update(self.processing_settings['Auto Read'])
                # Create GameReader subdirectory in Temp if it doesn't exist
                game_reader_dir = os.path.join(tempfile.gettempdir(), 'GameReader')
                os.makedirs(game_reader_dir, exist_ok=True)
                temp_path = os.path.join(game_reader_dir, 'auto_read_settings.json')
                # Save the settings
                with open(temp_path, 'w') as f:
                    json.dump(settings, f, indent=4)
                # Show status message
                if hasattr(self, 'status_label'):
                    self.status_label.config(text="Auto Read area settings saved")
                    if hasattr(self, '_feedback_timer') and self._feedback_timer:
                        self.root.after_cancel(self._feedback_timer)
                    self._feedback_timer = self.root.after(2000, lambda: self.status_label.config(text=""))
            save_button = tk.Button(area_frame, text="Save", command=save_auto_read_settings)
            save_button.pack(side="left")
            tk.Label(area_frame, text="").pack(side="left")  # No symbol for last separator; empty label

        self.areas.append((area_frame, hotkey_button, set_area_button, area_name_var, preprocess_var, voice_var, speed_var))
        print("Added new read area.\n--------------------------")
        
        # Bind events to update window size live
        def bind_resize_events(widget):
            if isinstance(widget, tk.Entry):
                widget.bind('<KeyRelease>', lambda e: self.resize_window())
                widget.bind('<FocusOut>', lambda e: self.resize_window())
            if isinstance(widget, ttk.Combobox):
                widget.bind('<<ComboboxSelected>>', lambda e: self.resize_window())
            elif isinstance(widget, tk.OptionMenu):
                widget.bind('<<ComboboxSelected>>', lambda e: self.resize_window())
            widget.bind('<Configure>', lambda e: self.resize_window())
        for widget in area_frame.winfo_children():
            bind_resize_events(widget)
        area_frame.bind('<Configure>', lambda e: self.resize_window())

        # Call resize_window to ensure the window properly resizes when new areas are added
        self.resize_window(force=True)

    def remove_area(self, area_frame, area_name):
        # Find and clean up the hotkey for this area
        for area in self.areas:
            if area[0] == area_frame:  # Found matching frame
                hotkey_button = area[1]  # Get the hotkey button
                
                # Clean up keyboard hook if it exists
                if hasattr(hotkey_button, 'keyboard_hook') and hotkey_button.keyboard_hook:
                    try:
                        # Debug: Log the type of object we're dealing with
                        hook_type = type(hotkey_button.keyboard_hook).__name__
                        hook_value = str(hotkey_button.keyboard_hook)[:100]  # Limit length for logging
                        print(f"Cleaning up keyboard hook - Type: {hook_type}, Value: {hook_value}")
                        
                        # Check if this is a custom ctrl hook or a regular add_hotkey hook
                        if hasattr(hotkey_button.keyboard_hook, 'remove'):
                            # This is an add_hotkey hook
                            keyboard.remove_hotkey(hotkey_button.keyboard_hook)
                            print(f"Successfully removed hotkey-based keyboard hook")
                        else:
                            # This is a custom on_press hook
                            keyboard.unhook(hotkey_button.keyboard_hook)
                            print(f"Successfully unhooked custom keyboard hook")
                    except Exception as e:
                        print(f"Warning: Error cleaning up keyboard hook: {e}")
                    finally:
                        # Always set to None to prevent future errors
                        hotkey_button.keyboard_hook = None

                # Clean up mouse hook if it exists
                if hasattr(hotkey_button, 'mouse_hook'):
                    try:
                        # Only try to unhook if the hook exists and is not None
                        if hotkey_button.mouse_hook:
                            # Debug: Log the type of object we're dealing with
                            hook_type = type(hotkey_button.mouse_hook).__name__
                            hook_value = str(hotkey_button.mouse_hook)[:100]  # Limit length for logging
                            print(f"Cleaning up mouse hook - Type: {hook_type}, Value: {hook_value}")
                            
                            # Check if it's a hook ID
                            if hasattr(hotkey_button, 'mouse_hook_id') and hotkey_button.mouse_hook_id:
                                try:
                                    mouse.unhook(hotkey_button.mouse_hook_id)
                                    print(f"Successfully unhooked mouse hook ID")
                                except Exception:
                                    print(f"Failed to unhook mouse hook ID")
                                    pass
                            # Clean up the handler function reference
                            if hasattr(hotkey_button, 'mouse_hook'):
                                hotkey_button.mouse_hook = None
                    except Exception as e:
                        print(f"Warning: Error cleaning up mouse hook: {e}")
                    finally:
                        # Always set to None to prevent future errors
                        hotkey_button.mouse_hook = None
                        
                try:
                    self.latest_images[area_name].close()
                    del self.latest_images[area_name]
                except:
                    pass
        
        # Remove the area frame from the UI
        area_frame.destroy()
        # Remove the area from the list of areas
        self.areas = [area for area in self.areas if area[0] != area_frame]
        print(f"Removed area: {area_name}\n--------------------------")
        
        # Resize the window after removing an area to ensure proper sizing
        self.resize_window(force=True)

    def resize_window(self, force: bool = False):
        """Resize the window based on current content.
        If force is True, actively set the window geometry to fit the content (used after loading a layout)."""
        # Ensure positions/sizes are current
        self.root.update_idletasks()

        # Dynamically compute the non-scrollable portion height (everything above the areas canvas)
        try:
            base_top = self.area_canvas.winfo_rooty() - self.root.winfo_rooty()
        except Exception:
            base_top = 210
        base_height = max(150, base_top + 20)  # add small bottom margin
        min_width = 850
        max_width = 1000
        area_frame_height = 0
        if len(self.areas) > 0:
            self.area_frame.update_idletasks()
            area_frame_height = self.area_frame.winfo_height()
        # Determine total height needed for all current areas
        area_row_height = 60  # Approx row height (for fallback)
        # Number of areas in the scroll field (exclude the Auto Read area at index 0)
        num_scroll_areas = max(0, len(self.areas) - 1)
        content_height = area_frame_height if area_frame_height > 0 else num_scroll_areas * area_row_height
        total_height_unconstrained = base_height + content_height
        total_height = max(total_height_unconstrained, 250)
        # Screen-constrained maximum height
        try:
            screen_h = self.root.winfo_screenheight()
        except Exception:
            screen_h = 1000
        vertical_margin = 140  # Keep some space from screen edges
        max_allowed_height = max(300, screen_h - vertical_margin)
        # Decide if scrollbar should be shown either due to screen limit or explicit row limit (>9)
        show_scroll_due_to_count = num_scroll_areas > 9
        show_scroll_due_to_screen = total_height_unconstrained > max_allowed_height
        show_scroll = show_scroll_due_to_count or show_scroll_due_to_screen

        # Compute target height of the window
        if show_scroll_due_to_count:
            # Cap visible rows to 9 when there are more than 9 areas
            visible_rows = 9
            # Do not grow the window further when crossing the threshold; keep current height or smaller cap
            cur_h_for_cap = self.root.winfo_height()
            desired_height_cap = min(base_height + visible_rows * area_row_height, max_allowed_height)
            target_height = min(cur_h_for_cap, desired_height_cap)
        else:
            # Otherwise try to fit all content within the screen
            target_height = min(total_height_unconstrained, max_allowed_height)
        
        # Determine the widest area
        widest = min_width
        for area in self.areas:
            frame = area[0]
            frame.update_idletasks()
            frame_left = frame.winfo_rootx()
            farthest_right = frame_left
            for child in frame.winfo_children():
                child.update_idletasks()
                child_right = child.winfo_rootx() + child.winfo_width()
                if child_right > farthest_right:
                    farthest_right = child_right
            area_width = farthest_right - frame_left
            if area_width > widest:
                widest = area_width
        widest += 60
        window_width = max(min_width, min(max_width, widest))
        
        # Apply scrollbar logic based on whether all content fits vertically
        if hasattr(self, 'area_scrollbar') and hasattr(self, 'area_canvas'):
            if show_scroll:
                # Need scrolling
                self.area_scrollbar.pack(side='right', fill='y')
                self.area_canvas.configure(yscrollcommand=self.area_scrollbar.set)
                if show_scroll_due_to_count:
                    canvas_height = max(100, min(target_height - base_height, 9 * area_row_height))
                else:
                    canvas_height = max(100, target_height - base_height)
                self.area_canvas.config(height=canvas_height)
            else:
                # All content fits; no scrollbar
                self.area_scrollbar.pack_forget()
                # Expand canvas to show all content when it fits
                self.area_canvas.config(height=area_frame_height)
        
        # Set minimums (use a constant min width so user can resize horizontally).
        # Ensure minimum width is sufficient to keep the single-line options from truncating.
        min_required_width = max(min_width, 1125) 
        self.root.minsize(min_required_width, 250)
        
        # Optionally force window geometry (used when loading a layout)
        cur_width = self.root.winfo_width()
        cur_height = self.root.winfo_height()
        if force:
            # To ensure Tk applies the new size reliably even when shrinking, call geometry twice
            self.root.geometry(f"{window_width}x{target_height}")
            self.root.update_idletasks()
            self.root.geometry(f"{window_width}x{target_height}")

        self.root.update_idletasks()  # Ensure geometry is applied


    
    def set_area(self, frame, area_name_var, set_area_button):
        # Check if another area selection is already in progress
        if hasattr(self, 'area_selection_in_progress') and self.area_selection_in_progress:
            print("Another area selection is already in progress. Please wait for it to complete.")
            if hasattr(self, 'status_label'):
                self.status_label.config(text="⚠️  Another area selection is already in progress")
                if hasattr(self, '_feedback_timer') and self._feedback_timer:
                    self.root.after_cancel(self._feedback_timer)
                self._feedback_timer = self.root.after(3000, lambda: self.status_label.config(text=""))
            return
        
        # Mark that area selection is now in progress
        self.area_selection_in_progress = True
        
        x1, y1, x2, y2 = 0, 0, 0, 0
        selection_cancelled = False
        
        # Store the current mouse hooks to restore them later
        self.saved_mouse_hooks = []
        if hasattr(self, 'mouse_hooks') and self.mouse_hooks:
            self.saved_mouse_hooks = self.mouse_hooks.copy()
        
        # --- Disable all hotkeys before starting area selection ---
        try:
            # Only unhook keyboard hotkeys, leave mouse hooks alone
            keyboard.unhook_all()
            # Clear the mouse hooks list but don't unhook them yet
            if hasattr(self, 'mouse_hooks'):
                self.mouse_hooks.clear()
        except Exception as e:
            print(f"Error disabling hotkeys for area selection: {e}")
        
        self.hotkeys_disabled_for_selection = True

        def on_drag(event):
            if not selection_cancelled:
                # Only allow interaction if window is ready
                if not hasattr(select_area_window, 'window_ready') or not select_area_window.window_ready:
                    return
                    
                # Use event coordinates directly for canvas drawing
                canvas_x = event.x
                canvas_y = event.y
                
                # Update both rectangles with canvas coordinates
                coords = (
                    min(canvas_x, x1), 
                    min(canvas_y, y1),
                    max(canvas_x, x1), 
                    max(canvas_y, y1)
                )
                
                # Update both rectangles
                canvas.coords(border, *coords)
                canvas.coords(border_outline, *coords)
                
                # Debug: Show current drag coordinates (only print occasionally to avoid spam)
                if hasattr(on_drag, 'last_debug_time'):
                    if time.time() - on_drag.last_debug_time > 0.5:  # Print every 0.5 seconds
                        print(f"Debug: Dragging - Current: ({canvas_x}, {canvas_y}), Start: ({x1}, {y1})")
                        on_drag.last_debug_time = time.time()
                else:
                    on_drag.last_debug_time = time.time()

        def on_click(event):
            nonlocal x1, y1
            # Only allow interaction if window is ready
            if not hasattr(select_area_window, 'window_ready') or not select_area_window.window_ready:
                print("Debug: Ignoring click - window not ready yet")
                return
                
            # Store canvas coordinates
            x1 = event.x
            y1 = event.y
            print(f"Debug: Mouse click - Canvas coordinates: ({x1}, {y1})")
            canvas.bind("<B1-Motion>", on_drag)
            canvas.bind("<ButtonRelease-1>", on_release)
            # Initialize both rectangles at click point
            canvas.coords(border, x1, y1, x1, y1)
            canvas.coords(border_outline, x1, y1, x1, y1)

        def on_release(event):
            nonlocal x1, y1, x2, y2
            if not selection_cancelled:
                # Only allow interaction if window is ready
                if not hasattr(select_area_window, 'window_ready') or not select_area_window.window_ready:
                    print("Debug: Ignoring release - window not ready yet")
                    return
                    
                try:
                    # Stop speech on mouse release if the checkbox is checked
                    if hasattr(self, 'interrupt_on_new_scan_var') and self.interrupt_on_new_scan_var.get():
                        self.stop_speaking()
                    
                    # Convert canvas coordinates to screen coordinates for the final area
                    # Canvas coordinates are relative to the selection window, which is positioned at (window_x, window_y)
                    # But we need to convert to actual screen coordinates using the original min_x, min_y
                    x2 = event.x + min_x  # Convert canvas to screen coordinates
                    y2 = event.y + min_y
                    x1_screen = x1 + min_x
                    y1_screen = y1 + min_y
                    
                    print(f"Debug: Mouse release - Canvas: ({event.x}, {event.y}), Screen: ({x2}, {y2}), Start: ({x1_screen}, {y1_screen})")
                    
                    # Only set coordinates if we have a valid selection (not a click)
                    # Check minimum drag distance using canvas coordinates for consistency
                    if abs(event.x - x1) > 5 and abs(event.y - y1) > 5:  # Minimum 5px drag
                        final_coords = (
                            min(x1_screen, x2), 
                            min(y1_screen, y2),
                            max(x1_screen, x2), 
                            max(y1_screen, y2)
                        )
                        frame.area_coords = final_coords
                        print(f"Debug: Area selection coordinates - Canvas: ({x1}, {y1}), Screen: ({x1_screen}, {y1_screen}), Final: {final_coords}")
                    else:
                        # If it's just a click, don't update the coordinates
                        frame.area_coords = getattr(frame, 'area_coords', (0, 0, 0, 0))
                    
                    # If this is the Auto Read area, trigger reading immediately and keep button label as 'Select Area'
                    is_auto_read = hasattr(area_name_var, 'get') and area_name_var.get() == "Auto Read"
                    
                    # Release grabs/bindings before destroying the overlay
                    try:
                        select_area_window.grab_release()
                    except Exception:
                        pass
                    try:
                        self.root.unbind_all("<Escape>")
                    except Exception:
                        pass
                    try:
                        canvas.unbind("<Button-1>")
                        canvas.unbind("<B1-Motion>")
                        canvas.unbind("<Escape>")
                    except Exception:
                        pass
                    # Destroy the selection window to restore normal mouse handling
                    select_area_window.destroy()
                    
                    if is_auto_read:
                        # Read after a short delay so overlay is gone
                        self.root.after(100, lambda: self.read_area(frame))
                    
                    # Only prompt for name if it's not Auto Read and has default name
                    current_name = area_name_var.get()
                    if not is_auto_read:
                        if current_name == "Area Name":
                            # Create custom dialog that stays on top and auto-focuses
                            area_name = self._create_area_name_dialog()
                            if area_name and area_name.strip():
                                area_name_var.set(area_name)
                                print(f"Set area: {frame.area_coords} with name {area_name_var.get()}\n--------------------------")
                            else:
                                messagebox.showerror("Error", "Area name cannot be empty.")
                                print("Error: Area name cannot be empty, Using Default: Area Name")
                                # Still restore hotkeys even if name is invalid
                                self._restore_hotkeys_after_selection()
                                return
                        else:
                            print(f"Updated area: {frame.area_coords} with existing name {current_name}\n--------------------------")
                        if set_area_button is not None:
                            set_area_button.config(text="Edit Area")
                    else:
                        # Always keep button label as 'Select Area' for Auto Read
                        if set_area_button is not None:
                            set_area_button.config(text="Select Area")
                    
                    # Mark that we have unsaved changes
                    self._set_unsaved_changes()
                    
                except Exception as e:
                    print(f"Error during area selection: {e}")
                finally:
                    # Always ensure hotkeys are restored
                    self._restore_hotkeys_after_selection()

        def on_escape(event):
            nonlocal selection_cancelled
            selection_cancelled = True
            if not hasattr(frame, 'area_coords'):
                frame.area_coords = (0, 0, 0, 0)
            
            # Release grabs/bindings before destroying the overlay
            try:
                select_area_window.grab_release()
            except Exception:
                pass
            try:
                self.root.unbind_all("<Escape>")
            except Exception:
                pass
            try:
                canvas.unbind("<Button-1>")
                canvas.unbind("<B1-Motion>")
                canvas.unbind("<Escape>")
            except Exception:
                pass
            # Destroy the selection window to restore normal mouse handling
            select_area_window.destroy()
            
            # Use our helper method to ensure consistent hotkey restoration
            self._restore_hotkeys_after_selection()
            print("Area selection cancelled\n--------------------------")

        # Create fullscreen window that spans all monitors
        select_area_window = tk.Toplevel(self.root)
        select_area_window.overrideredirect(True)
        
        # Add a protocol handler to reset the flag if the window is destroyed unexpectedly
        def on_window_destroy():
            if hasattr(self, 'area_selection_in_progress'):
                self.area_selection_in_progress = False
            # Ensure hotkeys are restored
            self._restore_hotkeys_after_selection()
        
        select_area_window.protocol("WM_DELETE_WINDOW", on_window_destroy)
        
        # Get the true multi-monitor dimensions using win32api.GetSystemMetrics
        # This ensures consistency with capture_screen_area function
        min_x = win32api.GetSystemMetrics(win32con.SM_XVIRTUALSCREEN)  # Leftmost x (can be negative)
        min_y = win32api.GetSystemMetrics(win32con.SM_YVIRTUALSCREEN)  # Topmost y (can be negative)
        virtual_width = win32api.GetSystemMetrics(win32con.SM_CXVIRTUALSCREEN)
        virtual_height = win32api.GetSystemMetrics(win32con.SM_CYVIRTUALSCREEN)
        max_x = min_x + virtual_width
        max_y = min_y + virtual_height
        
        print(f"Debug: Area selection - Virtual screen bounds: ({min_x}, {min_y}, {max_x}, {max_y})")
        print(f"Debug: Area selection - Window size: {virtual_width}x{virtual_height}")
        
        # Set window to cover entire virtual screen
        # Use the actual virtual screen coordinates, even if negative
        # Windows should handle negative coordinates for multi-monitor setups
        select_area_window.geometry(f"{virtual_width}x{virtual_height}+{min_x}+{min_y}")
        
        print(f"Debug: Window positioned at ({min_x}, {min_y}) with size {virtual_width}x{virtual_height}")
        
        # Create canvas first
        canvas = tk.Canvas(select_area_window, 
                          cursor="cross",
                          width=virtual_width,
                          height=virtual_height,
                          highlightthickness=0,
                          bg='white')
        canvas.pack(fill="both", expand=True)
        
        # Set window properties
        select_area_window.attributes("-alpha", 0.5)  
        select_area_window.attributes("-topmost", True)  # Keep window on top
        
        # Force the window to be positioned correctly with multiple attempts
        def ensure_proper_positioning():
            try:
                select_area_window.update_idletasks()
                select_area_window.lift()
                select_area_window.focus_force()
                
                # Check if positioning worked
                actual_x = select_area_window.winfo_x()
                actual_y = select_area_window.winfo_y()
                
                if abs(actual_x - min_x) > 10 or abs(actual_y - min_y) > 10:
                    print(f"Debug: Window positioning failed, retrying... Expected: ({min_x}, {min_y}), Got: ({actual_x}, {actual_y})")
                    # Try to reposition
                    select_area_window.geometry(f"{virtual_width}x{virtual_height}+{min_x}+{min_y}")
                    select_area_window.update_idletasks()
                    
                    # Check again
                    actual_x = select_area_window.winfo_x()
                    actual_y = select_area_window.winfo_y()
                    print(f"Debug: After retry - Position: ({actual_x}, {actual_y})")
                else:
                    print(f"Debug: Window positioned correctly at ({actual_x}, {actual_y})")
                    
            except Exception as e:
                print(f"Debug: Error ensuring proper positioning: {e}")
        
        # Ensure proper positioning with a delay
        self.root.after(100, ensure_proper_positioning)
        
        # Create border rectangle with more visible red border
        border = canvas.create_rectangle(0, 0, 0, 0,
                                       outline='red',
                                       width=3,  # Increased width
                                       dash=(8, 4))  # Longer dashes, shorter gaps
        
        # Wait for proper positioning before binding events
        def bind_events_after_positioning():
            try:
                # Only bind events if window is properly positioned
                actual_x = select_area_window.winfo_x()
                actual_y = select_area_window.winfo_y()
                
                if abs(actual_x - min_x) <= 10 and abs(actual_y - min_y) <= 10:
                    print("Debug: Binding events - window properly positioned")
                    # Bind events
                    canvas.bind("<Button-1>", on_click)
                    canvas.bind("<Escape>", on_escape)
                    select_area_window.bind("<Escape>", on_escape)
                    # Capture Escape at the application level to ensure it works even if focus is lost
                    try:
                        self.root.bind_all("<Escape>", on_escape)
                    except Exception:
                        pass
                    
                    # Add focus and key bindings
                    select_area_window.focus_force()
                    # Grab all events so Escape is reliably received
                    try:
                        select_area_window.grab_set()
                    except Exception:
                        pass
                    select_area_window.bind("<FocusOut>", lambda e: select_area_window.focus_force())
                    select_area_window.bind("<Key>", lambda e: on_escape(e) if e.keysym == "Escape" else None)
                    
                    # Mark that the window is ready for interaction
                    select_area_window.window_ready = True
                    print("Debug: Area selection window is ready for interaction")
                    
                else:
                    print(f"Debug: Window not properly positioned yet, retrying... ({actual_x}, {actual_y}) vs ({min_x}, {min_y})")
                    # Retry after a short delay
                    self.root.after(50, bind_events_after_positioning)
                    
            except Exception as e:
                print(f"Debug: Error binding events: {e}")
        
        # Bind events after positioning is confirmed
        self.root.after(150, bind_events_after_positioning)
        
        # Add timeout to prevent hanging if window never gets positioned
        def timeout_handler():
            if not hasattr(select_area_window, 'window_ready') or not select_area_window.window_ready:
                print("Debug: Timeout - forcing window to be ready")
                select_area_window.window_ready = True
                # Force bind events
                try:
                    canvas.bind("<Button-1>", on_click)
                    canvas.bind("<Escape>", on_escape)
                    select_area_window.bind("<Escape>", on_escape)
                    self.root.bind_all("<Escape>", on_escape)
                    select_area_window.focus_force()
                    select_area_window.grab_set()
                    select_area_window.bind("<FocusOut>", lambda e: select_area_window.focus_force())
                    select_area_window.bind("<Key>", lambda e: on_escape(e) if e.keysym == "Escape" else None)
                    print("Debug: Events bound after timeout")
                except Exception as e:
                    print(f"Debug: Error binding events after timeout: {e}")
        
        # Set timeout to 2 seconds
        self.root.after(2000, timeout_handler)
        
        # Create second border for better visibility
        border_outline = canvas.create_rectangle(0, 0, 0, 0,
                                          outline='red',
                                          width=3,
                                          dash=(8, 4),
                                          dashoffset=6)  # Offset to create alternating pattern

    def _restore_hotkeys_after_selection(self):
        """Helper method to restore hotkeys after area selection"""
        if not hasattr(self, 'hotkeys_disabled_for_selection') or not self.hotkeys_disabled_for_selection:
            return
            
        try:
            self.restore_all_hotkeys()
            self.hotkeys_disabled_for_selection = False
            print("Hotkeys re-enabled after area selection")
            
            # Reset the area selection flag
            if hasattr(self, 'area_selection_in_progress'):
                self.area_selection_in_progress = False
            
            # Force focus back to the main window
            if hasattr(self, 'root') and self.root.winfo_exists():
                self.root.focus_force()
                
        except Exception as e:
            print(f"Error restoring hotkeys after area selection: {e}")
            # Ensure the flags are cleared even if there's an error
            self.hotkeys_disabled_for_selection = False
            if hasattr(self, 'area_selection_in_progress'):
                self.area_selection_in_progress = False

    def _create_area_name_dialog(self):
        """Create a custom dialog for naming areas that stays on top and auto-focuses"""
        # Create the dialog window
        dialog = tk.Toplevel(self.root)
        dialog.title("Area Name")
        dialog.geometry("250x120")
        dialog.resizable(False, False)
        
        # Make dialog stay on top of main window
        dialog.transient(self.root)
        dialog.grab_set()  # Make it modal
        
        # Center the dialog on the main window
        dialog.geometry("+%d+%d" % (
            self.root.winfo_rootx() + self.root.winfo_width()//2 - 150,
            self.root.winfo_rooty() + self.root.winfo_height()//2 - 60))
        
        # Set the window icon
        try:
            icon_path = os.path.join(os.path.dirname(__file__), 'Assets', 'icon.ico')
            if os.path.exists(icon_path):
                dialog.iconbitmap(icon_path)
        except Exception as e:
            print(f"Error setting dialog icon: {e}")
        
        # Create and pack the label
        label = tk.Label(dialog, text="Enter a name for this area:", pady=10)
        label.pack()
        
        # Create and pack the entry field
        entry = tk.Entry(dialog, width=30)
        entry.pack(pady=10)
        
        # Create button frame
        button_frame = tk.Frame(dialog)
        button_frame.pack(pady=10)
        
        # Variable to store the result
        result = [None]
        
        def on_ok():
            result[0] = entry.get()
            dialog.destroy()
        
        def on_enter(event):
            on_ok()
        
        # Bind Enter key
        entry.bind('<Return>', on_enter)
        
        # Create OK button only
        ok_button = tk.Button(button_frame, text="OK", command=on_ok, width=8)
        ok_button.pack()
        
        # Pre-fill with "Area Name" and focus the entry field
        entry.insert(0, "Area Name")
        entry.focus_set()
        entry.select_range(0, tk.END)
        
        # Wait for the dialog to close
        dialog.wait_window()
        
        return result[0]

    def disable_all_hotkeys(self):
        """Disable all hotkeys for keyboard, mouse, and controller."""
        try:
            # Unhook all keyboard and mouse hooks
            keyboard.unhook_all()
            mouse.unhook_all()
            
            # Stop controller monitoring to disable controller hotkeys
            if hasattr(self, 'controller_handler') and self.controller_handler:
                self.controller_handler.stop_monitoring()
            
            # Only clear lists if they exist and are not empty
            if hasattr(self, 'keyboard_hooks') and self.keyboard_hooks:
                self.keyboard_hooks.clear()
            if hasattr(self, 'mouse_hooks') and self.mouse_hooks:
                self.mouse_hooks.clear()
            if hasattr(self, 'hotkeys') and self.hotkeys:
                self.hotkeys.clear()
            
            # Reset hotkey setting state
            self.setting_hotkey = False
            
        except Exception as e:
            print(f"Warning: Error during hotkey cleanup: {e}")
            print(f"Current state - keyboard_hooks: {len(getattr(self, 'keyboard_hooks', []))}")
            print(f"Current state - mouse_hooks: {len(getattr(self, 'mouse_hooks', []))}")
            print(f"Current state - hotkeys: {len(getattr(self, 'hotkeys', []))}")
            # Don't fail the entire operation if cleanup fails

    def unhook_mouse(self):
        try:
            # Only attempt to unhook and clear if mouse_hooks exists and has items
            if hasattr(self, 'mouse_hooks') and self.mouse_hooks:
                mouse.unhook_all()
                self.mouse_hooks.clear()
        except Exception as e:
            print(f"Warning: Error during mouse hook cleanup: {e}")
            print(f"Mouse hooks list state: {len(getattr(self, 'mouse_hooks', []))}")

    def restore_all_hotkeys(self):
        """Restore all area and stop hotkeys after area selection is finished/cancelled."""
        # First, clean up any existing hooks
        try:
            keyboard.unhook_all()
            if hasattr(self, 'mouse_hooks'):
                self.mouse_hooks.clear()
        except Exception as e:
            print(f"Error cleaning up hooks during restore: {e}")
        
        # Restore the saved mouse hooks
        if hasattr(self, 'saved_mouse_hooks'):
            for hook in self.saved_mouse_hooks:
                try:
                    mouse.hook(hook)
                    self.mouse_hooks.append(hook)
                except Exception as e:
                    print(f"Error restoring mouse hook: {e}")
            # Clean up the saved hooks
            delattr(self, 'saved_mouse_hooks')
        
        # Re-register all hotkeys for areas
        for area_tuple in getattr(self, 'areas', []):
            area_frame, hotkey_button, _, area_name_var, _, _, _ = area_tuple
            if hasattr(hotkey_button, 'hotkey'):
                try:
                    self.setup_hotkey(hotkey_button, area_frame)
                except Exception as e:
                    print(f"Error re-registering hotkey: {e}")
        
        # Re-register stop hotkey if it exists
        if hasattr(self, 'stop_hotkey_button') and hasattr(self.stop_hotkey_button, 'mock_button'):
            try:
                self.setup_hotkey(self.stop_hotkey_button.mock_button, None)
            except Exception as e:
                print(f"Error re-registering stop hotkey: {e}")
        
        # Restart controller monitoring if there are any controller hotkeys
        if hasattr(self, 'controller_handler') and self.controller_handler:
            # Check if any areas have controller hotkeys
            has_controller_hotkeys = False
            for area_tuple in getattr(self, 'areas', []):
                area_frame, hotkey_button, _, area_name_var, _, _, _ = area_tuple
                if hasattr(hotkey_button, 'hotkey') and hotkey_button.hotkey and hotkey_button.hotkey.startswith('controller_'):
                    has_controller_hotkeys = True
                    break
            
            # Also check if stop hotkey is a controller hotkey
            if hasattr(self, 'stop_hotkey') and self.stop_hotkey and self.stop_hotkey.startswith('controller_'):
                has_controller_hotkeys = True
            
            # Start controller monitoring if needed
            if has_controller_hotkeys and not self.controller_handler.running:
                self.controller_handler.start_monitoring()

    def set_hotkey(self, button, area_frame):
        # Clean up temporary hooks and disable all hotkeys
        try:
            if hasattr(button, 'keyboard_hook_temp'):
    
                delattr(button, 'keyboard_hook_temp')
            
                if hasattr(button, 'mouse_hook_temp'):
                    try:
                        mouse.unhook(button.mouse_hook_temp)
                    except Exception:
                        pass
                    delattr(button, 'mouse_hook_temp')
            
            self.disable_all_hotkeys()
        except Exception as e:
            print(f"Warning: Error cleaning up temporary hooks: {e}")

        self._hotkey_assignment_cancelled = False  # Guard flag to block late events
        self.setting_hotkey = True
        print(f"Hotkey assignment mode started for button: {button}")

        def finish_hotkey_assignment():
            # --- Re-enable all hotkeys after hotkey assignment is finished/cancelled ---
            try:
                self.restore_all_hotkeys()
            except Exception as e:
                print(f"Warning: Error restoring hotkeys: {e}")
            # Stop live preview updater if running
            try:
                if hasattr(self, '_hotkey_preview_job') and self._hotkey_preview_job:
                    self.root.after_cancel(self._hotkey_preview_job)
                    self._hotkey_preview_job = None
            except Exception:
                pass

        # Track whether a non-modifier key was pressed during capture (to distinguish bare modifiers)
        combo_state = {'non_modifier_pressed': False, 'held_modifiers': set()}

        # Live preview of currently held modifiers while waiting for a non-modifier key
        def _update_hotkey_preview():
            if self._hotkey_assignment_cancelled or not self.setting_hotkey:
                return
            try:
                mods = []
                # Use scan code detection for more reliable left/right distinction
                left_ctrl_pressed, right_ctrl_pressed = detect_ctrl_keys()
                
                if left_ctrl_pressed or right_ctrl_pressed: mods.append('CTRL')
                if keyboard.is_pressed('shift'): mods.append('SHIFT')
                if keyboard.is_pressed('left alt'): mods.append('L-ALT')
                if keyboard.is_pressed('right alt'): mods.append('R-ALT')
                if keyboard.is_pressed('left windows') or keyboard.is_pressed('right windows') or keyboard.is_pressed('windows'):
                    mods.append('WIN')
                preview = " + ".join(mods)
                if preview:
                    button.config(text=f"Set Hotkey: [ {preview} + ]")
                else:
                    button.config(text=f"Set Hotkey: [  ]")
            except Exception:
                pass
            # Schedule next update
            try:
                self._hotkey_preview_job = self.root.after(80, _update_hotkey_preview)
            except Exception:
                pass

        def on_key_press(event):
            if self._hotkey_assignment_cancelled or not self.setting_hotkey:
                return
            # Ignore Escape
            if event.scan_code == 1:
                return
            
            print(f"Key press event received: {event.name} (type: {type(event).__name__})")
            
            # Use scan code and virtual key code for consistent behavior across keyboard layouts
            scan_code = getattr(event, 'scan_code', None)
            vk_code = getattr(event, 'vk_code', None)
            
            # Get current keyboard layout for consistency
            current_layout = get_current_keyboard_layout()
            
            # Determine key name based on scan code and virtual key code for consistency
            name = None
            side = None
            
            # Handle modifier keys consistently
            if scan_code == 29:  # Left Ctrl
                name = 'ctrl'
                side = 'left'
            elif scan_code == 157:  # Right Ctrl
                name = 'ctrl'
                side = 'right'
            elif scan_code == 42:  # Left Shift
                name = 'shift'
                side = 'left'
            elif scan_code == 54:  # Right Shift
                name = 'shift'
                side = 'right'
            elif scan_code == 56:  # Left Alt
                name = 'left alt'
                side = 'left'
            elif scan_code == 184:  # Right Alt
                name = 'right alt'
                side = 'right'
            elif scan_code == 91:  # Left Windows
                name = 'windows'
                side = 'left'
            elif scan_code == 92:  # Right Windows
                name = 'windows'
                side = 'right'
            else:
                # For non-modifier keys, use the event name but normalize it
                raw_name = (event.name or '').lower()
                name = normalize_key_name(raw_name)
                
                # Enhanced support for function keys and special keys
                if scan_code in self.numpad_scan_codes:
                    sym = self.numpad_scan_codes[scan_code]
                    name = f"num_{sym}"
                elif scan_code in [59, 60, 61, 62, 63, 64, 65, 66, 67, 68, 87, 88]:  # F1-F12
                    f_num = scan_code - 58 if scan_code <= 68 else scan_code - 76
                    name = f"f{f_num}"
                elif scan_code == 69:  # Num Lock
                    name = "num lock"
                elif scan_code == 70:  # Scroll Lock
                    name = "scroll lock"
                elif scan_code == 83:  # Insert
                    name = "insert"
                elif scan_code == 71:  # Home
                    name = "home"
                elif scan_code == 79:  # End
                    name = "end"
                elif scan_code == 73:  # Page Up
                    name = "page up"
                elif scan_code == 81:  # Page Down
                    name = "page down"
                elif scan_code == 82:  # Delete
                    name = "delete"
                elif scan_code == 15:  # Tab
                    name = "tab"
                elif scan_code == 28:  # Enter
                    name = "enter"
                elif scan_code == 14:  # Backspace
                    name = "backspace"
                elif scan_code == 57:  # Space
                    name = "space"

            # Debug: Show what name was determined
            print(f"Debug: Final determined name: '{name}' (scan code: {scan_code})")
            if scan_code in [29, 157]:
                print(f"Debug: Ctrl key detection - scan code {scan_code} -> '{name}'")
                if scan_code in [29, 157] and name != 'ctrl':
                    print(f"ERROR: Ctrl scan code {scan_code} detected but name is '{name}' instead of 'ctrl'")
            
            # Track modifiers as they're pressed and mark non-modifiers
            if name not in ('ctrl','alt','left alt','right alt','shift','windows'):
                combo_state['non_modifier_pressed'] = True
                print(f"Debug: Non-modifier key detected: '{name}'")
            else:
                # Add modifier to our tracking set
                combo_state['held_modifiers'].add(name)
                print(f"Debug: Modifier key detected: '{name}', held modifiers: {combo_state['held_modifiers']}")
            if name in ('ctrl', 'shift', 'alt', 'left alt', 'right alt', 'windows'):
                # Allow assigning a bare modifier when released, if user doesn't press another key
                # Start a short timer to check if still only this modifier is held
                def _assign_bare_modifier(modifier_name):
                    if self._hotkey_assignment_cancelled or not self.setting_hotkey:
                        return
                    try:
                        held = []
                        # Use scan code detection for more reliable left/right distinction
                        left_ctrl_pressed, right_ctrl_pressed = detect_ctrl_keys()
                        
                        if left_ctrl_pressed or right_ctrl_pressed: held.append('ctrl')
                        if keyboard.is_pressed('shift'): held.append('shift')
                        if keyboard.is_pressed('left alt'): held.append('left alt')
                        if keyboard.is_pressed('right alt'): held.append('right alt')
                        if keyboard.is_pressed('left windows') or keyboard.is_pressed('right windows') or keyboard.is_pressed('windows'):
                            held.append('windows')
                        # Only proceed if exactly one modifier is still held and matches side/base
                        if len(held) == 1:
                            only = held[0]
                            # Determine base from modifier_name
                            base = None
                            if 'ctrl' in modifier_name: base = 'ctrl'
                            elif 'alt' in modifier_name: base = 'alt'
                            elif 'shift' in modifier_name: base = 'shift'
                            elif 'windows' in modifier_name: base = 'windows'
                            
                            # Accept if same base and, when available, same side
                            if (base == 'ctrl' and only == 'ctrl') or \
                               (base == 'alt' and (only in ['left alt','right alt'])) or \
                               (base == 'shift' and only == 'shift') or \
                               (base == 'windows' and only == 'windows'):
                                key_name = only
                            else:
                                return

                            # Prevent duplicates: Stop hotkey
                            if getattr(self, 'stop_hotkey', None) == key_name:
                                self.setting_hotkey = False
                                self._hotkey_assignment_cancelled = True
                                try:
                                    if hasattr(button, 'keyboard_hook_temp'):
                                        keyboard.unhook(button.keyboard_hook_temp)
                                        delattr(button, 'keyboard_hook_temp')
                                    if hasattr(button, 'mouse_hook_temp'):
                                        mouse.unhook(button.mouse_hook_temp)
                                        delattr(button, 'mouse_hook_temp')
                                except Exception:
                                    pass
                                finish_hotkey_assignment()
                                try:
                                    messagebox.showwarning("Hotkey In Use", "This hotkey is already assigned to: Stop Hotkey")
                                except Exception:
                                    pass
                                return

                            # Prevent duplicates: other areas
                            for area in self.areas:
                                if area[1] is not button and hasattr(area[1], 'hotkey') and area[1].hotkey == key_name:
                                    self.setting_hotkey = False
                                    self._hotkey_assignment_cancelled = True
                                    try:
                                        if hasattr(button, 'keyboard_hook_temp'):
                                            keyboard.unhook(button.keyboard_hook_temp)
                                            delattr(button, 'keyboard_hook_temp')
                                        if hasattr(button, 'mouse_hook_temp'):
                                            mouse.unhook(button.mouse_hook_temp)
                                            delattr(button, 'mouse_hook_temp')
                                    except Exception:
                                        pass
                                    finish_hotkey_assignment()
                                    area_name = area[3].get() if hasattr(area[3], 'get') else "Unknown Area"
                                    show_thinkr_warning(self, area_name)
                                    return

                            button.hotkey = key_name
                            # Display mapping
                            disp = key_name.upper().replace('LEFT ALT','L-ALT').replace('RIGHT ALT','R-ALT') \
                                                .replace('WINDOWS','WIN').replace('CTRL','CTRL')
                            display_name = disp
                            button.config(text=f"Set Hotkey: [ {display_name} ]")
                            self.setup_hotkey(button, area_frame)
                            # Cleanup temp hooks and preview
                            try:
                                if hasattr(button, 'keyboard_hook_temp'):
                                    keyboard.unhook(button.keyboard_hook_temp)
                                    delattr(button, 'keyboard_hook_temp')
                                if hasattr(button, 'mouse_hook_temp'):
                                    mouse.unhook(button.mouse_hook_temp)
                                    delattr(button, 'mouse_hook_temp')
                            except Exception:
                                pass
                            finish_hotkey_assignment()
                            self.setting_hotkey = False
                            return
                    except Exception:
                        pass
                # Delay a bit to allow combination keys; if user presses another key quickly, normal path will handle it
                try:
                    print(f"Debug: Setting timer for _assign_bare_modifier with modifier_name: '{name}'")
                    self.root.after(200, lambda: _assign_bare_modifier(name))
                    print(f"Debug: Timer set successfully for {name}")
                except Exception as e:
                    print(f"Debug: Error setting timer: {e}")
                return

            # Only build combination if a non-modifier key was pressed
            # (Modifier keys alone are handled by the timer above)
            if not combo_state['non_modifier_pressed']:
                print(f"Debug: Skipping combination building - only modifier key pressed")
                return
            
            print(f"Debug: Building combination for non-modifier key")

            # Build combination string from tracked held modifiers + key
            try:
                # Use our tracked modifiers instead of keyboard.is_pressed()
                mods = list(combo_state['held_modifiers'])
                
                # Debug output
                print(f"Debug: Pressed key '{name}', tracked modifiers: {mods}")
            except Exception:
                mods = []

            # Determine base key (support numpad keys explicitly)
            base_key = name
            if event.scan_code in self.numpad_scan_codes:
                sym = self.numpad_scan_codes[event.scan_code]
                base_key = f"num_{sym}"

            # If base_key itself is a modifier, include it if not already in mods; otherwise avoid duplicate
            if base_key in ("ctrl", "shift", "alt", "windows", "left alt", "right alt"):
                combo_parts = (mods + [base_key]) if base_key not in mods else mods[:]
            else:
                combo_parts = mods + [base_key]
            key_name = "+".join(p for p in combo_parts if p)
            
            # Debug output
            print(f"Debug: Final key combination: '{key_name}' (from parts: {combo_parts})")

            # Prevent duplicates against Stop hotkey
            if getattr(self, 'stop_hotkey', None) == key_name:
                # Unhook temp hooks and set flags BEFORE showing the warning
                self.setting_hotkey = False
                self._hotkey_assignment_cancelled = True
                if hasattr(button, 'keyboard_hook_temp'):
                    try:
                        keyboard.unhook(button.keyboard_hook_temp)
                    except Exception:
                        pass
                    delattr(button, 'keyboard_hook_temp')
                if hasattr(button, 'mouse_hook_temp'):
                    try:
                        mouse.unhook(button.mouse_hook_temp)
                    except Exception:
                        pass
                    delattr(button, 'mouse_hook_temp')
                finish_hotkey_assignment()
                try:
                    messagebox.showwarning("Hotkey In Use", "This hotkey is already assigned to: Stop Hotkey")
                except Exception:
                    pass
                # Reset label text
                if hasattr(button, 'hotkey') and button.hotkey:
                    disp_prev = button.hotkey.replace('num_', 'num:') if button.hotkey.startswith('num_') else button.hotkey
                    button.config(text=f"Set Hotkey: [ {disp_prev.upper()} ]")
                else:
                    button.config(text="Set Hotkey")
                return

            # Disallow duplicate hotkeys
            duplicate_found = False
            for area in self.areas:
                if area[1] is not button and hasattr(area[1], 'hotkey') and area[1].hotkey == key_name:
                    duplicate_found = True
                    break

            if duplicate_found:
    
                # Unhook temp hooks and set flags BEFORE showing the warning
                self.setting_hotkey = False
                self._hotkey_assignment_cancelled = True  # Block all further events
                if hasattr(button, 'keyboard_hook_temp'):
                    try:
                        keyboard.unhook(button.keyboard_hook_temp)
                    except Exception:
                        pass
                    delattr(button, 'keyboard_hook_temp')
                if hasattr(button, 'mouse_hook_temp'):
                    try:
                        mouse.unhook(button.mouse_hook_temp)
                    except Exception:
                        pass
                    delattr(button, 'mouse_hook_temp')
                finish_hotkey_assignment()
                # Now show the warning dialog (no hooks are active)
                if hasattr(button, 'hotkey'):
                    display_name = self._get_display_hotkey(button)
                    button.config(text=f"Set Hotkey: [ {display_name.upper()} ]")
                else:
                    button.config(text="Set Hotkey")
                # Find the area name that's using this hotkey
                for area in self.areas:
                    if area[1] is not button and hasattr(area[1], 'hotkey') and area[1].hotkey == key_name:
                        area_name = area[3].get() if hasattr(area[3], 'get') else "Unknown Area"
                        break
                show_thinkr_warning(self, area_name)

                return  # Keep the return but without False since we want to show the warning
                
            # Only proceed with setting hotkey if no duplicate was found

            button.hotkey = key_name
            # Display: make NUMPAD look nice and uppercase
            # Display nicer labels for sided modifiers
            # Convert numpad keys to display format
            if key_name.startswith('num_'):
                display_name = self._convert_numpad_to_display(key_name)
            elif key_name.startswith('controller_'):
                # Extract controller button name for display
                button_name = key_name.replace('controller_', '')
                # Handle D-Pad names specially
                if button_name.startswith('dpad_'):
                    dpad_name = button_name.replace('dpad_', '')
                    display_name = f"D-Pad {dpad_name.title()}"
                else:
                    display_name = button_name
            else:
                display_name = key_name.replace('numpad ', 'NUMPAD ') \
                                                     .replace('ctrl','CTRL') \
                .replace('left alt','L-ALT').replace('right alt','R-ALT') \
                .replace('windows','WIN') \
                .replace('multiply', '*').replace('add', '+').replace('subtract', '-').replace('divide', '/')
            button.config(text=f"Set Hotkey: [ {display_name.upper()} ]")

            self.setup_hotkey(button, area_frame)
            self.setting_hotkey = False
            print(f"Hotkey assignment completed for button: {button}")

            # Unhook both temp hooks if they exist
            if hasattr(button, 'keyboard_hook_temp'):
                try:
                    keyboard.unhook(button.keyboard_hook_temp)
                except Exception:
                    pass
                delattr(button, 'keyboard_hook_temp')
            if hasattr(button, 'mouse_hook_temp'):
                try:
                    mouse.unhook(button.mouse_hook_temp)
                except Exception:
                    pass
                delattr(button, 'mouse_hook_temp')

            finish_hotkey_assignment()
            # Guard: prevent any further hotkey assignment callbacks
            self.setting_hotkey = False

            return

        def on_mouse_click(event):
            # Only handle button down events when in hotkey setting mode
            if not self.setting_hotkey or not isinstance(event, mouse.ButtonEvent) or event.event_type != mouse.DOWN:
                return
            
            # List of all potential names for left and right mouse buttons
            LEFT_MOUSE_BUTTONS = [
                '1', 'left', 'primary', 'select', 'action', 'button1', 'mouse1'
            ]
            RIGHT_MOUSE_BUTTONS = [
                '2', 'right', 'secondary', 'context', 'alternate', 'button2', 'mouse2'
            ]
            
            # Get the button name from the event
            button_name = str(event.button).lower()
            
            # Use the button identifier directly from the mouse library
            # This could be a number (1, 2, 3) or a string ('x', 'wheel', etc.)
            button_identifier = event.button
            
            # Check if this is a left or right mouse button
            is_left_button = button_identifier == 1 or str(button_identifier).lower() in ['left', 'primary', 'select', 'action', 'button1', 'mouse1']
            is_right_button = button_identifier == 2 or str(button_identifier).lower() in ['right', 'secondary', 'context', 'alternate', 'button2', 'mouse2']
            
            # Check if this is a left/right mouse button
            if is_left_button or is_right_button:
                # Get the current state of the allow_mouse_buttons checkbox
                allow_mouse_buttons = False
                if hasattr(self, 'allow_mouse_buttons_var'):
                    try:
                        allow_mouse_buttons = self.allow_mouse_buttons_var.get()
                    except Exception as e:
                        print(f"Error getting allow_mouse_buttons_var: {e}")
                

                
                if not allow_mouse_buttons:

                    if not hasattr(self, '_mouse_button_error_shown'):
                        messagebox.showwarning("Warning", "Left and right mouse buttons cannot be used as hotkeys.\nCheck 'Allow mouse left/right:' to enable them.")
                        self._mouse_button_error_shown = True
                    return
                
                # If we get here, mouse buttons are allowed
                button_name = f"button{button_identifier}"
    
                # Create a mock keyboard event for the mouse button
                mock_event = type('MockEvent', (), {
                    'name': button_name,
                    'scan_code': None,
                    'event_type': 'down'
                })
                
                # Store the original button identifier for the mouse hook
                button.original_button_id = button_identifier
                
                on_key_press(mock_event)
                return
            
            # Create a mock keyboard event
            mock_event = type('MockEvent', (), {
                'name': f'button{button_identifier}',  # Use the actual button identifier
                'scan_code': None
            })
            
            # Store the original button identifier for the mouse hook
            button.original_button_id = button_identifier
            
            on_key_press(mock_event)

        def on_controller_button_press(event):
            """Controller support disabled - pygame removed to reduce Windows security flags"""
            print("Controller support disabled - pygame removed to reduce Windows security flags")

        def on_controller_hat_press(event):
            """Controller support disabled - pygame removed to reduce Windows security flags"""
            print("Controller support disabled - pygame removed to reduce Windows security flags")



        # Clean up previous hooks
        if hasattr(button, 'keyboard_hook'):
            try:
                # For add_hotkey hooks, we need to use the remove method from the hook object
                if hasattr(button.keyboard_hook, 'remove'):
                    button.keyboard_hook.remove()
                else:
                    # Fallback to unhook if remove method doesn't exist
                    keyboard.unhook(button.keyboard_hook)
                delattr(button, 'keyboard_hook')
            except Exception as e:
                print(f"Error cleaning up keyboard hook: {e}")
        if hasattr(button, 'mouse_hook_id'):
            try:
                mouse.unhook(button.mouse_hook_id)
                delattr(button, 'mouse_hook_id')
            except Exception as e:
                print(f"Error cleaning up mouse hook ID: {e}")
        if hasattr(button, 'mouse_hook'):
            try:
                delattr(button, 'mouse_hook')
            except Exception as e:
                print(f"Error cleaning up mouse hook function: {e}")
        button.config(text="Press any key or combination...")
        
        # Temporarily disable all existing hotkey handlers to prevent conflicts during assignment
        print("Debug: Temporarily disabling existing hotkey handlers during assignment")
        try:
            keyboard.unhook_all()
            print("Debug: All existing keyboard hooks disabled")
        except Exception as e:
            print(f"Debug: Error disabling keyboard hooks: {e}")
        
        self.setting_hotkey = True  # Enable hotkey assignment mode before installing hooks
        
        # Live preview of currently held modifiers
        def _update_hotkey_preview():
            if self._hotkey_assignment_cancelled or not self.setting_hotkey:
                return
            try:
                mods = []
                # Use scan code detection for more reliable left/right distinction
                left_ctrl_pressed, right_ctrl_pressed = detect_ctrl_keys()
                
                if left_ctrl_pressed or right_ctrl_pressed: mods.append('CTRL')
                if keyboard.is_pressed('shift'): mods.append('SHIFT')
                if keyboard.is_pressed('left alt'): mods.append('L-ALT')
                if keyboard.is_pressed('right alt'): mods.append('R-ALT')
                if keyboard.is_pressed('left windows') or keyboard.is_pressed('right windows') or keyboard.is_pressed('windows'):
                    mods.append('WIN')
                preview = " + ".join(mods)
                if preview:
                    button.config(text=f"Press key: [ {preview} + ]")
                else:
                    button.config(text="Press any key or combination...")
            except Exception:
                pass
            # Schedule next update
            try:
                self._hotkey_preview_job = self.root.after(80, _update_hotkey_preview)
            except Exception:
                pass
        
        # Start live preview
        try:
            self._hotkey_preview_job = self.root.after(80, _update_hotkey_preview)
        except Exception:
            pass
        
        button.keyboard_hook_temp = keyboard.on_press(on_key_press)
        button.mouse_hook_temp = mouse.hook(on_mouse_click)
        
        # Start controller monitoring for hotkey assignment if controller support is available
        if CONTROLLER_AVAILABLE:
            self._start_controller_hotkey_monitoring(button, area_frame, finish_hotkey_assignment)

        # Also listen for Shift key release to allow assigning bare SHIFT reliably
        def on_shift_release(_e):
            if self._hotkey_assignment_cancelled or not self.setting_hotkey:
                return
            if combo_state.get('non_modifier_pressed'):
                return
            # Determine which shift key was released from event name if available
            side_label = 'left'
            try:
                raw = (getattr(_e, 'name', '') or '').lower()
                if 'right' in raw or 'right shift' in raw:
                    side_label = 'right'
            except Exception:
                pass
            # Assign bare sided SHIFT
            key_name = f"{side_label} shift"
            # Prevent duplicates: Stop hotkey
            if getattr(self, 'stop_hotkey', None) == key_name:
                # Cleanup temp hooks and end assignment with warning
                try:
                    if hasattr(button, 'keyboard_hook_temp'):
                        keyboard.unhook(button.keyboard_hook_temp)
                        delattr(button, 'keyboard_hook_temp')
                    if hasattr(button, 'mouse_hook_temp'):
                        mouse.unhook(button.mouse_hook_temp)
                        delattr(button, 'mouse_hook_temp')
                    if hasattr(button, 'shift_release_hooks'):
                        for h in button.shift_release_hooks:
                            try:
                                keyboard.unhook(h)
                            except Exception:
                                pass
                        delattr(button, 'shift_release_hooks')
                    if hasattr(button, 'ctrl_release_hooks'):
                        for h in button.ctrl_release_hooks:
                            try:
                                keyboard.unhook(h)
                            except Exception:
                                pass
                        delattr(button, 'ctrl_release_hooks')
                except Exception:
                    pass
                self.setting_hotkey = False
                finish_hotkey_assignment()
                try:
                    messagebox.showwarning("Hotkey In Use", "This hotkey is already assigned to: Stop Hotkey")
                except Exception:
                    pass
                return
            # Prevent duplicates: other areas
            for area in self.areas:
                if area[1] is not button and hasattr(area[1], 'hotkey') and area[1].hotkey == key_name:
                    try:
                        if hasattr(button, 'keyboard_hook_temp'):
                            keyboard.unhook(button.keyboard_hook_temp)
                            delattr(button, 'keyboard_hook_temp')
                        if hasattr(button, 'mouse_hook_temp'):
                            mouse.unhook(button.mouse_hook_temp)
                            delattr(button, 'mouse_hook_temp')
                        if hasattr(button, 'shift_release_hooks'):
                            for h in button.shift_release_hooks:
                                try:
                                    keyboard.unhook(h)
                                except Exception:
                                    pass
                            delattr(button, 'shift_release_hooks')
                        if hasattr(button, 'ctrl_release_hooks'):
                            for h in button.ctrl_release_hooks:
                                try:
                                    keyboard.unhook(h)
                                except Exception:
                                    pass
                            delattr(button, 'ctrl_release_hooks')
                    except Exception:
                        pass
                    self.setting_hotkey = False
                    finish_hotkey_assignment()
                    area_name = area[3].get() if hasattr(area[3], 'get') else "Unknown Area"
                    show_thinkr_warning(self, area_name)
                    return
            button.hotkey = key_name
            button.config(text=f"Set Hotkey: [ {'L-SHIFT' if side_label=='left' else 'R-SHIFT'} ]")
            self.setup_hotkey(button, area_frame)
            # Clean up temp hooks (keyboard/mouse/shift release hooks)
            try:
                if hasattr(button, 'keyboard_hook_temp'):
                    keyboard.unhook(button.keyboard_hook_temp)
                    delattr(button, 'keyboard_hook_temp')
                if hasattr(button, 'mouse_hook_temp'):
                    mouse.unhook(button.mouse_hook_temp)
                    delattr(button, 'mouse_hook_temp')
                if hasattr(button, 'shift_release_hooks'):
                    for h in button.shift_release_hooks:
                        try:
                            keyboard.unhook(h)
                        except Exception:
                            pass
                    delattr(button, 'shift_release_hooks')
                if hasattr(button, 'ctrl_release_hooks'):
                    for h in button.ctrl_release_hooks:
                        try:
                            keyboard.unhook(h)
                        except Exception:
                            pass
                    delattr(button, 'ctrl_release_hooks')
            except Exception:
                pass
            self.setting_hotkey = False
            finish_hotkey_assignment()

        try:
            button.shift_release_hooks = [
                keyboard.on_release_key('left shift', on_shift_release),
                keyboard.on_release_key('right shift', on_shift_release),
            ]
        except Exception:
            button.shift_release_hooks = []
        
        # Also listen for Ctrl key release to allow assigning bare CTRL reliably
        def on_ctrl_release(_e):
            if self._hotkey_assignment_cancelled or not self.setting_hotkey:
                return
            if combo_state.get('non_modifier_pressed'):
                return
            # Determine which ctrl key was released using scan code for reliability
            side_label = 'left'
            try:
                scan_code = getattr(_e, 'scan_code', None)
                if scan_code == 157:  # Right Ctrl scan code
                    side_label = 'right'
                elif scan_code == 29:  # Left Ctrl scan code
                    side_label = 'left'
                else:
                    # Fallback to event name if scan code is not available
                    raw = (getattr(_e, 'name', '') or '').lower()
                    if 'right' in raw or 'right ctrl' in raw:
                        side_label = 'right'
            except Exception:
                pass
            # Assign bare CTRL (no longer sided)
            key_name = "ctrl"
            # Prevent duplicates: Stop hotkey
            if getattr(self, 'stop_hotkey', None) == key_name:
                # Cleanup temp hooks and end assignment with warning
                try:
                    if hasattr(button, 'keyboard_hook_temp'):
                        keyboard.unhook(button.keyboard_hook_temp)
                        delattr(button, 'keyboard_hook_temp')
                    if hasattr(button, 'mouse_hook_temp'):
                        mouse.unhook(button.mouse_hook_temp)
                        delattr(button, 'mouse_hook_temp')
                    if hasattr(button, 'ctrl_release_hooks'):
                        for h in button.ctrl_release_hooks:
                            try:
                                keyboard.unhook(h)
                            except Exception:
                                pass
                        delattr(button, 'ctrl_release_hooks')
                except Exception:
                    pass
                self.setting_hotkey = False
                finish_hotkey_assignment()
                try:
                    messagebox.showwarning("Hotkey In Use", "This hotkey is already assigned to: Stop Hotkey")
                except Exception:
                    pass
                return
            # Prevent duplicates: other areas
            for area in self.areas:
                if area[1] is not button and hasattr(area[1], 'hotkey') and area[1].hotkey == key_name:
                    try:
                        if hasattr(button, 'keyboard_hook_temp'):
                            keyboard.unhook(button.keyboard_hook_temp)
                            delattr(button, 'keyboard_hook_temp')
                        if hasattr(button, 'mouse_hook_temp'):
                            mouse.unhook(button.mouse_hook_temp)
                            delattr(button, 'mouse_hook_temp')
                        if hasattr(button, 'ctrl_release_hooks'):
                            for h in button.ctrl_release_hooks:
                                try:
                                    keyboard.unhook(h)
                                except Exception:
                                    pass
                            delattr(button, 'ctrl_release_hooks')
                    except Exception:
                        pass
                    self.setting_hotkey = False
                    finish_hotkey_assignment()
                    area_name = area[3].get() if hasattr(area[3], 'get') else "Unknown Area"
                    show_thinkr_warning(self, area_name)
                    return
            button.hotkey = key_name
            button.config(text=f"Set Hotkey: [ CTRL ]")
            self.setup_hotkey(button, area_frame)
            # Clean up temp hooks (keyboard/mouse/ctrl release hooks)
            try:
                if hasattr(button, 'keyboard_hook_temp'):
                    keyboard.unhook(button.keyboard_hook_temp)
                    delattr(button, 'keyboard_hook_temp')
                if hasattr(button, 'mouse_hook_temp'):
                    mouse.unhook(button.mouse_hook_temp)
                    delattr(button, 'mouse_hook_temp')
                if hasattr(button, 'ctrl_release_hooks'):
                    for h in button.ctrl_release_hooks:
                        try:
                            keyboard.unhook(h)
                        except Exception:
                            pass
                    delattr(button, 'ctrl_release_hooks')
            except Exception:
                pass
            finish_hotkey_assignment()
            self.setting_hotkey = False
        
        try:
            button.ctrl_release_hooks = [
                keyboard.on_release_key('ctrl', on_ctrl_release),
            ]
        except Exception:
            button.ctrl_release_hooks = []
        
        # Set 3-second timeout for hotkey setting
        def unhook_mouse():
            try:
                # Safely clean up mouse hook
                if hasattr(button, 'mouse_hook_temp') and button.mouse_hook_temp is not None:
                    try:
                        # Best-effort unhook if possible
                        try:
                            mouse.unhook(button.mouse_hook_temp)
                        except Exception:
                            pass
                    except Exception as e:
                        print(f"Warning: Error unhooking mouse: {e}")
                    finally:
                        # Always clean up the attribute to prevent memory leaks
                        if hasattr(button, 'mouse_hook_temp'):
                            delattr(button, 'mouse_hook_temp')
                
                # Safely clean up keyboard hook
                if hasattr(button, 'keyboard_hook_temp') and button.keyboard_hook_temp is not None:
                    try:
                        # Check if the hook is still active before trying to remove it
                        if hasattr(keyboard, '_listener') and hasattr(keyboard._listener, 'running') and keyboard._listener.running:
                            keyboard.unhook(button.keyboard_hook_temp)
                    except Exception as e:
                        print(f"Warning: Error unhooking keyboard: {e}")
                    finally:
                        # Always clean up the attribute to prevent memory leaks
                        if hasattr(button, 'keyboard_hook_temp'):
                            delattr(button, 'keyboard_hook_temp')
                # Clean up shift release hooks
                if hasattr(button, 'shift_release_hooks'):
                    try:
                        for h in button.shift_release_hooks:
                            try:
                                keyboard.unhook(h)
                            except Exception:
                                pass
                    finally:
                        delattr(button, 'shift_release_hooks')
                
                # Clean up ctrl release hooks
                if hasattr(button, 'ctrl_release_hooks'):
                    try:
                        for h in button.ctrl_release_hooks:
                            try:
                                keyboard.unhook(h)
                            except Exception:
                                pass
                    finally:
                        delattr(button, 'ctrl_release_hooks')
                
                self.setting_hotkey = False
                self._hotkey_assignment_cancelled = True
                if not hasattr(button, 'hotkey'):
                    button.config(text="Set Hotkey")
            except Exception as e:
                print(f"Warning: Error during hook cleanup: {e}")
                button.config(text="Set Hotkey")
        self.root.after(3000, unhook_mouse)

    def _start_controller_hotkey_monitoring(self, button, area_frame, finish_hotkey_assignment):
        """Start monitoring controller input for hotkey assignment"""
        if not CONTROLLER_AVAILABLE:
            return
            
        def monitor_controller():
            try:
                button_name = self.controller_handler.wait_for_button_press(timeout=15)
                if button_name and not self._hotkey_assignment_cancelled:
                    key_name = f"controller_{button_name}"
                    
                    # Check if this controller button is already used by any area
                    for area in self.areas:
                        if area[1] is not button and hasattr(area[1], 'hotkey') and area[1].hotkey == key_name:
                            show_thinkr_warning(self, area[3].get())
                            self._hotkey_assignment_cancelled = True
                            finish_hotkey_assignment()
                            return
                    
                    # Check if it conflicts with stop hotkey
                    if getattr(self, 'stop_hotkey', None) == key_name:
                        messagebox.showwarning("Hotkey In Use", "This hotkey is already assigned to: Stop Hotkey")
                        self._hotkey_assignment_cancelled = True
                        finish_hotkey_assignment()
                        return
                    
                    # Set the hotkey
                    button.hotkey = key_name
                    button.config(text=f"Hotkey: [ Controller {button_name} ]")
                    self.setup_hotkey(button, area_frame)
                    print(f"Set hotkey: {key_name}\n--------------------------")
                    
                    finish_hotkey_assignment()
                else:
                    # Timeout or cancelled - do nothing, let keyboard/mouse handle it
                    pass
            except Exception as e:
                print(f"Error in controller monitoring: {e}")
                # Don't call finish_hotkey_assignment here, let keyboard/mouse handle it
        
        # Start controller monitoring in background
        threading.Thread(target=monitor_controller, daemon=True).start()

    def _start_controller_stop_hotkey_monitoring(self, finish_hotkey_assignment):
        """Start monitoring controller input for stop hotkey assignment"""
        if not CONTROLLER_AVAILABLE:
            return
            
        def monitor_controller():
            try:
                button_name = self.controller_handler.wait_for_button_press(timeout=15)
                if button_name and not self._hotkey_assignment_cancelled:
                    key_name = f"controller_{button_name}"
                    
                    # Check if this controller button is already used by any area
                    for area_frame, hotkey_button, _, area_name_var, _, _, _ in self.areas:
                        if hasattr(hotkey_button, 'hotkey') and hotkey_button.hotkey == key_name:
                            show_thinkr_warning(self, area_name_var.get())
                            self._hotkey_assignment_cancelled = True
                            finish_hotkey_assignment()
                            return
                    
                    # Remove existing stop hotkey if it exists
                    if hasattr(self, 'stop_hotkey'):
                        try:
                            if hasattr(self.stop_hotkey_button, 'mock_button'):
                                self._cleanup_hooks(self.stop_hotkey_button.mock_button)
                        except Exception as e:
                            print(f"Error cleaning up stop hotkey hooks: {e}")
                    
                    self.stop_hotkey = key_name
                    
                    # Create a mock button object to use with setup_hotkey
                    mock_button = type('MockButton', (), {'hotkey': key_name, 'is_stop_button': True})
                    self.stop_hotkey_button.mock_button = mock_button
                    
                    # Setup the hotkey
                    self.setup_hotkey(self.stop_hotkey_button.mock_button, None)
                    
                    display_name = f"Controller {button_name}"
                    self.stop_hotkey_button.config(text=f"Stop Hotkey: [ {display_name.upper()} ]")
                    print(f"Set Stop hotkey: {key_name}\n--------------------------")
                    
                    finish_hotkey_assignment()
                else:
                    # Timeout or cancelled - do nothing, let keyboard/mouse handle it
                    pass
            except Exception as e:
                print(f"Error in controller monitoring: {e}")
                # Don't call finish_hotkey_assignment here, let keyboard/mouse handle it
        
        # Start controller monitoring in background
        threading.Thread(target=monitor_controller, daemon=True).start()



    def _check_controller_hotkeys(self, button_name):
        """Check if a controller button press should trigger any hotkeys"""
        try:
            # Check area hotkeys
            for area_frame, hotkey_button, _, area_name_var, _, _, _ in self.areas:
                if hasattr(hotkey_button, 'hotkey') and hotkey_button.hotkey is not None and hotkey_button.hotkey.startswith('controller_'):
                    controller_button = hotkey_button.hotkey.replace('controller_', '')
                    if controller_button == button_name:
                        print(f"Controller hotkey triggered for area: {area_name_var.get()}")
                        # Trigger the hotkey action
                        if hasattr(hotkey_button, 'controller_hook'):
                            hotkey_button.controller_hook()
                        break
            
            # Check stop hotkey
            if hasattr(self, 'stop_hotkey') and self.stop_hotkey is not None and self.stop_hotkey.startswith('controller_'):
                controller_button = self.stop_hotkey.replace('controller_', '')
                if controller_button == button_name:
                    print(f"Controller stop hotkey triggered")
                    self.stop_speaking()
                    
        except Exception as e:
            print(f"Error checking controller hotkeys: {e}")

    def save_layout(self):
        # Check if there are no areas
        if not self.areas:
            messagebox.showerror("Error", "There is nothing to save.")
            return
            
        # Reset unsaved changes flag
        self._has_unsaved_changes = False

        # Check if all areas have coordinates set, but ignore Auto Read
        for area_frame, _, _, area_name_var, _, _, _ in self.areas:
            if area_name_var.get() == "Auto Read":
                continue
            if not hasattr(area_frame, 'area_coords'):
                messagebox.showerror("Error", f"Area '{area_name_var.get()}' does not have a defined area, remove it or configure before saving.")
                return

        layout = {
            "version": APP_VERSION,
            "bad_word_list": self.bad_word_list.get(),
            "ignore_usernames": self.ignore_usernames_var.get(),
            "ignore_previous": self.ignore_previous_var.get(),
            "ignore_gibberish": self.ignore_gibberish_var.get(),
            "pause_at_punctuation": self.pause_at_punctuation_var.get(),
            "better_unit_detection": self.better_unit_detection_var.get(),
            "read_game_units": self.read_game_units_var.get(),
            "fullscreen_mode": self.fullscreen_mode_var.get(),
            "stop_hotkey": self.stop_hotkey,
            "volume": self.volume.get(),
            "areas": []
        }
        for area_frame, hotkey_button, _, area_name_var, preprocess_var, voice_var, speed_var in self.areas:
            # Skip Auto Read area as it has its own save file
            if area_name_var.get() == "Auto Read":
                continue
                
            if hasattr(area_frame, 'area_coords'):
                area_info = {
                    "coords": area_frame.area_coords,
                    "name": area_name_var.get(),
                    "hotkey": hotkey_button.hotkey if hasattr(hotkey_button, 'hotkey') else None,
                    "preprocess": preprocess_var.get(),
                    "voice": voice_var.get(),
                    "speed": speed_var.get(),
                    "settings": self.processing_settings.get(area_name_var.get(), {})
                }
                layout["areas"].append(area_info)

        # Get the current file path and initialdir for the save dialog
        current_file = self.layout_file.get()
        initial_dir = os.path.dirname(current_file) if current_file else os.getcwd()
        initial_file = os.path.basename(current_file) if current_file else ""

        # Show Save As dialog with the current file pre-selected
        file_path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json")],
            initialdir=initial_dir,
            initialfile=initial_file
        )

        if not file_path:  # User cancelled
            return

        try:
            # Save the layout
            with open(file_path, 'w') as f:
                json.dump(layout, f, indent=4)
            
            # Store the full path in layout_file
            self.layout_file.set(file_path)
            
            # Show feedback in status label
            if hasattr(self, '_feedback_timer') and self._feedback_timer:
                self.root.after_cancel(self._feedback_timer)
            
            # Show save success message
            self.status_label.config(text=f"Layout saved to: {os.path.basename(file_path)}")
            self._feedback_timer = self.root.after(2000, lambda: self.status_label.config(text=""))
            
            print(f"Layout saved to {file_path}\n--------------------------")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save layout: {str(e)}")
            print(f"Error saving layout: {e}")

    def load_game_units(self):
        """Load game units from JSON file in GameReader directory."""
        import tempfile, os, json, re
        temp_path = os.path.join(tempfile.gettempdir(), 'GameReader')
        os.makedirs(temp_path, exist_ok=True)
        
        file_path = os.path.join(temp_path, 'gamer_units.json')
        if os.path.exists(file_path):
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    # Remove comments from JSON file before parsing
                    content = f.read()
                    # Remove single-line comments (// ...)
                    content = re.sub(r'//.*?$', '', content, flags=re.MULTILINE)
                    # Remove multi-line comments (/* ... */)
                    content = re.sub(r'/\*.*?\*/', '', content, flags=re.DOTALL)
                    # Parse the cleaned JSON
                    return json.loads(content)
            except (json.JSONDecodeError, UnicodeDecodeError) as e:
                print(f"Warning: Error reading game units file: {e}, using default units")
        
        # Create default game units if file doesn't exist or is invalid
        default_units = {
            'xp': 'Experience Points',
            'hp': 'Health Points',
            'mp': 'Mana Points',
            'gp': 'Gold Pieces',
            'pp': 'Platinum Pieces',
            'sp': 'Skill Points',
            'ep': 'Energy Points',
            'ap': 'Action Points',
            'bp': 'Battle Points',
            'lp': 'Loyalty Points',
            'cp': 'Challenge Points',
            'vp': 'Victory Points',
            'rp': 'Reputation Points',
            'tp': 'Talent Points',
            'ar': 'Armor Rating',
            'dmg': 'Damage',
            'dps': 'Damage Per Second',
            'def': 'Defense',
            'mat': 'Materials',
            'exp': 'Exploration Points',
            '§': 'Simoliance',
            'v-bucks': 'Virtual Bucks',
            'r$': 'Robux',
            'nmt': 'Nook Miles Tickets',
            'be': 'Blue Essence',
            'radianite': 'Radianite Points',
            'ow coins': 'Overwatch Coins',
            '₽': 'PokeDollars',
            '€$': 'Eurodollars',
            'z': 'Zenny',
            'l': 'Lunas',
            'e': 'Eve',
            'i': 'Isk',
            'j': 'Jewel',
            'sc': 'Star Coins',
            'o2': 'Oxygen',
            'pu': 'Power Units',
            'mc': 'Mana Crystals',
            'es': 'Essence',
            'sh': 'Shards',
            'st': 'Stars',
            'mu': 'Munny',
            'b': 'Bolts',
            'r': 'Rings',
            'ca': 'Caps',
            'rns': 'Runes',
            'sl': 'Souls',
            'fav': 'Favor',
            'am': 'Amber',
            'cc': 'Crystal Cores',
            'fg': 'Fragments'
        }
        
        # Save default units to file
        with open(file_path, 'w', encoding='utf-8') as f:
            header = '''//  Game Units Configuration
//  Format: "short_name": "Full Name"
//  Example: "xp" will be read as "Experience Points"
//  Enable "Read gamer units" in the main window to use this feature

'''
            f.write(header)
            json.dump(default_units, f, indent=4, ensure_ascii=False)
        
        return default_units

    def save_game_units(self):
        """Save game units to JSON file in GameReader directory."""
        import tempfile, os, json
        
        temp_path = os.path.join(tempfile.gettempdir(), 'GameReader')
        os.makedirs(temp_path, exist_ok=True)
        
        file_path = os.path.join(temp_path, 'game_units.json')
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                header = '''//  Game Units Configuration
//  Format: "short_name": "Full Name"
//  Example: "xp" will be read as "Experience Points"
//  Enable "Read gamer units" in the main window to use this feature

'''
                f.write(header)
                json.dump(self.game_units, f, indent=4, ensure_ascii=False)
            print(f"Game units saved to: {file_path}")
            
            # Show feedback in status label
            if hasattr(self, '_feedback_timer') and self._feedback_timer:
                self.root.after_cancel(self._feedback_timer)
            
            # Show load success message
            self.status_label.config(text="Game units saved successfully!")
            self._feedback_timer = self.root.after(3000, lambda: self.status_label.config(text=""))
            
            return True
        except Exception as e:
            print(f"Error saving game units: {e}")
            return False

    def open_game_reader_folder(self):
        """Open the GameReader folder in Windows Explorer."""
        import os
        import subprocess
        
        # Get the current username
        username = os.getlogin()
        # Construct the path to GameReader folder
        folder_path = os.path.join(os.getenv('LOCALAPPDATA'), 'Temp', 'GameReader')
        
        # Create folder if it doesn't exist
        os.makedirs(folder_path, exist_ok=True)
        
        # Open the folder in Windows Explorer
        try:
            subprocess.Popen(f'explorer "{folder_path}"')
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open folder: {str(e)}")

    def normalize_text(self, text):
        """Normalize text by removing punctuation and making it lowercase."""
        import string
        # Remove punctuation and make lowercase
        text = text.lower()
        text = text.translate(str.maketrans('', '', string.punctuation))
        # Remove extra whitespace
        text = ' '.join(text.split())
        return text

    def on_drop(self, event):
        """Handle file drop event"""
        try:
            # Get the file path from the drop event
            # On Windows, the path is wrapped in {}
            file_path = event.data.strip('{}')
            
            # Clean up the path (remove quotes if present)
            file_path = file_path.strip('\"\'')
            
            # Normalize the path (convert forward slashes to backslashes on Windows)
            file_path = os.path.normpath(file_path)
            
            # Check if the file exists and is a JSON file
            if not os.path.isfile(file_path) or not file_path.lower().endswith('.json'):
                messagebox.showerror("Error", "Please drop a valid JSON layout file")
                return
            
            # Check if we have a file already loaded
            if self.layout_file.get():
                # If it's the same file, just return
                if os.path.normpath(self.layout_file.get()) == file_path:
                    return
                    
                # If there are unsaved changes, show warning
                if self._has_unsaved_changes:
                    response = messagebox.askyesnocancel(
                        "Unsaved Changes",
                        f"You have unsaved changes in the current layout.\n\n"
                        f"Current: {os.path.basename(self.layout_file.get())}\n"
                        f"New: {os.path.basename(file_path)}\n\n"
                        "Do you want to save changes before loading the new layout?\n"
                        "(Yes = Save and load, No = Discard changes and load, Cancel = Do nothing)"
                    )
                    if response is None:  # Cancel
                        return
                    elif response:  # Yes - Save and load
                        self.save_layout()
                else:
                    # No unsaved changes, just confirm loading new file
                    if not messagebox.askyesno(
                        "Load New Layout",
                        f"Load new layout file?\n\n"
                        f"Current: {os.path.basename(self.layout_file.get())}\n"
                        f"New: {os.path.basename(file_path)}"
                    ):
                        return  # User chose not to load the new file
            
            # Load the new layout
            self._load_layout_file(file_path)
            
        except Exception as e:
            messagebox.showerror("Error", f"Error handling dropped file: {str(e)}")
            import traceback
            traceback.print_exc()

    def load_layout(self, file_path=None):
        """Show file dialog to load a layout file"""
        if not file_path:
            file_path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
            if not file_path:  # User cancelled
                return
        
        self._load_layout_file(file_path)

    def _set_unsaved_changes(self):
        """Mark that there are unsaved changes"""
        self._has_unsaved_changes = True
        
    def _validate_layout_data(self, layout):
        """Validate layout data for security and integrity"""
        if not isinstance(layout, dict):
            raise ValueError("Layout must be a dictionary")
        
        # Define expected structure and types
        expected_fields = {
            'version': str,
            'bad_word_list': str,
            'ignore_usernames': bool,
            'ignore_previous': bool,
            'ignore_gibberish': bool,
            'pause_at_punctuation': bool,
            'better_unit_detection': bool,
            'read_game_units': bool,
            'fullscreen_mode': bool,
            'stop_hotkey': (str, type(None)),
            'volume': str,
            'areas': list
        }
        
        # Validate top-level fields
        for field, expected_type in expected_fields.items():
            if field in layout:
                value = layout[field]
                if not isinstance(value, expected_type):
                    raise ValueError(f"Invalid type for {field}: expected {expected_type}, got {type(value)}")
                
                # Additional validation for specific fields
                if field == 'version':
                    if not value or len(value) > 10:  # Reasonable version string length
                        raise ValueError("Invalid version string")
                elif field == 'bad_word_list':
                    if len(value) > 10000:  # Reasonable limit for bad word list
                        raise ValueError("Bad word list too long")
                elif field == 'volume':
                    try:
                        vol_int = int(value)
                        if vol_int < 0 or vol_int > 100:
                            raise ValueError("Volume must be between 0 and 100")
                    except ValueError:
                        raise ValueError("Invalid volume value")
                elif field == 'stop_hotkey':
                    if value is not None and (not isinstance(value, str) or len(value) > 50):
                        raise ValueError("Invalid stop hotkey")
        
        # Validate areas array
        if 'areas' in layout:
            areas = layout['areas']
            if len(areas) > 50:  # Reasonable limit for number of areas
                raise ValueError("Too many areas defined")
            
            for i, area in enumerate(areas):
                if not isinstance(area, dict):
                    raise ValueError(f"Area {i} must be a dictionary")
                
                # Validate area fields
                area_fields = {
                    'coords': (list, tuple),
                    'name': str,
                    'hotkey': (str, type(None)),
                    'preprocess': bool,
                    'voice': str,
                    'speed': str,
                    'settings': dict
                }
                
                for field, expected_type in area_fields.items():
                    if field in area:
                        value = area[field]
                        if not isinstance(value, expected_type):
                            raise ValueError(f"Invalid type for area {i} {field}")
                        
                        # Additional validation
                        if field == 'name':
                            if not value or len(value) > 100:  # Reasonable name length
                                raise ValueError(f"Invalid area name in area {i}")
                            # Sanitize name - remove potentially dangerous characters
                            if any(char in value for char in ['<', '>', '"', "'", '&']):
                                raise ValueError(f"Area name contains invalid characters in area {i}")
                        elif field == 'coords':
                            if len(value) != 4:
                                raise ValueError(f"Coordinates must have exactly 4 values in area {i}")
                            for coord in value:
                                if not isinstance(coord, (int, float)) or coord < 0 or coord > 10000:
                                    raise ValueError(f"Invalid coordinate value in area {i}")
                        elif field == 'hotkey':
                            if value is not None and (not isinstance(value, str) or len(value) > 50):
                                raise ValueError(f"Invalid hotkey in area {i}")
                        elif field == 'voice' or field == 'speed':
                            if len(value) > 100:  # Reasonable limit
                                raise ValueError(f"Invalid {field} value in area {i}")
                        elif field == 'settings':
                            # Validate settings dictionary
                            if len(value) > 20:  # Reasonable number of settings
                                raise ValueError(f"Too many settings in area {i}")
                            for key, val in value.items():
                                if not isinstance(key, str) or len(key) > 50:
                                    raise ValueError(f"Invalid setting key in area {i}")
                                if not isinstance(val, (str, int, float, bool, type(None))) or (isinstance(val, str) and len(val) > 100):
                                    raise ValueError(f"Invalid setting value in area {i}")
        
        return True

    def _load_layout_file(self, file_path):
        """Internal method to load a layout file"""
        if file_path:
            try:
                # Basic file validation
                if not os.path.exists(file_path):
                    raise FileNotFoundError("Layout file does not exist")
                
                # Check file size (prevent loading extremely large files)
                file_size = os.path.getsize(file_path)
                if file_size > 10 * 1024 * 1024:  # 10MB limit
                    raise ValueError("Layout file is too large (max 10MB)")
                
                with open(file_path, 'r', encoding='utf-8') as f:
                    try:
                        layout = json.load(f)
                    except json.JSONDecodeError as e:
                        raise ValueError(f"Invalid JSON format: {str(e)}")
                
                # Validate the loaded data
                self._validate_layout_data(layout)
                
                # Only set the file path and reset unsaved changes AFTER successful validation
                self.layout_file.set(file_path)
                self._has_unsaved_changes = False
                    
                # Clear only user-added areas and processing settings (keep permanent area)
                if self.areas:
                    # Always keep the first (permanent) area
                    for area in self.areas[1:]:
                        # Clean up hotkeys before destroying the area
                        hotkey_button = area[1]
                        if hasattr(hotkey_button, 'keyboard_hook'):
                            try:
                                if hotkey_button.keyboard_hook:
                                    # Check if it's a callable (function) or a hook ID
                                    if callable(hotkey_button.keyboard_hook):
                                        # It's a function, try to unhook it
                                        try:
                                            keyboard.unhook(hotkey_button.keyboard_hook)
                                        except Exception:
                                            pass
                                    else:
                                        # Check if this is a custom ctrl hook, on_press_key hook, or a regular add_hotkey hook
                                        try:
                                            if hasattr(hotkey_button.keyboard_hook, 'remove'):
                                                # This is an add_hotkey hook
                                                keyboard.remove_hotkey(hotkey_button.keyboard_hook)
                                            elif hasattr(hotkey_button.keyboard_hook, 'unhook'):
                                                # This is an on_press_key hook
                                                hotkey_button.keyboard_hook.unhook()
                                            else:
                                                # This is a custom on_press hook
                                                keyboard.unhook(hotkey_button.keyboard_hook)
                                        except Exception:
                                            # Fallback to unhook if all methods fail
                                            keyboard.unhook(hotkey_button.keyboard_hook)
                            except Exception as e:
                                print(f"Warning: Error cleaning up keyboard hook: {e}")
                        if hasattr(hotkey_button, 'mouse_hook'):
                            try:
                                if hotkey_button.mouse_hook:
                                    # Check if it's a callable (function) or a hook ID
                                    if callable(hotkey_button.mouse_hook):
                                        # It's a function, try to unhook it
                                        try:
                                            mouse.unhook(hotkey_button.mouse_hook)
                                        except Exception:
                                            pass
                                    else:
                                        # It's a hook ID, try to remove the hotkey first
                                        try:
                                            if hasattr(hotkey_button, 'mouse_hook_id') and hotkey_button.mouse_hook_id:
                                                mouse.unhook(hotkey_button.mouse_hook_id)
                                        except Exception:
                                            pass
                            except Exception as e:
                                print(f"Warning: Error cleaning up mouse hook: {e}")
                        area[0].destroy()
                    self.areas = self.areas[:1]
                self.processing_settings.clear()

                save_version = layout.get("version", "0.0")
                current_version = "0.5"

                if tuple(map(int, save_version.split('.'))) < tuple(map(int, current_version.split('.'))):
                    messagebox.showerror("Error", "Cannot load older version save files.")
                    return

                # Extract just the filename from the full path
                file_name = os.path.basename(file_path)  # Just keep the original filename
                
                # Show feedback in status label
                if hasattr(self, '_feedback_timer') and self._feedback_timer:
                    self.root.after_cancel(self._feedback_timer)
                
                # Show load success message
                self.status_label.config(text=f"Layout loaded: {file_name}")
                self._feedback_timer = self.root.after(2000, lambda: self.status_label.config(text=""))

                # Actually load the layout data
                self.layout_file.set(file_name)
                self.bad_word_list.set(layout.get("bad_word_list", ""))
                self.ignore_usernames_var.set(layout.get("ignore_usernames", False))
                self.ignore_previous_var.set(layout.get("ignore_previous", False))
                self.ignore_gibberish_var.set(layout.get("ignore_gibberish", False))
                self.pause_at_punctuation_var.set(layout.get("pause_at_punctuation", False))
                self.better_unit_detection_var.set(layout.get("better_unit_detection", False))
                self.read_game_units_var.set(layout.get("read_game_units", False))
                self.fullscreen_mode_var.set(layout.get("fullscreen_mode", False))
                
                # Load volume setting
                saved_volume = layout.get("volume", "100")
                self.volume.set(saved_volume)
                try:
                    self.speaker.Volume = int(saved_volume)
                    print(f"Loaded volume setting: {saved_volume}%")
                except ValueError:
                    print("Invalid volume in save file, defaulting to 100%")
                    self.volume.set("100")
                    self.speaker.Volume = 100
                
                # Clean up existing areas and unhook all hotkeys
                # Clean up images
                for image in self.latest_images.values():
                    try:
                        image.close()
                    except:
                        pass
                self.latest_images.clear()
                
                # Unhook all existing hotkeys
                keyboard.unhook_all()
                mouse.unhook_all()
                
                # Set up stop hotkey first
                saved_stop_hotkey = layout.get("stop_hotkey")
                if saved_stop_hotkey:
                    self.stop_hotkey = saved_stop_hotkey
                    self.stop_hotkey_button.mock_button = type('MockButton', (), {
                        'hotkey': saved_stop_hotkey,
                        'is_stop_button': True
                    })
                    self.setup_hotkey(self.stop_hotkey_button.mock_button, None)  # Pass None as area_frame for stop hotkey
                    
                    # Update the button text
                    display_name = saved_stop_hotkey.replace('numpad ', 'NUMPAD ').replace('num_', 'num:') \
                                               .replace('ctrl','CTRL') \
                                               .replace('left alt','L-ALT').replace('right alt','R-ALT') \
                                               .replace('left shift','L-SHIFT').replace('right shift','R-SHIFT') \
                                               .replace('windows','WIN') \
                                               .replace('multiply', '*').replace('add', '+').replace('subtract', '-').replace('divide', '/')
                    self.stop_hotkey_button.config(text=f"Stop Hotkey: [ {display_name.upper()} ]")
                    print(f"Loaded Stop hotkey: {saved_stop_hotkey}")

                # --- Handle Auto Read hotkey ---
                auto_read_hotkey = None
                if self.areas and hasattr(self.areas[0][1], 'hotkey'):
                    auto_read_hotkey = self.areas[0][1].hotkey
                    # Clear the existing auto-read hotkey before loading new ones
                    if auto_read_hotkey:
                        try:
                            if hasattr(self.areas[0][1], 'hotkey_id') and hasattr(self.areas[0][1].hotkey_id, 'remove'):
                                # This is an add_hotkey hook
                                keyboard.remove_hotkey(self.areas[0][1].hotkey_id)
                            else:
                                # This is a custom on_press hook or doesn't exist
                                pass
                        except (KeyError, AttributeError):
                            pass
                        self.areas[0][1].hotkey = None
                        self.areas[0][1].config(text="Set Hotkey")
                
                # Check for conflicts with the auto-read hotkey
                conflict_area_name = None
                for area_info in layout.get("areas", []):
                    if auto_read_hotkey and area_info.get("hotkey") == auto_read_hotkey:
                        conflict_area_name = area_info["name"]
                        break
                # Load all the areas (under permanent area)
                areas_loaded = False
                for area_info in layout.get("areas", []):
                    # Create a new area using add_read_area (removable, editable, normal name)
                    self.add_read_area(removable=True, editable_name=True, area_name=area_info["name"])
                    
                    # Get the newly created area (last one in the list)
                    area_frame, hotkey_button, set_area_button, area_name_var, preprocess_var, voice_var, speed_var = self.areas[-1]
                    areas_loaded = True
                    
                    # Set the area coordinates
                    area_frame.area_coords = area_info["coords"]
                    
                    # Set the hotkey if it exists
                    if area_info["hotkey"]:
                        hotkey_button.hotkey = area_info["hotkey"]
                        display_name = area_info["hotkey"].replace('num_', 'num:').replace('multiply', '*').replace('add', '+').replace('subtract', '-').replace('divide', '/') if area_info["hotkey"].startswith('num_') else area_info["hotkey"].replace('multiply', '*').replace('add', '+').replace('subtract', '-').replace('divide', '/')
                        hotkey_button.config(text=f"Hotkey: [ {display_name.upper()} ]")
                        self.setup_hotkey(hotkey_button, area_frame)
                        
                        # Warn about special characters that may cause cross-language issues
                        if is_special_character(area_info["hotkey"]):
                            alternative = suggest_alternative_key(area_info["hotkey"])
                            if alternative:
                                print(f"WARNING: Area '{area_info['name']}' uses special character '{area_info['hotkey']}' which may not work on different keyboard layouts.")
                                print(f"Consider changing it to '{alternative}' for better compatibility.")
                    
                    # Set preprocessing and voice settings
                    preprocess_var.set(area_info.get("preprocess", False))
                    # Check if the saved voice exists in current SAPI voices and convert to display name
                    if hasattr(self, 'voices') and self.voices:
                        try:
                            saved_voice = area_info.get("voice")
                            if saved_voice and saved_voice != "Select Voice":
                                # Find the display name that corresponds to this full voice name
                                display_name = 'Select Voice'
                                for i, voice in enumerate(self.voices, 1):
                                    if hasattr(voice, 'GetDescription') and voice.GetDescription() == saved_voice:
                                        # Create the same abbreviated display name logic WITH numbering
                                        full_name = voice.GetDescription()
                                        if "Microsoft" in full_name and " - " in full_name:
                                            parts = full_name.split(" - ")
                                            if len(parts) == 2:
                                                voice_part = parts[0].replace("Microsoft ", "")
                                                lang_part = parts[1]
                                                display_name = f"{i}. {voice_part} ({lang_part})"
                                            else:
                                                display_name = f"{i}. {full_name}"
                                        elif " - " in full_name:
                                            parts = full_name.split(" - ")
                                            if len(parts) == 2:
                                                display_name = f"{i}. {parts[0]} ({lang_part})"
                                            else:
                                                display_name = f"{i}. {full_name}"
                                        else:
                                            display_name = f"{i}. {full_name}"
                                        break
                                
                                voice_var.set(display_name)
                                # Set the full name for the voice variable
                                voice_var._full_name = saved_voice
                            else:
                                voice_var.set('Select Voice')
                        except Exception as e:
                            print(f"Warning: Could not validate voice: {e}")
                            voice_var.set('Select Voice')
                    else:
                        voice_var.set('Select Voice')
                    speed_var.set(area_info.get("speed", "1.0"))
                    
                    # Load and store image processing settings
                    if "settings" in area_info:
                        self.processing_settings[area_info["name"]] = area_info["settings"].copy()
                        print(f"Loaded image processing settings for area: {area_info['name']}")
                        
                # Update preferred sizes during load
                self.resize_window()

                # Only process the last loaded area if any areas were loaded
                if areas_loaded and len(self.areas) > 1:
                    # Get coordinates from the last loaded area
                    x1, y1, x2, y2 = area_frame.area_coords
                    screenshot = capture_screen_area(x1, y1, x2, y2)

                    # Store original or processed image based on settings
                    if preprocess_var.get() and area_info["name"] in self.processing_settings:
                        settings = self.processing_settings[area_info["name"]]
                        processed_image = preprocess_image(
                            screenshot,
                            brightness=settings.get('brightness', 1.0),
                            contrast=settings.get('contrast', 1.0),
                            saturation=settings.get('saturation', 1.0),
                            sharpness=settings.get('sharpness', 1.0),
                            blur=settings.get('blur', 0.0),
                            threshold=settings.get('threshold', None) if settings.get('threshold_enabled', False) else None,
                            hue=settings.get('hue', 0.0),
                            exposure=settings.get('exposure', 1.0)
                        )
                        self.latest_images[area_name_var.get()] = processed_image
                    else:
                        self.latest_images[area_name_var.get()] = screenshot
                # --- Handle Auto Read hotkey state after loading ---
                # If no conflict and auto-read hotkey exists, re-register it
                if not conflict_area_name and auto_read_hotkey and self.areas and hasattr(self.areas[0][1], 'hotkey'):
                    try:
                        # Re-register the auto-read hotkey
                        self.areas[0][1].hotkey = auto_read_hotkey
                        display_name = auto_read_hotkey.replace('num_', 'num:').replace('multiply', '*').replace('add', '+').replace('subtract', '-').replace('divide', '/') if auto_read_hotkey.startswith('num_') else auto_read_hotkey.replace('multiply', '*').replace('add', '+').replace('subtract', '-').replace('divide', '/')
                        self.areas[0][1].config(text=f"Hotkey: [ {display_name.upper()} ]")
                        # Re-setup the hotkey
                        self.setup_hotkey(self.areas[0][1], self.areas[0][0])
                        print(f"Re-registered Auto Read hotkey: {auto_read_hotkey}")
                    except Exception as e:
                        print(f"Error re-registering Auto Read hotkey: {e}")
                # Show popup if conflict detected
                elif conflict_area_name:
                    hotkey_val = auto_read_hotkey if auto_read_hotkey else "?"
                    messagebox.showinfo(
                        "Hotkey Conflict",
                        f"Detected same Hotkey!\n\nAuto Read Hotkey = {hotkey_val}\n{conflict_area_name} Hotkey = {hotkey_val}\n\nPlease set a new hotkey for AutoRead if you still want this function.")
                    # Clear the Auto Read hotkey registration
                    if self.areas and hasattr(self.areas[0][1], 'hotkey'):
                        try:
                            if hasattr(self.areas[0][1], 'hotkey_id') and hasattr(self.areas[0][1].hotkey_id, 'remove'):
                                # This is an add_hotkey hook
                                keyboard.remove_hotkey(self.areas[0][1].hotkey_id)
                            else:
                                # This is a custom on_press hook or doesn't exist
                                pass
                        except (KeyError, AttributeError):
                            pass
                        self.areas[0][1].hotkey = None
                        self.areas[0][1].config(text="Set Hotkey")

                print(f"Layout loaded from {file_path}\n--------------------------")
                
                # Force-resize the window to fit the newly loaded layout
                self.resize_window(force=True)
                
            except (ValueError, FileNotFoundError) as e:
                # Handle validation and file errors with specific messages
                messagebox.showerror("Invalid Save File", f"The save file appears to be corrupted or malicious:\n\n{str(e)}\n\nPlease use a valid save file.")
                print(f"Security validation failed for layout file: {e}")
            except json.JSONDecodeError as e:
                messagebox.showerror("Invalid Save File", f"The save file contains invalid JSON format:\n\n{str(e)}\n\nThe file may be corrupted.")
                print(f"JSON decode error in layout file: {e}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load layout: {str(e)}")
                print(f"Error loading layout: {e}")

    def validate_speed_key(self, event, speed_var):
        """Additional validation for speed entry key presses"""
        if event.char.isdigit() or event.keysym in ('BackSpace', 'Delete', 'Left', 'Right'):
            return
        return 'break'

    def setup_hotkey(self, button, area_frame):
        """Enhanced hotkey setup supporting multi-key combinations like Ctrl+Shift+F1"""
        try:
            # Clean up any existing hooks for this button first
            if hasattr(button, 'keyboard_hook'):
                try:
                    # Check if this is a custom ctrl hook or a regular add_hotkey hook
                    if hasattr(button.keyboard_hook, 'remove'):
                        # This is an add_hotkey hook
                        keyboard.remove_hotkey(button.keyboard_hook)
                    else:
                        # This is a custom on_press hook
                        keyboard.unhook(button.keyboard_hook)
                    delattr(button, 'keyboard_hook')
                except Exception as e:
                    print(f"Error cleaning up keyboard hook: {e}")
            
            if hasattr(button, 'mouse_hook_id'):
                try:
                    mouse.unhook(button.mouse_hook_id)
                    delattr(button, 'mouse_hook_id')
                except Exception as e:
                    print(f"Error cleaning up mouse hook ID: {e}")
            if hasattr(button, 'mouse_hook'):
                try:
                    delattr(button, 'mouse_hook')
                except Exception as e:
                    print(f"Error cleaning up mouse hook function: {e}")
            
            # Store area_frame if this is not a stop button
            if not hasattr(button, 'is_stop_button') and area_frame is not None:
                button.area_frame = area_frame
                
            # Only proceed if we have a valid hotkey
            if not hasattr(button, 'hotkey') or not button.hotkey:
                print(f"No hotkey set for button: {button}")
                return False
                
            print(f"Setting up hotkey for: {button.hotkey}")
            
            # Define the hotkey handler
            def hotkey_handler():
                try:
                    if self.setting_hotkey:
                        return
                    
                    # Check if the button itself is still valid
                    if not hasattr(button, 'hotkey') or not button.hotkey:
                        print(f"Warning: Hotkey triggered for invalid button, ignoring")
                        return
                    
                    # Debug: Log which hotkey was triggered
                    if hasattr(button, 'hotkey'):
                        print(f"Hotkey triggered: '{button.hotkey}' (type: {type(button.hotkey).__name__}, bytes: {button.hotkey.encode('utf-8')})")
                        
                    # Handle stop button
                    if hasattr(button, 'is_stop_button'):
                        self.stop_speaking()
                        return
                        
                    # Handle Auto Read area
                    if hasattr(button, 'area_frame') and button.area_frame:
                        # Check if the area still exists in the areas list
                        area_exists = False
                        for area in self.areas:
                            if area[0] is button.area_frame:
                                area_exists = True
                                break
                        
                        if not area_exists:
                            print(f"Warning: Hotkey triggered for removed area, ignoring")
                            return
                        
                        area_info = self._get_area_info(button)
                        if area_info and area_info.get('name') == "Auto Read":
                            self.set_area(
                                area_info['frame'], 
                                area_info['name_var'], 
                                area_info['set_area_btn'])
                            return
                        
                        # Handle regular areas
                        if hasattr(button, 'area_frame'):
                            self.stop_speaking()
                            threading.Thread(
                                target=self.read_area, 
                                args=(button.area_frame,), 
                                daemon=True
                            ).start()
                            
                except Exception as e:
                    print(f"Error in hotkey handler: {e}")
            
            # Set up the appropriate hook based on hotkey type
            if button.hotkey.startswith('button'):
                try:
                    # For mouse buttons, we need to use mouse.hook() and track button states
                    # Extract button identifier from hotkey (e.g., "button1" -> 1, "buttonx" -> "x")
                    button_identifier = button.hotkey.replace('button', '')
                    
                    # Check if we have the original button identifier stored
                    if hasattr(button, 'original_button_id'):
                        print(f"Setting up mouse hook for button '{button.original_button_id}' (hotkey: {button.hotkey})")
                    else:
                        print(f"Setting up mouse hook for button '{button_identifier}' (hotkey: {button.hotkey})")
                    
                    # Create a mouse event handler for this specific button
                    def mouse_button_handler(event):
                        # Only process ButtonEvent objects, ignore MoveEvent, WheelEvent, etc.
                        if not isinstance(event, mouse.ButtonEvent):
                            return
                        
                        # Check if this is the right button by comparing with original_button_id
                        button_matches = False
                        
                        if hasattr(button, 'original_button_id'):
                            # Use the original button identifier for comparison
                            button_matches = (event.button == button.original_button_id)
                        else:
                            # Fall back to comparing with the extracted identifier
                            button_matches = (event.button == button_identifier)
                        
                        if (button_matches and 
                            hasattr(event, 'event_type') and event.event_type == mouse.DOWN):
                            print(f"Mouse button '{event.button}' matched for hotkey '{button.hotkey}'")
                            hotkey_handler()
                    
                    # Store the handler function so we can unhook it later
                    button.mouse_hook = mouse_button_handler
                    button.mouse_hook_id = mouse.hook(mouse_button_handler)
                    print(f"Mouse hook set up for {button.hotkey}")
                except Exception as e:
                    print(f"Error setting up mouse hook: {e}")
                    return False
            elif button.hotkey.startswith('controller_'):
                try:
                    # For controller buttons, we need to monitor controller input
                    # Extract button identifier from hotkey (e.g., "controller_A" -> "A")
                    button_identifier = button.hotkey.replace('controller_', '')
                    
                    print(f"Setting up controller hook for button '{button_identifier}' (hotkey: {button.hotkey})")
                    
                    # Create a controller event handler for this specific button
                    def controller_button_handler():
                        print(f"Controller button '{button_identifier}' pressed for hotkey '{button.hotkey}'")
                        hotkey_handler()
                    
                    # Store the handler function so we can access it later
                    button.controller_hook = controller_button_handler
                    
                    # Start controller monitoring if not already running
                    if not self.controller_handler.running:
                        self.controller_handler.start_monitoring()
                    
                    print(f"Controller hook set up for {button.hotkey}")
                    return True
                except Exception as e:
                    print(f"Error setting up controller hook: {e}")
                    return False
            else:
                try:
                    # Validate the hotkey before setting it up
                    if not button.hotkey or button.hotkey.strip() == '':
                        print(f"Error: Empty hotkey for button")
                        return False
                    
                    # Enhanced validation for multi-key combinations
                    hotkey_parts = button.hotkey.split('+')
                    if len(hotkey_parts) > 1:
                        print(f"Multi-key hotkey detected: {button.hotkey} ({len(hotkey_parts)} parts)")
                        
                        # Validate each part of the combination
                        valid_parts = []
                        for part in hotkey_parts:
                            part = part.strip().lower()
                            if part in ['ctrl', 'shift', 'alt', 'left alt', 'right alt', 'windows']:
                                valid_parts.append(part)
                            elif part.startswith('f') and part[1:].isdigit() and 1 <= int(part[1:]) <= 24:
                                valid_parts.append(part)  # Function keys F1-F24
                            elif part.startswith('num_'):
                                valid_parts.append(part)  # Numpad keys
                            elif len(part) == 1 and part.isalnum():
                                valid_parts.append(part)  # Single character keys
                            elif part in ['space', 'enter', 'tab', 'backspace', 'delete', 'insert', 'home', 'end', 'page up', 'page down']:
                                valid_parts.append(part)  # Special keys
                            else:
                                print(f"Warning: Unknown hotkey part '{part}' in '{button.hotkey}'")
                                valid_parts.append(part)  # Still allow it, but warn
                        
                        if len(valid_parts) != len(hotkey_parts):
                            print(f"Some hotkey parts could not be validated")
                    
                    # Check for special characters and warn about potential issues
                    is_special = is_special_character(button.hotkey)
                    if is_special:
                        print(f"WARNING: Hotkey '{button.hotkey}' contains special characters that may cause issues")
                        alternative = suggest_alternative_key(button.hotkey)
                        if alternative:
                            print(f"Consider using '{alternative}' instead for better compatibility")
                        
                        # Check if this hotkey would conflict with existing ones
                        if not self._check_hotkey_uniqueness(button.hotkey, button):
                            print(f"WARNING: Hotkey '{button.hotkey}' conflicts with an existing hotkey")
                            print(f"This may cause both hotkeys to trigger the same action")
                            print(f"Please choose a different hotkey to avoid conflicts")
                            return False
                    
                    # Check for problematic characters in numpad keys
                    if button.hotkey.startswith('num_'):
                        if len(button.hotkey) < 5:  # num_ + at least 1 character
                            print(f"Error: Invalid numpad hotkey format: '{button.hotkey}'")
                            return False
                        
                        # Additional validation for numpad keys
                        numpad_part = button.hotkey[4:]  # Get the part after 'num_'
                        if not numpad_part or numpad_part.strip() == '':
                            print(f"Error: Empty numpad key part in hotkey: '{button.hotkey}'")
                            return False
                        
                        # Check if the numpad key is valid
                        valid_numpad_keys = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'multiply', 'add', 'subtract', '.', 'divide', 'enter']
                        if numpad_part not in valid_numpad_keys:
                            print(f"Error: Invalid numpad key '{numpad_part}' in hotkey: '{button.hotkey}'")
                            print(f"Valid numpad keys: {valid_numpad_keys}")
                            return False
                        
                        # Note: Special characters (*, +, -, /) are now handled by using descriptive names
                        # in the numpad_scan_codes dictionary (multiply, add, subtract, divide)
                    
                    # Set up the keyboard hook (preserving original hotkey)
                    try:
                        # Debug: Log the exact hotkey being registered
                        print(f"Registering hotkey: '{button.hotkey}' (length: {len(button.hotkey)}, bytes: {button.hotkey.encode('utf-8')})")
                        
                        # Special handling for ctrl key to prevent cross-activation
                        if button.hotkey == 'ctrl':
                            # Use both scan codes to catch either left or right Ctrl
                            button.keyboard_hook = keyboard.add_hotkey('ctrl', hotkey_handler)
                            print(f"Ctrl hotkey hook set up for both left and right Ctrl")
                        else:
                            # Use add_hotkey for other hotkeys
                            button.keyboard_hook = keyboard.add_hotkey(button.hotkey, hotkey_handler)
                            
                            if len(hotkey_parts) > 1:
                                print(f"Multi-key hotkey registered successfully: '{button.hotkey}'")
                            elif is_special:
                                print(f"Keyboard hook set up for special character hotkey: '{button.hotkey}'")
                            else:
                                print(f"Keyboard hook set up for '{button.hotkey}'")
                    except Exception as e:
                        print(f"Error setting up keyboard hook: {e}")
                        print(f"Hotkey value: '{button.hotkey}' (length: {len(button.hotkey) if button.hotkey else 0})")
                        
                        # Try to provide helpful error messages for common issues
                        if "invalid" in str(e).lower() or "unknown" in str(e).lower():
                            print(f"This might be due to an unsupported key combination")
                            print(f"Try using simpler combinations or check key names")
                        elif "already" in str(e).lower() or "exists" in str(e).lower():
                            print(f"This hotkey might already be registered elsewhere")
                        
                        return False
                except Exception as e:
                    print(f"Error setting up keyboard hook: {e}")
                    print(f"Hotkey value: '{button.hotkey}' (length: {len(button.hotkey) if button.hotkey else 0})")
                    return False
                    
            return True
            
        except Exception as e:
            print(f"Error in setup_hotkey: {e}")
            return False
            
    def _convert_numpad_to_display(self, hotkey):
        """Convert numpad hotkey names to display symbols"""
        if not hotkey or not hotkey.startswith('num_'):
            return hotkey
        
        numpad_part = hotkey[4:]  # Get the part after 'num_'
        symbol_map = {
            'multiply': '*',
            'add': '+',
            'subtract': '-',
            'divide': '/',
            '0': '0', '1': '1', '2': '2', '3': '3', '4': '4',
            '5': '5', '6': '6', '7': '7', '8': '8', '9': '9',
            '.': '.', 'enter': 'Enter'
        }
        
        if numpad_part in symbol_map:
            return f"num:{symbol_map[numpad_part]}"
        return hotkey

    def _get_display_hotkey(self, button):
        """Get the display text for a hotkey"""
        if not hasattr(button, 'hotkey') or not button.hotkey:
            return "No Hotkey"
        
        return button.hotkey

    def _check_hotkey_uniqueness(self, new_hotkey, exclude_button=None):
        """Check if a hotkey is unique among all registered hotkeys"""
        if not new_hotkey:
            return True
        
        for area in self.areas:
            if area[1] is exclude_button:
                continue
            if hasattr(area[1], 'hotkey') and area[1].hotkey:
                if area[1].hotkey == new_hotkey:
                    return False
        return True

    def _normalize_hotkey(self, hotkey):
        """Normalize hotkey to prevent character encoding issues (for reference only)"""
        if not hotkey:
            return hotkey
        
        # Convert to lowercase for consistency
        normalized = hotkey.lower()
        
        # Handle common character normalizations that cause conflicts
        char_map = {
            'å': 'a',
            'ä': 'a',
            'ö': 'o',
            '¨': 'u',
            '´': "'",
            '`': "'",
            '~': '~',
            '^': '^'
        }
        
        for special_char, normal_char in char_map.items():
            if special_char in normalized:
                print(f"Normalizing '{special_char}' to '{normal_char}' in hotkey '{hotkey}'")
                normalized = normalized.replace(special_char, normal_char)
        
        return normalized

    def _cleanup_hooks(self, button):
        """Simple cleanup method for existing hooks"""
        try:
            # Clean up mouse hook if it exists
            if hasattr(button, 'mouse_hook_id'):
                try:
                    if button.mouse_hook_id:
                        mouse.unhook(button.mouse_hook_id)
                except Exception as e:
                    print(f"Warning: Error cleaning up mouse hook ID: {e}")
                finally:
                    # Always set to None to prevent future errors
                    button.mouse_hook_id = None
            
            if hasattr(button, 'mouse_hook'):
                try:
                    # Clean up the handler function reference
                    button.mouse_hook = None
                except Exception as e:
                    print(f"Warning: Error cleaning up mouse hook function: {e}")
            
            # Clean up keyboard hook if it exists
            if hasattr(button, 'keyboard_hook'):
                try:
                    if button.keyboard_hook:
                        # Check if it's a callable (function) or a hook ID
                        if callable(button.keyboard_hook):
                            # It's a function, try to unhook it
                            try:
                                keyboard.unhook(button.keyboard_hook)
                            except Exception:
                                pass
                        else:
                            # Check if this is a custom ctrl hook or a regular add_hotkey hook
                            try:
                                if hasattr(button.keyboard_hook, 'remove'):
                                    # This is an add_hotkey hook
                                    keyboard.remove_hotkey(button.keyboard_hook)
                                else:
                                    # This is a custom on_press hook
                                    keyboard.unhook(button.keyboard_hook)
                            except Exception:
                                # Fallback to unhook if both methods fail
                                keyboard.unhook(button.keyboard_hook)
                except Exception as e:
                    print(f"Warning: Error cleaning up keyboard hook: {e}")
                finally:
                    # Always set to None to prevent future errors
                    button.keyboard_hook = None
            
            # Clean up controller hook if it exists
            if hasattr(button, 'controller_hook'):
                try:
                    button.controller_hook = None
                except Exception as e:
                    print(f"Warning: Error cleaning up controller hook: {e}")
                        
        except Exception as e:
            print(f"Unexpected error in _cleanup_hooks: {e}")
            # Make sure we don't leave any attributes behind
            for attr in ['mouse_hook', 'keyboard_hook', 'controller_hook']:
                if hasattr(button, attr):
                    try:
                        delattr(button, attr)
                    except:
                        pass


            
    def _get_area_info(self, button):
        """Helper method to get area information for a button"""
        for area in self.areas:
            if area[1] is button:
                return {
                    'frame': area[0],
                    'name': area[3].get() if hasattr(area[3], 'get') else None,
                    'name_var': area[3],
                    'set_area_btn': area[2]
                }
        return None

    def read_area(self, area_frame):
        # First check if this is a stop button - if so, return immediately
        if area_frame is None:  # Stop button passes None as area_frame
            return

        # Check if the area frame still exists in the areas list
        area_exists = False
        for area in self.areas:
            if area[0] is area_frame:
                area_exists = True
                break
        
        if not area_exists:
            print(f"Warning: Attempted to read removed area, ignoring")
            return

        if not hasattr(area_frame, 'area_coords'):
            # Suppress error for Auto Read area
            area_info = None
            for area in self.areas:
                if area[0] is area_frame:
                    area_info = area
                    break
            if area_info and area_info[3].get() == "Auto Read":
                return
            messagebox.showerror("Error", "No area coordinates set. Click Set Area to set one.")
            return

        # Ensure speaker is initialized
        if not self.speaker:
            try:
                self.speaker = win32com.client.Dispatch("SAPI.SpVoice")
                self.speaker.Volume = int(self.volume.get())
            except Exception as e:
                print(f"Error initializing speaker: {e}")
                return

        # Get area info first
        area_info = None
        for area in self.areas:
            if area[0] is area_frame:
                area_info = area
                break
        
        if not area_info:
            print(f"Error: Could not determine area name for frame {area_frame}")
            return

        area_name = area_info[3].get()
        self.latest_area_name.set(area_name)
        voice_var = area_info[5]
        speed_var = area_info[6]
        preprocess = area_info[4].get()

        # Show processing feedback
        self.show_processing_feedback(area_name)

        # Capture screenshot
        x1, y1, x2, y2 = area_frame.area_coords
        screenshot = capture_screen_area(x1, y1, x2, y2)
        
        # Store original or processed image based on settings
        if preprocess and area_name in self.processing_settings:
            settings = self.processing_settings[area_name]
            processed_image = preprocess_image(
                screenshot,
                brightness=settings.get('brightness', 1.0),
                contrast=settings.get('contrast', 1.0),
                saturation=settings.get('saturation', 1.0),
                sharpness=settings.get('sharpness', 1.0),
                blur=settings.get('blur', 0.0),
                threshold=settings.get('threshold', None) if settings.get('threshold_enabled', False) else None,
                hue=settings.get('hue', 0.0),
                exposure=settings.get('exposure', 1.0)
            )
            self.latest_images[area_name] = processed_image
            # Use processed image for OCR
            text = pytesseract.image_to_string(processed_image)
            print("Image preprocessing applied.")
        else:
            self.latest_images[area_name] = screenshot
            # Use original image for OCR
            text = pytesseract.image_to_string(screenshot)

        # --- Better measurement unit detection logic ---
        import re
        if hasattr(self, 'better_unit_detection_var') and self.better_unit_detection_var.get():
            unit_map = {
                'l': 'Liters',
                'm': 'Meters',
                'in': 'Inches',
                'ml': 'Milliliters',
                'gal': 'Gallons',
                'g': 'Grams',
                'lb': 'Pounds',
                'ib': 'Pounds',  # Treat 'ib' as 'Pounds' due to OCR confusion
                'c': 'Celsius',
                'f': 'Fahrenheit',
                # Money units
                'kr': 'Crowns',
                'eur': 'Euros',
                'usd': 'US Dollars',
                'sek': 'Swedish Crowns',
                'nok': 'Norwegian Crowns',
                'dkk': 'Danish Crowns',
                '£': 'Pounds Sterling',
            }
            pattern = re.compile(r'(?<!\w)(\d+(?:\.\d+)?)(\s*)(l|m|in|ml|gal|g|lb|ib|c|f|kr|eur|usd|sek|nok|dkk|£)(?!\w)', re.IGNORECASE)
            def repl(match):
                value = match.group(1)
                space = match.group(2)
                unit = match.group(3).lower()
                if unit in ['lb', 'ib']:
                    return f"{value}{space}Pounds"
                if unit == '£':
                    return f"{value}{space}Pounds Sterling"
                return f"{value}{space}{unit_map.get(unit, unit)}"
            text = pattern.sub(repl, text)

        # --- Read game units logic ---
        if hasattr(self, 'read_game_units_var') and self.read_game_units_var.get():
            game_unit_map = self.game_units
            
            # Add default mappings for common game units
            default_mappings = {
                'xp': 'Experience Points',
                'hp': 'Health Points',
                'mp': 'Mana Points',
                'gp': 'Gold Pieces',
                'pp': 'Platinum Pieces',
                'sp': 'Skill Points',
                'ep': 'Energy Points',
                'ap': 'Action Points',
                'bp': 'Battle Points',
                'lp': 'Loyalty Points',
                'cp': 'Challenge Points',
                'vp': 'Victory Points',
                'rp': 'Reputation Points',
                'tp': 'Talent Points',
                'ar': 'Armor Rating',
                'dmg': 'Damage',
                'dps': 'Damage Per Second',
                'def': 'Defense',
                'mat': 'Materials',
                'exp': 'Exploration Points',
                '§': 'Simoliance',
                'v-bucks': 'Virtual Bucks',
                'r$': 'Robux',
                'nmt': 'Nook Miles Tickets',
                'be': 'Blue Essence',
                'radianite': 'Radianite Points',
                'ow coins': 'Overwatch Coins',
                '₽': 'PokeDollars',
                '€$': 'Eurodollars',
                'z': 'Zenny',
                'l': 'Lunas',
                'e': 'Eve',
                'i': 'Isk',
                'j': 'Jewel',
                'sc': 'Star Coins',
                'o2': 'Oxygen',
                'pu': 'Power Units',
                'mc': 'Mana Crystals',
                'es': 'Essence',
                'sh': 'Shards',
                'st': 'Stars',
                'mu': 'Munny',
                'b': 'Bolts',
                'r': 'Rings',
                'ca': 'Caps',
                'rns': 'Runes',
                'sl': 'Souls',
                'fav': 'Favor',
                'am': 'Amber',
                'cc': 'Crystal Cores',
                'fg': 'Fragments'
            }
            
            # Update game units with default mappings if they don't exist
            for key, value in default_mappings.items():
                if key not in game_unit_map:
                    game_unit_map[key] = value
            # Sort by length descending to match longer units first (e.g., 'gp' before 'g')
            sorted_units = sorted(game_unit_map.keys(), key=len, reverse=True)
            # Build regex pattern for all units (word boundaries, case-insensitive)
            pattern = re.compile(r'(?<!\w)(\d+(?:\.\d+)?)(\s*)(' + '|'.join(map(re.escape, sorted_units)) + r')(?!\w)', re.IGNORECASE)
            def game_repl(match):
                value = match.group(1)
                space = match.group(2)
                unit = match.group(3).lower()
                return f"{value}{space}{game_unit_map.get(unit, unit)}"
            text = pattern.sub(game_repl, text)

        print(f"Processing Area with name '{area_name}' Output Text: \n {text}\n--------------------------")

        # Handle text history if ignore previous is enabled
        if self.ignore_previous_var.get():
            # Limit history size to prevent memory growth
            max_history_size = 1000  # Adjust as needed
            if area_name in self.text_histories and len(self.text_histories[area_name]) > max_history_size:
                # Keep only the most recent entries
                self.text_histories[area_name] = set(list(self.text_histories[area_name])[-max_history_size:])

        # Split text into lines to handle usernames
        lines = text.split('\n')
        filtered_lines = []
        
        import re  # Move import to here, outside the loop
        for line in lines:
            if not line.strip():  # Skip empty lines
                continue
            words = line.split()
            if words:
                # Filter out usernames if enabled
                if self.ignore_usernames_var.get():
                    # Check for username pattern (word followed by : or ;)
                    filtered_words = []
                    i = 0
                    while i < len(words):
                        # Check if current word is part of a username pattern
                        if i < len(words) - 1 and words[i + 1] in [':', ';']:
                            i += 2 if words[i + 1] in [':', ';'] else 1
                        else:
                            filtered_words.append(words[i])
                        i += 1
                    line = ' '.join(filtered_words)

                ignore_items = [item.strip().lower() for item in self.bad_word_list.get().split(',') if item.strip()]

                def normalize_text(text):
                    # Remove punctuation, normalize spaces, and make lowercase
                    text = re.sub(r'[\W_]+', ' ', text)  # Replace all non-word chars with space
                    text = re.sub(r'\s+', ' ', text)     # Collapse multiple spaces
                    return text.strip().lower()

                def is_ignored(text, ignore_items, normalized_line=None):
                    """
                    Returns True if the text matches any ignored word or phrase (case-insensitive, ignores punctuation and extra spaces).
                    - For single words: matches if the word matches (case-insensitive)
                    - For phrases: matches if the phrase appears exactly (case-insensitive, ignores punctuation)
                    """
                    text_norm = normalize_text(text)
                    for item in ignore_items:
                        if ' ' in item:
                            # Phrase: match as exact phrase anywhere in the normalized line
                            if normalized_line and normalize_text(item) in normalized_line:
                                return True
                        else:
                            # Single word: match as a word (not substring)
                            if text_norm == normalize_text(item):
                                return True
                    return False

                # Normalize the line for phrase matching
                normalized_line = normalize_text(line)

                # Remove ignored phrases (with spaces)
                for item in ignore_items:
                    if ' ' in item:
                        norm_phrase = normalize_text(item)
                        # Remove all occurrences of the phrase from the normalized line
                        while norm_phrase in normalized_line:
                            # Find the phrase in the original line (approximate position)
                            # Replace in the original line (case-insensitive, ignoring punctuation)
                            # We'll use regex for robust matching
                            pattern = re.compile(r'\b' + re.escape(item) + r'\b', re.IGNORECASE)
                            line = pattern.sub(' ', line)
                            # Re-normalize after replacement
                            normalized_line = normalize_text(line)

                # Now split and filter out single ignored words
                filtered_words = [word for word in line.split() if not any(normalize_text(word) == normalize_text(item) for item in ignore_items if ' ' not in item)]

                # Only skip if the line is empty after filtering
                if not filtered_words:
                    continue
                
                # Apply gibberish filtering if enabled
                if self.ignore_gibberish_var.get():
                    vowels = set('aeiouAEIOU')
                    def is_not_gibberish(word):
                        if any(c.isalpha() for c in word) or any(c.isdigit() for c in word):
                            if len(word) <= 3:
                                return True
                            # Allow if word contains a vowel (for longer words)
                            return any(c in vowels for c in word) or any(c.isdigit() for c in word)
                        return False
                    filtered_words = [word for word in filtered_words if is_not_gibberish(word)]
                
                if filtered_words:  # Only add non-empty lines
                    filtered_lines.append(' '.join(filtered_words))
        # Join lines with proper spacing
        filtered_text = ' '.join(filtered_lines)

        if self.pause_at_punctuation_var.get():
            # Replace punctuation with itself plus a pause marker
            for punct in ['.', '!', '?']:
                filtered_text = filtered_text.replace(punct, punct + ' ... ')
            # Add smaller pauses for commas and semicolons
            for punct in [',', ';']:
                filtered_text = filtered_text.replace(punct, punct + ' .. ')

        # Set the voice and speed for SAPI
        if voice_var:
            # Check both SAPI voices and our combined voice list (includes mock voices)
            selected_voice = None
            
            # Get the actual voice name (full name, not display name)
            actual_voice_name = getattr(voice_var, '_full_name', voice_var.get())
            
            # First try to find in SAPI voices
            try:
                voices = self.speaker.GetVoices()
                for voice in voices:
                    if voice.GetDescription() == actual_voice_name:
                        selected_voice = voice
                        break
            except Exception as e:
                print(f"Error getting SAPI voices: {e}")
            
            # If not found in SAPI, try our combined voice list (includes mock voices)
            if not selected_voice and hasattr(self, 'voices'):
                for voice in self.voices:
                    if hasattr(voice, 'GetDescription') and voice.GetDescription() == actual_voice_name:
                        # Check if this is a real SAPI voice object
                        if hasattr(voice, 'GetId') and hasattr(voice, 'GetToken'):  # Working OneCore voice object
                            print(f"Found working OneCore voice: {actual_voice_name}")
                            selected_voice = voice
                            break
                        elif hasattr(voice, 'GetId'):  # Real SAPI voice object
                            print(f"Found real voice in combined list: {actual_voice_name}")
                            selected_voice = voice
                            break
                        else:
                            # For mock voices, we can't set them directly, so just continue
                            print(f"Found mock voice: {actual_voice_name}")
                            selected_voice = "mock_voice"  # Mark as found but don't set
                            break
            
            if selected_voice and selected_voice != "mock_voice":
                try:
                    # If this is a OneCore voice, route through UWP immediately for reliability
                    if hasattr(selected_voice, 'GetToken'):
                        print(f"Using OneCore voice via Narrator: {actual_voice_name}")
                        if _ensure_uwp_available():
                            try:
                                loop = asyncio.new_event_loop()
                                asyncio.set_event_loop(loop)
                                loop.run_until_complete(self._speak_with_uwp(filtered_text, preferred_desc=actual_voice_name))
                                loop.close()
                                return
                            except Exception as _e:
                                print(f"UWP fallback failed: {_e}")
                                import traceback; traceback.print_exc()
                        else:
                            print("UWP TTS not available. Install with: pip install winsdk")
                        # If UWP not available or failed, we fall through and try SAPI default
                    else:
                        # Regular SAPI voice
                        self.speaker.Voice = selected_voice
                        print(f"Successfully set voice to: {selected_voice.GetDescription()}")
                except Exception as set_voice_e:
                    print(f"Error setting voice: {set_voice_e}")
                    messagebox.showerror("Error", f"Could not set voice: {set_voice_e}")
                    return
            elif selected_voice == "mock_voice":
                # For mock voices (OneCore), use UWP path if available
                print(f"Using OneCore (mock) voice via Narrator: {actual_voice_name}")
                if _ensure_uwp_available():
                    try:
                        loop = asyncio.new_event_loop()
                        asyncio.set_event_loop(loop)
                        loop.run_until_complete(self._speak_with_uwp(filtered_text, preferred_desc=actual_voice_name))
                        loop.close()
                        return
                    except Exception as _e:
                        print(f"UWP fallback failed: {_e}")
                        import traceback; traceback.print_exc()
                else:
                    print("UWP TTS not available. Install with: pip install winsdk")
                # If UWP not available, inform and abort
                messagebox.showerror("Error", "Selected voice requires Windows Narrator TTS. Please install 'winsdk' (pip install winsdk) or choose a SAPI voice.")
                return
            else:
                messagebox.showerror("Error", "No voice selected. Please select a voice.")
                print("Error: Did not speak, Reason: No selected voice.")
                return

        # Update speed for win32com - Convert from percentage to rate (-10 to 10)
        if speed_var:
            try:
                speed = int(speed_var.get())
                if speed > 0:
                    # Convert speed percentage to SAPI rate (-10 to 10)
                    self.speaker.Rate = (speed - 100) // 10
            except ValueError:
                pass  # Invalid speed value, ignore

        # Set volume and speak text
        try:
            # Set volume
            try:
                vol = int(self.volume.get())
                if 0 <= vol <= 100:
                    self.speaker.Volume = vol
                else:
                    self.volume.set("100")
                    self.speaker.Volume = 100
            except ValueError:
                self.volume.set("100")
                self.speaker.Volume = 100

            # Ensure speech engine is ready before speaking
            self._ensure_speech_ready()
            
            # Speak the text
            self.is_speaking = True
            self.speaker.Speak(filtered_text, 1)  # 1 is SVSFlagsAsync
            print("Speech started.\n--------------------------")
        except Exception as e:
            print(f"Error during speech: {e}")
            self.is_speaking = False
            try:
                # Try to reinitialize the speaker
                self.speaker = win32com.client.Dispatch("SAPI.SpVoice")
                self.speaker.Volume = int(self.volume.get())
            except Exception as e2:
                print(f"Error reinitializing speaker: {e2}")
                self.is_speaking = False

    def cleanup(self):
        """Proper cleanup method for the application"""
        print("Performing cleanup...")
        try:
            # Stop UWP worker
            try:
                if hasattr(self, '_uwp_thread_stop'):
                    self._uwp_thread_stop.set()
                if hasattr(self, '_uwp_interrupt'):
                    try:
                        self._uwp_interrupt.set()
                    except Exception:
                        pass
                if hasattr(self, '_uwp_queue'):
                    try:
                        self._uwp_queue.put_nowait(("STOP", None, None))
                    except Exception:
                        pass
                if hasattr(self, '_uwp_thread') and self._uwp_thread:
                    self._uwp_thread.join(timeout=0.5)
            except Exception:
                pass
            # First, clean up the debug console if it exists
            if hasattr(self, 'console_window'):
                try:
                    self.console_window.window.destroy()
                    delattr(self, 'console_window')
                except:
                    pass

            # Restore stdout to its original state
            if hasattr(sys, 'stdout_original'):
                sys.stdout = sys.stdout_original

            # Cleanup hotkeys
            try:
                self.disable_all_hotkeys()
            except Exception as e:
                print(f"Warning: Error cleaning up hotkeys: {e}")
            
            # Cleanup Tesseract engine
            if hasattr(self, 'engine'):
                self.engine.endLoop()
                del self.engine
            
            # Cleanup speaker
            if hasattr(self, 'speaker'):
                del self.speaker
            
            # Cleanup hotkey scancodes and settings
            self.hotkey_scancodes.clear()
            self.processing_settings.clear()
            self.text_histories.clear()
            
            # Clear all variables
            self.hotkeys.clear()
            self.areas.clear()
            self.latest_images.clear()
            
            # Reset flags
            self.is_speaking = False
            self.setting_hotkey = False
            
        except Exception as e:
            print(f"Error during cleanup: {e}")
        finally:
            print("Cleanup completed")

    def __del__(self):
        """Cleanup when the object is destroyed."""
        self.cleanup()

    def is_valid_text(self, text):
        """Check if text appears to be valid (not gibberish)."""
        # Skip empty text
        if not text.strip():  # Skip empty lines
            return False
            
        # Count valid vs invalid characters
        valid_chars = 0
        invalid_chars = 0
        
        for char in text:
            # Count letters, numbers, and common punctuation as valid
            if char.isalnum() or char in ".,!?'\"- ":
                valid_chars += 1
            else:
                invalid_chars += 1
        
        # If there are too many invalid characters relative to valid ones, consider it gibberish
        if invalid_chars > valid_chars / 2:
            return False
            
        # Check for repeated symbols which often appear in OCR artifacts
        if any(symbol * 2 in text for symbol in "/\\|[]{}=<>+*"):
            return False
            
        # Check minimum length after stripping special charactersa
        clean_text = ''.join(c for c in text if c.isalnum() or c.isspace())
        if len(clean_text.strip()) < 2:  # Require at least 2 alphanumeric characters
            return False
            
        return True

    def show_processing_feedback(self, area_name):
        """Show processing feedback with text only"""
        # Cancel any existing feedback clear timer
        if hasattr(self, '_feedback_timer') and self._feedback_timer:
            self.root.after_cancel(self._feedback_timer)
        
        # Update status text
        self.status_label.config(text=f"Processing Area: {area_name}")
        
        # Set timer to clear the text after 0.5 seconds
        self._feedback_timer = self.root.after(1300, lambda: self.status_label.config(text=""))


# Add this function near the top of the file, after the imports
def open_url(url):
    """Helper function to open URLs in the default browser"""
    try:
        print(f"Attempting to open URL: {url}")
        result = webbrowser.open(url)
        if result:
            print(f"Successfully opened URL: {url}")
        else:
            print(f"Failed to open URL: {url} - webbrowser.open returned False")
    except Exception as e:
        print(f"Error opening URL {url}: {e}")
        # Try alternative method
        try:
            import subprocess
            import platform
            if platform.system() == "Windows":
                subprocess.run(["start", url], shell=True, check=True)
                print(f"Opened URL using Windows start command: {url}")
            elif platform.system() == "Darwin":  # macOS
                subprocess.run(["open", url], check=True)
                print(f"Opened URL using macOS open command: {url}")
            else:  # Linux
                subprocess.run(["xdg-open", url], check=True)
                print(f"Opened URL using xdg-open: {url}")
        except Exception as e2:
            print(f"Alternative method also failed: {e2}")

def capture_screen_area(x1, y1, x2, y2):
    """Capture screen area across multiple monitors using win32api"""
    # Get virtual screen bounds
    min_x = win32api.GetSystemMetrics(win32con.SM_XVIRTUALSCREEN)  # Leftmost x (can be negative)
    min_y = win32api.GetSystemMetrics(win32con.SM_YVIRTUALSCREEN)  # Topmost y (can be negative)
    total_width = win32api.GetSystemMetrics(win32con.SM_CXVIRTUALSCREEN)
    total_height = win32api.GetSystemMetrics(win32con.SM_CYVIRTUALSCREEN)
    max_x = min_x + total_width
    max_y = min_y + total_height
    
    print(f"Debug: Screenshot capture - Input coords: ({x1}, {y1}, {x2}, {y2})")
    print(f"Debug: Virtual screen bounds: ({min_x}, {min_y}, {max_x}, {max_y})")

    # Clamp coordinates to virtual screen bounds
    x1 = max(min_x, min(max_x, x1))
    y1 = max(min_y, min(max_y, y1))
    x2 = max(min_x, min(max_x, x2))
    y2 = max(min_y, min(max_y, y2))
    
    print(f"Debug: Clamped coordinates: ({x1}, {y1}, {x2}, {y2})")

    # Ensure valid area (swap if necessary and check size)
    x1, x2 = min(x1, x2), max(x1, x2)
    y1, y2 = min(y1, y2), max(y1, y2)
    width = x2 - x1
    height = y2 - y1
    if width <= 0 or height <= 0:
        return Image.new('RGB', (1, 1))  # Return a blank 1x1 image for invalid areas

    # Get DC from entire virtual screen
    hwin = win32gui.GetDesktopWindow()
    hwindc = win32gui.GetWindowDC(hwin)
    srcdc = win32ui.CreateDCFromHandle(hwindc)
    memdc = srcdc.CreateCompatibleDC()

    # Create bitmap for capture area
    bmp = win32ui.CreateBitmap()
    bmp.CreateCompatibleBitmap(srcdc, width, height)
    memdc.SelectObject(bmp)

    # Copy screen into bitmap
    memdc.BitBlt((0, 0), (width, height), srcdc, (x1, y1), win32con.SRCCOPY)

    # Convert bitmap to PIL Image
    bmpinfo = bmp.GetInfo()
    bmpstr = bmp.GetBitmapBits(True)
    img = Image.frombuffer(
        'RGB',
        (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
        bmpstr, 'raw', 'BGRX', 0, 1
    )

    # Clean up
    memdc.DeleteDC()
    win32gui.ReleaseDC(hwin, hwindc)
    win32gui.DeleteObject(bmp.GetHandle())

    return img

if __name__ == "__main__":
    import tempfile, os, json
    
    # Use TkinterDnD's Tk if available, otherwise fall back to regular tkinter
    if TKDND_AVAILABLE:
        try:
            root = TkinterDnD.Tk()
        except Exception as _tkdnd_error:
            print(f"Warning: TkinterDnD failed to initialize ({_tkdnd_error}). Falling back to Tk.")
            root = tk.Tk()
    else:
        root = tk.Tk()
    
    # Hide the window during setup to prevent the "stretching" effect
    root.withdraw()
    
    # Set the window icon
    try:
        icon_path = os.path.join(os.path.dirname(__file__), 'Assets', 'icon.ico')
        if os.path.exists(icon_path):
            root.iconbitmap(icon_path)
            print(f"Set window icon to: {icon_path}")
        else:
            print(f"Icon file not found at: {icon_path}")
    except Exception as e:
        print(f"Error setting window icon: {e}")
    
    app = GameTextReader(root)
    # Create permanent area at the top
    app.add_read_area(removable=False, editable_name=False, area_name="Auto Read")
    
    # Set the proper window size before it becomes visible
    app.root.update_idletasks()  # Ensure all widgets are properly sized
    app.resize_window(force=True)  # Calculate and set the optimal window size
    
    # Now show the window at the correct size
    app.root.deiconify()
    # Try to load settings for Auto Read area from temp folder
    temp_path = os.path.join(tempfile.gettempdir(), 'GameReader', 'auto_read_settings.json')
    if os.path.exists(temp_path) and app.areas:
        try:
            # Basic file validation
            file_size = os.path.getsize(temp_path)
            if file_size > 1024 * 1024:  # 1MB limit for auto-read settings
                print("Warning: Auto-read settings file is too large, skipping load")
            else:
                with open(temp_path, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                
                # Validate auto-read settings
                if not isinstance(settings, dict):
                    print("Warning: Invalid auto-read settings format, skipping load")
                else:
                    # Validate expected fields
                    valid_settings = True
                    for key, value in settings.items():
                        if key == 'preprocess' and not isinstance(value, bool):
                            valid_settings = False
                            break
                        elif key == 'voice' and not isinstance(value, str):
                            valid_settings = False
                            break
                        elif key == 'speed' and not isinstance(value, str):
                            valid_settings = False
                            break
                        elif key == 'hotkey' and value is not None and (not isinstance(value, str) or len(value) > 50):
                            valid_settings = False
                            break
                        elif key == 'stop_read_on_select' and not isinstance(value, bool):
                            valid_settings = False
                            break
                        elif key == 'processing' and not isinstance(value, dict):
                            valid_settings = False
                            break
                    
                    if not valid_settings:
                        print("Warning: Invalid auto-read settings values, skipping load")
                    else:
                        # Find the permanent area and set its settings
                        area_frame, hotkey_button, set_area_button, area_name_var, preprocess_var, voice_var, speed_var = app.areas[0]
                        
                        # Get preprocess state from root level
                        preprocess_state = settings.get('preprocess', False)
                        
                        # Set UI controls from root level settings
                        preprocess_var.set(preprocess_state)
            
                        # Convert full voice name back to display name for the UI
                        saved_voice = settings.get('voice', 'Select Voice')
                        if saved_voice != 'Select Voice':
                            # Find the display name that corresponds to this full voice name
                            display_name = 'Select Voice'
                            for i, voice in enumerate(app.voices, 1):
                                if hasattr(voice, 'GetDescription') and voice.GetDescription() == saved_voice:
                                    # Create the same abbreviated display name logic WITH numbering
                                    full_name = voice.GetDescription()
                                    if "Microsoft" in full_name and " - " in full_name:
                                        parts = full_name.split(" - ")
                                        if len(parts) == 2:
                                            voice_part = parts[0].replace("Microsoft ", "")
                                            lang_part = parts[1]
                                            display_name = f"{i}. {voice_part} ({lang_part})"
                                        else:
                                            display_name = f"{i}. {full_name}"
                                    elif " - " in full_name:
                                        parts = full_name.split(" - ")
                                        if len(parts) == 2:
                                            display_name = f"{i}. {parts[0]} ({parts[1]})"
                                        else:
                                            display_name = f"{i}. {full_name}"
                                    else:
                                        display_name = f"{i}. {full_name}"
                                    break
                            voice_var.set(display_name)
                            # Set the full name for the voice variable
                            voice_var._full_name = saved_voice
                        else:
                            voice_var.set('Select Voice')
            
            speed_var.set(settings.get('speed', '100'))
            
            # Set hotkey if it exists
            if settings.get('hotkey'):
                hotkey_button.hotkey = settings['hotkey']
                display_name = settings['hotkey'].replace('num_', 'num:').replace('multiply', '*').replace('add', '+').replace('subtract', '-').replace('divide', '/') if settings['hotkey'].startswith('num_') else settings['hotkey'].replace('multiply', '*').replace('add', '+').replace('subtract', '-').replace('divide', '/')
                hotkey_button.config(text=f"Set Hotkey: [ {display_name.upper()} ]")
                app.setup_hotkey(hotkey_button, area_frame)
            
            # Get processing settings from the 'processing' dictionary
            processing_settings = settings.get('processing', {})
            brightness = processing_settings.get('brightness', 1.0)
            contrast = processing_settings.get('contrast', 1.0)
            saturation = processing_settings.get('saturation', 1.0)
            sharpness = processing_settings.get('sharpness', 1.0)
            blur = processing_settings.get('blur', 0.0)
            hue = processing_settings.get('hue', 0.0)
            exposure = processing_settings.get('exposure', 1.0)
            threshold = processing_settings.get('threshold', 128)
            threshold_enabled = processing_settings.get('threshold_enabled', False)
            
            # Store all settings in processing_settings
            app.processing_settings['Auto Read'] = {
                'brightness': brightness,
                'contrast': contrast,
                'saturation': saturation,
                'sharpness': sharpness,
                'blur': blur,
                'hue': hue,
                'exposure': exposure,
                'threshold': threshold,
                'threshold_enabled': threshold_enabled,
                'preprocess': preprocess_state  # Include preprocess flag from root level
            }
            
            # Set interrupt on new scan setting
            app.interrupt_on_new_scan_var.set(settings.get('stop_read_on_select', False))
            
            # If the UI for these exists, set them as well
            if hasattr(app, 'processing_settings_widgets'):
                widgets = app.processing_settings_widgets.get('Auto Read', {})
                if 'brightness' in widgets:
                    widgets['brightness'].set(brightness)
                if 'contrast' in widgets:
                    widgets['contrast'].set(contrast)
                if 'saturation' in widgets:
                    widgets['saturation'].set(saturation)
                if 'sharpness' in widgets:
                    widgets['sharpness'].set(sharpness)
                if 'blur' in widgets:
                    widgets['blur'].set(blur)
                if 'hue' in widgets:
                    widgets['hue'].set(hue)
                if 'exposure' in widgets:
                    widgets['exposure'].set(exposure)
                if 'threshold' in widgets:
                    widgets['threshold'].set(threshold)
                if 'threshold_enabled' in widgets:
                    widgets['threshold_enabled'].set(threshold_enabled)
                print("Loaded Auto Read settings successfully")
        except Exception as e:
            print(f"Error loading Auto Read settings: {e}")
            # If there was an error, initialize with default settings
            if 'Auto Read' not in app.processing_settings:
                app.processing_settings['Auto Read'] = {
                    'brightness': 1.0,
                    'contrast': 1.0,
                    'saturation': 1.0,
                    'sharpness': 1.0,
                    'blur': 0.0,
                    'hue': 0.0,
                    'exposure': 1.0,
                    'threshold': 128,
                    'threshold_enabled': False,
                }

    root.mainloop()
