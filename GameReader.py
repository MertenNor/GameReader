###
###  I know.. This code is.. well not great, all made with AI, but it works. feel free to make any changes!
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
from PIL import Image, ImageEnhance, ImageFilter, ImageGrab, ImageTk
import ctypes
import winsound
import asyncio
import queue
try:
    # winsdk is the package name; modules generally import from winrt.*
    import importlib
    UWP_TTS_AVAILABLE = False
    _uwp_import_error = None
    try:
        from winrt.windows.media.speechsynthesis import SpeechSynthesizer
        from winrt.windows.storage.streams import DataReader
        UWP_TTS_AVAILABLE = True
    except Exception as e:
        _uwp_import_error = e
        try:
            # Attempt to import winsdk meta package and retry
            importlib.import_module('winsdk')
            from winrt.windows.media.speechsynthesis import SpeechSynthesizer
            from winrt.windows.storage.streams import DataReader
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
            from winrt.windows.media.speechsynthesis import SpeechSynthesizer as _SS  # noqa: F401
            from winrt.windows.storage.streams import DataReader as _DR  # noqa: F401
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

APP_VERSION = "0.8.5"

CHANGELOG = """
- Fixed an issue where numpad hotkeys didn’t work in fullscreen mode for some users.
- Properly added support for voices installed via Windows settings.
- Updated the Info/Help window.
- Made minor UI improvements.


Thanks to everyone who has sent in feedback and bug reports!

Note: Some users have requested to add AI Voices this will the next main priority.

- Thanks for using GameReader!
 
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

class ImageProcessingWindow:
    def __init__(self, root, area_name, latest_images, settings, game_text_reader):
        self.window = tk.Toplevel(root)
        self.window.title(f"Image Processing for: {area_name}")
        self.area_name = area_name
        self.latest_images = latest_images
        self.settings = settings
        self.game_text_reader = game_text_reader

        # Check if there is an image for the area
        if area_name not in latest_images:
            messagebox.showerror("Error", "No image to process, generate an image by pressing the hotkey.")
            self.window.destroy()
            return

        self.image = latest_images[area_name]
        self.processed_image = self.image.copy()

        # Create a canvas to display the image
        self.image_frame = ttk.Frame(self.window)
        self.image_frame.grid(row=0, column=0, columnspan=5, padx=10, pady=10)
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
        control_frame.grid(row=1, column=0, columnspan=5, pady=10)

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

        self.create_slider("Brightness", self.brightness_var, 0.1, 2.0, 1.0, 2, 0)
        self.create_slider("Contrast", self.contrast_var, 0.1, 2.0, 1.0, 2, 1)
        self.create_slider("Saturation", self.saturation_var, 0.1, 2.0, 1.0, 2, 2)
        self.create_slider("Sharpness", self.sharpness_var, 0.1, 2.0, 1.0, 2, 3)
        self.create_slider("Blur", self.blur_var, 0.0, 10.0, 0.0, 2, 4)
        self.create_slider("Threshold", self.threshold_var, 0, 255, 128, 3, 0, self.threshold_enabled_var)
        self.create_slider("Hue", self.hue_var, -1.0, 1.0, 0.0, 3, 1)
        self.create_slider("Exposure", self.exposure_var, 0.1, 2.0, 1.0, 3, 2)

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
    match = re.search(r'CHANGELOG\s*=\s*([ru]?)(["\\']{3})(.*?)\2', code, re.DOTALL)
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

class GameTextReader:
    def __init__(self, root):
        self.root = root
        self.root.title(f"Game Reader v{APP_VERSION}")
        # --- Update check on startup ---
        local_version = APP_VERSION
        FORCE_UPDATE_CHECK = False  # Set to True to force update popup, False for normal behavior
        threading.Thread(target=lambda: check_for_update(local_version, force=FORCE_UPDATE_CHECK), daemon=True).start()
        # --- End update check ---
        self.root.geometry("1115x260")  # Initial window size (height reduced for less vertical tallness)
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
            55: '*',     # Numpad *
            78: '+',     # Numpad +
            74: '-',     # Numpad -
            83: '.',     # Numpad .
            53: '/',     # Numpad /
            28: 'enter'  # Numpad Enter
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
            # Method 4: Try to force Windows to load OneCore voices directly (quiet)
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
        
        # Enable drag and drop using TkinterDnD2
        root.drop_target_register(DND_FILES)
        root.dnd_bind('<<Drop>>', self.on_drop)
        root.dnd_bind('<<DropEnter>>', lambda e: 'break')
        root.dnd_bind('<<DropPosition>>', lambda e: 'break')

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

    async def _speak_with_uwp(self, text: str, preferred_desc: str = None):
        """Speak using UWP Narrator (OneCore) via Windows.Media.SpeechSynthesis.
        This plays audio directly and does not integrate with SAPI voices. Used as a fallback
        to get OneCore/Narrator voices speaking when SAPI can't set them.
        """
        if not UWP_TTS_AVAILABLE:
            return
        # Import lazily to avoid hard dependency at import time
        try:
            from winrt.windows.media.speechsynthesis import SpeechSynthesizer  # type: ignore
        except Exception:
            try:
                from winsdk.windows.media.speechsynthesis import SpeechSynthesizer  # type: ignore
            except Exception:
                return
        try:
            from winrt.windows.media.playback import MediaPlayer  # type: ignore
            from winrt.windows.media.core import MediaSource  # type: ignore
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
                from winrt.windows.media.playback import MediaPlayer
                from winrt.windows.media.core import MediaSource
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
                                from winrt.windows.media.playback import MediaPlaybackState  # type: ignore
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
            
    def restart_tesseract(self):
        """Forcefully stop the speech and reinitialize the system."""
        print("Forcing stop...")
        try:
            self.stop_speaking()  # Stop the speech
            print("System reinitialized. Audio stopped.")
        except Exception as e:
            print(f"Error during forced stop: {e}")

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
        
        debug_button = tk.Button(buttons_frame, text="Debug Window", command=self.show_debug)
        debug_button.pack(side='left', padx=5)
        
        info_button = tk.Button(buttons_frame, text="Info/Help", command=self.show_info)
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
        
        save_button = tk.Button(layout_buttons_frame, text="Save Layout", command=self.save_layout)
        save_button.pack(side='left', padx=5)
        
        load_button = tk.Button(layout_buttons_frame, text="Load Layout..", command=self.load_layout)
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
        add_area_button = tk.Button(add_area_frame, text="Add Read Area", 
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

        # --- Disable all hotkeys while info window is open ---
        try:
            keyboard.unhook_all()
            mouse.unhook_all()
        except Exception as e:
            print(f"Error unhooking hotkeys for info window: {e}")

        # On close, re-enable all hotkeys
        def on_info_close():
            # Restore hotkeys for all areas
            for area in self.areas:
                area_frame, hotkey_button, _, area_name_var, _, _, _ = area
                if hasattr(hotkey_button, 'hotkey'):
                    self.setup_hotkey(hotkey_button, area_frame)
            # Restore stop hotkey if present
            if hasattr(self, 'stop_hotkey_button') and hasattr(self.stop_hotkey_button, 'mock_button'):
                self.setup_hotkey(self.stop_hotkey_button.mock_button, None)
            info_window.destroy()

        info_window.protocol("WM_DELETE_WINDOW", on_info_close)
        info_window.bind('<Escape>', lambda e: on_info_close())
        
        # Set window icon if available
        try:
            info_window.iconbitmap('icon.ico')  # You would need to add an icon file
        except:
            pass
        
        # Main container with padding
        main_frame = ttk.Frame(info_window, padding="20 20 20 10")
        main_frame.pack(fill='both', expand=True)
        
        # Title section
        title_frame = ttk.Frame(main_frame)
        title_frame.pack(fill='x', pady=(0, 20))
        
        title_label = ttk.Label(title_frame, 
                               text=f"GameReader v{APP_VERSION}", 
                               font=("Helvetica", 16, "bold"))
        title_label.pack(side='left')
        
        # Credits/Links Area replaced by clickable images
        credits_frame = ttk.Frame(main_frame)
        credits_frame.pack(fill='x', pady=(0, 20))

        images_row = ttk.Frame(credits_frame)
        images_row.pack(fill='x', pady=(0, 5))

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

            # Smooth shrink animation on mouse leave
            def _cancel_anim(c):
                if hasattr(c, "_anim_job") and c._anim_job:
                    try:
                        c.after_cancel(c._anim_job)
                    except Exception:
                        pass
                    c._anim_job = None

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
                cursor='hand2'
            )
            coffee_canvas.pack(side='left', padx=10)
            coffee_img_id = coffee_canvas.create_image(coffee_cw // 2, coffee_ch // 2, image=info_window.coffee_photo)
            coffee_canvas.bind("<Button-1>", lambda e: open_url("https://buymeacoffee.com/mertennor"))
            coffee_canvas.bind(
                "<Enter>",
                lambda e, c=coffee_canvas, iid=coffee_img_id: c.itemconfig(iid, image=info_window.coffee_photo_hover)
            )
            coffee_canvas.bind(
                "<Leave>",
                lambda e, c=coffee_canvas, iid=coffee_img_id: animate_to_normal(c, iid, info_window.coffee_pil)
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
                cursor='hand2'
            )
            google_canvas.pack(side='left', padx=10)
            google_img_id = google_canvas.create_image(google_cw // 2, google_ch // 2, image=info_window.google_photo)
            google_canvas.bind("<Button-1>", lambda e: open_url("https://forms.gle/8YBU8atkgwjyzdM79"))
            google_canvas.bind(
                "<Enter>",
                lambda e, c=google_canvas, iid=google_img_id: c.itemconfig(iid, image=info_window.google_photo_hover)
            )
            google_canvas.bind(
                "<Leave>",
                lambda e, c=google_canvas, iid=google_img_id: animate_to_normal(c, iid, info_window.google_pil)
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
                cursor='hand2'
            )
            github_canvas.pack(side='left', padx=10)
            github_img_id = github_canvas.create_image(github_cw // 2, github_ch // 2, image=info_window.github_photo)
            github_canvas.bind("<Button-1>", lambda e: open_url("https://github.com/MertenNor/GameReader"))
            github_canvas.bind(
                "<Enter>",
                lambda e, c=github_canvas, iid=github_img_id: c.itemconfig(iid, image=info_window.github_photo_hover)
            )
            github_canvas.bind(
                "<Leave>",
                lambda e, c=github_canvas, iid=github_img_id: animate_to_normal(c, iid, info_window.github_pil)
            )
        except Exception as e:
            # Fallback text if images can't be displayed
            fallback = ttk.Label(credits_frame, text=f"Error displaying info images: {e}", foreground='red')
            fallback.pack(anchor='w')


        # Spacer before Tesseract warning
        ttk.Label(credits_frame, text="").pack()
        tesseract_frame = ttk.Frame(credits_frame)
        tesseract_frame.pack(fill='x', pady=(10, 0))

        # First line of the message
        tesseract_label = ttk.Label(
            tesseract_frame,
            text=(
                "! IMPORTANT ! This program requires Tesseract OCR to function ( default installation: C:\Program Files ) to process text in images."
            ),
            font=("Helvetica", 10, "bold"),
            foreground='red'
        )
        tesseract_label.pack(anchor='w', pady=(0, 5))

        # Second line with a clickable link
        download_label = ttk.Label(
            tesseract_frame,
            text="Download the latest version here:",
            font=("Helvetica", 10, "bold"),
            foreground='red'
        )
        # Download instruction and clickable URLs on the same line
        download_row = ttk.Frame(tesseract_frame)
        download_row.pack(anchor='w')
        download_label = ttk.Label(download_row,
                                   text="Download the latest version here: ",
                                   font=("Helvetica", 10, "bold"),
                                   foreground='red')
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
                               text="Direct link to Tesseract installer v.5.5.0",
                               font=("Helvetica", 10),
                               foreground='blue',
                               cursor='hand2')
        direct_link.pack(side='left')
        direct_link.bind("<Button-1>", lambda e: open_url("https://github.com/tesseract-ocr/tesseract/releases/download/5.5.0/tesseract-ocr-w64-setup-5.5.0.20241111.exe"))
        direct_link.bind("<Enter>", lambda e: direct_link.configure(font=("Helvetica", 10, "underline")))
        direct_link.bind("<Leave>", lambda e: direct_link.configure(font=("Helvetica", 10)))
        # Explanatory label below
        tesseract_note = ttk.Label(tesseract_frame,
                                   text="( yes, you need this if you want this program to read the text for you. )",
                                   font=("Helvetica", 11, "bold"),
                                   foreground='red')
        tesseract_note.pack(anchor='w', padx=(0,0), pady=(2, 6))

        
        # Create a frame with scrollbar for the main content
        content_frame = ttk.Frame(main_frame)
        content_frame.pack(fill='both', expand=True)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(content_frame)
        scrollbar.pack(side='right', fill='y')
        
        # Create text widget with custom styling - make it selectable
        text_widget = tk.Text(content_frame, 
                             wrap=tk.WORD, 
                             yscrollcommand=scrollbar.set,
                             font=("Helvetica", 10),
                             padx=10,
                             pady=10,
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
                                   .replace('left ctrl','L-CTRL').replace('right ctrl','R-CTRL') \
                                   .replace('left alt','L-ALT').replace('right alt','R-ALT') \
                                   .replace('left shift','L-SHIFT').replace('right shift','R-SHIFT') \
                                   .replace('windows','WIN')
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
            # Normalize name and side
            raw_name = (event.name or '').lower()
            side = None
            if any(s in raw_name for s in ["right", "høyre", "rechts", "derecha", "droite", "destra", "dereito", "höger"]):
                side = 'right'
            if any(s in raw_name for s in ["left", "venstre", "links", "izquierda", "gauche", "sinistra", "esquerda", "vänster"]):
                side = 'left'
            base = None
            if any(x in raw_name for x in ["ctrl", "control", "strg"]):
                base = 'ctrl'
            elif any(x in raw_name for x in ["altgr", "option", "menu"]) or raw_name.strip() == 'alt':
                base = 'alt'
            elif any(x in raw_name for x in ["shift", "skift", "umschalt"]):
                base = 'shift'
            elif any(x in raw_name for x in ["windows", "win", "super", "meta", "cmd", "command", "lwin", "rwin"]):
                base = 'windows'
            if base:
                if base in ('ctrl','alt'):
                    name = f"{side or 'left'} {base}"
                elif base in ('shift','windows'):
                    name = base
                else:
                    name = base
            else:
                name = raw_name

            # Non-modifier pressed
            if name not in ('ctrl','left ctrl','right ctrl','alt','left alt','right alt','shift','windows'):
                combo_state['non_modifier_pressed'] = True
            # Bare modifier assignment path
            if name in ('ctrl','shift','alt','windows'):
                def _assign_bare_modifier():
                    if self._hotkey_assignment_cancelled or not self.setting_hotkey:
                        return
                    try:
                        held = []
                        if keyboard.is_pressed('left ctrl'): held.append('left ctrl')
                        if keyboard.is_pressed('right ctrl'): held.append('right ctrl')
                        if keyboard.is_pressed('shift'): held.append('shift')
                        if keyboard.is_pressed('left alt'): held.append('left alt')
                        if keyboard.is_pressed('right alt'): held.append('right alt')
                        if keyboard.is_pressed('left windows') or keyboard.is_pressed('right windows') or keyboard.is_pressed('windows'):
                            held.append('windows')
                        if len(held) == 1:
                            only = held[0]
                            if (base == 'ctrl' and (only in ['left ctrl','right ctrl'])) or \
                               (base == 'alt' and (only in ['left alt','right alt'])) or \
                               (base == 'shift' and only == 'shift') or \
                               (base == 'windows' and only == 'windows'):
                                key_name_local = only
                                _assign_stop_hotkey_and_register(key_name_local)
                                return
                    except Exception:
                        pass
                try:
                    self.root.after(200, _assign_bare_modifier)
                except Exception:
                    pass
                return

            # Build combo from held modifiers + base key
            try:
                mods = []
                if keyboard.is_pressed('left ctrl'): mods.append('left ctrl')
                if keyboard.is_pressed('right ctrl'): mods.append('right ctrl')
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
                base_key = f"numpad {sym}"
            if base_key in ("ctrl", "shift", "alt", "windows", "left ctrl", "right ctrl", "left alt", "right alt"):
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

        # Set button to indicate we're waiting for input
        self.stop_hotkey_button.config(text="Press any key or mouse button...")
        
        # Set up temporary hooks for key and mouse input
        try:
            # Store the hooks as attributes of the button for cleanup
            self.stop_hotkey_button.keyboard_hook_temp = keyboard.on_press(on_key_press, suppress=True)
            self.stop_hotkey_button.mouse_hook_temp = mouse.hook(on_mouse_click)
            # Start live preview polling
            try:
                self._hotkey_preview_job = self.root.after(80, _update_hotkey_preview)
            except Exception:
                pass
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

        # Set a timer to reset the button if no key is pressed
        def reset_button():
            if not hasattr(self, 'stop_hotkey') or not self.stop_hotkey:
                self.stop_hotkey_button.config(text="Set Stop Hotkey")
            self._hotkey_assignment_cancelled = True
            self.setting_hotkey = False
            finish_hotkey_assignment()
            
        self.unhook_timer = self.root.after(5000, reset_button)

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
        # Get voice descriptions for the dropdown menu
        voice_names = []
        if hasattr(self, 'voices') and self.voices:
            try:
                voice_names = [voice.GetDescription() for voice in self.voices]
            except Exception as e:
                print(f"Warning: Could not get voice descriptions: {e}")
                voice_names = []
        voice_menu = tk.OptionMenu(area_frame, voice_var, "Select Voice", *voice_names)
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
                settings = {
                    'preprocess': preprocess_var.get(),
                    'voice': voice_var.get(),
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
            if isinstance(widget, tk.OptionMenu) or isinstance(widget, ttk.Combobox):
                widget.bind('<<ComboboxSelected>>', lambda e: self.resize_window())
            widget.bind('<Configure>', lambda e: self.resize_window())
        for widget in area_frame.winfo_children():
            bind_resize_events(widget)
        area_frame.bind('<Configure>', lambda e: self.resize_window())

        # Automatically recompute layout without forcing geometry, so default window size is preserved
        self.resize_window(force=True)

    def remove_area(self, area_frame, area_name):
        # Find and clean up the hotkey for this area
        for area in self.areas:
            if area[0] == area_frame:  # Found matching frame
                hotkey_button = area[1]  # Get the hotkey button
                # Clean up keyboard hook if it exists
                if hasattr(hotkey_button, 'keyboard_hook'):
                    try:
                        # Only try to unhook if the hook exists and is not None
                        if hotkey_button.keyboard_hook:
                            keyboard.unhook(hotkey_button.keyboard_hook)
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
                            mouse.unhook(hotkey_button.mouse_hook)
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
        min_width = 950
        max_width = 1600
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
        min_required_width = max(min_width, 1087) 
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

        def on_click(event):
            nonlocal x1, y1
            # Store canvas coordinates
            x1 = event.x
            y1 = event.y
            canvas.bind("<B1-Motion>", on_drag)
            canvas.bind("<ButtonRelease-1>", on_release)
            # Initialize both rectangles at click point
            canvas.coords(border, x1, y1, x1, y1)
            canvas.coords(border_outline, x1, y1, x1, y1)

        def on_release(event):
            nonlocal x1, y1, x2, y2
            if not selection_cancelled:
                try:
                    # Stop speech on mouse release if the checkbox is checked
                    if hasattr(self, 'interrupt_on_new_scan_var') and self.interrupt_on_new_scan_var.get():
                        self.stop_speaking()
                    
                    # Convert canvas coordinates to screen coordinates for the final area
                    x2 = event.x_root
                    y2 = event.y_root
                    x1_screen = x1 + min_x
                    y1_screen = y1 + min_y
                    
                    # Only set coordinates if we have a valid selection (not a click)
                    if abs(x2 - x1_screen) > 5 and abs(y2 - y1_screen) > 5:  # Minimum 5px drag
                        frame.area_coords = (
                            min(x1_screen, x2), 
                            min(y1_screen, y2),
                            max(x1_screen, x2), 
                            max(y1_screen, y2)
                        )
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
                            area_name = simpledialog.askstring("Area Name", "Enter a name for this area:")
                            if area_name and area_name.strip():
                                area_name_var.set(area_name)
                                print(f"Set area: {frame.area_coords} with name {area_name_var.get()}\n--------------------------")
                            else:
                                messagebox.showerror("Error", "Area name cannot be empty.")
                                print("Error: Area name cannot be empty.")
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
        
        # Get the true multi-monitor dimensions using win32api
        monitors = win32api.EnumDisplayMonitors()
        min_x = min_y = max_x = max_y = 0
        
        for monitor in monitors:
            monitor_info = win32api.GetMonitorInfo(monitor[0])
            monitor_area = monitor_info['Monitor']
            min_x = min(min_x, monitor_area[0])
            min_y = min(min_y, monitor_area[1])
            max_x = max(max_x, monitor_area[2])
            max_y = max(max_y, monitor_area[3])
        
        virtual_width = max_x - min_x
        virtual_height = max_y - min_y
        
        # Set window to cover entire virtual screen
        select_area_window.geometry(f"{virtual_width}x{virtual_height}+{min_x}+{min_y}")
        
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
        
        # Create border rectangle with more visible red border
        border = canvas.create_rectangle(0, 0, 0, 0,
                                       outline='red',
                                       width=3,  # Increased width
                                       dash=(8, 4))  # Longer dashes, shorter gaps
        
        # Create second border for better visibility
        border_outline = canvas.create_rectangle(0, 0, 0, 0,
                                          outline='red',
                                          width=3,
                                          dash=(8, 4),
                                          dashoffset=6)  # Offset to create alternating pattern
        
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

    def _restore_hotkeys_after_selection(self):
        """Helper method to restore hotkeys after area selection"""
        if not hasattr(self, 'hotkeys_disabled_for_selection') or not self.hotkeys_disabled_for_selection:
            return
            
        try:
            self.restore_all_hotkeys()
            self.hotkeys_disabled_for_selection = False
            print("Hotkeys re-enabled after area selection")
            
            # Force focus back to the main window
            if hasattr(self, 'root') and self.root.winfo_exists():
                self.root.focus_force()
                
        except Exception as e:
            print(f"Error restoring hotkeys after area selection: {e}")
            # Ensure the flag is cleared even if there's an error
            self.hotkeys_disabled_for_selection = False

    def disable_all_hotkeys(self):
        """Disable all hotkeys for keyboard and mouse."""
        try:
            # Unhook all keyboard and mouse hooks
            keyboard.unhook_all()
            mouse.unhook_all()
            
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
        combo_state = {'non_modifier_pressed': False}

        # Live preview of currently held modifiers while waiting for a non-modifier key
        def _update_hotkey_preview():
            if self._hotkey_assignment_cancelled or not self.setting_hotkey:
                return
            try:
                mods = []
                if keyboard.is_pressed('left ctrl'): mods.append('L-CTRL')
                if keyboard.is_pressed('right ctrl'): mods.append('R-CTRL')
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
            # Normalize key name and detect side for modifiers (handles localized names)
            raw_name = (event.name or '').lower()
            side = None
            # Detect side hints
            if any(s in raw_name for s in ["right", "høyre", "rechts", "derecha", "droite", "destra", "dereito", "höger"]):
                side = 'right'
            if any(s in raw_name for s in ["left", "venstre", "links", "izquierda", "gauche", "sinistra", "esquerda", "vänster"]):
                side = 'left'
            # Base modifier detection
            base = None
            if any(x in raw_name for x in ["ctrl", "control", "strg"]):
                base = 'ctrl'
            elif any(x in raw_name for x in ["altgr", "option", "menu"]) or raw_name.strip() == 'alt':
                base = 'alt'
            elif any(x in raw_name for x in ["shift", "skift", "umschalt"]):
                base = 'shift'
            elif any(x in raw_name for x in ["windows", "win", "super", "meta", "cmd", "command", "lwin", "rwin"]):
                base = 'windows'
            # Compose canonical name for modifier or keep raw for normal keys
            if base:
                if base in ('ctrl','alt'):
                    name = f"{side or 'left'} {base}"  # default to left when side unknown
                elif base in ('shift','windows'):
                    name = base
                else:
                    name = base
            else:
                name = raw_name

            # Mark that a non-modifier was pressed
            if name not in ('ctrl','left ctrl','right ctrl','alt','left alt','right alt','shift','windows'):
                combo_state['non_modifier_pressed'] = True
            if name in ('ctrl', 'shift', 'alt', 'windows'):
                # Allow assigning a bare modifier when released, if user doesn't press another key
                # Start a short timer to check if still only this modifier is held
                def _assign_bare_modifier():
                    if self._hotkey_assignment_cancelled or not self.setting_hotkey:
                        return
                    try:
                        held = []
                        if keyboard.is_pressed('left ctrl'): held.append('left ctrl')
                        if keyboard.is_pressed('right ctrl'): held.append('right ctrl')
                        if keyboard.is_pressed('shift'): held.append('shift')
                        if keyboard.is_pressed('left alt'): held.append('left alt')
                        if keyboard.is_pressed('right alt'): held.append('right alt')
                        if keyboard.is_pressed('left windows') or keyboard.is_pressed('right windows') or keyboard.is_pressed('windows'):
                            held.append('windows')
                        # Only proceed if exactly one modifier is still held and matches side/base
                        if len(held) == 1:
                            only = held[0]
                            # Accept if same base and, when available, same side
                            if (base == 'ctrl' and (only in ['left ctrl','right ctrl'])) or \
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
                            disp = key_name.upper().replace('LEFT CTRL','L-CTRL').replace('RIGHT CTRL','R-CTRL') \
                                                .replace('LEFT ALT','L-ALT').replace('RIGHT ALT','R-ALT') \
                                                .replace('WINDOWS','WIN')
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
                    self.root.after(200, _assign_bare_modifier)
                except Exception:
                    pass
                return

            # Build combination string from held modifiers + key
            try:
                mods = []
                if keyboard.is_pressed('left ctrl'): mods.append('left ctrl')
                if keyboard.is_pressed('right ctrl'): mods.append('right ctrl')
                if keyboard.is_pressed('shift'): mods.append('shift')
                if keyboard.is_pressed('left alt'): mods.append('left alt')
                if keyboard.is_pressed('right alt'): mods.append('right alt')
                if keyboard.is_pressed('left windows') or keyboard.is_pressed('right windows') or keyboard.is_pressed('windows'):
                    mods.append('windows')
            except Exception:
                mods = []

            # Determine base key (support numpad keys explicitly)
            base_key = name
            if event.scan_code in self.numpad_scan_codes:
                sym = self.numpad_scan_codes[event.scan_code]
                base_key = f"numpad {sym}"

            # If base_key itself is a modifier, include it if not already in mods; otherwise avoid duplicate
            if base_key in ("ctrl", "shift", "alt", "windows", "left ctrl", "right ctrl", "left alt", "right alt"):
                combo_parts = (mods + [base_key]) if base_key not in mods else mods[:]
            else:
                combo_parts = mods + [base_key]
            key_name = "+".join(p for p in combo_parts if p)

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
                    display_name = button.hotkey.replace('num_', 'num:') if button.hotkey.startswith('num_') else button.hotkey
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
            display_name = key_name.replace('numpad ', 'NUMPAD ').replace('num_', 'num:') \
                                 .replace('left ctrl','L-CTRL').replace('right ctrl','R-CTRL') \
                                 .replace('left alt','L-ALT').replace('right alt','R-ALT') \
                                 .replace('windows','WIN')
            button.config(text=f"Set Hotkey: [ {display_name.upper()} ]")

            self.setup_hotkey(button, area_frame)
            self.setting_hotkey = False

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
                
            # Only proceed with setting hotkey if no duplicate was found

            button.hotkey = key_name
            display_name = key_name.replace('num_', 'num:') if key_name.startswith('num_') else key_name
            button.config(text=f"Set Hotkey: [ {display_name.upper()} ]")

            self.setup_hotkey(button, area_frame)
            self.setting_hotkey = False

            # Unhook both temp hooks if they exist
            if hasattr(button, 'keyboard_hook_temp'):

                keyboard.unhook(button.keyboard_hook_temp)
                delattr(button, 'keyboard_hook_temp')
            if hasattr(button, 'mouse_hook_temp'):

                mouse.unhook(button.mouse_hook_temp)
                delattr(button, 'mouse_hook_temp')

            finish_hotkey_assignment()
            # Guard: prevent any further hotkey assignment callbacks
            self.setting_hotkey = False

            return

            button.hotkey = key_name
            display_name = key_name.replace('num_', 'num:') if key_name.startswith('num_') else key_name
            button.config(text=f"Set Hotkey: [ {display_name.upper()} ]")

            self.setup_hotkey(button, area_frame)
            self.setting_hotkey = False

            # Unhook both temp hooks if they exist
            if hasattr(button, 'keyboard_hook_temp'):

                keyboard.unhook(button.keyboard_hook_temp)
                delattr(button, 'keyboard_hook_temp')
            if hasattr(button, 'mouse_hook_temp'):

                mouse.unhook(button.mouse_hook_temp)
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
            
            # Check if this is a left or right mouse button
            is_left_button = button_name in LEFT_MOUSE_BUTTONS or event.button == 1
            is_right_button = button_name in RIGHT_MOUSE_BUTTONS or event.button == 2
            
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
                button_name = f"button{event.button}"
    
                # Create a mock keyboard event for the mouse button
                mock_event = type('MockEvent', (), {
                    'name': button_name,
                    'scan_code': None,
                    'event_type': 'down'
                })
                on_key_press(mock_event)
                return
            
            # Create a mock keyboard event
            mock_event = type('MockEvent', (), {
                'name': f'button{event.button}',  # Format: button3, button4, etc.
                'scan_code': None
            })
            on_key_press(mock_event)
            
            # Check for hotkey conflicts
            key_name = f'button{event.button}'
            for area_frame2, hotkey_button2, _, area_name_var2, _, _, _ in self.areas:
                if hotkey_button2 is not button and hasattr(hotkey_button2, 'hotkey') and hotkey_button2.hotkey == key_name:
                    show_thinkr_warning(self, area_name_var2.get())
                    if hasattr(button, 'hotkey'):
                        display_name = button.hotkey.replace('num_', 'num:') if button.hotkey.startswith('num_') else button.hotkey
                        button.config(text=f"Set Hotkey: [ {display_name.upper()} ]")
                    else:
                        button.config(text="Set Hotkey")
                    return
            
            # Set the new hotkey
            button.hotkey = key_name
            display_name = key_name.replace('num_', 'num:') if key_name.startswith('num_') else key_name
            button.config(text=f"Set Hotkey: [ {display_name.upper()} ]")
            
            # Setup the hotkey
            self.setup_hotkey(button, area_frame)
            
            # Clean up temporary hooks
            if hasattr(button, 'keyboard_hook_temp'):
                keyboard.unhook(button.keyboard_hook_temp)
                delattr(button, 'keyboard_hook_temp')
            if hasattr(button, 'mouse_hook_temp'):
                mouse.unhook(button.mouse_hook_temp)
                delattr(button, 'mouse_hook_temp')
            
            # Finish hotkey assignment
            self.setting_hotkey = False
            self._hotkey_assignment_cancelled = False
            finish_hotkey_assignment()
            return
        # Clean up previous hooks
        if hasattr(button, 'keyboard_hook'):
            try:
                # button.keyboard_hook from keyboard.add_hotkey returns an id; prefer remove_hotkey
                try:
                    keyboard.remove_hotkey(button.keyboard_hook)
                except Exception:
                    try:
                        keyboard.unhook(button.keyboard_hook)
                    except Exception:
                        pass
                delattr(button, 'keyboard_hook')
            except Exception as e:
                print(f"Error cleaning up keyboard hook: {e}")
        if hasattr(button, 'mouse_hook'):
            try:
                mouse.unhook(button.mouse_hook)
                delattr(button, 'mouse_hook')
            except Exception as e:
                print(f"Error cleaning up mouse hook: {e}")
        button.config(text="Press any key...")
        self.setting_hotkey = True  # Enable hotkey assignment mode before installing hooks
        button.keyboard_hook_temp = keyboard.on_press(on_key_press)
        button.mouse_hook_temp = mouse.hook(on_mouse_click)

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
                
                self.setting_hotkey = False
                self._hotkey_assignment_cancelled = True
                if not hasattr(button, 'hotkey'):
                    button.config(text="Set Hotkey")
            except Exception as e:
                print(f"Warning: Error during hook cleanup: {e}")
                button.config(text="Set Hotkey")
        self.root.after(3000, unhook_mouse)

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
        
    def _load_layout_file(self, file_path):
        """Internal method to load a layout file"""
        if file_path:
            try:
                # Store the file path before loading
                self.layout_file.set(file_path)
                # Reset unsaved changes when loading a new file
                self._has_unsaved_changes = False
                
                with open(file_path, 'r') as f:
                    layout = json.load(f)
                    
                # Clear only user-added areas and processing settings (keep permanent area)
                if self.areas:
                    # Always keep the first (permanent) area
                    for area in self.areas[1:]:
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
                                               .replace('left ctrl','L-CTRL').replace('right ctrl','R-CTRL') \
                                               .replace('left alt','L-ALT').replace('right alt','R-ALT') \
                                               .replace('left shift','L-SHIFT').replace('right shift','R-SHIFT') \
                                               .replace('windows','WIN')
                    self.stop_hotkey_button.config(text=f"Stop Hotkey: [ {display_name.upper()} ]")
                    print(f"Loaded Stop hotkey: {saved_stop_hotkey}")

                # --- Handle Auto Read hotkey ---
                auto_read_hotkey = None
                if self.areas and hasattr(self.areas[0][1], 'hotkey'):
                    auto_read_hotkey = self.areas[0][1].hotkey
                    # Clear the existing auto-read hotkey before loading new ones
                    if auto_read_hotkey:
                        try:
                            keyboard.remove_hotkey(self.areas[0][1].hotkey_id)
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
                for area_info in layout.get("areas", []):
                    # Create a new area using add_read_area (removable, editable, normal name)
                    self.add_read_area(removable=True, editable_name=True, area_name=area_info["name"])
                    
                    # Get the newly created area (last one in the list)
                    area_frame, hotkey_button, set_area_button, area_name_var, preprocess_var, voice_var, speed_var = self.areas[-1]
                    
                    # Set the area coordinates
                    area_frame.area_coords = area_info["coords"]
                    
                    # Set the hotkey if it exists
                    if area_info["hotkey"]:
                        hotkey_button.hotkey = area_info["hotkey"]
                        display_name = area_info["hotkey"].replace('num_', 'num:') if area_info["hotkey"].startswith('num_') else area_info["hotkey"]
                        hotkey_button.config(text=f"Hotkey: [ {display_name.upper()} ]")
                        self.setup_hotkey(hotkey_button, area_frame)
                    
                    # Set preprocessing and voice settings
                    preprocess_var.set(area_info.get("preprocess", False))
                    # Check if the saved voice exists in current SAPI voices
                    if hasattr(self, 'voices') and self.voices:
                        try:
                            current_voice_names = [voice.GetDescription() for voice in self.voices]
                            if area_info.get("voice") in current_voice_names:
                                voice_var.set(area_info["voice"])
                        except Exception as e:
                            print(f"Warning: Could not validate voice: {e}")
                    speed_var.set(area_info.get("speed", "1.0"))
                    
                    # Load and store image processing settings
                    if "settings" in area_info:
                        self.processing_settings[area_info["name"]] = area_info["settings"].copy()
                        print(f"Loaded image processing settings for area: {area_info['name']}")
                        
                # Update preferred sizes during load
                self.resize_window()

                # Get coordinates from the loaded area
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
                        display_name = auto_read_hotkey.replace('num_', 'num:') if auto_read_hotkey.startswith('num_') else auto_read_hotkey
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
                            if hasattr(self.areas[0][1], 'hotkey_id'):
                                keyboard.remove_hotkey(self.areas[0][1].hotkey_id)
                        except (KeyError, AttributeError):
                            pass
                        self.areas[0][1].hotkey = None
                        self.areas[0][1].config(text="Set Hotkey")

                print(f"Layout loaded from {file_path}\n--------------------------")
                
                # Force-resize the window to fit the newly loaded layout
                self.resize_window(force=True)
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load layout: {str(e)}")
                print(f"Error loading layout: {e}")

    def validate_speed_key(self, event, speed_var):
        """Additional validation for speed entry key presses"""
        if event.char.isdigit() or event.keysym in ('BackSpace', 'Delete', 'Left', 'Right'):
            return
        return 'break'

    def setup_hotkey(self, button, area_frame):
        try:
            # Clean up any existing hooks for this button first
            self._cleanup_hooks(button)
            
            # Store area_frame if this is not a stop button
            if not hasattr(button, 'is_stop_button') and area_frame is not None:
                button.area_frame = area_frame
                
            # Only proceed if we have a valid hotkey
            if not hasattr(button, 'hotkey') or not button.hotkey:
                print(f"No hotkey set for button: {button}")
                return False
                
            print(f"Setting up hotkey for: {button.hotkey}")
            
            # Define the hotkey handler function (accept optional event for keyboard lib compatibility)
            def hotkey_handler(event=None, *args, **kwargs):
                try:
                    if self.setting_hotkey:
                        return
                    # If this is a single-key hotkey (no '+') and not a pure modifier, ignore when any modifier is held
                    try:
                        hk = getattr(button, 'hotkey', '') or ''
                        if '+' not in hk and hk not in ['ctrl','left ctrl','right ctrl','alt','left alt','right alt','shift','left shift','right shift','windows']:
                            if (keyboard.is_pressed('left ctrl') or keyboard.is_pressed('right ctrl') or
                                keyboard.is_pressed('shift') or keyboard.is_pressed('left alt') or
                                keyboard.is_pressed('right alt') or keyboard.is_pressed('left windows') or
                                keyboard.is_pressed('right windows') or keyboard.is_pressed('windows')):
                                return
                    except Exception:
                        pass
                        
                    # Handle stop button
                    if hasattr(button, 'is_stop_button'):
                        self.root.after_idle(self.stop_speaking)
                        return
                        
                    # Handle Auto Read area
                    area_info = self._get_area_info(button)
                    if area_info and area_info.get('name') == "Auto Read":
                        self.root.after_idle(lambda: self.set_area(
                            area_info['frame'], 
                            area_info['name_var'], 
                            area_info['set_area_btn']))
                        return
                        
                    # Prevent multiple rapid triggers
                    if hasattr(button, '_is_processing') and button._is_processing:
                        return
                        
                    button._is_processing = True
                    self.stop_speaking()
                    threading.Thread(
                        target=self.read_area, 
                        args=(button.area_frame,), 
                        daemon=True
                    ).start()
                    self.root.after(100, lambda: setattr(button, '_is_processing', False))
                except Exception as e:
                    print(f"Error in hotkey_handler: {e}")
                
            # Define the mouse click handler
            def on_mouse_click(event):
                try:
                    if (self.setting_hotkey or 
                        getattr(self, 'hotkeys_disabled_for_selection', False) or 
                        not isinstance(event, mouse.ButtonEvent) or 
                        event.event_type != mouse.DOWN):
                        return False
                        
                    current_button = f'button{event.button}'
                    if current_button != button.hotkey:
                        return False
                        
                    # Handle stop button
                    if hasattr(button, 'is_stop_button'):
                        self.root.after_idle(self.stop_speaking)
                        return True
                        
                    # Handle Auto Read area
                    area_info = self._get_area_info(button)
                    if area_info and area_info.get('name') == "Auto Read":
                        self.root.after_idle(lambda: self.set_area(
                            area_info['frame'], 
                            area_info['name_var'], 
                            area_info['set_area_btn']))
                        return True
                        
                    # Prevent multiple rapid triggers
                    if hasattr(button, '_is_processing') and button._is_processing:
                        return True
                        
                    button._is_processing = True
                    self.stop_speaking()
                    threading.Thread(
                        target=self.read_area, 
                        args=(button.area_frame,), 
                        daemon=True
                    ).start()
                    self.root.after(100, lambda: setattr(button, '_is_processing', False))
                    return True
                except Exception as e:
                    print(f"Error in on_mouse_click: {e}")
                return False
                
            # Set up the appropriate hook based on hotkey type
            if button.hotkey.startswith('button'):
                try:
                    button.mouse_hook = mouse.hook(on_mouse_click)
                    print(f"Mouse hook set up for {button.hotkey}")
                except Exception as e:
                    print(f"Error setting up mouse hook: {e}")
                    return False
            else:
                try:
                    # Use low-level hook for numpad keys to avoid mis-triggers from function keys
                    hotkey_str = button.hotkey
                    num_key = None
                    pure_modifier = hotkey_str in ['left ctrl','right ctrl','left alt','right alt','shift','left shift','right shift','windows']
                    if pure_modifier:
                        # Use low-level hook and normalize event names to match exact sided modifiers
                        def on_kb_event_mod(e):
                            try:
                                if getattr(e, 'event_type', None) != 'down' or self.setting_hotkey or getattr(self, 'hotkeys_disabled_for_selection', False):
                                    return
                                raw = (getattr(e, 'name', '') or '').lower()
                                side = None
                                if any(s in raw for s in ["right", "høyre", "rechts", "derecha", "droite", "destra", "dereito", "höger"]):
                                    side = 'right'
                                if any(s in raw for s in ["left", "venstre", "links", "izquierda", "gauche", "sinistra", "esquerda", "vänster"]):
                                    side = 'left'
                                base = None
                                if any(x in raw for x in ["ctrl", "control", "strg"]):
                                    base = 'ctrl'
                                elif any(x in raw for x in ["altgr", "option", "menu"]) or raw.strip() == 'alt':
                                    base = 'alt'
                                elif any(x in raw for x in ["shift", "skift", "umschalt"]):
                                    base = 'shift'
                                elif any(x in raw for x in ["windows", "win", "super", "meta", "cmd", "command", "lwin", "rwin"]):
                                    base = 'windows'
                                # For Shift, infer side by scan_code if not in the raw name
                                if base == 'shift' and not side:
                                    sc = getattr(e, 'scan_code', None)
                                    if sc == 42:
                                        side = 'left'
                                    elif sc == 54:
                                        side = 'right'
                                name_ev = (
                                    f"{side or 'left'} {base}" if base in ('ctrl','alt','shift')
                                    else (base or raw)
                                )
                                if name_ev == hotkey_str:
                                    hotkey_handler()
                            except Exception:
                                pass
                        button.keyboard_hook = keyboard.hook(on_kb_event_mod)
                        button.keyboard_hook_is_lowlevel = True
                        button._hotkey_callback = hotkey_handler
                        print(f"Keyboard low-level sided modifier hook set up for {hotkey_str}")
                        return True
                    if hotkey_str.startswith('num_'):
                        num_key = hotkey_str.replace('num_', '')
                        # Map to scancode
                        try:
                            num_scancode = next(sc for sc, sym in self.numpad_scan_codes.items() if sym == num_key)
                        except StopIteration:
                            num_scancode = None

                        if num_scancode is not None:
                            def on_kb_event(e):
                                try:
                                    if (getattr(e, 'event_type', None) == 'down' and
                                        getattr(e, 'scan_code', None) == num_scancode and
                                        not self.setting_hotkey and
                                        not getattr(self, 'hotkeys_disabled_for_selection', False)):
                                        # Ignore when modifier keys are held to prevent false triggers
                                        try:
                                            if (keyboard.is_pressed('ctrl') or
                                                keyboard.is_pressed('alt') or
                                                keyboard.is_pressed('shift') or
                                                keyboard.is_pressed('windows')):
                                                return
                                        except Exception:
                                            pass
                                        hotkey_handler()
                                except Exception:
                                    pass
                            button.keyboard_hook = keyboard.hook(on_kb_event)
                            button.keyboard_hook_is_lowlevel = True
                            print(f"Keyboard low-level hook set up for numpad scancode {num_scancode}")
                        else:
                            # Fallback to string hotkey if scancode not found
                            button.keyboard_hook = keyboard.add_hotkey(hotkey_str, hotkey_handler, suppress=False)
                            print(f"Keyboard hotkey set up for {hotkey_str}")
                    else:
                        # Non-numpad hotkeys use add_hotkey
                        button.keyboard_hook = keyboard.add_hotkey(hotkey_str, hotkey_handler, suppress=False)
                        print(f"Keyboard hotkey set up for {hotkey_str}")

                    # Store handler for potential fallback
                    button._hotkey_callback = hotkey_handler

                    # Start polling fallback for numpad keys for fullscreen robustness
                    try:
                        if num_key is not None:
                            vk = self.numpad_vk_codes.get(num_key)
                            if vk is not None:
                                self._start_numpad_fallback(button, vk, hotkey_handler)
                    except Exception as _e:
                        pass
                    
                    # Test if the hotkey is working by trying to register it again
                    # If it fails, it means the hotkey is already registered and working
                    try:
                        test_hook = keyboard.add_hotkey(hotkey_str, lambda: None, suppress=False)
                        keyboard.remove_hotkey(hotkey_str)
                        print(f"Hotkey {hotkey_str} is working properly")
                        
                        # Show status message for successful hotkey setup
                        if hasattr(self, 'status_label'):
                            self.status_label.config(text=f"Hotkey '{hotkey_str}' set successfully")
                            if hasattr(self, '_feedback_timer') and self._feedback_timer:
                                self.root.after_cancel(self._feedback_timer)
                            self._feedback_timer = self.root.after(3000, lambda: self.status_label.config(text=""))
                            
                    except Exception as test_e:
                        print(f"Warning: Hotkey {hotkey_str} might not work in all applications: {test_e}")
                        # Show warning message
                        if hasattr(self, 'status_label'):
                            self.status_label.config(text=f"Warning: Hotkey '{hotkey_str}' may not work in all applications")
                            if hasattr(self, '_feedback_timer') and self._feedback_timer:
                                self.root.after_cancel(self._feedback_timer)
                            self._feedback_timer = self.root.after(5000, lambda: self.status_label.config(text=""))
                        
                except Exception as e:
                    print(f"Error setting up keyboard hotkey: {e}")
                    # Try alternative method for problematic hotkeys
                    try:
                        print(f"Trying alternative hotkey method for {button.hotkey}")
                        button.keyboard_hook = keyboard.on_press(hotkey_handler)
                        print(f"Alternative keyboard hook set up for {button.hotkey}")
                    except Exception as alt_e:
                        print(f"Alternative method also failed: {alt_e}")
                        return False
                    
            return True
            
        except Exception as e:
            print(f"Error in setup_hotkey: {e}")
            return False
            
    def _cleanup_hooks(self, button):
        """Helper method to clean up existing hooks for a button"""
        try:
            # Stop any fullscreen fallback polling
            try:
                self._stop_numpad_fallback(button)
            except Exception:
                pass
            # Clean up mouse hook if it exists
            if hasattr(button, 'mouse_hook'):
                try:
                    # mouse._listener may be None or missing; guard access
                    handlers = getattr(getattr(mouse, '_listener', None), 'handlers', None)
                    if handlers is not None and button.mouse_hook in handlers:
                        mouse.unhook(button.mouse_hook)
                except Exception as e:
                    # Swallow any unhook errors; we'll just remove the attribute
                    print(f"Warning: Mouse unhook issue (ignored): {e}")
                finally:
                    try:
                        delattr(button, 'mouse_hook')
                    except Exception:
                        pass
            
            # Clean up keyboard hotkey if it exists
            if hasattr(button, 'keyboard_hook') or hasattr(button, 'keyboard_hook_type'):
                try:
                    # For keyboard hotkeys registered with add_hotkey, use remove_hotkey
                    if hasattr(button, 'hotkey') and button.hotkey:
                        hotkey_str = button.hotkey
                        # Special cleanup for press_key modifiers
                        if hasattr(button, 'keyboard_hook_type') and getattr(button, 'keyboard_hook_type', '') == 'press_key':
                            try:
                                keyboard.unhook_key(getattr(button, 'keyboard_hotkey_name', hotkey_str))
                            except Exception:
                                pass
                            for attr in ['keyboard_hook_type', 'keyboard_hotkey_name', 'keyboard_hook']:
                                if hasattr(button, attr):
                                    try:
                                        delattr(button, attr)
                                    except Exception:
                                        pass
                            return
                        # If we installed a low-level hook for numpad, unhook it directly
                        if hasattr(button, 'keyboard_hook_is_lowlevel') and button.keyboard_hook_is_lowlevel:
                            try:
                                keyboard.unhook(button.keyboard_hook)
                            except Exception:
                                pass
                            if hasattr(button, 'keyboard_hook_is_lowlevel'):
                                delattr(button, 'keyboard_hook_is_lowlevel')
                            delattr(button, 'keyboard_hook')
                            return
                        if hotkey_str.startswith('num_'):
                            # Convert numpad format to standard format for removal
                            num_key = hotkey_str.replace('num_', '')
                            if num_key in ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']:
                                hotkey_str = f"numpad {num_key}"
                            elif num_key == '+':
                                hotkey_str = "numpad +"
                            elif num_key == '-':
                                hotkey_str = "numpad -"
                            elif num_key == '*':
                                hotkey_str = "numpad *"
                            elif num_key == '/':
                                hotkey_str = "numpad /"
                            elif num_key == 'enter':
                                hotkey_str = "numpad enter"
                            elif num_key == '.':
                                hotkey_str = "numpad ."
                        
                        keyboard.remove_hotkey(hotkey_str)
                    delattr(button, 'keyboard_hook')
                except (ValueError, AttributeError) as e:
                    # Handle case where hotkey was already removed or is invalid
                    print(f"Keyboard hotkey already removed or invalid: {e}")
                except Exception as e:
                    print(f"Unexpected error cleaning up keyboard hotkey: {e}")
                    # Ensure we still remove the attribute even if removal fails
                    if hasattr(button, 'keyboard_hook'):
                        delattr(button, 'keyboard_hook')
                        
        except Exception as e:
            print(f"Unexpected error in _cleanup_hooks: {e}")
            # Make sure we don't leave any attributes behind
            for attr in ['mouse_hook', 'keyboard_hook']:
                if hasattr(button, attr):
                    try:
                        delattr(button, attr)
                    except:
                        pass

    def _start_numpad_fallback(self, button, vk_code, handler):
        """Start a polling fallback for numpad key press using GetAsyncKeyState.
        Only used when fullscreen mode is enabled and for numpad keys.
        """
        try:
            # If already running, stop first
            self._stop_numpad_fallback(button)
        except Exception:
            pass

        stop_event = threading.Event()
        setattr(button, '_fallback_stop_event', stop_event)
        setattr(button, '_fallback_last_down', False)
        setattr(button, '_fallback_vk_code', vk_code)

        def poll_loop():
            try:
                while not stop_event.is_set():
                    try:
                        state = ctypes.windll.user32.GetAsyncKeyState(vk_code)
                        is_down = bool(state & 0x8000)
                        last_down = getattr(button, '_fallback_last_down', False)
                        if is_down and not last_down:
                            # Edge: key pressed
                            setattr(button, '_fallback_last_down', True)
                            # Avoid triggering during hotkey-setting or selection
                            if getattr(self, 'setting_hotkey', False) or getattr(self, 'hotkeys_disabled_for_selection', False):
                                time.sleep(0.02)
                                continue
                            try:
                                self.root.after_idle(handler)
                            except Exception as _e:
                                pass
                        elif not is_down and last_down:
                            setattr(button, '_fallback_last_down', False)
                    except Exception:
                        pass
                    time.sleep(0.02)
            except Exception:
                pass

        t = threading.Thread(target=poll_loop, daemon=True)
        setattr(button, '_fallback_thread', t)
        try:
            t.start()
            print(f"Started fullscreen fallback for VK={vk_code}")
        except Exception as e:
            print(f"Could not start fullscreen fallback: {e}")

    def _stop_numpad_fallback(self, button):
        """Stop polling fallback if running for this button."""
        if hasattr(button, '_fallback_stop_event'):
            try:
                button._fallback_stop_event.set()
            except Exception:
                pass
        if hasattr(button, '_fallback_thread'):
            try:
                # Do not block forever
                button._fallback_thread.join(timeout=0.2)
            except Exception:
                pass
        for attr in ['_fallback_stop_event', '_fallback_thread', '_fallback_last_down', '_fallback_vk_code']:
            if hasattr(button, attr):
                try:
                    delattr(button, attr)
                except Exception:
                    pass

    # Removed fullscreen mode toggle handler; fallback is always managed per hotkey
            
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
            # Check both SAPI voices and our combined voice list (including mock voices)
            selected_voice = None
            
            # First try to find in SAPI voices
            try:
                voices = self.speaker.GetVoices()
                for voice in voices:
                    if voice.GetDescription() == voice_var.get():
                        selected_voice = voice
                        break
            except Exception as e:
                print(f"Error getting SAPI voices: {e}")
            
            # If not found in SAPI, try our combined voice list (includes mock voices)
            if not selected_voice and hasattr(self, 'voices'):
                for voice in self.voices:
                    if hasattr(voice, 'GetDescription') and voice.GetDescription() == voice_var.get():
                        # Check if this is a real SAPI voice object
                        if hasattr(voice, 'GetId') and hasattr(voice, 'GetToken'):  # Working OneCore voice object
                            print(f"Found working OneCore voice: {voice_var.get()}")
                            selected_voice = voice
                            break
                        elif hasattr(voice, 'GetId'):  # Real SAPI voice object
                            print(f"Found real voice in combined list: {voice_var.get()}")
                            selected_voice = voice
                            break
                        else:
                            # For mock voices, we can't set them directly, so just continue
                            print(f"Found mock voice: {voice_var.get()}")
                            selected_voice = "mock_voice"  # Mark as found but don't set
                            break
            
            if selected_voice and selected_voice != "mock_voice":
                try:
                    # If this is a OneCore voice, route through UWP immediately for reliability
                    if hasattr(selected_voice, 'GetToken'):
                        print(f"Using OneCore voice via Narrator: {voice_var.get()}")
                        if _ensure_uwp_available():
                            try:
                                loop = asyncio.new_event_loop()
                                asyncio.set_event_loop(loop)
                                loop.run_until_complete(self._speak_with_uwp(filtered_text, preferred_desc=voice_var.get()))
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
                print(f"Using OneCore (mock) voice via Narrator: {voice_var.get()}")
                if _ensure_uwp_available():
                    try:
                        loop = asyncio.new_event_loop()
                        asyncio.set_event_loop(loop)
                        loop.run_until_complete(self._speak_with_uwp(filtered_text, preferred_desc=voice_var.get()))
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
        webbrowser.open(url)
    except Exception as e:
        print(f"Error opening URL: {e}")

def capture_screen_area(x1, y1, x2, y2):
    """Capture screen area across multiple monitors using win32api"""
    # Get virtual screen bounds
    min_x = win32api.GetSystemMetrics(win32con.SM_XVIRTUALSCREEN)  # Leftmost x (can be negative)
    min_y = win32api.GetSystemMetrics(win32con.SM_YVIRTUALSCREEN)  # Topmost y (can be negative)
    total_width = win32api.GetSystemMetrics(win32con.SM_CXVIRTUALSCREEN)
    total_height = win32api.GetSystemMetrics(win32con.SM_CYVIRTUALSCREEN)
    max_x = min_x + total_width
    max_y = min_y + total_height

    # Clamp coordinates to virtual screen bounds
    x1 = max(min_x, min(max_x, x1))
    y1 = max(min_y, min(max_y, y1))
    x2 = max(min_x, min(max_x, x2))
    y2 = max(min_y, min(max_y, y2))

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
    
    # Use TkinterDnD's Tk if available, otherwise use regular tkinter
    if TKDND_AVAILABLE:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
    app = GameTextReader(root)
    # Create permanent area at the top
    app.add_read_area(removable=False, editable_name=False, area_name="Auto Read")
    # Try to load settings for Auto Read area from temp folder
    temp_path = os.path.join(tempfile.gettempdir(), 'GameReader', 'auto_read_settings.json')
    if os.path.exists(temp_path) and app.areas:
        try:
            with open(temp_path, 'r') as f:
                settings = json.load(f)
            # Find the permanent area and set its settings
            area_frame, hotkey_button, set_area_button, area_name_var, preprocess_var, voice_var, speed_var = app.areas[0]
            
            # Get preprocess state from root level
            preprocess_state = settings.get('preprocess', False)
            
            # Set UI controls from root level settings
            preprocess_var.set(preprocess_state)
            voice_var.set(settings.get('voice', 'Select Voice'))
            speed_var.set(settings.get('speed', '100'))
            
            # Set hotkey if it exists
            if settings.get('hotkey'):
                hotkey_button.hotkey = settings['hotkey']
                display_name = settings['hotkey'].replace('num_', 'num:') if settings['hotkey'].startswith('num_') else settings['hotkey']
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
    root.mainloop()
