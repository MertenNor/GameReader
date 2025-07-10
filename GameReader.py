###
###  I know.. This code is.. well not great, all made with AI, but it works. feel free to make any changes!
###

# Standard library imports
import datetime
import io
import json
import os
import datetime
import queue
import re
import sys
import threading
import time
import webbrowser
import tempfile
from typing import Optional
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
import subprocess

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)  # FIX DPI ON WINDOWS
except AttributeError:
    ctypes.windll.user32.SetProcessDPIAware()
except Exception as e:
    print(f"Warning: Could not set DPI awareness: {e}")

APP_VERSION = "0.8.2"

CHANGELOG = """
- Better handling of stopping speech and reinitializing the system.
- More robust handling of mouse buttons for hotkeys.
- URLs in the info window were broken. They now direct to your default browser when clicked.
"""


# Create a StringIO buffer to capture print statements
log_buffer = io.StringIO()

# Redirect standard output to the StringIO buffer
sys.stdout = log_buffer

# --- Custom Hotkey Conflict Warning Dialog (No Symbol, Styled OK Button) ---
def show_thinkr_warning(game_reader, area_name):
    """Display a warning dialog for hotkey conflicts and manage hotkey states."""
    # Disable all hotkeys when dialog is shown
    try:
        keyboard.unhook_all()
        mouse.unhook_all()
    except Exception as e:
        print(f"Error disabling hotkeys for warning dialog: {e}")

    # Create and configure the dialog window
    win = tk.Toplevel(game_reader.root)
    win.title("Hotkey Conflict Detected!")
    win.geometry("370x170")
    win.resizable(False, False)
    win.grab_set()  # Make dialog modal
    win.transient(game_reader.root)  # Tie to parent window

    # Center the dialog relative to the main window
    win.update_idletasks()
    x = game_reader.root.winfo_rootx() + (game_reader.root.winfo_width() // 2 - 185)
    y = game_reader.root.winfo_rooty() + (game_reader.root.winfo_height() // 2 - 85)
    win.geometry(f"370x170+{x}+{y}")

    # Remove any warning icon
    for child in win.winfo_children():
        if isinstance(child, tk.Label) and child.cget("image"):
            child.destroy()

    # Add message label
    msg = tk.Label(
        win,
        text=f"This key is already used by area:\n'{area_name}'.\n\nPlease choose a different hotkey.",
        font=("Helvetica", 12),
        wraplength=340,
        justify="center"
    )
    msg.pack(pady=(28, 6))

    # Add OK button
    btn = tk.Button(
        win,
        text="OK",
        width=12,
        height=1,
        font=("Helvetica", 11, "bold"),
        relief="raised",
        bd=2
    )
    btn.pack(pady=(6, 10))
    btn.focus_set()  # Focus for keyboard accessibility

    # Define close handler to restore hotkeys and destroy window
    def on_close():
        try:
            game_reader.restore_all_hotkeys()
        except Exception as e:
            print(f"Error restoring hotkeys: {e}")
        win.destroy()

    # Set up close actions
    btn.config(command=on_close)
    win.protocol("WM_DELETE_WINDOW", on_close)
    win.bind("<Return>", lambda e: on_close())

def restore_all_hotkeys(self):
    """
    Restore all hotkeys for the application, including area hotkeys and the stop hotkey.
    """
    try:
        # Restore hotkeys for each area
        for area_frame, hotkey_button, *rest in self.areas:
            area_frame, hotkey_button, *rest = area
            if hasattr(hotkey_button, 'hotkey'):
                try:
                    self.setup_hotkey(hotkey_button, area[0])
                except ValueError as e:
                    print(f"Warning: Error restoring hotkey for area: {e}")
        
        # Restore stop hotkey if it exists
        if hasattr(self, 'stop_hotkey_button') and hasattr(self.stop_hotkey_button, 'mock_button'):
            try:
                self.setup_hotkey(self.stop_hotkey_button.mock_button, None)
            except ValueError as e:
                print(f"Warning: Error restoring stop hotkey: {e}")
    
    except Exception as e:
        print(f"Warning: Error in restore_all_hotkeys: {e}")

class ConsoleWindow:
    def __init__(self, root, log_buffer, layout_file_var, latest_images, latest_area_name_var):
        self.window = tk.Toplevel(root)
        self.window.title("Debug Console")
        self.window.geometry("690x500")
        self.latest_images = latest_images
        self.layout_file_var = layout_file_var
        self.latest_area_name_var = latest_area_name_var
        self.photo = None
        self.MAX_LINES = 250
        self.log_lines = log_buffer.getvalue().splitlines()[-self.MAX_LINES:] if log_buffer else []
        
        self._create_top_frame()
        self._create_image_frame()
        self._create_log_frame()
        self.update_console()

    def _create_top_frame(self):
        top_frame = tk.Frame(self.window)
        top_frame.pack(fill='x', padx=10, pady=5)
        self.show_image_var = tk.BooleanVar(value=True)
        self.image_checkbox = tk.Checkbutton(
            top_frame, text="Show last processed image", variable=self.show_image_var,
            command=self.update_image_display
        )
        self.image_checkbox.pack(side='left')
        scale_frame = tk.Frame(top_frame)
        scale_frame.pack(side='left', padx=10)
        tk.Label(scale_frame, text="Scale:").pack(side='left')
        self.scale_var = tk.StringVar(value="100")
        scales = [str(i) for i in range(10, 101, 10)]
        tk.OptionMenu(scale_frame, self.scale_var, *scales, command=self.update_image_display).pack(side='left')
        tk.Label(scale_frame, text="%").pack(side='left')
        tk.Button(top_frame, text="Save Log", command=self.save_log).pack(side='left', padx=(10, 0))
        tk.Button(top_frame, text="Clear Log", command=self.clear_log).pack(side='left', padx=(10, 0))
        tk.Button(top_frame, text="Save Image", command=self.save_image).pack(side='left', padx=(10, 0))

    def _create_image_frame(self):
        image_frame = tk.Frame(self.window)
        image_frame.pack(fill='x', padx=10, pady=5)
        self.image_label = tk.Label(image_frame)
        self.image_label.pack(fill='x')

    def _create_log_frame(self):
        log_frame = tk.Frame(self.window)
        log_frame.pack(fill='both', expand=True, padx=10, pady=5)
        self.text_widget = tk.Text(log_frame)
        self.text_widget.pack(fill='both', expand=True)
        self.text_widget.config(state=tk.DISABLED)
        def _on_mousewheel_debug(event):
            self.text_widget.yview_scroll(int(-1 * (event.delta / 120)), 'units')
            return "break"
        def _bind_mousewheel_debug(event):
            self.text_widget.bind_all('<MouseWheel>', _on_mousewheel_debug)
        def _unbind_mousewheel_debug(event):
            self.text_widget.unbind_all('<MouseWheel>')
        self.text_widget.bind('<Enter>', _bind_mousewheel_debug)
        self.text_widget.bind('<Leave>', _unbind_mousewheel_debug)
        self.context_menu = tk.Menu(self.text_widget, tearoff=0)
        self.context_menu.add_command(label="Copy", command=self.copy_selection)
        self.context_menu.add_command(label="Select All", command=self.select_all)
        self.text_widget.bind("<Button-3>", self.show_context_menu)

    def show_context_menu(self, event):
        try:
            self.context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.context_menu.grab_release()

    def copy_selection(self):
        try:
            selected_text = self.text_widget.get("sel.first", "sel.last")
            self.window.clipboard_clear()
            self.window.clipboard_append(selected_text)
        except tk.TclError:
            pass

    def select_all(self):
        self.text_widget.tag_add("sel", "1.0", "end")

    def update_image_display(self, *args):
        if not self.window.winfo_exists():
            return
        area_name = self.latest_area_name_var.get()
        if self.show_image_var.get() and area_name in self.latest_images:
            image = self.latest_images[area_name]
            scale_factor = int(self.scale_var.get()) / 100
            if scale_factor != 1:
                new_width = int(image.width * scale_factor)
                new_height = int(image.height * scale_factor)
                image = image.resize((new_width, new_height), Image.Resampling.LANCZOS)
            window_height = max(500, image.height + 300)
            window_x, window_y, window_width = self.window.winfo_x(), self.window.winfo_y(), self.window.winfo_width()
            self.window.geometry(f"{window_width}x{window_height}+{window_x}+{window_y}")
            self.photo = ImageTk.PhotoImage(image)
            if self.image_label.winfo_exists():
                self.image_label.config(image=self.photo)
        else:
            if self.image_label.winfo_exists():
                self.image_label.config(image='')
            self.photo = None

    def update_console(self):
        if not self.text_widget.winfo_exists():
            return
        self.text_widget.config(state=tk.NORMAL)
        self.text_widget.delete(1.0, tk.END)
        text = '\n'.join(self.log_lines)
        self.text_widget.insert(tk.END, text)
        self.text_widget.config(state=tk.DISABLED)
        self.text_widget.see(tk.END)

    def write(self, message):
        if not self.window.winfo_exists():
            return
        lines = message.splitlines()
        for line in lines:
            self.log_lines.append(line)
        while len(self.log_lines) > self.MAX_LINES:
            self.log_lines.pop(0)
        self.update_console()
        if self.show_image_var.get():
            self.update_image_display()

    def flush(self):
        pass

    def save_log(self):
        current_time = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        save_file_name = os.path.splitext(os.path.basename(self.layout_file_var.get()))[0]
        suggested_name = f"Log_{save_file_name}_{current_time}.txt"
        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt", initialfile=suggested_name, filetypes=[("Text files", "*.txt")]
        )
        if file_path:
            with open(file_path, 'w') as f:
                f.write('\n'.join(self.log_lines))
            print(f"Log saved to {file_path}\n--------------------------")

    def save_image(self):
        if not self.window.winfo_exists():
            return
        area_name = self.latest_area_name_var.get()
        latest_image = self.latest_images.get(area_name)
        if not isinstance(latest_image, Image.Image):
            messagebox.showerror("Error", "No image to save.")
            return
        current_time = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        suggested_name = f"{area_name}_{current_time}.png"
        file_path = filedialog.asksaveasfilename(
            defaultextension=".png", initialfile=suggested_name, filetypes=[("PNG files", "*.png")]
        )
        if file_path:
            latest_image.save(file_path, "PNG")
            print(f"Image saved to {file_path}\n--------------------------")

    def clear_log(self):
        self.log_lines.clear()
        self.update_console()

class ImageProcessingWindow:
    def __init__(self, root, area_name, latest_images, settings, game_text_reader):
        self.window = tk.Toplevel(root)
        self.window.title(f"Image Processing for: {area_name}")
        self.area_name = area_name
        self.latest_images = latest_images
        self.settings = settings
        self.game_text_reader = game_text_reader
        self.presets_dir = os.path.join(tempfile.gettempdir(), 'GameReader', 'presets')
        os.makedirs(self.presets_dir, exist_ok=True)

        if area_name not in latest_images:
            messagebox.showerror("Error", "No image to process, generate an image by pressing the hotkey.")
            self.window.destroy()
            return

        self.image = latest_images[area_name]
        self.processed_image = self.image.copy()

        self.image_frame = ttk.Frame(self.window)
        self.image_frame.grid(row=0, column=0, columnspan=5, padx=10, pady=10)
        self.canvas = tk.Canvas(self.image_frame, width=self.image.width, height=self.image.height)
        self.canvas.pack()

        self.photo_image = ImageTk.PhotoImage(self.image)
        self.image_on_canvas = self.canvas.create_image(0, 0, anchor=tk.NW, image=self.photo_image)
        
        info_text = f"Showing previous image captured in area: {area_name}\n\nProcessing applies to unprocessed images; results may differ if the preview is already processed."
        info_label = ttk.Label(self.image_frame, text=info_text, font=("Helvetica", 12), justify='center')
        info_label.pack(pady=(10, 0), fill='x')

        control_frame = ttk.Frame(self.window)
        control_frame.grid(row=1, column=0, columnspan=5, pady=10)
        
        # Presets frame
        presets_frame = ttk.Frame(control_frame)
        presets_frame.pack(side='left', padx=10)

		# Add new buttons
        ttk.Button(presets_frame, text="Load Presets", command=self._load_preset_from_file).pack(side='left', padx=5)
        ttk.Button(presets_frame, text="Open Folder", command=self._open_presets_folder).pack(side='left', padx=5)
        
        ttk.Label(presets_frame, text="Presets:").pack(side='left')
        self.preset_var = tk.StringVar()
        self.preset_combobox = ttk.Combobox(presets_frame, textvariable=self.preset_var, state='readonly', width=20)
        self.preset_combobox.pack(side='left')
        self.preset_combobox['values'] = self._get_preset_list()
        self.preset_combobox.bind('<<ComboboxSelected>>', self._load_preset)
        ttk.Button(presets_frame, text="Save As...", command=self._save_preset_as).pack(side='left', padx=5)

        scale_frame = ttk.Frame(control_frame)
        scale_frame.pack(side='left', padx=10)
        
        ttk.Label(scale_frame, text="Preview Scale:").pack(side='left')
        self.scale_var = tk.StringVar(value="100")
        scales = [str(i) for i in range(10, 101, 10)]
        scale_menu = tk.OptionMenu(scale_frame, self.scale_var, *scales, command=lambda *args: self.display_image())
        scale_menu.pack(side='left')
        ttk.Label(scale_frame, text="%").pack(side='left')

        ttk.Button(control_frame, text="Apply img. processing", command=self.save_settings).pack(side='left', padx=10)
        ttk.Button(control_frame, text="Reset to default", command=self.reset_all).pack(side='left', padx=10)

        slider_configs = [
            {"label": "Brightness", "var_name": "brightness", "from_": 0.1, "to": 2.0, "initial": 1.0, "row": 2, "col": 0},
            {"label": "Contrast", "var_name": "contrast", "from_": 0.1, "to": 2.0, "initial": 1.0, "row": 2, "col": 1},
            {"label": "Saturation", "var_name": "saturation", "from_": 0.1, "to": 2.0, "initial": 1.0, "row": 2, "col": 2},
            {"label": "Sharpness", "var_name": "sharpness", "from_": 0.1, "to": 2.0, "initial": 1.0, "row": 2, "col": 3},
            {"label": "Blur", "var_name": "blur", "from_": 0.0, "to": 10.0, "initial": 0.0, "row": 2, "col": 4},
            {"label": "Threshold", "var_name": "threshold", "from_": 0, "to": 255, "initial": 128, "row": 3, "col": 0, "enabled_var_name": "threshold_enabled"},
            {"label": "Hue", "var_name": "hue", "from_": -1.0, "to": 1.0, "initial": 0.0, "row": 3, "col": 1},
            {"label": "Exposure", "var_name": "exposure", "from_": 0.1, "to": 2.0, "initial": 1.0, "row": 3, "col": 2},
        ]

        self.slider_vars = {}
        self.initializing = True  # Set flag to True during initialization
        for config in slider_configs:
            var_name = config["var_name"]
            initial = settings.get(var_name, config["initial"])
            if var_name == "threshold":
                self.slider_vars[var_name] = tk.IntVar(value=initial)
            else:
                self.slider_vars[var_name] = tk.DoubleVar(value=initial)
            if "enabled_var_name" in config:
                self.slider_vars[config["enabled_var_name"]] = tk.BooleanVar(value=settings.get(config["enabled_var_name"], False))
            self.create_slider(
                config["label"],
                self.slider_vars[var_name],
                config["from_"],
                config["to"],
                config["initial"],
                config["row"],
                config["col"],
                self.slider_vars.get(config.get("enabled_var_name"))
            )
        self.initializing = False  # Set flag to False after all sliders are created
        print(f"Created sliders for {len(slider_configs)} configurations")  # Debug print to verify slider creation
        self.update_image()  # Optional: Update image with initial settings

    def _get_preset_list(self):
        """Retrieve the list of available presets from the presets directory."""
        if not os.path.exists(self.presets_dir):
            return []
        return [f[:-7] for f in os.listdir(self.presets_dir) if f.endswith('.preset')]

    def _load_preset(self, event=None):
        """Load the selected preset and apply its settings to the sliders."""
        selected = self.preset_var.get()
        if not selected:
            return
        preset_path = os.path.join(self.presets_dir, f"{selected}.preset")
        if os.path.exists(preset_path):
            try:
                with open(preset_path, 'r') as f:
                    preset_data = json.load(f)
                for key, value in preset_data.items():
                    if key in self.slider_vars:
                        self.slider_vars[key].set(value)
                self.update_image()
                print(f"Loaded preset: {selected}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load preset: {str(e)}")
        else:
            messagebox.showerror("Error", "Selected preset not found.")

    def _save_preset_as(self):
        """Save the current slider settings as a new preset."""
        preset_name = simpledialog.askstring("Save Preset", "Enter a name for the preset:")
        if not preset_name:
            return
        preset_path = os.path.join(self.presets_dir, f"{preset_name}.preset")
        if os.path.exists(preset_path):
            if not messagebox.askyesno("Overwrite", f"Preset '{preset_name}' already exists. Overwrite?"):
                return
        try:
            preset_data = {key: var.get() for key, var in self.slider_vars.items()}
            with open(preset_path, 'w') as f:
                json.dump(preset_data, f, indent=4)
            print(f"Saved preset: {preset_name}")
            self.preset_combobox['values'] = self._get_preset_list()
            self.preset_var.set(preset_name)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save preset: {str(e)}")
    
    def _load_preset_from_file(self):
        file_path = filedialog.askopenfilename(
            initialdir=self.presets_dir,
            title="Select Preset File",
            filetypes=(("Preset files", "*.preset"), ("All files", "*.*"))
        )
        if not file_path:
            return

        try:
            with open(file_path, 'r') as f:
                new_preset_data = json.load(f)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load preset: {str(e)}")
            return

        preset_name = os.path.splitext(os.path.basename(file_path))[0]
        existing_presets = self._get_preset_list()

        if preset_name in existing_presets:
            existing_path = os.path.join(self.presets_dir, f"{preset_name}.preset")
            with open(existing_path, 'r') as f:
                existing_data = json.load(f)
            if existing_data == new_preset_data:
                print(f"Preset '{preset_name}' already exists with same configuration.")
                return
            else:
                # Find a new unique name
                i = 1
                while True:
                    new_name = f"{preset_name} ({i})"
                    if new_name not in existing_presets:
                        break
                    i += 1
                preset_name = new_name

        # Save the new preset to the presets directory
        new_preset_path = os.path.join(self.presets_dir, f"{preset_name}.preset")
        with open(new_preset_path, 'w') as f:
            json.dump(new_preset_data, f, indent=4)

        # Update the combobox with the new preset list
        self.preset_combobox['values'] = self._get_preset_list()
        self.preset_var.set(preset_name)
        print(f"Loaded and saved new preset: {preset_name}")
        
    def _open_presets_folder(self):
        if os.path.exists(self.presets_dir):
            try:
                os.startfile(self.presets_dir)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to open folder: {str(e)}")
        else:
            messagebox.showerror("Error", "Presets directory does not exist.")
    
    def create_slider(self, label, variable, from_, to, initial, row, col, enabled_var=None):
        frame = ttk.Frame(self.window)
        frame.grid(row=row, column=col, padx=10, pady=5)

        label_frame = ttk.LabelFrame(frame, text=label)
        label_frame.pack(fill='both', expand=True)
    
        ttk.Label(label_frame, text=label).pack()

        entry_var = tk.StringVar(value=f'{initial:.2f}')
        variable.trace_add('write', lambda *args: entry_var.set(f'{variable.get():.2f}'))

        slider = ttk.Scale(label_frame, from_=from_, to=to, orient='horizontal', variable=variable, command=self.update_image)
        slider.set(initial)
        slider.pack()

        entry = ttk.Entry(label_frame, textvariable=entry_var)
        entry.pack()
        
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
        self.slider_vars["brightness"].set(1.0)
        self.slider_vars["contrast"].set(1.0)
        self.slider_vars["saturation"].set(1.0)
        self.slider_vars["sharpness"].set(1.0)
        self.slider_vars["blur"].set(0.0)
        self.slider_vars["threshold"].set(128)
        self.slider_vars["hue"].set(0.0)
        self.slider_vars["exposure"].set(1.0)
        self.slider_vars["threshold_enabled"].set(False)
        self.update_image()

    def update_image(self, _=None):
        if self.initializing or not self.image:  # Skip if initializing or no image
            return
        if self.processed_image:
            self.processed_image.close()
        self.processed_image = self.image.copy()

        enhancer = ImageEnhance.Brightness(self.processed_image)
        self.processed_image = enhancer.enhance(self.slider_vars["brightness"].get())
        enhancer = ImageEnhance.Contrast(self.processed_image)
        self.processed_image = enhancer.enhance(self.slider_vars["contrast"].get())
        enhancer = ImageEnhance.Color(self.processed_image)
        self.processed_image = enhancer.enhance(self.slider_vars["saturation"].get())
        enhancer = ImageEnhance.Sharpness(self.processed_image)
        self.processed_image = enhancer.enhance(self.slider_vars["sharpness"].get())
        if self.slider_vars["blur"].get() > 0:
            self.processed_image = self.processed_image.filter(ImageFilter.GaussianBlur(self.slider_vars["blur"].get()))
        if self.slider_vars["threshold_enabled"].get():
            self.processed_image = self.processed_image.point(lambda p: p > self.slider_vars["threshold"].get() and 255)
        self.processed_image = self.processed_image.convert('HSV')
        channels = list(self.processed_image.split())
        channels[0] = channels[0].point(lambda p: (p + int(self.slider_vars["hue"].get() * 255)) % 256)
        self.processed_image = Image.merge('HSV', channels).convert('RGB')
        enhancer = ImageEnhance.Brightness(self.processed_image)
        self.processed_image = enhancer.enhance(self.slider_vars["exposure"].get())

        self.display_image()

    def display_image(self):
        scale_factor = int(self.scale_var.get()) / 100
        if scale_factor != 1:
            new_width = int(self.processed_image.width * scale_factor)
            new_height = int(self.processed_image.height * scale_factor)
            display_image = self.processed_image.resize((new_width, new_height), Image.Resampling.LANCZOS)
        else:
            display_image = self.processed_image
        self.photo_image = ImageTk.PhotoImage(display_image)
        self.canvas.itemconfig(self.image_on_canvas, image=self.photo_image)
        self.canvas.config(width=display_image.width, height=display_image.height)

    def save_settings(self):
        self.settings['brightness'] = self.slider_vars["brightness"].get()
        self.settings['contrast'] = self.slider_vars["contrast"].get()
        self.settings['saturation'] = self.slider_vars["saturation"].get()
        self.settings['sharpness'] = self.slider_vars["sharpness"].get()
        self.settings['blur'] = self.slider_vars["blur"].get()
        self.settings['hue'] = self.slider_vars["hue"].get()
        self.settings['exposure'] = self.slider_vars["exposure"].get()
        self.settings['threshold'] = self.slider_vars["threshold"].get() if self.slider_vars["threshold_enabled"].get() else None
        self.settings['threshold_enabled'] = self.slider_vars["threshold_enabled"].get()

        self.game_text_reader.processing_settings[self.area_name] = self.settings.copy()

        for area_frame, _, _, area_name_var, preprocess_var, *rest in self.game_text_reader.areas:
            if area_name_var.get() == self.area_name:
                preprocess_var.set(True)
                break

        if self.area_name == "Auto Read":
            self.save_auto_read_settings()
            self.window.destroy()
        else:
            if not self.game_text_reader.layout_file.get():
                self.prompt_save_layout()
            self.game_text_reader.save_layout()
            self.window.destroy()

    def save_auto_read_settings(self):
        preprocess = None
        voice = None
        speed = None
        hotkey = None
        for area_frame, _, _, area_name_var, preprocess_var, voice_var, speed_var, *rest in self.game_text_reader.areas:
            if area_name_var.get() == self.area_name:
                preprocess = preprocess_var.get() if hasattr(preprocess_var, 'get') else preprocess_var
                voice = voice_var.get() if hasattr(voice_var, 'get') else voice_var
                speed = speed_var.get() if hasattr(speed_var, 'get') else speed_var
                break
        for area_frame, hotkey_button, _, area_name_var, *rest in self.game_text_reader.areas:
            if area_name_var.get() == self.area_name:
                hotkey = getattr(hotkey_button, 'hotkey', None)
                break
        settings = {
            'preprocess': preprocess,
            'voice': voice,
            'speed': speed,
            'brightness': self.slider_vars["brightness"].get(),
            'contrast': self.slider_vars["contrast"].get(),
            'saturation': self.slider_vars["saturation"].get(),
            'sharpness': self.slider_vars["sharpness"].get(),
            'blur': self.slider_vars["blur"].get(),
            'hue': self.slider_vars["hue"].get(),
            'exposure': self.slider_vars["exposure"].get(),
            'threshold': self.slider_vars["threshold"].get() if self.slider_vars["threshold_enabled"].get() else None,
            'threshold_enabled': self.slider_vars["threshold_enabled"].get(),
            'hotkey': hotkey,
            'stop_read_on_select': getattr(self.game_text_reader, 'interrupt_on_new_scan_var', tk.BooleanVar(value=True)).get()
        }
        temp_path = os.path.join(tempfile.gettempdir(), 'auto_read_settings.json')
        with open(temp_path, 'w') as f:
            json.dump(settings, f)
        if hasattr(self.game_text_reader, 'status_label'):
            self.game_text_reader.status_label.config(text="Auto Read area settings saved (auto)")
            if hasattr(self.game_text_reader, '_feedback_timer') and self.game_text_reader._feedback_timer:
                self.game_text_reader.root.after_cancel(self.game_text_reader._feedback_timer)
            self.game_text_reader._feedback_timer = self.game_text_reader.root.after(2000, lambda: self.game_text_reader.status_label.config(text=""))

    def prompt_save_layout(self):
        dialog = tk.Toplevel(self.window)
        dialog.title("No Save File")
        dialog.geometry("400x150")
        dialog.transient(self.window)
        dialog.grab_set()
        message = tk.Label(dialog, text="No save file exists. You need to save the layout\nto preserve these settings.\n\nCreate save file now?", pady=20)
        message.pack()
        button_frame = tk.Frame(dialog)
        button_frame.pack(pady=10)
        yes_button = tk.Button(button_frame, text="Yes", command=lambda: [dialog.destroy(), self.game_text_reader.save_layout()], width=10)
        yes_button.pack(side='left', padx=10)
        no_button = tk.Button(button_frame, text="No", command=dialog.destroy, width=10)
        no_button.pack(side='left', padx=10)
        dialog.update_idletasks()
        width = dialog.winfo_width()
        height = dialog.winfo_height()
        x = (dialog.winfo_screenwidth() // 2) - (width // 2)
        y = (dialog.winfo_screenheight() // 2) - (height // 2)
        dialog.geometry(f'{width}x{height}+{x}+{y}')
        self.window.wait_window(dialog)

    def update_preview(self, *args):
        self.display_image()

def preprocess_image(image: Image, brightness: float = 1.0, contrast: float = 1.0, saturation: float = 1.0, sharpness: float = 1.0, blur: float = 0.0, threshold: Optional[int] = None, hue: float = 0.0, exposure: float = 1.0) -> Image:
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

    # Apply blur if specified
    if blur > 0:
        image = image.filter(ImageFilter.GaussianBlur(blur))

    # Apply threshold to grayscale image if specified
    if threshold is not None:
        image = image.convert('L')
        image = image.point(lambda p: 255 if p > threshold else 0)

    # Apply hue adjustment if specified
    if hue != 0.0:
        image = image.convert('HSV')
        channels = list(image.split())
        channels[0] = channels[0].point(lambda p: (p + int(hue * 255)) % 256)
        image = Image.merge('HSV', channels).convert('RGB')

    # Apply exposure
    enhancer = ImageEnhance.Brightness(image)
    image = enhancer.enhance(exposure)

    return image

def extract_changelog_from_code(code):
    """Extracts the CHANGELOG string from the code."""
    match = re.search(r'CHANGELOG\s*=\s*([ru]?)(["\\']{3})(.*?)\2', code, re.DOTALL)
    if match:
        return match.group(3).strip()
    return None

def check_for_update(local_version, force=False):  # for testing the update window. False for release.
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
                popup = tk.Toplevel()
                popup.title("Update Available")
                popup.geometry("750x350")  # Set initial size
                popup.minsize(400, 150)    # Set minimum size
                
                # Make window resizable
                popup.resizable(True, comfortably=True)
                
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

def remove_json_comments(content):
    """Remove single-line (//) and multi-line (/* */) comments from JSON content."""
    content = re.sub(r'//.*?$', '', content, flags=re.MULTILINE)
    content = re.sub(r'/\*.*?\*/', '', content, flags=re.DOTALL)
    return content


class GameTextReader:
    def __init__(self, root):
        """Initialize the GameTextReader application with GUI, TTS, and hotkey functionality."""
        self.root = root
        self.root.title(f"Game Reader v{APP_VERSION}")
        # Update check on startup
        local_version = APP_VERSION
        FORCE_UPDATE_CHECK = False
        threading.Thread(target=lambda: check_for_update(local_version, force=FORCE_UPDATE_CHECK), daemon=True).start()

        self.root.geometry("1115x180")
        self.layout_file = tk.StringVar()
        self.latest_images = {}
        self.latest_area_name = tk.StringVar()
        self.areas = []
        self.stop_hotkey = None

        # TTS initialization
        self.engine = None
        self.engine_lock = threading.Lock()
        self.speaker = None
        self.volume = tk.StringVar(value="100")
        try:
            self.engine = pyttsx3.init()
            _ = self.engine.getProperty('rate')
            self.speaker = win32com.client.Dispatch("SAPI.SpVoice")
            self.speaker.Volume = int(self.volume.get())
            # Check for RealtimeTTS availability
            self.realtimetts_available = False
            self.available_realtimetts_engines = []
            self.current_stream = None  # To manage RealtimeTTS stream
            try:
                import RealtimeTTS
                self.realtimetts_available = True
                possible_engines = [
                    "SystemEngine", "AzureEngine", "ElevenlabsEngine", "CoquiEngine",
                    "OpenAIEngine", "GTTSEngine", "EdgeEngine", "ParlerEngine",
                    "StyleTTSEngine", "PiperEngine", "KokoroEngine", "OrpheusEngine"
                ]
                for engine_name in possible_engines:
                    try:
                        engine_class = getattr(RealtimeTTS, engine_name)
                        self.available_realtimetts_engines.append(engine_name)
                    except AttributeError:
                        pass
            except ImportError:
                pass

            # TTS engine selection variables
            self.tts_engine_var = tk.StringVar(value="Windows TTS")
            self.realtimetts_engine_var = tk.StringVar(value="")
            if self.realtimetts_available and self.available_realtimetts_engines:
                self.realtimetts_engine_var.set(self.available_realtimetts_engines[0])
        except Exception as e:
            print(f"Warning: Could not initialize text-to-speech: {e}")
            print("Text-to-speech functionality may be limited.")
            self.engine = None if not self.engine else self.engine

        self.bad_word_list = tk.StringVar()
        self.hotkeys = set()
        self.is_speaking = False
        self.processing_settings = {}

        # Checkbox variables
        self.ignore_usernames_var = tk.BooleanVar(value=False)
        self.ignore_previous_var = tk.BooleanVar(value=False)
        self.ignore_gibberish_var = tk.BooleanVar(value=False)
        self.pause_at_punctuation_var = tk.BooleanVar(value=False)
        self.fullscreen_mode_var = tk.BooleanVar(value=False)
        self.better_unit_detection_var = tk.BooleanVar(value=False)
        self.read_game_units_var = tk.BooleanVar(value=False)
        self.allow_mouse_buttons_var = tk.BooleanVar(value=False)

        # Hotkey management
        self.setting_hotkey = False
        self.unhook_timer = None
        self.keyboard_hooks = []
        self.mouse_hooks = []
        pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
        self.numpad_scan_codes = {
            82: '0', 79: '1', 80: '2', 81: '3', 75: '4',
            76: '5', 77: '6', 71: '7', 72: '8', 73: '9',
            55: '*', 78: '+', 74: '-', 83: '.', 53: '/',
            28: 'enter'
        }  # Maps scan codes to numpad keys

        self.MODIFIERS = ['ctrl', 'shift', 'alt']  # Modifiers to support
        self.ALL_MODIFIERS = ['ctrl', 'shift', 'alt']  # Used for exact modifier matching
        self.DISPLAY_NAMES = {
            'ctrl': 'Ctrl',
            'shift': 'Shift',
            'alt': 'Alt',
            'mouse_left': 'Left Click',
            'mouse_right': 'Right Click',
            'mouse_middle': 'Middle Click',
            'mouse_x': 'Mouse Button 4',
            'mouse_x2': 'Mouse Button 5',
            }

        # Initialize text_histories as a dictionary and add a lock for thread safety
        self.text_histories = {}
        self.history_lock = threading.Lock()
        self.tts_queue = queue.Queue()
        
        self.game_units = self.load_game_units()

        self.setup_gui()
        self.voices = self.engine.getProperty('voices') if self.engine else []
        self.current_voices = self.voices

        self.stop_keyboard_hook = None
        self.stop_mouse_hook = None
        self.setting_hotkey_mouse_hook = None
        self.unhook_timer = None

        root.protocol("WM_DELETE_WINDOW", lambda: (self.cleanup(), root.destroy()))
        self._has_unsaved_changes = False

        root.drop_target_register(DND_FILES)
        root.dnd_bind('<<Drop>>', self.on_drop)
        root.dnd_bind('<<DropEnter>>', lambda e: 'break')
        root.dnd_bind('<<DropPosition>>', lambda e: 'break')
        self.root.after(100, self._process_tts_queue)

    def set_hotkey_if_not_setting(self, button, area_frame):
        if not self.setting_hotkey:
            self.set_hotkey(button, area_frame)

    def set_stop_hotkey_if_not_setting(self):
        if not self.setting_hotkey:
            self.set_stop_hotkey()

    def get_display_name(self, hotkey):
        parts = hotkey.split('+')
        display_parts = []
        for part in parts:
            if part in self.DISPLAY_NAMES:
                display_parts.append(self.DISPLAY_NAMES[part])
            elif part.startswith('num '):
                display_parts.append('Num ' + part[4:])
            else:
                display_parts.append(part.title())
        return '+'.join(display_parts)

    def speak_text(self, text, voice, speed):
        """Speak text using the selected TTS engine."""
        if self.tts_engine_var.get() == "Windows TTS":
            payload = (text, voice, speed)
            self.tts_queue.put(("SPEAK", payload))
        else:
            # RealtimeTTS
            if not self.realtimetts_available:
                messagebox.showerror("Error", "RealtimeTTS is not installed.")
                return
            engine_name = self.realtimetts_engine_var.get()
            if not engine_name:
                messagebox.showerror("Error", "No RealtimeTTS engine selected.")
                return
            try:
                import RealtimeTTS
                # Handle engines requiring specific parameters
                if engine_name == "AzureEngine":
                    speech_key = os.environ.get("AZURE_SPEECH_KEY")
                    region = os.environ.get("AZURE_SPEECH_REGION")
                    if not speech_key or not region:
                        raise ValueError("AZURE_SPEECH_KEY and AZURE_SPEECH_REGION must be set in environment variables.")
                    engine = RealtimeTTS.AzureEngine(speech_key, region)
                elif engine_name == "ElevenlabsEngine":
                    api_key = os.environ.get("ELEVENLABS_API_KEY")
                    if not api_key:
                        raise ValueError("ELEVENLABS_API_KEY must be set in environment variables.")
                    engine = RealtimeTTS.ElevenlabsEngine(api_key)
                elif engine_name == "OpenAIEngine":
                    # Assumes OPENAI_API_KEY is set in environment
                    engine = RealtimeTTS.OpenAIEngine()
                else:
                    engine_class = getattr(RealtimeTTS, engine_name)
                    engine = engine_class()

                if voice != "Select Voice":
                    engine.set_voice(voice)

                # Handle interrupt or ongoing speech
                if getattr(self, 'interrupt_on_new_scan_var', None) and self.interrupt_on_new_scan_var.get():
                    self.stop_speaking()
                elif self.is_speaking:
                    print("Already speaking. Stop current speech to proceed.")
                    return

                stream = RealtimeTTS.TextToAudioStream(engine)
                self.current_stream = stream
                stream.feed(text)
                stream.play_async()
                self.is_speaking = True
                print("RealtimeTTS speech started.\n--------------------------")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to use RealtimeTTS engine {engine_name}: {str(e)}")
                self.is_speaking = False

    def stop_speaking(self):
        """Stop ongoing speech for the selected TTS engine."""
        if self.tts_engine_var.get() == "Windows TTS":
            if not hasattr(self, 'speaker') or self.speaker is None:
                self.is_speaking = False
                return
            try:
                self.speaker.Speak("", 2)  # SVSFPurgeBeforeSpeak = 2
                self.is_speaking = False
                print("Speech stopped.\n--------------------------")
            except AttributeError as e:
                print(f"Error stopping speech: Invalid speaker object - {e}")
                self.is_speaking = False
                try:
                    self.speaker = win32com.client.Dispatch("SAPI.SpVoice")
                    self.speaker.Volume = int(self.volume.get())
                except ValueError as e2:
                    print(f"Warning: Invalid volume value '{self.volume.get()}': {e2}")
                    self.speaker.Volume = 100
                except Exception as e2:
                    print(f"Warning: Could not reinitialize speech engine: {e2}")
                    self.speaker = None
            except Exception as e:
                print(f"Unexpected error stopping speech: {e}")
                self.is_speaking = False
        else:
            if hasattr(self, 'current_stream') and self.current_stream:
                try:
                    self.current_stream.stop()
                    print("RealtimeTTS stream stopped.\n--------------------------")
                except Exception as e:
                    print(f"Error stopping RealtimeTTS stream: {e}")
                finally:
                    self.current_stream = None
            self.is_speaking = False

    def restart_tesseract(self):
        """Forcefully stop the speech and reinitialize the system."""
        print("Forcing stop...")
        try:
            self.stop_speaking()
            print("System reinitialized. Audio stopped.\n--------------------------")
        except Exception as e:
            print(f"Error during forced stop: {e}\n--------------------------")

    def setup_gui(self):
        """Set up the GUI for GameTextReader."""
        # Main frames
        top_frame = tk.Frame(self.root)
        top_frame.pack(fill='x', padx=10, pady=5)
        
        control_frame = tk.Frame(self.root)
        control_frame.pack(fill='x', padx=10, pady=5)
        
        # TTS engine selection frame
        tts_frame = tk.Frame(self.root)
        tts_frame.pack(fill='x', padx=10, pady=5)

        tk.Label(tts_frame, text="TTS Engine:").pack(side='left')

        tts_options = ["Windows TTS"]
        if self.realtimetts_available:
            tts_options.append("RealtimeTTS")

        self.tts_engine_combobox = ttk.Combobox(tts_frame, textvariable=self.tts_engine_var, values=tts_options, state='readonly')
        self.tts_engine_combobox.pack(side='left', padx=5)
        self.tts_engine_combobox.bind('<<ComboboxSelected>>', self.on_tts_engine_selected)

        # RealtimeTTS engine selection (hidden initially)
        self.realtimetts_engine_label = tk.Label(tts_frame, text="RealtimeTTS Engine:")
        self.realtimetts_engine_combobox = ttk.Combobox(tts_frame, textvariable=self.realtimetts_engine_var, 
                                                        values=self.available_realtimetts_engines, state='readonly')
        self.realtimetts_engine_combobox.bind('<<ComboboxSelected>>', self.on_realtimetts_engine_selected)
        self.realtimetts_engine_label.pack_forget()
        self.realtimetts_engine_combobox.pack_forget()

        # Show RealtimeTTS engine dropdown if RealtimeTTS is selected
        if self.tts_engine_var.get() == "RealtimeTTS":
            self.realtimetts_engine_label.pack(side='left', padx=(10, 0))
            self.realtimetts_engine_combobox.pack(side='left', padx=5)
        
        options_frame = tk.Frame(self.root)
        options_frame.pack(fill='x', padx=10, pady=5)
        
        # Top frame: Title and volume controls
        tk.Label(top_frame, text=f"GameReader v{APP_VERSION}", font=("Helvetica", 12, "bold")).pack(side='left', padx=(0, 20))
        
        volume_frame = tk.Frame(top_frame)
        volume_frame.pack(side='left', padx=10)
        tk.Label(volume_frame, text="Volume %:").pack(side='left')
        vcmd = (self.root.register(lambda P: self.validate_numeric_input(P, is_speed=False)), '%P')
        tk.Entry(volume_frame, textvariable=self.volume, width=4, validate='all', validatecommand=vcmd).pack(side='left', padx=5)
        tk.Button(volume_frame, text="Set", command=self.set_volume).pack(side='left', padx=5)
        
        # Top frame: Right-aligned buttons
        buttons_frame = tk.Frame(top_frame)
        buttons_frame.pack(side='right')
        tk.Button(buttons_frame, text="Debug Window", command=self.show_debug).pack(side='left', padx=5)
        tk.Button(buttons_frame, text="Info/Help", command=self.show_info).pack(side='left', padx=5)
        
        # Control frame: Layout info and buttons
        layout_frame = tk.Frame(control_frame)
        layout_frame.pack(side='left', fill='x', expand=True)
        tk.Label(layout_frame, text="Loaded Layout:").pack(side='left')
        tk.Label(layout_frame, textvariable=self.layout_file, font=("Helvetica", 10, "bold")).pack(side='left', padx=5)
        
        layout_buttons_frame = tk.Frame(control_frame)
        layout_buttons_frame.pack(side='right')
        tk.Button(layout_buttons_frame, text="Program Saves...", command=self.open_game_reader_folder).pack(side='left', padx=5)
        tk.Button(layout_buttons_frame, text="Save Layout", command=self.save_layout).pack(side='left', padx=5)
        tk.Button(layout_buttons_frame, text="Load Layout..", command=self.load_layout).pack(side='left', padx=5)
        
        # Options frame: Word filtering
        filter_frame = tk.Frame(options_frame)
        filter_frame.pack(fill='x', pady=5)
        tk.Label(filter_frame, text="Ignored Word List:").pack(side='left')
        self.bad_word_entry = ttk.Entry(filter_frame, textvariable=self.bad_word_list)
        self.bad_word_entry.pack(side='left', fill='x', expand=True)
        
        # Context menu for bad word entry
        self.bad_word_menu = tk.Menu(self.root, tearoff=0)
        for label, cmd in [("Cut", '<<Cut>>'), ("Copy", '<<Copy>>'), ("Paste", '<<Paste>>')]:
            self.bad_word_menu.add_command(label=label, command=lambda c=cmd: self.bad_word_entry.event_generate(c))
        self.bad_word_menu.add_separator()
        self.bad_word_menu.add_command(label="Select All", command=lambda: self.bad_word_entry.selection_range(0, 'end'))
        self.bad_word_entry.bind('<Button-3>', lambda event: self.bad_word_menu.post(event.x_root, event.y_root))
        
        # Options frame: Checkboxes
        checkbox_frame = tk.Frame(options_frame)
        checkbox_frame.pack(fill='x', pady=5)
        for text, var in [
            ("Ignore usernames:", self.ignore_usernames_var),
            ("Ignore previous spoken words:", self.ignore_previous_var),
            ("Ignore gibberish:", self.ignore_gibberish_var),
            ("Better unit detection:", self.better_unit_detection_var),
            ("Read gamer units:", self.read_game_units_var),
            ("Fullscreen mode:", self.fullscreen_mode_var),
            ("Allow mouse left/right:", self.allow_mouse_buttons_var)
        ]:
            self.create_checkbox(checkbox_frame, text, var, side='left', padx=5)
        self.allow_mouse_buttons_var.set(False)  # Initialize after creation
        
        # Add area frame
        add_area_frame = tk.Frame(self.root)
        add_area_frame.pack(fill='x', padx=10, pady=5)
        tk.Button(add_area_frame, text="Add Read Area", command=self.add_read_area, font=("Helvetica", 10)).pack(side='left')
        
        self.status_frame = tk.Frame(add_area_frame)
        self.status_frame.pack(side='left', fill='x', expand=True, padx=10)
        self.status_label = tk.Label(self.status_frame, text="", font=("Helvetica", 10), fg="black")
        self.status_label.pack(side='top')
        self.stop_hotkey_button = tk.Button(add_area_frame, text="Set Stop Hotkey", command=lambda: self.set_stop_hotkey_if_not_setting())
        
        self.stop_hotkey_button = tk.Button(add_area_frame, text="Set Stop Hotkey", command=self.set_stop_hotkey)
        self.stop_hotkey_button.pack(side='right')
        
        # Scrollable area frame
        self.area_outer_frame = tk.Frame(self.root)
        self.area_outer_frame.pack(fill='both', expand=True, pady=5)
        self.area_canvas = tk.Canvas(self.area_outer_frame, highlightthickness=0)
        self.area_canvas.pack(side='left', fill='both', expand=True)
        self.area_scrollbar = tk.Scrollbar(self.area_outer_frame, orient='vertical', command=self.area_canvas.yview)
        self.area_scrollbar.pack(side='right', fill='y')
        
        # Optimized mouse wheel scrolling
        def _on_mousewheel(event):
            if self.area_canvas.bbox('all') and self.area_canvas.winfo_height() < (self.area_canvas.bbox('all')[3] - self.area_canvas.bbox('all')[1]):
                self.area_canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')
            return "break"
        self.area_canvas.bind_all('<MouseWheel>', _on_mousewheel)
        
        self.area_frame = tk.Frame(self.area_canvas)
        self.area_window = self.area_canvas.create_window((0, 0), window=self.area_frame, anchor='nw')
        self.area_canvas.configure(yscrollcommand=self.area_scrollbar.set)
        
        # Bind resizing
        def on_frame_configure(event):
            self.area_canvas.configure(scrollregion=self.area_canvas.bbox('all'))
            canvas_width = self.area_canvas.winfo_width()
            self.area_canvas.itemconfig(self.area_window, width=canvas_width)
        self.area_frame.bind('<Configure>', on_frame_configure)
        self.area_canvas.bind('<Configure>', on_frame_configure)
        
        self.area_scrollbar.pack_forget()  # Hide scrollbar initially
        self.root.bind("<Button-1>", self.remove_focus)
        
        print("GUI setup complete.")

    def on_tts_engine_selected(self, event):
        selected = self.tts_engine_var.get()
        if selected == "RealtimeTTS":
            self.realtimetts_engine_label.pack(side='left', padx=(10, 0))
            self.realtimetts_engine_combobox.pack(side='left', padx=5)
        else:
            self.realtimetts_engine_label.pack_forget()
            self.realtimetts_engine_combobox.pack_forget()
        self.update_current_voices()
        self.update_voice_dropdowns()

    def on_realtimetts_engine_selected(self, event):
        self.update_current_voices()
        self.update_voice_dropdowns()

    def update_current_voices(self):
        if self.tts_engine_var.get() == "Windows TTS":
            self.current_voices = self.voices
        elif self.tts_engine_var.get() == "RealtimeTTS":
            engine_name = self.realtimetts_engine_var.get()
            if engine_name:
                try:
                    import RealtimeTTS
                    if engine_name == "AzureEngine":
                        speech_key = os.environ.get("AZURE_SPEECH_KEY")
                        region = os.environ.get("AZURE_SPEECH_REGION")
                        if not speech_key or not region:
                            raise ValueError("AZURE_SPEECH_KEY and AZURE_SPEECH_REGION must be set.")
                        engine = RealtimeTTS.AzureEngine(speech_key, region)
                    elif engine_name == "ElevenlabsEngine":
                        api_key = os.environ.get("ELEVENLABS_API_KEY")
                        if not api_key:
                            raise ValueError("ELEVENLABS_API_KEY must be set.")
                        engine = RealtimeTTS.ElevenlabsEngine(api_key)
                    elif engine_name == "OpenAIEngine":
                        engine = RealtimeTTS.OpenAIEngine()
                    else:
                        engine_class = getattr(RealtimeTTS, engine_name)
                        engine = engine_class()
                    self.current_voices = engine.get_voices()
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to get voices for {engine_name}: {str(e)}")
                    self.current_voices = []
            else:
                self.current_voices = []
        else:
            self.current_voices = []

    def update_voice_dropdowns(self):
        for area in self.areas:
            _, _, _, _, _, voice_var, _, voice_menu = area
            voice_menu['menu'].delete(0, 'end')
            if self.current_voices:
                voice_menu['menu'].add_command(label="Select Voice", command=lambda: voice_var.set("Select Voice"))
                for voice in self.current_voices:
                    voice_menu['menu'].add_command(label=voice.name, command=lambda v=voice.name: voice_var.set(v))
                voice_var.set("Select Voice")
            else:
                voice_menu['menu'].add_command(label="No voices available", command=lambda: voice_var.set("No voices available"))
                voice_var.set("No voices available")

    def create_checkbox(self, parent, text, variable, side='top', padx=0, pady=2):
        """Helper method to create and return a checkbox with a label.
        
        Args:
            parent: The parent widget to contain the checkbox frame.
            text: The label text for the checkbox.
            variable: Tkinter variable (e.g., BooleanVar) for the checkbox state.
            side: Packing side for the frame (default: 'top').
            padx: Horizontal padding for the frame (default: 0).
            pady: Vertical padding for the frame (default: 2).
        
        Returns:
            tk.Frame: The frame containing the checkbox and label.
        """
        frame = tk.Frame(parent)
        frame.pack(side=side, padx=padx, pady=pady)
        
        checkbox = tk.Checkbutton(frame, variable=variable)
        checkbox.pack(side='right')
        
        label = tk.Label(frame, text=text)
        label.pack(side='right')
        
        return frame

    def set_volume(self):
        """Set the speaker volume based on the volume StringVar value."""
        try:
            vol = int(self.volume.get())
            if 0 <= vol <= 100:
                self.speaker.Volume = vol
                print(f"Program volume set to {vol}%\n--------------------------")
            else:
                raise ValueError("Volume out of range")
        except ValueError as e:
            self.volume.set("100")
            self.speaker.Volume = 100
            print(f"Invalid volume: {e}, set to 100")
        except Exception as e:
            print(f"Error setting volume: {e}")

    def remove_focus(self, event):
        """Remove focus from entry fields when clicking outside them.
        
        Args:
            event: The Tkinter event object from a mouse click.
        """
        if not isinstance(event.widget, tk.Entry):
            self.root.focus()
    
    def show_info(self):
        # Create Tkinter window with a modern look
        info_window = tk.Toplevel(self.root)
        style = ttk.Style()
        style.configure('TNotebook.Tab', font=('Helvetica', 14), padding=[10, 5])
        info_window.title("GameReader - Information")
        info_window.geometry("900x600")
        info_window.resizable(True, True)
        info_window.configure(bg='#f0f0f0')

        # Disable hotkeys to prevent interference
        try:
            keyboard.unhook_all()
            mouse.unhook_all()
        except Exception as e:
            print(f"Error unhooking hotkeys for info window: {e}")

        # Restore hotkeys on window close
        def on_info_close():
            for area in self.areas:
                area_frame, hotkey_button, *rest = area
                if hasattr(hotkey_button, 'hotkey'):
                    self.setup_hotkey(hotkey_button, area_frame)
            if hasattr(self, 'stop_hotkey_button') and hasattr(self.stop_hotkey_button, 'mock_button'):
                self.setup_hotkey(self.stop_hotkey_button.mock_button, None)
            info_window.destroy()

        info_window.protocol("WM_DELETE_WINDOW", on_info_close)
        info_window.bind('<Escape>', lambda e: on_info_close())

        # Set window icon silently if available
        try:
            info_window.iconbitmap('icon.ico')
        except:
            pass

        # Main frame with padding and background
        main_frame = ttk.Frame(info_window, padding="20", style='Custom.TFrame')
        style.configure('Custom.TFrame', background='#f0f0f0')
        main_frame.pack(fill='both', expand=True)

        # Create notebook for tabs above the content
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill='both', expand=True, pady=10)

        # About tab
        about_frame = ttk.Frame(notebook)
        notebook.add(about_frame, text="About")
        self.build_about_section(about_frame)

        # Instructions tab
        instructions_frame = ttk.Frame(notebook)
        notebook.add(instructions_frame, text="Instructions")
        self.build_content_text(instructions_frame)

        # Add close button at the bottom with styling
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(fill='x', pady=(10, 0))
        close_button = ttk.Button(bottom_frame, text="Close", command=info_window.destroy, style='Custom.TButton')
        style.configure('Custom.TButton', font=('Helvetica', 10, 'bold'))
        close_button.pack(side='right')
        close_button.bind("<Enter>", lambda e: close_button.config(style='Hover.TButton'))
        close_button.bind("<Leave>", lambda e: close_button.config(style='Custom.TButton'))
        style.configure('Hover.TButton', font=('Helvetica', 10, 'bold'), background='#d0d0d0')

        # Make window modal
        info_window.transient(self.root)
        info_window.grab_set()

    # Helper function for clickable labels
    def create_clickable_label(self, parent, text, url, font=("Helvetica", 10), foreground='black'):
        label = ttk.Label(parent, text=text, font=font, foreground=foreground, cursor='hand2')
        label.bind("<Button-1>", lambda e: open_url(url))
        label.bind("<Enter>", lambda e: label.configure(font=(*font, "underline")))
        label.bind("<Leave>", lambda e: label.configure(font=font))
        return label

    # Title section
    def build_title_section(self, parent):
        title_frame = ttk.Frame(parent)
        title_frame.pack(fill='x', pady=(0, 20))
        ttk.Label(title_frame, text=f"GameReader v{APP_VERSION}", font=("Helvetica", 16, "bold")).pack(side='left')
        return title_frame

    # Credits and links section
    def build_credits_section(self):
        content = [
            ("Program Information\n", 'bold'),
            ("---\n", 'normal'),
            ("Designer: MertenNor\n", 'normal'),
            ("Coder: Various AIs via ", 'normal'),
            ("Cursor", ('link', 'link_cursor')),
            (" and ", 'normal'),
            ("Windsurf", ('link', 'link_windsurf')),
            ("\n\n", 'normal'),
            ("Official Links\n", 'bold'),
            ("---\n", 'normal'),
            ("GitHub: ", 'bold'),
            ("GitHub.com/mertennor/gamereader", ('link', 'link_github')),
            ("\n\n", 'normal'),
            ("Support & Feedback\n", 'bold'),
            ("---\n", 'normal'),
            ("Buy me a Coffee ☕: ", 'bold'),
            ("BuyMeaCoffee.com/mertennor", ('link', 'link_coffee')),
            ("\n", 'normal'),
            ("Feedback: Want features or found bugs? Use this form: ", 'bold'),
            ("Forms.Gle/8YBU8atkgwjyzdM79", ('link', 'link_form')),
            ("\n\n", 'normal'),
        ]
        links = {
            'link_cursor': "https://www.cursor.com/",
            'link_windsurf': "https://windsurf.com/",
            'link_github': "https://github.com/MertenNor/GameReader",
            'link_coffee': "https://www.buymeacoffee.com/mertennor",
            'link_form': "https://forms.gle/8YBU8atkgwjyzdM79",
        }
        return content, links

    # Tesseract warning section
    def build_tesseract_warning(self):
        content = [
            ("⚠ IMPORTANT NOTICE\n", ('bold', 'red')),
            ("This program requires Tesseract OCR to function properly.\n", ('bold', 'red')),
            ("Default installation path: C:\\Program Files\n", ('bold', 'red')),
            ("Download the latest version here:\n", ('bold', 'red')),
            ("www.github.com/tesseract-ocr/tesseract/releases", ('link', 'link_tesseract')),
            ("\n", 'normal'),
            ("Note: Tesseract OCR is essential for text recognition.\n\n", ('bold', 'red')),
        ]
        links = {
            'link_tesseract': "https://github.com/tesseract-ocr/tesseract/releases",
        }
        return content, links

    def build_about_section(self, parent):
        # Create text widget with scrollbar
        text_frame = ttk.Frame(parent)
        text_frame.pack(fill='both', expand=True)
        scrollbar = ttk.Scrollbar(text_frame)
        scrollbar.pack(side='right', fill='y')
        text_widget = tk.Text(text_frame, wrap=tk.WORD, yscrollcommand=scrollbar.set, font=("Helvetica", 10))
        text_widget.pack(side='left', fill='both', expand=True)
        scrollbar.config(command=text_widget.yview)
        
        # Define tags
        text_widget.tag_configure('title', font=("Helvetica", 16, "bold"), justify='center')
        text_widget.tag_configure('normal', justify='left')
        text_widget.tag_configure('bold', font=("Helvetica", 10, "bold"))
        text_widget.tag_configure('red', foreground='red')
        text_widget.tag_configure('link', foreground='blue', underline=True)
        text_widget.tag_configure('toc_header', font=("Helvetica", 12, "bold"))
        
        # Insert title
        text_widget.insert('end', f"GameReader v{APP_VERSION}\n", 'title')
        # Set mark after title
        text_widget.mark_set('after_title', 'end')
        text_widget.mark_gravity('after_title', 'left')
        # Insert description
        description = "GameReader is a tool designed to read text from game screens and convert it to speech, making games more accessible.\n\n"
        text_widget.insert('end', description, 'normal')
        
        # Insert content with marks for headers
        credits_content, credits_links = self.build_credits_section()
        tesseract_content, tesseract_links = self.build_tesseract_warning()
        all_content = credits_content + tesseract_content
        all_links = {**credits_links, **tesseract_links}
        headers = []
        for i, (text, tags) in enumerate(all_content):
            if isinstance(tags, tuple) and 'bold' in tags or tags == 'bold':
                mark_name = f"about_header_{i}"
                text_widget.mark_set(mark_name, 'end')
                text_widget.mark_gravity(mark_name, 'left')
                headers.append((text.strip(), mark_name))
                print(f"Set mark {mark_name} at position {text_widget.index(mark_name)}")  # Debug mark position
            text_widget.insert('end', text, tags)
        
        # Bind events for content links
        def on_link_click(event):
            index = text_widget.index(f"@{event.x},{event.y}")
            tags = text_widget.tag_names(index)
            for tag in tags:
                if tag in all_links:
                    open_url(all_links[tag])
                    break
        text_widget.bind("<Button-1>", on_link_click)
        
        # Change cursor on hover for content links
        for link_tag in all_links.keys():
            text_widget.tag_bind(link_tag, "<Enter>", lambda e: text_widget.config(cursor="hand2"))
            text_widget.tag_bind(link_tag, "<Leave>", lambda e: text_widget.config(cursor=""))
        
        # Make text read-only
        text_widget.config(state='disabled')

    # Content text section (scrollable)
    def build_content_text(self, parent):
        # Main content frame
        content_frame = ttk.Frame(parent)
        content_frame.pack(fill='both', expand=True)

        # TOC frame on the left
        toc_frame = ttk.Frame(content_frame, width=200)
        toc_frame.pack(side='left', fill='y')
        toc_listbox = tk.Listbox(toc_frame, width=30, font=("Helvetica", 10))
        toc_listbox.pack(fill='both', expand=True)

        # Text frame on the right
        text_frame = ttk.Frame(content_frame)
        text_frame.pack(side='right', fill='both', expand=True, padx=10, pady=10)
        scrollbar = ttk.Scrollbar(text_frame)
        scrollbar.pack(side='right', fill='y')
        text_widget = tk.Text(text_frame, wrap=tk.WORD, yscrollcommand=scrollbar.set, font=("Helvetica", 10))
        text_widget.pack(side='left', fill='both', expand=True)
        scrollbar.config(command=text_widget.yview)

        # Define tags
        text_widget.tag_configure('bold', font=("Helvetica", 10, "bold"))

        # Insert instructional text and collect headers
        headers = []
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

        headers = []
        for i, (text, tag) in enumerate(info_text):
            if tag == 'bold':
                header_tag = f"header_{i}"
                text_widget.insert('end', text, ('bold', header_tag))
                headers.append((text.strip(), header_tag))
            else:
                text_widget.insert('end', text, tag)

        # Populate TOC listbox
        for header, _ in headers:
            toc_listbox.insert('end', header)

        # Bind TOC selection
        def on_toc_select(event):
            selection = toc_listbox.curselection()
            if selection:
                index = selection[0]
                _, header_tag = headers[index]
                ranges = text_widget.tag_ranges(header_tag)
                if ranges:
                    start_index = ranges[0]
                    text_widget.update_idletasks()
                    pixels = text_widget.tk.call(text_widget._w, 'count', '-ypixels', '1.0', start_index)
                    total_pixels = text_widget.tk.call(text_widget._w, 'count', '-ypixels', '1.0', 'end')
                    frac = float(pixels) / float(total_pixels) if total_pixels != '0' else 0.0
                    text_widget.yview_moveto(frac)
        toc_listbox.bind('<<ListboxSelect>>', on_toc_select)

        # Add context menu
        self._add_context_menu(text_widget)

        # Make text read-only
        text_widget.config(state='disabled')

    def _create_text_widget(self, parent):
        """Create and configure the text widget with scrollbar."""
        scrollbar = ttk.Scrollbar(parent)
        scrollbar.pack(side='right', fill='y')
        
        text_widget = tk.Text(parent,
                              wrap=tk.WORD,
                              yscrollcommand=scrollbar.set,
                              font=("Helvetica", 10),
                              padx=10,
                              pady=10,
                              spacing1=2,
                              spacing2=2,
                              background='#f5f5f5',
                              border=1,
                              state='normal',
                              cursor='xterm',
                              selectbackground='#0078d7',
                              selectforeground='white')
        text_widget.pack(side='left', fill='both', expand=True)
        
        scrollbar.config(command=text_widget.yview)
        text_widget.tag_configure('bold', font=("Helvetica", 10, "bold"))
        
        return text_widget, scrollbar

    def _insert_info_text(self, text_widget):
        """Insert instructional text into the text widget with a table of contents."""
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
        
        # Insert content with marks for headers
        headers = []
        for i, (text, tag) in enumerate(info_text):
            if tag == 'bold':
                mark_name = f"header_{i}"
                text_widget.mark_set(mark_name, 'end')
                text_widget.mark_gravity(mark_name, 'left')
                headers.append((text.strip(), mark_name))
            text_widget.insert('end', text, tag)
        
        # Insert TOC at the beginning
        toc_lines = ["Table of Contents\n"] + [f"• {header_text}\n" for header_text, _ in headers] + ["\n═══════════════════════════════\n"]
        toc_text = ''.join(toc_lines)
        text_widget.insert('1.0', toc_text)
        
        # Add toc_header tag to "Table of Contents"
        toc_header_start = '1.0'
        toc_header_end = text_widget.index('1.0 lineend')
        text_widget.tag_add('toc_header', toc_header_start, toc_header_end)
        text_widget.tag_configure('toc_header', font=("Helvetica", 12, "bold"))
        
        # Add link tags to TOC entries
        current_line = text_widget.index('2.0')  # Line after "Table of Contents"
        for header_text, mark_name in headers:
            link_tag = f"toc_link_{mark_name}"
            line_start = current_line
            line_end = text_widget.index(f"{line_start} lineend")
            text_widget.tag_add(link_tag, line_start, line_end)
            text_widget.tag_configure(link_tag, foreground='blue', underline=True)
            text_widget.tag_bind(link_tag, '<Button-1>', lambda e, m=mark_name: text_widget.see(m))
            text_widget.tag_bind(link_tag, '<Enter>', lambda e: text_widget.config(cursor='hand2'))
            text_widget.tag_bind(link_tag, '<Leave>', lambda e: text_widget.config(cursor=''))
            current_line = text_widget.index(f"{line_start} +1l")
        
        # Enable text selection even when disabled
        def enable_text_selection(event=None):
            return 'break'
        
        text_widget.bind('<Key>', enable_text_selection)
        text_widget.bind('<Control-c>', lambda e: text_widget.event_generate('<<Copy>>') or 'break')
        text_widget.bind('<Control-a>', lambda e: (text_widget.tag_add('sel', '1.0', 'end'), 'break'))

    def _add_context_menu(self, text_widget):
        """Add a right-click context menu to the text widget."""
        context_menu = tk.Menu(text_widget, tearoff=0)
        context_menu.add_command(label="Copy", command=lambda: text_widget.event_generate('<<Copy>>'))
        context_menu.add_command(label="Select All", command=lambda: text_widget.tag_add('sel', '1.0', 'end'))
        text_widget.bind("<Button-3>", lambda event: context_menu.tk_popup(event.x_root, event.y_root))

    def show_debug(self):
        # Ensure stdout_original is set only once, ideally in __init__
        if not hasattr(self, 'console_window') or not self.console_window.window.winfo_exists():
            self.console_window = ConsoleWindow(self.root, log_buffer, self.layout_file, self.latest_images, self.latest_area_name)
        else:   
            self.console_window.update_console()
        sys.stdout = self.console_window
        
    def customize_processing(self, area_name_var):
        area_name = area_name_var.get()
        if not area_name:
            messagebox.showerror("Error", "Invalid area name.")
            return
        if area_name not in self.latest_images:
            messagebox.showerror("Error", "No image to process yet. Please generate an image by pressing the hotkey.")
            return
        self.processing_settings.setdefault(area_name, {})
        ImageProcessingWindow(self.root, area_name, self.latest_images, self.processing_settings[area_name], self)
        
    def set_stop_hotkey(self):
        try:
            self.disable_all_hotkeys()
        except Exception as e:
            print(f"Warning: Error disabling hotkeys: {e}")

        self._hotkey_assignment_cancelled = False
        self.setting_hotkey = True

        # Store the previous hotkey and button text
        prev_hotkey = getattr(self, 'stop_hotkey', None)
        prev_button_text = self.stop_hotkey_button.cget('text')

        def finish_hotkey_assignment():
            try:
                self.restore_all_hotkeys()
            except Exception as e:
                print(f"Warning: Error restoring hotkeys: {e}")
            self.stop_speaking()

        def on_key_press(event):
            if self._hotkey_assignment_cancelled or not self.setting_hotkey or event.scan_code == 1:
                return
            if event.scan_code is None:  # Mouse button
                key_name = event.name
            elif event.name in self.MODIFIERS:
                return
            else:
                if event.scan_code in self.numpad_scan_codes:
                    key_name = 'num ' + self.numpad_scan_codes[event.scan_code]
                else:
                    key_name = event.name

            pressed_modifiers = [mod for mod in self.MODIFIERS if keyboard.is_pressed(mod)]
            hotkey = '+'.join(pressed_modifiers + [key_name])

            # Check for duplicate hotkeys
            for area_frame, hotkey_button, _, area_name_var, *rest in self.areas:
                if hasattr(hotkey_button, 'hotkey') and hotkey_button.hotkey == hotkey:
                    show_thinkr_warning(self, area_name_var.get())
                    self._hotkey_assignment_cancelled = True
                    self.setting_hotkey = False
                    # Restore previous hotkey text if any
                    if prev_hotkey:
                        display_name = self.get_display_name(prev_hotkey)
                        self.stop_hotkey_button.config(text=f"Stop Hotkey: [ {display_name} ]")
                    else:
                        self.stop_hotkey_button.config(text="Set Stop Hotkey")
                    finish_hotkey_assignment()
                    return

            # Handle mouse button restrictions
            parts = hotkey.split('+')
            if parts[-1] in ['mouse_left', 'mouse_right'] and not self.allow_mouse_buttons_var.get():
                messagebox.showwarning("Warning", "Left and right mouse buttons cannot be used as hotkeys.\nCheck 'Allow mouse left/right:' to enable them.")
                self._hotkey_assignment_cancelled = True
                self.setting_hotkey = False
                # Restore previous hotkey text if any
                if prev_hotkey:
                    display_name = self.get_display_name(prev_hotkey)
                    self.stop_hotkey_button.config(text=f"Stop Hotkey: [ {display_name} ]")
                else:
                    self.stop_hotkey_button.config(text="Set Stop Hotkey")
                finish_hotkey_assignment()
                return

            if hasattr(self.stop_hotkey_button, 'mock_button'):
                self._cleanup_hooks(self.stop_hotkey_button.mock_button)
            self.stop_hotkey = hotkey
            mock_button = type('MockButton', (), {'hotkey': hotkey, 'is_stop_button': True})
            self.stop_hotkey_button.mock_button = mock_button
            self.setup_hotkey(mock_button, None)
            display_name = self.get_display_name(hotkey)
            self.stop_hotkey_button.config(text=f"Stop Hotkey: [ {display_name} ]")
            print(f"Set Stop hotkey: {hotkey}\n--------------------------")
            self.setting_hotkey = False
            self._hotkey_assignment_cancelled = True
            finish_hotkey_assignment()

        def on_mouse_click(event):
            if (self._hotkey_assignment_cancelled or
                not self.setting_hotkey or
                not isinstance(event, mouse.ButtonEvent) or
                event.event_type != mouse.DOWN):
                return
            mock_event = type('MockEvent', (), {'name': f"mouse_{event.button}", 'scan_code': None})
            on_key_press(mock_event)

        # Clean up previous hooks if any
        if hasattr(self.stop_hotkey_button, 'keyboard_hook_temp'):
            try:
                keyboard.unhook(self.stop_hotkey_button.keyboard_hook_temp)
            except Exception:
                pass
            delattr(self.stop_hotkey_button, 'keyboard_hook_temp')
        if hasattr(self.stop_hotkey_button, 'mouse_hook_temp'):
            try:
                mouse.unhook(self.stop_hotkey_button.mouse_hook_temp)
            except Exception:
                pass
            delattr(self.stop_hotkey_button, 'mouse_hook_temp')

        self.stop_hotkey_button.config(text="Press any key (10s)...")
        self.setting_hotkey = True
        self.stop_hotkey_button.keyboard_hook_temp = keyboard.on_press(on_key_press)
        self.stop_hotkey_button.mouse_hook_temp = mouse.hook(on_mouse_click)
        self.stop_hotkey_button.countdown_remaining = 10
        self.stop_hotkey_button.config(text=f"Press key ({self.stop_hotkey_button.countdown_remaining}s)")

        def update_countdown():
            if not self.setting_hotkey or not hasattr(self.stop_hotkey_button, 'countdown_remaining') or self.stop_hotkey_button.countdown_remaining <= 0:
                return
            self.stop_hotkey_button.countdown_remaining -= 1
            if self.stop_hotkey_button.countdown_remaining > 0:
                self.stop_hotkey_button.config(text=f"Press key ({self.stop_hotkey_button.countdown_remaining}s)")
                self.stop_hotkey_button.countdown_timer = self.root.after(1000, update_countdown)
            else:
                # Time's up, cancel hotkey setting
                if prev_hotkey:
                    display_name = self.get_display_name(prev_hotkey)
                    self.stop_hotkey_button.config(text=f"Stop Hotkey: [ {display_name} ]")
                else:
                    self.stop_hotkey_button.config(text="Set Stop Hotkey")
                self.setting_hotkey = False
                self._hotkey_assignment_cancelled = True
                # Unhook temporary hooks
                if hasattr(self.stop_hotkey_button, 'keyboard_hook_temp'):
                    keyboard.unhook(self.stop_hotkey_button.keyboard_hook_temp)
                    delattr(self.stop_hotkey_button, 'keyboard_hook_temp')
                if hasattr(self.stop_hotkey_button, 'mouse_hook_temp'):
                    mouse.unhook(self.stop_hotkey_button.mouse_hook_temp)
                    delattr(self.stop_hotkey_button, 'mouse_hook_temp')
                finish_hotkey_assignment()

        self.stop_hotkey_button.countdown_timer = self.root.after(1000, update_countdown)

        # Add Escape key binding to cancel hotkey assignment
        def on_escape(event):
            if prev_hotkey:
                display_name = self.get_display_name(prev_hotkey)
                self.stop_hotkey_button.config(text=f"Stop Hotkey: [ {display_name} ]")
            else:
                self.stop_hotkey_button.config(text="Set Stop Hotkey")
            self.setting_hotkey = False
            self._hotkey_assignment_cancelled = True
            # Unhook temporary hooks
            if hasattr(self.stop_hotkey_button, 'keyboard_hook_temp'):
                keyboard.unhook(self.stop_hotkey_button.keyboard_hook_temp)
                delattr(self.stop_hotkey_button, 'keyboard_hook_temp')
            if hasattr(self.stop_hotkey_button, 'mouse_hook_temp'):
                mouse.unhook(self.stop_hotkey_button.mouse_hook_temp)
                delattr(self.stop_hotkey_button, 'mouse_hook_temp')
            finish_hotkey_assignment()
            # Unbind Escape after use
            self.root.unbind('<Escape>')

        self.root.bind('<Escape>', on_escape)

    def validate_numeric_input(self, P: str, is_speed: bool = False) -> bool:
        """Validate input to only allow numbers with different limits for speed and volume."""
        if P == "":  # Allow empty field
            return True
        if not P.isdigit():  # Only digits allowed
            return False
        value = int(P)
        return value >= 0 if is_speed else 0 <= value <= 100

    def add_read_area(self, removable=True, editable_name=True, area_name="Area Name"):
        area_frame = tk.Frame(self.area_frame)
        area_frame.pack(pady=(4, 0), anchor='center')
        area_name_var = tk.StringVar(value=area_name)
        area_name_label = tk.Label(area_frame, textvariable=area_name_var)
        area_name_label.pack(side="left")
        
        # Editable name with right-click prompt (disabled for Auto Read)
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
            area_name_label.bind('<Button-3>', prompt_edit_area_name)

        # Set Area button (None for Auto Read)
        if not removable and area_name == "Auto Read":
            set_area_button = None
        else:
            set_area_button = tk.Button(area_frame, text="Set Area")
            set_area_button.pack(side="left")
        tk.Label(area_frame, text=" ⏐ ").pack(side="left")  # Separator
        if set_area_button is not None:
            set_area_button.config(command=partial(self.set_area, area_frame, area_name_var, set_area_button))

        # Hotkey button (always present)
        hotkey_button = tk.Button(area_frame, text="Set Hotkey")
        hotkey_button.config(command=lambda: self.set_hotkey(hotkey_button, area_frame))
        hotkey_button.pack(side="left")
        hotkey_button.config(command=lambda: self.set_hotkey_if_not_setting(hotkey_button, area_frame))
        tk.Label(area_frame, text=" ⏐ ").pack(side="left")  # Separator

        # Image Processing controls
        customize_button = tk.Button(area_frame, text="Img. Processing...", command=partial(self.customize_processing, area_name_var))
        customize_button.pack(side="left")
        tk.Label(area_frame, text=" Enable:").pack(side="left")
        preprocess_var = tk.BooleanVar()
        preprocess_checkbox = tk.Checkbutton(area_frame, variable=preprocess_var)
        preprocess_checkbox.pack(side="left")
        tk.Label(area_frame, text=" ⏐ ").pack(side="left")  # Separator

        # Voice selection
        voice_var = tk.StringVar(value="Select Voice")
        voice_menu = tk.OptionMenu(area_frame, voice_var, "Select Voice", *[voice.name for voice in self.current_voices])
        voice_menu.pack(side="left")
        tk.Label(area_frame, text=" ⏐ ").pack(side="left")  # Separator

        # Reading speed entry
        speed_var = tk.StringVar(value="100")
        tk.Label(area_frame, text="Reading Speed % :").pack(side="left")
        vcmd = (self.root.register(lambda P: self.validate_numeric_input(P, is_speed=True)), '%P')
        speed_entry = tk.Entry(area_frame, textvariable=speed_var, width=5, validate='all', validatecommand=vcmd)
        speed_entry.pack(side="left")
        tk.Label(area_frame, text=" ⏐ ").pack(side="left")  # Separator
        
        # Add Show History button
        history_button = tk.Button(area_frame, text="Show History", command=lambda: self.show_history(area_name_var.get()))
        history_button.pack(side="left")
        tk.Label(area_frame, text=" ⏐ ").pack(side="left")  # Separator
        
        speed_entry.bind('<Control-v>', lambda e: 'break')
        speed_entry.bind('<Control-V>', lambda e: 'break')
        speed_entry.bind('<Key>', lambda e: self.validate_speed_key(e, speed_var))

        # Conditional buttons (Remove or Save/Checkbox for Auto Read)
        if removable:
            remove_area_button = tk.Button(area_frame, text="Remove Area", command=lambda: self.remove_area(area_frame, area_name_var.get()))
            remove_area_button.pack(side="left")
        else:
            self.interrupt_on_new_scan_var = tk.BooleanVar(value=True)
            stop_read_checkbox = tk.Checkbutton(area_frame, text="Stop read on select", variable=self.interrupt_on_new_scan_var)
            stop_read_checkbox.pack(side="left", padx=(10, 2))
            def save_auto_read_settings():
                import tempfile, os, json
                hotkey = None
                for area in self.areas:
                    if area[0] == area_frame and area[3].get() == "Auto Read":
                        hotkey = getattr(area[1], 'hotkey', None)
                        break
                settings = {
                    'preprocess': preprocess_var.get(),
                    'voice': voice_var.get(),
                    'speed': speed_var.get(),
                    'hotkey': hotkey,
                    'stop_read_on_select': self.interrupt_on_new_scan_var.get(),
                }
                if 'Auto Read' in self.processing_settings:
                    settings['processing'] = self.processing_settings['Auto Read'].copy()
                game_reader_dir = os.path.join(tempfile.gettempdir(), 'GameReader')
                os.makedirs(game_reader_dir, exist_ok=True)
                temp_path = os.path.join(game_reader_dir, 'auto_read_settings.json')
                with open(temp_path, 'w') as f:
                    json.dump(settings, f, indent=4)
                if hasattr(self, 'status_label'):
                    self.status_label.config(text="Auto Read area settings saved")
                    if hasattr(self, '_feedback_timer') and self._feedback_timer:
                        self.root.after_cancel(self._feedback_timer)
                    self._feedback_timer = self.root.after(2000, lambda: self.status_label.config(text=""))
            save_button = tk.Button(area_frame, text="Save", command=save_auto_read_settings)
            save_button.pack(side="left")

        self.areas.append((area_frame, hotkey_button, set_area_button, area_name_var, preprocess_var, voice_var, speed_var, voice_menu))
        print("Added new read area.\n--------------------------")
        
        # Bind resize events to widgets
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

        self.resize_window()

    def remove_area(self, area_frame, area_name):
        """Remove a specified read area, cleaning up its hotkeys and image."""
        # Find the area and clean up its hotkey
        for area in self.areas:
            if area[0] == area_frame:
                hotkey_button = area[1]
                self._cleanup_hooks(hotkey_button)  # Use existing method to clean up hooks
                break

        # Clean up the associated image
        try:
            if area_name in self.latest_images:
                self.latest_images[area_name].close()
                del self.latest_images[area_name]
        except KeyError:
            print(f"Warning: No image found for area '{area_name}'")
        except Exception as e:
            print(f"Warning: Error closing image for area '{area_name}': {e}")

        # Remove the area from UI and internal list
        area_frame.destroy()
        self.areas = [area for area in self.areas if area[0] != area_frame]
        
        # Mark unsaved changes
        self._set_unsaved_changes()
        
        print(f"Removed area: {area_name}\n--------------------------")

    def resize_window(self):
        """Resize the window based on the number of areas and the widest area line. Caps height after 10 areas, enabling scrollbar."""
        BASE_HEIGHT = 210  # Height for main controls and padding
        MIN_WIDTH = 950
        MAX_WIDTH = 1600
        AREA_ROW_HEIGHT = 60  # 55 for frame, 5 for padding
        SCROLLBAR_THRESHOLD = 8  # Show scrollbar after 8 areas

        # Calculate total area frame height
        area_frame_height = self.area_frame.winfo_height() if self.areas else 0
        self.area_frame.update_idletasks()

        # Cap visible areas at 10 for height calculation
        visible_area_count = min(10, len(self.areas))
        fixed_canvas_height = visible_area_count * AREA_ROW_HEIGHT
        total_height = min(BASE_HEIGHT + (fixed_canvas_height if len(self.areas) > 10 else area_frame_height), BASE_HEIGHT + 10 * AREA_ROW_HEIGHT)
        total_height = max(total_height, 250)

        # Find the widest area frame
        widest = MIN_WIDTH
        for area in self.areas:
            frame = area[0]
            frame.update_idletasks()
            frame_left = frame.winfo_rootx()
            farthest_right = max((child.winfo_rootx() + child.winfo_width() for child in frame.winfo_children()), default=frame_left)
            widest = max(widest, farthest_right - frame_left + 60)

        window_width = max(MIN_WIDTH, min(MAX_WIDTH, widest))

        # Set minimum window size, capping height
        self.root.minsize(window_width, min(total_height, BASE_HEIGHT + 5 * AREA_ROW_HEIGHT))

        # Manage scrollbar visibility
        if hasattr(self, 'area_scrollbar'):
            canvas_height = fixed_canvas_height if len(self.areas) > SCROLLBAR_THRESHOLD else area_frame_height
            if len(self.areas) > SCROLLBAR_THRESHOLD:
                self.area_scrollbar.pack(side='right', fill='y')
                self.area_canvas.configure(yscrollcommand=self.area_scrollbar.set)
            else:
                self.area_scrollbar.pack_forget()
            self.area_canvas.config(height=canvas_height)

        self.root.update_idletasks()  # Ensure geometry is applied

    def set_area(self, frame, area_name_var, set_area_button):
        # Store current mouse hooks
        self.saved_mouse_hooks = getattr(self, 'mouse_hooks', []).copy()

        # Disable all hotkeys
        try:
            keyboard.unhook_all()
            if hasattr(self, 'mouse_hooks'):
                self.mouse_hooks.clear()
        except Exception as e:
            print(f"Error disabling hotkeys for area selection: {e}")

        self.hotkeys_disabled_for_selection = True

        # Create fullscreen selection window
        select_area_window = tk.Toplevel(self.root)
        select_area_window.overrideredirect(True)

        # Calculate virtual screen dimensions across all monitors
        monitors = win32api.EnumDisplayMonitors()
        min_x = min(monitor[2][0] for monitor in monitors)
        min_y = min(monitor[2][1] for monitor in monitors)
        max_x = max(monitor[2][2] for monitor in monitors)
        max_y = max(monitor[2][3] for monitor in monitors)
        virtual_width = max_x - min_x
        virtual_height = max_y - min_y

        select_area_window.geometry(f"{virtual_width}x{virtual_height}+{min_x}+{min_y}")

        # Set up canvas
        canvas = tk.Canvas(select_area_window, cursor="cross", width=virtual_width, height=virtual_height,
                           highlightthickness=0, bg='white')
        canvas.pack(fill="both", expand=True)

        select_area_window.attributes("-alpha", 0.5)
        select_area_window.attributes("-topmost", True)

        # Add user instructions
        tk.Label(select_area_window, text="Drag to select area, press Escape to cancel",
                 font=("Helvetica", 12), bg='white').place(x=10, y=10)

        # Create selection rectangles
        border = canvas.create_rectangle(0, 0, 0, 0, outline='red', width=3, dash=(8, 4))
        border_outline = canvas.create_rectangle(0, 0, 0, 0, outline='red', width=3, dash=(8, 4), dashoffset=6)

        # Initialize coordinates in screen space
        x1, y1 = 0, 0

        def on_click(event):
            nonlocal x1, y1
            x1 = event.x_root
            y1 = event.y_root
            canvas.bind("<B1-Motion>", on_drag)
            canvas.bind("<ButtonRelease-1>", on_release)
            canvas_x = x1 - min_x
            canvas_y = y1 - min_y
            canvas.coords(border, canvas_x, canvas_y, canvas_x, canvas_y)
            canvas.coords(border_outline, canvas_x, canvas_y, canvas_x, canvas_y)

        def on_drag(event):
            current_x = event.x_root
            current_y = event.y_root
            canvas_x1 = x1 - min_x
            canvas_y1 = y1 - min_y
            canvas_x2 = current_x - min_x
            canvas_y2 = current_y - min_y
            coords = (min(canvas_x1, canvas_x2), min(canvas_y1, canvas_y2),
                      max(canvas_x1, canvas_x2), max(canvas_y1, canvas_y2))
            canvas.coords(border, *coords)
            canvas.coords(border_outline, *coords)

        def on_release(event):
            try:
                if hasattr(self, 'interrupt_on_new_scan_var') and self.interrupt_on_new_scan_var.get():
                    self.stop_speaking()
                x2 = event.x_root
                y2 = event.y_root
                if abs(x2 - x1) > 5 and abs(y2 - y1) > 5:  # Minimum 5px drag
                    frame.area_coords = (min(x1, x2), min(y1, y2), max(x1, x2), max(y1, y2))
                else:
                    frame.area_coords = getattr(frame, 'area_coords', (0, 0, 0, 0))

                is_auto_read = area_name_var.get() == "Auto Read"
                select_area_window.destroy()

                if is_auto_read:
                    self.root.after(100, lambda: self.read_area(frame))
                else:
                    current_name = area_name_var.get()
                    if current_name == "Area Name":
                        area_name = simpledialog.askstring("Area Name", "Enter a name for this area:")
                        if area_name and area_name.strip():
                            area_name_var.set(area_name)
                            print(f"Set area: {frame.area_coords} with name {area_name_var.get()}\n--------------------------")
                        else:
                            messagebox.showerror("Error", "Area name cannot be empty.")
                            print("Error: Area name cannot be empty.")
                            self._restore_hotkeys_after_selection()
                            return
                    else:
                        print(f"Updated area: {frame.area_coords} with existing name {current_name}\n--------------------------")
                    if set_area_button is not None:
                        set_area_button.config(text="Edit Area")
                if set_area_button is not None and is_auto_read:
                    set_area_button.config(text="Select Area")

                self._set_unsaved_changes()
            except Exception as e:
                print(f"Error during area selection: {e}")
            finally:
                self._restore_hotkeys_after_selection()

        def on_escape(event):
            canvas.unbind("<B1-Motion>")
            canvas.unbind("<ButtonRelease-1>")
            select_area_window.destroy()
            self._restore_hotkeys_after_selection()
            print("Area selection cancelled\n--------------------------")

        # Bind events
        canvas.bind("<Button-1>", on_click)
        canvas.bind("<Escape>", on_escape)
        select_area_window.bind("<Escape>", on_escape)
        select_area_window.focus_force()
        select_area_window.bind("<FocusOut>", lambda e: select_area_window.focus_force())
        select_area_window.bind("<Key>", lambda e: on_escape(e) if e.keysym == "Escape" else None)

    def _restore_hotkeys_after_selection(self):
        """Helper method to restore hotkeys after area selection"""
        # Check if hotkeys were disabled; return early if not
        if not getattr(self, 'hotkeys_disabled_for_selection', False):
            return

        try:
            self.restore_all_hotkeys()
            print("Hotkeys re-enabled after area selection")
        except (AttributeError, RuntimeError) as e:
            print(f"Error restoring hotkeys: {e}")
        finally:
            # Always reset the flag, even on error
            self.hotkeys_disabled_for_selection = False
            # Restore focus if root exists
            if hasattr(self, 'root') and self.root.winfo_exists():
                self.root.focus_force()

    def disable_all_hotkeys(self):
        """Disable all hotkeys for keyboard and mouse."""
        try:
            keyboard.unhook_all()
            mouse.unhook_all()
        except Exception as e:
            print(f"Warning: Error during hotkey cleanup: {e}")
        self.keyboard_hooks.clear()
        self.mouse_hooks.clear()
        self.hotkeys.clear()
        self.setting_hotkey = False

    def unhook_mouse(self):
        if hasattr(self, 'mouse_hooks') and self.mouse_hooks:
            try:
                mouse.unhook_all()
                self.mouse_hooks.clear()
            except Exception as e:
                print(f"Warning: Error during mouse hook cleanup: {e}")
                print(f"Mouse hooks list state: {len(self.mouse_hooks)}")

    def restore_all_hotkeys(self):
        """Restore all area and stop hotkeys after area selection is finished/cancelled."""
        try:
            keyboard.unhook_all()
            if hasattr(self, 'mouse_hooks'):
                self.mouse_hooks.clear()
        except Exception as e:
            print(f"Error cleaning up hooks during restore: {e}")
        
        if hasattr(self, 'saved_mouse_hooks'):
            if not hasattr(self, 'mouse_hooks'):
                self.mouse_hooks = []
            for hook in self.saved_mouse_hooks:
                try:
                    mouse.hook(hook)
                    self.mouse_hooks.append(hook)
                except Exception as e:
                    print(f"Error restoring mouse hook: {e}")
            delattr(self, 'saved_mouse_hooks')
        
        for area in self.areas:
            area_frame, hotkey_button, *rest = area
            if hasattr(hotkey_button, 'hotkey'):
                try:
                    self.setup_hotkey(hotkey_button, area_frame)
                except Exception as e:
                    print(f"Error re-registering hotkey: {e}")
        
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
                delattr(button, 'mouse_hook_temp')
            self.disable_all_hotkeys()
        except Exception as e:
            print(f"Warning: Error cleaning up temporary hooks: {e}")

        self._hotkey_assignment_cancelled = False  # Guard flag to block late events
        self.setting_hotkey = True

        def finish_hotkey_assignment():
            """Re-enable all hotkeys after assignment is finished or cancelled."""
            try:
                self.restore_all_hotkeys()
            except Exception as e:
                print(f"Warning: Error restoring hotkeys: {e}")

        def on_key_press(event):
            """Handle key press events during hotkey assignment."""
            if self._hotkey_assignment_cancelled or not self.setting_hotkey:
                return
            if event.scan_code == 1:  # Ignore Escape
                return

            if event.scan_code is None:  # Mouse button
                key_name = event.name  # 'button1', etc.
            elif event.name in self.MODIFIERS:
                return  # Ignore modifier keys alone
            else:
                if event.scan_code in self.numpad_scan_codes:
                    key_name = 'num ' + self.numpad_scan_codes[event.scan_code]
                else:
                    key_name = event.name

            pressed_modifiers = [mod for mod in self.MODIFIERS if keyboard.is_pressed(mod)]
            hotkey = '+'.join(pressed_modifiers + [key_name])

            # Check for duplicate hotkeys
            for area in self.areas:
                if area[1] is not button and hasattr(area[1], 'hotkey') and area[1].hotkey == hotkey:
                    self.setting_hotkey = False
                    self._hotkey_assignment_cancelled = True
                    if hasattr(button, 'keyboard_hook_temp'):
                        keyboard.unhook(button.keyboard_hook_temp)
                        delattr(button, 'keyboard_hook_temp')
                    if hasattr(button, 'mouse_hook_temp'):
                        mouse.unhook(button.mouse_hook_temp)
                        delattr(button, 'mouse_hook_temp')
                    finish_hotkey_assignment()
                    if hasattr(button, 'hotkey'):
                        display_name = self.get_display_name(button.hotkey)
                        button.config(text=f"Set Hotkey: [ {display_name} ]")
                    else:
                        button.config(text="Set Hotkey")
                    area_name = area[3].get() if hasattr(area[3], 'get') else "Unknown Area"
                    show_thinkr_warning(self, area_name)
                    return

            # Handle mouse button restrictions
            parts = hotkey.split('+')
            if parts[-1] in ['mouse_left', 'mouse_right'] and not self.allow_mouse_buttons_var.get():
                messagebox.showwarning("Warning", "Left and right mouse buttons cannot be used as hotkeys.\nCheck 'Allow mouse left/right:' to enable them.")
                self.setting_hotkey = False
                self._hotkey_assignment_cancelled = True
                finish_hotkey_assignment()
                button.config(text="Set Hotkey")
                return

            # Set the hotkey and finalize assignment
            button.hotkey = hotkey
            display_name = self.get_display_name(hotkey)
            button.config(text=f"Set Hotkey: [ {display_name} ]")
            self.setup_hotkey(button, area_frame)
            # Stop countdown
            if hasattr(button, 'countdown_timer'):
                self.root.after_cancel(button.countdown_timer)
                delattr(button, 'countdown_timer')
            if hasattr(button, 'countdown_remaining'):
                delattr(button, 'countdown_remaining')
            self.setting_hotkey = False
            if hasattr(button, 'keyboard_hook_temp'):
                keyboard.unhook(button.keyboard_hook_temp)
                delattr(button, 'keyboard_hook_temp')
            if hasattr(button, 'mouse_hook_temp'):
                mouse.unhook(button.mouse_hook_temp)
                delattr(button, 'mouse_hook_temp')
            finish_hotkey_assignment()

        def on_mouse_click(event):
            """Handle mouse click events during hotkey assignment."""
            if not self.setting_hotkey or not isinstance(event, mouse.ButtonEvent) or event.event_type != mouse.DOWN:
                return
            mock_event = type('MockEvent', (), {'name': f"mouse_{event.button}", 'scan_code': None})
            on_key_press(mock_event)

        # Clean up previous hooks
        if hasattr(button, 'keyboard_hook'):
            try:
                keyboard.unhook(button.keyboard_hook)
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
        self.setting_hotkey = True
        button.keyboard_hook_temp = keyboard.on_press(on_key_press)
        button.mouse_hook_temp = mouse.hook(on_mouse_click)
        button.countdown_remaining = 10
        button.config(text=f"Press key ({button.countdown_remaining}s)")

        # Set 3-second timeout for hotkey setting
        def update_countdown():
            if not self.setting_hotkey or not hasattr(button, 'countdown_remaining') or button.countdown_remaining <= 0:
                return
            button.countdown_remaining -= 1
            if button.countdown_remaining > 0:
                button.config(text=f"Press key ({button.countdown_remaining}s)")
                button.countdown_timer = self.root.after(1000, update_countdown)
            else:
                # Time's up, cancel hotkey setting
                button.config(text="Set Hotkey")
                self.setting_hotkey = False
                self._hotkey_assignment_cancelled = True
                # Unhook temporary hooks
                if hasattr(button, 'keyboard_hook_temp'):
                    keyboard.unhook(button.keyboard_hook_temp)
                    delattr(button, 'keyboard_hook_temp')
                if hasattr(button, 'mouse_hook_temp'):
                    mouse.unhook(button.mouse_hook_temp)
                    delattr(button, 'mouse_hook_temp')
                # Restore all hotkeys
                finish_hotkey_assignment()

        button.countdown_timer = self.root.after(1000, update_countdown)

        # Add Escape key binding to cancel hotkey assignment for area hotkeys
        def on_escape(event):
            if hasattr(button, 'hotkey') and button.hotkey:
                display_name = self.get_display_name(button.hotkey)
                button.config(text=f"Set Hotkey: [ {display_name} ]")
            else:
                button.config(text="Set Hotkey")
            self.setting_hotkey = False
            self._hotkey_assignment_cancelled = True
            # Unhook temporary hooks
            if hasattr(button, 'keyboard_hook_temp'):
                keyboard.unhook(button.keyboard_hook_temp)
                delattr(button, 'keyboard_hook_temp')
            if hasattr(button, 'mouse_hook_temp'):
                mouse.unhook(button.mouse_hook_temp)
                delattr(button, 'mouse_hook_temp')
            finish_hotkey_assignment()
            # Unbind Escape after use
            self.root.unbind('<Escape>')

        self.root.bind('<Escape>', on_escape)

    def _update_status(self, message, duration=2000):
        """Update the status label with a message and clear it after a duration."""
        if hasattr(self, 'status_label'):
            # Cancel any existing timer
            if hasattr(self, '_feedback_timer') and self._feedback_timer:
                self.root.after_cancel(self._feedback_timer)
            self.status_label.config(text=message)
            if duration > 0:
                self._feedback_timer = self.root.after(duration, lambda: self.status_label.config(text=""))

    def save_layout(self):
        # Early return if no areas exist
        if not self.areas:
            messagebox.showerror("Error", "There is nothing to save.")
            return

        # Reset unsaved changes flag
        self._has_unsaved_changes = False

        # Validate coordinates for all areas except "Auto Read"
        for area_frame, _, _, area_name_var, _, _, _ in self.areas:
            if area_name_var.get() != "Auto Read" and not hasattr(area_frame, 'area_coords'):
                messagebox.showerror("Error", f"Area '{area_name_var.get()}' does not have a defined area, remove it or configure before saving.")
                return

        # Build layout dictionary with settings
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

        # Populate areas list, excluding "Auto Read"
        for area_frame, hotkey_button, _, area_name_var, preprocess_var, voice_var, speed_var, _ in self.areas:
            if area_name_var.get() == "Auto Read":
                continue
            if hasattr(area_frame, 'area_coords'):
                layout["areas"].append({
                    "coords": area_frame.area_coords,
                    "name": area_name_var.get(),
                    "hotkey": getattr(hotkey_button, 'hotkey', None),
                    "preprocess": preprocess_var.get(),
                    "voice": voice_var.get(),
                    "speed": speed_var.get(),
                    "settings": self.processing_settings.get(area_name_var.get(), {})
                })

        # Prepare save dialog parameters
        current_file = self.layout_file.get()
        initial_dir = os.path.dirname(current_file) if current_file else os.getcwd()
        initial_file = os.path.basename(current_file) if current_file else ""

        # Prompt user for file path
        file_path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json")],
            initialdir=initial_dir,
            initialfile=initial_file
        )
        if not file_path:
            return

        # Save layout to file with specific error handling
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(layout, f, indent=4)
            self.layout_file.set(file_path)
            self._update_status(f"Layout saved to: {os.path.basename(file_path)}")
            print(f"Layout saved to {file_path}\n--------------------------")
        except PermissionError:
            messagebox.showerror("Error", "Permission denied. Please choose a different location or run as administrator.")
            print(f"Permission error saving layout to {file_path}")
        except IOError as e:
            messagebox.showerror("Error", f"Failed to save layout: {str(e)}")
            print(f"IO error saving layout: {e}")
        except Exception as e:
            messagebox.showerror("Error", f"Unexpected error saving layout: {str(e)}")
            print(f"Unexpected error saving layout: {e}")

    def load_game_units(self):
        """Load game units from JSON file in GameReader directory."""
        import tempfile, os, json, re
        temp_path = os.path.join(tempfile.gettempdir(), 'GameReader')
        os.makedirs(temp_path, exist_ok=True)
        
        file_path = os.path.join(temp_path, 'gamer_units.json')
        if os.path.exists(file_path):
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                content = remove_json_comments(content)
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
        temp_path = os.path.join(tempfile.gettempdir(), 'GameReader')
        os.makedirs(temp_path, exist_ok=True)
        
        file_path = os.path.join(temp_path, 'game_units.json')
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                header = [
                    '//  Game Units Configuration',
                    '//  Format: "short_name": "Full Name"',
                    '//  Example: "xp" will be read as "Experience Points"',
                    '//  Enable "Read gamer units" in the main window to use this feature',
                    ''
                ]
                f.write('\n'.join(header))
                json.dump(self.game_units, f, indent=4, ensure_ascii=False)
            print(f"Game units saved to: {file_path}")
            
            # Update status label with feedback
            if hasattr(self, '_feedback_timer') and self._feedback_timer:
                self.root.after_cancel(self._feedback_timer)
            self.status_label.config(text="Game units saved successfully!")
            self._feedback_timer = self.root.after(3000, lambda: self.status_label.config(text=""))
            
            return True
        except Exception as e:
            print(f"Error saving game units: {e}")
            return False

    def open_game_reader_folder(self):
        """Open the GameReader folder in Windows Explorer."""
        folder_path = os.path.join(tempfile.gettempdir(), 'GameReader')
        os.makedirs(folder_path, exist_ok=True)
        try:
            subprocess.Popen(f'explorer "{folder_path}"')
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open folder: {str(e)}")

    def normalize_text(self, text):
        """Normalize text by removing punctuation and making it lowercase."""
        text = text.lower()
        text = re.sub(r'[^\w\s]', '', text)  # Remove punctuation
        text = re.sub(r'\s+', ' ', text).strip()  # Remove extra whitespace
        return text

    def on_drop(self, event):
        """Handle file drop event"""
        try:
            file_path = self._normalize_path(event.data)
            if not self._is_valid_json_file(file_path):
                messagebox.showerror("Error", "Please drop a valid JSON layout file")
                return

            current_layout = self.layout_file.get()
            if current_layout:
                if os.path.normpath(current_layout) == file_path:
                    return
                if self._has_unsaved_changes:
                    response = messagebox.askyesnocancel(
                        "Unsaved Changes",
                        f"You have unsaved changes in the current layout.\n\n"
                        f"Current: {os.path.basename(current_layout)}\n"
                        f"New: {os.path.basename(file_path)}\n\n"
                        "Do you want to save changes before loading the new layout?\n"
                        "(Yes = Save and load, No = Discard changes and load, Cancel = Do nothing)"
                    )
                    if response is None:
                        return
                    elif response:
                        self.save_layout()
                else:
                    if not messagebox.askyesno(
                        "Load New Layout",
                        f"Load new layout file?\n\n"
                        f"Current: {os.path.basename(current_layout)}\n"
                        f"New: {os.path.basename(file_path)}"
                    ):
                        return

            self._load_layout_file(file_path)
        except Exception as e:
            messagebox.showerror("Error", f"Error handling dropped file: {str(e)}")
            import traceback
            traceback.print_exc()

    def _normalize_path(self, path):
        """Normalize the file path"""
        path = path.strip('{}').strip('\"\'')
        return os.path.normpath(path)

    def _is_valid_json_file(self, file_path):
        """Check if the file is a valid JSON file"""
        return os.path.isfile(file_path) and file_path.lower().endswith('.json')

    def _show_status_feedback(self, message):
        """Display temporary status feedback"""
        if hasattr(self, '_feedback_timer') and self._feedback_timer:
            self.root.after_cancel(self._feedback_timer)
        self.status_label.config(text=message)
        self._feedback_timer = self.root.after(2000, lambda: self.status_label.config(text=""))

    def _cleanup_resources(self):
        """Clean up images and hotkeys"""
        for image in self.latest_images.values():
            try:
                image.close()
            except:
                pass
        self.latest_images.clear()
        keyboard.unhook_all()
        mouse.unhook_all()

    def _clear_auto_read_hotkey(self):
        """Clear existing Auto Read hotkey and return its value"""
        auto_read_hotkey = None
        if self.areas and hasattr(self.areas[0][1], 'hotkey'):
            auto_read_hotkey = self.areas[0][1].hotkey
            if auto_read_hotkey:
                try:
                    keyboard.remove_hotkey(self.areas[0][1].hotkey_id)
                except (KeyError, AttributeError):
                    pass
                self.areas[0][1].hotkey = None
                self.areas[0][1].config(text="Set Hotkey")
        return auto_read_hotkey

    def _check_hotkey_conflict(self, auto_read_hotkey, areas):
        """Check for hotkey conflicts with Auto Read"""
        for area_info in areas:
            if auto_read_hotkey and area_info.get("hotkey") == auto_read_hotkey:
                return area_info["name"]
        return None

    def _load_area(self, area_info):
        """Load a single area from layout data"""
        self.add_read_area(removable=True, editable_name=True, area_name=area_info["name"])
        area_frame, hotkey_button, set_area_button, area_name_var, preprocess_var, voice_var, speed_var, _ = self.areas[-1]
        
        area_frame.area_coords = area_info["coords"]
        if area_info["hotkey"]:
            hotkey_button.hotkey = area_info["hotkey"]
            display_name = area_info["hotkey"].replace('num_', 'num:') if area_info["hotkey"].startswith('num_') else area_info["hotkey"]
            hotkey_button.config(text=f"Hotkey: [ {display_name} ]")
            self.setup_hotkey(hotkey_button, area_frame)
        
        preprocess_var.set(area_info.get("preprocess", False))
        if area_info.get("voice") in [voice.name for voice in self.voices]:
            voice_var.set(area_info["voice"])
        speed_var.set(area_info.get("speed", "1.0"))
        
        if "settings" in area_info:
            self.processing_settings[area_info["name"]] = area_info["settings"].copy()
            print(f"Loaded image processing settings for area: {area_info['name']}")
        
        x1, y1, x2, y2 = area_frame.area_coords
        screenshot = capture_screen_area(x1, y1, x2, y2)
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

    def _handle_auto_read_hotkey(self, auto_read_hotkey, conflict_area_name):
        """Handle Auto Read hotkey post-loading"""
        if not conflict_area_name and auto_read_hotkey and self.areas and hasattr(self.areas[0][1], 'hotkey'):
            try:
                self.areas[0][1].hotkey = auto_read_hotkey
                display_name = auto_read_hotkey.replace('num_', 'num:') if auto_read_hotkey.startswith('num_') else auto_read_hotkey
                self.areas[0][1].config(text=f"Hotkey: [ {display_name} ]")
                self.setup_hotkey(self.areas[0][1], self.areas[0][0])
                print(f"Re-registered Auto Read hotkey: {auto_read_hotkey}")
            except Exception as e:
                print(f"Error re-registering Auto Read hotkey: {e}")
        elif conflict_area_name:
            hotkey_val = auto_read_hotkey if auto_read_hotkey else "?"
            messagebox.showinfo(
                "Hotkey Conflict",
                f"Detected same Hotkey!\n\nAuto Read Hotkey = {hotkey_val}\n{conflict_area_name} Hotkey = {hotkey_val}\n\nPlease set a new hotkey for AutoRead if you still want this function."
            )
            if self.areas and hasattr(self.areas[0][1], 'hotkey'):
                try:
                    if hasattr(self.areas[0][1], 'hotkey_id'):
                        keyboard.remove_hotkey(self.areas[0][1].hotkey_id)
                except (KeyError, AttributeError):
                    pass
                self.areas[0][1].hotkey = None
                self.areas[0][1].config(text="Set Hotkey")

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
        try:
            # Store the full file path and reset unsaved changes flag
            self.layout_file.set(file_path)
            self._has_unsaved_changes = False

            # Load JSON data from file
            with open(file_path, 'r') as f:
                layout = json.load(f)

            # Preserve permanent area (first area) and clear others
            if self.areas:
                permanent_area = self.areas[0]
                for area in self.areas[1:]:
                    area[0].destroy()
                self.areas = [permanent_area]
            self.processing_settings.clear()

            # Check version compatibility
            save_version = layout.get("version", "0.0")
            current_version = "0.5"
            if tuple(map(int, save_version.split('.'))) < tuple(map(int, current_version.split('.'))):
                messagebox.showerror("Error", "Cannot load older version save files.")
                return

            # Extract filename for display
            file_name = os.path.basename(file_path)
            self._show_status_feedback(f"Layout loaded: {file_name}")

            # Load layout settings
            self.layout_file.set(file_name)
            self.bad_word_list.set(layout.get("bad_word_list", ""))
            self.ignore_usernames_var.set(layout.get("ignore_usernames", False))
            self.ignore_previous_var.set(layout.get("ignore_previous", False))
            self.ignore_gibberish_var.set(layout.get("ignore_gibberish", False))
            self.pause_at_punctuation_var.set(layout.get("pause_at_punctuation", False))
            self.better_unit_detection_var.set(layout.get("better_unit_detection", False))
            self.read_game_units_var.set(layout.get("read_game_units", False))
            self.fullscreen_mode_var.set(layout.get("fullscreen_mode", False))

            # Load and apply volume setting
            saved_volume = layout.get("volume", "100")
            self.volume.set(saved_volume)
            try:
                self.speaker.Volume = int(saved_volume)
                print(f"Loaded volume setting: {saved_volume}%")
            except ValueError:
                print("Invalid volume in save file, defaulting to 100%")
                self.volume.set("100")
                self.speaker.Volume = 100

            # Clean up existing resources
            self._cleanup_resources()

            # Load stop hotkey
            saved_stop_hotkey = layout.get("stop_hotkey")
            if saved_stop_hotkey:
                self.stop_hotkey = saved_stop_hotkey
                self.stop_hotkey_button.mock_button = type('MockButton', (), {
                    'hotkey': saved_stop_hotkey,
                    'is_stop_button': True
                })
                self.setup_hotkey(self.stop_hotkey_button.mock_button, None)
                display_name = saved_stop_hotkey.replace('num_', 'num:') if saved_stop_hotkey.startswith('num_') else saved_stop_hotkey
                self.stop_hotkey_button.config(text=f"Stop Hotkey: [ {display_name} ]")
                print(f"Loaded Stop hotkey: {saved_stop_hotkey}")

            # Handle Auto Read hotkey
            auto_read_hotkey = self._clear_auto_read_hotkey()
            conflict_area_name = self._check_hotkey_conflict(auto_read_hotkey, layout.get("areas", []))

            # Load areas from layout
            for area_info in layout.get("areas", []):
                self._load_area(area_info)

            # Post-load Auto Read hotkey handling
            self._handle_auto_read_hotkey(auto_read_hotkey, conflict_area_name)

            print(f"Layout loaded from {file_path}\n--------------------------")
            self.resize_window()  # Resize once after all areas are loaded

        except Exception as e:
            messagebox.showerror("Error", f"Failed to load layout: {str(e)}")
            print(f"Error loading layout: {e}")

    def validate_speed_key(self, event, speed_var):
        """Validate key presses for speed entry, allowing digits and navigation keys."""
        allowed_keys = {'BackSpace', 'Delete', 'Left', 'Right'}
        return None if (event.char.isdigit() or event.keysym in allowed_keys) else 'break'

    def setup_hotkey(self, button, area_frame):
        try:
            self._cleanup_hooks(button)
            if not hasattr(button, 'is_stop_button') and area_frame is not None:
                button.area_frame = area_frame
            if not hasattr(button, 'hotkey') or not button.hotkey:
                print(f"No hotkey set for button: {button}")
                return False
            print(f"Setting up hotkey for: {button.hotkey}")

            parts = button.hotkey.split('+')
            if parts[-1].startswith('mouse_'):
                # Mouse hotkey
                modifiers = parts[:-1]
                mouse_button = parts[-1][6:]  # 'button1', 'button2', etc.

                def on_mouse_event(event):
                    if isinstance(event, mouse.ButtonEvent) and event.event_type == mouse.DOWN:
                        pressed_modifiers = [mod for mod in self.ALL_MODIFIERS if keyboard.is_pressed(mod)]
                        if event.button == mouse_button and pressed_modifiers == modifiers:
                            handle_hotkey_action()

                button.mouse_hook = mouse.hook(on_mouse_event)
                print(f"Mouse hook set up for {button.hotkey}")
            else:
                # Keyboard hotkey
                hotkey_str = button.hotkey
                def callback():
                    handle_hotkey_action()
                button.keyboard_hotkey_id = keyboard.add_hotkey(hotkey_str, callback)
                print(f"Keyboard hotkey set up for {hotkey_str}")

            def handle_hotkey_action():
                if hasattr(button, 'is_stop_button'):
                    self.root.after_idle(self.stop_speaking)
                    return True
                area_info = self._get_area_info(button)
                if area_info and area_info.get('name') == "Auto Read":
                    self.root.after_idle(lambda: self.set_area(
                        area_info['frame'],
                        area_info['name_var'],
                        area_info['set_area_btn']))
                    return True
                if getattr(button, '_is_processing', False):
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

            return True
        except Exception as e:
            print(f"Error in setup_hotkey: {e}")
            return False

    def _cleanup_hooks(self, button):
        try:
            if hasattr(button, 'mouse_hook'):
                try:
                    if button.mouse_hook in mouse._listener.handlers:
                        mouse.unhook(button.mouse_hook)
                except Exception as e:
                    print(f"Warning: Error cleaning up mouse hook: {e}")
                finally:
                    delattr(button, 'mouse_hook')
            if hasattr(button, 'keyboard_hotkey_id'):
                try:
                    keyboard.remove_hotkey(button.keyboard_hotkey_id)
                except Exception as e:
                    print(f"Warning: Error removing hotkey: {e}")
                finally:
                    delattr(button, 'keyboard_hotkey_id')
        except Exception as e:
            print(f"Unexpected error in _cleanup_hooks: {e}")

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
        # Handle stop button
        if area_frame is None:
            return

        # Validate area coordinates
        if not hasattr(area_frame, 'area_coords'):
            area_info = next((area for area in self.areas if area[0] is area_frame), None)
            if area_info and area_info[3].get() == "Auto Read":
                return
            messagebox.showerror("Error", "No area coordinates set. Click Set Area to set one.")
            return

        # Initialize speaker if necessary
        if not self.speaker:
            try:
                self.speaker = win32com.client.Dispatch("SAPI.SpVoice")
                self.speaker.Volume = int(self.volume.get())
            except Exception as e:
                print(f"Error initializing speaker: {e}")
                return

        # Retrieve area information efficiently
        area_info = next((area for area in self.areas if area[0] is area_frame), None)
        if not area_info:
            print(f"Error: Could not determine area name for frame {area_frame}")
            return
        _, _, _, area_name_var, preprocess_var, voice_var, speed_var, *rest = area_info
        area_name = area_name_var.get()
        self.latest_area_name.set(area_name)
        preprocess = preprocess_var.get()

        # Provide processing feedback
        self.show_processing_feedback(area_name)

        # Capture and process screenshot
        x1, y1, x2, y2 = area_frame.area_coords
        screenshot = capture_screen_area(x1, y1, x2, y2)
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
            text = pytesseract.image_to_string(processed_image)
            print("Image preprocessing applied.")
        else:
            self.latest_images[area_name] = screenshot
            text = pytesseract.image_to_string(screenshot)

        # Apply unit detection
        if self.better_unit_detection_var.get():
            unit_map = {
                'l': 'Liters', 'm': 'Meters', 'in': 'Inches', 'ml': 'Milliliters', 'gal': 'Gallons',
                'g': 'Grams', 'lb': 'Pounds', 'ib': 'Pounds', 'c': 'Celsius', 'f': 'Fahrenheit',
                'kr': 'Crowns', 'eur': 'Euros', 'usd': 'US Dollars', 'sek': 'Swedish Crowns',
                'nok': 'Norwegian Crowns', 'dkk': 'Danish Crowns', '£': 'Pounds Sterling',
            }
            pattern = re.compile(r'(?<!\w)(\d+(?:\.\d+)?)(\s*)(l|m|in|ml|gal|g|lb|ib|c|f|kr|eur|usd|sek|nok|dkk|£)(?!\w)', re.IGNORECASE)
            def repl(match):
                value, space, unit = match.group(1), match.group(2), match.group(3).lower()
                if unit in ['lb', 'ib']:
                    return f"{value}{space}Pounds"
                if unit == '£':
                    return f"{value}{space}Pounds Sterling"
                return f"{value}{space}{unit_map.get(unit, unit)}"
            text = pattern.sub(repl, text)

        if self.read_game_units_var.get():
            game_unit_map = self.game_units.copy()
            default_mappings = {
                'xp': 'Experience Points', 'hp': 'Health Points', 'mp': 'Mana Points', 'gp': 'Gold Pieces',
                'pp': 'Platinum Pieces', 'sp': 'Skill Points', 'ep': 'Energy Points', 'ap': 'Action Points',
                'bp': 'Battle Points', 'lp': 'Loyalty Points', 'cp': 'Challenge Points', 'vp': 'Victory Points',
                'rp': 'Reputation Points', 'tp': 'Talent Points', 'ar': 'Armor Rating', 'dmg': 'Damage',
                'dps': 'Damage Per Second', 'def': 'Defense', 'mat': 'Materials', 'exp': 'Exploration Points',
                '§': 'Simoliance', 'v-bucks': 'Virtual Bucks', 'r$': 'Robux', 'nmt': 'Nook Miles Tickets',
                'be': 'Blue Essence', 'radianite': 'Radianite Points', 'ow coins': 'Overwatch Coins',
                '₽': 'PokeDollars', '€$': 'Eurodollars', 'z': 'Zenny', 'l': 'Lunas', 'e': 'Eve', 'i': 'Isk',
                'j': 'Jewel', 'sc': 'Star Coins', 'o2': 'Oxygen', 'pu': 'Power Units', 'mc': 'Mana Crystals',
                'es': 'Essence', 'sh': 'Shards', 'st': 'Stars', 'mu': 'Munny', 'b': 'Bolts', 'r': 'Rings',
                'ca': 'Caps', 'rns': 'Runes', 'sl': 'Souls', 'fav': 'Favor', 'am': 'Amber', 'cc': 'Crystal Cores',
                'fg': 'Fragments'
            }
            game_unit_map.update({k: v for k, v in default_mappings.items() if k not in game_unit_map})
            sorted_units = sorted(game_unit_map.keys(), key=len, reverse=True)
            pattern = re.compile(r'(?<!\w)(\d+(?:\.\d+)?)(\s*)(' + '|'.join(map(re.escape, sorted_units)) + r')(?!\w)', re.IGNORECASE)
            text = pattern.sub(lambda m: f"{m.group(1)}{m.group(2)}{game_unit_map.get(m.group(3).lower(), m.group(3))}", text)

        print(f"Processing Area with name '{area_name}' Output Text: \n {text}\n--------------------------")

        # Manage text history
        if self.ignore_previous_var.get():
            max_history_size = 1000
            if area_name in self.text_histories and len(self.text_histories[area_name]) > max_history_size:
                self.text_histories[area_name] = set(list(self.text_histories[area_name])[-max_history_size:])

        # Filter text
        lines = text.split('\n')
        filtered_lines = []
        ignore_items = [item.strip().lower() for item in self.bad_word_list.get().split(',') if item.strip()]
        for line in lines:
            if not line.strip():
                continue
            words = line.split()
            if self.ignore_usernames_var.get():
                filtered_words = []
                i = 0
                while i < len(words):
                    if i < len(words) - 1 and words[i + 1] in [':', ';']:
                        i += 2
                    else:
                        filtered_words.append(words[i])
                        i += 1
                line = ' '.join(filtered_words)

            normalized_line = self.normalize_text(line)
            for item in ignore_items:
                if ' ' in item:
                    norm_phrase = self.normalize_text(item)
                    while norm_phrase in normalized_line:
                        pattern = re.compile(r'\b' + re.escape(item) + r'\b', re.IGNORECASE)
                        line = pattern.sub(' ', line)
                        normalized_line = self.normalize_text(line)
            filtered_words = [word for word in line.split() if not any(self.normalize_text(word) == self.normalize_text(item) for item in ignore_items if ' ' not in item)]
            if not filtered_words:
                continue

            if self.ignore_gibberish_var.get():
                vowels = set('aeiouAEIOU')
                def is_not_gibberish(word):
                    if any(c.isalpha() for c in word) or any(c.isdigit() for c in word):
                        if len(word) <= 3:
                            return True
                        return any(c in vowels for c in word) or any(c.isdigit() for c in word)
                    return False
                filtered_words = [word for word in filtered_words if is_not_gibberish(word)]
            if filtered_words:
                filtered_lines.append(' '.join(filtered_words))

        filtered_text = ' '.join(filtered_lines)
        if self.pause_at_punctuation_var.get():
            for punct in ['.', '!', '?']:
                filtered_text = filtered_text.replace(punct, punct + ' ... ')
            for punct in [',', ';']:
                filtered_text = filtered_text.replace(punct, punct + ' .. ')

        # Add filtered text to history with thread safety
        with self.history_lock:
            if area_name not in self.text_histories:
                self.text_histories[area_name] = []
            self.text_histories[area_name].append(filtered_text)
            # Limit history to the last 50 entries
            if len(self.text_histories[area_name]) > 50:
                self.text_histories[area_name] = self.text_histories[area_name][-50:]

        # Configure and speak text
        if self.tts_engine_var.get() == "Windows TTS":
            if voice_var.get() != "Select Voice":
                voices = self.speaker.GetVoices()
                for voice in voices:
                    if voice.GetDescription() == voice_var.get():
                        self.speaker.Voice = voice
                        break
                else:
                    messagebox.showerror("Error", "Selected voice not found.")
                    return

            try:
                speed = int(speed_var.get())
                if speed > 0:
                    self.speaker.Rate = (speed - 100) // 10
            except ValueError:
                pass

        self.speak_text(filtered_text, voice_var.get(), speed_var.get())

    def show_history(self, area_name):
        """Display the history of read texts for the specified area in a new window with enhanced features."""
        # Check if there's any history to show
        with self.history_lock:
            if area_name not in self.text_histories or not self.text_histories[area_name]:
                messagebox.showinfo("History", "No history available for this area.")
                return

        # Create a new top-level window for history
        history_window = tk.Toplevel(self.root)
        history_window.title(f"History for {area_name}")
        history_window.geometry("600x400")
        history_window.transient(self.root)  # Tie to parent window
        history_window.grab_set()  # Make modal

        # Create a frame for search bar
        search_frame = tk.Frame(history_window)
        search_frame.pack(fill='x', padx=10, pady=5)
        tk.Label(search_frame, text="Search:").pack(side='left')
        search_entry = tk.Entry(search_frame)
        search_entry.pack(side='left', fill='x', expand=True)

        # Create a frame for the text widget and scrollbar
        text_frame = tk.Frame(history_window)
        text_frame.pack(fill='both', expand=True, padx=10, pady=10)

        # Create the text widget and scrollbar inside the frame
        text_widget = tk.Text(text_frame, wrap=tk.WORD, height=20, width=80)
        scrollbar = tk.Scrollbar(text_frame, command=text_widget.yview)

        # Configure grid layout
        text_widget.grid(row=0, column=0, sticky='nsew')
        scrollbar.grid(row=0, column=1, sticky='ns')
        text_frame.grid_rowconfigure(0, weight=1)
        text_frame.grid_columnconfigure(0, weight=1)

        # Link the text widget and scrollbar
        text_widget.config(yscrollcommand=scrollbar.set)

        # Configure tags for styling
        text_widget.tag_configure('header', font=('Arial', 10, 'bold'), justify='left')
        text_widget.tag_configure('box', background='#f0f0f0', relief='groove',
                                 borderwidth=2, spacing1=10, spacing3=10, justify='left')
        text_widget.tag_configure('highlight', background='yellow', justify='left')
        text_widget.tag_configure('content', justify='left', lmargin1=10, lmargin2=10, rmargin=10)
        text_widget.tag_configure('hidden', elide=True)  # Tag to hide content

        # Dictionary to store the start and end indices of each box
        box_indices = {}

        # Populate the text widget with history, applying justification
        with self.history_lock:
            for idx, text in enumerate(self.text_histories[area_name], 1):
                # Insert header
                header_start = text_widget.index(tk.END)
                text_widget.insert(tk.END, f"Reading {idx}:\n", 'header')
                # Insert content with justification
                content_start = text_widget.index(tk.END)
                text_widget.insert(tk.END, f"{text}\n\n", ('box', 'content'))
                content_end = text_widget.index(tk.END)
                # Store the start and end indices of the current box
                box_indices[idx] = (header_start, content_end)
                # Apply content tag to the content portion
                text_widget.tag_add('content', content_start, content_end)

        # Make text read-only
        text_widget.config(state=tk.DISABLED)

        def search_text(event=None):
            # Clear previous highlights and hidden tags
            text_widget.tag_remove('highlight', '1.0', tk.END)
            text_widget.tag_remove('hidden', '1.0', tk.END)
            # Get search query
            query = search_entry.get().strip()
            if not query:
                # If the search query is empty, show all boxes
                text_widget.config(state=tk.NORMAL)
                text_widget.tag_remove('hidden', '1.0', tk.END)
                text_widget.config(state=tk.DISABLED)
                return
            # List to store indices of boxes that contain the query
            matching_boxes = set()
            # Search through each box for the query
            for idx, (start, end) in box_indices.items():
                # Get the text of the current box
                box_text = text_widget.get(start, end)
                if query.lower() in box_text.lower():
                    matching_boxes.add(idx)
            if matching_boxes:
                all_boxes_to_show = matching_boxes  # Only show matching boxes, hide all others
                # Hide boxes that are not in all_boxes_to_show
                text_widget.config(state=tk.NORMAL)
                for idx, (start, end) in box_indices.items():
                    if idx not in all_boxes_to_show:
                        text_widget.tag_add('hidden', start, end)
                text_widget.config(state=tk.DISABLED)
                # Highlight the query in the visible boxes
                for idx in all_boxes_to_show:
                    start, end = box_indices[idx]
                    text_widget.config(state=tk.NORMAL)
                    text_widget.tag_remove('highlight', start, end)
                    # Search for the query within the current box
                    idx_pos = start
                    while True:
                        idx_pos = text_widget.search(query, idx_pos, nocase=1, stopindex=end)
                        if not idx_pos:
                            break
                        lastidx = f"{idx_pos}+{len(query)}c"
                        text_widget.tag_add('highlight', idx_pos, lastidx)
                        idx_pos = lastidx
                    text_widget.config(state=tk.DISABLED)
            else:
                # If no matches, hide all boxes
                text_widget.config(state=tk.NORMAL)
                for idx, (start, end) in box_indices.items():
                    text_widget.tag_add('hidden', start, end)
                text_widget.config(state=tk.DISABLED)

        # Bind the search_text function to the KeyRelease event for live search
        search_entry.bind('<KeyRelease>', search_text)

        # Add context menu for copying text
        context_menu = tk.Menu(text_widget, tearoff=0)
        context_menu.add_command(label="Copy", command=lambda: text_widget.event_generate('<<Copy>>'))
        context_menu.add_command(label="Select All", command=lambda: text_widget.tag_add('sel', '1.0', 'end'))
        text_widget.bind("<Button-3>", lambda event: context_menu.tk_popup(event.x_root, event.y_root))

        # Center the window relative to the main window
        history_window.update_idletasks()
        x = self.root.winfo_rootx() + (self.root.winfo_width() // 2 - history_window.winfo_width() // 2)
        y = self.root.winfo_rooty() + (self.root.winfo_height() // 2 - history_window.winfo_height() // 2)
        history_window.geometry(f"+{x}+{y}")

        # Ensure hotkeys are restored on close
        def on_close():
            history_window.destroy()
            self.restore_all_hotkeys()

        history_window.protocol("WM_DELETE_WINDOW", on_close)
        history_window.bind("<Escape>", lambda e: on_close())

    def cleanup(self):
        """Proper cleanup method for the application"""
        print("Performing cleanup...")
        try:
            # Close console window if it exists
            if hasattr(self, 'console_window'):
                try:
                    self.console_window.window.destroy()
                    self.console_window = None  # Set to None instead of delattr for simplicity
                except Exception:
                    pass  # Silent failure is acceptable during cleanup

            # Restore original stdout
            if hasattr(sys, 'stdout_original'):
                sys.stdout = sys.stdout_original

            # Explicitly close PIL Image objects in latest_images
            for img in self.latest_images.values():
                if isinstance(img, Image.Image):
                    try:
                        img.close()
                    except Exception:
                        pass  # Silent failure for image closing
            self.latest_images.clear()

            # Cleanup hotkeys
            try:
                self.disable_all_hotkeys()
            except Exception as e:
                print(f"Warning: Error cleaning up hotkeys: {e}")

            # Cleanup TTS engine
            if hasattr(self, 'engine') and self.engine is not None:
                try:
                    self.engine.endLoop()
                except Exception:
                    pass  # Silent failure if engine cleanup fails
                self.engine = None

            # Cleanup speaker
            if hasattr(self, 'speaker'):
                self.speaker = None  # Set to None instead of del for consistency

            # Clear data structures
            for collection in (self.hotkey_scancodes, self.processing_settings, 
                             self.text_histories, self.hotkeys, self.areas):
                collection.clear()

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
        """Check if text is valid and not gibberish, based on character analysis."""
        # Skip empty or whitespace-only text
        if not text.strip():
            return False

        # Define valid characters
        VALID_CHARS = set(".,!?'\"- ")  # Common punctuation and space

        # Count valid and invalid characters efficiently
        valid_chars = sum(1 for char in text if char.isalnum() or char in VALID_CHARS)
        invalid_chars = len(text) - valid_chars

        # Reject if invalid characters exceed half of valid ones
        if invalid_chars > valid_chars / 2:
            return False

        # Check for repeated OCR artifact symbols
        OCR_ARTIFACTS = "/\\|[]{}=<>+*"
        if any(symbol * 2 in text for symbol in OCR_ARTIFACTS):
            return False

        # Ensure cleaned text has at least 2 alphanumeric characters
        clean_text = ''.join(c for c in text if c.isalnum() or c.isspace())
        if len(clean_text.strip()) < 2:
            return False

        return True

    def _process_tts_queue(self):
        try:
            while not self.tts_queue.empty():
                command, data = self.tts_queue.get_nowait()
                if command == "SPEAK":
                    text, voice_name, speed = data
                    if self.is_speaking:
                        try:
                            self.speaker.Speak("", 2)  # Purge before speaking new text
                        except Exception as e:
                            print(f"Error purging speech: {e}")
                            self.reinit_speaker()
                    if not self.speaker:
                        print("Cannot speak, TTS speaker is not available.")
                        continue
                    try:
                        self.speaker.Resume()
                        if voice_name != "Select Voice":
                            for voice in self.speaker.GetVoices():
                                if voice.GetDescription() == voice_name:
                                    self.speaker.Voice = voice
                                    break
                        self.speaker.Rate = (int(speed) - 100) // 10
                        self.speaker.Speak(text, 1)
                        self.is_speaking = True
                        print(f"Speech started via queue for: \"{text[:30]}...\"")
                    except Exception as e:
                        print(f"Error processing SPEAK command: {e}")
                        self.is_speaking = False
                        self.reinit_speaker()
                elif command == "STOP":
                    if self.speaker:
                        try:
                            self.speaker.Pause()
                            self.speaker.Speak("", 2)
                            print("Windows TTS speech stopped successfully.")
                        except Exception as e:
                            print(f"Error stopping speech: {e}")
                            self.reinit_speaker()
                    self.is_speaking = False
                    while not self.tts_queue.empty():
                        try:
                            self.tts_queue.get_nowait()
                        except queue.Empty:
                            break
        except queue.Empty:
            pass
        finally:
            self.root.after(100, self._process_tts_queue)

    def reinit_speaker(self):
        print("Attempting to re-initialize the Windows TTS speaker...")
        try:
            self.speaker = win32com.client.Dispatch("SAPI.SpVoice")
            self.speaker.Volume = int(self.volume.get())
            self.is_speaking = False
            print("Speaker re-initialized successfully.")
        except Exception as e:
            print(f"FATAL: Failed to re-initialize speaker: {e}")
            self.speaker = None

    def show_processing_feedback(self, area_name):
        """Display temporary processing feedback for the specified area."""
        # Cancel any existing feedback timer
        if hasattr(self, '_feedback_timer') and self._feedback_timer:
            self.root.after_cancel(self._feedback_timer)

        # Update status label with processing message
        self.status_label.config(text=f"Processing Area: {area_name}")

        # Schedule clearing the label after 1.3 seconds
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
    from tkinterdnd2 import DND_FILES, TkinterDnD
    
    # Initialize Tkinter with drag-and-drop support
    root = TkinterDnD.Tk()
    app = GameTextReader(root)
    app.add_read_area(removable=False, editable_name=False, area_name="Auto Read")
    
    # Load Auto Read settings from temp file if available
    TEMP_DIR = os.path.join(tempfile.gettempdir(), 'GameReader')
    temp_path = os.path.join(TEMP_DIR, 'auto_read_settings.json')
    if os.path.exists(temp_path) and app.areas:
        try:
            with open(temp_path, 'r') as f:
                settings = json.load(f)
            
            # Unpack the permanent "Auto Read" area tuple
            area_frame, hotkey_button, set_area_button, area_name_var, preprocess_var, voice_var, speed_var = app.areas[0]
            
            # Apply UI settings with defaults
            preprocess_var.set(settings.get('preprocess', False))
            voice_var.set(settings.get('voice', 'Select Voice'))
            speed_var.set(settings.get('speed', '100'))
            
            # Configure hotkey if present
            if 'hotkey' in settings:
                hotkey_button.hotkey = settings['hotkey']
                display_name = settings['hotkey'].replace('num_', 'num:')
                hotkey_button.config(text=f"Set Hotkey: [ {display_name} ]")
                app.setup_hotkey(hotkey_button, area_frame)
            
            # Load processing settings with defaults in a concise way
            defaults = {
                'brightness': 1.0, 'contrast': 1.0, 'saturation': 1.0, 'sharpness': 1.0,
                'blur': 0.0, 'hue': 0.0, 'exposure': 1.0, 'threshold': 128,
                'threshold_enabled': False, 'preprocess': settings.get('preprocess', False)
            }
            app.processing_settings['Auto Read'] = {
                key: settings.get('processing', {}).get(key, default) for key, default in defaults.items()
            }
            
            # Set interrupt on new scan option
            app.interrupt_on_new_scan_var.set(settings.get('stop_read_on_select', False))
            
            # Update processing widgets if they exist
            if hasattr(app, 'processing_settings_widgets'):
                widgets = app.processing_settings_widgets.get('Auto Read', {})
                for key, value in app.processing_settings['Auto Read'].items():
                    if key in widgets:
                        widgets[key].set(value)
                print("Loaded Auto Read settings successfully")
        except (json.JSONDecodeError, Exception) as e:
            print(f"Error loading Auto Read settings: {e}")
            # Set default processing settings on failure
            if 'Auto Read' not in app.processing_settings:
                app.processing_settings['Auto Read'] = {
                    'brightness': 1.0, 'contrast': 1.0, 'saturation': 1.0, 'sharpness': 1.0,
                    'blur': 0.0, 'hue': 0.0, 'exposure': 1.0, 'threshold': 128,
                    'threshold_enabled': False
                }
    
    root.mainloop()