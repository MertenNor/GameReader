###
###  I know.. This code is.. well not great, all made with AI, but it works. feel free to make any changes!
###

import tkinter as tk
from tkinter import filedialog, ttk, messagebox, simpledialog
from PIL import ImageGrab, Image, ImageTk, ImageEnhance, ImageFilter
from functools import partial
import pytesseract
import pyttsx3
import json
import keyboard
import threading
import sys
import io
import datetime
import os
import win32com.client
import mouse
import win32gui
import time
import webbrowser
import win32api
import win32ui
import win32con
import requests  # For checking updates
import re  # For version extraction

APP_VERSION = "0.7.6"

CHANGELOG = """
- Changelog -  
 - Fixed a bug where the Auto Read function didn't work with fullscreen applications.
 - Improved overall implementation of the AutoRead feature.
"""


# Create a StringIO buffer to capture print statements
log_buffer = io.StringIO()

# Redirect standard output to the StringIO buffer
sys.stdout = log_buffer

# --- Custom Hotkey Conflict Warning Dialog (No Symbol, Styled OK Button) ---
def show_thinkr_warning(game_reader, area_name):
    import keyboard
    import mouse
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

def restore_all_hotkeys():
    """
    Restore all hotkeys for the application. This function should be adapted to your app's hotkey storage logic.
    """
    import keyboard
    import mouse
    # Example: restore hotkeys for each area if you have a list/dict of them
    # Replace the following with your actual hotkey setup logic
    if hasattr(globals(), 'app_hotkeys'):
        for hotkey, callback in globals()['app_hotkeys'].items():
            keyboard.add_hotkey(hotkey, callback)
    # If you use mouse hotkeys, add them similarly
    # mouse.add_hotkey(...)

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

        self.log_buffer = log_buffer
        self.layout_file_var = layout_file_var
        self.latest_area_name_var = latest_area_name_var
        self.photo = None  # Keep a reference to prevent garbage collection

        # Add line limit constant
        self.MAX_LINES = 100
        
        self.update_console()

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
        self.window.title(f"Customize Processing for {area_name}")
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
        info_text = f"Processing applies to unprocessed images; results may differ if the preview is already processed.\n \nShowing previous image captured in area: {area_name}"
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

        entry = ttk.Entry(label_frame, textvariable=entry_var)
        entry.pack()
    
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

        # Find and enable the preprocess checkbox for this area
        for area_frame, _, _, area_name_var, preprocess_var, _, _ in self.game_text_reader.areas:
            if area_name_var.get() == area_name:
                preprocess_var.set(True)  # Enable the checkbox
                break

        # Check if there's a current layout file
        if not self.game_text_reader.layout_file.get():
            # Create custom dialog
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
            
            # Variable to store result
            self.dialog_result = False
            
            # Button callbacks
            def on_save():
                self.dialog_result = True
                dialog.destroy()
                
            def on_cancel():
                self.dialog_result = False
                dialog.destroy()
            
            # Add buttons
            save_button = tk.Button(button_frame, text="Save...", command=on_save, width=10)
            save_button.pack(side=tk.LEFT, padx=10)
            
            cancel_button = tk.Button(button_frame, text="Cancel", command=on_cancel, width=10)
            cancel_button.pack(side=tk.LEFT, padx=10)
            
            # Wait for dialog to close
            self.window.wait_window(dialog)
            
            if not self.dialog_result:
                return  # User clicked Cancel

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

def check_for_update(local_version):
    """
    Fetch the remote GameReader.py from GitHub, extract version and changelog, compare to local_version.
    If remote version is newer, show a popup.
    """
    GITHUB_RAW_URL = "https://raw.githubusercontent.com/MertenNor/GameReader/main/GameReader.py"
    try:
        resp = requests.get(GITHUB_RAW_URL, timeout=5)
        if resp.status_code == 200:
            remote_content = resp.text
            remote_version = extract_version_from_code(remote_content)
            remote_changelog = extract_changelog_from_code(remote_content)
            if remote_version and version_tuple(remote_version) > version_tuple(local_version):
                # Show popup
                import tkinter as tk
                from tkinter import messagebox
                def open_github():
                    import webbrowser
                    webbrowser.open('https://github.com/MertenNor/GameReader')
                root = tk._default_root or tk.Tk()
                if not tk._default_root:
                    root.withdraw()
                msg = f"A new version of GameReader is available!\n\nLocal version: {local_version}\nLatest version: {remote_version}"
                if remote_changelog:
                    msg += f"\n\nWhat's new:\n{remote_changelog}"
                msg += "\n\nVisit GitHub to download the update."
                if messagebox.askyesno("Update Available", msg + "\n\nOpen GitHub page now?"):
                    open_github()
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
        self.root.title(f"Game Text Reader v{APP_VERSION}")
        # --- Update check on startup ---
        local_version = APP_VERSION
        threading.Thread(target=lambda: check_for_update(local_version), daemon=True).start()
        # --- End update check ---
        self.root.geometry("1100x180")  # Initial window size (height reduced for less vertical tallness)
        self.layout_file = tk.StringVar()
        self.latest_images = {}  # Use a dictionary to store images for each area
        self.latest_area_name = tk.StringVar()  # Ensure this is defined
        self.areas = []
        self.stop_hotkey = None  # Variable to store the STOP hotkey
        self.engine = pyttsx3.init()
        self.engine_lock = threading.Lock()  # Lock for the text-to-speech engine
        self.bad_word_list = tk.StringVar()  # StringVar for the bad word list
        self.hotkeys = set()  # Track registered hotkeys
        self.is_speaking = False  # Flag to track if the engine is speaking
        self.processing_settings = {}  # Dictionary to store processing settings for each area
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

        self.hotkey_scancodes = {}  # Dictionary to store scan codes for hotkeys
        self.setting_hotkey = False  # Flag to track if we're in hotkey setting mode
        self.unhook_timer = None  # Add this line
        
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

        self.text_histories = {}  # Dictionary to store text history for each area
        self.ignore_previous_var = tk.BooleanVar(value=False)  # Variable for the checkbox
        self.ignore_gibberish_var = tk.BooleanVar(value=False)  # Variable for the gibberish checkbox
        self.pause_at_punctuation_var = tk.BooleanVar(value=False)  # Variable for punctuation pauses
        self.fullscreen_mode_var = tk.BooleanVar(value=False)  # Variable for fullscreen mode

        self.setup_gui()
        self.voices = self.engine.getProperty('voices')  # Get available voices
        
        self.stop_keyboard_hook = None
        self.stop_mouse_hook = None
        self.setting_hotkey_mouse_hook = None
        self.unhook_timer = None
        
        # Add this line to handle window closing
        root.protocol("WM_DELETE_WINDOW", lambda: (self.cleanup(), root.destroy()))

    def speak_text(self, text):
        """Speak text using win32com.client (SAPI.SpVoice)."""
        # Always check and stop speech if interrupt is enabled
        if hasattr(self, 'interrupt_on_new_scan_var') and self.interrupt_on_new_scan_var.get():
            self.stop_speaking()  # Always attempt to stop, even if not currently speaking
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

    def stop_speaking(self):
        """Stop the ongoing speech immediately."""
        try:
            if self.speaker:
                self.speaker.Speak("", 2)  # Use SVSFPurgeBeforeSpeak flag
            self.is_speaking = False
            # Reinitialize speaker
            self.speaker = win32com.client.Dispatch("SAPI.SpVoice")
            self.speaker.Volume = int(self.volume.get())
            print("Speech stopped.\n--------------------------")
        except Exception as e:
            print(f"Error stopping speech: {e}")
            try:
                self.speaker = win32com.client.Dispatch("SAPI.SpVoice")
                self.speaker.Volume = int(self.volume.get())
                self.is_speaking = False
            except Exception as e2:
                print(f"Error reinitializing speaker: {e2}")

    def setup_gui(self):
        # Create main frames for organization
        top_frame = tk.Frame(self.root)
        top_frame.pack(fill='x', padx=10, pady=5)
        
        control_frame = tk.Frame(self.root)
        control_frame.pack(fill='x', padx=10, pady=5)
        
        options_frame = tk.Frame(self.root)
        options_frame.pack(fill='x', padx=10, pady=5)
        
        # Top frame contents - Title and buttons
        title_label = tk.Label(top_frame, text=f"Game Text Reader v{APP_VERSION}", font=("Helvetica", 12, "bold"))
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
        tk.Label(layout_frame, textvariable=self.layout_file, font=("Helvetica", 10, "bold")).pack(side='left', padx=5)
        
        # Layout control buttons
        layout_buttons_frame = tk.Frame(control_frame)
        layout_buttons_frame.pack(side='right')
        
        save_button = tk.Button(layout_buttons_frame, text="Save Layout", command=self.save_layout)
        save_button.pack(side='left', padx=5)
        
        load_button = tk.Button(layout_buttons_frame, text="Load Layout..", command=self.load_layout)
        load_button.pack(side='left', padx=5)
        
        # Options frame contents
        # Word filtering frame
        filter_frame = tk.Frame(options_frame)
        filter_frame.pack(fill='x', pady=5)
        
        tk.Label(filter_frame, text="Ignored Word List:").pack(side='left')
        self.bad_word_entry = tk.Entry(filter_frame, textvariable=self.bad_word_list)
        self.bad_word_entry.pack(side='left', fill='x', expand=True)
        
        # Single line of checkboxes
        checkbox_frame = tk.Frame(options_frame)
        checkbox_frame.pack(fill='x', pady=5)
        
        # Create all checkboxes in a single line
        self.create_checkbox(checkbox_frame, "Ignore usernames:", self.ignore_usernames_var, side='left', padx=5)
        self.create_checkbox(checkbox_frame, "Ignore previous spoken words:", self.ignore_previous_var, side='left', padx=5)
        self.create_checkbox(checkbox_frame, "Ignore gibberish:", self.ignore_gibberish_var, side='left', padx=5)
        self.create_checkbox(checkbox_frame, "Pause at punctuation:", self.pause_at_punctuation_var, side='left', padx=5)
        self.create_checkbox(checkbox_frame, "Fullscreen mode:", self.fullscreen_mode_var, side='left', padx=5)

        
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
        
        # Frame for the areas - now with scrollable canvas
        self.area_outer_frame = tk.Frame(self.root)
        self.area_outer_frame.pack(fill='both', expand=True, pady=5)

        self.area_canvas = tk.Canvas(self.area_outer_frame, highlightthickness=0)
        self.area_canvas.pack(side='left', fill='both', expand=True)
        self.area_scrollbar = tk.Scrollbar(self.area_outer_frame, orient='vertical', command=self.area_canvas.yview)
        self.area_scrollbar.pack(side='right', fill='y')

        # Enable mouse wheel scrolling for the canvas
        def _on_mousewheel(event):
            # Windows scroll direction
            self.area_canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')
        self.area_canvas.bind_all('<MouseWheel>', _on_mousewheel)
        
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
        info_window.title("Game Text Reader - Information")
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
                area_frame, hotkey_button, _, _, _, _, _ = area
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
                               text=f"Game Text Reader v{APP_VERSION}", 
                               font=("Helvetica", 16, "bold"))
        title_label.pack(side='left')
        
        # Credits and Links Section - Improved Layout
        credits_frame = ttk.Frame(main_frame)
        credits_frame.pack(fill='x', pady=(0, 20))

        # --- Horizontal Frame for Program Info and Official Links ---
        credits_row = ttk.Frame(credits_frame)
        credits_row.pack(fill='x', pady=(0, 5))

        # Program Info (left)
        proginfo_frame = ttk.Frame(credits_row)
        proginfo_frame.pack(side='left', padx=0, anchor='n')
        proginfo_label = ttk.Label(proginfo_frame, text="Program Information", font=("Helvetica", 11, "bold"))
        proginfo_label.pack(side='top', anchor='w')
        designer_label = ttk.Label(proginfo_frame, text="Designer: MertenNor", font=("Helvetica", 10))
        designer_label.pack(side='left', padx=(0, 15))
        # Coder line with embedded Cursor link
        coder_frame = ttk.Frame(proginfo_frame)
        coder_frame.pack(side='left')
        coder_label1 = ttk.Label(coder_frame, text="Coder: ChatGPT / Claude-3.5 via ", font=("Helvetica", 10))
        coder_label1.pack(side='left')
        cursor_link = ttk.Label(coder_frame, text="Cursor", font=("Helvetica", 10, "underline"), foreground='black', cursor='hand2')
        cursor_link.pack(side='left')
        cursor_link.bind("<Button-1>", lambda e: open_url("https://www.cursor.com/"))
        cursor_link.bind("<Enter>", lambda e: cursor_link.configure(font=("Helvetica", 10, "underline")))
        cursor_link.bind("<Leave>", lambda e: cursor_link.configure(font=("Helvetica", 10)))

        # Official Links (right)
        links_frame = ttk.Frame(credits_row)
        links_frame.pack(side='left', padx=(80, 0), anchor='n')
        links_label = ttk.Label(links_frame, text="Official Links", font=("Helvetica", 11, "bold"))
        links_label.pack(side='top', anchor='w')

        # GitHub link
        github_frame = ttk.Frame(links_frame)
        github_frame.pack(fill='x', pady=(2, 0))
        github_label = ttk.Label(github_frame, text="GitHub: ", font=("Helvetica", 10, "bold"))
        github_label.pack(side='left')
        github_link = ttk.Label(github_frame, text="GitHub.com/mertennor/gamereader", font=("Helvetica", 10), foreground='black', cursor='hand2')
        github_link.pack(side='left')
        github_link.bind("<Button-1>", lambda e: open_url("https://github.com/MertenNor/GameReader"))
        github_link.bind("<Enter>", lambda e: github_link.configure(font=("Helvetica", 10, "underline")))
        github_link.bind("<Leave>", lambda e: github_link.configure(font=("Helvetica", 10)))

        # --- Section: Support & Feedback ---
        support_frame = ttk.Frame(credits_frame)
        support_frame.pack(fill='x', pady=(10, 5))
        support_label = ttk.Label(support_frame, text="Support & Feedback", font=("Helvetica", 11, "bold"))
        support_label.pack(side='top', anchor='w')

        # Coffee link
        coffee_frame = ttk.Frame(support_frame)
        coffee_frame.pack(fill='x', pady=(2, 0))
        coffee_label = ttk.Label(coffee_frame, text="Buy me a Coffee: ", font=("Helvetica", 10, "bold"))
        coffee_label.pack(side='left')
        support_link = ttk.Label(coffee_frame, text="BuyMeaCoffee.com/mertennor ☕", font=("Helvetica", 10), foreground='black', cursor='hand2')
        support_link.pack(side='left')
        support_link.bind("<Button-1>", lambda e: open_url("https://www.buymeacoffee.com/mertennor"))
        support_link.bind("<Enter>", lambda e: support_link.configure(font=("Helvetica", 10, "underline")))
        support_link.bind("<Leave>", lambda e: support_link.configure(font=("Helvetica", 10)))

        # Feedback link
        feedback_frame = ttk.Frame(support_frame)
        feedback_frame.pack(fill='x', pady=(2, 0))
        feedback_label = ttk.Label(feedback_frame, text="Bug Reports & Feedback via Google Form: ", font=("Helvetica", 10, "bold"))
        feedback_label.pack(side='left')
        feedback_link = ttk.Label(feedback_frame, text="Forms.Gle/8YBU8atkgwjyzdM79 🐞", font=("Helvetica", 10), foreground='black', cursor='hand2')
        feedback_link.pack(side='left')
        feedback_link.bind("<Button-1>", lambda e: open_url("https://forms.gle/8YBU8atkgwjyzdM79"))
        feedback_link.bind("<Enter>", lambda e: feedback_link.configure(font=("Helvetica", 10, "underline")))
        feedback_link.bind("<Leave>", lambda e: feedback_link.configure(font=("Helvetica", 10)))

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
        # Download instruction and clickable URL on the same line
        download_row = ttk.Frame(tesseract_frame)
        download_row.pack(anchor='w')
        download_label = ttk.Label(download_row,
                                   text="Download the latest version here: ",
                                   font=("Helvetica", 10, "bold"),
                                   foreground='red')
        download_label.pack(side='left')
        tesseract_link = ttk.Label(download_row,
                                   text="www.gitub.com/tesseract-ocr/tesseract/releases",
                                   font=("Helvetica", 10),
                                   foreground='blue',
                                   cursor='hand2')
        tesseract_link.pack(side='left')
        tesseract_link.bind("<Button-1>", lambda e: open_url("https://github.com/tesseract-ocr/tesseract/releases"))
        tesseract_link.bind("<Enter>", lambda e: tesseract_link.configure(font=("Helvetica", 10, "underline")))
        tesseract_link.bind("<Leave>", lambda e: tesseract_link.configure(font=("Helvetica", 10)))
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
        
        # Create text widget with custom styling
        text_widget = tk.Text(content_frame, 
                             wrap=tk.WORD, 
                             yscrollcommand=scrollbar.set,
                             font=("Helvetica", 10),
                             padx=10,
                             pady=10,
                             spacing1=2,  # Space between lines
                             spacing2=2,  # Space between paragraphs
                             background='#f5f5f5',  # Light gray background
                             border=1)
        text_widget.pack(fill='both', expand=True)
        
        # Configure scrollbar
        scrollbar.config(command=text_widget.yview)
        
        # Create a bold font tag
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


            ("AutoRead\n", 'bold'),
            ("------------------------\n", None),
("When assigned a hotkey, the program will automatically read the text in the selected area.\n", None),
("The save button here will save the settings for the AutoRead area only.\n", None),
("Note! This works best with applications in windowed borderless mode.\n", None),
("This save file can be found here: C:\\Users\\<username>\\AppData\\Local\\Temp\nFilename: auto_read_settings.json\n\n", None),
            
            ("Add read area\n", 'bold'),
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
            ("A list of words or phrases to ignore while reading text. Example: Chocolate, Apple, Banana\n", None),
            ("This will then be ignored in all areas.\n\n\n", None),
            
            ("CHECKBOX OPTIONS\n", 'bold'),
            ("════════════════\n\n", None),
            
            ("Ignore usernames *EXPERIMENTAL* \n", 'bold'),
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
            ("Feature for capturing text from fullscreen applications. May cause brief screen flicker during capture for the program to take an updated screenshot.\n\n\n", None),

            ("TIPS AND TRICKS\n", 'bold'),
            ("═════════════\n\n", None),
            
            ("• Use image processing for areas with difficult-to-read text\n\n", None),
            
            ("• Create two identical areas with different hotkeys: assign one a male voice and the other a female voice.\n", None),
            ("  This lets you easily switch between male and female voices for text, ideal for game dialogue.\n\n", None),
            ("• Experiment with different preprocessing settings for optimal text recognition in your specific use case.\n\n", None),
            ("• Want more Voices? add them in windows. \n", None),
            ("  Go to Settings > Time & Language > Speech > Manage voices, then click 'Add voices' to install new ones. \n", None),
            ("  you can also find other Narrator voices online that you can add to windows. \n", None),
        ]
        
        # Insert text with tags
        for text, tag in info_text:
            text_widget.insert('end', text, tag)
        
        # Make text widget read-only
        text_widget.config(state='disabled')
        
        # Add bottom frame for close button with padding
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(fill='x', pady=(20, 0))
        
        # Add close button
        close_button = ttk.Button(bottom_frame, 
                                 text="wait?? what is this button doing down here?", 
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
        def on_key_press(event):
            # Set flag to ignore hotkey triggers
            self.setting_hotkey = True
            
            # Get the key name
            key_name = event.name
            if event.scan_code in self.numpad_scan_codes:
                key_name = f"num_{self.numpad_scan_codes[event.scan_code]}"
            
            # Check if this key is already used by any area
            for area_frame, hotkey_button, _, area_name_var, _, _, _ in self.areas:
                if hasattr(hotkey_button, 'hotkey') and hotkey_button.hotkey == key_name:
                    messagebox.showwarning("Hotkey Conflict", 
                        f"This key is already used by area '{area_name_var.get()}'.\nPlease choose a different key.")
                    # Reset button text to previous state if it exists, or default text
                    if hasattr(self, 'stop_hotkey'):
                        display_name = self.stop_hotkey.replace('num_', 'num:') if self.stop_hotkey.startswith('num_') else self.stop_hotkey
                        self.stop_hotkey_button.config(text=f"Stop Hotkey: [ {display_name} ]")
                    else:
                        self.stop_hotkey_button.config(text="Set Stop Hotkey")
                    self.setting_hotkey = False
                    return
            
            # Remove existing stop hotkey if it exists
            if hasattr(self, 'stop_hotkey'):
                try:
                    if hasattr(self.stop_hotkey_button, 'mock_button'):
                        if hasattr(self.stop_hotkey_button.mock_button, 'keyboard_hook'):
                            keyboard.unhook(self.stop_hotkey_button.mock_button.keyboard_hook)
                        if hasattr(self.stop_hotkey_button.mock_button, 'mouse_hook'):
                            mouse.unhook(self.stop_hotkey_button.mock_button.mouse_hook)
                except:
                    pass
            
            self.stop_hotkey = key_name
            
            # Create a mock button object to use with setup_hotkey
            mock_button = type('MockButton', (), {'hotkey': key_name, 'is_stop_button': True})
            self.stop_hotkey_button.mock_button = mock_button  # Store reference to mock button
            self.setup_hotkey(self.stop_hotkey_button.mock_button, None)  # Pass None as area_frame for stop hotkey
            
            display_name = key_name.replace('num_', 'num:') if key_name.startswith('num_') else key_name
            self.stop_hotkey_button.config(text=f"Stop Hotkey: [ {display_name} ]")
            print(f"Set Stop hotkey: {key_name}\n--------------------------")
            
            # Clean up the temporary hooks
            keyboard.unhook(self.temp_keyboard_hook)
            if hasattr(self, 'setting_hotkey_mouse_hook'):
                mouse.unhook(self.setting_hotkey_mouse_hook)
            
            self.setting_hotkey = False
            return

        def on_mouse_click(event):
            if not isinstance(event, mouse.ButtonEvent) or event.event_type != mouse.DOWN:
                return

            # Early return for left/right mouse buttons
            if event.button in [1, 2]:
                messagebox.showerror("Error", "Left and right mouse buttons cannot be used as hotkeys.")
                # Reset button text to previous state if it exists, or default text
                if hasattr(self, 'stop_hotkey'):
                    display_name = self.stop_hotkey.replace('num_', 'num:') if self.stop_hotkey.startswith('num_') else self.stop_hotkey
                    self.stop_hotkey_button.config(text=f"Stop Hotkey: [ {display_name} ]")
                else:
                    self.stop_hotkey_button.config(text="Press any key...")
                return

            key_name = f'button{event.button}'
            print(f"Setting mouse button {key_name} as stop hotkey")
            
            # Remove existing stop hotkey if it exists
            if hasattr(self, 'stop_hotkey'):
                try:
                    keyboard.unhook_all()
                    mouse.unhook_all()
                except:
                    pass
            
            self.stop_hotkey = key_name
            
            # Create a mock button object to use with setup_hotkey
            mock_button = type('MockButton', (), {'hotkey': key_name})
            self.setup_hotkey(mock_button, None)  # Pass None as area_frame for stop hotkey
            
            self.stop_hotkey_button.config(text=f"Stop Hotkey: [ {key_name} ]")
            print(f"Set Stop hotkey: {key_name}")
            self.setting_hotkey = False

        # Clean up only the stop hotkey hooks before setting new ones
        try:
            if hasattr(self, 'stop_hotkey_button') and hasattr(self.stop_hotkey_button, 'mock_button'):
                if hasattr(self.stop_hotkey_button.mock_button, 'keyboard_hook'):
                    keyboard.unhook(self.stop_hotkey_button.mock_button.keyboard_hook)
                if hasattr(self.stop_hotkey_button.mock_button, 'mouse_hook'):
                    mouse.unhook(self.stop_hotkey_button.mock_button.mouse_hook)
        except Exception as e:
            print(f"Error cleaning up stop hotkey hooks: {e}")

        self.stop_hotkey_button.config(text="Press any key...")
        # Store references to temporary hooks so we can clean them up later
        self.temp_keyboard_hook = keyboard.on_press(on_key_press)
        self.setting_hotkey_mouse_hook = mouse.hook(on_mouse_click)

    def restart_tesseract(self):
        """Forcefully stop the speech and reinitialize the system."""
        print("Forcing stop...")
        try:
            self.stop_speaking()  # Stop the speech
            print("System reinitialized. Audio stopped.")
        except Exception as e:
            print(f"Error during forced stop: {e}")

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
        area_frame = tk.Frame(self.area_frame)
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
        voice_menu = tk.OptionMenu(area_frame, voice_var, "Select Voice", *[voice.name for voice in self.voices])
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
                settings = {
                    'preprocess': preprocess_var.get(),
                    'voice': voice_var.get(),
                    'speed': speed_var.get(),
                    'hotkey': hotkey,
                    'stop_read_on_select': self.interrupt_on_new_scan_var.get(),
                }
                # Optionally add more settings as needed
                temp_path = os.path.join(tempfile.gettempdir(), 'auto_read_settings.json')
                with open(temp_path, 'w') as f:
                    json.dump(settings, f)
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

        # Automatically resize the window
        self.resize_window()

    def remove_area(self, area_frame, area_name):
        # Find and clean up the hotkey for this area
        for area in self.areas:
            if area[0] == area_frame:  # Found matching frame
                hotkey_button = area[1]  # Get the hotkey button
                # Clean up keyboard hook if it exists
                if hasattr(hotkey_button, 'keyboard_hook'):
                    try:
                        keyboard.unhook(hotkey_button.keyboard_hook)
                    except Exception as e:
                        print(f"Error cleaning up keyboard hook: {e}")
                # Clean up mouse hook if it exists
                if hasattr(hotkey_button, 'mouse_hook'):
                    try:
                        mouse.unhook(hotkey_button.mouse_hook)
                    except Exception as e:
                        print(f"Error cleaning up mouse hook: {e}")
                break

        # Clean up image if it exists
        if area_name in self.latest_images:
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

    def resize_window(self):
        """Resize the window based on the number of areas and the longest area line. Keeps window height fixed after 10 areas, enabling scrollbar."""
        base_height = 210  # Height for main controls, padding, etc.
        min_width = 950
        max_width = 1600
        area_frame_height = 0
        if len(self.areas) > 0:
            self.area_frame.update_idletasks()
            area_frame_height = self.area_frame.winfo_height()
        # Calculate area height for up to 10 areas
        visible_area_count = min(10, len(self.areas))
        area_row_height = 60  # 55 for frame, 5 for padding
        fixed_canvas_height = visible_area_count * area_row_height
        # The total height for window (fixed after 10 areas)
        # If more than 10 areas, cap height strictly to 10 area rows
        if len(self.areas) > 10:
            total_height = base_height + fixed_canvas_height
        else:
            total_height = base_height + area_frame_height
        # Never allow window to grow vertically beyond this cap
        total_height = min(total_height, base_height + 10 * area_row_height)
        total_height = max(total_height, 250)
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
        # Set minimum window size, but never force vertical growth beyond the cap
        self.root.minsize(window_width, min(total_height, base_height + 5 * area_row_height))
        # Scrollbar logic
        if hasattr(self, 'area_scrollbar'):
            if len(self.areas) > 8:
                self.area_scrollbar.pack(side='right', fill='y')
                self.area_canvas.configure(yscrollcommand=self.area_scrollbar.set)
                self.area_canvas.config(height=fixed_canvas_height)
            else:
                self.area_scrollbar.pack_forget()
                self.area_canvas.config(height=area_frame_height)
        # Only increase window size if needed, never grow vertically past the cap
        cur_width = self.root.winfo_width()
        cur_height = self.root.winfo_height()
        max_height = base_height + 5 * area_row_height
        if cur_width < window_width or cur_height < total_height:
            # Always cap the height
            # self.root.geometry(f"{window_width}x{min(total_height, max_height)}")  # Disabled dynamic resizing
            pass  # No dynamic resizing, keep window fixed
        # Window size is now fixed from __init__

        self.root.update_idletasks()  # Ensure geometry is applied

    def set_area(self, frame, area_name_var, set_area_button):
        x1, y1, x2, y2 = 0, 0, 0, 0
        selection_cancelled = False
        # --- Disable all hotkeys before starting area selection ---
        try:
            keyboard.unhook_all()
            mouse.unhook_all()
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
                # Stop speech on mouse release if the checkbox is checked
                if hasattr(self, 'interrupt_on_new_scan_var') and self.interrupt_on_new_scan_var.get():
                    self.stop_speaking()
                # Convert canvas coordinates to screen coordinates for the final area
                x2 = event.x_root
                y2 = event.y_root
                x1_screen = x1 + min_x
                y1_screen = y1 + min_y
                frame.area_coords = (
                    min(x1_screen, x2), 
                    min(y1_screen, y2),
                    max(x1_screen, x2), 
                    max(y1_screen, y2)
                )
                # If this is the Auto Read area, trigger reading immediately and keep button label as 'Select Area'
                is_auto_read = hasattr(area_name_var, 'get') and area_name_var.get() == "Auto Read"
                if is_auto_read:
                    select_area_window.destroy()
                    # Read after a short delay so overlay is gone
                    self.root.after(100, lambda: self.read_area(frame))
                else:
                    select_area_window.destroy()

                # --- Re-enable hotkeys after area selection completes ---
                try:
                    self.restore_all_hotkeys()
                except Exception as e:
                    print(f"Error restoring hotkeys after area selection: {e}")
                self.hotkeys_disabled_for_selection = False

                # Only prompt for name if it's the default or not set, and not Auto Read
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
                            return
                    else:
                        print(f"Updated area: {frame.area_coords} with existing name {current_name}\n--------------------------")
                    if set_area_button is not None:
                        set_area_button.config(text="Edit Area")
                else:
                    # Always keep button label as 'Select Area' for Auto Read
                    if set_area_button is not None:
                        set_area_button.config(text="Select Area")

        def on_escape(event):
            nonlocal selection_cancelled
            selection_cancelled = True
            if not hasattr(frame, 'area_coords'):
                frame.area_coords = (0, 0, 0, 0)
            select_area_window.destroy()
            # --- Re-enable hotkeys if selection is cancelled ---
            try:
                self.restore_all_hotkeys()
            except Exception as e:
                print(f"Error restoring hotkeys after cancelling area selection: {e}")
            self.hotkeys_disabled_for_selection = False
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
        
        # Add focus and key bindings
        select_area_window.focus_force()
        select_area_window.bind("<FocusOut>", lambda e: select_area_window.focus_force())
        select_area_window.bind("<Key>", lambda e: on_escape(e) if e.keysym == "Escape" else None)

    def disable_all_hotkeys(self):
        """Disable all hotkeys for keyboard and mouse."""
        try:
            import keyboard
            import mouse
            keyboard.unhook_all()
            mouse.unhook_all()
        except Exception as e:
            print(f"Error disabling all hotkeys: {e}")

    def restore_all_hotkeys(self):
        """Restore all area and stop hotkeys after area selection is finished/cancelled."""
        # Re-register all hotkeys for areas
        for area_tuple in getattr(self, 'areas', []):
            area_frame, hotkey_button, _, _, _, _, _ = area_tuple
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
        # --- Disable all hotkeys before starting hotkey assignment ---
        try:
            keyboard.unhook_all()
            mouse.unhook_all()
        except Exception as e:
            print(f"Error disabling hotkeys for hotkey assignment: {e}")
        self.hotkeys_disabled_for_hotkey = True
        self._hotkey_assignment_cancelled = False  # Guard flag to block late events

        def finish_hotkey_assignment():
            # --- Re-enable all hotkeys after hotkey assignment is finished/cancelled ---
            try:
                self.restore_all_hotkeys()
            except Exception as e:
                print(f"Error restoring hotkeys after hotkey assignment: {e}")
            self.hotkeys_disabled_for_hotkey = False

        def on_key_press(event):
            if self._hotkey_assignment_cancelled or not self.setting_hotkey:
                return
            key_name = event.name
            if event.scan_code in self.numpad_scan_codes:
                key_name = f"num_{self.numpad_scan_codes[event.scan_code]}"
            # Disallow Esc key as a hotkey
            if key_name.lower() in ("esc", "escape"):
                # Unhook temp hooks and set flags BEFORE showing the warning
                self.setting_hotkey = False
                self._hotkey_assignment_cancelled = True  # Block all further events
                if hasattr(button, 'keyboard_hook_temp'):
                    keyboard.unhook(button.keyboard_hook_temp)
                    delattr(button, 'keyboard_hook_temp')
                if hasattr(button, 'mouse_hook_temp'):
                    mouse.unhook(button.mouse_hook_temp)
                    delattr(button, 'mouse_hook_temp')
                finish_hotkey_assignment()
                if hasattr(button, 'hotkey'):
                    display_name = button.hotkey.replace('num_', 'num:') if button.hotkey.startswith('num_') else button.hotkey
                    button.config(text=f"Set Hotkey: [ {display_name} ]")
                else:
                    button.config(text="Set Hotkey")
                # Show the same custom dialog as for mouse button hotkey errors
                dialog = tk.Toplevel(self.root)
                dialog.title("Invalid Hotkey")
                dialog.geometry("300x120")
                dialog.transient(self.root)
                dialog.grab_set()
                # Center the dialog on the screen
                dialog.update_idletasks()
                x = self.root.winfo_rootx() + self.root.winfo_width()//2 - 150
                y = self.root.winfo_rooty() + self.root.winfo_height()//2 - 60
                dialog.geometry(f"300x120+{x}+{y}")
                # Add message
                msg = tk.Label(dialog, text="Escape (Esc) cannot be used as a hotkey.\nPlease set a different hotkey.", pady=20)
                msg.pack()
                # Add OK button
                ok_btn = tk.Button(dialog, text="OK", command=dialog.destroy, width=10)
                ok_btn.pack(pady=10)
                ok_btn.focus_set()
                dialog.bind("<Return>", lambda e: dialog.destroy())
                dialog.protocol("WM_DELETE_WINDOW", dialog.destroy)
                return
            # Remove existing hook for this button only
            if hasattr(button, 'keyboard_hook'):
                try:
                    keyboard.unhook(button.keyboard_hook)
                    delattr(button, 'keyboard_hook')
                except:
                    pass
            # Check for hotkey conflicts across ALL areas, including Auto Read
            for area_frame2, hotkey_button2, _, area_name_var2, _, _, _ in self.areas:
                # Do not allow duplicate hotkeys for any area (including Auto Read)
                if hotkey_button2 is not button and hasattr(hotkey_button2, 'hotkey') and hotkey_button2.hotkey == key_name:
                    # Unhook temp hooks and set flags BEFORE showing the warning
                    self.setting_hotkey = False
                    self._hotkey_assignment_cancelled = True  # Block all further events
                    if hasattr(button, 'keyboard_hook_temp'):
                        keyboard.unhook(button.keyboard_hook_temp)
                        delattr(button, 'keyboard_hook_temp')
                    if hasattr(button, 'mouse_hook_temp'):
                        mouse.unhook(button.mouse_hook_temp)
                        delattr(button, 'mouse_hook_temp')
                    finish_hotkey_assignment()
                    # Now show the warning dialog (no hooks are active)
                    if hasattr(button, 'hotkey'):
                        display_name = button.hotkey.replace('num_', 'num:') if button.hotkey.startswith('num_') else button.hotkey
                        button.config(text=f"Set Hotkey: [ {display_name} ]")
                    else:
                        button.config(text="Set Hotkey")
                    show_thinkr_warning(self, area_name_var2.get())
                    return
            button.hotkey = key_name
            display_name = key_name.replace('num_', 'num:') if key_name.startswith('num_') else key_name
            button.config(text=f"Set Hotkey: [ {display_name} ]")
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
            if self._hotkey_assignment_cancelled or not self.setting_hotkey:
                return
            if not hasattr(event, 'event_type') or event.event_type != mouse.DOWN:
                return
            if getattr(event, 'button', None) in [1, 2]:
                messagebox.showerror("Error", "Left and right mouse buttons cannot be used as hotkeys.")
                if hasattr(button, 'hotkey'):
                    display_name = button.hotkey.replace('num_', 'num:') if button.hotkey.startswith('num_') else button.hotkey
                    button.config(text=f"Set Hotkey: [ {display_name} ]")
                else:
                    button.config(text="Set Hotkey")
                # Unhook temp hooks and set flags BEFORE showing the warning
                self.setting_hotkey = False
                self._hotkey_assignment_cancelled = True  # Block all further events
                if hasattr(button, 'keyboard_hook_temp'):
                    keyboard.unhook(button.keyboard_hook_temp)
                    delattr(button, 'keyboard_hook_temp')
                if hasattr(button, 'mouse_hook_temp'):
                    mouse.unhook(button.mouse_hook_temp)
                    delattr(button, 'mouse_hook_temp')
                finish_hotkey_assignment()
                if hasattr(button, 'hotkey'):
                    display_name = button.hotkey.replace('num_', 'num:') if button.hotkey.startswith('num_') else button.hotkey
                    button.config(text=f"Set Hotkey: [ {display_name} ]")
                else:
                    button.config(text="Set Hotkey")
                messagebox.showerror("Error", "Left and right mouse buttons cannot be used as hotkeys.")
                return
            key_name = f'button{event.button}'
            if hasattr(button, 'mouse_hook'):
                try:
                    mouse.unhook(button.mouse_hook)
                    delattr(button, 'mouse_hook')
                except:
                    pass
            for area_frame2, hotkey_button2, _, area_name_var2, _, _, _ in self.areas:
                if hotkey_button2 is not button and hasattr(hotkey_button2, 'hotkey') and hotkey_button2.hotkey == key_name:
                    show_thinkr_warning(self, area_name_var2.get())
                    if hasattr(button, 'hotkey'):
                        display_name = button.hotkey.replace('num_', 'num:') if button.hotkey.startswith('num_') else button.hotkey
                        button.config(text=f"Set Hotkey: [ {display_name} ]")
                    else:
                        button.config(text="Set Hotkey")
                    finish_hotkey_assignment()
                    return
            button.hotkey = key_name
            button.config(text=f"Set Hotkey: [ {key_name} ]")
            self.setup_hotkey(button, area_frame)
            # Unhook both temp hooks if they exist
            if hasattr(button, 'keyboard_hook_temp'):
                keyboard.unhook(button.keyboard_hook_temp)
                delattr(button, 'keyboard_hook_temp')
            if hasattr(button, 'mouse_hook_temp'):
                mouse.unhook(button.mouse_hook_temp)
                delattr(button, 'mouse_hook_temp')
            finish_hotkey_assignment()
            return
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
        self.setting_hotkey = True  # Enable hotkey assignment mode before installing hooks
        button.keyboard_hook_temp = keyboard.on_press(on_key_press)
        button.mouse_hook_temp = mouse.hook(on_mouse_click)

    def save_layout(self):
        # Check if there are no areas
        if not self.areas:
            messagebox.showerror("Error", "There is nothing to save.")
            return

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
            "fullscreen_mode": self.fullscreen_mode_var.get(),
            "stop_hotkey": self.stop_hotkey,
            "volume": self.volume.get(),
            "areas": []
        }
        for area_frame, hotkey_button, _, area_name_var, preprocess_var, voice_var, speed_var in self.areas:
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

    def load_layout(self):
        file_path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if file_path:
            try:
                # Store the file path before loading
                self.layout_file.set(file_path)
                
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
                    display_name = saved_stop_hotkey.replace('num_', 'num:') if saved_stop_hotkey.startswith('num_') else saved_stop_hotkey
                    self.stop_hotkey_button.config(text=f"Stop Hotkey: [ {display_name} ]")
                    print(f"Loaded Stop hotkey: {saved_stop_hotkey}")

                # --- Check for Auto Read hotkey conflicts before registering ---
                auto_read_hotkey = None
                if self.areas and hasattr(self.areas[0][1], 'hotkey'):
                    auto_read_hotkey = self.areas[0][1].hotkey
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
                        hotkey_button.config(text=f"Hotkey: [ {display_name} ]")
                        self.setup_hotkey(hotkey_button, area_frame)
                    
                    # Set preprocessing and voice settings
                    preprocess_var.set(area_info.get("preprocess", False))
                    if area_info.get("voice") in [voice.name for voice in self.voices]:
                        voice_var.set(area_info["voice"])
                    speed_var.set(area_info.get("speed", "1.0"))
                    
                    # Load and store image processing settings
                    if "settings" in area_info:
                        self.processing_settings[area_info["name"]] = area_info["settings"].copy()
                        print(f"Loaded image processing settings for area: {area_info['name']}")
                        
                    # Update window size after loading each area
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
                # --- Show popup if conflict detected ---
                if conflict_area_name:
                    # Restore to messagebox.showwarning with requested text
                    hotkey_val = auto_read_hotkey if auto_read_hotkey else "?"
                    messagebox.showinfo(
                        "Hotkey Conflict",
                        f"Detected same Hotkey!\n\nAuto Read Hotkey = {hotkey_val}\n{conflict_area_name} Hotkey = {hotkey_val}\n\nPlease set a new hotkey for AutoRead if you still want this function.")
                    # Optionally: clear the Auto Read hotkey registration
                    if self.areas and hasattr(self.areas[0][1], 'hotkey'):
                        del self.areas[0][1].hotkey
                        self.areas[0][1].config(text="Set Hotkey")

                print(f"Layout loaded from {file_path}\n--------------------------")
                
                # Automatically resize the window
                self.resize_window()
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load layout: {str(e)}")
                print(f"Error loading layout: {e}")

    def validate_speed_key(self, event, speed_var):
        """Additional validation for speed entry key presses"""
        if event.char.isdigit() or event.keysym in ('BackSpace', 'Delete', 'Left', 'Right'):
            return
        return 'break'

    def setup_hotkey(self, button, area_frame):
        # Add this check at the start of setup_hotkey
        if hasattr(button, 'hotkey') and button.hotkey.startswith('button'):
            # Check for both number format and left/right text format
            button_id = button.hotkey.replace('button', '')
            if button_id in ['1', '2', 'left', 'right']:  # Check for both formats
                print("Prevented left/right mouse button from being set as hotkey")
                
                # Create custom dialog
                dialog = tk.Toplevel(self.root)
                dialog.title("Invalid Hotkey")
                dialog.geometry("300x120")
                dialog.transient(self.root)  # Make dialog modal
                dialog.grab_set()  # Make dialog modal
                
                # Center the dialog on the screen
                dialog.geometry("+%d+%d" % (
                    self.root.winfo_rootx() + self.root.winfo_width()/2 - 150,
                    self.root.winfo_rooty() + self.root.winfo_height()/2 - 60))
                
                # Add message
                message = tk.Label(dialog, 
                    text=f"Mouse button {button_id} cannot be used as a hotkey.\nPlease set a different hotkey.",
                    pady=20)
                message.pack()
                
                # Add OK button
                ok_button = tk.Button(dialog, text="OK", command=dialog.destroy, width=10)
                ok_button.pack(pady=10)
                
                return False
        
       # print(f"Setting up hotkey for {button.hotkey}...")

        # Clean up existing hooks for this specific button only
        if hasattr(button, 'mouse_hook'):
            try:
                mouse.unhook(button.mouse_hook)
                delattr(button, 'mouse_hook')
            except Exception as e:
                print(f"Error cleaning up mouse hook: {e}")

        if hasattr(button, 'keyboard_hook'):
            try:
                keyboard.unhook(button.keyboard_hook)
                delattr(button, 'keyboard_hook')
            except Exception as e:
                print(f"Error cleaning up keyboard hook: {e}")

        # Only store area_frame if this is NOT a stop button
        if not hasattr(button, 'is_stop_button'):
            button.area_frame = area_frame

        def on_hotkey(e):
            if self.setting_hotkey:
                return False

            # Only process keyboard events if it's not a mouse button hotkey
            if not button.hotkey.startswith('button'):
                # Handle numpad keys
                if button.hotkey.startswith('num_'):
                    key_name = button.hotkey.replace('num_', '')
                    if not (e.name == key_name and e.scan_code in self.numpad_scan_codes):
                        return False
                # Handle regular keys
                elif e.name != button.hotkey:
                    return False

                # Separate handling for stop button
                if hasattr(button, 'is_stop_button'):
                    self.root.after_idle(self.stop_speaking)
                    return True

                # Special handling for Auto Read area: simulate Select Area button press
                area_name = None
                for area in self.areas:
                    if area[1] is button:
                        area_name = area[3].get() if hasattr(area[3], 'get') else None
                        set_area_btn = area[2]
                        break
                if area_name == "Auto Read":
                    # Simulate Select Area button press (call set_area as the button would)
                    self.root.after_idle(lambda: self.set_area(area[0], area[3], set_area_btn))
                    return True

                # Only handle area reading for non-stop buttons (all other areas)
                if hasattr(button, '_is_processing') and button._is_processing:
                    return True
                button._is_processing = True
                self.stop_speaking()
                threading.Thread(target=self.read_area, args=(area_frame,), daemon=True).start()
                self.root.after(100, lambda: setattr(button, '_is_processing', False))
                return True

        def on_mouse_click(event):
            if self.setting_hotkey:
                return

            if not isinstance(event, mouse.ButtonEvent) or event.event_type != mouse.DOWN:
                return

            current_button = f'button{event.button}'
            if current_button != button.hotkey:
                return

            # print(f"Mouse button {current_button} pressed")

            # Separate handling for stop button
            if hasattr(button, 'is_stop_button'):
                # print(f"Stop hotkey mouse button '{button.hotkey}' pressed!")
                self.root.after_idle(self.stop_speaking)
                return
            
            # Only handle area reading for non-stop buttons
            else:
                if hasattr(button, '_is_processing') and button._is_processing:
                    return
                    
                button._is_processing = True
                self.stop_speaking()
                threading.Thread(target=self.read_area, args=(area_frame,), daemon=True).start()
                self.root.after(100, lambda: setattr(button, '_is_processing', False))

        # Set up appropriate hooks based on hotkey type
        if button.hotkey.startswith('button'):
            button.mouse_hook = mouse.hook(on_mouse_click)
         #   print(f"Mouse hook set up for {button.hotkey}")
        else:
            button.keyboard_hook = keyboard.on_press(on_hotkey)
          #  print(f"Keyboard hook set up for {button.hotkey}")

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
                
                # Apply bad word filtering while preserving order
                bad_words = [word.strip().lower() for word in self.bad_word_list.get().split(',')]
                filtered_words = [word for word in line.split() if word.lower() not in bad_words]
                
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
            # Get all available SAPI voices
            voices = self.speaker.GetVoices()
            selected_voice = None
            # Find the voice with matching name
            for voice in voices:
                if voice.GetDescription() == voice_var.get():
                    selected_voice = voice
                    break
            if selected_voice:
                self.speaker.Voice = selected_voice
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

            # Clean up hooks for each area
            for area_frame, hotkey_button, _, _, _, _, _ in self.areas:
                if hasattr(hotkey_button, 'keyboard_hook'):
                    try:
                        keyboard.unhook(hotkey_button.keyboard_hook)
                    except:
                        pass
                    
                if hasattr(hotkey_button, 'mouse_hook'):
                    try:
                        mouse.unhook(hotkey_button.mouse_hook)
                    except:
                        pass
            
            # Clean up images
            for image in self.latest_images.values():
                try:
                    image.close()
                except:
                    pass
            self.latest_images.clear()
            
            # Clean up text histories
            self.text_histories.clear()
            
            # Clean up speech - modified to check if speaker exists and is not None
            if hasattr(self, 'speaker') and self.speaker is not None:
                try:
                    self.speaker.Speak("", 2)  # Stop any ongoing speech
                except:
                    pass
                self.speaker = None
            
            # Clean up engine
            if hasattr(self, 'engine'):
                try:
                    self.engine.stop()
                except:
                    pass
            
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
    import win32gui
    import win32ui
    import win32con
    from PIL import Image
    
    # Get DC from entire virtual screen
    hwin = win32gui.GetDesktopWindow()
    hwindc = win32gui.GetWindowDC(hwin)
    srcdc = win32ui.CreateDCFromHandle(hwindc)
    memdc = srcdc.CreateCompatibleDC()
    
    # Create bitmap for entire capture area
    width = x2 - x1
    height = y2 - y1
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
        bmpstr, 'raw', 'BGRX', 0, 1)
    
    # Clean up
    memdc.DeleteDC()
    win32gui.ReleaseDC(hwin, hwindc)
    win32gui.DeleteObject(bmp.GetHandle())
    
    return img

if __name__ == "__main__":
    import tempfile, os, json
    root = tk.Tk()
    app = GameTextReader(root)
    # Create permanent area at the top
    app.add_read_area(removable=False, editable_name=False, area_name="Auto Read")
    # Try to load settings for Auto Read area from temp folder
    temp_path = os.path.join(tempfile.gettempdir(), 'auto_read_settings.json')
    if os.path.exists(temp_path) and app.areas:
        try:
            with open(temp_path, 'r') as f:
                settings = json.load(f)
            # Find the permanent area and set its settings
            area_frame, hotkey_button, set_area_button, area_name_var, preprocess_var, voice_var, speed_var = app.areas[0]
            preprocess_var.set(settings.get('preprocess', False))
            voice_var.set(settings.get('voice', 'Select Voice'))
            speed_var.set(settings.get('speed', '100'))
            # Restore image processing settings
            brightness = settings.get('brightness', 1.0)
            contrast = settings.get('contrast', 1.0)
            saturation = settings.get('saturation', 1.0)
            sharpness = settings.get('sharpness', 1.0)
            blur = settings.get('blur', 0.0)
            hue = settings.get('hue', 0.0)
            exposure = settings.get('exposure', 1.0)
            threshold = settings.get('threshold', 128)
            threshold_enabled = settings.get('threshold_enabled', False)
            # Set these in the processing_settings dict
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
            }
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
            # Restore hotkey if present
            hotkey = settings.get('hotkey')
            if hotkey:
                hotkey_button.hotkey = hotkey
                display_name = hotkey.replace('num_', 'num:') if hotkey.startswith('num_') else hotkey
                hotkey_button.config(text=f"Hotkey: [ {display_name} ]")
                app.setup_hotkey(hotkey_button, area_frame)
            # Restore 'Stop read on select' checkbox if present
            if hasattr(app, 'interrupt_on_new_scan_var'):
                app.interrupt_on_new_scan_var.set(settings.get('stop_read_on_select', True))
        except Exception as e:
            print(f"Failed to load Auto Read settings: {e}")
    root.mainloop()
