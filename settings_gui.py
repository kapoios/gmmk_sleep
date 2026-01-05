import json
import os
import tkinter as tk
from tkinter import ttk, messagebox
import ctypes


class SettingsWindow:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("GMMK Sleep Settings")
        
        # Get screen dimensions
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        
        # Calculate window size as percentage of screen (with min/max limits)
        # 20% width, 25% height with higher max for 4K/8K displays
        window_width = min(1200, max(500, int(screen_width * 0.20)))
        window_height = min(900, max(400, int(screen_height * 0.25)))
        
        # Center the window on screen
        x_position = (screen_width - window_width) // 2
        y_position = (screen_height - window_height) // 2
        
        self.window.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")
        self.window.resizable(True, True)
        self.window.minsize(500, 350)
        
        # Load the hidapi.dll from the project directory
        dll_path = os.path.join(os.path.dirname(__file__), "hidapi.dll")
        ctypes.CDLL(dll_path)
        
        # Import hid after loading the DLL
        import hid
        self.hid = hid
        
        self.settings_path = os.path.join(os.path.dirname(__file__), "settings.json")
        self.devices = []
        
        self.create_widgets()
        self.current_vid = None
        self.current_pid = None
        self.load_current_settings()
        self.load_devices()
        
    def create_widgets(self):
        # Title
        title_label = tk.Label(
            self.window, 
            text="Select USB HID Device", 
            font=("Arial", 14, "bold")
        )
        title_label.pack(pady=10)
        
        # Instructions
        instructions = tk.Label(
            self.window,
            text="Choose the HID device that corresponds to your GMMK keyboard:",
            wraplength=550
        )
        instructions.pack(pady=5)
        
        # Frame for device list
        list_frame = tk.Frame(self.window)
        list_frame.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)
        
        # Scrollbar
        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Listbox for devices
        self.device_listbox = tk.Listbox(
            list_frame,
            yscrollcommand=scrollbar.set,
            font=("Courier", 9),
            selectmode=tk.SINGLE
        )
        self.device_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.device_listbox.yview)
        
        # Buttons frame
        button_frame = tk.Frame(self.window)
        button_frame.pack(pady=10)
        
        # Refresh button
        refresh_btn = tk.Button(
            button_frame,
            text="Refresh Devices",
            command=self.load_devices,
            width=15
        )
        refresh_btn.pack(side=tk.LEFT, padx=5)
        
        # Save button
        save_btn = tk.Button(
            button_frame,
            text="Save",
            command=self.save_settings,
            width=15,
            bg="#4CAF50",
            fg="white"
        )
        save_btn.pack(side=tk.LEFT, padx=5)
        
        # Cancel button
        cancel_btn = tk.Button(
            button_frame,
            text="Cancel",
            command=self.window.destroy,
            width=15
        )
        cancel_btn.pack(side=tk.LEFT, padx=5)
        
        # Current settings label
        self.current_label = tk.Label(
            self.window,
            text="",
            font=("Arial", 9),
            fg="blue"
        )
        self.current_label.pack(pady=5)
        
    def load_devices(self):
        """Enumerate all HID devices and display them (keyboards only, no duplicates)"""
        self.device_listbox.delete(0, tk.END)
        self.devices = []
        seen_devices = set()
        
        try:
            all_devices = self.hid.enumerate()
            
            for device in all_devices:
                vendor_id = device.get('vendor_id', 0)
                product_id = device.get('product_id', 0)
                manufacturer = device.get('manufacturer_string', 'Unknown')
                product = device.get('product_string', 'Unknown')
                interface = device.get('interface_number', -1)
                usage_page = device.get('usage_page', 0)
                usage = device.get('usage', 0)
                
                # Filter for keyboards only - be strict
                # Only accept devices with standard keyboard usage OR vendor-specific with "keyboard" in name
                is_keyboard = (usage_page == 0x01 and usage == 0x06)  # Standard keyboard usage
                
                # For vendor-specific pages, require "keyboard" in the product name
                if usage_page == 0xFF01 and 'keyboard' in product.lower():
                    is_keyboard = True
                
                if not is_keyboard:
                    continue
                
                # Create unique identifier to show only one entry per keyboard (ignore interface)
                device_key = (vendor_id, product_id)
                
                if device_key in seen_devices:
                    continue
                
                seen_devices.add(device_key)
                
                # Create display string
                display_str = f"VID: 0x{vendor_id:04X} | PID: 0x{product_id:04X} | {manufacturer} - {product}"
                
                self.device_listbox.insert(tk.END, display_str)
                self.devices.append(device)
            
            if not self.devices:
                self.device_listbox.insert(tk.END, "No keyboard devices found")
            else:
                # Highlight the currently configured device if it exists
                self.highlight_current_device()
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to enumerate devices: {e}")
    
    def load_current_settings(self):
        """Load and display current settings"""
        try:
            with open(self.settings_path, 'r') as f:
                settings = json.load(f)
            
            vendor_id = settings.get('VENDOR_ID', '0x0000')
            product_id = settings.get('PRODUCT_ID', '0x0000')
            
            # Store current VID/PID for highlighting
            self.current_vid = int(vendor_id, 16)
            self.current_pid = int(product_id, 16)
            
            self.current_label.config(
                text=f"Current Settings: VID={vendor_id}, PID={product_id}"
            )
                    
        except Exception as e:
            print(f"Could not load current settings: {e}")
    
    def highlight_current_device(self):
        """Highlight the currently configured device in the list"""
        if self.current_vid is None or self.current_pid is None:
            return
        
        for i, device in enumerate(self.devices):
            vid = device.get('vendor_id', 0)
            pid = device.get('product_id', 0)
            
            if vid == self.current_vid and pid == self.current_pid:
                self.device_listbox.selection_clear(0, tk.END)
                self.device_listbox.selection_set(i)
                self.device_listbox.see(i)
                self.device_listbox.itemconfig(i, bg='lightblue')
                break
    
    def save_settings(self):
        """Save selected device to settings.json"""
        selection = self.device_listbox.curselection()
        
        if not selection:
            messagebox.showwarning("No Selection", "Please select a device first")
            return
        
        selected_index = selection[0]
        
        if selected_index >= len(self.devices):
            messagebox.showerror("Error", "Invalid device selection")
            return
        
        device = self.devices[selected_index]
        vendor_id = device.get('vendor_id', 0)
        product_id = device.get('product_id', 0)
        
        # Confirm with user
        confirm = messagebox.askyesno(
            "Confirm Save",
            f"Save the following settings?\n\n"
            f"Vendor ID: 0x{vendor_id:04X}\n"
            f"Product ID: 0x{product_id:04X}\n"
            f"Device: {device.get('manufacturer_string', 'Unknown')} - {device.get('product_string', 'Unknown')}\n\n"
            f"Note: You may need to restart the application for changes to take effect."
        )
        
        if not confirm:
            return
        
        try:
            settings = {
                "VENDOR_ID": f"0x{vendor_id:04X}",
                "PRODUCT_ID": f"0x{product_id:04X}"
            }
            
            with open(self.settings_path, 'w') as f:
                json.dump(settings, f, indent=2)
            
            messagebox.showinfo(
                "Success",
                "Settings saved successfully!\n\n"
                "Please restart the application for changes to take effect."
            )
            self.window.destroy()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save settings: {e}")
    
    def run(self):
        """Run the settings window"""
        self.window.mainloop()


def open_settings():
    """Function to open the settings window"""
    app = SettingsWindow()
    app.run()


if __name__ == "__main__":
    open_settings()
