import ctypes
from ctypes import wintypes
import json
import os
import re
import subprocess
import sys
import threading
import time

import pystray
from PIL import Image
import settings_gui

# Define GUID structure for Power APIs
class GUID(ctypes.Structure):
    _fields_ = [
        ("Data1", ctypes.c_ulong),
        ("Data2", ctypes.c_ushort),
        ("Data3", ctypes.c_ushort),
        ("Data4", ctypes.c_ubyte * 8),
    ]

# GUID_VIDEO_SUBGROUP: {7516b95f-f776-4464-8c53-06167f40cc99}
GUID_VIDEO_SUBGROUP = GUID(0x7516b95f, 0xf776, 0x4464, (ctypes.c_ubyte * 8)(0x8c, 0x53, 0x06, 0x16, 0x7f, 0x40, 0xcc, 0x99))
# GUID_VIDEO_POWERDOWN_TIMEOUT: {3c0bc021-c8a8-4e07-a973-6b14cbcb2b7e}
GUID_VIDEO_POWERDOWN_TIMEOUT = GUID(0x3c0bc021, 0xc8a8, 0x4e07, (ctypes.c_ubyte * 8)(0xa9, 0x73, 0x6b, 0x14, 0xcb, 0xcb, 0x2b, 0x7e))

# For Power Status
class SYSTEM_POWER_STATUS(ctypes.Structure):
    _fields_ = [
        ('ACLineStatus', ctypes.c_byte),
        ('BatteryFlag', ctypes.c_byte),
        ('BatteryLifePercent', ctypes.c_byte),
        ('Reserved1', ctypes.c_byte),
        ('BatteryLifeTime', ctypes.c_uint32),
        ('BatteryFullLifeTime', ctypes.c_uint32),
    ]

# Load the hidapi.dll from the project directory
dll_path = os.path.join(os.path.dirname(__file__), "hidapi.dll")
ctypes.CDLL(dll_path)

import hid


# Load VENDOR_ID and PRODUCT_ID from settings.json
def load_ids():
    settings_path = os.path.join(os.path.dirname(__file__), "settings.json")
    with open(settings_path, 'r') as f:
        data = json.load(f)
    vendor_id = int(data["VENDOR_ID"], 16)
    product_id = int(data["PRODUCT_ID"], 16)
    return vendor_id, product_id


VENDOR_ID, PRODUCT_ID = load_ids()
INTERFACE = 2  # Interface from Wireshark
REPORT_LENGTH = 256  # wLength from Wireshark

stop_event = threading.Event()


def on_exit(icon):
    icon.stop()
    stop_event.set()


def on_settings(icon):
    """Open the settings GUI in a separate thread"""
    settings_thread = threading.Thread(target=settings_gui.open_settings, daemon=True)
    settings_thread.start()


def create_tray_icon():
    icon_image = Image.open(os.path.join(os.path.dirname(__file__), "icon.png"))
    menu = pystray.Menu(
        pystray.MenuItem("Settings", on_settings),
        pystray.MenuItem("Exit", on_exit)
    )
    icon = pystray.Icon("gmmk_sleep", icon_image, "GMMK Sleep!", menu=menu)
    icon.run()


def get_display_timeout():
    """Retrieve the display timeout (in seconds) using official Windows Power APIs."""
    timeout = None
    try:
        powrprof = ctypes.windll.powrprof
        kernel32 = ctypes.windll.kernel32
        
        # Check power status
        status = SYSTEM_POWER_STATUS()
        kernel32.GetSystemPowerStatus(ctypes.byref(status))
        on_ac = status.ACLineStatus != 0
        
        active_guid_ptr = ctypes.POINTER(GUID)()
        if powrprof.PowerGetActiveScheme(None, ctypes.byref(active_guid_ptr)) == 0:
            try:
                timeout_val = wintypes.DWORD()
                read_func = powrprof.PowerReadACValueIndex if on_ac else powrprof.PowerReadDCValueIndex
                if read_func(None, active_guid_ptr, ctypes.byref(GUID_VIDEO_SUBGROUP),
                           ctypes.byref(GUID_VIDEO_POWERDOWN_TIMEOUT), ctypes.byref(timeout_val)) == 0:
                    timeout = timeout_val.value
            finally:
                kernel32.LocalFree(active_guid_ptr)
    except Exception:
        print(f"Registry access failed ({e}), attempting powercfg fallback...")
        try:
            # Fallback to powercfg
            cmd = "powercfg /query SCHEME_CURRENT SUB_VIDEO VIDEOIDLE"
            result = subprocess.run(cmd, capture_output=True, text=True, check=True, 
                                   creationflags=subprocess.CREATE_NO_WINDOW)
            
            status = SYSTEM_POWER_STATUS()
            ctypes.windll.kernel32.GetSystemPowerStatus(ctypes.byref(status))
            on_ac = status.ACLineStatus != 0
            
            matches = dict(re.findall(r"(AC|DC) Setting Index: (0x[0-9a-fA-F]+)", result.stdout))
            val = matches.get("AC" if on_ac else "DC")
            if val:
                timeout = int(val, 16)
        except Exception:
            print(f"powercfg failed too...")
            pass
    if timeout is not None:
        print("Display timeout: " + str(timeout) + " seconds")
    else:
        print("Display timeout not found, using 15 minutes as default")
        timeout = 15 * 60
    # Convert seconds to milliseconds
    return timeout * 1000


def is_system_active(timeout):
    class LASTINPUTINFO(ctypes.Structure):
        _fields_ = [
            ('cbSize', ctypes.c_uint),
            ('dwTime', ctypes.c_uint)
        ]

    last_input = LASTINPUTINFO()
    last_input.cbSize = ctypes.sizeof(LASTINPUTINFO)

    kernel32 = ctypes.windll.kernel32
    user32 = ctypes.windll.user32
    
    # Specify return type as unsigned 32-bit to handle values > 2^31 (after ~24.8 days uptime)
    kernel32.GetTickCount.restype = ctypes.c_uint

    current_tick = kernel32.GetTickCount()
    user32.GetLastInputInfo(ctypes.byref(last_input))
    idle_time = current_tick - last_input.dwTime
    
    #print(f"Debug: idle_time={idle_time}ms, timeout={timeout}ms, idle_time < timeout = {idle_time < timeout}")
    
    return idle_time < timeout


def find_device_path():
    devices = hid.enumerate(VENDOR_ID, PRODUCT_ID)
    for device in devices:
        if (device['interface_number'] == INTERFACE and
                device['usage_page'] == 0xFF01):
            print(
                f"Found device: Interface={device.get('interface_number', -1)}, Usage Page={device.get('usage_page', 0):04x}")
            return device['path']
    return None


def send_report(report):
    device_path = find_device_path()
    if not device_path:
        # Try without usage page check as fallback
        devices = hid.enumerate(VENDOR_ID, PRODUCT_ID)
        for device in devices:
            if device['interface_number'] == INTERFACE:
                device_path = device['path']
                break

    if not device_path:
        print(f"Warning: Device with interface {INTERFACE} not found - keyboard may be disconnected")
        return False

    try:
        device = hid.Device(path=device_path)
        device.send_feature_report(bytes(report))
        print("Feature report sent successfully")
        device.close()
        return True

    except Exception as e:
        print(f"Warning: Could not send report to device - {e}")
        return False


def main_loop():
    display_timeout = get_display_timeout()
    last_state = is_system_active(display_timeout)
    device_connected = True
    reconnect_attempts = 0
    
    print(f"System state: {'ACTIVE' if last_state else 'IDLE'}")
    find_device_path()
    
    while not stop_event.is_set():
        current_state = is_system_active(display_timeout)
        print(f"System state: {'ACTIVE' if current_state else 'IDLE'}")

        if current_state != last_state:
            print("System activity changed, updating keyboard lighting...")
            report = [0x07, 0x01, 0x01, 0x01] if current_state else [0x07, 0x01, 0x02, 0x01]
            report += [0x00] * (REPORT_LENGTH - 4)
            
            success = send_report(report)
            if success:
                last_state = current_state
                print("Keyboard lighting updated successfully")
                device_connected = True
                reconnect_attempts = 0
            else:
                if device_connected:
                    print("Device disconnected - will retry when reconnected")
                    device_connected = False
                reconnect_attempts += 1
                
                # Try to reconnect less frequently after multiple failures
                if reconnect_attempts > 10:
                    time.sleep(5)  # Wait longer between reconnection attempts
        else:
            print("No change in system activity")
            
            # If device was disconnected, periodically try to reconnect
            if not device_connected and reconnect_attempts % 5 == 0:
                print("Attempting to reconnect to device...")
                # Try sending current state to check if device is back
                report = [0x07, 0x01, 0x01, 0x01] if current_state else [0x07, 0x01, 0x02, 0x01]
                report += [0x00] * (REPORT_LENGTH - 4)
                
                if send_report(report):
                    print("Device reconnected successfully!")
                    device_connected = True
                    reconnect_attempts = 0
                    last_state = current_state
                    
        time.sleep(2)
    sys.exit(0)


def run_script():
    # Start the tray icon in a separate thread
    tray_thread = threading.Thread(target=create_tray_icon, daemon=True)
    tray_thread.start()
    main_loop()


if __name__ == "__main__":
    try:
        run_script()
    except Exception as e:
        print(f"Error: {e}")
