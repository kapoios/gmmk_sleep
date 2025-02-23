import ctypes
import json
import os
import sys
import threading
import time
import winreg

import pystray
from PIL import Image

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


def create_tray_icon():
    icon_image = Image.open(os.path.join(os.path.dirname(__file__), "icon.png"))
    menu = pystray.Menu(pystray.MenuItem("Exit", on_exit))
    icon = pystray.Icon("gmmk_sleep", icon_image, "GMMK Sleep!", menu=menu)
    icon.run()


def get_display_timeout():
    # Open the main power schemes key
    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                        r"System\CurrentControlSet\Control\Power\User\PowerSchemes") as power_key:
        # Read the active power scheme GUID
        active_scheme, _ = winreg.QueryValueEx(power_key, "ActivePowerScheme")

    # Combine paths for the display timeout setting (AC)
    subkey_path = (
        f"System\\CurrentControlSet\\Control\\Power\\User\\PowerSchemes\\{active_scheme}\\"
        "7516b95f-f776-4464-8c53-06167f40cc99\\3c0bc021-c8a8-4e07-a973-6b14cbcb2b7e"
    )

    # Read the ACSettingIndex as the display timeout (in seconds)
    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, subkey_path) as display_key:
        ac_timeout, _ = winreg.QueryValueEx(display_key, "ACSettingIndex")

    if ac_timeout:
        print("Display timeout: " + str(ac_timeout) + " seconds")
    else:
        print("Display timeout not found, using 15 minutes as default")
        ac_timeout = 15 * 60
    # Convert seconds to milliseconds
    return ac_timeout * 1000


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

    current_tick = kernel32.GetTickCount()
    user32.GetLastInputInfo(ctypes.byref(last_input))
    idle_time = current_tick - last_input.dwTime

    return idle_time < timeout


def find_device_path():
    devices = hid.enumerate(VENDOR_ID, PRODUCT_ID)
    for device in devices:
        print(
            f"Found device: Interface={device.get('interface_number', -1)}, Usage Page={device.get('usage_page', 0):04x}")
        if (device['interface_number'] == INTERFACE and
                device['usage_page'] == 0xFF01):
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
        raise Exception(f"Device with interface {INTERFACE} not found")

    try:
        device = hid.Device(path=device_path)
        device.send_feature_report(bytes(report))
        print("Feature report sent successfully")

        device.close()

    except Exception as e:
        raise Exception(f"HID error: {e}")


def main_loop():
    display_timeout = get_display_timeout()
    last_state = is_system_active(display_timeout)
    print(f"System state: {'ACTIVE' if last_state else 'IDLE'}")
    while not stop_event.is_set():
        current_state = is_system_active(display_timeout)
        print(f"System state: {'ACTIVE' if current_state else 'IDLE'}")

        if current_state != last_state:
            print("System activity changed, updating keyboard lighting...")
            report = [0x07, 0x01, 0x01, 0x01] if current_state else [0x07, 0x01, 0x02, 0x01]
            report += [0x00] * (REPORT_LENGTH - 4)
            send_report(report)
            last_state = current_state
            print("Keyboard lighting updated successfully")
        else:
            print("No change in system activity")
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
