# GMMK Sleep!
Turn off the GMMK2 lights when your display turns off.

My GMMK 2 keyboard would not turn off when my displays were off so I created this.
It has only been tested on GMMK 2 Full (96%) UK ISO, using Windows 11.
It may work on other GMMK keyboards assuming the protocol is the same.

**This was done in an afternoon to cover my use case, I share because why not? Feel free to fork if you want different functionality**

# How to use
Create two profiles in Glorious Core, one with your RGB on and a second with the RGB off.

**Download the bundled zip from the releases tab, extract somewhere**

Make sure the VID and PID match yours in the settings.json file, update if not.

**Run the .exe**

If you want to exit it, you can right click the tray icon.

# How it works:
It checks the display timeout from the current power plan and then checks if the system is IDLE for the same amount of time.
If no screen timeout is set, it defaults to 15 minutes.
It switches the profile to profile 2 if the system is IDLE for that long, switches back to profile 1 when the system stops being IDLE.

# How to run locally:
Clone this repo, add hidapi.dll from [HERE](https://github.com/libusb/hidapi/releases) to the root of the script.

Run with:
```
python main.py
```

Packaged with pyinstaller, I used the following command:
```
pyinstaller --noconsole --exclude-module numpy --add-binary hidapi.dll:. --add-data icon.png:. --add-data settings.json:. main.py
```
