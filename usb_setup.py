import os
import time
import datetime
import ctypes
from pathlib import Path

import serial.tools.list_ports
try:
    import pyautogui
except Exception:
    pyautogui = None

LOG_DIR = Path(r"C:\SetupLogs")
LOG_FILE = LOG_DIR / "setup.log"


def log(message: str) -> None:
    """Write a timestamped message to the log file."""
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with LOG_FILE.open("a", encoding="utf-8") as f:
        f.write(f"{timestamp} - {message}\n")


def show_popup(title: str, text: str) -> None:
    """Display a Windows message box."""
    ctypes.windll.user32.MessageBoxW(0, text, title, 0)


def take_screenshot(path: Path) -> None:
    """Save a screenshot to the specified path."""
    if pyautogui is None:
        log("pyautogui not available; screenshot skipped")
        return
    try:
        image = pyautogui.screenshot()
        image.save(str(path))
        log(f"Screenshot saved to {path}")
    except Exception as e:
        log(f"Failed to take screenshot: {e}")


def get_connected_ports():
    """Return a dict mapping COM port to hardware id."""
    ports = {}
    for p in serial.tools.list_ports.comports():
        ports[p.device] = p.hwid
    return ports


def wait_for_new_port(existing_ports):
    """Wait until a new COM port appears and return (port, hwid)."""
    log("Waiting for new COM port...")
    while True:
        ports = get_connected_ports()
        for device, hwid in ports.items():
            if device not in existing_ports:
                return device, hwid
        time.sleep(1)


def setup_device(name: str, expected_port: str, existing_ports):
    """Prompt the user to connect a device and detect its COM port."""
    show_popup("Device Setup", f"Please plug in {name} now.")
    log(f"Prompted user to connect {name}")
    port, hwid = wait_for_new_port(existing_ports)
    log(f"Detected {name} on {port} (HWID: {hwid})")
    screenshot_name = (
        f"{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}_"
        f"{name.replace(' ', '_')}.png"
    )
    screenshot_path = LOG_DIR / screenshot_name
    take_screenshot(screenshot_path)
    show_popup("Device Detected", f"{name} detected on {port}\nHWID: {hwid}")
    existing_ports.add(port)
    return port


def main():
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    log("==== Starting USB Device Setup ====")
    existing_ports = set(get_connected_ports().keys())

    devices = [
        ("SICK Hand Scanner", "COM7"),
        ("Boarding Pass Barcode Scanner", "COM8"),
    ]

    detected = {}
    try:
        for name, expected in devices:
            port = setup_device(name, expected, existing_ports)
            detected[name] = port
        show_popup("Setup Complete", "All devices have been detected and logged.")
        log("Setup completed successfully.")
    except Exception as e:
        log(f"Error during setup: {e}")
        show_popup("Setup Error", str(e))


if __name__ == "__main__":
    main()
