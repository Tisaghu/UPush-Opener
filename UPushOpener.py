import pyautogui as pag
import time
import os
import psutil
import pygetwindow as gw
import win32com.client

# Configurable Constants
EXCEL_SHORTCUT_PATH = "C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\Excel.lnk"
UPUSH_WINDOW_TITLE = "upush_template_physicalasset2013"
LOGIN_WINDOW_TITLE = "Sign in to your account"
DEFAULT_WAIT_TIME = 5
MAX_WAIT_TIME = 60  # Timeout for waiting loops


def automate_excel_task():
    try:
        open_excel()

        time.sleep(5)  # Wait for Excel to load

        print("Opening a blank workbook...")
        ensure_window_active("Excel")
        pag.hotkey('enter')

        navigate_to_plugins()
        click_upush_button()
        
        wait_for_window(UPUSH_WINDOW_TITLE)
        handle_login_window()

        close_initial_window()

        print("Task completed!")
    except Exception as e:
        print(f"An error occurred: {e}")



def close_initial_window():
    print("Closing the initial window...")

    try:
        excel = win32com.client.Dispatch("Excel.Application")
        for wb in excel.Workbooks:
            if wb.Name == "Book1.xlsx" or wb.Name.startswith("Book1"):  # Adjust if necessary
                wb.Close(SaveChanges=False)
                print("Closed the blank workbook.")
                return
    except Exception as e:
        print(f"Error closing initial window: {e}")


def open_excel():
    print("Opening Excel...")
    os.startfile(EXCEL_SHORTCUT_PATH)
    wait_for_window("Excel")


def ensure_window_active(window_title):
    windows = [w for w in gw.getAllWindows() if window_title in w.title]
    if not windows:
        raise Exception(f"{window_title} window not found")

    window = windows[0]
    if window.isMinimized:
        window.restore()
        print(f"Restored: {window.title}")
    if gw.getActiveWindow() != window:
        window.activate()
        print(f"Activated: {window.title}")


def navigate_to_plugins():
    ensure_window_active("Excel")
    print("Navigating to the Plug-ins tab...")
    pag.hotkey('alt', 'x')


def click_upush_button():
    ensure_window_active("Excel")
    print("Clicking on the UPush button...")
    pag.hotkey('y', '1')
    pag.hotkey('enter')


def wait_for_window(title, timeout=MAX_WAIT_TIME):
    print(f"Waiting for window with title '{title}'...")
    start_time = time.time()
    while time.time() - start_time < timeout:
        windows = gw.getWindowsWithTitle(title)
        if windows:
            return windows[0]
        time.sleep(1)
    raise TimeoutError(f"Window '{title}' did not appear within {timeout} seconds.")


def handle_login_window():
    try:
        login_window = wait_for_window(LOGIN_WINDOW_TITLE, timeout=DEFAULT_WAIT_TIME)
        time.sleep(5)  # Wait for the window to load
        login_window.activate()
        print("Handling login window...")
        pag.hotkey('tab')
        pag.hotkey('enter')
    except TimeoutError:
        print("Login window did not appear, proceeding...")


if __name__ == "__main__":
    automate_excel_task()
