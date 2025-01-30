# pip install pyautogui pygetwindow

import pyautogui as pag
import time
import os
import psutil
import pygetwindow as gw


def automate_excel_task():
    # Open Excel
    print("Opening Excel...")
    os.startfile("C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\Excel.lnk")  # Replace with your Excel shortcut path
    time.sleep(5)  # Wait for Excel to open

    # Make sure that excel is open
    excel_windows = [w for w in gw.getAllWindows() if "Excel" in w.title]
    while not excel_windows:
        time.sleep(1)
        excel_windows = [w for w in gw.getAllWindows() if "Excel" in w.title]

    # Ensure Excel is active before performing actions
    ensure_excel_active()

    # Open a blank workbook
    print("Opening a blank workbook...")
    pag.hotkey('enter')

    # Ensure Excel is active before performing actions
    ensure_excel_active()

    # Navigate to the Plug-ins tab
    print("Navigating to the Plug-ins tab...")
    pag.hotkey('alt', 'x')  

    # Ensure Excel is active before performing actions
    ensure_excel_active()

    # Click on the UPush button
    print("Clicking on the UPush button...")
    pag.hotkey('y','1')
    pag.hotkey('enter')

    #wait until an excel window with title containing "upush_template_physicalasset2013" is found
    print("Waiting for UPush window to load...")
    upush_window = None
    while not upush_window:
        upush_window = gw.getWindowsWithTitle("upush_template_physicalasset2013")
        if not upush_window:
            time.sleep(1)

    time.sleep(5)  # Wait for the login window to load

    #login popup window may appear here, handle it if needed
    #if window title starts with "Sign in to your account", handle it
    print("Handling login popup window...")
    login_window = None
    while not login_window:
        login_window = gw.getWindowsWithTitle("Sign in to your account")
        if not login_window:
            time.sleep(1)

    
    active_window = gw.getActiveWindow()

    if login_window[0].title != active_window.title:
        login_window[0].activate()

    # Press tab then enter to select first account
    #TODO: handle multiple accounts or no accounts - maybe a pop up to pause here and let user select account?
    pag.hotkey('tab')
    pag.hotkey('enter')


    #TODO: possibly find the best time to close the unnecessary excel window by cpu usage since UPush takes a minute to load
    process_name = "upush_template_physicalasset2013 - Excel"
    process = get_process_by_window_title(process_name)
    if process:
        process_name = process.name()
        print(f"Process found: {process_name}")
    
    return

    #TODO:Close the original Excel window
    #close_original_excel_window()

    print("Task completed!")


def ensure_excel_active():
    excel_windows = [w for w in gw.getAllWindows() if "Excel" in w.title]
    if not excel_windows:
        raise Exception("Excel window not found")
    
    if excel_windows[0].isMinimized:
        excel_windows[0].restore()
        print(f"Restored: {excel_windows[0].title}")

    active_window = gw.getActiveWindow()
    if "Excel" not in active_window.title:
        excel_windows[0].activate()
        print(f"Activated: {excel_windows[0].title}")


def get_process_by_window_title(title):
    for proc in psutil.process_iter(['pid', 'name']):
        try:
            if title in proc.name():  # Match process name
                return proc
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            print("Could not access process")
            continue
    return None


def close_original_excel_window():
    print("Closing original Excel window...")

    # Wait 2 minutes for UPush to finish loading
    time.sleep(120)

    # Look for all open Excel windows
    excel_windows = [w for w in gw.getAllWindows() if "Excel" in w.title]

    # Find the excel window that has title starting with "Book"
    excel_windows = [w for w in excel_windows if w.title.startswith("Book")]

    # Close that window
    book_window = excel_windows[0]
    print(f"Closing: {book_window.title}")
    book_window.close()

if __name__ == "__main__":
    automate_excel_task()

