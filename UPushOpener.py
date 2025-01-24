# pip install pyautogui pygetwindow

import pyautogui as pag
import time
import os
import pygetwindow as gw

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


def automate_excel_task():
    # Open Excel
    os.startfile("C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\Excel.lnk")  # Replace with your Excel shortcut path
    time.sleep(5)  # Wait for Excel to open

    # Open a blank workbook
    pag.hotkey('enter')

    # Navigate to the Plug-ins tab
    pag.hotkey('alt', 'x')  

    # Click on the UPush button
    pag.hotkey('y','1')
    #pag.hotkey('1')

    pag.hotkey('enter')

    #wait until an excel window with title containing "upush_template_physicalasset2013" is found
    upush_window = None
    while not upush_window:
        upush_window = gw.getWindowsWithTitle("upush_template_physicalasset2013")
        if not upush_window:
            time.sleep(1)

    time.sleep(5)  # Wait for the login window to load

    #login popup window may appear here, handle it if needed
    #if window title starts with "Sign in to your account", handle it
    login_window = None
    while not login_window:
        login_window = gw.getWindowsWithTitle("Sign in to your account")
        if not login_window:
            time.sleep(1)

    
    active_window = gw.getActiveWindow()

    if login_window[0].title != active_window.title:
        login_window[0].activate()

    # Press tab then enter to select first account
    pag.hotkey('tab')
    pag.hotkey('enter')

    # Close the original Excel window
    #close_original_excel_window()

    print("Task completed!")

if __name__ == "__main__":
    automate_excel_task()

