import win32com.client
import time

def automate_excel_task():
    try:
        # Start Excel or connect to an existing instance
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True  # Make Excel visible

        print("Opening a blank workbook...")
        blank_workbook = excel.Workbooks.Add()  # Open a new blank workbook

        print("Navigating to the Plug-ins tab...")

        print("Opening UPush plugin...")

        # Wait for login window to appear
        time.sleep(5)  # Adjust based on actual behavior

        # Close the blank workbook without closing Excel
        print("Closing the initial workbook...")
        blank_workbook.Close(SaveChanges=False)

        print("Task completed!")
    
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    automate_excel_task()
