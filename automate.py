import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import pandas as pd
import pyautogui
import time
from openpyxl import Workbook, load_workbook
from datetime import datetime
import os
import logging
import threading
import keyboard  # Ensure you have the `keyboard` package installed (pip install keyboard)
import cv2  # Ensure you have the `opencv-python` package installed (pip install opencv-python)

# Configure logging
logging.basicConfig(filename='automation_log.txt', level=logging.DEBUG, 
                    format='%(asctime)s - %(levelname)s - %(message)s')

# Global variable to store data
data = None

# Function to load data from an Excel file
def load_data():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if not file_path:
        return
    try:
        global data
        data = pd.read_excel(file_path, header=None)  # Read without headers
        logging.debug(f"Loaded data:\n{data}")  # Debug: Log the entire loaded data

        # Check the number of columns and assign column names dynamically
        if len(data.columns) == 2:
            data.columns = ['Date', 'Invoice Number']
        elif len(data.columns) == 3:
            data.columns = ['Date', 'Invoice Number', 'Serial']
        else:
            raise ValueError("Unexpected number of columns in the loaded data.")

        logging.debug(f"Columns: {data.columns}")  # Debug: Log the column names
        if data.empty:
            logging.warning("DataFrame is empty. Please check the file content.")
            messagebox.showwarning("No Data", "The loaded data is empty. Please check the file content.")
        else:
            messagebox.showinfo("Data Loaded", "Data loaded successfully!")
    except Exception as e:
        logging.error(f"Error loading data: {e}")  # Debug
        messagebox.showerror("Error", f"Failed to load data: {e}")

# Function to check for the pop-up and close it if it appears
def check_and_close_popup(image_path, retries=5, interval=1):
    for _ in range(retries):
        try:
            popup_location = pyautogui.locateOnScreen(image_path, confidence=0.9)
            if popup_location:
                logging.debug("Pop-up detected. Closing it.")  # Debug: Log message when pop-up is detected
                pyautogui.press('enter')
                time.sleep(interval)
                return True
        except Exception as e:
            logging.error(f"Error checking pop-up: {e}")  # Debug
        time.sleep(interval)
    logging.debug("Pop-up not detected.")  # Debug: Log message when pop-up is not detected
    return False

# Global flag to stop automation
stop_automation_flag = False

# Function to stop the automation process
def stop_automation():
    global stop_automation_flag
    stop_automation_flag = True
    logging.debug("Stop automation flag set.")  # Debug: Log when stop is requested

# Function to automate data entry
def automate_entry():
    global stop_automation_flag
    stop_automation_flag = False  # Ensure the flag is reset at the start
    try:
        logging.debug("Starting automation...")  # Debug
        if data is None:
            messagebox.showerror("No Data", "Please load the data first.")
            return

        # Ensure data is not empty
        if data.empty:
            logging.warning("Data is empty. Automation cannot proceed.")
            messagebox.showwarning("No Data", "The loaded data is empty. Please check the file content.")
            return

        def monitor_stop_key():
            keyboard.wait('`')
            stop_automation()

        monitor_thread = threading.Thread(target=monitor_stop_key)
        monitor_thread.start()

        for index, row in data.iterrows():
            if stop_automation_flag:
                logging.debug("Stop automation flag detected. Ending automation.")
                break

            logging.debug(f"Processing row {index}: {row}")  # Debug: Log the current row being processed
            user_date = row['Date']
            invoice_number = row['Invoice Number']
            serial_number = row[2] if len(row) > 2 and not pd.isna(row[2]) else ''

            if pd.isna(invoice_number) or pd.isna(user_date):
                logging.debug(f"Skipping row {index} due to missing data")  # Debug: Log message if data is missing
                continue

            # Convert invoice_number to string if it's a number
            invoice_number = str(invoice_number)
            logging.debug(f"Invoice Number: {invoice_number}")  # Debug: Log the invoice number

            # Focus on the application window
            logging.debug("Focusing on application window...")  # Debug
            pyautogui.click(624, 229)  # Click to focus on the application window (coordinates may need adjustment)
            time.sleep(1)  # Wait for the application window to focus

            # Click on the "Order #" field (coordinates may need adjustment)
            logging.debug("Clicking on Order # field...")  # Debug
            pyautogui.click(689, 365)  # Adjust these coordinates to the "Order #" field
            time.sleep(1)
            pyautogui.write(invoice_number)  # Enter the invoice number
            logging.debug(f"Entered Invoice Number: {invoice_number}")  # Debug: Confirm invoice number entry
            time.sleep(0.1)
            pyautogui.press('tab')
            time.sleep(1)

            # Check for the pop-up and close it if it appears
            logging.debug("Checking for pop-up...")  # Debug
            check_and_close_popup('credit_limit_exceeded.png', retries=2, interval=1)
            time.sleep(2)

            # Change Mode/Prev/Disp: O to I, continue regardless of pop-up detection
            logging.debug("Changing Mode/Prev/Disp to I...")  # Debug
            pyautogui.click(660, 360)  # Ensure correct field is focused
            pyautogui.press('i')
            pyautogui.press('tab')
            time.sleep(1)
            pyautogui.press('enter')
            time.sleep(3)  # Wait for the new window to appear

            # If there is a serial number in column C
            if serial_number:
                logging.debug(f"Entering Serial Number: {serial_number}")  # Debug
                time.sleep(1)
                pyautogui.write(serial_number)
                time.sleep(0.5)
                pyautogui.hotkey('alt', 'r')  # Close the screen with ALT+R
                time.sleep(1)

            # Enter the date in the "Ship Date" field on the second screen
            logging.debug("Entering Ship Date...")  # Debug
            pyautogui.click(661, 522)  # Adjust these coordinates to the "Ship Date" field
            time.sleep(1)
            pyautogui.write(user_date)
            logging.debug(f"Entered Ship Date: {user_date}")  # Debug: Confirm date entry
            pyautogui.click(1169, 795)
            time.sleep(2)
            pyautogui.press('enter')

            # Log progress
            logging.debug("Automation step completed successfully.")  # Debug

    except Exception as e:
        logging.error(f"Error during automation: {e}")  # Debug: Log the exception
        messagebox.showerror("Error", f"Failed to complete automation: {e}")

    finally:
        monitor_thread.join()

# Function to write data to an Excel sheet
def write_to_excel(event=None):
    user_date = date_entry.get()
    if not user_date:
        messagebox.showerror("No Date", "Please enter the date first.")
        return

    try:
        # Format the date as MM/DD/YY
        formatted_date = datetime.strptime(user_date, "%Y-%m-%d").strftime("%m/%d/%y")
    except ValueError as e:
        logging.error(f"Error parsing date: {e}")  # Debug
        messagebox.showerror("Invalid Date", "Please enter the date in the format YYYY-MM-DD.")
        return

    invoice_number = invoice_entry.get()
    if not invoice_number:
        messagebox.showerror("No Invoice Number", "Please enter the invoice number.")
        return

    serial_value = serial_var.get()
    serial_text = ""

    if serial_value:
        serial_text = simpledialog.askstring("Serial Number", "Please enter the serial number:")
        if not serial_text:
            messagebox.showerror("No Serial Number", "Please enter the serial number.")
            return

    try:
        file_path = "automate.xlsx"

        if os.path.exists(file_path):
            workbook = load_workbook(file_path)
            sheet = workbook.active
        else:
            workbook = Workbook()
            sheet = workbook.active

        # Add data to the next empty row
        next_row = sheet.max_row + 1
        row_data = [formatted_date, invoice_number]
        if serial_value:
            row_data.append(serial_text)
        else:
            row_data.append("")

        sheet.append(row_data)
        logging.debug(f"Writing row: {row_data}")  # Debug: Log the row data being written

        workbook.save(file_path)
        invoice_entry.delete(0, tk.END)
        serial_checkbox.deselect()
    except Exception as e:
        logging.error(f"Error writing to Excel: {e}")  # Debug
        messagebox.showerror("Error", f"Failed to write to Excel: {e}")

# Function to purge data from the Excel sheet
def purge_data():
    try:
        file_path = "automate.xlsx"

        if os.path.exists(file_path):
            workbook = load_workbook(file_path)
            sheet = workbook.active

            # Clear all rows, including headers
            for row in sheet.iter_rows():
                for cell in row:
                    cell.value = None
                    
            workbook.save(file_path)
            messagebox.showinfo("Purge Complete", "All data has been purged from the Excel sheet.")
        else:
            messagebox.showinfo("File Not Found", "The Excel file does not exist.")
    except Exception as e:
        logging.error(f"Error purging data: {e}")  # Debug
        messagebox.showerror("Error", f"Failed to purge data: {e}")

# Main GUI setup
root = tk.Tk()
root.title("Data Entry Automation")

data = None

# Create and place the date entry widget
date_label = tk.Label(root, text="Enter the date (YYYY-MM-DD):")
date_label.pack(pady=5)
date_entry = tk.Entry(root)
date_entry.pack(pady=5)

# Create and place the invoice number entry widget
invoice_label = tk.Label(root, text="Enter the invoice number:")
invoice_label.pack(pady=5)
invoice_entry = tk.Entry(root)
invoice_entry.pack(pady=5)

# Create and place the serial checkbox
serial_var = tk.IntVar()
serial_checkbox = tk.Checkbutton(root, text="Serial", variable=serial_var)
serial_checkbox.pack(pady=5)

# Create and place the load data button
load_button = tk.Button(root, text="Load Data", command=load_data)
load_button.pack(pady=10)

# Create and place the start automation button
start_button = tk.Button(root, text="Start Automation", command=automate_entry)
start_button.pack(pady=10)

# Create and place the stop automation button
stop_button = tk.Button(root, text="Stop Automation", command=stop_automation)
stop_button.pack(pady=10)

# Create and place the write to Excel button
write_button = tk.Button(root, text="Write to Excel", command=write_to_excel)
write_button.pack(pady=10)

# Create and place the purge data button
purge_button = tk.Button(root, text="Purge Data", command=purge_data)
purge_button.pack(pady=10)

# Bind the Enter key to write to Excel
root.bind('<Return>', write_to_excel)

# Run the application
root.mainloop()
