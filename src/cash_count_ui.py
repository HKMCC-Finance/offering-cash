import tkinter as tk
from tkinter import messagebox
import os
import time
import pyautogui
from datetime import datetime
import glob
import pdfplumber
import pandas as pd
from openpyxl import load_workbook
import subprocess

def pdf_to_dataframe(pdf_path):
    """ Extracts table data from a single-page PDF and returns it as a pandas DataFrame. """
    df = pd.DataFrame()
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        table = page.extract_table()
        if table:
            df = pd.DataFrame(table[4:], columns=table[3])
    return df


def find_latest_pdf(data_folder, date_str):
    """ Finds the latest PDF file in the specified directory """
    directory_path = os.path.join(data_folder, date_str)
    pdf_files = glob.glob(os.path.join(directory_path, "*.pdf"))
    if not pdf_files:
        messagebox.showerror("Error", "No PDF files found in the directory.")
        return None
    latest_pdf = max(pdf_files, key=os.path.getctime)
    return latest_pdf


def process_pdf(pdf_file, output_dir, is_second_offering=False, mass_time=None, populate_header=False):
    """ Processes the PDF and saves formatted data to an Excel file """
    df = pdf_to_dataframe(pdf_file)
    df = df[:-1]
    df['DENO'] = df['DENO'].astype(int)
    df['AMT'] = df['AMT'].astype(int)
    df['QTY'] = df['QTY'].astype(int)
    df = df.sort_values(by='DENO', ascending=True)
    df.rename(columns={'AMT': 'Amount'}, inplace=True)
    df.reset_index(inplace=True, drop=True)
    df.drop(columns='DENO', inplace=True)
    
    workbook = load_workbook(output_dir)
    sheet = workbook.active 
    
    # Populate date and time if this is the first time
    if populate_header and mass_time:
        current_date = datetime.today()
        mass_date = current_date.strftime("%m/%d/%y")
        sheet.cell(row=3, column=4, value=mass_date)
        sheet.cell(row=3, column=6, value=mass_time)
    
    # First offering: columns C-D (col_idx 3-4), starting at row 7
    # Second offering: columns E-F (col_idx 5-6), starting at row 7
    # Both use the same row offset (5), but different column offsets
    row_offset = 5
    col_offset = 2 if not is_second_offering else 4  # First offering: +2, Second offering: +4
    
    for row_idx, row in enumerate(df.itertuples(index=False), start=2):
        for col_idx, value in enumerate(row, start=1):
            sheet.cell(row=row_idx+row_offset, column=col_idx+col_offset, value=value)
    
    workbook.save(output_dir)


def open_output_directory(output_folder):
    """ Opens the folder where the processed file is saved """
    subprocess.Popen(f'explorer "{output_folder}"', shell=True)


def process_cash_count_data(is_second_offering=False):
    """ Executes the full process of locating and processing the PDF """
    cash_run_date = datetime.today().strftime("%Y%m%d")
    data_folder = "E:\\CashCounting\\BC-40 UpperMonitor v13\\Release\\Data"
    pdf_file = find_latest_pdf(data_folder, cash_run_date)
    
    if pdf_file:
        run_date = datetime.today().strftime("%m-%d-%Y")
        output_folder = os.path.join("E:\\헌금보고서", run_date)
        os.makedirs(output_folder, exist_ok=True)
        output_dir = os.path.join(output_folder, f'헌금보고서_{run_date}_{mass_time_var.get()}미사.xlsx')
        
        # If it's the second offering, the file should already exist from the first offering
        if is_second_offering and not os.path.exists(output_dir):
            messagebox.showerror("Error", "첫 번째 헌금 보고서 파일을 찾을 수 없습니다. 먼저 1차 헌금을 처리하십시오.")
            return
        
        # If it's the first offering and the file doesn't exist, copy the template
        if not is_second_offering and not os.path.exists(output_dir):
            template_path = "E:\\헌금보고서\\헌금보고서_양식.xlsx"
            if os.path.exists(template_path):
                import shutil
                shutil.copy(template_path, output_dir)
        
        process_pdf(pdf_file, output_dir, is_second_offering, mass_time_var.get(), not is_second_offering)
        
        offering_type = "2차" if is_second_offering else "1차"
        success_message = f"{offering_type} 헌금 현금 부분의 헌금보고서 생성이 완료됐습니다."
        
        success_win = tk.Toplevel(root)
        success_win.title("Success")
        success_win.geometry("600x300")
        success_label = tk.Label(success_win, text=success_message, font=("Arial", 18, "bold"), wraplength=550)
        success_label.pack(pady=30)
        
        # After first offering, check if second offering is needed
        if not is_second_offering and has_second_offering_var.get():
            continue_button = tk.Button(
                success_win, 
                text="2차 헌금 카운팅 시작", 
                command=lambda: [success_win.destroy(), start_second_offering()], 
                font=("Arial", 16, "bold"),
                width=20,
                height=2,
                bg="#4CAF50",
                fg="white",
                cursor="hand2"
            )
            continue_button.pack(pady=20)
        else:
            # Final completion - show option to open directory
            open_dir_button = tk.Button(
                success_win, 
                text="보고서 폴더 열기", 
                command=lambda: open_output_directory(output_folder), 
                font=("Arial", 16, "bold"),
                width=20,
                height=2,
                bg="#2196F3",
                fg="white",
                cursor="hand2"
            )
            open_dir_button.pack(pady=20)


def confirm_completion(is_second_offering=False):
    """ Creates a pop-up window asking user to confirm cash counting completion """
    confirm_win = tk.Toplevel(root)
    confirm_win.title("Confirmation")
    confirm_win.geometry("700x350")
    
    offering_type = "2차" if is_second_offering else "1차"
    label_text = f"{offering_type} 헌금 현금 카운팅 완료 후,\n현금을 기계 하단부에서 빼시고,\n아래 '완료' 버튼을 클릭하십시오."
    label = tk.Label(confirm_win, text=label_text, font=("Arial", 18, "bold"), justify="center")
    label.pack(pady=40)
    
    def on_confirm():
        confirm_win.destroy()
        process_message = f"{offering_type} 헌금 현금 보고서 생성을 시작합니다."
        messagebox.showinfo("Processing", process_message)
        process_cash_count_data(is_second_offering)
    
    confirm_button = tk.Button(
        confirm_win, 
        text="완료", 
        command=on_confirm, 
        font=("Arial", 18, "bold"),
        width=15,
        height=2,
        bg="#4CAF50",
        fg="white",
        cursor="hand2"
    )
    confirm_button.pack(pady=30)


def launch_app(app_path, is_second_offering=False):
    """ Launch the cash counting application and set parameters """
    try:
        os.startfile(app_path)
        time.sleep(0.5)

        # Set the Port as "COM3"
        pyautogui.click(x=803, y=226)
        time.sleep(0.25)
        pyautogui.click(x=803, y=241)

        # Set the Baud Rate as 115200
        pyautogui.click(x=943, y=222)
        time.sleep(0.25)
        pyautogui.click(x=934, y=244)

        # Set the connection value to Open
        time.sleep(0.25)
        pyautogui.click(x=1020, y=224)
        
        confirm_completion(is_second_offering)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to launch application: {e}")


def start_second_offering():
    """ Starts the second offering cash counting process """
    messagebox.showinfo("2차 헌금", "2차 헌금 카운팅을 시작합니다. 'OK' 버튼을 클릭하십시오...")
    # Skip launching the app for second offering - just wait for confirmation
    confirm_completion(is_second_offering=True)


def ask_second_offering():
    """ Creates a window to ask if there is a second offering """
    second_offering_win = tk.Toplevel(root)
    second_offering_win.title("2차 헌금 확인")
    second_offering_win.geometry("600x350")
    
    label = tk.Label(second_offering_win, text="2차 헌금이 있습니까?", font=("Arial", 20, "bold"))
    label.pack(pady=50)
    
    def on_yes():
        has_second_offering_var.set(True)
        second_offering_win.destroy()
        proceed_to_counting()
    
    def on_no():
        has_second_offering_var.set(False)
        second_offering_win.destroy()
        proceed_to_counting()
    
    button_frame = tk.Frame(second_offering_win)
    button_frame.pack(pady=30)
    
    yes_button = tk.Button(
        button_frame, 
        text="예", 
        command=on_yes, 
        font=("Arial", 18, "bold"), 
        width=12,
        height=2,
        bg="#4CAF50",
        fg="white",
        cursor="hand2"
    )
    yes_button.pack(side=tk.LEFT, padx=20)
    
    no_button = tk.Button(
        button_frame, 
        text="아니오", 
        command=on_no, 
        font=("Arial", 18, "bold"), 
        width=12,
        height=2,
        bg="#f44336",
        fg="white",
        cursor="hand2"
    )
    no_button.pack(side=tk.LEFT, padx=20)


def proceed_to_counting():
    """ Proceeds to start the first offering cash counting """
    mass_time = mass_time_var.get()
    second_offering_status = "있습니다" if has_second_offering_var.get() else "없습니다"
    
    messagebox.showinfo(
        "Selection Confirmed", 
        f"{mass_time} 미사가 선택됐습니다.\n2차 헌금: {second_offering_status}\n\n1차 헌금 카운팅을 시작합니다. 'OK' 버튼을 클릭하십시오..."
    )
    
    app_path = "E:\\CashCounting\\BC-40 UpperMonitor v13\\Release\\UpperMonitor.exe"
    launch_app(app_path, is_second_offering=False)


def start_cash_counting():
    """ Retrieves user selection and stores the mass time """
    mass_time = mass_time_var.get()
    if not mass_time:
        messagebox.showwarning("Warning", "Please select a mass time.")
        return
    
    # Ask about second offering before proceeding
    ask_second_offering()


# Create the main Tkinter window
root = tk.Tk()
root.title("Cash Counting System")
root.geometry("600x500")

title_label = tk.Label(root, text="미사 시간을 선택하십시오.", font=("Arial", 22, "bold"))
title_label.pack(pady=30)

mass_time_var = tk.StringVar()
mass_time_var.set(None)

has_second_offering_var = tk.BooleanVar()
has_second_offering_var.set(False)

options = ["7시반", "9시", "11시", "17시"]
for option in options:
    tk.Radiobutton(
        root, 
        text=option, 
        variable=mass_time_var, 
        value=option, 
        font=("Arial", 18, "bold"),
        indicatoron=1,
        cursor="hand2"
    ).pack(pady=8)

start_button = tk.Button(
    root, 
    text="현금 카운팅 시작", 
    command=start_cash_counting, 
    font=("Arial", 18, "bold"),
    width=18,
    height=2,
    bg="#2196F3",
    fg="white",
    cursor="hand2"
)
start_button.pack(pady=40)

# Run the main loop
root.mainloop()