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


def process_pdf(pdf_file, output_dir):
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
    workbook = load_workbook("E:\\헌금보고서\\헌금보고서_양식.xlsx")
    sheet = workbook.active 
    for row_idx, row in enumerate(df.itertuples(index=False), start=2):
        for col_idx, value in enumerate(row, start=1):
            sheet.cell(row=row_idx+5, column=col_idx+2, value=value)
    workbook.save(output_dir)


def open_output_directory(output_folder):
    """ Opens the folder where the processed file is saved """
    subprocess.Popen(f'explorer "{output_folder}"', shell=True)


def process_cash_count_data():
    """ Executes the full process of locating and processing the PDF """
    cash_run_date = datetime.today().strftime("%Y%m%d")
    data_folder = "E:\\CashCounting\\BC-40 UpperMonitor v13\\Release\\Data"
    pdf_file = find_latest_pdf(data_folder, cash_run_date)
    if pdf_file:
        run_date = datetime.today().strftime("%m-%d-%Y")
        output_folder = os.path.join("E:\\헌금보고서", run_date)
        os.makedirs(output_folder, exist_ok=True)
        output_dir = os.path.join(output_folder, f'헌금보고서_{run_date}_{mass_time_var.get()}미사.xlsx')
        process_pdf(pdf_file, output_dir)
        
        success_win = tk.Toplevel(root)
        success_win.title("Success")
        success_win.geometry("400x200")
        success_label = tk.Label(success_win, text="현금 부분의 헌금보고서 생성이 완료됐습니다.", font=("Arial", 12), wraplength=350)
        success_label.pack(pady=10)
        
        open_dir_button = tk.Button(success_win, text="보고서 폴더 열기", command=lambda: open_output_directory(output_folder), font=("Arial", 12))
        open_dir_button.pack(pady=10)


def confirm_completion():
    """ Creates a pop-up window asking user to confirm cash counting completion """
    confirm_win = tk.Toplevel(root)
    confirm_win.title("Confirmation")
    confirm_win.geometry("500x200")
    
    label = tk.Label(confirm_win, text="현금 카운팅 완료 후 아래 버튼을 클릭하십시오.", font=("Arial", 12))
    label.pack(pady=10)
    
    def on_confirm():
        confirm_win.destroy()
        messagebox.showinfo("Processing", "현금 보고서 생성을 시작합니다.")
        process_cash_count_data()
    
    confirm_button = tk.Button(confirm_win, text="Yes", command=on_confirm, font=("Arial", 12))
    confirm_button.pack(pady=10)


def launch_app(app_path):
    """ Launch the cash counting application and set parameters """
    try:
        os.startfile(app_path)
        time.sleep(3)

        # Set the Port as "COM3"
        pyautogui.click(x=803, y=226)
        time.sleep(1)
        pyautogui.click(x=803, y=241)

        # Set the Baud Rate as 115200
        pyautogui.click(x=943, y=222)
        time.sleep(1)
        pyautogui.click(x=934, y=244)

        # Set the connection value to Open
        time.sleep(1)
        pyautogui.click(x=1020, y=224)
        
        confirm_completion()
    except Exception as e:
        messagebox.showerror("Error", f"Failed to launch application: {e}")


def start_cash_counting():
    """ Retrieves user selection and stores the mass time """
    mass_time = mass_time_var.get()
    if not mass_time:
        messagebox.showwarning("Warning", "Please select a mass time.")
        return
    
    messagebox.showinfo("Selection Confirmed", f"{mass_time} 미사가 선택됐습니다. 아래 'OK' 버튼을 클릭하십시오...")
    
    # Define the application path
    app_path = "E:\\CashCounting\\BC-40 UpperMonitor v13\\Release\\UpperMonitor.exe"
    launch_app(app_path)

# Create the main Tkinter window
root = tk.Tk()
root.title("Cash Counting System")
root.geometry("400x250")

title_label = tk.Label(root, text="미사 시간을 선택하십시오.", font=("Arial", 14))
title_label.pack(pady=10)

mass_time_var = tk.StringVar()
mass_time_var.set(None)  # Set to None to deselect all

options = ["7시반", "9시", "11시", "17시"]
for option in options:
    tk.Radiobutton(root, text=option, variable=mass_time_var, value=option, font=("Arial", 12)).pack()

start_button = tk.Button(root, text="현금 카운팅 시작", command=start_cash_counting, font=("Arial", 12))
start_button.pack(pady=20)

# Run the main loop
root.mainloop()