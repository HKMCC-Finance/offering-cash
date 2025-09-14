# Cash Counting Application

A Python-based GUI application for automating cash counting processes and generating donation reports.

## Features

- **Mass Time Selection**: Choose from different mass times (7시반, 9시, 11시, 17시)
- **Automated Cash Counting**: Launches and configures the BC-40 UpperMonitor application
- **PDF Processing**: Extracts data from cash counting PDF reports
- **Excel Report Generation**: Creates formatted donation reports in Excel format
- **File Management**: Automatically organizes reports by date

## Requirements

- Python 3.x
- Required packages:
  - tkinter
  - pyautogui
  - pdfplumber
  - pandas
  - openpyxl

## Usage

1. Run the application:
   ```
   python src/cash_count_ui.py
   ```

2. Select the appropriate mass time from the radio buttons

3. Click "현금 카운팅 시작" to begin the cash counting process

4. The application will:
   - Launch the BC-40 UpperMonitor
   - Configure connection settings automatically
   - Process the generated PDF data
   - Create a formatted Excel report

## File Structure

- `src/cash_count_ui.py` - Main GUI application
- `coordinate_finder.py` - Utility for finding UI coordinates
- `coordinate_capture.py` - Interactive coordinate capture tool

## Notes

- Ensure the BC-40 UpperMonitor application is installed in the specified path
- PDF files are expected in the Data folder with date-based organization
- Output reports are saved to the 헌금보고서 directory