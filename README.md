This is a Python script that performs several tasks related to handling Excel files, downloading files from URLs, extracting data from Excel sheets, and sending emails using Outlook. Let's go through the script step by step to understand what it does:

The script imports various modules like openpyxl, datetime, os, urllib, time, win32com.client, xlrd, and webbrowser.

The script sets up file paths for different Excel files that will be downloaded and processed later.

It loads an existing Excel workbook (final_workbook) using the openpyxl library.

The script proceeds to download three Excel files from URLs (excel_67_url, excel_file_url, and stock_value_url) using the urlretrieve function from urllib.request. If any of these files already exist, they will be removed before downloading the latest version.

After downloading the Excel file (excel_file_path), the script opens it using the default web browser.

The script then selects a sheet named after yesterday's date from the final_workbook and unhides it if it was hidden.

Data is copied from the source sheet (source_sheet) of the downloaded Excel file to the final sheet (final_sheet) of the final_workbook. The script loops through different ranges in the source sheet and copies the data to the corresponding cells in the final sheet.

The script then looks for the latest email with the subject prefix "FW: Отправка: Справка о зольности" and saves any attachments with that subject prefix to the local directory.

The script opens another Excel file named "Справка о зольности добытых и отгруженных углей за 29.06.2023г.xls" (xls_file_path) using the xlrd library and reads specific cell values from the sheet named "Зольность доб. и отгр.углей". It also opens another Excel file (stock_value_path) and reads cell values from it.

The data from the second Excel file (xls_file_path) and the data from the third Excel file (stock_value_path) are then copied to specific cells in the final sheet of the final_workbook.

The script saves the changes to the final_workbook.

After all the processing is done, the script waits for 10 seconds and then sends an email using Outlook. The email contains the previously saved final_workbook as an attachment, and the recipients are listed in the recipients_list.

Finally, the script calculates the execution time and prints it.
