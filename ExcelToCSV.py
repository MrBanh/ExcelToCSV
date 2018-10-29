#! python3

# ExcelToCSV.py - Reads all excel files in current working directory and outputs
# them as CSV files

import csv, os, openpyxl

desktop = os.path.join(os.environ['USERPROFILE'], 'Desktop\\')
os.chdir(desktop)
os.makedirs('Excel_to_CSV_Files', exist_ok=True)
locatedCSVFiles = os.path.join(desktop, 'Excel_to_CSV_Files')
locatedExcelFiles = os.path.join(desktop, 'excelSpreadsheets')
os.chdir(locatedExcelFiles)

for excelFile in os.listdir('.'):
    # Skip non-xlsx files, load the workbook object
    if not excelFile.endswith('xlsx'):
        continue
    wb = openpyxl.load_workbook(excelFile)

    # Loop through every sheet in the workbook
    for sheetName in wb.sheetnames:
        sheet = wb[sheetName]

        # Create the CSV filename from Excel filename and sheet title
        csvFileName = f'{excelFile.split(".xlsx")[0]}_{sheetName}.csv'
        csvFile = open(os.path.join(locatedCSVFiles, csvFileName), 'w', newline='')

        # Create the csv.writer object for this CSV File
        csvWriter = csv.writer(csvFile)

        # Loop through every row in the sheet
        for rowNum in range(1, sheet.max_row + 1):
            rowData = []    # append each cell to this list
            # Loop through each cell in the row
            for colNum in range(1, sheet.max_column + 1):
                # Append each cell's data to rowData
                    rowData.append(sheet.cell(row=rowNum, column=colNum).value)
                
            # Write the rowData list to the CSV file
            csvWriter.writerow(rowData)

        csvFile.close()