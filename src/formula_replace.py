import xlwings as xw
import openpyxl


def pop_rows(workbook):
    #do a sunburst graph, parent: Tot_ASSM, Tot_RECC Child: median, mean, max, min Total saving type Primary, Secondary, Tertiary, Quaternary
    wb = workbook
    destination_sheet = wb['RECC']
    #Add recommended savings, plant energy costs and number of reccs
    count = 0
    column_to_check = "G"
    populated_rows = 0
    # Iterate over rows in the column and count populated ones
    for row in destination_sheet.iter_rows(min_row=1, max_row=destination_sheet.max_row, min_col=1, max_col=1):
        cell_value = row[0].value
        if cell_value is not None and str(cell_value).strip() != '':
            populated_rows += 1
    return populated_rows

    


def replace_cell_value(file_path, sheet_name, cell_address):
    # Connect to the Excel application
    app = xw.App()

    # Open the Excel file
    wb = app.books.open(file_path)

    # Select the worksheet
    sheet = wb.sheets[sheet_name]

    # Replace the value of the cell
    x = sheet.range(cell_address).value
    sheet.range(cell_address).value = x
    # Save the changes
    wb.save()

    # Close the workbook and Excel application
    wb.close()
    app.quit()



