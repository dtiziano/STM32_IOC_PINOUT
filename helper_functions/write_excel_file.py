import pandas as pd
import openpyxl
import openpyxl.styles


def write_excel_file(excel_file_path, pin_data, additional_info, duplicate_EXTI_error=True):
    sheet_name = None
    row_counter = 1
    row_header = None
    # Export the DataFrame to Excel with alternating row colors, borders, and adjusted column widths
    with pd.ExcelWriter(excel_file_path, engine='openpyxl') as writer:
        pin_data.to_excel(writer, sheet_name='Sheet1', index=False)

        # Access the Excel writer and the sheet
        workbook = writer.book
        sheet = writer.sheets['Sheet1']
        sheet_name =  sheet


        # Define the colors
        no_error_color = '00FF00' # Light green
        error_color = 'FF0000' #red
        color_header = '4285F4'  # Dark blue
        color_white = 'FFFFFF'  # White
        color_light_blue = 'DDEBF7'  # Light blue

         # Insert additional info
        sheet.insert_rows(0)
        sheet["A1"] = "MCU: " + additional_info['McuName']        
        sheet["B1"] = "CPN: " + additional_info['McuCPN']
        sheet["C1"] = "Footprint: " + additional_info['McuPackage']

        # for col_idx in range(1, sheet.max_column + 1):
        #     sheet.cell(row=1, column=col_idx).fill = openpyxl.styles.PatternFill(
        #         start_color = error_color if duplicate_EXTI_error else no_error_color, end_color = error_color if duplicate_EXTI_error else no_error_color, fill_type='solid'
        #     )
        row_counter += 1

        # Insert the EXTI check into the excel
        sheet.insert_rows(0)
        sheet["A1"] = "Error: Duplicate EXTI signal found" if duplicate_EXTI_error else "All good: no duplicated EXTI lines"
        for col_idx in range(1, sheet.max_column + 1):
            sheet.cell(row=1, column=col_idx).fill = openpyxl.styles.PatternFill(
                start_color = error_color if duplicate_EXTI_error else no_error_color, end_color = error_color if duplicate_EXTI_error else no_error_color, fill_type='solid'
            )
        row_counter += 1

       
        row_header = row_counter
        # Apply color to header
        for col_idx in range(1, sheet.max_column + 1):
            sheet.cell(row=row_counter, column=col_idx).fill = openpyxl.styles.PatternFill(
                start_color = color_header, end_color = color_header, fill_type='solid'
            )
        row_counter += 1
        # Apply alternating row colors starting from the second row
        for row_idx in range(row_counter, sheet.max_row + 1, 2):
            for col_idx in range(1, sheet.max_column + 1):
                fill_color = color_white if row_idx % 4 == 2 else color_light_blue
                sheet.cell(row=row_idx, column=col_idx).fill = openpyxl.styles.PatternFill(
                    start_color=fill_color, end_color=fill_color, fill_type='solid'
                )

        # Adjust column widths to fit the data
        for col in sheet.columns:
            max_length = 0
            column = col[0].column_letter  # Get the column name
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2)
            sheet.column_dimensions[column].width = adjusted_width

        # Add thin lines separating rows and columns
        thin_border = openpyxl.styles.Side(border_style='thin', color='000000')
        border = openpyxl.styles.Border(left=thin_border, right=thin_border, top=thin_border, bottom=thin_border)

        for row in sheet.iter_rows():
            for cell in row:
                cell.border = border

        print(f'DataFrame exported to {excel_file_path} with alternating row colors and{" without" if not duplicate_EXTI_error else ""} the first line.')

    wb = openpyxl.load_workbook(filename = excel_file_path)
    tab = openpyxl.worksheet.table.Table(displayName="pin_data", ref='A'+str(row_header)+f':{openpyxl.utils.get_column_letter(pin_data.shape[1])}{len(pin_data)+1}')
    wb[sheet_name.title].add_table(tab)
    wb.save(excel_file_path)