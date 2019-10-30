import xlsxwriter
from collections import OrderedDict


def wtfl(file,sheets):
    # Create a workbook and add a worksheet.
    workbook = xlsxwriter.Workbook(file)

    ## maintain order
    sheets = OrderedDict(sorted(sheets.items(), key=lambda t: t[0]))

    ## formatting params

    format_date = workbook.add_format()
    format_date.set_num_format('yyyy-mm-dd')
    format_date.set_font_name('Calibri')

    format_header = workbook.add_format()
    format_header.set_font_color('#0095C5')
    format_date.set_font_name('Calibri')
    format_header.set_bold()

    default_format = workbook.add_format()
    default_format.set_font_name('Calibri')
    
    worksheet = {}
    
    for sheet in sheets:
        worksheet[sheet] = workbook.add_worksheet(sheet)
        data = sheets[sheet]

        # Start from the first cell. Rows and columns are zero indexed.
        row = 0

        # Iterate over the data and write it out row by row.
        for line in data:
            col = 0
            for field in line:
                c_format = default_format
                if row == 0:
                    c_format = format_header
                elif 'datetime' in str(type(field)):
                    c_format = format_date
                worksheet[sheet].write(row, col, field, c_format)
                col += 1
            row += 1

    workbook.close()

## example
##wtfl('test.xlsx',{'1. Standard':[['1','vnfdkvondfiovbndfnovbidf','3'],['1','2','3'],['1','2','3']],'2. Detailed':[['1','2','3'],['1','2','3'],['1','2','3']]})
