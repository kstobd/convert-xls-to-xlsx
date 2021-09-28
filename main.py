import xlrd
import openpyxl


def cvt_xls_to_xlsx(*args, **kw):
    """Open and convert XLS file to openpyxl.workbook.Workbook object

    @param args: args for xlrd.open_workbook
    @param kw: kwargs for xlrd.open_workbook
    @return: openpyxl.workbook.Workbook
    """

    book_xls = xlrd.open_workbook(*args, formatting_info=True, ragged_rows=True, **kw)      #open xls book
    book_xlsx = openpyxl.workbook.Workbook()    #create xlsx book

    sheet_names = book_xls.sheet_names()    #create name for xls book
    for sheet_index in range(len(sheet_names)):
        sheet_xls = book_xls.sheet_by_name(sheet_names[sheet_index])
        #choose sheet index
        if sheet_index == 0:
            sheet_xlsx = book_xlsx.active
            sheet_xlsx.title = sheet_names[sheet_index]
        else:
            sheet_xlsx = book_xlsx.create_sheet(title=sheet_names[sheet_index])

        for crange in sheet_xls.merged_cells:
            rlo, rhi, clo, chi = crange

            sheet_xlsx.merge_cells(
                start_row=rlo + 1, end_row=rhi,
                start_column=clo + 1, end_column=chi,
            )

            #fill format book by value
            for row in range(0, sheet_xls.nrows):
                for col in range(0, sheet_xls.ncols):
                    try:
                        sheet_xlsx.cell(row = row+1 , column = col+1).value = sheet_xls.cell_value(row, col)
                    except BaseException:
                        pass

    return book_xlsx

fname = "Расписание занятий ф-т ИТ -  02.03.02, 09.03.01, 09.03.02 - 1, 2 курс - семестр 1,3 -2021-22 уч г (0916).xls"
cvt_xls_to_xlsx(fname).save(filename="outF.xlsx")