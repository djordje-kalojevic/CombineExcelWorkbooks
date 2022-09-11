"""Allows the user to combine Excel workbooks.
Note that the modules used to do allow this by default,
as such, the results are not always perfect.
Simpler workbooks should pose no issue though."""

from tkinter import Tk
from tkinter.messagebox import showerror
from tkinter.filedialog import askopenfile, asksaveasfilename
from copy import copy
from os import system
from os.path import isfile
import sys
from typing import IO
from openpyxl import load_workbook, Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils.cell import get_column_letter
from openpyxl.cell.cell import Cell
from PIL import Image


def create_program_mainloop(transparent_icon_location: str) -> Tk:
    """Creates an instance of tk.Tk and replaces its icon with a transparent one."""

    if not isfile(transparent_icon_location):
        transparent_icon = Image.new('RGBA', (16, 16), (0, 0, 0, 0))
        transparent_icon.save(transparent_icon_location, 'ICO')

    root = Tk()
    root.title('')
    # hides its window
    root.withdraw()
    root.iconbitmap(True, transparent_icon_location)

    return root


def copy_sheet(source_sheet: Worksheet, target_sheet: Worksheet):
    """Copies a sheet from one workbook to another, cell by cell."""

    copy_cells(source_sheet, target_sheet)
    copy_sheet_attributes(source_sheet, target_sheet)
    copy_row_column_dimensions(source_sheet, target_sheet)


def copy_cells(source_sheet: Worksheet, target_sheet: Worksheet):
    """Copies cell contents as well as its properties from source to target sheet."""
    for (row, col), source_cell in source_sheet._cells.items():  #pylint: disable=protected-access
        target_cell: Cell = target_sheet.cell(column=col, row=row)
        source_cell: Cell
        # copies cell content
        target_cell._value = source_cell._value  #pylint: disable=protected-access
        target_cell.data_type = source_cell.data_type

        if source_cell.has_style:
            target_cell.font = copy(source_cell.font)
            target_cell.border = copy(source_cell.border)
            target_cell.fill = copy(source_cell.fill)
            target_cell.number_format = copy(source_cell.number_format)
            target_cell.protection = copy(source_cell.protection)
            target_cell.alignment = copy(source_cell.alignment)

        if source_cell.hyperlink:
            target_cell._hyperlink = copy(source_cell.hyperlink)  #pylint: disable=protected-access

        if source_cell.comment:
            target_cell.comment = copy(source_cell.comment)


def copy_sheet_attributes(source_sheet: Worksheet, target_sheet: Worksheet):
    """Copies sheet attributes from source to target sheet."""
    target_sheet.auto_filter = copy(source_sheet.auto_filter)
    target_sheet.sheet_format = copy(source_sheet.sheet_format)
    target_sheet.sheet_properties = copy(source_sheet.sheet_properties)
    target_sheet.merged_cells = copy(source_sheet.merged_cells)
    target_sheet.page_margins = copy(source_sheet.page_margins)
    target_sheet.page_setup = copy(source_sheet.page_setup)
    target_sheet.print_options = copy(source_sheet.print_options)
    target_sheet.freeze_panes = copy(source_sheet.freeze_panes)


def copy_row_column_dimensions(source_sheet: Worksheet,
                               target_sheet: Worksheet):
    """Copies row hights and column width from source to target sheet.
    Note: this is not always perfect."""
    for row in range(source_sheet.max_row + 1):
        target_sheet.row_dimensions[row].height = (
            source_sheet.row_dimensions[row].height)

    # column indexes start with 1
    for index in range(source_sheet.max_column + 1)[1:]:
        letter = get_column_letter(index)
        target_sheet.column_dimensions[letter].width = (
            source_sheet.column_dimensions[letter].width)


def browse_excel(title: str | None = None) -> IO:
    """Prompts user to select .xls or preferably .xlsx file."""
    if title is None:
        title = "Choose a file"
    file = askopenfile(mode='rb',
                       title=title,
                       filetypes=[("Excel (.xlsx) file", "*.xlsx"),
                                  ("Excel (.xls) file", "*.xls")])
    if not file:
        sys.exit()

    return file


def save_excel(file: Workbook) -> str:
    """Saves the combined workbook."""
    file_saved = False
    while not file_saved:
        try:
            file_name = asksaveasfilename(defaultextension=".xlsx",
                                          filetypes=[("Excel file (.xlsx)",
                                                      "*.xlsx")])
            file_saved = True
        except PermissionError:
            showerror(
                title="Error occurred!",
                message=
                ("File could not be saved because it is already opened by another process "
                 "(likely Excel.exe). Please close it before continuing."))

    if not file_name:
        sys.exit()

    file.save(file_name)

    return file_name


def combine_workbooks():
    """Combines two workbooks by copying sheets.
    This is done CELL by CELL and can take a little bit for larger workbooks.
    Although it should not provide any issues on modern systems."""

    create_program_mainloop(transparent_icon_location='icon.ico')
    #current_dir = pathlib.Path().absolute()

    source_file = browse_excel(title="Please select the source workbook")

    target_file = browse_excel(title="Please select the target workbook")

    file_saved = False
    while not file_saved:
        try:
            file_name = asksaveasfilename(defaultextension=".xlsx",
                                          filetypes=[("Excel file (.xlsx)",
                                                      "*.xlsx")])
            file_saved = True
        except PermissionError:
            showerror(
                title="Error occurred!",
                message=
                ("File could not be saved because it is already opened by another process "
                 "(likely Excel.exe). Please close it before continuing."))

    if not file_name:
        sys.exit()

    source_workbook: Worksheet = load_workbook(source_file)
    target_workbook: Worksheet = load_workbook(target_file)

    for sheet_name in source_workbook.sheetnames:
        source_sheet = source_workbook[sheet_name]
        target_sheet = target_workbook.create_sheet(sheet_name)
        copy_sheet(source_sheet, target_sheet)

    target_workbook.save(file_name)

    # opens the saved workbook
    system(f'"{file_name}"')


if __name__ == '__main__':
    combine_workbooks()
