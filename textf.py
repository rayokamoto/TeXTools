""" TeX Table Formatter """

import os
import xlwings as xw
import sys
import time


__version__ = '0.2.0'


def progress_bar(current, total):
    """
    Calculate progress of task.\n
    Takes two args: `current` - current value and `total` - total values, then
    calculates percentage. A progress bar is subsequently formed.
    """

    percent = (current/total)
    bar_len = 50
    bar = f"[{'#'*int(percent*bar_len)}{' '*(bar_len-int(percent*bar_len))}] {int(percent*100)}%"

    return bar



def write(text):
    """
    Write text to file.
    """

    path = os.path.join(os.getcwd(), 'tables')
    if not os.path.isdir(path):
        os.makedirs(path) # saves to current working directory (dir where cmd is run)
        print(f"New directory \"{path}\" has been made.")
    
    date = time.strftime('%Y-%m-%d_%H-%M-%S')
    filename = f"TeXtablef_{date}.txt"
    full_path = os.path.join(path, filename)
    with open(full_path, 'w', encoding='utf-8') as f:
        f.write(text)
    
    print(f"Data written to \"{full_path}\".")



def format(workbook, sheet, cells):
    """
    Retrieve data from Excel file, then format 
    """

    prompt = f"{'-'*60}\n File: {workbook.name}\nSheet: {sheet.name}\nCells: {cells}\n{'-'*60}"
    print(prompt)
    proceed = input("Proceed [y/n]? ")

    if proceed.lower() not in ["y", "yes"]:
        print("Operation has been cancelled.")
        return

    cells_data = sheet.range(cells)
    column_count = cells_data.columns.count
    #row_count = cells_data.rows.count
    total_cells = cells_data.count
    current_cell = 0

    data = ""

    # we don't want to take first iteration in the for loop as that will add "\\" to the start
    # and an if statement negating that won't be efficient
    cells = iter(sheet.range(cells))
    cell = next(cells)
    prev_row = cell.row # get initial row
    data += str(cell.value)
    current_cell+=1

    for cell in cells:
        current_cell+=1
        if prev_row < cell.row:
            data += r" \\" + "\n"
        else:
            data += " & "
        
        prev_row = cell.row
        cell_val = cell.value
        if cell.value is None:
            cell_val = ""
        data += str(cell_val)

        prog_bar = progress_bar(current_cell, total_cells)
        bar = f"   {prog_bar}"
        if current_cell == total_cells:
            print(" "*len(bar), end='\r')
            print(f"Reading data from {workbook.name}:")
            print(bar)
        else:
            print(f"Reading data from {workbook.name}:", end='\r')
            print(bar, end='\r')
        
    data += r" \\" + "\n"

    text = ""
    columns = "l"*column_count
    begin_t = "\\begin{table}[]\n\\begin{tabular}{" + columns + "}\n"
    end_t = "\\end{tabular}\n\\end{table}"
    
    text += begin_t + data + end_t
    write(text)



def help():
    """ 
    #### General help command 
    * Usage info
    * Arguments/commands available
    """

    print(f"TeXtf - TeX table formatter {__version__}\n")
    print("Usage: <filename> <command> [options]\n")
    print(
        "Arguments:\n"
        "  -h, --help   Shows help\n"
        "  format       Format tables. Provide arguments as follows: <file name> <sheet name> <selected cells>"
    )



def get_fmtdata(argc, argv):
    """
    Get data required for formatting, such as file, sheet, cells, etc.
    """

    if argc == 1:
        print("An argument must be provided.")
        return
    
    # handle errors when args not provided
    
    try:
        workbook = xw.Book(argv[1])
    except FileNotFoundError:
        print(f"File {argv[1]} could not be found. If you have spaces in the file name, enclose the whole file name in quotes (\") or (').")
    except Exception as e:
        print(f"An error occured:")
        raise e
    
    try:
        sheet = workbook.sheets[argv[2]]
    except Exception as e:
        print("Sheet name must be valid.")
    
    try:
        cells = argv[3]
    except Exception as e:
        print("Cells must be valid.")
    
    format(workbook, sheet, cells)



# ! ADD ERROR HANDLING
def main(argc, argv):
    if argc == 0:
        help()        
    elif argv[0].lower() in ["--help", "-h"]:
        help()
    elif argv[0].lower() in ["help"]: # arg for help cmd was provided
        if argc == 1: # i.e. only help was called
            help()
        else: 
            print("Currently, detailed help commands do not exist.")
    elif argv[0].lower() in ["format"]:
        get_fmtdata(argc, argv)
    else:
        print(f"Command \"{argv[0]}\" does not exist.")



def debug_mode(argc, argv):
    if argc == 0:
        workbook = xw.Book('table_data.xlsx')
        format(workbook, workbook.sheets['sheet1'], 'a1:b10')
        return



if __name__ == "__main__":
    argv = sys.argv[1:] # disregard file itself
    argc = len(argv)
    main(argc, argv)
    #debug_mode(argc, argv)
