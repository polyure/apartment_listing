A python script that calls Oikotie API, saves found apartments on a map, and appends them to an .xlsx file. 

Note, xlwings requires Excel installation on the computer, and I'd rather use openpyxl, but that corrupts the resulting file since it cannot handle the indirect-function and turns them into array formulae. If you find a workaround, let me know.