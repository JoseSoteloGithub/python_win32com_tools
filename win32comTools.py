def get_range_whole(findString, currentSheet, withinRange=None, afterRange=None):

    if len(findString) <= 255:

        if (withinRange is None) and (afterRange is None):
            getPosition = currentSheet.Cells.Find(What=findString, After=currentSheet.Range("A1"), LookAt=constants.xlWhole, SearchDirection=constants.xlNext)
    
        if (not withinRange is None) and (afterRange is None):
            getPosition = currentSheet.Range(withinRange.Address).Find(What=findString, After=currentSheet.Range("A" + withinRange.Row), LookAt=constants.xlWhole, SearchDirection=constants.xlNext)
    
        elif (not withinRange is None) and (not afterRange is None):
            getPosition = currentSheet.Range(withinRange.Address).Find(What=findString, After=currentSheet.Range(afterRange.Address), LookAt=constants.xlWhole, SearchDirection=constants.xlNext)
        
        elif (withinRange is None) and (not afterRange is None):
            getPosition = currentSheet.Cells.Find(What=findString, After=currentSheet.Range(afterRange.Address), LookAt=constants.xlWhole, SearchDirection=constants.xlNext)
    
    get_range_whole = getPosition
    
    return get_range_whole

def find_nth(haystack, needle, n):
    start = haystack.find(needle)
    while start >= 0 and n > 1:
        start = haystack.find(needle, start+len(needle))
        n -= 1
    return start
def get_column_letter(worksheet_obj, substring_str, header_row_int, worksheet_column_letters_dict={}):
# Returns column letter or -1 if not found
    print('worksheet_obj:', worksheet_obj)
    print('worksheet_column_letters_dict:', worksheet_column_letters_dict)
    try:
        column_letter = worksheet_column_letters_dict[substring_str]
        print('Returned from try: on worksheet_column_letters_dict')
        return column_letter
    except:
        last_column_letter = get_last_column_letter(worksheet_obj, header_row_int)
        if header_row_int == 0:
            column_header = worksheet_obj.Cells.Find(What=substring_str, SearchOrder=constants.xlByRows, 
                                                            SearchDirection=constants.xlNext)        
        else:
            column_header = worksheet_obj.Range('A' + str(header_row_int) + ':' + last_column_letter + str(header_row_int)).Find(What=substring_str, SearchOrder=constants.xlByRows, 
                                                            SearchDirection=constants.xlNext)

        #print('column_header:', column_header)
        if not column_header is None:
            first_dolla = find_nth(column_header.Address, '$', 1)
            second_dolla = find_nth(column_header.Address, '$', 2)
            column_letter = column_header.Address[first_dolla + 1:second_dolla]
        else:
            column_letter = -1
        #print(f'{worksheet_obj.Parent.Name}.{worksheet_obj.Name}.{substring_str}: ', column_letter)
        return column_letter

def get_last_column_letter(worksheet_obj, header_row_int=0):
# Returns last column letter
    print('worksheet_obj.Name:', worksheet_obj.Name)
    print('header_row_int:', header_row_int)
    if header_row_int != 0:
        last_column_letter_address = worksheet_obj.Rows(header_row_int).Find(What='*', SearchOrder=constants.xlByColumns, SearchDirection=constants.xlPrevious).Address
    else:
        last_column_letter_address = worksheet_obj.Cells.Find(What='*', SearchOrder=constants.xlByColumns, SearchDirection=constants.xlPrevious).Address
    last_column_letter = last_column_letter_address.split('$')[1]
    return(last_column_letter)

def get_last_column_index(worksheet_obj, header_row_int=0):
# Returns last column as index
    if header_row_int != 0:
        last_column_range = worksheet_obj.Rows(header_row_int).Find(What='*', SearchOrder=constants.xlByColumns, SearchDirection=constants.xlPrevious)
    else:
        last_column_range = worksheet_obj.Cells.Find(What='*', SearchOrder=constants.xlByColumns, SearchDirection=constants.xlPrevious)

    return(last_column_range.Column)

def get_header_row(worksheet_obj, search_string_str):
# returns header row as int, -1 if not found
    header_range = worksheet_obj.Rows.Find(What=search_string_str, SearchOrder=constants.xlByRows, 
                                                        SearchDirection=constants.xlNext)
    if not header_range is None:
        header_row = header_range.Row
    else:
        header_row = -1
    header_row = str(header_row)
    return header_row

def get_last_row(worksheet_obj):
# Returns last row as int
    return(worksheet_obj.Cells.Find(What='*', SearchOrder=constants.xlByRows, SearchDirection=constants.xlPrevious).Row)

def sheet_exist(workbook, string):
    for ws_target in wb.Worksheets:
        if string in ws_target.Name:
            return True
    return False

def create_sheets():
    # Loop through list of name, check if there's a worksheet for the name, 
    # if not then copy the main worksheet and paste it at the end, rename it, apply filter for that name
    for name2 in name2_li:
        name2_sheet_exist = sheet_exist(wb, name2)
        if name2_sheet_exist == False:
            ws.Copy(After=wb.Sheets(len(wb.Worksheets)))
            wb.Worksheets(len(wb.Worksheets)).Name = name2
            wb.Worksheets(len(wb.Worksheets)).ListObjects(1).Range.AutoFilter(Field=name2_column_number, Criteria1= \
            name2)

import win32com.client
constants = win32com.client.constants
