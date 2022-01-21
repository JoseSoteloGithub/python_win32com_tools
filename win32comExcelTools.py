def find_nth(haystack, needle, n):
    start = haystack.find(needle)
    while start >= 0 and n > 1:
        start = haystack.find(needle, start+len(needle))
        n -= 1
    return start
def get_column_letter(worksheet, substring, header_row, worksheet_column_letters_dict={}):
# Returns column letter or -1 if not found
    print('worksheet_column_letters_dict:', worksheet_column_letters_dict)
    try:
        column_letter = worksheet_column_letters_dict[substring]
        print('Returned from try: on worksheet_column_letters_dict')
        return column_letter
    except:
        last_column_letter = get_last_column_letter(worksheet, header_row)
        column_header = worksheet.Range('A' + str(header_row) + ':' + last_column_letter + str(header_row)).Find(What=substring, SearchOrder=constants.xlByRows, 
                                                        SearchDirection=constants.xlNext)

        #print('column_header:', column_header)
        if not column_header is None:
            first_dolla = find_nth(column_header.Address, '$', 1)
            second_dolla = find_nth(column_header.Address, '$', 2)
            column_letter = column_header.Address[first_dolla + 1:second_dolla]
        else:
            column_letter = -1
        #print(f'{worksheet.Parent.Name}.{worksheet.Name}.{substring}: ', column_letter)
        return column_letter
def get_last_column_letter(worksheet, header_row=0):
# Returns last column letter
    print('header_row:', header_row)
    if header_row != 0:
        last_column_letter_address = worksheet.Rows(header_row).Find(What='*', SearchOrder=constants.xlByColumns, SearchDirection=constants.xlPrevious).Address
    else:
        last_column_letter_address = worksheet.Cells.Find(What='*', SearchOrder=constants.xlByColumns, SearchDirection=constants.xlPrevious).Address
    last_column_letter = last_column_letter_address.split('$')[1]
    return(last_column_letter)
