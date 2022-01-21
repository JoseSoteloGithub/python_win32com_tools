def find_nth(haystack_str, needle_str, n_int):
    start = haystack_str.find(needle_str)
    while start >= 0 and n_int > 1:
        start = haystack_str.find(needle_str, start+len(needle_str))
        n_int -= 1
    return start
def get_column_letter(worksheet_obj, substring_str, header_row_int, worksheet_column_letters_dict={}):
# Returns column letter or -1 if not found
    print('worksheet_column_letters_dict:', worksheet_column_letters_dict)
    try:
        column_letter = worksheet_column_letters_dict[substring_str]
        print('Returned from try: on worksheet_column_letters_dict')
        return column_letter
    except:
        last_column_letter = get_last_column_letter(worksheet_obj, header_row_int)
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
def get_last_column_letter(worksheet_ojb, header_row_int=0):
# Returns last column letter
    print('header_row_int:', header_row_int)
    if header_row_int != 0:
        last_column_letter_address = worksheet_obj.Rows(header_row_int).Find(What='*', SearchOrder=constants.xlByColumns, SearchDirection=constants.xlPrevious).Address
    else:
        last_column_letter_address = worksheet_obj.Cells.Find(What='*', SearchOrder=constants.xlByColumns, SearchDirection=constants.xlPrevious).Address
    last_column_letter = last_column_letter_address.split('$')[1]
    return(last_column_letter)
