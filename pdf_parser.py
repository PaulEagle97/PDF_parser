"""
This script was written for parsing two specific PDF documents:
'The_Oxford_3000'
'The_Oxford_5000'
It extracts all of the word columns from them, 
breaks each line into several duplicates based on the number of parts of speech,
and then exports them into .csv and .xlsx files as a table with 4 rows: 
'Word', 'Comments', 'Part of Speech', 'Level', 'Duplicate'

So an original line 'lie (tell a lie) v., n. B1' is separated into 2 rows (due to 2 parts of speech):
lie, (tell a lie), v., B1
lie, (tell a lie), n., B1
"""
import os
import re
import csv
from openpyxl import Workbook
from PyPDF2 import PdfReader


def visitor_body_1st_page(text, cm, tm ,fontDict, fontSize):
    """
    Helper function for parsing the first page
    Crops the header and footer texts
    """
    y_coord = tm[5]
    if y_coord < 730 and y_coord > 30:
        valid_parts.append(text)


def visitor_body_other_pages(text, cm, tm ,fontDict, fontSize):
    """
    Helper function for parsing every other page
    Crops the footer text
    """
    y_coord = tm[5]
    if y_coord > 30:
        valid_parts.append(text)    


def text_splitter(a_list):
    '''
    Separate lines at the beginning and end of each column of text 
    that got merged together due to parsing errors.
    '''
    pattern = '[A-C][1-2].+'
    text_split = []
    for line in a_list:
        lines, to_split = recurs_line_splitter(line, pattern)
        
        if to_split:
            text_split.extend(lines)
        else:
            text_split.append(line)
    
    return text_split


def recurs_line_splitter(line, pattern):
    '''
    Recursive function that splits a merged (line)
    based on the (pattern). Returns a list of lines.
    '''
    lines = []
    to_split = False
    for idx, string in enumerate(line):
        if re.match(pattern, string):
            to_split = True
            split_idx = idx
            split_str = string
            break
    
    if not to_split:
        return [line], to_split
    
    else:
        before_split = [split_str[:2]]
        after_split = [split_str[2:]]
        line_1 = line[ :split_idx]
        line_1.extend(before_split)
        lines.append(line_1)  
        after_split.extend(line[(split_idx + 1): ])
        other_splits = recurs_line_splitter(after_split, pattern)[0]
        lines.extend(other_splits)

        return lines, to_split


def str_merger(a_list):
    '''
    Merges together several occasions 
    when strings [A-C][1-2] got separated
    '''
    pattern = r'\b[A-C]{1}\b'
    for line in a_list:
        for idx, string in enumerate(line):
            if re.match(pattern, string):
                new_str = line[idx] + line[idx + 1]
                line.pop(idx)
                line.pop(idx)
                line.insert(idx, new_str)


def text_merger(a_list):
    '''
    Merge lines that got separated 
    due to a word being transposed to two rows
    '''
    pattern = '[A-C][1-2]'
    text_merge = []
    merged = False
    for idx, line in enumerate(a_list):
        if not merged:
            if idx < len(a_list) - 1:
                if len(a_list[idx + 1]) >= 3 and re.match(pattern, line[-1]):
                    text_merge.append(line)
                else:
                    line_merged = line + a_list[idx + 1]
                    text_merge.append(line_merged)
                    merged = True                        
            else:
                text_merge.append(line)
        else:
            merged = False   
    
    return text_merge


def cleaner_func(a_list):
    ''' 
    Clean all of the lines from redundant ('') and other strings
    '''
    pattern = '[1-9]'
    pattern_2 = '.+[A-C][1-2]'
    redundant_strs = {'', '.', ','}
    for idx, line in enumerate(a_list):
        for str in redundant_strs:
            c = line.count(str)
            for dummy_i in range(c):
                line.remove(str)
        if re.match(pattern, line[1]):
            line.pop(1)
        if re.match(pattern, line[0][-1]):
            a_word = a_list[idx].pop(0)
            a_list[idx].insert(0, a_word[ :-1])
        if re.match(pattern_2, line[-1]):
            str_1 = line[-1][ :-2]
            str_2 = line[-1][-2: ]
            line.pop()
            line.append(str_1)
            line.append(str_2)


def bracket_fixer(a_list):
    '''
    Fix lines with multiple word expressions inside brackets
    '''
    text_fixed = []
    pattern_1 = '\((.+)'
    pattern_2 = '(.+)\)'
    for line in a_list:
        if re.match(pattern_1, line[1]) and line[1][-1] != ")":            
            c = 2
            while not re.match(pattern_2, line[c]):
                c += 1
            fixed_brackets = [line[1]]
            for idx in range(2, c):
                fixed_brackets.append(line[idx])
            fixed_brackets.append(line[c])
            fixed_brackets = ' '.join(fixed_brackets)
            new_line = [line[0]] + [fixed_brackets] + line[(c + 1): ]
            text_fixed.append(new_line)
        else:
            text_fixed.append(line)
    
    return text_fixed


def line_dublicater(a_list):
    '''
    Takes a list with lines containing several parts of speech and/or levels 
    and return a list with word duplicates for each part of speech/level
    '''
    dub_lst = []
    pattern = '\((.+)\)'
    for line in a_list:
        if re.match(pattern, line[1]):
            bracket_val = 1
        else:
            bracket_val = 0
        #print(line)
        dub_line = line_part(line, bracket_val)
        if len(dub_line) > 1:
            for idx, dub_word in enumerate(dub_line):
                if idx > 0:
                    dub_word.append('*')
        dub_lst.extend(dub_line)

    return dub_lst


def line_part(line, bracket_val):
    '''
    Recursive function that breaks one line with (n) parts of speech
    and (m) levels into (m * n) duplicates of the same word
    '''
    pattern = '[A-C][1-2]'
    for idx, elem in enumerate(line):
        if re.match(pattern, elem):
            endpoint = idx
            a_level = elem
            break
    
    if endpoint == 2 + bracket_val and len(line) == (3 + bracket_val):
        return [line]
    
    an_instance = line[ :(2 + bracket_val)] + [a_level]
    a_word = line[ :(1 + bracket_val)]
    if endpoint > (2 + bracket_val):
        return [an_instance] + line_part(a_word + line[(2 + bracket_val): ], bracket_val)
    
    return [an_instance] + line_part(a_word + line[(endpoint + 1): ], bracket_val)


def line_alligner(a_list):
    '''
    Inserts an empty string element into every line 
    which doesn't have extra comments in "()"
    '''
    pattern = '\((.+)\)'
    for line in a_list:
        if not re.match(pattern, line[1]):
            line.insert(1, '')


def main(filename):
    '''
    The main body of the script
    '''
    global valid_parts

    print(f'\nParsing the file ---> "{filename}"')

    #compute paths
    curr_dir = os.getcwd()
    file_abs_path = os.path.join(curr_dir, "PDF_parser", "Original", filename)

    #initialize objects for further parsing
    reader = PdfReader(file_abs_path)
    text_1st_page = reader.pages[0]
    text_other_pages = reader.pages[1:]
    valid_parts = []

    #this block converts (reader) objects into multiline strings
    ################################################################
    #parse the first PDF page
    text_1st_page.extract_text(visitor_text = visitor_body_1st_page)
    valid_parts += "\n"

    #parse all the rest of pages
    for page in text_other_pages:
        page.extract_text(visitor_text = visitor_body_other_pages) + "\n"
    text_body = "".join(valid_parts)
    ################################################################

    #partition the multiline text string into a list of lists (line = list) with each word separated
    text_lst = [ [s.strip(",") for s in line.split() if s] for line in text_body.split('\n') if line]

    #substitute a specific line in ver. 3000
    if filename == 'The_Oxford_3000.pdf':
        text_lst[0] = ['a', 'indefinite article', 'A1']

    #the following block of functions resolves different specific types of inaccuracies
    #that occurred during the first phase of parsing by matching and substituting patterns
    ##################################
    text_lst = text_splitter(text_lst)

    str_merger(text_lst)  

    text_lst = text_merger(text_lst)    

    cleaner_func(text_lst) 

    text_lst = bracket_fixer(text_lst)
    ##################################

    #this function breaks one line into several duplicate lines
    #in case a word has several parts of speech mentioned
    text_lst = line_dublicater(text_lst)

    #alligns each line for further export to the table
    #by inserting empty string elements where needed
    line_alligner(text_lst)

    #creates the final csv file from the (text_lst) list of lists
    csv_filename = filename[: -4] + '.csv'
    csv_abs_path = os.path.join(curr_dir, "PDF_parser", "Parsed", csv_filename)
    header = ['Word', 'Comments', 'Part of Speech', 'Level', 'Duplicate']
    with open(csv_abs_path, 'w', encoding='UTF8', newline='') as f:
        writer = csv.writer(f)
        # write the header
        writer.writerow(header)
        # write multiple rows
        writer.writerows(text_lst)

    #ask the user whether to export to the Excel format
    valid_input = False
    while not valid_input:
        user_input = input('Do you want to export the file to the Excel format?\nValid entries: "yes" or "no"\n')
        valid_input = user_input in {'yes', 'no'}    
    
    if user_input == 'yes':
        #exports csv data to the excel format file
        wb = Workbook()
        ws = wb.active
        excel_filename = filename[: -4] + '.xlsx'
        xlsx_abs_path = os.path.join(curr_dir, "PDF_parser", "Parsed", excel_filename)
        with open(csv_abs_path, 'r', encoding='UTF8') as f:
            for row in csv.reader(f):
                ws.append(row)
        wb.save(xlsx_abs_path)

    print(f'\n"{filename}" has been successfully parsed')     


if __name__ == "__main__":
    '''
    Choose which file to parse and run (main) 
    '''
    print('<<< SCRIPT START >>>\n')

    #ask user to choose the dictionary for parsing
    valid_input = False
    while not valid_input:
        user_input = input('Choose which dictionary to parse\nValid entries: "3000" or "5000"\n')
        valid_input = user_input in {'3000', '5000'}
    
    filename = 'The_Oxford_' + user_input + '.pdf'
    
    main(filename)

    print('\n<<< SCRIPT END >>>\n')

