# -*- coding: utf-8 -*-
"""
Created on Tue Jun  1 10:43:44 2021

@author: CWilson
"""

import glob, os



def find_string_in_the_dir(string_to_find):
    function_name_as_str = "Production_Dashboard_temp_files"
    function_name_as_str = string_to_find
    # directory to search within
    pydir = 'c://users/cwilson/documents/python/'
    # all the files within the directory
    files = [y for x in os.walk(pydir) for y in glob.glob(os.path.join(x[0], '*.py'))]
    # function/string to find within the python files
    
    
    
    print('The string in question:\t\t"' + function_name_as_str + '"', end='\n\n')
    
    # dict with file names & line numbers the function_name_as_str appears on
    specific = {}
    
    for file in files:
        
        if file.endswith('find_string_within_python_directory.py'):
            continue
        
        with open(file) as f:
            
            lines = f.readlines()
            
            for i,line in enumerate(lines):
                
                if function_name_as_str in line:
                    
                    print('In the file:\t\t\t\t' + os.path.basename(file))
                    print('The string is on line #:\t' + str(i))
                    if file not in specific:
                        specific[file] = []
                        specific[file].append(i)
                        
                    else:
                        specific[file].append(i)
                

find_string_in_the_dir("by_pcmark['bins']")
