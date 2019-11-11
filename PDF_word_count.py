#!/usr/bin/env python3

# TODO find a better way of dealing with encrypted files.

import PyPDF2
import glob
import xlsxwriter
import time


# ------------------------ Input file path ------------------------ #
# Note use \**\ to access subfolders
input_file_path = r"C:\input_file_path\*.pdf"


# ------------------------ Output file path ----------------------- #
output_file_path = r"C:\Users\goldsby_c\Documents\myPython\Pearson\word_count.xlsx"
 

def globing():
    print('PyPDF2 import complete')
    glob_list = []
    for pdfFileObj in glob.iglob(input_file_path):
        glob_list.append(pdfFileObj)    
    return glob_list   
        
        
def main():
    # start timer
    start = time.time() 
    # lists
    total_words_list = []
    total_pages_list = []
    file_name_list = []
    from_glob = globing() 
    for location in from_glob:         
        try:
            # save the file path to list
            file_name_list.append(location)
            ##To get a PdfFileReader object that represents this PDF, call PyPDF2.PdfFileReader() and pass it pdfFileObj. Store this PdfFileReader object in pdfReader.
            pdfReader = PyPDF2.PdfFileReader(location)
            number_of_pages = pdfReader.getNumPages()
            saved_text = []
            # using the 'pdfreader.pages' method saves time and lines of code.
            for page in pdfReader.pages:
                # Once you have your Page object, call its extractText() method to return a string of the page’s text ❸. Note the text extraction isn’t perfect. 
                page_content = page.extractText()
                # appned to list
                saved_text.append(page_content)
            # count the total number of pages on the document
            total_pages = len(saved_text)
            # add number of pages from document 'total_pages' to total_pages_list which will keep a record for all documents.
            total_pages_list.append(total_pages)
            # print number of pages
            print("total number of pages are: ", total_pages)
            # create new var
            word_count_total = 0
            # loop count words and output total words on document
            for y in saved_text:
                # Generator to total up the count for each PDF.
                word_count_total = sum(len(page_text.split()) for page_text in saved_text)
            # takes total word count for each document and adds to a list
            total_words_list.append(word_count_total)
            # Print the count to terminal
            print("Total word count for your file is: ", word_count_total)
            # Print file path of PDF to terminal
            print(location)

            ##### -----  stop timer ------ ######
            end = time.time()
            minutes, seconds = divmod(end-start, 60)
            print("[m:s.ms] [{:0>2}:{:05.2f}]".format(int(minutes),seconds))
            
            
        except Exception: 
            print("------there was a problem with the file, it will be skipped-------")
            globing() # Deals with ENCRYPTED files.
    print("next to save...")
    return[total_words_list, total_pages_list, file_name_list]

# ----------------- function to save to file ------------------------ #
def save_to_file (output_file_path):
    # Pulls the retuned values from main() through.
    returned_values = main()
    #print (returned_values)
    print("output to excel started")
    # makes a new excel file
    workbook = xlsxwriter.Workbook(output_file_path)
    # add_worksheet method called on workbook object
    worksheet = workbook.add_worksheet()
    # start from first row
    row = 1
    # loop word count, page count and file location then write to file
    for total_words_list, total_pages_list, file_name_list in zip(returned_values[0], returned_values[1], returned_values[2]):
        # write to file using write method
        worksheet.write(row, 0, total_words_list )
        worksheet.write(row, 1, total_pages_list )
        worksheet.write(row, 2, file_name_list )
        row += 1
    workbook.close()
    print("---output to excel file complete, program all finished---")

# call save_to_file rather than main.
save_to_file(output_file_path)

