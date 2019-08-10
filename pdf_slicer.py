#!/bin/python
# A program to slice sections of pages from .PDF files

try:
    import PyPDF2
except:
    print("Error: PyPDF2 is not installed, yet is required.")
    exit()

import argparse

# Initialize variables
filename = ""
slice_start = 0
slice_split = 1
slice_step = 0
slice_stop = 0
verbose = False

def main():
    """ Main function """

    desc_str = """
    This program splits up a .pdf file into multiple pages.
    When run with no arguments, this program splits it into single pages.
    Starting page and number of pages per split can be set using arguments.
    """

    # Initialize the global variables
    global filename
    global slice_start
    global slice_split
    global slice_step
    global slice_stop
    global verbose

    parser = argparse.ArgumentParser(description=desc_str)

    # Add arguments
    parser.add_argument("filename", nargs=1, 
                        help="The name of the .pdf file to be processed.")
    parser.add_argument('-f', '--from', nargs=1, dest='start', 
                        type=int, required=False,
                        help="Starting page number, counting up from 0.")
    parser.add_argument('-p', '--pages', nargs=1, dest='split', 
                        type=int, required=False,
                        help="How many pages to slice the pdf into.")
    parser.add_argument('-s', '--skip', nargs=1, dest='step', 
                        type=int, required=False,
                        help="Set this to skip pages.")
    parser.add_argument('-t', '--to', nargs=1, dest='stop', 
                        type=int, required=False,
                        help="Page number to stop at, defaulting at the end.")
    parser.add_argument('-v', '--verbose', dest='verbose', action='store_true',
                        required=False,
                        help="Produces more feedback while running.")
    args = parser.parse_args()

    # Process arguments
    if args.verbose: 
        verbose = args.verbose
    filename = args.filename[0]
    if verbose:
        print("Splicing file {}.".format(filename))
    if filename[-4:].lower() != '.pdf':
        print("Supplied file is not recognized as a .pdf - exiting.")
        exit()    
    if args.start:
        slice_start = args.start[0] - 1
        if verbose:
            print("Starting at page {}".format(slice_start + 1))
    if args.split:
        slice_split = args.split[0]
        if verbose:
            print("Taking {} page(s) per slice.".format(slice_split))
    if args.step:
        slice_step = args.step[0]
        if verbose:
            print("Skipping {} page(s) in between splits.".format(slice_step))
    if args.stop:
        slice_stop = args.stop[0]
        if verbose:
            print("Stopping at page {}.".format(slice_stop))
    if args.verbose:
        verbose = args.verbose

    # Check for any possible conflicts
    if slice_start > slice_stop:
        print("Error: indicated start is after stopping point.")
        exit()

def pdf_slice(filename,start,stop,number):
    """ Splice the given file from start to stop, appending number to name """

    # Load the original file
    global pdf_original_file
    global pdf_original_reader

    # Intialize the writer
    pdf_writer = PyPDF2.PdfFileWriter()

    # Loop through pages and slice out a selection
    for pagenum in range(start, stop):
        # Snatch the single page
        page_obj = pdf_original_reader.getPage(pagenum)

        # Add the page to a file to write
        pdf_writer.addPage(page_obj)

    # Write the object to a file, using the file name minus extension
    pdf_output_file = open(filename[:-4] + str(number) + '.pdf', 'wb')
    pdf_writer.write(pdf_output_file)
    
    pdf_output_file.close()

# If run as a program, process command-line arguments
if __name__ == '__main__':
    main()

# Load the original file
pdf_original_file = open(filename, 'rb')
pdf_original_reader = PyPDF2.PdfFileReader(pdf_original_file)

# Set the stopping point at the total number of pager
if slice_stop == 0:
    slice_stop = pdf_original_reader.numPages

# Variable to track number to add to file
iteration = 0

# Set up variable to iterate through the PDF file
page_start = slice_start

# Loop through the page range
for page in range(slice_start,slice_stop,slice_split):
    iteration += 1

    # Calculate end of new slice, and limit if needed
    page_stop = page_start + slice_split
    if page_stop > pdf_original_reader.numPages:
        page_stop = pdf_original_reader.numPages

    # Call on the slice function to do the actual work
    if verbose:
        print("Slicing pages {} to {} from file.".format(page_start+1,page_stop))
    pdf_slice(filename,page_start,page_stop,iteration)
    
    # Increase the iteration
    page_start += slice_split + slice_step
    if page_start > pdf_original_reader.numPages:
        page_start = pdf_original_reader.numPages

    if page_stop == slice_stop:
        break

pdf_original_file.close()
