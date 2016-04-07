#!/usr/bin/python
"""
Takes a csv file containing paths to pdf files as a command-line argument and  
prints each of the pdf files on your default printer from Adobe Acrobat. Assumes
csv file has ONE pdf filename per line and that you have Adobe Acrobat installed.

USAGE: python smartPrint.py csv_filename
"""
import _winreg as winreg  
import time, os, subprocess  
import sys, csv, argparse

def smart_print(pdf):
    '''
    Parameters: pdf - type: string
    
    Prints the .pdf file pdf from Adobe Acrobat with the default system printer.
    '''

    # dynamically get path to AcroRD32.exe  
    AcroRD32Path = winreg.QueryValue(winreg.HKEY_CLASSES_ROOT,'Software\\Adobe\\Acrobat\Exe')  
    acroread = AcroRD32Path  
  
    print 'acroread = {0}'.format(acroread)  
  
    # the last set of double quotes leaves the printer blank, defaulting to the default printer for the system
    cmd = '{0} /N /T "{1}" ""'.format(acroread,pdf)
  
    # see what the command line will look like before execution  
    #print(cmd)  
  
    # open command line in a different process
    proc = subprocess.Popen(cmd)  
  
    # need sleep here so the command line has time to open the pdf and spool job to printer
    time.sleep(5)  
  
    # kill AcroRD32.exe from Task Manager
    os.system("TASKKILL /F /IM AcroRD32.exe")


def main():
    
    # ensure correct usage
    assert len(sys.argv) != 1, \
        "smartPrint.py takes one command-line argument: a csv file containing the paths to pdf files to be printed.\n" \
        "USEAGE: python smartPrint.py csv_filename"
    
    try:
        # get csv filename with paths to files to be printed
        input = open(sys.argv[1], 'rb')
        
        # try to read it
        reader = csv.reader(input)

    except IOError:
        print "Aaaaah what file are you feeding me :(", sys.exc_info()[0]
        raise

    else:
        # print 'em all
        for row in reader:
            
            # specify the path to pdf, preceded by r or R for string literal (otherwise need to escape backslashes)
            pdf = row[0]
            
            # print it
            smart_print(pdf)

    # kthxbai
    input.close()
    

if __name__ == "__main__":
    main()
