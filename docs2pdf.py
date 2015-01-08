#!C:\Python27\python.exe
#---------------------------------------------------------------------------
# Copyright Aeronautics sys. 2015
# All Rights Reserved
#---------------------------------------------------------------------------
#########################################################################################
# TITLE: docx2pdf.py
#
# Important:
# This script works on win OS with additional installed 'pywin32' module
#
# Description: Find DOCX document in recursive way through the source path.
# and convert them to PDF format with the same fylesystem structure in the
# output given path.
# Writen by: Ori Tolev (oritol@aeronautics-sys.com)
# Date: 07/01/2015
##########################################################################################

from os import chdir,path
from time import strftime
from win32com import client
import os
import argparse
import fnmatch
import ntpath
import platform


def find(pattern, root_path, dest_path):

        for root, dirs, files in os.walk(root_path):
             head, tail = ntpath.split(root_path)
             chdir(root)
             diff = os.path.relpath(root, root_path)
             target_path = os.path.join(dest_path, diff)
             if not os.path.exists(target_path):
                os.makedirs(target_path)
             else:
                print(" Directory {0} already exists".format(target_path))
                exit()
             word = client.DispatchEx("Word.Application")
             for name in files:
                 if fnmatch.fnmatch(name, pattern):
                     in_file = path.abspath(root + "\\" + name)
                     try:
                         os.rename(in_file,in_file+"_")
                         print "Access on file \"" + str(in_file) +"\" is available!"
                         os.rename(in_file+"_",in_file)
                     except OSError as e:
                         message = "Access-error on file \"" + str(in_file) + "\"!!! \n" + str(e)
                         print message
                         break
                     try:
                        word.Visible = False
                        new_name = name.replace(".docx", r".pdf")
                        print strftime ("%H:%M:%S"), " Found docx ",  path.abspath(in_file)
                        new_file = path.abspath(target_path + "\\" + new_name)
                        if not os.path.isfile(new_file):
                            doc = word.Documents.Open(in_file)
                            print strftime("%H:%M:%S"), " Saving pdf ... ",  path.abspath(new_file)
                            doc.SaveAs(new_file, FileFormat = 17)
                            doc.Close()
                     except Exception, e:
                        message = "Converting  file error \"" + str(in_file) + "\"!!! \n" + str(e)
                        print message
             word.Quit()
        return

parser = argparse.ArgumentParser(description="Docx2pdf")
parser.add_argument("-s", metavar='c:\Version_3.5', help="Source DOCX")
parser.add_argument("-d", metavar='c:\PDF_Version_3.5', help="Destination PDF")
platform = platform.system()
if (platform != "Windows" ):
    exit()
args = parser.parse_args()
find('*.docx', args.s, args.d)
