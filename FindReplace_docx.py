#!/usr/bin/env python

# Shashwat Deepali Nagar, 2016
# Jordan Lab, Georgia Tech

# Python script to find and replace words/phrases in a Docx document.

import os

inputFile = ["SampleDoc.docx"] # Populate this list with filenames.

find = "TestWord" # Substitute with word to be replaced
replace = "Genius" # Substitute with replacement

for files in inputFile:
    copyCmd = "cp %s sample.zip"%(files)
    print copyCmd, "\n"
    os.system(copyCmd)
    os.system("unzip sample.zip")
    sedCmd = "sed -i 's/%s/%s/g' word/document.xml"%(find, replace)
    print sedCmd, "\n"
    os.system(sedCmd)
    os.system('zip -r opFile.zip word/ docProps/ \[Content_Types\].xml _rels/')
    opCmd = "cp opFile.zip %s"%(files)
    os.system(opCmd)
    # Delete all generated files
    os.system("rm -r _rels/")
    os.system("rm -r docProps/")
    os.system("rm opFile.zip")
    os.system("rm sample.zip")
    os.system("rm -r word/")
    os.system("rm \[Content_Types\].xml")
