#!/usr/bin/env perl

# Shashwat Deepali Nagar, 2016
# Jordan Lab, Georgia Tech

# Perl script to find and replace words/phrases in a Docx document

use strict;

my @fileList = ("SampleDoc.docx"); # Populate this list with filenames

my $find = "TestWord"; # Substitute with word to be replaced
my $replace = "Genius"; # Substitute with replacement

foreach my $file (@fileList) {
  system("cp $file sample.zip");
  system("unzip sample.zip");
  system("sed -i 's/$find/$replace/g' word/document.xml");
  system("zip -r opFile.zip word/ docProps/ \[Content_Types\].xml _rels/");
  system("cp opFile.zip $file");
  # Delete all generated files
  system("rm -r _rels/");
  system("rm -r docProps/");
  system("rm opFile.zip");
  system("rm sample.zip");
  system("rm -r word/");
  system("rm \[Content_Types\].xml");
}
