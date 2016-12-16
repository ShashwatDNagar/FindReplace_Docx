# Find and Replace for Docx documents

Two scripts (in *Python* and *Perl*, choose either) that find and replace all instances
of a word in a Docx document.

The two scripts use system commands to unzip, modify, and re-zip the Docx file with the desired substitutions.

Currently, the scripts only support words (not phrases - spaces in the find and replace will yield unexpected results). Functionality that deals with this will be added soon.

## Input

Please specify the input files in the list/array called _fileList_.

Change the value of the variable _find_ with the word to be replaced and the
variable _replace_ with the replacement word.

*DO NOT USE WORDS WITH SPACES*

## Output

The input file will be edited in-place. Please make a copy of the file you are
editing before running these scripts on it.

### Future Functionality

 - Take directory as user input
 - Take _find_ as user input
 - Take _replace_ as user input
