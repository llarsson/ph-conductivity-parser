ph-conductivity-parser
======================

Parses pH and conductivity records from a particular proprietary format and outputs it in Microsoft Excel format.

Usage
-----

1. Clone this respositry (alternatively, download the files manually).
2. Ensure that Python 2.x is installed, and that the [xlwt
   library](http://pypi.python.org/pypi/xlwt) is installed as well.
3. Run the program as a normal Python script, and supply a list of files
   to convert as command parameters.

Example:

    python PhConductivityParser.py file1.txt file2.txt

Running the above will cause file1.txt and file2.txt to be parsed and
the output will be in file1.txt.xls and file2.txt.xls.
