# tm-dups
Finds duplicate TM entries in Excel list created with Olifant

Created in Python 3.7.1

Non-standard dependent libraries: openpyxl

I wrote this script to help me find entries in a very large Trados translation memory where the same source segment had multiple target segments with attributes indicating who translated them. It's written to be used on a Trados translation memory that has been converted into an Excel file using the program Olifant, with the source segment in column 4, the target segment in column 5, and the attribute in column 6. 

The duplicate segment pairs and their attributes are output in a new Excel file so that they can be easily viewed and submitted to an editor/client to determine which translations should be used.
