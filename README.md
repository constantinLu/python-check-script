# Python_checkScript
verification of the attribute table using python

 ## [1. checkScript](https://github.com/gizet/Python_checkScript/tree/master/check/original)

This script compares the input feature classes with the specified requirements and outputs an excel report with the problems.
It also checks the geometry of each shape and the result can be seen in the excel report (sheet3)

////

First Parameter: Input of the featureClasses
   - can contain a single featureClass (shape or gdb)
   - can contain a folder with shapeFiles
   - can contain a folder with shapeFiles and featureClass from a geodatabase.
 
////

Second Parameter:
   - is the Template excel which is already made
   - add the location of the Template excel

////

Third Parameter:
   - specify the location folder for the results
   - the script will generate a excel file for each featureClass


PYHON MODULE USED:

- arcpy
- os
- xlsx
- win32.client
- sys
