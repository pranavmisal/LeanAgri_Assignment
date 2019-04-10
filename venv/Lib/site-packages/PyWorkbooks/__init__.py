'''
Copyright (C) 2011 by Garrett Berg - cloudform511@gmail.com
This code is part of PyWorkbooks Project under the MIT license .
A copy of the MIT license  should be included in the install folder where
you obtained this file.

Obtain PyWorkbooks at https://sourceforge.net/projects/pyworkbooks/files/ 

This Package has several Modules:

ExWorkbook: Module containing the class "ExWorkbook" for seamlessly interfacing
with open Excel spreadsheets

GnWorkbook: Module containing the class "GnWorkbook" for seamlessly interfacing
with open Gnumeric spreadsheets

PyWorkbook: Module containing the base class for the above classes, as well as
documentation strings for helping you make your own.

DataHandler: Base data handling module for all of the above classes.  Can interface
natively with 2D lists or tuples (cannot change values in tuples though).


Other classes not mentioned above are not fully tested and not recommended for
general use.
'''
author = 'Garrett Berg'
website = 'https://sourceforge.net/projects/pyworkbooks/files/'

import DataHandler, PyWorkbook, GnWorkbook, ExWorkbook
