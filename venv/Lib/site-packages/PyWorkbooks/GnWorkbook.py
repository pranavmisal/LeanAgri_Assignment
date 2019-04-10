#! /usr/bin/env python
# -*- coding: utf-8 -*-

'''
Copyright (C) 2011 by Garrett Berg - cloudform511@gmail.com
This code is part of PyWorkbooks Project.  Obtain PyWorkbooks at
https://sourceforge.net/projects/pyworkbooks/files/ 

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.


EXAMPLE CODE / DOCTEST:

This example is also in the Standard Documentation with more comments

Note: to run this doctest, open a new Gnumeric document and save it as "Book1.gnumeric"
(should be automatic)

Then go to tools -> Python Console.
Import this module run the doctest on it:
import doctest
doctest.testmod(GnWorkbook)

Common usage:
>>> B = GnWorkbook()
>>> B.change_workbook('Book1.gnumeric')
>>> B.current_workbook_name
u'Book1.gnumeric'
>>> B.change_sheet('Sheet1')
>>> mymatrix = [[n*p for n in range(100)] for p in range(100)]  
>>> B[0,0] = mymatrix   # writing data
>>> data = B['A1:F20']    # reading data
>>> print len(data)      # gets a generator with a length
20
>>> print data.next()   # works like a generator
[0.0, 0.0, 0.0, 0.0, 0.0, 0.0]
>>> print len(data)
19
>>> print data[0]   # same as data.next()
[0.0, 1.0, 2.0, 3.0, 4.0, 5.0]
>>> print len(data)
18
>>> print data[:2]   # returns array of values
[[0.0, 2.0, 4.0, 6.0, 8.0, 10.0], [0.0, 3.0, 6.0, 9.0, 12.0, 15.0]]
>>> print len(data)
16
>>> throw = data[:]
>>> print len(data)
0
>>> data = B['A1: F20']
>>> len(data)
20
>>> print data[1]
[0.0, 1.0, 2.0, 3.0, 4.0, 5.0]
>>> len(data)
18
>>> data = B[0:20:2, 0:6]   # slicing
>>> len(data)   # half the length
10
>>> data.next()   
[0.0, 0.0, 0.0, 0.0, 0.0, 0.0]
>>> data.next()
[0.0, 2.0, 4.0, 6.0, 8.0, 10.0]
>>> len(data)
8
>>> B.change_dtype(list)   # changing data type
>>> B[:10, 1]
[0.0, 1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0, 9.0]
>>> B.change_dtype(tuple)
>>> B[:10, 1]
(0.0, 1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0, 9.0)
>>> B.change_dtype(int)   # numpy stuff
>>> B[:10, 1]
array([0, 1, 2, 3, 4, 5, 6, 7, 8, 9])
>>> B.change_dtype(float)
>>> B[:10, 1]
array([ 0.,  1.,  2.,  3.,  4.,  5.,  6.,  7.,  8.,  9.])
>>> B.change_dtype(int)
>>> data = B['A1: F20']
>>> data[:4]
array([[ 0,  0,  0,  0,  0,  0],
       [ 0,  1,  2,  3,  4,  5],
       [ 0,  2,  4,  6,  8, 10],
       [ 0,  3,  6,  9, 12, 15]])
>>> B.change_sheet('Sheet2')   # some additional simple tests
>>> B.current_sheet_name
u'Sheet2'
>>> B[4,4] = 90
>>> B.current_sheet_name
u'Sheet2'
>>> B[4,4]
90.0
>>> B.change_sheet(2)
>>> B.current_sheet_name
u'Sheet3'
>>> B[4,4,'Sheet2']
90.0
>>> B['E5', 'Sheet2']
90.0
>>> B.change_dtype(list)
>>> data = B[:20, :10, 'Sheet1']
>>> data[:3:2]   # stepping in datatype
[[0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0], [0.0, 2.0, 4.0, 6.0, 8.0, 10.0, 12.0, 14.0, 16.0, 18.0]]
>>> data = B[:20:2, :10, 'Sheet1']   # same as stepping in generator.  This one is faster though
>>> data[:2]   
[[0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0], [0.0, 2.0, 4.0, 6.0, 8.0, 10.0, 12.0, 14.0, 16.0, 18.0]]
'''
import sys
from time import time, sleep
   
try:
   import numpy as np
except ImportError: pass


import PyWorkbook
reload(PyWorkbook)
from PyWorkbook import PyWorkbook, _getsetfunc_USER_, __getitemUSER__

import DataHandler
reload(DataHandler)
from DataHandler import DataHandlerUser


class GnumericException(Exception):
   pass

class WriteProtectException(GnumericException):
   pass

class GnWorkbook(PyWorkbook):
   ''' this is the GnWorkbook for interfacing with gnumeric.  See http://cloudform511.posterous.com/
   for more information
   
   Typical use:
   def receive_workbook(nRange):   # pass in a cell range
      myWorkbook = GnWorkbook
      myWorkbook.change_sheet(nRange)   # change to sheet of that cell
      myWorkbook[0,0] = 'an external python wrote in cell A1"
      myWorkbook[4:12, 0] = 'an external python changed ALL of these!
      
      add_to_four = myWorkbook[4,4] + 4  # grab cell at E5 and add it to four
      return add_to_four                 # display the value
      
   member functions:
   change_sheet - pass in an integer, a string, or a cell array to change the sheet to that reference.
     defaults to the first sheet in the workbook
     
   change_cellref(tuple or nRange, use_end = False) - pass in a tuple or nRange to set the range of
      data to use.
      
   '''
   def __init__(self):
      self.Gnumeric = sys.modules['Gnumeric']
      '''
      # getting the working directory... very hackish, but cool
      
      sheet = self.Gnumeric.workbooks()[0].sheets()[0]
      prevtext = sheet[0, 0].get_entered_text()
      sheet[0, 0].set_text('=cell("filename", A1:A20)')      
      directory = str(sheet[0, 0].get_value())
      self.working_directory = '/'.join(directory[7:].split('/')[:-1]) + '/'
      sheet[0, 0].set_text(str(prevtext))      
      
      if self.working_directory not in sys.path: sys.path.append(self.working_directory)
      '''
      self.current_workbook_object = None
      self.current_workbook_name = None
      self.current_sheet_object = None
      self.current_sheet_name = None
      self._update_workbooks_()
      self._update_sheets_() # automatically done in _update_workbooks_, but put here anyway
      
      self.gFun = self.Gnumeric.functions
      
   class Sheets(object):
      ''' Sheet class works like dictionary that filters the input and returns proper
      sheet object'''
      def __init__(self, workbook): 
         self.workbook = workbook
      
      def __getitem__(self, sheet):
         if type(sheet) is int:
            return self.workbook.sheet_objects[sheet]
         elif type(sheet) in (str, unicode):
            return self.workbook.sheets_library[unicode(sheet)]
         elif type(sheet) == type(self.workbook.sheet_objects[0]):  # if it is sheet-type
            return sheet   # simply return the sheet.  This is useful only as a check
         else: # assume it is an nRange.
            sheetname = unicode(self.workbook.gFun['cell']('sheetname', sheet))   # gets the name of the current sheet from the range
            return self.workbook.sheets_library[sheetname]  # sets S equal to that sheet
   
      def sheet_index(self, sheet):
         sheet = self.__getitem__(sheet)
         return self.workbook.sheet_objects.index(sheet)
      
      def sheet_name(self, sheet):
         sheet = self.__getitem__(sheet)
         return unicode(sheet.get_name_unquoted())
   
   def change_sheet(self, sheet):
      self._update_sheets_()
      self.current_sheet_object = self.sheets[sheet]
      self.current_sheet_name = unicode(self.current_sheet_object.get_name_unquoted())
      
   def create_sheet(self, sheetname):
      if unicode(sheetname) in self.sheets_library.keys(): raise ValueError('sheet already exists')
      sheet = self.current_workbook_object.sheet_add
      sheet.rename(unicode(sheetname))
      self._update_sheets_()
   
   def rename_sheet(self, oldsheet, newsheetname):
      if unicode(newsheetname) in self.sheets_library.keys(): raise ValueError('sheetname already exists')
      oldsheet = self.sheets[oldsheet]
      oldsheet.rename(unicode(newsheetname))
      self._update_sheets_()
      
   def _update_sheets_(self):
      self.sheet_objects = self.Gnumeric.workbooks()[0].sheets()   # uses numerical indexes
      
      self.sheets_library = {}
      for sheet in self.sheet_objects:
         self.sheets_library[unicode(sheet.get_name_unquoted())] = sheet
      
      self.sheets = self.Sheets(self)
      
      if self.current_sheet_name not in self.sheets_library.keys():   # the objects themselves might change when shifting workbooks
         self.current_sheet_object = self.sheets[0]
         self.current_sheet_name = unicode(self.current_sheet_object.get_name_unquoted())
   
   def _get_slices_from_nRange_(self, nRange):
      start, end = self._interpret_nRange_(nRange)
      (start_row, start_col) = start
      (end_row, end_col) = end
      return slice(start_row, end_row + 1), slice(start_col, end_col + 1)
   
   def _interpret_nRange_(self, nRange):
      '''Finds the starting row, col and ending row,col of a range of cells
      and returns two tuples (start_row, start_col), (end_row, end_col)
      If you are building off of this class, overload this function to handle
      your own cell range input objects'''
      row = self.gFun['row'](nRange)
      col = self.gFun['column'](nRange)
      
      if type(row) in (float, int):
         start_row = int(row) - 1
         end_row = start_row
      else: # else it is a two deep array.  Super hackish
         start_row = int(row[0][0]) - 1 
         end_row = int(row[0][-1]) - 1
         
      if type(col) in (float, int):
         start_col = int(col) - 1
         end_col = start_col
         
      else: # else it is an array of arrays... why oh why did they do this?
         start_col = int(col[0][0]) - 1
         end_col = int(col[-1][0]) - 1  
   
      return (start_row, start_col), (end_row, end_col)

   def change_workbook(self, workbook, skipupdate = False):
      if skipupdate is False:
         self._update_workbooks_()
         
      if type(workbook) is int:
         new_workbook = self.workbook_objects[workbook]
      elif type(workbook) in (str, unicode):
         new_workbook = self.workbooks_library[unicode(workbook)]
      elif type(workbook) == type(self.workbook_objects[0]):  # if it is workbook-type
         new_workbook = workbook
      else: raise IndexError(workbook)
   
      self.current_workbook_object = self.workbook_objects[0]
      for workbook_name, workbook_object in self.workbooks_library.items():
         if workbook_object == self.current_workbook_object:
            self.current_workbook_name = workbook_name
            break
      else: raise Exception('internal error')
      
      sheet = self.current_workbook_object.sheets()[0]
      prevtext = sheet[0, 0].get_entered_text()
      sheet[0, 0].set_text('=cell("filename", A1:A20)')      
      directory = str(sheet[0, 0].get_value())
      self.working_directory = '/'.join(directory[7:].split('/')[:-1]) + '/'
      sheet[0, 0].set_text(str(prevtext))
      
      self._update_sheets_()
            
   def _update_workbooks_(self):
      '''updates the list of all workbook objects available'''
      self.workbook_objects = []
      self.workbooks_library = {}
      for workbook in self.Gnumeric.workbooks():
         self.workbook_objects.append(workbook)
         # have to do a hack to get the name
         sheet = workbook.sheets()[0]

         prevtext = sheet[0, 0].get_entered_text()
         sheet[0, 0].set_text('=cell("filename", A1:A20)')   
         sleep(.005)   
         filepath = str(sheet[0, 0].get_value())
         filename = ''.join(filepath.split('/')[-1])
         sheet[0, 0].set_text(str(prevtext))
         
         self.workbooks_library[unicode(filename)] = workbook
      
      
      if self.current_workbook_object not in self.workbook_objects:
         self.change_workbook(0, skipupdate = True)
      else:
         self._update_sheets_()
   

class GnWorkbookUser(GnWorkbook, DataHandlerUser):
   ''' CURRENTLY IN DEVELOPMENT
   Extends the GnWorkbook class to allow for undo and redo operations.
   Mostly useful for interactive sessions.  Slows down setting operations
   significantly.'''
   def _getsetfunc_(self, items, sheet, value, getting):
      return _getsetfunc_USER_(self, items, sheet, value, getting)

   def __getitem__(self, lots_of_items, sheet = None):
      return __getitemUSER__(self, lots_of_items, sheet)
   
''' standard development

import sys; sys.path.append('/home/garrett/Projects/EclipseWorkspace/Playground/src/PyWorkbooksFiles/PyWorkbooks'); import GnWorkbook; reload(GnWorkbook); B = GnWorkbook.GnWorkbook()

Some Notes for the future:

dir(Gnumeric)
['Boolean', 'CellPos', 'FALSE', 'GnmStyle', 'GnumericError', 
'GnumericErrorDIV0', 'GnumericErrorNA', 'GnumericErrorNAME', 
'GnumericErrorNULL', 'GnumericErrorNUM', 'GnumericErrorREF', 
'GnumericErrorVALUE', 'Range', 'TRUE', '__doc__', '__name__', 
'__package__', 'functions', 'plugin_info', 'workbook_new', 
'workbooks']

dir(Gnumeric.workbooks()[0].sheets()[0])
['cell_fetch', 'get_extent', 'get_name_unquoted', 'rename', 
'style_apply_range', 'style_get', 'style_set_pos', 
'style_set_range']

>>> dir(Gnumeric.workbooks()[0])
['gui_add', 'sheet_add', 'sheets']
'''

