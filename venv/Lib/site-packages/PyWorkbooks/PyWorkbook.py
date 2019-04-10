#! /usr/bin/env python
# -*- coding: utf-8 -*-
"""
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
"""

# making sure entire dependency tree reloads during development (this IS Beta people!)
import DataHandler
reload(DataHandler)

from DataHandler import DataHandler, DataHandlerUser



from abc import abstractmethod

class WorkbookException(Exception):
   pass

class NotInitialized(WorkbookException):
   pass

class WriteProtectException(WorkbookException):
   pass
      
class PyWorkbook(DataHandler):
   '''
   This is a base class for spreadsheet classes to overload.  All you need to sucessfully use this
   to interface with your own spreadsheet (or any data-base with three dimensional addressing) is:
   
   ************
   Some way (ANY WAY) to change and retrieve from a cell to/from an arbitrary sheet/object/dimension
   ************
   
   And that's it.  If you can change or retrieve a cell on an arbirary sheet/dimension in ANY WAY, 
   then you should be set to have the same functionality as GnWorkbook and ExWorkbook
   see the link posted at the top for the latest updates on this class
   
   Need to overload:
   * __init__ -- must call "_update_sheets_" when your module starts 
      (and _update_workbooks_ if you have it)
   
   * The Sheets class
      The Sheets object __getitem__ takes any value (integer, string, or whatever else you want) and
      returns a "sheet object"  The sheet object has to have the following functionality:
         sheet[col, row].get_value()  # gets the value from the cell at col, row
         sheet[col, row].set_text(string)  # sets the value at col, row to string
         
         OR
         
         you can overload the get_item and set_item to do what you need done -
         but you will probably still need a Sheet class.  See ExWorkbook for examples
   
   * _update_sheets_ - overload so that it will create the following objects:
      self.sheet_objects -- a list or tuple containing all possible sheet objects in order
      self.sheets_library -- a dictionary containing all sheet objects with their text-names as keys
      self.sheets = self.Sheets(self) -- the above Sheets object
      self.current_sheet_object -- the sheet object currently being addressed
      self.current_sheet_name -- the name of the sheet object currently being addressed
      
      # you might also want to define the following, if you have this functionality
      * for these to work you need to make the sheets class return a workbook.sheet object
         (the object needs to specify the workbook), or your sheet objects need to not ever change
      self.workbook_objects
      self.workbooks_library
      self.workbooks   -- an instance of the Workbooks class
      self.current_workbook_object
      self.current_workbook_name
      
      See the class "GnWorkbook" or "ExWorkbook" for examples on how to implement all of this.
   
   * self.change_sheet -- needs to change the current_sheet_object and current_sheet_name to
         their correct values
         
   ###  PLEASE copy-paste the code in either GnWorkbook or ExWorkbook and edit it to your liking!
   '''
   protected_sheets = []
   
   ##############################################
   ## Overload These - E
   # 
   # You don't NEED to overload Sheets, but you WILL need to make something nearly identical.
   # here is the skeleton.  See ExWorkbook and GnWorkbook for examples
   class Sheets(object):
      '''Sheet class works like dictionary that filters the input and returns proper
      sheet object (and workbook object if your application can work with multiple
      workbooks.'''
      def __init__(self, workbook): self.workbook = workbook
      
      def __getitem__(self, sheet): pass
   
      def sheet_index(self, sheet): pass
   
   def __init__(self, child):
      ''' you must call this in your own function, passing in self'''
      pass #self.child = child
      
   @abstractmethod
   def change_sheet(self, sheet):
      raise NotImplementedError
   
   @abstractmethod
   def _update_sheets_(self):
      raise NotImplementedError
   
   ##
   ##############################################
   
   ##########################
   ## You might want to override the following
   def _interpret_nRange_(self, nRange):
      raise NotImplementedError("Invalid input request {0}".format(nRange))
      '''Finds the starting row, col and ending row,col of a range of cells
      and returns two tuples (start_row, start_col), (end_row, end_col)
      If you are building off of this class, overload this function to handle
      your own cell range input objects'''
   
   def _update_workbooks_(self):
      raise NotImplementedError
   
   def change_workbook(self, workbook):
      raise NotImplementedError
      
   def _set_point_(self, row, col, value, sheet):
      sheet[col, row].set_text(str(value))

   def _get_point_(self, row, col, sheet):
      return sheet[col, row].get_value()
   
   def _get_slices_from_nRange_(self, nRange):
      '''If you have special inputs, like a range of cells'''
      raise NotImplementedError
   
   '''
   These functions are called by __getitem__ in data handler.  It will pass in any
   variables it received beyond two.  In the case of GnWorkbook or ExWorkbook, args
   will contain either none or the variable to determine the sheet.
   
   If the spreadsheet program you are trying to design for can access arrays of data
   faster than single points, you can override these to improve performance (as I have
   done for ExWorkbook)
   '''
   def _set_single_point_(self, row, col, value, args):
      sheet = self._find_sheet_(args)
      return self._set_point_(row, col, value, sheet)
   def _get_single_point_(self, row, col, args):
      sheet = self._find_sheet_(args)
      return self._get_point_(row, col, sheet)
   
   def _set_array_(self, rows, cols, value, args): 
      sheet = self._find_sheet_(args)
      return DataHandler._set_array_(self, rows, cols, value, sheet)
   def _get_array_(self, rows, cols, args, generate = True): 
      sheet = self._find_sheet_(args)
      return DataHandler._get_array_(self, rows, cols, sheet, generate)
   
   def _set_matrix_(self, rows, cols, value, args): 
      sheet = self._find_sheet_(args)
      return DataHandler._set_matrix_(self, rows, cols, value, sheet)
   def _get_matrix_(self, rows, cols, args):    # do not override
      sheet = self._find_sheet_(args)           # returns generators (which are great)
      return DataHandler._get_matrix_(self, rows, cols, sheet)
   ##
   ############################
   
   #######################
   # you probably don't want to override these
   def _find_sheet_(self, args):
      if type(args) is tuple:
         sheet, = args
         if sheet is None: sheet = self.current_sheet_object
         else: sheet = self.sheets[sheet]
      elif type(args) == type(self.current_sheet_object):
         sheet = args   # may have been converted by a previous find_sheet call/invalid
      else: raise Exception('internal error')
      return sheet
   
   def protect_sheet(self, protect_sheet):
      protect_sheet = self.sheets[protect_sheet] # get the sheet object
      if protect_sheet not in self.protected_sheets:
         self.protected_sheets.append(protect_sheet)
   
   def unprotect_sheet(self, unprotect_sheet):
      '''not yet implemented'''
      unprotect_sheet = self.sheets[unprotect_sheet] # get the sheet object
      if unprotect_sheet in self.protected_sheets:
         del self.protected_sheets[self.protected_sheets.index(unprotect_sheet)]
   
   def change_cellref(self, item0, item1 = None):
      '''input a range in the form of two ints, two tuples, or a range of cells'''
      if item1 is not None:
         DataHandler.change_cellref(self, item0, item1)
      else:
         start, end = self._interpret_nRange_(item0)
         (self.start_row, self.start_col) = start
         (self.end_row, self.end_col) = end
   
   def __str__(self):
      return ("%s : sheetname = '%s', cellref = (row %s, col %s)"
              % (self.__repr__(), self.current_sheet_name, self.start_row, self.start_col,))
   
   def _interpret_slices_(self, item, data = None):
      '''overrides member class function to allow for nRange items'''
      if type(item) in (list, tuple, str):   # i.e. standard
         return DataHandler._interpret_slices_(self, item, data)
      else:
         return DataHandler._interpret_slices_(self, self._get_slices_from_nRange_(item), data)
   
   def rowcol_to_string(self, row, col):
      '''converts a row or column to spreadsheet indicies
      '''
      rowName = str(row + 1)
      dividend = col + 1
      columnName = ''
      while dividend > 0:
         modulo = (dividend - 1) % 26
         columnName = str(chr(ord('A') + modulo)) + columnName
         dividend = int((dividend - modulo) / 26)
      
      return columnName + rowName
      
   def rowscols_to_string(self, rows, cols, data):
      '''for converting from rows cols (the internal standard) to a spreadsheet index.
      ex: (1,2,3) , (0,1,2,3) => "A2:D4"
      '''
      # first re-find splice
      stepping = False
      if rows[-1] - rows[0] != len(rows) - 1:
         stepping = True
         return None, stepping
      if cols[-1] - cols[0] != len(cols) - 1:
         stepping = True
         return None, stepping
      
      temp_rows, temp_cols = self._determine_rowscols_(rows, cols, data)
      
      startrow, startcol = temp_rows[0], cols[0]
      endrow, endcol = temp_rows[-1], temp_cols[-1]
      
      # now convert to string
      cellrange = (self.rowcol_to_string(startrow, startcol) + ':'
                    + self.rowcol_to_string(endrow, endcol))
      return cellrange, stepping # string range, horizontal step, verticle step


'''Functions for external User classes to call'''
def _getsetfunc_USER_(self, items, sheet, value, getting):
  if getting:
     return DataHandler.__getitem__(self, items, sheet)
  else:
     sheetname = self.sheets.sheet_name(sheet) # passes the string on so it is pickleable.
                                                # this will (unfortunatley) make it so that
                                                # the user can't change sheet names, but oh
                                                # well
                                                
     return DataHandlerUser.__setitem__(self, items, value, sheetname)  # function will then call _set_point_, passing on sheet in a tuple

def __getitemUSER__(self, lots_of_items, sheet = None):
   ''' in the user routines, Data handler will call __getitem__ before it writes, but it
   will pass external arguments (i.e. the sheet object).  However, PyWorkbook's __getitem__ routine
   expects the sheet to be specified by the user, so it will be in lots_of_items.
   What this means is that if sheet exists, it has to be packaged in a tuple before it moves on.'''
   if sheet is not None:
      lots_of_items += (sheet,)
   self._update_sheets_()
   return self.getitems_setitems_base(lots_of_items, None, True)

   
def PyWorkbookUser(DataHandlerUser):
   '''overrides  _getsetfunc_ with something that uses DataHandlerUser 
   instead of DataHanlder for setting items (needed to undo), plus inherits
   all the other functionality of DataHandlerUser (undo and redo)'''
   pass
   
def __for_doctest():
   class tester(PyWorkbook):
      def change_sheet(self, sheet): pass
      def _update_sheets_(self): pass
   return tester(None)

if __name__ == '__main__':
   import doctest
   doctest.testmod()
    
