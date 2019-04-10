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


Purpose:
Seamless pythonic interfacing between any open excel document and any python instance


Common usage:
>>> B = ExWorkbook()
>>> B.change_workbook('Book1')
>>> B.current_workbook_name
u'Book1'
>>> B.change_sheet('Sheet1')
>>> mymatrix = [[n*p for n in xrange(100)] for p in xrange(100)]  
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
>>> B[1, :1000] = ['what', 'is', 'a', 'man?']
>>> print B[1, :5]
[u'what', u'is', u'a', u'man?', 1.978]
'''

import sys

from PyWorkbook import PyWorkbook, PyWorkbookUser

if sys.platform == 'win32':		# why must it be an == sign???  I just don't know... I thought imutable types could use 'is'
   from win32com.client import Dispatch
   from win32com.client import GetActiveObject

try:
   import psyco
   psyco.full()
except ImportError:
   pass

class GnumericException(Exception):
   pass

class WriteProtectException(GnumericException):
   pass

class ExWorkbook(PyWorkbook):
   ''' this is the ExWorkbook for interfacing with an excel file.  See the lines of code in dev and dev2
   to get a basic idea of common use.
   '''
   
   def __init__(self):
      if sys.platform != 'linux2':
         self.xl_window = GetActiveObject("Excel.Application")
         self.xlApp = Dispatch(self.xl_window).Application
         self.current_workbook_object = None
         self._update_workbooks_()
         self.current_sheet_object = None
         self._update_sheets_()
      
   class Sheets(object):
      ''' Sheet class works like dictionary that filters the input and returns proper
      sheet object tied to the correct workbook object'''
      def __init__(self, workbook): 
         self.workbook = workbook
      
      def __getitem__(self, sheet):
         if type(sheet) is int:
            return self.workbook.sheet_objects[sheet]
         elif type(sheet) in (str, unicode):
            return self.workbook.sheets_library[unicode(sheet)]
         elif sheet in self.sheet_objects:  # if it is in the sheets
            return sheet   # simply return the sheet.
         
         raise NotImplementedError(sheet)
   
      def sheet_index(self, sheet):
         sheet = self.__getitem__(sheet)
         return self.workbook.sheet_objects.index(sheet)
   
   def change_sheet(self, sheet):
      self._update_workbooks_()
      self.current_sheet_object = self.sheets[sheet]
      self.current_sheet_name = self.current_sheet_object.Name
   
   def change_workbook(self, workbook):
      self._update_workbooks_()
      if type(workbook) is int:
         new_workbook = self.workbook_objects[workbook]
      elif type(workbook) in (str, unicode):
         new_workbook = self.workbooks_library[unicode(workbook)]
      elif type(workbook) == type(self.workbook_objects[0]):  # if it is sheet-type
         new_workbook = workbook
      
      self.current_workbook_object = new_workbook
      self.current_workbook_name = new_workbook.Name
   
   def create_sheet(self, sheetname):
      if unicode(sheetname) in self.sheets_library.keys(): raise ValueError('sheet already exists')
      sheet = self.current_workbook_object.Worksheets.Add()
      sheet.Name = unicode(sheetname)
      self._update_sheets_()
   
   def rename_sheet(self, oldsheet, newsheetname):
      if unicode(newsheetname) in self.sheets_library.keys(): raise ValueError('sheetname already exists')
      oldsheet = self.sheets[oldsheet]
      oldsheet.Name = unicode(newsheetname)
      self._update_sheets_()
      
   def _update_workbooks_(self):
      '''updates the list of all workbook objects available'''
      self.workbook_objects = []
      self.workbooks_library = {}
      for workbook in self.xlApp.Workbooks:
         self.workbook_objects.append(workbook)
         self.workbooks_library[workbook.Name] = workbook
      
      if self.current_workbook_object not in self.workbook_objects:
         self.current_workbook_name = self.xlApp.ActiveWorkbook.Name
         self.current_workbook_object = self.workbooks_library[self.current_workbook_name]
         
   def _update_sheets_(self):
      '''updates the sheets on the current workbook object'''
      self.sheet_objects = []
      self.sheets_library = {}
      
      for sheet in self.current_workbook_object.Worksheets:
         self.sheet_objects.append(sheet)
         self.sheets_library[sheet.Name] = sheet
      self.sheets = self.Sheets(self)
         
      if self.current_sheet_object not in self.sheet_objects:
         self.current_sheet_object = self.sheets[0]
         self.current_sheet_name = self.sheets[0].Name
   
   def _interpret_nRange_(self, nRange):
      '''Finds the starting row, col and ending row,col of a range of cells
      and returns two tuples (start_row, start_col), (end_row, end_col)
      If you are building off of this class, overload this function to handle
      your own cell range input objects'''
      # eventually I want this to be able to select the currently active cells if they
      # pass in a None value
      raise NotImplementedError
      
   def _set_point_(self, row, col, value, sheet):
      sheet.Cells(row + 1, col + 1).Value = value

   def _get_point_(self, row, col, sheet):
      return sheet.Cells(row + 1, col + 1).Value
   
   def _set_single_point_(self, row, col, value, args):
      sheet = self._find_sheet_(args)
      return self._set_point_(row, col, value, sheet)
   def _get_single_point_(self, row, col, args):
      sheet = self._find_sheet_(args)
      return self._get_point_(row, col, sheet)
      
   ## fun over-rides to increase performance
   def _get_array_(self, rows, cols, args, generate = True):
      sheet = self._find_sheet_(args)
      rangeString, stepping = self.rowscols_to_string(rows, cols, None)
      
      if stepping is True:
         return PyWorkbook._get_array_(self, rows, cols, args, generate)
      
      if len(cols) is 1:
         data = sheet.Range(rangeString).Value
         myiter = (n[0] for n in data)
         return self.convert_to_dtype(myiter, length = len(data))
         
      if len(rows) is 1: # vector
         return self.convert_to_dtype(sheet.Range(rangeString).Value[0])

      else: raise Exception('internal error')   # it is a matrix and should have been called
         # by the matrix routine
   
   def _set_array_(self, rows, cols, value, args):
      # note that value is a generator or a single item.
      sheet = self._find_sheet_(args)
      if hasattr(value, 'next'): value = tuple(value)   # iterators not allowed.  We are going to need to anyway
      rangeString, stepping = self.rowscols_to_string(rows, cols, value)
     
      if not hasattr(value, '__iter__'):  # single item/string
         if stepping is False:
            sheet.Range(rangeString).Value = value
            return
         else:
            return PyWorkbook._set_array_(self, rows, cols, value, args)
      
      else:
         if stepping is False:
            sheet.Range(rangeString).Value = value
            return
         else:
            return PyWorkbook._set_array_(self, rows, cols, value, args)
   
   def get_active_cell(self):
      cellstr = str(self.xlApp.ActiveCell.Address.replace('$', ''))
      return self.strindex_to_stdindex(cellstr)
   
   def get_active_sheet(self):
      return unicode(self.xlApp.ActiveSheet.Name)
   
#   def _set_matrix_(self, rows, cols, value, args):
#      return self._set_array_(rows, cols, value, args)


   
   

def dev1():
   B = ExWorkbook()
   B.change_sheet('Sheet1')
   B.change_dtype(list)
   B[1, :1000] = ['index', 'Ith', 'L(Iop)', 'Lmax', 'Vth', 'SE(Iop)', 'SEmax',
                                 'ISEmax', 'dSEdImax', 'WPEmax', 'IWPEmax', 'LWPEmax', 'WPEIop']
   data = B[1, :20]
   B.change_dtype(float)
   print data
   
if __name__ == '__main__':
   import doctest
   #dev1()
   print doctest.testmod()

