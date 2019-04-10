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


Typical usage
>>> data = [[p * n for n in range(20)] for p in range(10)]
>>> handler = DataHandler(data)

>>> handler[4, :15]
[0, 4, 8, 12, 16, 20, 24, 28, 32, 36, 40, 44, 48, 52, 56]

>>> handler[4:7, 1]
[4, 5, 6]
>>> handler[0:7, 3]
[0, 3, 6, 9, 12, 15, 18]
>>> handler[:10, :5][:4]
[[0, 0, 0, 0, 0], [0, 1, 2, 3, 4], [0, 2, 4, 6, 8], [0, 3, 6, 9, 12]]
>>> gen = handler[:10, :5]
>>> gen.next()
[0, 0, 0, 0, 0]
>>> gen[:4]
[[0, 1, 2, 3, 4], [0, 2, 4, 6, 8], [0, 3, 6, 9, 12], [0, 4, 8, 12, 16]]
>>> gen[:]
[[0, 5, 10, 15, 20], [0, 6, 12, 18, 24], [0, 7, 14, 21, 28], [0, 8, 16, 24, 32], [0, 9, 18, 27, 36]]
>>> handler.change_dtype(float)
>>> gen = handler[:10, :5]
>>> gen[:4]
array([[  0.,   0.,   0.,   0.,   0.],
       [  0.,   1.,   2.,   3.,   4.],
       [  0.,   2.,   4.,   6.,   8.],
       [  0.,   3.,   6.,   9.,  12.]])
>>> gen = handler[:10:3, :5:2]   # slice in two places
>>> gen[:3]
array([[  0.,   0.,   0.],
       [  0.,   6.,  12.],
       [  0.,  12.,  24.]])
>>> handler[0, :] = [123, 456, 789]   # empty addressing for arrays, doesn't work yet
>>> handler[0, :3]
array([ 123.,  456.,  789.])
>>> handler[1,1]
1
>>> handler.change_dtype(list)
>>> handler[0, 0:20] = ['hello', 'there', 'bob']
>>> strings = handler[0:3, :3]
>>> handler.change_dtype(float)
>>> print strings.next()
['hello', 'there', 'bob']


####################   DataHandlerUser
### Doesn't currently work !!!
Doctest for DataHandlerUser, with undo and redo functionality
(still in beta, do NOT use more than one instance at once.
Also, don't have any files called "undoredo.pyworkbooks" in
your folder that are important to you :D)

#Typical Usage:
#>>> mymatrix = [[p * n for n in xrange(0, 100)] for p in xrange(0, 100)]
#>>> handler = DataHandlerUser(mymatrix)
#>>> print handler[1, :10]
#[0, 1, 2, 3, 4, 5, 6, 7, 8, 9]
#>>> handler[1, :10] = 44
#>>> print handler[1, :10]
#[44, 44, 44, 44, 44, 44, 44, 44, 44, 44]
#>>> handler.undo()
#>>> print handler[1, :10]
#[0, 1, 2, 3, 4, 5, 6, 7, 8, 9]
#>>> print handler[:5, :5][:]
#[[0, 0, 0, 0, 0], [0, 1, 2, 3, 4], [0, 2, 4, 6, 8], [0, 3, 6, 9, 12], [0, 4, 8, 12, 16]]
#>>> smalmat = [[p * n for n in xrange(-10, 0)] for p in xrange(0, 10)]
#>>> handler[0, 0] = smalmat
#>>> print handler[:5, :5][:]
#[[0, 0, 0, 0, 0], [-10, -9, -8, -7, -6], [-20, -18, -16, -14, -12], [-30, -27, -24, -21, -18], [-40, -36, -32, -28, -24]]
#>>> handler[0, (0, 2, 5)] = 'I was changed here'
#>>> print handler[:5, :5][:]
#[['I was changed here', 0, 'I was changed here', 0, 0], [-10, -9, -8, -7, -6], [-20, -18, -16, -14, -12], [-30, -27, -24, -21, -18], [-40, -36, -32, -28, -24]]
#>>> handler.undo()
#>>> print handler[:5, :5][:]
#[[0, 0, 0, 0, 0], [-10, -9, -8, -7, -6], [-20, -18, -16, -14, -12], [-30, -27, -24, -21, -18], [-40, -36, -32, -28, -24]]
#>>> handler.undo()
#>>> print handler[:5, :5][:]
#[[0, 0, 0, 0, 0], [0, 1, 2, 3, 4], [0, 2, 4, 6, 8], [0, 3, 6, 9, 12], [0, 4, 8, 12, 16]]
#>>> handler.redo()
#>>> print handler[:5, :5][:]
#[[0, 0, 0, 0, 0], [-10, -9, -8, -7, -6], [-20, -18, -16, -14, -12], [-30, -27, -24, -21, -18], [-40, -36, -32, -28, -24]]
#>>> handler.redo()
#>>> print handler[:5, :5][:]
#[['I was changed here', 0, 'I was changed here', 0, 0], [-10, -9, -8, -7, -6], [-20, -18, -16, -14, -12], [-30, -27, -24, -21, -18], [-40, -36, -32, -28, -24]]
#>>> handler.undo()
#>>> print handler[:5, :5][:]
#[[0, 0, 0, 0, 0], [-10, -9, -8, -7, -6], [-20, -18, -16, -14, -12], [-30, -27, -24, -21, -18], [-40, -36, -32, -28, -24]]
#
#
## Doesn't yet work with User
#>>> handler[0, :] = [123, 456, 789]   # empty addressing for arrays, doesn't with User yet
#>>> handler[0, :3]
[123, 456, 789]
'''


try:
   import numpy as np
except ImportError: pass


from time import time
import shelve
import os
def reconstruct_iter(first, myiter): # reconstructing the iterator
   yield first
   while True:
      yield myiter.next()
               
class GeneratedData(object):
   '''generator-like object containing data from a matrix call to a DataHandler object.  
   Call len to retrieve the length, and get_rows to recieve the
   requested number of rows in your chosen data type (and subtracting the number of rows remaining).
   This code is intended to be the fastest possible way to perform these kinds of requests and
   has been tested fairly extensively.
   '''
   def __init__(self, workbook, iterated_data, dtype, length):
      self.__iterated_data = iterated_data.__iter__()
      self.__workbook = workbook  # data type to use
      self._dtype_ = dtype
      self.length = length
      
   def __iter__(self): return self
   
   def next(self):
      self.length -= 1
      
      if self._dtype_ != self.__workbook._dtype_:
         prevdtype = self.__workbook._dtype_
         self.__workbook._dtype_ = self._dtype_
         
         out = self.__iterated_data.next()
         
         self.__workbook._dtype_ = prevdtype
      else:
         out = self.__iterated_data.next()
      return out
   
   def __len__(self):
      '''returns the remaining amount of length.'''
      return self.length
   
   def __morenext(self, amount):
      # internal function used to clear a certain amount of data
      # note, doesn't keep track of length.
      for n in xrange(amount):
         self.__iterated_data.next()
         
   def get_data(self, amount, step = None):
      '''gets the data specified by amount and returns them in the specified data.  If all the (remaining) length
      are desired, set amount = -1 '''
      if amount == -1: amount = self.length
      elif self.length - amount < 0: raise IndexError('Not that much data left available')
      step = step if step is not None else 0 # convert None to 0
      
      if amount is 1:   # single element
         return self.next()
         
      if self._dtype_ in (int, float):
         if step is 0:
            myray = tuple((self.__iterated_data.next() for _n in xrange(amount)))
            rows = amount
         else:
            myray = []
            for n in xrange(0, amount, step):
               myray.append(self.__iterated_data.next())
               self.__morenext(step - 1)
               rows = amount / step
               
         self.length -= amount
         if type(myray[0]) in (int, str, float):
            return np.array(myray, dtype = self._dtype_)
         return np.reshape(np.concatenate(myray), (rows, len(myray[0])))
      
      if self._dtype_ is list:
         if step is 0:
            mylist = [self.__iterated_data.next() for _n in xrange(amount)]
         else:
            mylist = []
            for n in xrange(0, amount, step):
               mylist.append(self.__iterated_data.next())
               self.__morenext(step - 1)
         self.length -= amount   # done like this (instead of using self.next() for increased performance
         return mylist
      
      if self._dtype_ is tuple:
         if step is 0:
            mytup = tuple((self.__iterated_data.next() for _n in xrange(amount)))
         else:
            mytup = [] # yes I know it's a list... fastest way to do this I believe
            for n in xrange(0, amount, step):
               mytup.append(self.__iterated_data.next())
               self.__morenext(step - 1)
            mytup = tuple(mytup)
            
         self.length -= amount
         return mytup
         
   def __getitem__(self, item):
      if type(item) is int:
         self.length -= item
         self.__morenext(item)
         return self.get_data(1)
      
      if type(item) is slice:
         start, stop, step = item.start, item.stop, item.step
         start = start if start is not None else 0 # convert start = None to start = 0
         
         if start < 0: raise IndexError(item)
         if start != 0:
            self.length -= start
            self.__morenext(start - 1)
         if stop is None: return self.get_data(-1, step)
         if stop - start < 0: raise IndexError(item)
         return self.get_data(stop - (start if start is not None else 0), step)
   
   def __str__(self):
      return '<Generator like object> with {0} values remaining and dtype of {1}'.format(self.length, self._dtype_)
   
   def __repr__(self):
      return object.__repr__(self) + 'custom generator like object for DataHandler class'

class DataHandler(object):
   ''' 
   Allows for arbirary 2D data handling.  Replace the __init__ function and create two 
   functions in order to handle any set of arbirary data that you can address with two points
   
   These are the functions you need to re-write:  (currently created for handling lists)
   def list__get_point_(row, col):
         return self.data[row][col]
   def list__set_point_(row, col, value):
      self.data[row][col] = value
   
   '''
   # relative indexes
   start_row = 0
   start_col = 0
   end_row = 0
   end_col = 0
   
   _dtype_ = list   # standard return data type
   
   '''Set to True to make it output a generator for single array requests'''
   ALWAYS_GENERATE = False  # normally does not generate single arrays
   
   def __init__(self, input_data):
      self.data = input_data
      self.current_sheet_name = 'debuging with numpy!'
      
   
   def _get_point_(self, row, col, args):
      ''' Override this function to interface with your data'''
      return self.data[row][col]
   
   def _set_point_(self, row, col, value, args):
      ''' Override this function to interface with your data'''
      self.data[row][col] = value
   
   def _set_single_point_(self, row, col, value, args):
      '''different from _set_point_ in that you are guaranteed that the user wanted
      to set only ONE Point, whereas _set_point_ may be called many times sequentially by
      set array.  Use this to intercept calls and analyze whatever is in the 'args'
      variable'''
      return self._set_point_(row, col, value, args)
   
   def _get_single_point_(self, row, col, args):
      '''different from _get_point_ in that you are guaranteed that the user wanted
      to set only ONE Point.  _set_point_ can be called any number of times by
      set array.  Use this to intercept calls and analyze whatever is in the 'args'
      variable'''
      return self._get_point_(row, col, args)
   
   def parse_items(self, items):
      '''grabs valid items, passes rest on.  args will always be a tuple containing at least
      one item (None if there were None)'''
      if type(items) is str:
         args = (None,)
      elif len(items) > 1:
         if type(items[0]) is str:
            items, args = items[0], items[1:] 
         elif len(items) > 2: # first two are std input
            items, args = items[:2], items[2:]
         else: args = (None,)
      else: raise IndexError('Invalid input {0}'.format(items))
      return items, args
      
   def __getitem__(self, items):
      '''standard method call.  If len(items) > 2 they are packaged into args, and passed
      on to later functions.'''      
      items, args = self.parse_items(items)
      
      rows, cols = self._interpret_slices_(items)
      
      if len(rows) is 1 and len(cols) is 1: # it is a single element.  Return it
         return self._get_single_point_(rows[0]  , cols[0] , args)
      if len(rows) is 1:   # they want a single row, return a 1D array
         return self._get_array_(rows, cols, args)
      if len(cols) is 1:   # they want a single column, return a 1D array
         return self._get_array_(rows, cols, args)
      # else it is some kind of matrix
      return self._get_matrix_(rows, cols, args)
      
   def __setitem__(self, items, value):  
      '''standard method call.  If len(items) > 2 they are packaged into args, and passed
      on to later functions.'''
      items, args = self.parse_items(items)
      
      rows, cols = self._interpret_slices_(items, value)
      
      if not hasattr(value, '__iter__'):
         if len(rows) is 1 and len(cols) is 1: # it is a single element.  Change it
            self._set_single_point_(rows[0] , cols[0] , value, args)     
         elif len(rows) is 1:   # they want a single row, change the 1D array
            self._set_array_(rows, cols, value, args)
         elif len(cols) is 1:   # they want a single column, change the 1D array
            self._set_array_(rows, cols, value, args)
         else:
            self._set_matrix_(rows, cols, value, args)
      else: # the value is iterable
         ### This could be made faster.  I think there is a standard module for
         ### reconstructing iterators, and my method probably lowers performance
         if hasattr(value, 'next'): # if it is an itterator
            myiter = value.__iter__()  # 
            firstval = myiter.next()   # take a peek
            reiter = reconstruct_iter(firstval, myiter)
         else: 
            firstval = value[0]
            reiter = value
         if hasattr(firstval, '__iter__'): # then it is a matrix!
            self._set_matrix_(rows, cols, reiter, args) 
         elif len(rows) is 1 and len(cols) is 1: # it is a single element.  Bad unless matrix
            raise IndexError("must put in two corrdinates for array modifications (which direction do you want?")
         elif len(rows) is 1:   # change down the row, 
            self._set_array_(rows, cols, reiter, args)
         elif len(cols) is 1:   # change down the column
            self._set_array_(rows, cols, reiter, args)      
         else:
            raise IndexError("Only one coordinate must be sliced if not a matrix (to determine direction of array)")
   
   def _get_array_(self, rows, cols, args, generate = True):
      '''gets a range of values passed in in rows, cols format (they are arrays containing the rows needed
      and the cols needed).  In this case, it has to get the parts of the array point by point.
      Member functions can override this if they have a faster method to access arrays.
      
      - if generate is False it overrides ALWAYS_GENERATE (for use in matrix call programs)'''
      if len(rows) is 1:   # they want a single row, return a 1D array
         out = GeneratedData(self, (self._get_point_(rows[0] , col , args) for col in cols),
                             self._dtype_, len(cols))
         if self.ALWAYS_GENERATE == True and generate == True:
            return out
         else:
            return out[:]
         
      if len(cols) is 1:   # they want a single column, return a 1D array
         out = GeneratedData(self, (self._get_point_(row , cols[0] , args) for row in rows),
                             self._dtype_, len(rows))
         if self.ALWAYS_GENERATE == True and generate == True:
            return out
         else:
            return out[:]
      
      raise Exception('Internal Error')
   
   def _get_matrix_(self, rows, cols, args):
      ''' uses the _get_array_ function to get a generator for a matrix of values'''
      return GeneratedData(self, (self._get_array_((row,), cols, args, False) for row in rows)
                           , self._dtype_, len(rows))
      
   def _set_array_(self, rows, cols, value, args):
      ''' gets rows, cols each in tuple format, and a value as either a generator or
      a single value (int, string, float, etc).  Input is guaranteed to be an array
      (not a matrix)
      - args can contain whatever was additionally passed to __setitem__ in a tuple'''
      if not hasattr(value, '__iter__'): # it is a single value/string
         if len(rows) is 1 and len(cols) is 1: # it is a single element.
            raise Exception('internal error')     
         
         elif len(rows) is 1:   # they want a single row, change the 1D array
            for col in cols:
               self._set_point_(rows[0] , col  , value, args)
               
         elif len(cols) is 1:   # they want a single column, change the 1D array
            for row in rows:
               self._set_point_(row , cols[0]  , value, args)
            
         else: # it is a matrix, shouldn't be in here.
            raise Exception('internal error')
      
      else:
         if len(rows) is 1:   # change down the row, 
            for col_count, value_part in enumerate(value):
               self._set_point_(rows[0]  , cols[0] + col_count , value_part, args)
            
         elif len(cols) is 1:   # change down the column
            for row_count, value_part in enumerate(value):
               self._set_point_(rows[0] + row_count , cols[0]  , value_part, args)            
            
         else:
            raise IndexError("Only one coordinate must be sliced if not a matrix (to determine direction of array)")
   
   def _set_matrix_(self, rows, cols, value, args):
      ''' uses the _set_array_ function to set a matrix of values.  You may want to override'''
      if not hasattr(value, '__iter__'): # copy a single element over an array
         for row in rows:
            self._set_array_((row,), cols, value, args)
      else:    # the value itself is a matrix of values
         for row_count, row_data in enumerate(value): # note, cols has to specify the direction 
                                                      # to write in, which is why it is passed as it is
            self._set_array_((row_count + rows[0],), (cols[0], cols[0] + 1), row_data, args)
      
   def _interpret_slices_(self, item_in, data = None):
      '''internal member function to interpret a large variety of slice inputs'''
      if type(item_in) in (tuple, list):
         rows, cols = item_in
      elif type(item_in) is str:
         rows, cols = self.strindex_to_stdindex(item_in)
      else: 
         raise IndexError("invalid index")  
      
      error = IndexError('cannot input negative indecies')
      rows = self.stdindex_to_gen(rows, data)
      rows = tuple((n + self.start_row for n in rows))
      if min(rows) < 0: raise error
      
      cols = self.stdindex_to_gen(cols, data)      
      cols = tuple((n + self.start_col for n in cols))
      if min(cols) < 0: raise error
      
      # by the end they will all be tuples containing the necessary extractions
      
      return rows, cols
   
   def stdindex_to_gen(self, stdindex, data = None):
      '''converts from standard index to a tuple.  Also checks if data will
      determine the length'''
      if hasattr(stdindex, '__iter__'):
         outgen = stdindex.__iter__()
      else:
         if type(stdindex) is slice:
            start = stdindex.start if stdindex.start is not None else 0
            if stdindex.stop is None:   # the stopping point is none, only valid if data is a matrix/array
               if hasattr(data, '__iter__'):
                  outgen = (start, start + 1).__iter__()
               else: raise IndexError('Need a stop point for slice')
            elif stdindex.stop == start: raise IndexError("Cannot slice from the same place")
            else: outgen = xrange(start, stdindex.stop, (1 if stdindex.step is None else stdindex.step))
         elif type(stdindex) is not tuple:
            outgen = (stdindex,).__iter__()
      return outgen
   
   def strindex_to_stdindex(self, strindex):
      '''converts from standard spreadsheet string index into stdindex'''
      strindex = strindex.upper().replace(' ', '')
      
      item = strindex.split(':')
      output = []
      if len(item) > 3: raise IndexError(strindex)
      for index, cell_str in enumerate(item):
         if index > 1:  # if we are on the third element, then that is the step
            output.append(int(cell_str))
            break
         
         # cell_str will be, say 'AA100'.  Find the break point
         for n, char in enumerate(cell_str):
            if char < 'A':break
         
         row = int(cell_str[n:]) - 1
         
         col_list = list(cell_str[:n])
         col_list.reverse()   # put it so least value first
         col = 0
         for n, c in enumerate(col_list): # convert to number
            #col += (ord(c) - ord('A') + 1) * (26 ** n) - 1  
            col += (ord(c) - ord('A') + 1) * (26 ** n) - 1 + n  
            
         output.append((row, col))
         
      # we now have an output with length of 1, 2 or 3
      if len(output) is 1: # single point
         rows, cols = output[0]
         
      elif len(output) in (2, 3):   # start, end, ?step?
         # unpack the tuples
         if len(output) is 2:
            startpoint, endpoint = output
            step = None
         else:
            startpoint, endpoint, step = output
         
         row_st, col_st = startpoint
         row_end, col_end = endpoint
         
         if row_st == row_end: # same row, no slice
            rows = row_st
         else:
            rows = slice(row_st, row_end + 1, step)
         
         if col_st == col_end: # same col, no slice
            cols = col_st
         else:
            cols = slice(col_st, col_end + 1, step)
      return rows, cols
   
   def change_cellref(self, item0, item1):
      '''allows the item of a tuple to set the range.  Should be a 1D tuple if no
      end range is desired, or a 2D tuple if an endrange is desired
      example:
      change_cellref( (1,2) )  # starting point at row 1, col 2.  No end range desired
      change_cellref( ( (1,2) , (4,5) )   # starting point at row 1, col 2.  End at row 4 col 5
      '''
      error = ValueError("item must be 2 ints or 2 tupples")
      if type(item0) is int:
         if type(item1) is not int: raise error
         # it is a 1D tuple and has no endrange
         self.start_row = item0
         self.start_col = item1
         self.end_row = None
         self.end_col = None
         self.isMatrix = True
      
      elif type(item0) in (tuple, list):
         if type(item1) not in (tuple, list): raise error
         # it is a 2D tuple, so it contains an end-range
         self.start_row = item0[0]
         self.start_col = item0[1]
         self.end_row = item1[0]
         self.end_col = item1[1]
               
      else: raise error
   
   def change_dtype(self, dtype):
      '''use list or tuple to recieve a python lists or python tuples.  Use int and float to
      receive numpy arrays of integers or floating point numbers'''
      if dtype in (list, tuple, int, float):
         self._dtype_ = dtype
      else:
         raise ValueError('Requested data return type not valid')
   
   def convert_to_dtype(self, data, length = None):
      '''converts from any data with a length to the current dtype.  Should only ever receive arrays
      or iterators of specified length.
      Note: might not be as fast as through the GeneratedData class, but still useful when
      you already have an array of data.'''
      if self._dtype_ in (int, float):
         if length is not None:
            return np.fromiter(data.__iter__(), self._dtype_, length)
         else:
            return np.array(data, dtype = self._dtype_)
         
      if self._dtype_ is tuple:
         return tuple(data)
      if self._dtype_ is list:
         return list(data)
   
   def _determine_rowscols_(self, rows, cols, data):
      '''can be used to reverse the operations done to find rows, cols by DataHandler.
      takes into account whether the data itself is a matrix.
      Data needs to be indexible (not an iterator)'''
      temp_rows = rows
      temp_cols = cols
      
      if hasattr(data, '__iter__'):
         if hasattr(data[0], '__iter__'): # it is a matrix
            temp_rows = range(rows[0], rows[-1] + len(data))
            temp_cols = range(cols[0], cols[-1] + len(data[0]))
         else: # it is an array
            if len(cols) is 1:      # verticle array
               temp_rows = range(rows[0], rows[-1] + len(data))
            if len(temp_rows) is 1:
               temp_cols = range(cols[0], cols[0] + len(data))
      
      return tuple(temp_rows), tuple(temp_cols)
   
class UndoException(Exception):
   pass

class DataHandlerUser(DataHandler):
   ''' -- NOT FINISHED -- But usable with single terminal
   This is an extension module for any class that is built on DataHandler
   to provide undo/redo functionality.'''
   def __init__(self, data):
      self.__unredofilename = '.undoredo' + str(time()) + '.pyworkbooks'
      try:
         os.remove(self.__unredofilename)
      except OSError: pass
      
      self._undoredoBuffer = shelve.open(self.__unredofilename)
      self._undoindex = 0
      self.maxundo = 100
      self.actioncount = 0
      
      DataHandler.__init__(self, data)
   
   def get_current_data(self, items, value, *args):
      rows, cols = self._interpret_slices_(items)
      rows, cols = self._determine_rowscols_(rows, cols, value)
      
      dtype = self._dtype_
      self.change_dtype(tuple)
      stored_data = self.__getitem__((rows, cols), *args)   # its a good thing it accepts tuples!
      if type(stored_data) is GeneratedData:
         stored_data = stored_data[:]
      self.change_dtype(dtype)
      commands = items, args
      return commands, stored_data
   
   def __setitem__(self, items, value, *args):
      # determine if value is a matrx/array.  If so, it determines length.
      # we are going to interpret ahead of time, so we know if it will set 
      # a matrix/array or only what is in rows/cols
      commands, stored_data = self.get_current_data(items, value, *args)
      
      self.__clearredo()
      
      self.__putbuffer(commands, stored_data)
      return DataHandler.__setitem__(self, items, value, *args)
   
   def __putbuffer(self, commands, data):
      '''receives commands in a buffer, args should be packaged in a tuple, and the data
      that needs to be written by those commands if undo is called'''
      if len(self._undoredoBuffer) > self.maxundo:
         oldkey = sorted(self._undoredoBuffer.keys())[0]
         del self._undoredoBuffer[oldkey]
      self._undoredoBuffer[str(self.actioncount)] = (commands, data)
      self.actioncount += 1

   def __clearredo(self):
      allkeys = self._undoredoBuffer.keys()
      allkeys.sort()
      while self._undoindex > 0:
         del self._undoredoBuffer[allkeys.pop()]
         self._undoindex -= 1
      
   def undo(self):
      if len(self._undoredoBuffer) is 0:
         raise UndoException('Buffer Empty')
      newestkey = sorted(self._undoredoBuffer.keys())[-self._undoindex - 1]
      self._undoindex += 1
      undocommands, undodata = self._undoredoBuffer[newestkey]
      
      # now we must take the data that we put down before, so that we can redo!
      items, args = undocommands
      redocommands, redodata = self.get_current_data(items, undodata, *args)
      self._undoredoBuffer[newestkey] = (redocommands, redodata)
      # basically we are just flipping that data in the worksheet with the data in the buffer     
      
      return DataHandler.__setitem__(self, items, undodata, *args)
   
   def redo(self):
      if self._undoindex is 0:
         raise UndoException('No more operations to redo')
      newestkey = sorted(self._undoredoBuffer.keys())[-self._undoindex]
      self._undoindex -= 1
      redocommands, redodata = self._undoredoBuffer[newestkey]
      
      # now we must take the data that we put down before, so that we can undo again
      items, args = redocommands
      undocommands, undodata = self.get_current_data(items, redodata, *args)
      self._undoredoBuffer[newestkey] = (undocommands, undodata)
      # basically we are just flipping that data.undoredo in the worksheet with the data in the buffer     
      
      DataHandler.__setitem__(self, items, redodata, *args)
      
   def __del__(self):
      if '.undoredo' not in self.__unredofilename or '.pyworkbooks' not in self.__unredofilename:
         raise Exception('Critical Error: Internal filename for undo/redo operations tampered with')
      try:
         os.remove(self.__unredofilename)
      except OSError: pass

def dev1():
   mymatrix = [[p * n for n in xrange(0, 100)] for p in xrange(0, 100)]
   handler = DataHandler(mymatrix)
   handler[1, :] = range(-10, 0)

   handler.change_dtype(float)
   handler[0, :] = [123, 456, 789]   # empty addressing for arrays, doesn't work yet
   print handler[0, :3]
   
   smallmatrix = [[p * n for n in xrange(-10, 0)] for p in xrange(-10, 0)]
   handler[0, 0] = smallmatrix
   print handler[0, :5]
   
   handler.change_dtype(list)
   handler[0, 0:20] = ['hello', 'there', 'bob']
   strings = handler[0:3, :3]
   handler.change_dtype(float)
   print strings[0]
   ['hello', 'there', 'bob']
   data = handler[1:10000, 2:60 + 1]

if __name__ == '__main__':
   dev1()
   import doctest
   doctest.testmod()

