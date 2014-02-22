Notes
=====

For Users
---------

This directory defines a Python package_.  Each file (or module in Python parlance) in this directory more or less corresponds to a section in the HFSS Scripting guide.

.. _package: http://docs.python.org/2/tutorial/modules.html#packages

You can import this package in a few different ways.  See here_ for an old but still pertinent discussion of module importing styles in Python.  You should generally follow their advice.

.. _here: http://effbot.org/zone/import-confusion.htm

1. Just as in the examples
   
   .. sourcecode:: python

        import hycohanz as hfss

   This basically imports all of hycohanz into a single flat namespace under the alias hfss.  
   This is useful for conciseness, but as hycohanz has grown to implement a large number of
   HFSS scripting functions, this flat namespace may seem cumbersome to some users.  For this
   reason hycohanz has been broken down into submodules as mentioned above.  This affords us a
   second way to import hycohanz as discussed next.

2. As submodules

   .. sourcecode:: python

       import hycohanz.modeler3d as hfssm3d
       import hycohanz.fieldscalculator as hfssfc

   ... etc.  The only problem here is that it's perhaps overly verbose for most purposes.

3. Cherry-picking

   .. sourcecode:: python

       from hycohanz.modeler3d import create_circle, create_polyline

4. Kitchen-sink style, which I recommend against

   .. sourcecode:: python

       from hycohanz import *

   For those coming over from MATLAB this will seem the most natural since MATLAB encourages 
   users to have lots of functions in the global namespace.  The main reasons I don't like 
   this one is that it 1. pollutes the global namespace, which
   will cause you headaches as your scripts become more complex, and 2. makes    
   your code harder for static analysis tools to understand, and static analysis is very   
   useful.  

As of this writing, #1 is my current preferred importing style for day-to-day use, but as 
stated earlier this may change over time.

For Developers
--------------

The way the code is currently structured, if you're adding functions to the hyohanz submodules, you should remember to import them in hycohanz.py so that they show up in the flat namespace of import style #1 above.  Otherwise, you'll force your users to use import styles #2 or #3.

 