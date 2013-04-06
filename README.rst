hycohanz
========

hycohanz is an Open Source (BSD license) Python wrapper interface 
to the ANSYS HFSS Windows COM API, enabling you to control HFSS 
from Python.

Warning
-------

hycohanz is pre-alpha software.  The interfaces can be expected to 
break constantly.  If you are looking for a stable, reliable 
function library for HFSS, you must look elsewhere for now.

Documentation
-------------

See http://mradway.github.io/hycohanz/ for installation instructions.  
This page also has a minimal example that can be used to test hycohanz 
once installation is complete.  Several more basic examples can be 
found in the examples directory.

Most wrapper functions are documented with useful docstrings, and in most 
cases the their interfaces tend to follow the HFSS API fairly closely.

For best use of this library you really should familiarize yourself with the 
information in the HFSS Scripting Guide, available in the HFSS GUI under 
Help->Scripting Contents.  The library is intended to be used in consultation 
with this resource.

If the docstrings and examples are not sufficient, you will find that 
many functions consist of five or fewer lines of simple (almost trivial) 
code that are easily understood.

Contributing
------------

Often one finds that this library is missing a wrapper for a particular 
function.  Fortunately it's often quite easy to add, usually taking 
only a few minutes.  Most of the time it's a quick modification of 
an existing function.  Many functions can be implemented in five 
lines of code or less.  If you do add a feature to the code, please 
consider contributing it back to this project.