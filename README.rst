hycohanz
========

hycohanz_ is an Open Source (BSD license) Python wrapper interface to the ANSYS HFSS Windows COM API, 
enabling you to control HFSS from Python.  
hycohanz simplifies control of HFSS from Python for RF, microwave, and antenna engineers.

.. _hycohanz:  http://mradway.github.io/hycohanz/

Minimal Example
---------------

.. sourcecode:: python

    import hycohanz as hfss

    [oAnsoftApp, oDesktop] = hfss.setup_interface()

    raw_input('Press "Enter" to quit HFSS.>')

    hfss.quit_application(oDesktop)

    del oDesktop
    del oAnsoftApp

Dozens more examples_ are included in the examples directory of the source distribution.

.. _examples:  https://github.com/mradway/hycohanz/tree/devel/examples


Quick Install
-------------

Installation is easy if you already have HFSS, Python, and the pywin32 Python package:

1. Download the `.zip file`_ from Github.

.. _`.zip file`:  https://github.com/mradway/hycohanz/archive/devel.zip

2. Unzip to a convenient location.

3. At the Windows command shell prompt, run::

    > C:\Python27\python setup.py install

See `Detailed Installation`_ if you don't already have Python installed.

.. _`Detailed Installation`:  http://mradway.github.io/hycohanz/install.html

Problems, Bugs, Questions, and Feature Requests
-----------------------------------------------
These are currently handled via the hycohanz issue tracker https://github.com/mradway/hycohanz/issues.  

This issue tracker is useful if you 

- run into problems with installing or running hycohanz, 
- find a bug, 
- have a question,
- would like to see a feature implemented.

Features
--------
hycohanz provides convenience functions for the following:

- Starting, connecting to, and closing HFSS
- Creating design variables
- Manipulating HFSS expressions
- Creating 3D models using polylines, circles, rectangles, spheres, etc.
- Querying objects and groups of objects
- Object manipulation via unite, subtract, imprint, mirror, move, cut, paste, rotate, scale, sweep, etc.
- Assigning boundary conditions
- Manipulating projects and designs
- Creating analysis setups and frequency sweeps

Examples
--------
Dozens of examples_ are included in the examples directory of the source distribution.

.. _examples:  https://github.com/mradway/hycohanz/tree/devel/examples

Warning
-------

hycohanz is pre-alpha software and is in active development.  
The hycohanz function interfaces can be expected to change frequently, with little concern for backwards compatibility.
This situation is expected to resolve as the project approaches a more mature state.  
However, if today you require a stable, reliable, and correct function library for HFSS, unfortunately this library is probably not for you in its current form.

See Also
--------
scikit-rf_:  An actively-developed library for performing common tasks in RF, providing functionality analogous to that provided by the MATLAB RF Toolbox.  If you're working with RF or microwave you should consider getting it.

PyVISA_:  Enables control of instrumentation via Python.

matplotlib_:  Excellent Python 2-D plotting library.

numpy_:  Fundamental functions for manipulating arrays and matrices and performing linear algebra in Python.  

scipy_:  Builds upon numpy_ to enable MATLAB-like functionality in Python.

sympy_:  Implements analogous functionality to the MATLAB Symbolic Toolbox.

.. _scikit-rf:  http://scikit-rf.org/
.. _PyVISA:  http://pyvisa.sourceforge.net/
.. _matplotlib:  http://matplotlib.org/
.. _numpy:  http://www.numpy.org/
.. _scipy:  http://www.scipy.org/
.. _sympy:  http://sympy.org/en/index.html

Download
--------

A zip file of the development branch can be downloaded from 
https://github.com/mradway/hycohanz/archive/devel.zip

Of course, one can also pull the source tree in the usual way using git.

Installation
------------
See http://mradway.github.io/hycohanz/ for detailed installation instructions.  
This page also has a minimal example that can be used to test hycohanz 
once installation is complete.  

Documentation
-------------

Several basic examples can be found in the examples directory.

Most wrapper functions are documented with useful docstrings, and in most 
cases their interfaces tend to follow the HFSS API fairly closely.

For best use of this library you should familiarize yourself with the 
information in the HFSS Scripting Guide, available in the HFSS GUI under 
Help->Scripting Contents.  The library is intended to be used in consultation 
with this resource.

If the docstrings and examples are not sufficient, you will find that 
many functions consist of five or fewer lines of simple (almost trivial) 
code that are easily understood.

Frequently Asked Questions
--------------------------

:Q: Why not write scripts using Visual Basic for Applications (VBA) or JavaScript (JS)?
:A: I've found that programming in Python is generally much, much easier and more 
    powerful than in either of these languages.  Plus, I've generally found that 
    Visual Basic scripts run inside HFSS tend to break without useful error 
    messages, or worse, crash HFSS entirely.  hycohanz can also crash HFSS. But 
    when it does, the Python interpreter gives you a nice stack trace, allowing 
    you to determine what went wrong.

:Q: Why use Windows COM instead of .NET?
:A: As I understand it, the Visual Basic examples in the HFSS Scripting Guide 
    use Windows COM, so that's what I use.  If you're using IronPython, then 
    accessing .NET resources should be trivial.  However, I don't use IronPython 
    since I make extensive use in my daily work of numpy, scipy, matplotlib, 
    h5py, etc., and IronPython has had issues integrating with these tools 
    in the past.

:Q: Why not metaprogram VBA or JS?  Then I could use this library on Linux.
:A: That was my initial approach, because I wanted cross-platform capability.  
    Compared to the Windows COM approach, it's a lot more time-consuming, and 
    it has all of the drawbacks of the first question.

:Q: Why did you use Python instead of MATLAB?
:A: I'm a recent convert to Python, so I now use Python in my daily workflow 
    whenever it's convenient (that means about 99.9% of the time). Python 
    gives you keyword arguments, which helps keep the average length in characters 
    of a hycohanz function call to a minimum, while minimizing implementation 
    overhead compared to MATLAB.

:Q: Why not skip the HFSS interface entirely and directly emit a .hfss file?  Then 
    I could use this library on Linux.
:A: I've also considered this approach.  As you may know, .hfss files are 
    quasi-human-readable text files with a file format that could in principle be 
    reasonably parsed and emitted.  However, the expected implementation effort 
    would have been quite a bit higher than I wanted.  Not to mention that the format is not 
    (to my knowledge) static, nor is it publicly specified or documented.  Thus, an 
    implementation of this approach would be expected to be fragile, crash HFSS 
    frequently, and leave non-useful error messages.

Contributing
------------

Often one finds that this library is missing a wrapper for a particular 
function.  Fortunately it's often quite easy to add, usually taking 
only a few minutes.  Most of the time it's a quick modification of 
an existing function.  Many functions can be implemented in five 
lines of code or less.  If you do add a feature to the code, please 
consider contributing it back to this project.
