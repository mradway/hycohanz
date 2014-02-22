Examples
========

In this directory are a number of very simple example scripts.  Basically, each script is intended to highlight basic usage of a single hycohanz_ function given by the name of the script itself.  This is intended to make it easy for one to find an example script for an unfamiliar function.  

.. _hycohanz:  http://mradway.github.io/hycohanz/

If you find a script that doesn't work and you believe it's a bug, please create an issue in the `issue tracker`_.  Likewise, if you find a hycohanz function that doesn't have an example here, the issue tracker is a great place to report it.  Of course, if you just want to email_ it to me directly, that's fine too.  And of course, if you've written an example yourself we'd like to see it!

.. _`issue tracker`: https://github.com/mradway/hycohanz/issues
.. _email:  mailto:mradway@gmail.com

Important Note on Windows COM Object Bookkeeping
------------------------------------------------

To keep the examples as simple as possible, almost all of these scripts interact with HFSS in the following manner:

.. sourcecode:: python

    import hycohanz as hfss

    [oAnsoftApp, oDesktop] = hfss.setup_interface()

    raw_input('Press "Enter" to quit HFSS.>')

    hfss.quit_application(oDesktop)

    del oDesktop
    del oAnsoftApp

Note that the last two lines are very important; they delete the Windows COM objects that have been created to enable script interaction between HFSS and Python.  For some (as yet unknown) reason, Windows fails to deallocate these objects properly upon script completion unless the objects are explicitly deallocated within the script.  This does not present a problem unless the script abnormally exits (crashes) in the middle of execution, prior to the deallocation.

Unfortunately, your only indication that this did not happen properly is that HFSS begins to exhibit strange behavior.  For instance, if you attempt to shut down HFSS after a running script crashes it may appear to do so properly; however, if you examine the process list in the Windows Task Manager, you may find hfss.exe still running.  We'll refer to that process as a zombie.  Furthermore, if you restart HFSS, you may find that it appears to start up fine, but when you try to run scripts again you may find that they no longer work.  Upon examining the process list again you may find that there are in fact *two* HFSS processes running, the zombie and the new process.  To get your scripts working again is a simple matter of deleting the zombie process from the Windows Task Manager process listing.

Of course, you can attempt to deal with this issue by wrapping each hycohanz call in a Python try...except...finally block, but that gets a cumbersome quickly.  Fortunately, Python has the built-in concept of a `context manager`_ using the "with" statement (note: this statement has no obvious similarity to the Visual Basic "with" statement).  hycohanz has implemented a basic complement of context managers that can make your HFSS scripting life much easier.  

.. _`context manager`: http://legacy.python.org/dev/peps/pep-0343/

The new usage pattern ends up looking like the following (adapted from add_property.py):

.. sourcecode:: python

    import hycohanz as hfss
    
    with hfss.App() as App:       
        with hfss.NewProject(App.oDesktop) as P:
            with hfss.InsertDesign(P.oProject, "HFSSDesign1", "DrivenModal") as D:
                hfss.add_property(D.oDesign, "length", hfss.Expression("1m"))
            
                raw_input('Press "Enter" to quit HFSS.>')

Note that all of the explicit setup and teardown code in the old version is missing in the new version; yet, all allocation is handled properly (that is, assuming the hycohanz context managers are implemented correctly).  The setup and teardown logic has been "factored out" (to use a software engineering term) to the hycohanz library, making your scripts cleaner and clearer.
