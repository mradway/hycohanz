# -*- coding: utf-8 -*-
"""
Created on Mon Jun 24 19:43:08 2013

@author: radway
"""
from __future__ import division, print_function, unicode_literals, absolute_import

import win32com.client

def setup_interface():
    """
    Set up the COM interface to the running HFSS process.
    
    Returns
    -------
    oAnsoftApp : pywin32 COMObject
        Handle to the HFSS application interface
    oDesktop : pywin32 COMObject
        Handle to the HFSS desktop interface
    
    Examples
    --------
    >>> import Hyphasis as hfss
    >>> [oAnsoftApp, oDesktop] = hfss.setup_interface()
    
    """
    # I'm still looking for a better way to do this.  This attaches to an 
    # existing HFSS process instead of creating a new one.  I would highly 
    # prefer that a new process is created.  Apparently 
    # win32com.client.DispatchEx() doesn't work here either.
    oAnsoftApp = win32com.client.Dispatch('AnsoftHfss.HfssScriptInterface')

    oDesktop = oAnsoftApp.GetAppDesktop()

    return [oAnsoftApp, oDesktop]