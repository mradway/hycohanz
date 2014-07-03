# -*- coding: utf-8 -*-
"""
Functions in this module correspond more or less to the functions described 
in the HFSS Scripting Guide, Section "Design Object Script Commands".

At last count there were 2 functions implemented out of 27.
"""
from __future__ import division, print_function, unicode_literals, absolute_import

def get_module(oDesign, ModuleName):
    """
    Get a module handle for the given module.
    
    Parameters
    ----------
    oDesign : pywin32 COMObject
        The HFSS design object upon which to operate.
    ModuleName : str
        Name of the module to return.  One of 
            - "BoundarySetup"
            - "MeshSetup"
            - "AnalysisSetup"
            - "Optimetrics"
            - "Solutions"
            - "FieldsReporter"
            - "RadField"
            - "UserDefinedSolutionModule"
        
    Returns
    -------
    oModule : pywin32 COMObject
        Handle to the given module
        
    """
    oModule = oDesign.GetModule(ModuleName)
    
    return oModule

def set_active_editor(oDesign, editorname="3D Modeler"):
    """
    Set the active editor.
    
    Parameters
    ----------
    oDesign : pywin32 COMObject
        The HFSS design upon which to operate.
    editorname : str
        Name of the editor to set as active.  As of this writing "3D Modeler" 
        is the only known valid value.
    
    Returns
    -------
    oEditor : pywin32 COMObject
        The HFSS Editor object.
        
    """
    oEditor = oDesign.SetActiveEditor(editorname)
    
    return oEditor

def solve(oDesign,setup_name_list):
    """
    Solve Setup.

    Parameters
    ----------
    oDesktop : pywin32 COMObject
        HFSS Desktop object.
    
    Returns
    -------
    None
    
    Examples
    --------
    >>> import Hyphasis as hfss
    >>> 
    
    """
    return oDesign.Solve([setup_name_list])
