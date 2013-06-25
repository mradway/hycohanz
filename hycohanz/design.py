# -*- coding: utf-8 -*-
"""
Created on Mon Jun 24 20:32:40 2013

@author: radway
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

