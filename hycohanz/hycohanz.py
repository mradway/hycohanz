# -*- coding: utf-8 -*-
"""
Expose the HFSS Windows COM API.

Example Usage
-------------
>>> import hycohanz as hfss
>>> [oAnsoftApp, oDesktop] = hfss.setup_interface()
>>> hfss.quit_application(oDesktop)

"""

from __future__ import division, print_function, unicode_literals, absolute_import

import warnings

import win32com.client

warnings.simplefilter('default')

def setup_interface():
    """
    Set up the COM interface to the running HFSS process.
    
    Returns
    -------
    oAnsoftApp : pywin32 COMObject
        Handle to the HFSS application interface
    oDesktop : pywin32 COMObject
        Handle to the HFSS desktop interface
    
    Example Usage
    -------------
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
    
def quit_application(oDesktop):
    """
    Exit HFSS.

    Parameters
    ----------
    oDesktop : pywin32 COMObject
        HFSS Desktop object.
    
    Returns
    -------
    None
    
    Example Usage
    -------------
    >>> import Hyphasis as hfss
    >>> [oAnsoftApp, oDesktop] = hfss.setup_interface()
    >>> hfss.quit_application(oDesktop)
    
    """
    oDesktop.QuitApplication()

def new_project(oDesktop):
    """
    Create a new project.  The new project becomes the active project.
    
    Parameters
    ----------
    oDesktop : pywin32 COMObject
        HFSS Desktop object in which to create the new project.
        
    Returns
    -------
    oProject : pywin32 COMObject
        The created HFSS project.
        
    """
    oProject = oDesktop.NewProject()
    
    return oProject

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

def insert_design(oProject, designname, solutiontype):
    """
    Insert an HFSS design.  The scripting interface doesn't appear to support 
    creation of HFSS-IE designs at this time, or is undocumented.
    
    Parameters
    ----------
    oProject : pywin32 COMObject
        The HFSS project in which the operation will be performed.
    designname : str
        Name of the design to insert.
    solutiontype : str
        Name of the solution type.  One of ("DrivenModal", 
                                            "DrivenTerminal", 
                                            "Eigenmode")
        
    Returns
    -------
    oDesign : pywin32 COMObject
        The created HFSS design.
        
    """
    oDesign = oProject.InsertDesign("HFSS", designname, solutiontype, "")
    
    return oDesign

def get_project_name(oProject):
    """
    Get the name of the specified project.
    
    Parameters
    ----------
    oProject : pywin32 COMObject
        The HFSS project object upon which to operate.
        
    Returns
    -------
    str
        The name of the project.
        
    """
    return oProject.GetName()

def close_project_byname(oDesktop, projectname):
    """
    Close the specified project using the given HFSS project object.
    
    Parameters
    ----------
    oDesktop : pywin32 COMObject
        The HFSS desktop object upon which to operate.
    projectname : string
        The name of the HFSS project to close.
        
    Returns
    -------
    None
    
    """
    oDesktop.CloseProject(projectname)

def close_project_byhandle(oDesktop, oProject):
    """
    Close the specified project using the given HFSS project object.
    
    Parameters
    ----------
    oDesktop : pywin32 COMObject
        The HFSS desktop object upon which to operate.
    oProject : pywin32 COMObject
        The HFSS project object upon which to operate.
        
    Returns
    -------
    None
    
    """
    oDesktop.CloseProject(get_project_name(oProject))

def get_active_project(oDesktop):
    """
    Get a handle for the active project.
    
    Parameters
    ----------
    oDesktop : pywin32 COMObject
        The HFSS desktop object upon which to operate.
        
    Returns
    -------
    oProject : pywin32 COMObject
        The HFSS project object upon which to operate.
    
    """
    return oDesktop.GetActiveProject()

