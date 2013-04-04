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

def close_current_project(oDesktop):
    """
    Close the active project.
    
    Parameters
    ----------
    oDesktop : pywin32 COMObject
        The HFSS desktop object upon which to operate.
        
    Returns
    -------
    None
    
    """
    oProject = get_active_project(oDesktop)
    projectname = get_project_name(oProject)
    oDesktop.CloseProject(projectname)

def get_projects(oDesktop):
    """
    Get the list of open projects.
    
    Parameters
    ----------
    oDesktop : pywin32 COMObject
        The HFSS desktop object upon which to operate.
        
    Returns
    -------
    oProject : list of pywin32 COMObjects
        The HFSS desktop object upon which to operate.
        
    """
    return oDesktop.GetProjects()

def close_all_projects(oDesktop):
    """
    Close all open projects.
    
    Parameters
    ----------
    oDesktop : pywin32 COMObject
        The HFSS desktop object upon which to operate.
        
    Returns
    -------
    None
    
    """
    objlist = get_projects(oDesktop)
    
    for item in objlist:
        close_project_byhandle(oDesktop, item)

def close_all_projects_except_current(oDesktop):
    """
    Close all open projects except the active project.
    
    Parameters
    ----------
    oDesktop : pywin32 COMObject
        The HFSS desktop object upon which to operate.
        
    Returns
    -------
    None
    
    """
    currproj = get_project_name(get_active_project(oDesktop))
    projlist = get_projects(oDesktop)
    
    for item in projlist:
        if get_project_name(item) != currproj:
            close_project_byhandle(oDesktop, item)

def insert_frequency_sweep(oDesign,
                           setupname,
                           sweepname,
                           startvalue,
                           stopvalue,
                           stepsize,
                           IsEnabled=True,
                           SetupType="LinearStep",
                           Type="Discrete",
                           SaveFields=True,
                           ExtrapToDC=False):
    """
    Insert an HFSS frequency sweep.
    
    Warning
    -------
    The API interface for this function is very susceptible to change!  It 
    currently only works for Discrete sweeps using Linear Steps.  Contributions 
    are encouraged.
    
    Parameters
    ----------
    oAnalysisSetup : pywin32 COMObject
        The HFSS Analysis Setup Module in which to insert the sweep.
    setupname : string
        The name of the setup to add
    sweepname : string
        The desired name of the sweep
    startvalue : float
        Lowest frequency in Hz.
    stopvalue : float
        Highest frequency in Hz.
    stepsize : flot
        The frequency increment in Hz.
    IsEnabled : bool
        Whether the sweep is enabled.
    SetupType : string
        The type of sweep setup to add.  One of "LinearStep", "LinearCount", 
        or "SinglePoints".  Currently only "LinearStep" is supported.
    Type : string
        The type of sweep to perform.  One of "Discrete", "Fast", or 
        "Interpolating".  Currently only "Discrete" is supported.
    Savefields : bool
        Whether to save the fields.
    ExtrapToDC : bool
        Whether extrapolation to DC is enabled.
        
    Returns
    -------
    None
    
    """
    oAnalysisSetup = oDesign.GetModule("AnalysisSetup")
    return oAnalysisSetup.InsertFrequencySweep(setupname, 
                                        ["NAME:" + sweepname, 
                                         "IsEnabled:=", IsEnabled, 
                                         "SetupType:=", SetupType, 
                                         "StartValue:=", str(startvalue) + "Hz", 
                                         "StopValue:=", str(stopvalue) + "Hz", 
                                         "StepSize:=", str(stepsize) + "Hz", 
                                         "Type:=", Type, 
                                         "SaveFields:=", SaveFields, 
                                         "ExtrapToDC:=", ExtrapToDC])

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
