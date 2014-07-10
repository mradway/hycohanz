# -*- coding: utf-8 -*-
"""
HFSS functions that use the Project module. Functions in this module correspond 
more or less to the functions described in the HFSS Scripting Guide, 
Section "Project Object Script Commands".

At last count there were 3 functions implemented out of 17.
"""
from __future__ import division, print_function, unicode_literals, absolute_import

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

def set_active_design(oProject, designname):
    """
    Set the active design.
    
    Parameters
    ----------
    oProject : pywin32 COMObject
        The HFSS design upon which to operate.
    designname : str
        Name of the design to set as active.
    
    Returns
    -------
    oDesign : pywin32 COMObject
        The HFSS Design object.
        
    """
    oEditor = oProject.SetActiveDesign(designname)
    
    return oEditor
    
def insert_design(oProject, designname, solutiontype):
    """
    Insert an HFSS design.  The inserted design becomes the active design.
    
    Note:  The scripting interface doesn't appear to support 
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

def get_design(oProject, design_name):
    """
    Returns specified design.
    
    Parameters
    ----------
    oProject : pywin32 COMObject
        The HFSS project in which the operation will be performed.
    designname : str
        Name of the design to insert.
        
    Returns
    -------
    oDesign : pywin32 COMObject
        The specified HFSS design.
        
    """
    oDesign = oProject.GetDesign(design_name)
    return oDesign

def get_top_design_list(oProject):
    """
    Returns a list of the names of the top-level designs.
    
    Parameters
    ----------
    oProject : pywin32 COMObject
        The HFSS project in which the operation will be performed.
    designname : str
        Name of the design to insert.
        
    Returns
    -------
    design_list : list of str
        The top-level design list.
        
    """
    design_list = list(oProject.GetTopDesignList())
    return map(str,design_list)

