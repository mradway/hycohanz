# -*- coding: utf-8 -*-
"""
Functions in this module correspond more or less to the functions described 
in the HFSS Scripting Guide, Section "Boundary and Excitation Module Script 
Commands".

At last count there were 4 functions implemented out of 20.
"""
from __future__ import division, print_function, unicode_literals, absolute_import

from hycohanz.design import get_module

def assign_perfect_e(oDesign, boundaryname, facelist, InfGroundPlane=False):
    """
    Create a perfect E boundary.
    
    Parameters
    ----------
    oDesign : pywin32 COMObject
        The HFSS design to which this function is applied.
    boundaryname : str
        The name to give this boundary in the Boundaries tree.
    facelist : list of ints
        The faces to assign to this boundary condition.
    
    Returns
    -------
    None
    """
    oBoundarySetupModule = get_module(oDesign, "BoundarySetup")
    oBoundarySetupModule.AssignPerfectE(["Name:" + boundaryname, "Faces:=", facelist, "InfGroundPlane:=", InfGroundPlane])

def assign_radiation(oDesign, 
                     faceidlist, 
                     IsIncidentField=False, 
                     IsEnforcedField=False, 
                     IsFssReference=False, 
                     IsForPML=False,
                     UseAdaptiveIE=False,
                     IncludeInPostproc=True,
                     Name='Rad1'):
    """
    Assign a radiation boundary on the given faces.
    
    Parameters
    ----------
    oDesign : pywin32 COMObject
        The HFSS design to which this function is applied.
    faceidlist : list of ints
        The faces to assign to this boundary condition.
    IsIncidentField : bool
        If True, same as checking the "Incident Field" radio button in the 
        Radiation Boundary setup dialog.  Mutually-exclusive with 
        IsEnforcedField
    IsEnforcedField : bool
        If True, same as checking the "Enforced Field" radio button in the 
        Radiation Boundary setup dialog.  Mutually-exclusive with 
        IsIncidentField
    IsFssReference : bool
        If IsEnforcedField is False, is equivalent to checking the 
        "Reference for FSS" check box i the Radiation Boundary setup dialog.
    IsForPML : bool
        Not explored at this time.  Likely use case is when defining a 
        radiation boundary in conjuction with PMLs where the boundary lies on 
        the surface between the PML and the PML base object.
    UseAdaptiveIE : bool
        Not explored at this time.  It is likely that setting this to True is 
        equivalent to selecting the "Model exterior as HFSS-IE domain" check 
        box in the Radiation Boundary setup dialog.
    IncludeInPostproc : bool
        Not explored at this time.  Likely use case is to remove certain 
        boundaries from consideration during certain postprocessing 
        operations, such as when computing the radiation pattern.
    Name : str
        The name to assign the boundary.
    Returns
    -------
    None
    """
    oBoundarySetupModule = get_module(oDesign, "BoundarySetup")
    arg = ["NAME:{0}".format(Name), 
           "Faces:=", faceidlist, 
           "IsIncidentField:=", IsIncidentField, 
           "IsEnforcedField:=", IsEnforcedField, 
           "IsFssReference:=", IsFssReference, 
           "IsForPML:=", IsForPML, 
           "UseAdaptiveIE:=", UseAdaptiveIE, 
           "IncludeInPostproc:=", IncludeInPostproc]
    
    oBoundarySetupModule.AssignRadiation(arg)

def assign_perfect_h(oDesign, boundaryname, facelist):
    """
    Create a perfect H boundary.
    
    Parameters
    ----------
    oDesign : pywin32 COMObject
        The HFSS design to which this function is applied.
    boundaryname : str
        The name to give this boundary in the Boundaries tree.
    facelist : list of ints
        The faces to assign to this boundary condition.
    
    Returns
    -------
    None
    """
    oBoundarySetupModule = get_module(oDesign, "BoundarySetup")
    oBoundarySetupModule.AssignPerfectH(["Name:" + boundaryname, "Faces:=", facelist])

def assign_waveport_multimode(oDesign, 
                              portname, 
                              faceidlist, 
                              Nmodes=1,
                              RenormalizeAllTerminals=True,
                              UseLineAlignment=False,
                              DoDeembed=False,
                              ShowReporterFilter=False,
                              ReporterFilter=[True],
                              UseAnalyticAlignment=False):
    """
    Assign a waveport excitation using multiple modes.
    
    Parameters
    ----------
    oDesign : pywin32 COMObject
        The HFSS design to which this function is applied.
    portname : str
        Name of the port to create.
    faceidlist : list
        List of face id integers.
    Nmodes : int
        Number of modes with which to excite the port.
        
    Returns
    -------
    None
    """
    oBoundarySetupModule = get_module(oDesign, "BoundarySetup")
    modesarray = ["NAME:Modes"]
    for n in range(0, Nmodes):
        modesarray.append(["NAME:Mode" + str(n + 1),
                           "ModeNum:=", n + 1,
                           "UseIntLine:=", False])

    waveportarray = ["NAME:" + portname, 
                     "Faces:=", faceidlist, 
                     "NumModes:=", Nmodes, 
                     "RenormalizeAllTerminals:=", RenormalizeAllTerminals, 
                     "UseLineAlignment:=", UseLineAlignment, 
                     "DoDeembed:=", DoDeembed, 
                     modesarray, 
                     "ShowReporterFilter:=", ShowReporterFilter, 
                     "ReporterFilter:=", ReporterFilter, 
                     "UseAnalyticAlignment:=", UseAnalyticAlignment]

    oBoundarySetupModule.AssignWavePort(waveportarray)





        
