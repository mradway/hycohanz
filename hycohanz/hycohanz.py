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

from hycohanz.expression import Expression
from hycohanz.modeler3d import *
from hycohanz.fieldscalculator import *

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
    
    Examples
    --------
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

def insert_analysis_setup(oDesign, 
                          Frequency,
                          PortsOnly=True,
                          MaxDeltaS=0.02,                       
                          Name='Setup1',
                          UseMatrixConv=False,
                          MaximumPasses=20,
                          MinimumPasses=2,
                          MinimumConvergedPasses=2,
                          PercentRefinement=30,
                          IsEnabled=True,
                          BasisOrder=2,
                          UseIterativeSolver=False,
                          DoLambdaRefine=True,
                          DoMaterialLambda=True,
                          SetLambdaTarget=True,
                          Target=0.6667,
                          UseMaxTetIncrease=False,
                          PortAccuracy=2,
                          UseABCOnPort=False,
                          SetPortMinMaxTri=False,
                          EnableSolverDomains=False,
                          ThermalFeedback=False,
                          NoAdditionalRefinementOnImport=False):
    """
    Insert an HFSS analysis setup.
    """
    oAnalysisSetup = get_module(oDesign, "AnalysisSetup")
    oAnalysisSetup.InsertSetup( "HfssDriven", 
                               ["NAME:" + Name, 
                                "Frequency:=", str(Frequency) +"Hz", 
                                "PortsOnly:=", PortsOnly, 
                                "MaxDeltaS:=", MaxDeltaS, 
                                "UseMatrixConv:=", UseMatrixConv, 
                                "MaximumPasses:=", MaximumPasses, 
                                "MinimumPasses:=", MinimumPasses, 
                                "MinimumConvergedPasses:=", MinimumConvergedPasses, 
                                "PercentRefinement:=", PercentRefinement, 
                                "IsEnabled:=", IsEnabled, 
                                "BasisOrder:=", BasisOrder, 
                                "UseIterativeSolver:=", UseIterativeSolver, 
                                "DoLambdaRefine:=", DoLambdaRefine, 
                                "DoMaterialLambda:=", DoMaterialLambda, 
                                "SetLambdaTarget:=", SetLambdaTarget, 
                                "Target:=", Target, 
                                "UseMaxTetIncrease:=", UseMaxTetIncrease, 
                                "PortAccuracy:=", PortAccuracy, 
                                "UseABCOnPort:=", UseABCOnPort, 
                                "SetPortMinMaxTri:=", SetPortMinMaxTri, 
                                "EnableSolverDomains:=", EnableSolverDomains, 
                                "ThermalFeedback:=", ThermalFeedback, 
                                "NoAdditionalRefinementOnImport:=", NoAdditionalRefinementOnImport])
                                
    return Name

def add_property(oDesign, name, value):
    """
    Add a design property.
    
    Parameters
    ----------
    oDesign : pywin32 COMObject
        The HFSS design from which to retrieve the module.
    name : str
        The name of the property to add.
    value : Hyphasis Expression object
        The value of the property.
        
    Returns
    -------
    None
    
    """
    propserversarray = ["NAME:PropServers", "LocalVariables"]

    newpropsarray = ["NAME:NewProps", ["NAME:" + name, 
                                       "PropType:=", "VariableProp", 
                                       "UserDef:=", True, 
                                       "Value:=", Expression(value).expr]]
    
    proptabarray = ["NAME:LocalVariableTab", propserversarray, newpropsarray]
          
    oDesign.ChangeProperty(["NAME:AllTabs", proptabarray])

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

def open_project(oDesktop, filename):
    """
    Open an HFSS project.
    
    Parameters
    ----------
    oDesktop : pywin32 COMObject
        The HFSS desktop object to which this function is applied.
    filename : str
        The name of the file to open.
    
    Returns
    -------
    oProject : pywin32 COMObject
        An handle to the opened project.
    
    """
    return oDesktop.OpenProject(filename)
    