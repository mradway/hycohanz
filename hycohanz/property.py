# -*- coding: utf-8 -*-
"""
HFSS functions that use the Property module. Functions in this module correspond 
more or less to the functions described in the HFSS Scripting Guide, 
Section "Property Script Commands".

At last count there were 1 functions implemented out of 7.
"""
from __future__ import division, print_function, unicode_literals, absolute_import

from hycohanz.expression import Expression

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

def set_variable(oProject, name, value):
    """
    Change a design property.  This function differs significantly from 
    SetVariableValue() in that it makes the reasonable assumption that 
    if the variable contains '$', then the variable is global; otherwise, 
    it is assumed to be a local variable.
    
    Parameters
    ----------
    oProject : pywin32 COMObject
        The HFSS design from which to retrieve the module.
    name : str
        The name of the property/variable to edit.
    value : Hyphasis Expression object
        The new value of the property.
        
    Returns
    -------
    None
    
    """
    if '$' in name: 
		oProject.SetVariableValue(name,Expression(value).expr)
    else:
		oDesign = oProject.GetActiveDesign()
		oDesign.SetVariableValue(name,Expression(value).expr)

def get_variables(oProject,oDesign=''):
    """
    get list of non-indexed variables.
    
    Parameters
    ----------
    oProject : pywin32 COMObject
        The HFSS design from which to retrieve the variables.
    oDesign : pywin32 COMObject
        Optional, if specified function returns variable list of oDesign.

        
    Returns
    -------
    variable_list: list of str
	list of non-indexed project/design variables
    
    """
    if oDesign=='':
        variable_list = list(oProject.GetVariables())
    else:
        variable_list = list(oDesign.GetVariables())
    return map(str,variable_list)


