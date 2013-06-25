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


