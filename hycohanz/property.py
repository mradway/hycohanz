# -*- coding: utf-8 -*-
"""
Created on Mon Jun 24 20:37:00 2013

@author: radway
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


