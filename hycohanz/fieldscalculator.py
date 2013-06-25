# -*- coding: utf-8 -*-
"""
HFSS Fields Calculator functions. Functions in this module correspond more or 
less to the functions described in the HFSS Scripting Guide, Section "Fields 
Calculator Script Commands".

At last count there were 5 functions implemented out of 28.
"""

from __future__ import division, print_function, unicode_literals, absolute_import

def enter_vol(oFieldsReporter, VolumeName):
    """
    Enters a volume defined in the 3D Modeler editor into the Fields Calculator.

    Parameters
    ----------
    oFieldsReporter : pywin32 COMObject
        An HFSS "FieldsReporter" module
    VolumeName : str
        Name of a volume defined in the 3D Modeler editor.

    Returns
    -------
    None
    """
    oFieldsReporter.EnterVol(VolumeName)



def calc_op(oFieldsReporter, OperationString):
    """
    Performs a calculator operation in the Fields Calculator.
    
    Parameters
    ----------
    oFieldsReporter : pywin32 COMObject
        An HFSS "FieldsReporter" module 
    OperationString : str
        The text on the corresponding calculator button.
        
    Returns
    -------
    None
    """
    oFieldsReporter.CalcOp(OperationString)
    


def clc_eval(oFieldsReporter, setupname, sweepname, freq, phase, variablesdict):
    """
    Evaluates the expression at the top of the Fields Calculator stack using 
    the provided solution name and variable values.

    Parameters
    ----------
    oFieldsReporter : pywin32 COMObject
        An HFSS "FieldsReporter" module 
    setupname : str
        Name of HFSS setup to use, for example "Setup1"
    sweepname : str
        Name of HFSS sweep to use, for example "LastAdaptive"
    freq : float
        Frequency at which to evaluate the calculator expression.
    phase : float
        Phase in degrees at which to evaluate the calculator expression.
    variablesdict : dict
        Dictionary listing the variables and their values that define the 
        design variation, except for 'Freq' and 'Phase'.  Variable names are the keys (strings), and the values are 
        the values of the variable (floats). For example: {'radius': 0.5, 'height': '2.0'}
        
    Returns
    -------
    None
    """
    solutionname = setupname + " : " + sweepname
    
    variablesarray = ["Freq:=", str(freq) + 'Hz', "Phase:=", str(phase) + 'deg']
    
    for key in variablesdict:
        variablesarray += [str(key) + ':=', str(variablesdict[key])]
        
    print('solutionname: ' + str(solutionname))
    print('variablesarray: ' + str(variablesarray))
    
    oFieldsReporter.ClcEval(solutionname, variablesarray)
    
    
    
def enter_qty(oFieldsReporter, FieldQuantityString):
    """
    Enter a field quantity into the Fields Calculator.
    
    Parameters
    ----------
    oFieldsReporter : pywin32 COMObject
        An HFSS "FieldsReporter" module 
    FieldQuantityString : str
        The field quantity to be entered onto the stack.
        
    Returns
    -------
    None
    """
    return oFieldsReporter.EnterQty(FieldQuantityString)
    
    
    
def get_top_entry_value(oModule, setupname, sweepname, freq, phase, variablesdict):
    """
    Evaluates the expression at the top of the stack using the provided 
    solution name and variable values.

    Parameters
    ----------
    oModule : pywin32 COMObject
        An HFSS "FieldsReporter" module 
    SolutionName : str
        Name of solution to use, for example "Setup1 : LastAdaptive"
    Freq : float
        Frequency at which to evaluate the calculator expression.
    Phase : float
        Phase in degrees at which to evaluate the calculator expression.
    variablesdict : dict
        Dictionary listing the variables and their values that define the 
        design variation, except for 'Freq' and 'Phase'.  Variable names are the keys (strings), and the values are 
        the values of the variable (floats). For example: {'radius': 0.5, 'height': '2.0'}
        
    Returns
    -------
    result : str
        A string representation of the value at the top of the Calculator stack.
        
    """
    solutionname = setupname + " : " + sweepname
    
    variablesarray = ["Freq:=", str(freq) + 'Hz', "Phase:=", str(phase) + 'deg']
    
    for key in variablesdict:
        variablesarray += [str(key) + ':=', str(variablesdict[key])]
        
    print('solutionname: ' + str(solutionname))
    print('variablesarray: ' + str(variablesarray))
    
    result = oModule.GetTopEntryValue(solutionname, variablesarray)

    return result

if __name__ == "__main__":
    import doctest
    doctest.testmod()