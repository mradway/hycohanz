# -*- coding: utf-8 -*-
"""
HFSS Fields Calculator functions.

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
    


def enter_qty(oFieldsReporter, FieldQuantityString):
    """
    Enter a field quantity.
    
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
    
    
    


if __name__ == "__main__":
    import doctest
    doctest.testmod()