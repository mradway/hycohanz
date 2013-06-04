# -*- coding: utf-8 -*-
"""
The HFSS expression generator.

"""

from __future__ import division, print_function, unicode_literals, absolute_import

import warnings

warnings.simplefilter('default')

class Expression(object):
    """
    An HFSS expression.
    
    This object enables manipulation of HFSS expressions using Python 
    arithmetic operators, which is much more convenient than manipulating 
    their string representation.
    
    Parameters
    ----------
    expr : str
        Initialize the expression using its string representation.
    
    Attributes
    ----------
    expr : str
        The string representation of the expression object.
    
    Raises
    ------
    NotImplementedError
        For operations involving floor division (Python 2 '/' or 
        Python 3 '//') 
    
    """
    def __init__(self, expr):
        if isinstance(expr, Expression):
            self.expr = expr.expr
        else:
            self.expr = str(expr)
        
    def __add__(self, y):
        """
        Overloads the addition (+) operator.
        """
        if isinstance(y, Expression):
            return Expression('(' + self.expr + ') + ' + str(y.expr))
        else:
            return Expression('(' + self.expr + ') + ' + str(y))
        
    def __sub__(self, y):
        """
        Overloads the subtraction (-) operator.
        """
        if isinstance(y, Expression):
            return Expression('(' + self.expr + ') - ' + str(y.expr))
        else:
            return Expression('(' + self.expr + ') - ' + str(y))
        
    def __mul__(self, y):
        """
        Overloads the multiplication (*) operator.
        """
        if isinstance(y, Expression):
            return Expression('(' + self.expr + ') * ' + str(y.expr))
        else:
            return Expression('(' + self.expr + ') * ' + str(y))
        
    def __truediv__(self, y):
        """
        Overloads the Python 3 division (/) operator.
        """
        if isinstance(y, Expression):
            return Expression('(' + self.expr + ') / ' + str(y.expr))
        else:
            return Expression('(' + self.expr + ') / ' + str(y))
            
    def __div__(self, y):
        """
        Overloads the Python 3 floor division (//) operator.
        """
        raise NotImplementedError(""""Classic" division is not implemented by 
design.  Please use from __future__ import division in the calling code.""")
        
    def __neg__(self):
        """
        Overloads the negation (-) operator.
        """
        return Expression('-(' + self.expr + ')')

if __name__ == "__main__":
    import doctest
    doctest.testmod()