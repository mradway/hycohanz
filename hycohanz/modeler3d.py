# -*- coding: utf-8 -*-
"""
HFSS functions that use the 3D modeler. Functions in this module correspond 
more or less to the functions described in the HFSS Scripting Guide, 
Section "3D Modeler Editor Script Commands".

At last count there were 31 functions implemented out of 93.
"""

from __future__ import division, print_function, unicode_literals, absolute_import

import warnings

from hycohanz.expression import Expression as Ex

warnings.simplefilter('default')

def get_matched_object_name(oEditor, name_filter="*"):
    """
    Returns a list of objects that match the input filter.
    
    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.
    name_filter : str
        Wildcard text to search.  Should contain '*' wildcard character for tuples.
        
    Returns
    -------
    part : list
        List of object names matched to the filter.
    
    """

    selections = oEditor.GetMatchedObjectName(name_filter)
    
    return list(selections)

def assign_material(oEditor, partlist, MaterialName="vacuum", SolveInside=True):
    """
    Assign a material to the specified objects. Only the MaterialName and 
    SolveInside parameters of <AttributesArray> are supported.
    
    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.
    partlist : list
        List of part name strings to which the material is applied.
    
    Returns
    -------
    None
    """
    selectionsarray = ["NAME:Selections", 
                       "Selections:=", ','.join(partlist)]
    
    attributesarray = ["NAME:Attributes", 
                       "MaterialName:=", MaterialName, 
                       "SolveInside:=", SolveInside]
    
    oEditor.AssignMaterial(selectionsarray, attributesarray)

def create_rectangle(   oEditor, 
                        xs, 
                        ys, 
                        zs, 
                        width, 
                        height, 
                        WhichAxis='Z', 
                        Name='Rectangle1', 
                        Flags='', 
                        Color=(132, 132, 193), 
                        Transparency=0, 
                        PartCoordinateSystem='Global',
                        UDMId='',
                        MaterialValue='"vacuum"',
                        SolveInside=True,
                        IsCovered=True):
    """
    Draw a rectangle.
    
    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.
    xs : float or hycohanz Expression object
    ys : float or hycohanz Expression object
    zs : float or hycohanz Expression object
        The x, y, and z coordinates of the center of the circle.
    width : float or hycohanz Expression object
        x-dimension of the rectangle
    height : float or hycohanz Expression object
        y-dimension of the rectangle
    WhichAxis : str
        The axis normal to the circle.  Can be 'X', 'Y', or 'Z'.
    Name : str
        The requested name of the object.  If this is not available, HFSS 
        will assign a different name, which is returned by this function.
    Flags : str
        Flags associated with this object.  See HFSS Scripting Guide for details.
    Color : tuple of length=3
        RGB components of the circle
    Transparency : float between 0 and 1
        Fractional transparency.  0 is opaque and 1 is transparent.
    PartCoordinateSystem : str
        The name of the coordinate system in which the object is drawn.
    MaterialName : str
        Name of the material to assign to the object.  Name must be surrounded 
        by double quotes.
    SolveInside : bool
        Whether to mesh the interior of the object and solve for the fields 
        inside.
    IsCovered : bool
        Whether the rectangle is has a surface or has only edges.
        
    Returns
    -------
    str
        The actual name of the created object.
        
    """
    RectangleParameters = [ "NAME:RectangleParameters",
                            "IsCovered:=", IsCovered,
                            "XStart:=", Ex(xs).expr,
                            "YStart:=", Ex(ys).expr,
                            "ZStart:=", Ex(zs).expr,
                            "Width:=", Ex(width).expr,
                            "Height:=", Ex(height).expr,
                            "WhichAxis:=", WhichAxis]

    Attributes = [  "NAME:Attributes",
                    "Name:=", Name,
                    "Flags:=", Flags,
                    "Color:=", "({r} {g} {b})".format(r=Color[0], g=Color[1], b=Color[2]),
                    "Transparency:=", Transparency,
                    "PartCoordinateSystem:=", PartCoordinateSystem,
                    "UDMId:=", UDMId,
                    "MaterialValue:=", MaterialValue,
                    "SolveInside:=", SolveInside]
                    
    return oEditor.CreateRectangle(RectangleParameters, Attributes)


def create_EQbasedcurve(   oEditor, 
                        xt, 
                        yt, 
                        zt, 
                        tstart, 
                        tend,
                        numpoints,
                        Version=1, 
                        Name='EQcurve1', 
                        Flags='', 
                        Color=(132, 132, 193), 
                        Transparency=0, 
                        PartCoordinateSystem='Global',
                        UDMId='',
                        MaterialValue='"vacuum"',
                        SolveInside=True):
    """
    Draw an equation based curve.
    
    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.
    xt : equation for the x axis, equation must be a string
    yt : equation for the y axis, equation must be a string
    zt : equation for the z axis, equation must be a string
    tstart : the start of variable _t, must be a string
    tend : then end of variable _t, must be a string
    numpoints : The number of points that the curve is made of,
        enter 0 for a continuous line, must be a string
    Version Number : ?
        Unsure what this does
    Name : str
        The requested name of the object.  If this is not available, HFSS 
        will assign a different name, which is returned by this function.
    Flags : str
        Flags associated with this object.  See HFSS help for details.
    Color : tuple of length=3
        RGB components of the circle
    Transparency : float between 0 and 1
        Fractional transparency.  0 is opaque and 1 is transparent.
    PartCoordinateSystem : str
        The name of the coordinate system in which the object is drawn.
    MaterialName : str
        Name of the material to assign to the object.  Name must be surrounded 
        by double quotes.
    SolveInside : bool
        Whether to mesh the interior of the object and solve for the fields 
        inside.
        
    Returns
    -------
    str
        The actual name of the created object.
        
    """
    EquationCurveParameters = [ "NAME:EquationBasedCurveParameters",
                            "XtFunction:=", xt,
                            "YtFunction:=", yt,
                            "ZtFunction:=", zt,
                            "tStart:=", tstart,
                            "tEnd:=", tend,
                            "NumOfPointsOnCurve:=", numpoints,
                            "Version:=", Version]

    Attributes = [  "NAME:Attributes",
                    "Name:=", Name,
                    "Flags:=", Flags,
                    "Color:=", "({r} {g} {b})".format(r=Color[0], g=Color[1], b=Color[2]),
                    "Transparency:=", Transparency,
                    "PartCoordinateSystem:=", PartCoordinateSystem,
                    "UDMId:=", UDMId,
                    "MaterialValue:=", MaterialValue,
                    "SolveInside:=", SolveInside]
                    
    return oEditor.CreateEquationCurve(EquationCurveParameters, Attributes)

def create_circle(oEditor, xc, yc, zc, radius, 
                  WhichAxis='Z', 
                  NumSegments=0,
                  Name='Circle1',
                  Flags='',
                  Color=(132, 132, 193),
                  Transparency=0,
                  PartCoordinateSystem='Global',
                  MaterialName='"vacuum"',
                  Solveinside=True):
    """
    Create a circle primitive.
    
    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.
    xc : float or hycohanz Expression object
    yc : float or hycohanz Expression object
    zc : float or hycohanz Expression object
        The x, y, and z coordinates of the center of the circle.
    radius : float
        The radius of the circle.
    WhichAxis : str
        The axis normal to the circle.  Can be 'X', 'Y', or 'Z'.
    NumSegments : int
        If 0, the circle is not segmented.  Otherwise, the circle is 
        segmented into NumSegments sides.
    Name : str
        The requested name of the object.  If this is not available, HFSS 
        will assign a different name, which is returned by this function.
    Flags : str
        Flags associated with this object.  See HFSS help for details.
    Color : tuple of length=3
        RGB components of the circle
    Transparency : float between 0 and 1
        Fractional transparency.  0 is opaque and 1 is transparent.
    PartCoordinateSystem : str
        The name of the coordinate system in which the object is drawn.
    MaterialName : str
        Name of the material to assign to the object.  Name must be surrounded 
        by double quotes.
    SolveInside : bool
        Whether to mesh the interior of the object and solve for the fields 
        inside.
        
    Returns
    -------
    str
        The actual name of the created object.
    """
    circleparams = ["NAME:CircleParameters", 
                    "XCenter:=", Ex(xc).expr, 
                    "YCenter:=", Ex(yc).expr, 
                    "ZCenter:=", Ex(zc).expr, 
                    "Radius:=", Ex(radius).expr, 
                    "WhichAxis:=", str(WhichAxis), 
                    "NumSegments:=", str(NumSegments)]

    attributesarray = ["NAME:Attributes", 
                       "Name:=", Name, 
                       "Flags:=", Flags, 
                       "Color:=", "({r} {g} {b})".format(r=Color[0], g=Color[1], b=Color[2]), 
                       "Transparency:=", str(Transparency), 
                       "PartCoordinateSystem:=", PartCoordinateSystem, 
                       "MaterialName:=", MaterialName, 
                       "Solveinside:=", Solveinside]

    return oEditor.CreateCircle(circleparams, attributesarray)

def create_sphere(oEditor, x, y, z, radius,
                  Name="Sphere1",
                  Flags="",
                  Color=(132, 132, 193),
                  Transparency=0,
                  PartCoordinateSystem="Global",
                  UDMId="",
                  MaterialValue='"vacuum"',
                  SolveInside=True):
    """
    Create a sphere primitive.
    
    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.
    x : float or hycohanz Expression object
        x position in Cartesian coordinates.
    y : float or hycohanz Expression object
        y position in Cartesian coordinates.
    z : float or hycohanz Expression object
        z position in Cartesian coordinates.
    radius : float
        The sphere radius.
    Name : str
        The requested name of the object.  HFSS doesn't necessarily honor this.
    Flags : str
        Flags to attach to the object.  See the HFSS help for explanation 
        of this parameter.
    Color : tuple of ints
        The RGB indices corresponding to the desired color of the object.
    Transparency : float between 0 and 1
        Fractional transparency.  0 is opaque and 1 is transparent.
    PartCoordinateSystem : str
        The name of the coordinate system in which the object is drawn.
    UDMId : str
        Unknown use.  See HFSS documentation for explanation.
    MaterialValue : str
        Name of the material to assign to the object
    SolveInside : bool
        Whether to mesh the interior of the object and solve for the fields 
        inside.
        
    Returns
    -------
    part : str
        The actual name assigned by HFSS to the part.
        
    """
    sphereparametersarray = ["NAME:SphereParameters", 
                             "XCenter:=", Ex(x).expr, 
                             "YCenter:=", Ex(y).expr, 
                             "ZCenter:=", Ex(z).expr, 
                             "Radius:=", Ex(radius).expr]
    
    attributesarray = ["NAME:Attributes", 
                       "Name:=",  Name, 
                       "Flags:=", Flags, 
                       "Color:=", "({r} {g} {b})".format(r=Color[0], g=Color[1], b=Color[2]), 
                       "Transparency:=", Transparency, 
                       "PartCoordinateSystem:=", PartCoordinateSystem, 
                       "UDMId:=", UDMId, 
                       "MaterialValue:=", MaterialValue, 
                       "SolveInside:=", SolveInside]
    
    part = oEditor.CreateSphere(sphereparametersarray, attributesarray)
    
    return part

def create_box( oEditor, 
                xpos, 
                ypos, 
                zpos, 
                xsize, 
                ysize,
                zsize,
                Name='Box1', 
                Flags='', 
                Color=(132, 132, 193), 
                Transparency=0, 
                PartCoordinateSystem='Global',
                UDMId='',
                MaterialValue='"vacuum"',
                SolveInside=True,
                IsCovered=True,
                ):
    """
    Draw a 3D box.

    Note:  This function was contributed by C. A. Donado Morcillo.

    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.
    xpos : float or hycohanz Expression object
    ypos : float or hycohanz Expression object
    zpos : float or hycohanz Expression object
        The x, y, and z coordinates of the base point of the box.
    xsize : float or hycohanz Expression object
    ysize : float or hycohanz Expression object
    zsize : float or hycohanz Expression object
        x-, y-, and z-dimensions of the box
    Name : str
        The requested name of the object.  If this is not available, HFSS 
        will assign a different name, which is returned by this function.
    Flags : str
        Flags associated with this object.  See HFSS Scripting Guide for details.
    Color : tuple of length=3
        RGB components of the circle
    Transparency : float between 0 and 1
        Fractional transparency.  0 is opaque and 1 is transparent.
    PartCoordinateSystem : str
        The name of the coordinate system in which the object is drawn.
    MaterialName : str
        Name of the material to assign to the object.  Name must be surrounded 
        by double quotes.
    SolveInside : bool
        Whether to mesh the interior of the object and solve for the fields 
        inside.
    IsCovered : bool
        Whether the rectangle is has a surface or has only edges.

    Returns
    -------
    str
        The actual name of the created object.

    """
    BoxParameters = [ "NAME:BoxParameters",
                    "XPosition:=", Ex(xpos).expr,
                    "YPosition:=", Ex(ypos).expr,
                    "ZPosition:=", Ex(zpos).expr,
                    "XSize:=", Ex(xsize).expr,
                    "YSize:=", Ex(ysize).expr,
                    "ZSize:=", Ex(zsize).expr]

    Attributes = [  "NAME:Attributes",
                    "Name:=", Name,
                    "Flags:=", Flags,
                    "Color:=", "({r} {g} {b})".format(r=Color[0], g=Color[1], b=Color[2]),
                    "Transparency:=", Transparency,
                    "PartCoordinateSystem:=", PartCoordinateSystem,
                    "UDMId:=", UDMId,
                    "MaterialValue:=", MaterialValue,
                    "SolveInside:=", SolveInside]

    return oEditor.CreateBox(BoxParameters, Attributes)    

def create_polyline(oEditor, x, y, z, Name="Polyline1", 
                                Flags="", 
                                Color="(132 132 193)", 
                                Transparency=0, 
                                PartCoordinateSystem="Global", 
                                UDMId="", 
                                MaterialValue='"vacuum"', 
                                SolveInside=True, 
                                IsPolylineCovered=True, 
                                IsPolylineClosed=True, 
                                XSectionBendType="Corner", 
                                XSectionNumSegments="0", 
                                XSectionHeight="0mm", 
                                XSectionTopWidth="0mm", 
                                XSectionWidth="0mm",
                                XSectionOrient="Auto",
                                XSectionType="None",
                                SegmentType="Line",
                                NoOfPoints=2):
    """ 
    Draw a polyline.
    
    Warning:  HFSS apparently throws an exception whenever IsPolylineClosed=False
    Warning:  HFSS 13 crashes when you click on the last segment in the model tree.
    
    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which to perform the operation
    x : array_like
        The x locations of the polyline vertices. Can have numeric or string elements.
    y : array_like
        The y locations of the polyline vertices. Can have numeric or string elements.
    z : array_like
        The z locations of the polyline vertices. Can have numeric or string elements.
    Name : str
        Requested name of the polyline
    Flags : str
        Certain flags that can be set.  See HFSS Scripting Manual for details.
    Transparency : float
        Fractional transparency of the object.  0 is opaque, 1 is transparent.
    PartCoordinateSystem : str
        Coordinate system to use in constructing the object.
    UDMId : str
        TODO:  Add documentation here.
    MaterialValue : str
        Name of the material to assign to the object.
    SolveInside : bool
        Whether fields are computed inside the object.
    IsPolylineCovered : bool
        Whether the polyline is covered.
    IsPolylineClosed : bool
        Whether the polyline should be considered closed.
    TODO:  finish documentation of this function.
        
    Returns
    -------
    polyname : str
        Actual name of the polyline
        
    
    x, y, and z are lists with numeric or string elements.
    
    Example Usage
    -------------
    >>> import Hyphasis as hfss
    >>> [oAnsoftApp, oDesktop] = hfss.setup_interface()
    >>> oProject = hfss.new_project(oDesktop)
    >>> oDesign = hfss.insert_design(oProject, "HFSSDesign1", "DrivenModal")
    >>> oEditor = hfss.set_active_editor(oDesign, "3D Modeler")
    >>> tri = hfss.create_polyline(oEditor, [0, 1, 0], [0, 0, 1], [0, 0, 0])
    """
    # Augment the polyline points vector by appending the first element to 
    # the last.  This gives polyline points and N - 1 segments
    xv = list(x) + [list(x)[0]]
    yv = list(y) + [list(y)[0]]
    zv = list(z) + [list(z)[0]]
    
#    print('xv:  ' + str(xv))
    
    Npts = len(xv)
#    print("Npts: " + str(Npts))
    
#    plpoint = []
    polylinepoints = ["NAME:PolylinePoints"]
    
    for n in range(0, Npts):
        if isinstance(xv[n], str):
            xpt = xv[n]
        elif isinstance(xv[n], (float, int)):
            xpt = str(xv[n]) + "meter"
        elif isinstance(xv[n], Ex):
            xpt = xv[n].expr
        else:
            raise TypeError('xv must be of type str, int, float, or Ex')
            
        if isinstance(yv[n], str):
            ypt = yv[n]
        elif isinstance(yv[n], (float, int)):
            ypt = str(yv[n]) + "meter"
        elif isinstance(yv[n], Ex):
            ypt = yv[n].expr
        else:
            raise TypeError('yv must be of type str, int, float, or Ex')
            
        if isinstance(zv[n], str):
            zpt = zv[n]
        elif isinstance(zv[n], (float, int)):
            zpt = str(zv[n]) + "meter"
        elif isinstance(zv[n], Ex):
            zpt = zv[n].expr
        else:
            raise TypeError('zv must be of type str, int, float, or Ex')
        
#        print('xpt:  ' + str(xpt))
#        print('ypt:  ' + str(ypt))
#        print('zpt:  ' + str(zpt))
        polylinepoints.append([["NAME:PLPoint", 
                        "X:=", xpt, 
                        "Y:=", ypt, 
                        "Z:=", zpt]])
    
#    plpoint[0] = ["NAME:PLPoint", "X:=", "-1mm", "Y:=", "0mm", "Z:=", "0mm"]
#    plpoint[1] = ["NAME:PLPoint", "X:=", "0mm", "Y:=", "1mm", "Z:=", "0mm"]
#    plpoint[2] = ["NAME:PLPoint", "X:=", "1mm", "Y:=", "0mm", "Z:=", "0mm"]

#    polylinepoints.append(plpoint)
#    print(polylinepoints)
    
##    polylinepoints = ["NAME:PolylinePoints", plpoint[0], plpoint[1], plpoint[2]]
#    plsegment = []
    polylinesegments = ["NAME:PolylineSegments"]
    if IsPolylineClosed == True:
        Nsegs = Npts - 1 
    else:
        Nsegs = Npts - 1
    
#    print("Nsegs: " + str(Nsegs))    
    for n in range(0, Nsegs):
        polylinesegments.append(["NAME:PLSegment", 
                                 "SegmentType:=", SegmentType, 
                                 "StartIndex:=", n, 
                                 "NoOfPoints:=", NoOfPoints])
        
#    print(plsegment)
        
#    plsegment1 = ["NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 0, "NoOfPoints:=", 2]
#    plsegment2 = ["NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 1, "NoOfPoints:=", 2]
#    plsegment3 = ["NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 2, "NoOfPoints:=", 2]
    
#    polylinesegments = ["NAME:PolylineSegments", plsegment1, plsegment2, plsegment3]

#    polylinesegments.append(plsegment)
    
#    print(polylinesegments)
    
    polylinexsection = ["NAME:PolylineXSection", 
                        "XSectionType:=", XSectionType, 
                        "XSectionOrient:=", XSectionOrient, 
                        "XSectionWidth:=", XSectionWidth, 
                        "XSectionTopWidth:=", XSectionTopWidth, 
                        "XSectionHeight:=", XSectionHeight, 
                        "XSectionNumSegments:=", XSectionNumSegments, 
                        "XSectionBendType:=", XSectionBendType]

    polylineparams = ["NAME:PolylineParameters", 
                      "IsPolylineCovered:=", IsPolylineCovered, 
                      "IsPolylineClosed:=", IsPolylineClosed, 
                      polylinepoints, 
                      polylinesegments]#, 
#                      polylinexsection]

    polylineattribs = ["NAME:Attributes", 
                       "Name:=", Name, 
                       "Flags:=", Flags, 
                       "Color:=", Color,
                       "Transparency:=", Transparency, 
                       "PartCoordinateSystem:=", PartCoordinateSystem, 
                       "UDMId:=", UDMId, 
                       "MaterialValue:=", MaterialValue, 
                       "SolveInside:=",  SolveInside]
    
    polyname = oEditor.CreatePolyline(polylineparams, polylineattribs)

    return polyname

def get_selections(oEditor):
    """
    Get a list of the currently-selected objects in the design.  
    
    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.
        
    Returns
    -------
    selectionlist : list
        List of the selectable objects in the design?
    """
    return oEditor.GetSelections()

def move(oEditor, partlist, x, y, z, NewPartsModelFlag="Model"):
    """
    Move specified parts.
    
    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.
    partlist : list
        List of part name strings to be mirrored.
    x : float
        x displacement in Cartesian coordinates.
    y : float
        y displacement in Cartesian coordinates.
    z : float
        z displacement in Cartesian coordinates.
        
    Returns
    -------
    None
    """
    selectionsarray = ["NAME:Selections", 
                       "Selections:=", ','.join(partlist), 
                       "NewPartsModelFlag:=", NewPartsModelFlag]
                      
    moveparametersarray = ["NAME:TranslateParameters", 
                           "TranslateVectorX:=", str(Ex(x).expr), 
                           "TranslateVectorY:=", str(Ex(y).expr), 
                           "TranslateVectorZ:=", str(Ex(z).expr)]
    
    oEditor.Move(selectionsarray, moveparametersarray)

def get_object_name(oEditor, index):
    """
    Return the object name corresponding to the zero-based creation-order index.
    
    Note:  This is NOT the inverse of get_object_id_by_name()!
    """
    return oEditor.GetObjectName(index)

def copy(oEditor, partlist):
    """
    Copy specified parts to the clipboard.
    
    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS Editor in which to perform the operation
    partlist : list of strings
        The parts to copy
        
    Returns
    -------
    None
    """
    selectionsarray = ["NAME:Selections", 
                       "Selections:=", ','.join(partlist), 
                       "NewPartsModelFlag:=", "Model"]
                       
    oEditor.Copy(selectionsarray)
    
def get_object_id_by_name(oEditor, objname):
    """
    Return the object ID of the specified part.
    
    Note:  This is NOT the inverse of get_object_name()!
    """
    return oEditor.GetObjectIDByName(objname)

def paste(oEditor):
    """
    Paste a design in the active project from the clipboard.
    
    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS Editor in which to perform the operation

    Returns
    -------
    pastelist : list
        List of parts that are pasted
    """
    pastelist = oEditor.Paste()
    return pastelist

def imprint(oEditor, blanklist, toollist, KeepOriginals=False):
    """
    Imprint an object onto another object.
    
    Note:  This function is undocumented in the HFSS Scripting Guide.
    """
    imprintselectionsarray = [ "NAME:Selections", 
                               "Blank Parts:=", ','.join(blanklist), 
                               "Tool Parts:=", ','.join(toollist)]
    
    imprintparams = ["NAME:ImprintParameters", 
                     "KeepOriginals:=", KeepOriginals]
    
    return oEditor.Imprint(imprintselectionsarray, imprintparams)

def mirror(oEditor, partlist, base, normal):
    """
    Mirror specified parts about a given base point with respect to a given 
    plane.
    
    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.
    partlist : list
        List of part name strings to be mirrored.
    base : list
        Mirror base point in Cartesian coordinates.
    normal : list
        Mirror plane normal in Cartesian coordinates.
        
    Returns
    -------
    None
    """
    selectionsarray = ["NAME:Selections", 
                       "Selections:=", ','.join(partlist), 
                       "NewPartsModelFlag:=", "Model"]
                       
    mirrorparamsarray = ["NAME:MirrorParameters", 
                         "MirrorBaseX:=", str(base[0]) + "meter", 
                         "MirrorBaseY:=", str(base[1]) + "meter", 
                         "MirrorBaseZ:=", str(base[2]) + "meter", 
                         "MirrorNormalX:=", str(normal[0]) + "meter", 
                         "MirrorNormalY:=", str(normal[1]) + "meter", 
                         "MirrorNormalZ:=", str(normal[2]) + "meter"]
                       
    oEditor.Mirror(selectionsarray, mirrorparamsarray)

def sweep_along_vector(oEditor, obj_name_list, x, y, z):
    """
    Sweeps the specified 1D or 2D parts along a vector.
    
    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.
    obj_name_list : list
        List of part name strings to be scaled.
    x : float
        x component of the sweep vector.
    y : float
        y component of the sweep vector.
    z : float
        z component of the sweep vector.
        
    Returns
    -------
    None
    """
    selections = ", ".join(obj_name_list)
    
#    print(selections)
    
    oEditor.SweepAlongVector(["NAME:Selections", 
                              "Selections:=", selections, 
                              "NewPartsModelFlag:=", "Model"], 
                             ["NAME:VectorSweepParameters", 
                              "DraftAngle:=", "0deg", 
                              "DraftType:=", "Round", 
                              "CheckFaceFaceIntersection:=", False, 
                              "SweepVectorX:=", Ex(x).expr, 
                              "SweepVectorY:=", Ex(y).expr, 
                              "SweepVectorZ:=", Ex(z).expr])

    return get_selections(oEditor)

def rotate(oEditor, partlist, axis, angle):
    """
    Rotate specified parts.
    
    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.
    partlist : list
        List of part name strings to be rotated.
    axis : str
        Rotation axis.
    angle : float
        Rotation angle in radians
   
    Returns
    -------
    None
    """
    selectionsarray = ["NAME:Selections", 
                       "Selections:=", ','.join(partlist), 
                       "NewPartsModelFlag:=", "Model"]
                       
    rotateparametersarray = ["NAME:RotateParameters", 
                             "RotateAxis:=", axis, 
                             "RotateAngle:=", Ex(angle).expr]
                             
    oEditor.Rotate(selectionsarray, rotateparametersarray)

def subtract(oEditor, blanklist, toollist, KeepOriginals=False):
    """
    Subtract the specified objects.
    
    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.        
    partlist : list
        List of part name strings to be subtracted.
    toollist : list
        List of part name strings to subtract from partlist.
    KeepOriginals : bool
        Whether to clone the tool parts.
    
    Returns
    -------
    objname : str
        Name of object created by the subtract operation
    """
#    blankliststr = ""
#    for item in blanklist:
#        blankliststr += (',' + item)
#        
#    toolliststr = ""
#    for item in toollist:
#        toolliststr += (',' + item)
    
    subtractselectionsarray = ["NAME:Selections", 
                               "Blank Parts:=", ','.join(blanklist), 
                               "Tool Parts:=", ','.join(toollist)]
    
    subtractparametersarray = ["NAME:SubtractParameters", 
                               "KeepOriginals:=", KeepOriginals]
    
    oEditor.Subtract(subtractselectionsarray, subtractparametersarray)
    
    return blanklist[0]

def unite(oEditor, partlist, KeepOriginals=False):
    """
    Unite the specified objects.
    
    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.        
    partlist : list
        List of part name strings to be united.
    KeepOriginals : bool
        Whether to keep the original parts for subsequent operations.
    
    Returns
    -------
    objname : str
        Name of object created by the unite operation
        
    Examples
    --------
    >>> import Hyphasis as hfss
    >>> [oAnsoftApp, oDesktop] = hfss.setup_interface()
    >>> oProject = hfss.new_project(oDesktop)
    >>> oDesign = hfss.insert_design(oProject, "HFSSDesign1", "DrivenModal")
    >>> oEditor = hfss.set_active_editor(oDesign, "3D Modeler")
    >>> tri1 = hfss.create_polyline(oEditor, [0, 1, 0], [0, 0, 1], [0, 0, 0])
    >>> tri2 = hfss.create_polyline(oEditor, [0, -1, 0], [0, 0, 1], [0, 0, 0])
    >>> tri3 = hfss.unite(oEditor, [tri1, tri2])
    """
#    partliststr = ""
#    for item in partlist:
#        partliststr += (',' + item)
    
    selectionsarray = ["NAME:Selections", "Selections:=", ','.join(partlist)]
    
    uniteparametersarray = ["NAME:UniteParameters", "KeepOriginals:=", KeepOriginals]
    
    oEditor.Unite(selectionsarray, uniteparametersarray)
    
    return partlist[0]

def scale(oEditor, partlist, x, y, z):
    """
    Scale specified parts.
    
    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.
    partlist : list
        List of part name strings to be scaled.
    x : float
        x scaling factor in Cartesian coordinates.
    y : float
        y scaling factor in Cartesian coordinates.
    z : float
        z scaling factor in Cartesian coordinates.
        
    Returns
    -------
    None
    """
    selections = ", ".join(partlist)
    selectionsarray = ["NAME:Selections", 
                       "Selections:=", selections, 
                       "NewPartsModelFlag:=", "Model"]
  
    scaleparametersarray = ["NAME:ScaleParameters", 
                            "ScaleX:=", str(x), 
                            "ScaleY:=", str(y), 
                            "ScaleZ:=", str(z)]
  
    oEditor.Scale(selectionsarray, scaleparametersarray)

def get_object_name_by_faceid(oEditor, faceid):
    """
    Return the object name corresponding to the given face ID.
    
    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.
    faceid : int
        The face ID of the given face.
    
    Returns
    -------
    objname : str
        The name of the object.

    """
    return oEditor.GetObjectNameByFaceID(faceid)

def import_model(oEditor, 
                 sourcefile,
                 HealOption=1,
                 CheckModel=False,
                 Options='-1',
                 FileType='UnRecognized',
                 MaxStitchTol=-1,
                 ImportFreeSurfaces=False):
    """
    Import a 3D model from a file.
    
    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.
    sourcefile : str
        Name of the 3D model file.
        
    Returns
    -------
    None
    
    Notes
    -----
    - This function is barely documented in the HFSS Scripting Guide.
    - No documentation of the optional arguments is given because their 
      equivalents are not documented in the HFSS Scripting Guide.
    - There is no documented way to request a particular name for the created 
      object.
    - There is no documented way to obtain the name assigned to the created 
      object.
    
    Examples
    --------
    >>> import Hyphasis as hfss
    >>> [oAnsoftApp, oDesktop] = hfss.setup_interface()
    >>> oProject = hfss.new_project(oDesktop)
    >>> oDesign = hfss.insert_design(oProject, "HFSSDesign1", "DrivenModal")
    >>> oEditor = hfss.set_active_editor(oDesign, "3D Modeler")
    >>> hfss.import_model(oEditor, "Z:\shared\Parts\MachinescrewCap4-40_375mil\91251A108.SAT")
    >>> hfss.import_model(oEditor, "Z:\shared\Parts\MachinescrewCap4-40_375mil\91251A108.IGS")
    >>> hfss.import_model(oEditor, "Z:\shared\Parts\MachinescrewCap4-40_375mil\91251A108.STEP")
    
    """
    import_params_array =["NAME:NativeBodyParameters", 
                          "HealOption:=", HealOption, 
                          "CheckModel:=", CheckModel, 
                          "Options:=", Options, 
                          "FileType:=", FileType, 
                          "MaxStitchTol:=", MaxStitchTol, 
                          "ImportFreeSurfaces:=", ImportFreeSurfaces, 
                          "SourceFile:=", sourcefile]
    
    oEditor.Import(import_params_array)
    
    return get_selections(oEditor)

def get_edge_by_position(oEditor, bodyname, x, y, z):
    """
    Get the edge of a given body that lies at a given position.
    
    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.
    bodyname : str
        Name of the body whose edge will be returned
    x : float
        x position of the edge
    y : float
        y position of the edge
    z : float
        z position of the edge
        
    Returns
    -------
    edgeid : int
        Id number of the edge.
    """
#    print(Ex(x).expr)
#    print(Ex(y).expr)
#    print(Ex(z).expr)
    positionparameters = ["NAME:EdgeParameters", 
                          "BodyName:=", bodyname,
                          "Xposition:=", Ex(x).expr,
                          "YPosition:=", Ex(y).expr,
                          "ZPosition:=", Ex(z).expr]

    edgeid = oEditor.GetEdgeByPosition(positionparameters)
    
    return edgeid
    
def fillet(oEditor, partlist, edgelist, radius, vertexlist=[], setback=0):
    """
    Create fillets on the given edges.
    
    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.
    partlist : list of strings
        List of part name strings to be filleted.
    edgelist : list of ints
        List of edge indexes to be filleted.
    radius : float
        Radius of the fillet.
    vertexlist : list
        List of vertices to chamfer
    setback : float
        The setback distance.  See the HFSS help for an explanation of this 
        parameter.
    
    Returns
    -------
    None
    
    Examples
    --------
    >>> import Hyphasis as hfss
    >>> [oAnsoftApp, oDesktop] = hfss.setup_interface()
    >>> oProject = hfss.new_project(oDesktop)
    >>> oDesign = hfss.insert_design(oProject, "HFSSDesign1", "DrivenModal")
    >>> oEditor = hfss.set_active_editor(oDesign, "3D Modeler")
    >>> box1 = hfss.create_box(oEditor, 1, 1, 1, 0, 0, 0)
    >>> edge1 = hfss.get_edge_by_position(oEditor, box1, 0.5, 0, 0)
    >>> hfss.fillet(oEditor, [box1], [edge1], 0.25)
    
    """
    selectionsarray = ["NAME:Selections", 
                       "Selections:=", ','.join(partlist), 
                       "NewPartsModelFlag:=", "Model"]
                            
    tempparams = ["NAME:FilletParameters", 
                  "Edges:=", edgelist, 
                  "Vertices:=", vertexlist, 
                  "Radius:=",  Ex(radius).expr, 
                  "Setback:=", str(setback)]
    
    filletparameters = ["NAME:Parameters", tempparams]
                            
    oEditor.Fillet(selectionsarray, filletparameters)
    
def separate_body(oEditor, partlist, NewPartsModelFlag="Model"):
    """
    Separate bodies of the specified multi-lump object
    
    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.
    partlist : list of str
        List of object names to operate upon
    NewPartsModelFlag : str
        See the HFSS Scripting Guide.
    
    Returns
    -------
    newpartlist : list of str
        List of objects created by the operation.
    
    """
    selectionsarray = ["NAME:Selections", 
                       "Selections:=", ",".join(partlist), 
                       "NewPartsModelFlag:=", NewPartsModelFlag]

    oEditor.SeparateBody(selectionsarray)
    
    return (partlist[0],) + get_selections(oEditor)
    
def delete(oEditor, partlist):
    """
    Delete selected objects, coordinate systems, points, planes, and others.
    
    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.
    partlist : list of str
        List of object names to delete
    
    Returns
    -------
    None
    
    """
    selectionsarray = ["NAME:Selections", 
                       "Selections:=", ','.join(partlist)]
                       
    return oEditor.Delete(selectionsarray)


def split(oEditor, partlist, 
          NewPartsModelFlag="Model", 
          SplitPlane='XY', 
          WhichSide="PositiveOnly", 
          SplitCrossingObjectsOnly=False, 
          DeleteInvalidObjects=True):
    """
    Splits specified objects along a plane.
    
    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.
    partlist : list
        List of part name strings to be mirrored.
    SplitPlane : str
        Plane in which to split the part(s).  Allowable values are "XY", "ZX", or "YZ"
    WhichSide : str
        Side to keep.  Allowable values are "Both", "PositiveOnly", or "NegativeOnly"
    SplitCrossingObjectsOnly : bool
        If True, only splits objects that actually cross the split plane.
    DeleteInvalidObjects : bool
        If True, invalid objects generated by the operation are deleted.
        
    Returns
    -------
    None
    """
    selectionsarray = ["NAME:Selections", 
                       "Selections:=", ",".join(partlist), 
                       "NewPartsModelFlag:=", NewPartsModelFlag]
    splittoparams = ["NAME:SplitToParameters", 
                     "SplitPlane:=", SplitPlane, 
                     "WhichSide:=", WhichSide, 
                     "SplitCrossingObjectsOnly:=", SplitCrossingObjectsOnly, 
                     "DeleteInvalidObjects:=", DeleteInvalidObjects]
                       
    return oEditor.Split(selectionsarray, splittoparams)

def get_face_by_position(oEditor, bodyname, x, y, z):
    """
    Get the face of a given body that lies at a given position.
    
    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.
    bodyname : str
        Name of the body whose face will be returned
    x : float
        x position of the face
    y : float
        y position of the face
    z : float
        z position of the face
        
    Returns
    -------
    faceid : int
        Id number of the face.
    """
    positionparameters = ["NAME:Parameters", 
                          "BodyName:=", bodyname,
                          "Xposition:=", Ex(x).expr,
                          "YPosition:=", Ex(y).expr,
                          "ZPosition:=", Ex(z).expr]
                          
    faceid = oEditor.GetFaceByPosition(positionparameters)
    
    return faceid
    
def uncover_faces(oEditor, partlist, dictoffacelists):
    """
    Uncover specified faces.
    
    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.
    partlist : list
        List of part name strings whose faces will be uncovered.
    dictoffacelists : dict
        Dict containing part names as the keys, and lists of integer face 
        ids as the values
        
    Returns
    -------
    None
    """
    selectionsarray = ["NAME:Selections", "Selections:=", ','.join(partlist)]

    uncoverparametersarray = ["NAME:Parameters"]
    for part in partlist:
        uncoverparametersarray += [["NAME:UncoverFacesParameters", "FacesToUncover:=", dictoffacelists[part]]]

    print('selectionsarray:  {s}'.format(s=selectionsarray))
    print('uncoverparametersarray:  {s}'.format(s=uncoverparametersarray))

    oEditor.UncoverFaces(selectionsarray, uncoverparametersarray)
    
def connect(oEditor, partlist):
    """
    Connects specified 1-D parts to form a sheet, or specified 2-D parts to 
    form a volume.:
        
    WARNING:  oEditor.Connect() is a very flaky operation, and the result can 
    depend on the order that the parts are given to the operation among other 
    seemingly random considerations.  It will very often fail on simple 
    connect operations for no apparent reason.  
    
    If you have difficulty with very strange-looking Connect() operations, 
    first try to make your parts have the same number of vertices.  Second, 
    try reversing the order that you give the parts to Connect().
    
    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.        
    partlist : list
        List of part name strings to be connected.
    
    Returns
    -------
    objname : str
        Name of object created by the connect operation
    """
    selectionsarray = ["NAME:Selections", "Selections:=", ','.join(partlist)]
    
    oEditor.Connect(selectionsarray)
    
    return partlist[0]

def rename_part(oEditor, oldname, newname):
    """
    Rename a part.
    
    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed. 
    oldname : str
        The name of the part to rename
    newname : str
        The new name to assign to the part
        
    Returns
    -------
    None
    
    """
    renameparamsarray = ["Name:Rename Data", "Old Name:=", oldname, "New Name:=", newname]
    
    return oEditor.RenamePart(renameparamsarray)

def get_face_ids(oEditor, body_name):
    """
    Get the face id list of a given body name.
    
    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.
    body_name : str
        Name of the body whose face id list will be returned
        
    Returns
    -------
    face_id_list : list of int
        list with face Id numbers of body_name
    """

    face_id_list = list(oEditor.GetFaceIDs(body_name))
    return map(int,face_id_list)

