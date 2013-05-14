# -*- coding: utf-8 -*-
"""
Created on Sun Jan 22 22:34:02 2012

@author: radway
"""

from __future__ import division, print_function, unicode_literals, absolute_import

import warnings

from hycohanz.expression import Expression as Ex

warnings.simplefilter('default')

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
