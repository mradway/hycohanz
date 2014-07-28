# -*- coding: utf-8 -*-
"""
Functions in this module correspond more or less to the functions described 
in the HFSS Scripting Guide (v 2013.11), Section "Material Script Commands".

At last count there were 2 functions implemented out of 5.
"""
from __future__ import division, print_function, unicode_literals, absolute_import

import warnings

from hycohanz.desktop import get_active_project

warnings.simplefilter('default')


def add_material(oDesktop,
                material_name,
                rel_permittivity=1,
                rel_permeability =1,
                cond=0,
                diel_loss_tan=0,
                mag_loss_tan=0,
                mag_saturation=0,
                lande_g=2,
                delta_h=0
                ):
    """
    Add Material.

    Parameters
    ----------
    oDesktop : pywin32 COMObject
        HFSS Desktop object.
    material_name : str
        Name of the added material.
    rel_permittivity : float
    rel_permeability : float
    cond : float
    diel_loss_tan : float
    mag_loss_tan : float
    mag_saturation : float
    lande_g : float
    delta_h : float
        The relative permittivity, relative permeability, electric 
        conductivity, dielectric loss tangent, magnetic loss tangent, 
        magnetic saturation, Lande G factor, and delta_h associated with 
        the added material.
    
    Returns
    -------
    None
    
    Examples
    --------
    >>> import Hyphasis as hfss
    >>> 
    """
    oProject = get_active_project(oDesktop)
    if does_material_exist(oProject,material_name):
        msg = material_name + " already exists in the local library. No material was created"
        warnings.warn(msg)
        return msg
    else:
        mat_param = ["NAME:"+material_name,
                    "permittivity:=", rel_permittivity,
                    "permeability:=", rel_permeability,
                    "conductivity:=", cond,
                    "dielectric_loss_tangent:=", diel_loss_tan, 
                    "magnetic_loss_tangent:=", mag_loss_tan, 
                    "saturation_mag:=", mag_saturation,
                    "lande_g_factor:=", lande_g,
                    "delta_H:=", delta_h]
        oDefinitionManager = oProject.GetDefinitionManager()
        return oDefinitionManager.AddMAterial(mat_param)


def does_material_exist(oProject,material_name):
    """
    Check if material exists.

    Parameters
    ----------
    oDesktop : pywin32 COMObject
        HFSS Desktop object.
    
    Returns
    -------
    Bool
    
    Examples
    --------
    >>> import Hyphasis as hfss
    >>> 
    
    """
    oDefinitionManager = oProject.GetDefinitionManager()
    return oDefinitionManager.DoesMaterialExist(material_name)
