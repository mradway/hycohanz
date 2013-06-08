from __future__ import division, print_function, unicode_literals, absolute_import

import hycohanz as hfss

raw_input('Press "Enter" to connect to HFSS.>')

[oAnsoftApp, oDesktop] = hfss.setup_interface()

raw_input('Press "Enter" to create a new project.>')

oProject = hfss.new_project(oDesktop)

raw_input('Press "Enter" to insert a new DrivenModal design named HFSSDesign1.>')

oDesign = hfss.insert_design(oProject, "HFSSDesign1", "DrivenModal")

raw_input('Press "Enter" to set the active editor to "3D Modeler" (The default and only known correct value).>')

oEditor = hfss.set_active_editor(oDesign)

raw_input('Press "Enter" to insert some circle properties into the design.>')

hfss.add_property(oDesign, "xcenter", hfss.Expression("1m"))
hfss.add_property(oDesign, "ycenter", hfss.Expression("2m"))
hfss.add_property(oDesign, "zcenter", hfss.Expression("0m"))
hfss.add_property(oDesign, "diam", hfss.Expression("6m"))

raw_input('Press "Enter" to draw a sphere using the properties.>')

objname = hfss.create_sphere(
    oEditor, 
    hfss.Expression("xcenter"), 
    hfss.Expression("ycenter"), 
    hfss.Expression("zcenter"), 
    hfss.Expression("diam")/2,
    )

raw_input('Press "Enter" to split the sphere into two hemispheres.>')

hfss.split(oEditor, [objname], SplitPlane='XY', WhichSide='Both')

raw_input('Press "Enter" to quit HFSS.>')

hfss.quit_application(oDesktop)

del oEditor
del oDesign
del oProject
del oDesktop
del oAnsoftApp
