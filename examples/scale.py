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

raw_input('Press "Enter" to draw a red three-vertex polyline named Triangle1.>')

obj = hfss.create_polyline(oEditor, [1, 0, -1], 
                                    [0, 1, 0], 
                                    [0, 0, 0], 
                                    Name='Triangle1',
                                    Color="(255 0 0)")

print("We'll now scale the triangle by a factor of 2 in the x-axis, 3 in the y-axis, and 4 in the z-axis.")

raw_input('Press "Enter" to scale the part.>')

hfss.scale(oEditor, [obj], 2, 3, 4)

raw_input('Press "Enter" to quit HFSS.>')

hfss.quit_application(oDesktop)

del oEditor
del oDesign
del oProject
del oDesktop
del oAnsoftApp

