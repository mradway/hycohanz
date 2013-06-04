import hycohanz as hfss

raw_input('Press "Enter" to connect to HFSS.>')

[oAnsoftApp, oDesktop] = hfss.setup_interface()

raw_input('Press "Enter" to create a new project.>')

oProject = hfss.new_project(oDesktop)

raw_input('Press "Enter" to insert a new DrivenModal design named HFSSDesign1.>')

oDesign = hfss.insert_design(oProject, "HFSSDesign1", "DrivenModal")

raw_input('Press "Enter" to set the active editor to "3D Modeler" (The default and only known correct value).>')

oEditor = hfss.set_active_editor(oDesign)

raw_input('Press "Enter" to draw a red rectangle named Rectangle1.>')

objname = hfss.create_rectangle(
    oEditor,  
    1, 
    2, 
    3, 
    4,
    5,
    Name='Rectangle1',
    Color=(255, 0, 0))

raw_input('Press "Enter" to sweep the rectangle.')

hfss.sweep_along_vector(oEditor, [objname], 0, 0, 1)

raw_input('Press "Enter" to grab one of the edges to fillet.> ')

edgeid = hfss.get_edge_by_position(oEditor, objname, 1, 7, 3.5)

print('edgeid: ' + str(edgeid))

raw_input('Press "Enter" to fillet the edge.> ')

hfss.fillet(oEditor, [objname], [edgeid], 1)

raw_input('Press "Enter" to quit HFSS.>')

hfss.quit_application(oDesktop)

del oEditor
del oDesign
del oProject
del oDesktop
del oAnsoftApp
