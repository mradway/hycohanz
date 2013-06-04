import hycohanz as hfss

raw_input('Press "Enter" to connect to HFSS.>')

[oAnsoftApp, oDesktop] = hfss.setup_interface()

raw_input('Press "Enter" to create a new project.>')

oProject = hfss.new_project(oDesktop)

raw_input('Press "Enter" to insert a new DrivenModal design named HFSSDesign1.>')

oDesign = hfss.insert_design(oProject, "HFSSDesign1", "DrivenModal")

raw_input('Press "Enter" to set the active editor to "3D Modeler" (The default and only known correct value).>')

oEditor = hfss.set_active_editor(oDesign)

raw_input('Press "Enter" to draw two red three-vertex polylines named Triangle1 and Triangle2.>')

objname1 = hfss.create_polyline(
    oEditor, [1, 0, -1], 
    [0, 1, 0], 
    [0, 0, 0], 
    Name='Triangle1',
    Color="(255 0 0)",
    )
    
objname2 = hfss.create_polyline(
    oEditor, [1, 0, -1], 
    [0, 1, 0], 
    [1, 1, 1], 
    Name='Triangle2',
    Color="(255 0 0)",
    )

raw_input('Press "Enter" create a volume from these polylines using the Connect operation.>')

hfss.connect(oEditor, [objname1, objname2])

raw_input('Press "Enter" to quit HFSS.>')

hfss.quit_application(oDesktop)

del oEditor
del oDesign
del oProject
del oDesktop
del oAnsoftApp
