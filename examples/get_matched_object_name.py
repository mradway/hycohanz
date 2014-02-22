import hycohanz as hfss

[oAnsoftApp, oDesktop] = hfss.setup_interface()

oProject = hfss.new_project(oDesktop)
oDesign = hfss.insert_design(oProject, "HFSSDesign1", "DrivenModal")
oEditor = hfss.set_active_editor(oDesign)

objname1 = hfss.create_polyline(
    oEditor, [1, 0], 
    [0, 0],
    [0, 0],
    Name='Path',
    Color="(255 0 0)",
    IsPolylineCovered=False,
    IsPolylineClosed=False)

objname2 = hfss.create_circle(oEditor, 0, 0, 0, 4, WhichAxis='X')

objname3 = hfss.create_cylinder(oEditor, 0, 0, 0, 10, 30, WhichAxis='Z')

objname4 = hfss.create_sphere(oEditor, 0, 0, 0, 3)

hfss.sweep_along_path(oEditor,[objname2, objname1])

hfss.subtract(oEditor, [objname3], [objname4], KeepOriginals=False)

merge1 = hfss.get_matched_object_name(oEditor, "Cyl*")
merge2 = hfss.get_matched_object_name(oEditor, "Cir*")
hfss.unite(oEditor, merge1 + merge2)

raw_input('Press "Enter" to quit HFSS.>')
hfss.close_current_project(oDesktop)

hfss.quit_application(oDesktop)

del oEditor
del oDesign
del oProject
del oDesktop
del oAnsoftApp