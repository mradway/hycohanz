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

hfss.create_rectangle(oEditor,  1, 
                                2, 
                                3, 
                                4,
                                5,
                                Name='Rectangle1',
                                Color=(255, 0, 0))

raw_input('Press "Enter" to quit HFSS.>')

hfss.quit_application(oDesktop)
