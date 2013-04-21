import hycohanz as hfss

raw_input('Press "Enter" to connect to HFSS.>')

[oAnsoftApp, oDesktop] = hfss.setup_interface()

raw_input('Press "Enter" to create a new project.>')

oProject = hfss.new_project(oDesktop)

raw_input('Create a project in the HFSS window and press "Enter">')

oProject2 = hfss.get_active_project(oDesktop)
project2name = hfss.get_project_name(oProject2)

print("You created a project with name {0}".format(project2name))

raw_input('Press "Enter" to close all projects.>')

hfss.close_all_projects(oDesktop)

raw_input('Press "Enter" to quit HFSS.>')

hfss.quit_application(oDesktop)

del oEditor
del oDesign
del oProject
del oDesktop
del oAnsoftApp
