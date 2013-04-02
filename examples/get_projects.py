import hycohanz as hfss

raw_input('Press "Enter" to connect to HFSS.>')

[oAnsoftApp, oDesktop] = hfss.setup_interface()

raw_input('Press "Enter" to create a new project.>')

oProject = hfss.new_project(oDesktop)

raw_input('Press "Enter" to create another project.>')

oProject2 = hfss.new_project(oDesktop)

raw_input('Press "Enter" to get a list of open projects.>')

projectlist = hfss.get_projects(oDesktop)

print(projectlist)

raw_input('Press "Enter" to quit HFSS.>')

hfss.quit_application(oDesktop)
