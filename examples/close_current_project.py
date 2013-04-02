import hycohanz as hfss

raw_input('Press "Enter" to connect to HFSS.>')

[oAnsoftApp, oDesktop] = hfss.setup_interface()

raw_input('Press "Enter" to create a new project.>')

oProject = hfss.new_project(oDesktop)

raw_input('Press "Enter" to create another project.>')

oProject2 = hfss.new_project(oDesktop)

raw_input('Press "Enter" to close the current project.>')

hfss.close_current_project(oDesktop)

raw_input('Press "Enter" to quit HFSS.>')

hfss.quit_application(oDesktop)
