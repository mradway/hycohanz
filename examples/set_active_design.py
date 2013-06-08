import hycohanz as hfss
import os.path

raw_input('Press "Enter" to connect to HFSS.>')

[oAnsoftApp, oDesktop] = hfss.setup_interface()

raw_input('Press "Enter" to open an example project.>')

filepath = os.path.join(os.path.abspath(os.path.curdir), 'WR284.hfss')

oProject = hfss.open_project(oDesktop, filepath)

raw_input('Press "Enter" to set the active design to HFSSDesign1.>')

oDesign = hfss.set_active_design(oProject, 'HFSSDesign1')

raw_input('Press "Enter" to close the example project.>')

hfss.close_project_byhandle(oDesktop, oProject)

raw_input('Press "Enter" to quit HFSS.>')

hfss.quit_application(oDesktop)

del oProject
del oDesktop
del oAnsoftApp
