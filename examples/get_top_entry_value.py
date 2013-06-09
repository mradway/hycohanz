import hycohanz as hfss
import os.path

raw_input('Press "Enter" to connect to HFSS.>')

[oAnsoftApp, oDesktop] = hfss.setup_interface()

raw_input('Press "Enter" to open an example project.>')

filepath = os.path.join(os.path.abspath(os.path.curdir), 'WR284.hfss')

oProject = hfss.open_project(oDesktop, filepath)

raw_input('Press "Enter" to set the active design to HFSSDesign1.>')

oDesign = hfss.set_active_design(oProject, 'HFSSDesign1')

raw_input('Press "Enter" to get a handle to the Fields Reporter module.>')

oFieldsReporter = hfss.get_module(oDesign, 'FieldsReporter')

raw_input('Press "Enter" to enter the field quantity "E" in the Fields Calculator.>')

hfss.enter_qty(oFieldsReporter, 'E')

raw_input('Open the calculator to verify that the Calculator input reads "CVc : <Ex,Ey,Ez>">')

hfss.calc_op(oFieldsReporter, 'Mag')

raw_input('Close and reopen Calculator to verify that the "Mag" function has been applied.>')

hfss.enter_vol(oFieldsReporter, 'Polyline1')

raw_input('Close and reopen Calculator to verify that the evaluation volume is entered.>')

hfss.calc_op(oFieldsReporter, 'Maximum')

raw_input('Close and reopen Calculator to verify that the "Maximum" function is entered.>')

hfss.clc_eval(
    oFieldsReporter, 
    'Setup1', 
    'LastAdaptive', 
    3.95e9, 
    0.0, 
    {},
    )

raw_input('Verify that the correct answer is at the top of the stack.>')

result = hfss.get_top_entry_value(
    oFieldsReporter, 
    'Setup1', 
    'LastAdaptive', 
    3.95e9, 
    0.0, 
    {},
    )

print('result: ' + str(result))

raw_input('Press "Enter" to close the example project.>')

hfss.close_project_byhandle(oDesktop, oProject)

raw_input('Press "Enter" to quit HFSS.>')

hfss.quit_application(oDesktop)

del oDesign
del oProject
del oDesktop
del oAnsoftApp
