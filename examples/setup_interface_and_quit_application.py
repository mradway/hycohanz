import hycohanz as hfss

[oAnsoftApp, oDesktop] = hfss.setup_interface()

raw_input('Press "Enter" to quit HFSS.>')

hfss.quit_application(oDesktop)

del oDesktop
del oAnsoftApp
