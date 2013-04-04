import hycohanz as hfss

raw_input('Press "Enter" to connect to HFSS.>')

[oAnsoftApp, oDesktop] = hfss.setup_interface()

raw_input('Press "Enter" to create a new project.>')

oProject = hfss.new_project(oDesktop)

raw_input('Press "Enter" to insert a new DrivenModal design named HFSSDesign1.>')

oDesign = hfss.insert_design(oProject, "HFSSDesign1", "DrivenModal")

oBoundarySetup = hfss.get_module(oDesign, "BoundarySetup")
oMeshSetup = hfss.get_module(oDesign, "MeshSetup")
oAnalysisSetup = hfss.get_module(oDesign, "AnalysisSetup")
oOptimetrics = hfss.get_module(oDesign, "Optimetrics")
oSolutions = hfss.get_module(oDesign, "Solutions")
oFieldsReporter = hfss.get_module(oDesign, "FieldsReporter")
oRadField = hfss.get_module(oDesign, "RadField")
oUserDefinedSolutionModule = hfss.get_module(oDesign, "UserDefinedSolutionModule")

raw_input('Press "Enter" to quit HFSS.>')

hfss.quit_application(oDesktop)
