import hycohanz as hfss

raw_input('Press "Enter" to connect to HFSS.>')

with hfss.App() as App:

    raw_input('Press "Enter" to create a new project.>')
    
    with hfss.NewProject(App.oDesktop) as P:
    
        raw_input('Press "Enter" to insert a new DrivenModal design named HFSSDesign1.>')
        
        with hfss.InsertDesign(P.oProject, "HFSSDesign1", "DrivenModal") as D:
        
            raw_input('Press "Enter" to quit HFSS.>')
