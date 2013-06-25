import hycohanz as hfss

raw_input('Press "Enter" to connect to HFSS.>')
with hfss.App() as App:
    raw_input('Press "Enter" to create a new project.>')
    with hfss.NewProject(App.oDesktop) as P:
        raw_input('Press "Enter" to insert a new design.>')
        with hfss.InsertDesign(P.oProject, "HFSSDesign1", "DrivenModal") as D:
            raw_input('Press "Enter" to set the active editor.>')
            with hfss.SetActiveEditor(D.oDesign) as E:
                print("Switch to HFSS and draw some random objects in the editor.")
                            
                print("When done, select some of them.")
                
                raw_input('Press "Enter" to print a list of the selected objects.>')
                
                objlist = hfss.get_selections(E.oEditor)
                
                print(objlist)
                
                raw_input('Press "Enter" to quit HFSS.>')
