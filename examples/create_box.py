import hycohanz as hfss

raw_input('Press "Enter" to connect to HFSS.>')

with hfss.App() as App:

    raw_input('Press "Enter" to create a new project.>')
    with hfss.NewProject(App.oDesktop) as P:

        raw_input('Press "Enter" to insert a new DrivenModal design named HFSSDesign1.>')
        with hfss.InsertDesign(P.oProject, "HFSSDesign1", "DrivenModal") as D:

            raw_input('Press "Enter" to set the active editor to "3D Modeler" (The default and only known correct value).>')
            with hfss.SetActiveEditor(D.oDesign) as E:

                raw_input('Press "Enter" to draw a red box named Box1.>')

                hfss.create_box(E.oEditor,  
                                1, 
                                2, 
                                3, 
                                4,
                                5,
                                Name='Box1',
                                Color=(255, 0, 0))

                raw_input('Press "Enter" to release HFSS.>')
