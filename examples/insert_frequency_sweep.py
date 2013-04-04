import hycohanz as hfss

raw_input('Press "Enter" to connect to HFSS.>')

[oAnsoftApp, oDesktop] = hfss.setup_interface()

raw_input('Press "Enter" to create a new project.>')

oProject = hfss.new_project(oDesktop)

raw_input('Press "Enter" to insert a new DrivenModal design named HFSSDesign1.>')

oDesign = hfss.insert_design(oProject, "HFSSDesign1", "DrivenModal")

setupname = hfss.insert_analysis_setup(oDesign, 1e9)

hfss.insert_frequency_sweep(oDesign,
                            setupname,
                            "Sweep1",
                            1e9,
                            2e9,
                            0.1e9,
                            IsEnabled=True,
                            SetupType="LinearStep",
                            Type="Discrete",
                            SaveFields=True,
                            ExtrapToDC=False)

raw_input('Press "Enter" to quit HFSS.>')

hfss.quit_application(oDesktop)
