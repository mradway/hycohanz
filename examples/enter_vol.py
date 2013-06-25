"""
Demonstrate usage of the enter_vol() function.
"""
import hycohanz as hfss
import os.path

raw_input('Press "Enter" to connect to HFSS.>')
with hfss.App() as App:
    raw_input('Press "Enter" to open an example project.>')
    
    filepath = os.path.join(os.path.abspath(os.path.curdir), 'WR284.hfss')
    with hfss.OpenProject(App.oDesktop, filepath) as P:
        raw_input('Press "Enter" to set the active design to HFSSDesign1.>')
        
        with hfss.SetActiveDesign(P.oProject, 'HFSSDesign1') as D:
            raw_input('Press "Enter" to get a handle to the Fields Reporter module.>')
            
            with hfss.GetModule(D.oDesign, 'FieldsReporter') as FR:
                raw_input(
                    'Press "Enter" to enter the field quantity "E" in the Fields Calculator.>')
                
                hfss.enter_qty(FR.oModule, 'E')
                
                raw_input(
                    """
                    Open the calculator to verify that the 
                    Calculator input reads "CVc : <Ex,Ey,Ez>">
                    """)
                
                hfss.calc_op(FR.oModule, 'Mag')
                
                raw_input(
                    """
                    Close and reopen Calculator to verify that 
                    the "Mag" function has been applied.>
                    """)
                
                hfss.enter_vol(FR.oModule, 'Polyline1')
                
                raw_input(
                    """
                    Close and reopen Calculator to verify 
                    that the evaluation volume is entered.>
                    """)
                
                hfss.calc_op(FR.oModule, 'Maximum')
                
                raw_input(
                    """
                    Close and reopen Calculator to verify 
                    that the "Maximum" function is entered.>
                    """)
                
                hfss.clc_eval(
                    FR.oModule, 
                    'Setup1', 
                    'LastAdaptive', 
                    3.95e9, 
                    0.0, 
                    {},
                    )
                
                raw_input('Verify that the correct answer is at the top of the stack.>')
                
                result = hfss.get_top_entry_value(
                    FR.oModule, 
                    'Setup1', 
                    'LastAdaptive', 
                    3.95e9, 
                    0.0, 
                    {},
                    )
                
                print('result: ' + str(result))
                
                raw_input('Press "Enter" to quit HFSS.>')
