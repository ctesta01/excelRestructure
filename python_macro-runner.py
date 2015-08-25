from __future__ import print_function
import unittest
import os.path
import win32com.client

"""
ExcelMacro
August 19th, 2015

ExcelMacro works by running a macro defined in Macros.xlsm.
The macro is projectsRestructure.
projectsRestructure runs on OIREProjects.csv.

This macro outputs a file called OIREProjects_Restructured_ddmmyy.xlsx,
where ddmmyy is replaced with the date.

If there are no other Excel workbooks open, then the macro will close Excel.
If there are, it will only close Macros.xlsm, OIREProjects.csv,
and OIREProjects_Restructured_ddmmyy.xlsx.

This is currently setup to run where python_macro-runner.py, OIREProjects.csv,
and Macros.xlsm are all in the same directory.

If you would like to change that, change macrosPath and projectsPath to the
correct directories for each respective file (wrapped in quotes), and the
filenames (set by macrosName and projectsName) as well, if necessary.

Ex:
macrosName = 'NewMacros.xlsm'
projectsName = 'OIREProjects2.csv'

macrosPath = "C:\Some\Directory\Goes\Here"
projectPath = "C:\You\Can\Use\A\Different\Directory\Here"
"""

class ExcelMacro(unittest.TestCase):
    def test_excel_macro(self):
        macrosName = 'Macros.xlsm'
        projectsName = 'OIREProjects.csv'
        # OIREProjects.csv is assumed as the name of the csv in the macro,
        # you should not change this unless you also change it in the macro.

        macrosPath = os.getcwd()
        projectsPath = os.getcwd()

        
        xlApp = win32com.client.DispatchEx('Excel.Application')
        xlsPath = os.path.join(macrosPath, macrosName)
        wb = xlApp.Workbooks.Open(Filename=xlsPath)
        xlsPath = os.path.join(projectsPath, projectsName)
        wb = xlApp.Workbooks.Open(Filename=xlsPath)

        xlApp.Run('Macros.xlsm!projectsRestructure')

        if xlApp.Workbooks.Count == 0:
            xlApp.Quit
        
        print("Macro ran successfully!")
if __name__ == "__main__":
    unittest.main()
