# ============================================================================================================
# ============================================================================================================
# ==                                                                                                        ==
# ==                                           CubeGeneration                             ==
# ==                                                                                                        ==
# ============================================================================================================
# ============================================================================================================
# ABOUT
# ============================================================================================================
# version v1.0
# Macro developed for FreeCAD (http://www.freecadweb.org/).
# This macro generates cube shapes using a Spreadsheet.
# 
# Usage: Create a spreadsheet where part labels are in Column A (rows 2 - 100) 
#   For each row add numbers in the appropriate column, for shape to create
#     Column B (width) Column C (length) Column D (height) 
#     Column E (Position X) Column F (Position Y) Column G (Position Z)
#     Column H (Rotation X) Column I (Rotation Y) Column J (Rotation Z)
#
# Note: Before execution, remove cubes which have already been generated, otherwise the new cubes will be generated with a different name
#
# LICENSE
# ============================================================================================================
#
# Copyright (c) 2019 Factorem Pty. Ltd. (Australia)
#
# This work is licensed under GNU Lesser General Public License (LGPL).
# To view a copy of this license, visit https://www.gnu.org/licenses/lgpl-3.0.html.
#
# ============================================================================================================
__title__   = "cubegenerate"
__author__  = "victorromeo"
__version__ = "01.00"
__date__    = "13/11/2019"
 
__Comment__ = "This macro helps managing aliases inside FreeCAD Spreadsheet workbench. It is able to create, delete, move aliases and create a 'part family' group of files"
 
__Status__ = "stable"
__Requires__ = "FreeCAD 0.18"

from PySide import QtGui, QtCore
from FreeCAD import Gui
import os
import string
import traceback

App = FreeCAD
Gui = FreeCADGui

try:

        doc = FreeCAD.ActiveDocument    
        if not doc.FileName:
            FreeCAD.Console.PrintError('\nMust save project first\n')
         
        for i in range(2,100):
            cell_reference = 'A'+ str(i)           
            FreeCAD.Console.PrintMessage("\ncell" + cell_reference + "\n")
            part_name = App.ActiveDocument.Spreadsheet.getContents(cell_reference)
            if part_name == None or part_name == '':
                break

            App.ActiveDocument.Spreadsheet.setAlias('B'+str(i), '')
            App.ActiveDocument.Spreadsheet.setAlias('C'+str(i), '')
            App.ActiveDocument.Spreadsheet.setAlias('D'+str(i), '')
            App.ActiveDocument.Spreadsheet.setAlias('E'+str(i), '')
            App.ActiveDocument.Spreadsheet.setAlias('F'+str(i), '')
            App.ActiveDocument.Spreadsheet.setAlias('G'+str(i), '')
            App.ActiveDocument.Spreadsheet.setAlias('H'+str(i), '')
            App.ActiveDocument.Spreadsheet.setAlias('I'+str(i), '')
            App.ActiveDocument.Spreadsheet.setAlias('J'+str(i), '')
            App.ActiveDocument.recompute()
            App.ActiveDocument.Spreadsheet.setAlias('B'+str(i), part_name + 'Width')
            App.ActiveDocument.Spreadsheet.setAlias('C'+str(i), part_name + 'Length')
            App.ActiveDocument.Spreadsheet.setAlias('D'+str(i), part_name + 'Height')
            App.ActiveDocument.Spreadsheet.setAlias('E'+str(i), part_name + 'PosX')
            App.ActiveDocument.Spreadsheet.setAlias('F'+str(i), part_name + 'PosY')
            App.ActiveDocument.Spreadsheet.setAlias('G'+str(i), part_name + 'PosZ')
            App.ActiveDocument.Spreadsheet.setAlias('H'+str(i), part_name + 'RotX')
            App.ActiveDocument.Spreadsheet.setAlias('I'+str(i), part_name + 'RotY')
            App.ActiveDocument.Spreadsheet.setAlias('J'+str(i), part_name + 'RotZ')
            App.ActiveDocument.recompute()

            o = App.ActiveDocument.addObject("Part::Box","Box")
            o.Label = part_name
            o.setExpression('Width',u'Shapes.' + part_name+'Width')
            o.setExpression('Length',u'Shapes.' + part_name+'Length')
            o.setExpression('Height',u'Shapes.' + part_name+'Height')
            o.setExpression('Placement.Base.x',u'Shapes.' + part_name+'PosX')
            o.setExpression('Placement.Base.y',u'Shapes.' + part_name+'PosY')
            o.setExpression('Placement.Base.z',u'Shapes.' + part_name+'PosZ')

        App.ActiveDocument.recompute()
        FreeCAD.Console.PrintMessage("\nPart family files generated\n")

except Exception as e:
    FreeCAD.Console.PrintError("\nUnable to complete task\n")
    FreeCAD.Console.PrintError(e)

    traceback.print_exc()
 
