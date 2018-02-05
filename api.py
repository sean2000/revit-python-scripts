"""
Add dll references and define APP, UIDOC and DOC constants
"""
import clr
#clr.AddReferenceToFileAndPath(r"C:\Program Files\Autodesk\Revit 2016\RevitAPI.dll")
#clr.AddReferenceToFileAndPath(r"C:\Program Files\Autodesk\Revit 201\RevitAPIUI.dll")
#clr.AddReferenceToFileAndPath(r"C:\Users\DodswoS\Dropbox\work\jacobs\RevitTools\packages\Moq.4.7.99\lib\net45\Moq.dll")

UIAPP = __revit__
APP = __revit__.Application
UIDOC = __revit__.ActiveUIDocument
DOC = UIDOC.Document



