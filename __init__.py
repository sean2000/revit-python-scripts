from Autodesk.Revit.DB import IFamilyLoadOptions


class FamilyOptions(IFamilyLoadOptions):

    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        return True

    def OnSharedFamilyFound(self, familyInUse, source, overwriteParameterValues):
        return True


