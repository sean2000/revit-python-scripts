"""Functions related to export dwg files, images etc."""

from api import *
from lib import Walker
from lib import get_folder

from Autodesk.Revit.DB import ExportRange
from Autodesk.Revit.DB import FilteredElementCollector
from Autodesk.Revit.DB import FitDirectionType
from Autodesk.Revit.DB import ImageExportOptions
from Autodesk.Revit.DB import ImageFileType
from Autodesk.Revit.DB import ImageResolution
from Autodesk.Revit.DB import View3D
from Autodesk.Revit.DB import ViewFamily
from Autodesk.Revit.DB import ViewFamilyType
from Autodesk.Revit.DB import ZoomFitType


def export_images(doc):
    """Export 3D images from all revit family files found under starting directory"""
    start_dir = get_folder('Enter start directory: ')
    out_dir = get_folder('Enter directory to save images: ')
    last_doc = None

    def cb(filepath):
        current_uidoc = __revit__.OpenAndActivateDocument(filepath)
        current_doc = current_uidoc.Document
        if last_doc:
            last_doc.Close(False)
            print('Closing {}'.format(last_doc))
        last_doc = current_doc

        current_dir, filename = os.path.split(current_doc.PathName)
        basename, ext = os.path.splitext(filename)

        view_type = [x for x in
            FilteredElementCollector(current_doc).OfClass(ViewFamilyType)
            if x.ViewFamily == ViewFamily.ThreeDimensional][0]

        img = ImageExportOptions()
        img.FilePath = os.path.join(out_dir, '{}.png'.format(basename))
        img.HLRandWFViewsFileType = ImageFileType.PNG
        img.ExportRange = ExportRange.CurrentView
        img.ZoomType = ZoomFitType.FitToPage
        img.FitDirection = FitDirectionType.Horizontal
        img.ImageResolution = ImageResolution.DPI_600

        cats = current_doc.Settings.Categories
        view = current_doc.ActiveView
        tr = transaction(current_doc, 'Change view settings')
        tr.Start()
        view = View3D.CreateIsometric(current_doc, view_type.Id)
        current_uidoc.ActiveView = view
        current_doc.ExportImage(img)
        tr.RollBack()

        print(basename)

    w = Walker(start_dir, '.*.rfa', cb)
    w.run()

