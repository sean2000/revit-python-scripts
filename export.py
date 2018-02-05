from api import *


def export_view_image(doc):
    out_dir = 'c:\\Users\\DodswoS'
    last_doc = None

    def cb(filepath):
        current_uidoc = __revit__.OpenAndActivateDocument(filepath)
        current_doc = current_uidoc.Document
        #if last_doc:
        #    last_doc.Close(False)
        #    print 'Closing {}'.format(last_doc)

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
        with transaction(current_doc, 'Change view settings'):
            view = View3D.CreateIsometric(current_doc, view_type.Id)
            current_uidoc.ActiveView = view
            #view.SetVisibility(cats.get_Item(BuiltInCategory.OST_Dimensions), False)
            #view.SetVisibility(cats.get_Item(BuiltInCategory.OST_TextNotes), False)
            current_doc.ExportImage(img)
        print basename

    w = Walker(r'C:\Users\DodswoS\NBIF', '.*.rfa', cb)
    w.run()



