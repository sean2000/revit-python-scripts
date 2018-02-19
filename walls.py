"""Functions relating to walls"""
import re

from api import *
from lib import transaction
import constants

from Autodesk.Revit.DB import BuiltInCategory
from Autodesk.Revit.DB import BuiltInParameter
from Autodesk.Revit.DB import CurveElement
from Autodesk.Revit.DB import ElementTypeGroup
from Autodesk.Revit.DB import Family
from Autodesk.Revit.DB import FamilyInstance
from Autodesk.Revit.DB import FamilySymbol
from Autodesk.Revit.DB import FilteredElementCollector
from Autodesk.Revit.DB import Level
from Autodesk.Revit.DB import LinePatternElement
from Autodesk.Revit.DB import TextNote
from Autodesk.Revit.DB import SymbolicCurve
from Autodesk.Revit.DB import View
from Autodesk.Revit.DB import ViewSheet
from Autodesk.Revit.DB import WallType
from Autodesk.Revit.DB import XYZ


def fix_walls(doc):
    """Rename wall types"""
    wall_types = FilteredElementCollector(doc).OfClass(WallType)
    pat = re.compile('A_Wall_Generic Dry Wall (\d{2,3})mm')
    with transaction(doc, 'Fix walls'):
        for wt in wall_types:
            try:
                type_name = wt.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
                width = int(round(wt.Width * constants.MM))
                #acoustic = wt.LookupParameter('AcousticRating_TX').AsString()
                #fire_rating = wt.LookupParameter('FireRating_TX').AsString()
                code = wt.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_MARK).AsString()
                description = wt.get_Parameter(BuiltInParameter.ALL_MODEL_DESCRIPTION).AsString()

                bits = pat.findall(type_name)
                print(bits)
                if bits:
                    wt.Name = 'GENERIC_DW_{}'.format(width, bits[0][1])
                    print(wt.Name)

                #fam.get_Parameter(BuiltInParameter.KEYNOTE_PARAM).Set('')
                #fam.LookupParameter("Acoustic Rating").Set(acoustic)
                #fam.LookupParameter("FireRating_TX").Set(fire_rating)
                #print('{} {}'.format(width, code))
            except Exception as e:
                print(e)


def place_all_walls(doc):
    """Place all walls"""
    levels = list(FilteredElementCollector(doc).OfClass(Level))
    wall_types = sorted(FilteredElementCollector(doc).OfClass(WallType), key=lambda x: x.Width)

    with transaction(doc, 'Place walls'):
        x = 0
        y = 0
        i = 0
        step = 10
        for wt in wall_types:
            if i % 10 == 0:
                x = 0
                y += step
            else:
                x += step
            start = XYZ(x, y, 0)
            end = XYZ(x, y + step, 0)
            line = Line.CreateBound(start, end)
            w = Wall.Create(doc, line, levels[0].Id, False)
            w.WallType = wt
            i += 1


def walls(doc):
    families = FilteredElementCollector(doc).OfClass(WallType)
    for fam in families:
        print '{}\t{}'.format(
            fam.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString(),
            fam.Width * MM)


