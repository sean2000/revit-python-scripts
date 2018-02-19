import os
import re
from contextlib import contextmanager
import tempfile

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference("System.Windows.Forms")

import System
from Autodesk.Revit.DB import BuiltInCategory
from Autodesk.Revit.DB import BuiltInParameter
from Autodesk.Revit.DB import DisplayUnitType
from Autodesk.Revit.DB import FamilyInstance
from Autodesk.Revit.DB import FamilySymbol
from Autodesk.Revit.DB import Family
from Autodesk.Revit.DB import FilteredElementCollector
from Autodesk.Revit.DB import ParameterFilterElement
from Autodesk.Revit.DB import ParameterType
from Autodesk.Revit.DB import StorageType
from Autodesk.Revit.DB import Transaction
from Autodesk.Revit.DB import View
from Autodesk.Revit.DB import ViewFamilyType
from Autodesk.Revit.DB import ViewSheet
from Autodesk.Revit.DB import ViewType

from System.Windows.Forms import FolderBrowserDialog
from System.Windows.Forms import DialogResult


def get_temp_path(filename):
    return os.path.join(tempfile.gettempdir(), filename)


def get_folder(prompt):
    dlg = FolderBrowserDialog()
    dlg.Description = prompt
    res = dlg.ShowDialog()
    if res == DialogResult.OK:
        return dlg.SelectedPath


def load_assembly(path):
    assembly_bytes = System.IO.File.ReadAllBytes(path)
    pdb_path = path[:-3] + 'pdb'
    if os.path.exists(pdb_path):
        pdb_bytes = System.IO.File.ReadAllBytes(pdb_path)
        assembly = System.Reflection.Assembly.Load(assembly_bytes, pdb_bytes)
    else:
        assembly = System.Reflection.Assembly.Load(assembly_bytes)
    folder = os.path.dirname(path)
    System.AppDomain.CurrentDomain.AssemblyResolve += resolve_assembly_generator(folder)
    return assembly


def resolve_assembly_generator(folder):
    def result(sender, args):
        name = args.Name.split(',')[0]
        try:
            path = os.path.join(folder, name + '.dll')
            if not os.path.exists(path):
                return None
            return load_assembly(path)
        except:
            import traceback
            traceback.print_exc()
            return None
    return result


def cd2pd(cd):
    """c# date to python date"""
    return datetime(cd.Year, cd.Month, cd.Day, cd.Hour, cd.Minute, cd.Second)


def pd2cd(pd):
    """python date to c# date"""
    return System.DateTime(
        pd.year,
        pd.month,
        pd.day,
        pd.hour,
        pd.minute,
        pd.second,
        System.DateTimeKind.Local)


@contextmanager
def transaction(doc, message):
    tr = Transaction(doc, message)
    tr.Start()
    yield
    tr.Commit()


@contextmanager
def rollback_transaction_group(doc, message):
    tg = TransactionGroup(doc, message)
    tg.Start()
    yield
    tg.RollBack()


class DocumentWrapper(object):
    """
    Class that wraps Revit Document and provides easy access to commonly used
    items such as walls, doors, views etc
    """

    def __init__(self, doc):
        self._doc = doc

    @property
    def _collector(self):
        return FilteredElementCollector(self._doc)

    @property
    def sheets(self):
        return list(self._collector.OfClass(ViewSheet))

    @property
    def views(self):
        return list(self._collector.OfClass(View))

    @property
    def view_templates(self):
        return [x for x in self.views if x.IsTemplate]

    @property
    def view_filters(self):
        return list(self._collector.OfClass(ParameterFilterElement))

    @property
    def view_plans(self):
        return [x for x in self.views if x.ViewType == ViewType.FloorPlan]

    @property
    def doors(self):
        return list(self._collector.OfClass(
            FamilyInstance).OfCategory(BuiltInCategory.OST_Doors))

    @property
    def rooms(self):
        return list(self._collector.OfCategory(BuiltInCategory.OST_Rooms))

    @property
    def view_types(self):
        return list(self._collector.OfClass(ViewFamilyType))

    @property
    def title_blocks(self):
        return list(self._collector.OfClass(
            FamilySymbol).OfCategory(BuiltInCategory.OST_TitleBlocks))

    @property
    def symbols(self):
        return list(self._collector.OfClass(FamilySymbol))

    @property
    def families(self):
        return list(self._collector.OfClass(Family))

    @property
    def instances(self):
        return list(self._collector.OfClass(FamilyInstance))

    @property
    def scope_boxes(self):
        return list(self._collector.OfCategory(
            BuiltInCategory.OST_VolumeOfInterest))

    @property
    def unused_scope_boxes(self):
        all_scope_box_ids = set(x.Id for x in self.scope_boxes)
        used_scope_box_ids = set()
        for view in self.views:
            try:
                used_scope_box_ids.add(
                    view.get_Parameter(
                        BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP
                    ).AsElementId())
            except:
                pass
        unused_scope_box_ids = all_scope_box_ids.difference(used_scope_box_ids)
        unused_scope_boxes = []
        for id in unused_scope_box_ids:
            unused_scope_boxes.append(this._doc.GetElement(id))
        return unused_scope_boxes


class Walker(object):
    """
    Walk file system from start_dir and run callback function on any files that
    match file_regex
    """
    def __init__(self, start_dir, file_regex, callback):
        self.start_dir = start_dir
        self.file_regex = re.compile(file_regex)
        self.callback = callback

    def run(self):
        for root, dirs, filenames in os.walk(self.start_dir):
            for filename in filenames:
                if not self.file_regex.match(filename):
                    continue
                filepath = os.path.join(root, filename)
                self.callback(filepath)


def get_param_value(param, doc):
    if not param:
        return
    try:
        if param.DisplayUnitType == DisplayUnitType.DUT_MILLIMETERS:
            units = MM
        elif param.DisplayUnitType == DisplayUnitType.DUT_METERS:
            units = MM/1000
        else:
            units = 1
    except:
        units = 1
    if param.StorageType == StorageType.ElementId:
        id = param.AsElementId()
        #if id.IntegerValue >= 0:
        #    return doc.GetElement(id).Name
        return id.IntegerValue
    elif param.StorageType == StorageType.Double:
        return param.AsDouble() * units
    elif param.StorageType == StorageType.Integer:
        value = param.AsInteger()
        if param.Definition.ParameterType == ParameterType.YesNo:
            if value == 0:
                return False
            else:
                return True
        else:
            return value * units
    else:
        try:
            return param.AsString().encode('utf-8')
        except:
            return param.AsString()

