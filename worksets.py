"""Functions related to worksets"""
from Autodesk.Revit.DB import BuiltInParameter
from Autodesk.Revit.DB import FamilyInstance
from Autodesk.Revit.DB import FilteredElementCollector
from Autodesk.Revit.DB import FilteredWorksetCollector
from Autodesk.Revit.DB import WorksetKind

from lib import transaction


def get_hosted_workset(instance):
    return instance.Host.WorksetId


CATEGORY_WORKSETS = {
    'Casework': lambda x: 'A_FFE',
    'Stairs': lambda x: 'A_Core',
    #'Ceilings': lambda x: 'A_Interior',
    'Electrical Equipment': lambda x: 'A_FFE',
    'Electrical Fixtures': lambda x: 'A_FFE',
    'Furniture': lambda x: 'A_FFE',
    'Furniture Systems': lambda x: 'A_FFE',
    'Lighting Fixtures': lambda x: 'A_FFE',
    'Nurse Call Devices': lambda x: 'A_FFE',
    'Floors': lambda x: 'A_Structure',
    'Structural Columns': lambda x: 'A_Structure',
    'Structural Framing': lambda x: 'A_Structure',
    'Windows': get_hosted_workset,
    'Doors': get_hosted_workset,
    'Specialty Equipment': lambda x: 'A_FFE',
    'Roofs': lambda x: 'A_Envelope',
}

STOP_WORKSETS = ['A_StructureEarlyWorks']


def fix_worksets(doc):
    """Apply correct worksets to elements based on their category"""
    workset_table = doc.GetWorksetTable()
    worksets = FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset)
    workset_map = {x.Name: x.Id.IntegerValue for x in worksets}
    elements = FilteredElementCollector(doc).OfClass(FamilyInstance)
    count = 0

    with transaction(doc, 'Fix worksets'):
        for element in elements:
            category = element.Symbol.Category.Name

            ws_getter = CATEGORY_WORKSETS.get(category)
            if not ws_getter:
                continue

            new_workset = ws_getter(element)
            if not new_workset:
                continue

            # new_workset can be either the workset name or the workset id
            if not isinstance(new_workset, str):
                new_workset = workset_table.GetWorkset(new_workset).Name

            workset_id = doc.GetWorksetId(element.Id)
            workset = workset_table.GetWorkset(workset_id).Name

            if workset in STOP_WORKSETS:
                continue

            if workset != new_workset:
                workset_param = element.get_Parameter(
                    BuiltInParameter.ELEM_PARTITION_PARAM)
                if workset_param.IsReadOnly:
                    continue

                workset_param.Set(workset_map.get(new_workset))
                count += 1

            if count and not count % 10:
                print('{} changed'.format(count))

    print('Total of {} elements changed'.format(count))

