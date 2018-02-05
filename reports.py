import os
import csv

from Autodesk.Revit.DB import BuiltInCategory
from Autodesk.Revit.DB import BuiltInParameter
from Autodesk.Revit.DB import ElementId
from Autodesk.Revit.DB import ElementMulticategoryFilter
from Autodesk.Revit.DB import ElementTypeGroup
from Autodesk.Revit.DB import Family
from Autodesk.Revit.DB import FamilyInstance
from Autodesk.Revit.DB import FamilySymbol
from Autodesk.Revit.DB import FilteredElementCollector
from Autodesk.Revit.DB import Level
from Autodesk.Revit.DB import LinePatternElement
from Autodesk.Revit.DB import ParameterFilterElement
from Autodesk.Revit.DB import TextNote
from Autodesk.Revit.DB import View
from Autodesk.Revit.DB import ViewType
from Autodesk.Revit.DB import ViewSheet
from Autodesk.Revit.DB import XYZ
from System.Collections.Generic import List

import lib
from api import *
import constants


def families_report(doc):
    """Generate report of families"""
    with open(os.path.join(constants.HOMEDIR, '{} - families.csv'.format(doc.Title)), 'w') as f:
        w = csv.writer(f, lineterminator='\n')
        w.writerow(('Category', 'Family'))
        for family in FilteredElementCollector(doc).OfClass(Family):
            try:
                w.writerow((family.FamilyCategory.Name, family.Name))
            except:
                print(family)


def view_template_report(doc):
    """Generate csv report of View Templates and their Category visibility"""

    views = FilteredElementCollector(doc).OfClass(View)
    rows = []
    templates = set()

    for view in views:
        if view.ViewTemplateId != ElementId.InvalidElementId:
            templates.add(view.ViewTemplateId)

    for template_id in templates:
        template = doc.GetElement(template_id)
        print(template.Name)
        for cat in doc.Settings.Categories:
            try:
                rows.append((
                    template.Name,
                    str(cat.CategoryType).split('.')[-1],
                    cat.Name,
                    cat.get_Visible(template)
                ))
            except:
                pass

    csv_name = os.path.join(constants.HOMEDIR, '{} - View Templates.csv'.format(doc.Title))
    with open(csv_name, 'w') as f:
        w = csv.writer(f, lineterminator='\n')
        w.writerow(('View Template', 'Type', 'Category', 'Visible'))
        w.writerows(rows)


def sheet_report(doc):
    """Report on sheets"""
    sheets = [x for x in FilteredElementCollector(doc).OfClass(ViewSheet)
            if x.LookupParameter('Identifier_TX').AsString() == 'RLS']

    with lib.transaction(doc, 'Add sheets to schedule'):
        for sheet in sheets:
            scheduled_param = sheet.get_Parameter(BuiltInParameter.SHEET_SCHEDULED)
            if not scheduled_param.AsInteger():
                print(sheet.SheetNumber)
                scheduled_param.Set(1)


def get_door_window_walls(doc):
    """
    Report all doors or windows with hosted wall types
    """

    categories = List[BuiltInCategory]()
    categories.Add(BuiltInCategory.OST_Doors)
    categories.Add(BuiltInCategory.OST_Windows)
    categories_filter = ElementMulticategoryFilter(categories)

    elements = FilteredElementCollector(doc).OfClass(FamilyInstance).WherePasses(categories_filter)
    level_map = {}
    rows = []
    for element in elements:
        wall = element.Host
        if not wall:
            continue
        if not wall.WallType:
            continue

        level_id = element.LevelId
        level = level_map.get(level_id)
        if not level:
            level = doc.GetElement(level_id)
            level_map[level_id] = level

        rows.append([
            element.Category.Name,
            level.Name,
            element.Symbol.FamilyName,
            element.Name,
            element.Symbol.LookupParameter('Descriptor_TX').AsString() or '',
            element.get_Parameter(BuiltInParameter.ALL_MODEL_MARK).AsString(),
            wall.WallType.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString(),
            wall.WallType.LookupParameter('Wall_Mark_TX').AsString() or '',
            wall.WallType.Width * constants.MM,
        ])

    csv_name = os.path.join(constants.HOMEDIR, '{} - DoorsWindows.csv'.format(doc.Title))

    with open(csv_name, 'w') as f:
        writer = csv.writer(f, lineterminator='\n')
        writer.writerow([
            'Category',
            'Level',
            'Family',
            'Type',
            'Code',
            'Mark',
            'Wall Type',
            'Wall Code',
            'Wall Width',
        ])
        writer.writerows(rows)


def export_rooms_csv(doc):
    rooms = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms)
    with open('rooms.csv', 'w') as f:
        w = csv.writer(f)
        w.writerow(['ID', 'Number', 'Name', 'Level'])
        for room in rooms:
            w.writerow([
                room.Id,
                room.Number,
                room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString(),
                room.Level.Name])
    print 'Done'


def import_rooms_csv(doc):
    rooms = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms)
    id_map = {}
    csvpath = r'I:\NBIF\Projects\NB98100\Deliverables\BIM\01_Architecture\A_WIP Models\rooms_updated.csv'
    with open(csvpath) as f:
        for row in csv.DictReader(f, ):
            id_map[row['ID']] = row

    tr = Transaction(doc, "Change room numbers")
    tr.Start()

    with open('rooms_changed.csv', 'w') as f:
        w = csv.writer(f)
        w.writerow(['ID', 'Name', 'Old Number', 'New Number', 'Level'])
        for room in rooms:
            id_row = id_map.get(str(room.Id))
            if id_row:
                assert id_row['Level'] not in ['Level 1', 'Level 4', 'Level 5']
                old_number = room.Number
                room.Number = id_row['Number']
                print '{} - {} -> {}'.format(room.Id, old_number, room.Number)
                w.writerow([
                    room.Id,
                    room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString(),
                    old_number,
                    room.Number,
                    room.Level.Name])

    tr.Commit()


def list_doors():
    doors = FilteredElementCollector(doc)\
        .OfClass(FamilyInstance)\
        .OfCategory(BuiltInCategory.OST_Doors)
    print 'ID           Mark    ToRoom  FromRoom'
    print '--           ------- ------- --------'
    for door in doors:
        phase = doc.GetElement(door.CreatedPhaseId)
        try:
            to_room = door.ToRoom[phase]
            from_room = door.FromRoom[phase]
            door_mark = door.get_Parameter(BuiltInParameter.ALL_MODEL_MARK).AsString()
            to_room_number = None
            from_room_number = None
            if to_room:
                to_room_number = to_room.Number
            if from_room:
                from_room_number = from_room.Number
            print '{}\t{}\t{}\t{}'.format(door.Id, door_mark[:7], to_room_number, from_room_number)
        except:
            pass
    print


def list_text():
    elements = FilteredElementCollector(doc).OfClass(TextElement)
    for x in elements:
        print x.Text


def export_families(doc):
    """Generate report of loaded families and types"""

    symbols = FilteredElementCollector(doc).OfClass(FamilySymbol)
    with open(os.path.join(constants.HOMEDIR, 'report.csv'), 'w') as f:
        w = csv.writer(f, lineterminator='\n')
        w.writerow(['Category', 'Family', 'Type', 'Descriptor_TX', 'BIM ID'])
        for symbol in symbols:
            name = symbol.get_Parameter(
                BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
            family = symbol.Family.Name
            category = symbol.Category.Name
            dtx_param = symbol.LookupParameter('Descriptor_TX')
            dtx = ''
            if dtx_param:
                dtx = dtx_param.AsString()
            bim_id_param = symbol.LookupParameter('BIM ID')
            if bim_id_param:
                bim_id = bim_id_param.AsString()
            else:
                bim_id = ''
            w.writerow([category, family, name, dtx, bim_id])


def parameter_report(doc):
    iterator = doc.ParameterBindings.ForwardIterator()
    while iterator.MoveNext():
        definition = iterator.Key
        binding = iterator.Current

        print definition.Name, definition.ParameterType, definition.ParameterGroup, definition.UnitType
        #for category in binding.Categories:
        #    print category.Name


def offsets_report(doc):
    out = []
    for instance in FilteredElementCollector(doc).OfClass(FamilyInstance):
        offset = instance.get_Parameter(
            BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM
        ).AsDouble() * constants.MM
        #if offset < 0:
        out.append((
            instance.Id.IntegerValue,
            instance.Symbol.Category.Name,
            instance.Symbol.Family.Name,
            instance.Symbol.get_Parameter(
                BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString(),
            doc.GetElement(instance.LevelId).Name if instance.LevelId != ElementId.InvalidElementId else '',
            offset,
        ))
    with open(os.path.join(HOMEDIR, 'offsets.csv'), 'w') as f:
        w = csv.writer(f, lineterminator='\n')
        w.writerow(['ID', 'Category', 'Family', 'Type', 'Level', 'Offset'])
        w.writerows(out)


def view_filter_report(doc):
    """
    Generate a csv of the form [View, Template Name, View Filter Name] for all
    views that support filters
    """

    view_types = [
        ViewType.AreaPlan,
        ViewType.CeilingPlan,
        ViewType.Elevation,
        ViewType.FloorPlan,
        ViewType.Section,
        ViewType.ThreeD,
    ]

    with open(os.path.join(constants.HOMEDIR, 'filters.csv'), 'w') as f:
        w = csv.writer(f, lineterminator='\n')
        w.writerow(['View', 'Template', 'Filter'])

        for flt in FilteredElementCollector(doc).OfClass(ParameterFilterElement):
            w.writerow(['', '', flt.Name])

        for view in FilteredElementCollector(doc).OfClass(View):
            if view.ViewType not in view_types:
                continue

            if view.ViewTemplateId != ElementId.InvalidElementId:
                template_name = doc.GetElement(view.ViewTemplateId).Name
            else:
                template_name = ''

            try:
                for fid in view.GetFilters():
                    fl = doc.GetElement(fid)
                    w.writerow([view.Name, template_name, fl.Name])
            except:
                print(view.Name, view.ViewType)


def views_report(doc):

    with open(os.path.join(HOMEDIR, 'views_on_sheets.csv'), 'w') as f:
        w = csv.writer(f, lineterminator='\n')
        w.writerow(('Sheet Number', 'Sheet Name', 'View Name', 'View Type'))
        for sheet in FilteredElementCollector(doc).OfClass(ViewSheet):
            for view in FilteredElementCollector(doc, sheet.Id).OfClass(View):
                w.writerow((sheet.SheetNumber, sheet.Name, view.Name, view.ViewType))


