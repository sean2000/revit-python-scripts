import re

from Autodesk.Revit.DB import FilteredElementCollector
from Autodesk.Revit.DB import ParameterFilterElement
from Autodesk.Revit.DB import Transaction
from Autodesk.Revit.DB import View
from Autodesk.Revit.DB import ViewType

from lib import transaction
from api import *


def set_has_layout(doc):
    pat = re.compile('^(\d{4}) - ')
    plans = (x for x in FilteredElementCollector(doc).OfClass(View) if x.ViewType == ViewType.FloorPlan)
    rooms = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms)
    room_map = {}
    for room in rooms:
        room_map[room.Number] = room

    tr = Transaction(doc, 'Set HasLayout')
    tr.Start()
    for plan in plans:
        bits = pat.findall(plan.Name)
        if len(bits) == 1:
            room_num = bits[0]
            room = room_map.get(room_num)
            if room:
                room.LookupParameter('HasLayout_YN').Set(True)
                print(room)
    tr.Commit()


def fix_view_templates(doc):
    dw = DocumentWrapper(doc)
    tr = Transaction(doc, 'Fix views templates')
    tr.Start()
    for vt in dw.view_templates:
        param = vt.LookupParameter(constants.IDENTIFIER_PARAM)
        #if param.AsString() == 'PLOT_':
        if param:
            param.Set('PLOT')
            print(vt.Name)
    tr.Commit()


def rename_filters(doc):
    """Rename View Filters with '(1)' at the end of the name"""
    pat = re.compile('(.*)\(\d\)$')
    filters = FilteredElementCollector(doc).OfClass(ParameterFilterElement)
    with transaction(doc, 'Rename filters'):
        for f in filters:
            bits = pat.findall(f.Name)
            if bits:
                f.Name = bits[0].strip()
                print(f.Name)


def view_params(doc):
    view = doc.ActiveView
    fields = {}
    for param in view.Parameters:
        fields[param.Definition.Name] = get_param_value(param, doc)
    pprint(fields)


def swap_view_templates(doc, from_name, to_name):

    all_views = [x for x in FilteredElementCollector(doc).OfClass(View) if not x.IsTemplate]
    all_templates = [x for x in FilteredElementCollector(doc).OfClass(View) if x.IsTemplate]
    from_template = [x for x in all_templates if x.Name == from_name][0]
    to_template = [x for x in all_templates if x.Name == to_name][0]

    with transaction(doc, "Swap view templates"):
        for view in all_views:
            if view.ViewTemplateId == from_template.Id:
                view.ViewTemplateId = to_template.Id
                print(view.Name)


def fix_schedules(doc):
    pat = re.compile('(\d{4})')
    schedules = [x for x in FilteredElementCollector(doc).OfClass(View)
        if x.ViewType == ViewType.Schedule and 'fitout' in x.Name.lower()]
    levels = list(FilteredElementCollector(doc).OfClass(Level))
    with transaction(doc, "Fix schedules"):
        for schedule in schedules:
            # Get room name from schedule name
            bits = pat.findall(schedule.Name)
            room_name = None
            if bits:
                room_name = bits[0]
            else:
                continue

            # Check if level field already added
            stop = False
            fields = [schedule.Definition.GetField(x) for x in
                schedule.Definition.GetFieldOrder()]
            for f in fields:
                if f.GetName() == 'Level':
                    stop = True
            if stop:
                print('{} has Level field'.format(schedule.Name))
                continue

            # Add level field
            field = None
            for sf in schedule.Definition.GetSchedulableFields():
                if sf.GetName(doc) == 'Level':
                    field = schedule.Definition.AddField(sf)
            if not field:
                continue
            field.IsHidden = True

            # Get level to filter by
            level_num = room_name[0]
            level = [x for x in levels if x.Name == 'LEVEL ' + level_num][0]

            # Add level filter
            filters = list(schedule.Definition.GetFilters())
            level_filter = ScheduleFilter(
                field.FieldId, ScheduleFilterType.Equal, level.Id)
            filters.insert(0, level_filter)
            schedule.Definition.SetFilters(filters)
            print(schedule.Name)


def rename_schedules(doc):
    pat = re.compile('(FITOUT)[- ]+(\d{4})[- ]+(.*)', re.IGNORECASE)
    schedules = [x for x in FilteredElementCollector(doc).OfClass(View)
        if x.ViewType == ViewType.Schedule and 'fitout' in x.Name.lower()]

    with transaction(doc, 'Fix schedule titles'):
        for schedule in schedules:
            match = pat.findall(schedule.Name)
            if match:
                bits = match[0]
                new_name = '{} - {} - {}'.format(*bits).upper()
                try:
                    if schedule.Name != new_name:
                        schedule.Name = new_name
                except:
                    print(schedule.Name, new_name)


def rename_elevations(doc):

    rooms = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms)
    room_map = {
        x.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString():
        x.get_Parameter(BuiltInParameter.ROOM_NAME).AsString() for x in rooms}
    views = [x for x in FilteredElementCollector(doc).OfClass(View)
        if x.ViewType == ViewType.FloorPlan]

    #for elev in elevations:
    #    identifierParam = elev.LookupParameter(constants.IDENTIFIER_PARAM)
    #    descriptorParam = elev.LookupParameter(constants.DESCRIPTOR_TX)
    #    if identifierParam and \
    #       identifierParam.AsString() == 'PLOT' and \
    #       descriptorParam and \
    #       descriptorParam.AsString():
    #    print(elev.Name)

    #pat = re.compile('(\d{4})_(\d)')
    pat = re.compile('^(\d{4})+')

    with transaction(doc, 'Rename views'):
        for view in views:
            bits = pat.findall(view.Name)
            if bits:
                number = bits[0]
                name = room_map.get(number, 'XXX')
                print(view.Name, '{} - {} - PLAN'.format(','.join(bits), name))


def fix_view_names(doc):
    with transaction(doc, 'Rename views'):
        for view in FilteredElementCollector(doc).OfClass(View):
            title_param = view.get_Parameter(BuiltInParameter.VIEW_DESCRIPTION)
            sheet_param = view.get_Parameter(BuiltInParameter.VIEWPORT_SHEET_NUMBER)
            if sheet_param.AsString() and title_param.AsString().strip() and sheet_param.AsString().startswith('RM'):
                #print(title_param.AsString(), sheet_param.AsString())
                print(view.Name)
                title_param.Set('')


def rename_zone_views(doc):
    pat = re.compile(r'RCP - LEVEL (\d) - 1-(\d{2,3}) Zone (\d)')
    views = [x for x in FilteredElementCollector(doc).OfClass(View) if x.ViewType == ViewType.CeilingPlan]
    with transaction(doc, "Rename views"):
        for view in views:
            bits = pat.findall(view.Name)
            if bits:
                bits = bits[0]
                new_name = 'RCP_{}_{}_Z{}'.format(*bits)
                view.Name = new_name
                print(view.Name)


def fix_scope_boxes(doc):
    dw = DocumentWrapper(doc)
    scope_box_map = {x.Name: x for x in dw.scope_boxes}
    tr = Transaction(doc, 'Apply scope boxes to zone views')
    tr.Start()
    for view in dw.views:
        if 'Zone' in view.Name:
            scope_box_param = view.LookupParameter('Scope Box')
            if not scope_box_param:
                continue
            scope_box_name = scope_box_param.AsValueString()
            if scope_box_name == 'None':
                bits = re.findall('Zone (\d)$', view.Name)
                if not bits:
                    continue
                new_scope_box_name = '1-{} Zone {}'.format(view.Scale, bits[0])
                new_scope_box = scope_box_map[new_scope_box_name]
                scope_box_param.Set(new_scope_box.Id)
                print(view.Name, scope_box_param.AsValueString())
    tr.Commit()


def delete_unused_scope_boxes():
    dw = DocumentWrapper(doc)
    tr = Transaction(doc, 'Delete unused scope boxes')
    tr.Start()
    for scope_box in dw.unused_scope_boxes:
        doc.Delete(scope_box.Id)
    tr.Commit()


def rename_scope_boxes(doc):
    scope_boxes = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_VolumeOfInterest)
    tr = Transaction(doc, 'Rename scope boxes')
    tr.Start()
    for box in scope_boxes:
        new_name = box.Name.rstrip(' Scope Box')
        box.get_Parameter(BuiltInParameter.VOLUME_OF_INTEREST_NAME).Set(new_name)
        print(new_name)
    tr.Commit()


