"""Functions related to annotation."""
from Autodesk.Revit.DB import BuiltInCategory
from Autodesk.Revit.DB import BuiltInParameter
from Autodesk.Revit.DB import FamilyInstance
from Autodesk.Revit.DB import FilteredElementCollector
from Autodesk.Revit.DB import FilteredWorksetCollector
from Autodesk.Revit.DB import SpatialElementTag
from Autodesk.Revit.DB import WorksetKind

from lib import transaction


def fix_room_tags(doc):
    """Find room tags that are outside their room and move them back"""
    room_tags = FilteredElementCollector(
        doc
    ).OfClass(
        SpatialElementTag
    ).OfCategory(
        BuiltInCategory.OST_RoomTags
    )

    with transaction(doc, 'Move room tags'):
        count = 0
        for tag in room_tags:
            room = tag.Room
            if not room:
                continue

            box = room.get_BoundingBox(None)
            if not box:
                continue

            tag_point = tag.Location.Point
            tag_is_in_room = tag_point.X > box.Min.X and \
                tag_point.X < box.Max.X and \
                tag_point.Y > box.Min.Y and \
                tag_point.Y < box.Max.Y

            if not tag_is_in_room:
                tag.Location.Move(room.Location.Point - tag.Location.Point)
                count += 1
                print(room.Level.Name, room.Number)

    print('{} room tags moved'.format(count))


def create_text_styles(doc):

    black = int('000000', 16)
    blue = int('ff0000', 16)
    std_sizes = [
        ('1.8mm', 1.8, black),
        ('1.8mm Blue', 1.8, blue),
        ('2.0mm', 2.0, black),
        ('2.5mm', 2.5, black),
        ('3.5mm', 3.5, black),
        ('5.0mm', 5.0, black),
        ('7.0mm', 7.0, black),
        ('10.0mm', 10, black),
    ]
    #size_pairs = [std_sizes[i:i + 2] for i in range(len(std_sizes))]
    #if len(size_pairs[-1]) == 1:
    #    size_pairs = size_pairs[:-1]

    text_types = FilteredElementCollector(doc).OfClass(TextNoteType)
    text_type = list(text_types)[0]
    with transaction(doc, 'Create text types'):
        for name, size, color in std_sizes:
            try:
                elem = text_type.Duplicate(name)
                elem.LookupParameter('Text Font').Set('Arial Narrow')
                elem.LookupParameter('Text Size').Set(size / MM)
                elem.LookupParameter('Color').Set(color)
                print(elem)
            except Exception as e:
                print(e)
                continue

    #for t in text_types:
    #    size_in_mm = t.LookupParameter('Text Size').AsValueString()
    #    size = float(size_in_mm.rstrip(' mm'))
    #    print(size)
    #    for lower, upper in size_pairs:
    #        if lower < size and size < upper:
    #            print('Change {} to {}'.format(size, lower))


def create_text(doc):
    type_id = doc.GetDefaultElementTypeId(ElementTypeGroup.TextNoteType)
    x = 0
    y = 0
    with transaction(doc, 'Create text'):
        text_types = FilteredElementCollector(doc).OfClass(TextNoteType)
        for tt in text_types:
            note = TextNote.Create(
                doc,
                doc.ActiveView.Id,
                XYZ(x, y, 0),
                tt.LookupParameter('Type Name').AsString(),
                tt.Id);
            y += 10 / MM


