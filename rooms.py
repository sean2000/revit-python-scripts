from api import *


def fix_room_names():
    rooms = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms)
    tr = Transaction(doc, 'Fix room names')
    tr.Start()
    for room in rooms:
        room_name = room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString()
        room_num = room.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString()
        if room_num in room_name:
            new_name = room_name.rstrip(room_num).strip()
            room.get_Parameter(BuiltInParameter.ROOM_NAME).Set(new_name)
            print room_name, new_name
    tr.Commit()


def renumber_rooms():
    """
    Renumber all rooms
    """
    rooms = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms)
    with transaction(doc, "Renumber rooms"):
        for room in rooms:
            param = room.get_Parameter(BuiltInParameter.ROOM_NUMBER)
            level = room.Level.Name
            print(param.AsString(), level)


def fix_rooms(doc):
    #with open(os.path.join(HOMEDIR, 'rooms.csv'), 'w') as f:
    #    w = csv.writer(f, lineterminator='\n')
    #    w.writerow(('ID', 'Room Name', 'Room Number', 'Level'))
    #    rooms = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms)
    #    for room in rooms:
    #        room_name = room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString()
    #        room_num = room.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString()
    #        w.writerow((room.UniqueId, room_name.encode('utf-8'), room_num, room.Level.Name))
    #        print(room_num)

    # ID,Level,Room Number,New Name,Old Name,Names Match

    rooms_dict = {}
    with open(os.path.join(HOMEDIR, 'RoomNames.csv')) as f:
        for row in csv.DictReader(f):
            if row['Names Match'] == 'FALSE':
                rooms_dict[row['ID']] = row

    with transaction(doc, "Rename rooms"):
        for room_id in rooms_dict:
            room_number = rooms_dict[room_id]['Room Number']
            incorrect_name = rooms_dict[room_id]['New Name']
            correct_name = rooms_dict[room_id]['Old Name'].upper()

            room = doc.GetElement(room_id)
            room.get_Parameter(BuiltInParameter.ROOM_NAME).Set(correct_name)

            print('{}: {} -> {}'.format(
                room_number, incorrect_name, correct_name))


def delete_unplaced_rooms(doc):
    rooms = list(FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms))
    count = 0
    with transaction(doc, "Delete unplaced rooms"):
        for room in rooms:
            if room.Location is None and room.Area == 0.0:
                print(room.Number)  #, room.Location, room.Area)
                doc.Delete(room.Id)
                count += 1
    print(count)


