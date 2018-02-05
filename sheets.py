from api import *


def set_titleblock(doc, identifier_tx, family_name, type_name):
    sheets = [x for x in FilteredElementCollector(doc).OfClass(ViewSheet) if x.LookupParameter('Identifier_TX').AsString() == 'DESIGN DEVELOPMENT']
    new_titleblock = [x for x in FilteredElementCollector(doc).OfClass(FamilySymbol).OfCategory(BuiltInCategory.OST_TitleBlocks)
        if x.FamilyName == family_name and x.Name == type_name][0]
    tr = Transaction(doc, "Titleblock change")
    tr.Start()
    for sheet in sheets:
        elements_on_sheet = FilteredElementCollector(doc, sheet.Id).OfClass(FamilyInstance).OfCategory(BuiltInCategory.OST_TitleBlocks)
        for element in elements_on_sheet:
            element.Symbol = new_titleblock
    tr.Commit()


def change_titleblocks(doc):
    sheets = [x for x in FilteredElementCollector(doc).OfClass(ViewSheet) if x.LookupParameter('Identifier_TX').AsString() == 'DESIGN DEVELOPMENT']
    tr = Transaction(doc, "Titleblock change")
    zone_pat = re.compile("ZONE (\d)$")
    tr.Start()
    for sheet in sheets:
        elements_on_sheet = FilteredElementCollector(doc, sheet.Id).OfClass(FamilyInstance).OfCategory(BuiltInCategory.OST_TitleBlocks)
        for element in elements_on_sheet:
            show_keyplan = element.LookupParameter("Show Key Plan")
            show_northpoint = element.LookupParameter("North Point Visibility")
            is_plan = 'PLAN' in sheet.Name.upper()
            if show_keyplan:
                print sheet.Name
                show_keyplan.Set(is_plan)
                show_northpoint.Set(is_plan)
                zone = zone_pat.findall(sheet.Name.upper())
                if zone:
                    print zone[0]
                    element.LookupParameter("Show Key Plan " + zone[0]).Set(True)
    tr.Commit()


def add_keyplan(doc):
    key_plan_name = 'BS2 Key Plan'
    sheets = [x for x in FilteredElementCollector(doc).OfClass(ViewSheet) if x.LookupParameter('Identifier_TX').AsString() == 'DESIGN DEVELOPMENT']
    keyplan = [x for x in FilteredElementCollector(doc).OfClass(FamilySymbol) if x.FamilyName == key_plan_name][0]
    location = XYZ(340/MM, 31/MM, 0)
    zone_pat = re.compile("ZONE (\d)$")

    tr = Transaction(doc, "Titleblock change")
    tr.Start()

    for sheet in sheets:
        is_plan = 'PLAN' in sheet.Name.upper()
        scale = sheet.get_Parameter(BuiltInParameter.SHEET_SCALE).AsString().strip()
        is_50 = scale == '1 : 50'
        is_100 = scale == '1 : 100'
        if is_plan:
            zone = zone_pat.findall(sheet.Name.upper())
            instance = doc.Create.NewFamilyInstance(location, keyplan, sheet)
            print sheet.Name, zone
            if is_50 and zone:
                zone_param = "1-50 Zone {}".format(zone[0])
                instance.LookupParameter(zone_param).Set(True)
            if is_100 and zone:
                zone_param = "1-100 Zone {}".format(zone[0])
                instance.LookupParameter(zone_param).Set(True)

    tr.Commit()


def add_note_to_sheets(doc):
    sheets = [x for x in FilteredElementCollector(doc).OfClass(ViewSheet) if x.LookupParameter('Description_TX').AsString() == 'Moved to Main']
    warning = [x for x in FilteredElementCollector(doc).OfClass(FamilySymbol) if x.FamilyName == 'Warning Note'][0]
    location = XYZ(500 / MM, 360 / MM, 0)

    tr = Transaction(doc, "Titleblock change")
    tr.Start()

    for sheet in sheets:
        instance = doc.Create.NewFamilyInstance(location, warning, sheet)

    tr.Commit()


def viewport(doc):
    viewportType = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Viewports)
    vp = doc.GetElement(uidoc.Selection.GetElementIds()[0])
    vp.get_Parameter(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM).Set(viewportType)


def hide_schedules_on_sheets(doc):
    logfilename = os.path.join(HOMEDIR, 'sheets_done.txt')
    sheets_done = [x.strip() for x in open(logfilename).readlines()]
    all_sheets = [x for x in FilteredElementCollector(doc).OfClass(ViewSheet)
        if x.SheetNumber.startswith('RM') and
              x.SheetNumber not in sheets_done][:20]
    print('{} sheets left to process'.format(len(all_sheets)))
    sheets = all_sheets[:20]

    with open(logfilename, 'a') as f:
        with transaction(doc, 'Hide schedules on room layout sheets'):
            for sheet in sheets:
                print(sheet.SheetNumber)
                schedules = list(FilteredElementCollector(
                    doc, sheet.Id
                ).OfClass(ScheduleSheetInstance))
                f.write('{}{}'.format(sheet.SheetNumber, os.linesep))

                schedule_ids = List[ElementId](
                    x.Id for x in schedules
                        if x.Name.upper().startswith('FITOUT'))
                if schedule_ids.Count == 0:
                    continue

                sheet.HideElements(schedule_ids)
                print('{} schedules hidden'.format(len(schedule_ids)))

        f.write('----------{}'.format(os.linesep))


def blank_sheet_parameters(doc):

    with open(os.path.join(HOMEDIR, 'sheet_parameters.txt'), 'wa') as f:
        for sheet in FilteredElementCollector(doc).OfClass(ViewSheet):
            param_names = [
                'Sheet Issue Date',
                'Drawn Date',
                'Drg Checked Date',
                'Design Date',
                'Design Check Date',
                'Zone Manager Date',
                'Design Manager Date',
            ]
            for name in param_names:
                param = sheet.LookupParameter(name)
                f.write("{};{};{};{};{}\n".format(
                    datetime.now(),
                    doc.PathName,
                    sheet.SheetNumber,
                    name,
                    param.AsString()))

def sheet_titles(doc):
    sheets = FilteredElementCollector(doc).OfClass(ViewSheet)

    with transaction(doc, 'Fix sheet titles'):
        for sheet in sheets:
            if len(sheet.Name) < 25:
                continue
            bits = sheet.Name.split(' - ')
            if len(bits) < 2:
                continue
            sheet.Name = bits[0]
            sheet.LookupParameter('Title_Line2').Set(' - '.join(bits[1:]))
            print(bits)


def viewports(doc):
    sheet = doc.ActiveView
    titleblock = list(FilteredElementCollector(doc, sheet.Id).OfClass(FamilyInstance).OfCategory(BuiltInCategory.OST_TitleBlocks))[0]
    width = titleblock.get_Parameter(BuiltInParameter.SHEET_WIDTH).AsDouble()
    height = titleblock.get_Parameter(BuiltInParameter.SHEET_HEIGHT).AsDouble()
    #print constants.MM
    #print width, height
    print width * constants.MM, height * constants.MM
    for id in sheet.GetAllViewports():
        vp = doc.GetElement(id)
        view = doc.GetElement(vp.ViewId);
        box = vp.GetBoxCenter()
        print view.Name, box.X * constants.MM, box.Y * constants.MM


