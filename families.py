from collections import defaultdict
import csv
import os

from Autodesk.Revit.DB import BuiltInParameter
from Autodesk.Revit.DB import BuiltInCategory
from Autodesk.Revit.DB import Family
from Autodesk.Revit.DB import FamilyInstance
from Autodesk.Revit.DB import FamilySymbol
from Autodesk.Revit.DB import FilteredElementCollector

from api import *
from lib import transaction
from lib import get_temp_path
import constants


def report_corrupt_families(doc):
    """Generate CSV of all loaded families and their corrupted status"""
    families = FilteredElementCollector(doc).OfClass(Family)
    rows = []
    category = raw_input('Enter category: ').strip()

    for family in families:
        if not family.IsEditable or family.IsInPlace:
            continue
        if category and family.FamilyCategory.Name != category:
            continue
        try:
            fam_doc = doc.EditFamily(family)
            fam_doc.Close(False)
            rows.append((family.FamilyCategory.Name, family.Name, True))
            print('{} - OK'.format(family.Name))
        except Exception:
            print('{} - Corrupt'.format(family.Name))
            rows.append((family.FamilyCategory.Name, family.Name, False))

    csv_path = get_temp_path('corrupt.csv')
    with open(csv_path, 'w') as f:
        w = csv.writer(f, lineterminator='\n')
        w.writerow(('Category', 'Family', 'OK'))
        w.writerows(rows)
    os.startfile(csv_path, 'open')


def fix_duplicate_marks(doc):
    """Find elements with duplicate marks and rename"""
    mark_map = defaultdict(list)
    instances = FilteredElementCollector(doc).OfClass(FamilyInstance)

    for instance in instances:
        mark = instance.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)
        if mark:
            mark_map[mark.AsString()].append(instance.Id)

    with transaction(doc, 'Fix duplicate marks'):
        for mark, ids in mark_map.items():
            if len(ids) < 2:
                continue
            if not mark:
                continue

            for i, id in enumerate(ids[1:], 1):
                new_mark = '{}-{}'.format(mark, i)
                instance = doc.GetElement(id)
                instance.get_Parameter(BuiltInParameter.ALL_MODEL_MARK).Set(new_mark)
                print(new_mark)


def rename_single_types(doc, type_name='Default'):
    """Rename single types"""
    families = FilteredElementCollector(doc).OfClass(Family)
    with transaction(doc, 'Rename types'):
        for family in families:
            symbols = list(family.GetFamilySymbolIds())
            if len(symbols) == 1:
                symbol = doc.GetElement(symbols[0])
                symbol.Name = type_name


def rename_families(doc):

    acronyms = ['MAT', 'SSTN', 'OR', 'SC', 'PAED', 'ICU', 'WUP', 'FHR', 'NB',
        'AV', 'FIP', 'ED', 'CU', 'NWOW', 'NBC', 'EDPAEDS', 'ISOL', 'IPU', 'PB',
        'PECC', 'UB', 'UP', 'CIS']
    remove_words = ['SPJN', 'HE', '3D']
    families = FilteredElementCollector(doc).OfClass(Family)

    with transaction(doc, 'Rename families'):
        for fam in families:
            bits = re.findall(r'(SPJN[0-9X]{2,5})\s*[-_]\s*(.*)', fam.Name)
            if not bits:
                continue

            code, name = bits[0]
            name_bits = re.sub(r'[-_\s]', ' ', name).split()

            out = []
            for bit in name_bits:
                if bit.upper() in remove_words:
                    continue
                if bit in acronyms:
                    out.append(bit)
                else:
                    out.append(bit.title())

            name = ' '.join(out)
            fam.Name = '{} - {}'.format(code, name)
            print(fam.Name)


def report_family_types(doc):
    """Generate CSV report of types and parameters of family"""
    fm = doc.FamilyManager

    types = []
    fieldnames = {'Type'}
    for ft in fm.Types:
        ft_dict = {'Type': ft.Name}
        for p in fm.GetParameters():
            p_value = ft.AsValueString(p) or ft.AsString(p)
            if not p_value:
                continue
            ft_dict[p.Definition.Name] = p_value
            fieldnames.add(p.Definition.Name)
        types.append(ft_dict)

    csv_path = get_temp_path('types.csv')
    with open(csv_path, 'w') as f:
        dw = csv.DictWriter(f, fieldnames, lineterminator='\n')
        dw.writeheader()
        dw.writerows(types)

    os.startfile(csv_path, 'open')


def report_families(doc):
    """Generate report of families"""

    #start_dir = r'H:\NBIF\Health Library\Revit'
    #start_dir = r'C:\Users\DodswoS\Documents\Revit testing\New Families\temp\GE'
    start_dir = r'C:\Users\DodswoS\Documents\Revit testing\New Families\temp\Sorted'
    out_dir = r'C:\Users\DodswoS\Documents\Revit testing\New Families\temp\Sorted'
    out = []

    outpath = os.path.join(constants.HOMEDIR, 'types.csv')

    for root, dirs, files in os.walk(start_dir):
        for fname in files:

            with open(outpath, 'a') as f:
                w = csv.writer(f, lineterminator='\n')

                basename, ext = os.path.splitext(fname)
                if ext != '.rfa':
                    continue
                filepath = os.path.join(root, fname)
                famdoc = APP.OpenDocumentFile(filepath)
                fm = famdoc.FamilyManager
                category = famdoc.OwnerFamily.FamilyCategory.Name

                for ft in fm.Types:
                    for p in fm.GetParameters():
                        p_value = ft.AsValueString(p) or ft.AsString(p)
                        if not p_value:
                            continue
                        row = (
                            start_dir,
                            category,
                            fname,
                            ft.Name,
                            p.Definition.Name,
                            p_value,
                        )
                        try:
                            w.writerow(row)
                        except Exception as e:
                            print(e)
                            print(row)

                famdoc.Close(False)

            # Move family to category folder
            #category_folder = os.path.join(out_dir, category)
            #if not os.path.exists(category_folder):
            #    os.mkdir(category_folder)
            #shutil.move(filepath, category_folder)
            #print('moved to {}'.format(category_folder))

    #with open(outpath, 'w') as f:
    #    w = csv.writer(f, lineterminator='\n')
    #    w.writerow((
    #        'Source', 'Category', 'Family', 'Type', 'Parameter Name', 'Parameter Value'))
    #    #w.writerows(out)
    #    for row in out:
    #        try:
    #            w.writerow(row)
    #        except Exception:
    #            print(row)

    os.startfile(outpath, 'open')


def fix_types(doc):
    families = set()
    with open(os.path.join(HOMEDIR, 'types1.csv')) as f:
        families = {row['Family'] for row in csv.DictReader(f)}

    start_dir = r'H:\NBIF\Health Library\Revit'
    #start_dir = r'I:\NBIF\Projects\NB98100\Deliverables\BIM\01_Architecture\F_Project Content\Exported Families 20170109'
    out = []

    for root, dirs, files in os.walk(start_dir):
        for fname in files:
            if fname not in families:
                print('Ignoring {}'.format(fname))
                continue
            print(fname)
            basename, ext = os.path.splitext(fname)
            if ext != '.rfa':
                continue
            filepath = os.path.join(root, fname)
            famdoc = app.OpenDocumentFile(filepath)
            fm = famdoc.FamilyManager

            for ft in fm.Types:
                for p in fm.GetParameters():
                    out.append(
                        (
                            start_dir,
                            fname,
                            ft.Name,
                            p.Definition.Name,
                            ft.AsValueString(p) or ft.AsString(p)
                        )
                    )
            famdoc.Close(False)

    with open(os.path.join(HOMEDIR, 'types.csv'), 'w') as f:
        w = csv.writer(f, lineterminator='\n')
        w.writerow((
            'Source', 'Family', 'Type', 'Parameter Name', 'Parameter Value'))
        w.writerows(out)


def duplicate_families_by_parameter(doc):
    #categories = ['Furniture', 'Furniture Systems', 'Specialty Equipment']
    instances = FilteredElementCollector(doc).OfClass(FamilyInstance)
    families = FilteredElementCollector(doc).OfClass(Family)
    instance_families = {x.Symbol.Family.Name for x in instances}
    #loaded_families = {x.Name for x in families if x.FamilyCategory and x.FamilyCategory.Name in categories}

    rows = []
    param_name = constants.DESCRIPTOR_TX
    param_map = defaultdict(set)

    symbols = FilteredElementCollector(doc).OfClass(FamilySymbol)
    for s in symbols:
        param = s.LookupParameter(param_name)
        if param and param.AsString():
            param_map[param.AsString()].add(s)

    for param, symbols in param_map.items():
        if len(symbols) > 1:
            for x in param_map[param]:
                rows.append([param, x.Category.Name, x.Family.Name, x.UniqueId, x.Family.Name in instance_families])

    with open(os.path.join(HOMEDIR, 'Duplicates by {}.csv'.format(param_name)), 'w') as f:
        w = csv.writer(f, lineterminator='\n')
        w.writerow([param_name, 'Category', 'Family', 'Type ID', 'Used'])
        w.writerows(rows)

    print('Done')


def fix_duplicate_families(doc):
    dupe_family_names = {}
    for fam in FilteredElementCollector(doc).OfClass(Family):
        if fam.Name.endswith('3D1'):
            dupe_family_names[fam.Name] = set()

    pprint(dupe_family_names)


def rename_doors(doc):
    doors = FilteredElementCollector(
        doc
    ).OfClass(
        FamilySymbol
    ).OfCategory(
        BuiltInCategory.OST_Doors
    )
    patt = re.compile(r'D(\d{1,3})([a-zA-Z]*) (2100)')

    with transaction(doc, 'Rename doors'):
        for symbol in doors:
            name_param = symbol.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
            name = name_param.AsString()
            bits = patt.findall(name)
            if bits:
                bits = bits[0]
                new_name = 'D{}{}'.format(bits[0], bits[1])
                symbol.Name = new_name
                print('{} -> {}'.format(name, new_name))


def set_groups(doc):
    csvfile = os.path.join(HOMEDIR, 'br2_groups.csv')
    bimid_dict = defaultdict(set)
    for item in csv.DictReader(open(csvfile)):
        bimid_dict[item['BIM ID']].add(item['Budget: Code'])

    all_symbols = FilteredElementCollector(doc).OfClass(FamilySymbol)
    param_dict = {}
    for s in all_symbols:
        param = s.LookupParameter(constants.DESCRIPTOR_TX)
        if param:
            param_dict[get_param_value(param, doc)] = s

    with transaction(doc, 'Set groups'):
        for id, groups in bimid_dict.items():
            if len(groups) != 1:
                continue
            matching_symbol = param_dict.get(id)
            if matching_symbol:
                group = list(groups)[0]
                param = matching_symbol.LookupParameter('InstallationGroup_TX')
                if param:
                    param.Set(group)
                    print(id, group)



def fix_hfbs(doc):
    symbols = FilteredElementCollector(doc).OfClass(FamilySymbol)
    tr = Transaction(doc, 'Fix HFBS codes')
    tr.Start()
    for symbol in symbols:
        hfbs_param = symbol.LookupParameter(constants.HFBS_CODE_PARAM)
        if hfbs_param:
            val = hfbs_param.AsString()
            if not val:
                continue
            bits = val.split('-')
            if len(bits) == 2:
                try:
                    prefix = bits[0]
                    if prefix == 'S':
                        prefix = 'SE'
                    codenum = int(bits[1])
                    new_val = '{}-{:06d}'.format(prefix, codenum)
                    if val != new_val:
                        hfbs_param.Set(new_val)
                        print('{} -> {}'.format(val, new_val))
                except ValueError:
                    pass
    tr.Commit()


def set_door_ids(doc):
    tr = Transaction(doc, 'Fix HFBS codes')
    tr.Start()
    dw = DocumentWrapper(doc)
    for door in dw.doors:
        door.LookupParameter('ElementID').Set(door.Id.IntegerValue)
        print(door.Id)
    tr.Commit()


def fix_doors(doc):
    dw = DocumentWrapper(doc)
    with transaction(doc, "Fix doors"):
        for door in dw.doors:
            for k in range(1, 4):
                param = door.LookupParameter("Kick_K{}_YN".format(k))
                if param:
                    param.Set(False)
            print(get_param_value(door.get_Parameter(BuiltInParameter.ALL_MODEL_MARK), doc))


def swap_symbol(doc, family_name, from_type_name, to_type_name):
    dw = DocumentWrapper(doc)
    to_symbol = [x for x in dw.symbols if
        x.FamilyName == family_name and x.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString() == to_type_name][0]
    instances = [x for x in dw.instances if
        x.Symbol.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString() == from_type_name and
        x.Symbol.FamilyName == family_name]
    with transaction(doc, "Swap symbols"):
        for instance in instances:
            if instance.GroupId.IntegerValue < 0:
                print('{} is in group {}'.format(instance.Id, instance.GroupId))
            else:
                print(instance.GroupId, instance.Id)
                instance.Symbol = to_symbol


def purge_unused_types(doc):
    """Remove unused symbols for all families matching name_pattern"""
    name_pattern = raw_input('Enter family name pattern: ')
    matching_families = [x for x in FilteredElementCollector(doc).OfClass(
        Family) if name_pattern in x.Name]
    instances = FilteredElementCollector(doc).OfClass(FamilyInstance)
    used_symbols = {x.Symbol.Id for x in instances}
    count = 0
    with transaction(doc, "Delete unused symbols"):
        for family in matching_families:
            for symbol_id in family.GetFamilySymbolIds():
                if symbol_id not in used_symbols:
                    doc.Delete(symbol_id)
                    count += 1

    print('{} symbols deleted in {} families'.format(
        count, len(matching_families)))

