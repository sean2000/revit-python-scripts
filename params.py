from collections import defaultdict

from api import *


def fix_marks(doc):
    """
    Renumber duplicate marks with an incremental suffix
    """
    marks_map = defaultdict(list)
    all_instances = FilteredElementCollector(doc).OfClass(FamilyInstance)

    for instance in all_instances:
        mark_param = instance.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)
        if not mark_param:
            continue
        mark_val = mark_param.AsString()
        if not mark_val:
            continue
        marks_map[mark_val].append(instance)

    count = 0
    with transaction(doc, 'Fix marks'):
        for mark, elems in marks_map.items():
            if len(elems) > 1:
                inc = 1
                for elem in elems:
                    new_mark = '{}-{}'.format(mark, inc)
                    elem.get_Parameter(BuiltInParameter.ALL_MODEL_MARK).Set(new_mark)
                    inc += 1
                    count += 1
                    print('{} {}'.format(count, new_mark))


def delete_params(doc):

    with transaction(doc, 'Delete parameters'):
        for param in doc.FamilyManager.Parameters:
            if param.Definition.Name.startswith('Status'):
                doc.Delete(param.Id)
            if param.Definition.Name.endswith('YN'):
                doc.Delete(param.Id)


def delete_parameters():
    csvpath = r'C:\Users\DodswoS\Documents\parameters_to_delete.csv'
    params_to_delete = []

    with open(csvpath) as f:
        for row in csv.reader(f):
            params_to_delete += row

    iterator = doc.ParameterBindings.ForwardIterator()
    definitions_to_delete = []
    while iterator.MoveNext():
        definition = iterator.Key
        if definition.Name in params_to_delete:
            definitions_to_delete.append(definition)

    tr = Transaction(doc, 'Delete parameters')
    tr.Start()
    for definition in definitions_to_delete:
        doc.ParameterBindings.Remove(definition)
        print definition.Name
    tr.Commit()


def copy_parameter_values(doc, src_param, dest_param):
    """Copy the value of one parameter to another parameter"""

    families = FilteredElementCollector(doc).OfClass(Family)

    with transaction(doc, 'Copy parameters'):
        for fam in families:
            print(fam.Name)
            for symbol_id in fam.GetFamilySymbolIds():
                symbol = doc.GetElement(symbol_id)
                src_param = symbol.LookupParameter(src_param)
                dest_param = symbol.LookupParameter(dest_param)
                if not src_param or not dest_param:
                    continue

                try:
                    dest_param.Set(src_param.AsString())
                except Exception, e:
                    print e


def set_parameter_all_types(doc):

    fm = doc.FamilyManager
    th = fm.get_Parameter('TransomHeight_LE')
    td = fm.get_Parameter('TransomDepth_LE')
    tv = fm.get_Parameter('ShowTransom_YN')
    #param = fm.get_Parameter('Descriptor_TX')
    #descriptions = {
    #    'HA': 'Hinge-side Away',
    #    'LA': 'Latch-side Away',
    #    'EA': 'Either side Away',
    #    'FA': 'Front Away',
    #    'HT': 'Hinge-side Towards',
    #    'LT': 'Latch-side Towards',
    #    'ET': 'Either side Towards',
    #    'FT': 'Front Towards',
    #}

    with transaction(doc, 'Set parameters'):
        for t in fm.Types:
            fm.CurrentType = t
            #code, size = t.Name.split()
            #description = descriptions[code]
            #fm.Set(param, '{} {}'.format(description, size))
            fm.Set(th, 1500 / MM)
            fm.Set(td, 100 / MM)
            fm.Set(tv, False)
            print(t.Name)
    print('Finished')



def fix_family_hfbs():
    fm = doc.FamilyManager
    hfbs_param = fm.get_Parameter(constants.HFBS_CODE_PARAM)
    print hfbs_param
    if hfbs_param:
        val = hfbs_param.AsString()
        bits = val.split('-')
        if len(bits) == 2:
            try:
                prefix = bits[0]
                if prefix == 'S':
                    prefix = 'SE'
                codenum = int(bits[1])
                new_val = '{}-{:06d}'.format(prefix, codenum)
                if val != new_val:
                    #hfbs_param.Set(new_val)
                    fm.Set(hfbs_param, new_val)
                    print '{} -> {}'.format(val, new_val)
            except ValueError:
                pass



def set_bim_id(doc):
    elements = FilteredElementCollector(doc).OfClass(FamilySymbol)
    with transaction(doc, 'Populate BIM ID parameter'):
        for elem in elements:
            dtx_param = elem.LookupParameter('Descriptor_TX')
            bid_param = elem.LookupParameter('BIM ID')
            if dtx_param and bid_param:
                bid_param.Set(dtx_param.AsString())
                print(dtx_param.AsString(), bid_param.AsString())


