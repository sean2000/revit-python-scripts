from collections import defaultdict
from itertools import groupby

from Autodesk.Revit.DB import Family
from Autodesk.Revit.DB import FamilyInstance
from Autodesk.Revit.DB import FamilySymbol
from Autodesk.Revit.DB import FilteredElementCollector

from api import DOC
import lib


def get_category_by_name(DOC, cat_name):
    for cat in DOC.Settings.Categories:
        if cat.Name == cat_name:
            return cat


def smart_purge(DOC):
    """
    27828407
      20206287 0
      20206291 0
    955
      20380607 0
    21208765
      20522972 4
    21165783
      21165966 2
    """
    symbol_map = {}
    family_map = defaultdict(dict)
    symbols = FilteredElementCollector(DOC).OfClass(FamilySymbol)
    instances = FilteredElementCollector(DOC).OfClass(FamilyInstance)

    instances_grouped = groupby(instances, lambda x: x.Symbol.Id)
    for symbol_id, instances in instances_grouped:
        symbol_map[symbol_id] = len(list(instances))

    symbols_grouped = groupby(symbols, lambda x: x.Family.Id)
    for family_id, symbols in symbols_grouped:
        print(family_id)
        for symbol in symbols:
            family_map[family_id][symbol.Id] = symbol_map.get(symbol.Id, 0)
            print('  {} {}'.format(symbol.Id, symbol_map.get(symbol.Id, 0)))

    with lib.transaction(DOC, "Smart purge"):
        for family_id, type_ids in family_map.items():
            has_instances = False
            for type_id, count in type_ids.items():
                if count > 0:
                    has_instance = True
                    continue
            #print family_id, has_instances
            if not has_instances:
                family = DOC.GetElement(family_id)
                if not family.IsCurtainPanelFamily:
                    print(family.Name)
                    DOC.Delete(family_id)


def smart_purge2(DOC):
    instances = FilteredElementCollector(DOC).OfClass(FamilyInstance)
    instance_families = {x.Symbol.Family.Id for x in instances}
    loaded_families = {x.Id for x in FilteredElementCollector(DOC).OfClass(Family)}
    unused_families = loaded_families.difference(instance_families)

    with transaction(DOC, "Smart purge"):
        for fam in unused_families:
            DOC.Delete(fam)


