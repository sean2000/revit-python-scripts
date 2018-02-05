
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


