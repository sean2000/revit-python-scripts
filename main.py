import os
import inspect
import importlib
from collections import OrderedDict
import xml.etree.ElementTree as ET

from api import DOC, APP

HEADER = r'''
 ____        _   _                   ____            _       _
|  _ \ _   _| |_| |__   ___  _ __   / ___|  ___ _ __(_)_ __ | |_ ___
| |_) | | | | __| '_ \ / _ \| '_ \  \___ \ / __| '__| | '_ \| __/ __|
|  __/| |_| | |_| | | | (_) | | | |  ___) | (__| |  | | |_) | |_\__ \
|_|    \__, |\__|_| |_|\___/|_| |_| |____/ \___|_|  |_| .__/ \__|___/
       |___/                                          |_|
'''


def get_functions(start_dir):
    functions = OrderedDict()
    stop_modules = ['scratch', 'main', 'test_runner', 'api']

    for mname in os.listdir(start_dir):
        basename, ext = os.path.splitext(mname)
        if not ext == '.py':
            continue
        if basename in stop_modules:
            continue
        m = importlib.import_module(basename)

        for name, obj in inspect.getmembers(m):
            if not inspect.isfunction(obj):
                continue
            if name.startswith('_'):
                continue
            if not obj.__doc__:
                continue

            if m not in functions:
                functions[m] = []
            functions[m].append(obj)

    return functions


def get_index_input(prompt):
    index = 0
    while 1:
        try:
            index = int(raw_input(prompt)) - 1
            break
        except ValueError:
            print('Invalid index')
    return index


def prettify(text):
    return (' '.join(text.split('_'))).title()


def ui(functions):
    print(HEADER)

    for i, module in enumerate(functions, 1):
        print('{} {}'.format(i, module.__name__.title()))

    module_index = get_index_input('Which module?: ')
    module = functions.keys()[module_index]

    print('')
    for i, func in enumerate(functions[module], 1):
        print('{} {}'.format(i, prettify(func.__name__)))
        print('  ' + (func.__doc__.strip() or ''))
        print('')

    func_index = get_index_input('Which function?: ')

    func = functions[module][func_index]
    print('')

    # Execute function
    try:
        retval = func(DOC)
        print('{} executed successfully'.format(prettify(func.__name__)))
    except Exception as e:
        print(e)

    print('')
    choice = raw_input('"q" to exit or any key to continue: ')
    if choice.lower() == 'q':
        __window__.Close()
        return
    else:
        ui(functions)


def get_curdir():
    curdir = None
    xml_path = os.path.join(
        os.environ['APPDATA'],
        'RevitPythonShell' + APP.VersionNumber,
        'RevitPythonShell.xml')
    root = ET.parse(open(xml_path)).getroot()
    for command in root.iter('Command'):
        if command.attrib['name'] == 'Revit Scripts':
            path = command.attrib['src']
            curdir = os.path.dirname(path)
    return curdir


if __name__ == '__main__':
    curdir = get_curdir()
    ui(get_functions(curdir))
    print('Finished')

