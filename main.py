import os
import inspect
import importlib
from collections import OrderedDict

from api import DOC

CURDIR = r'C:\Users\DodswoS\Dropbox\work\jacobs\revit_python'
HEADER = r'''
 ____        _   _                   ____            _       _
|  _ \ _   _| |_| |__   ___  _ __   / ___|  ___ _ __(_)_ __ | |_ ___
| |_) | | | | __| '_ \ / _ \| '_ \  \___ \ / __| '__| | '_ \| __/ __|
|  __/| |_| | |_| | | | (_) | | | |  ___) | (__| |  | | |_) | |_\__ \
|_|    \__, |\__|_| |_|\___/|_| |_| |____/ \___|_|  |_| .__/ \__|___/
       |___/                                          |_|
'''


def get_functions():
    functions = OrderedDict()
    stop_modules = ['scratch', 'main', 'test_runner', 'api']

    for mname in os.listdir(CURDIR):
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


def ui(functions):
    print(HEADER)

    for i, module in enumerate(functions, 1):
        print('{} {}'.format(i, module.__name__))

    module_index = int(raw_input('\nWhich module?: ')) - 1
    module = functions.keys()[module_index]

    for i, func in enumerate(functions[module], 1):
        print('{} {}'.format(i, func.__name__.title()))
        print('\t' + (func.__doc__ or ''))

    func_index = int(raw_input('\nWhich function?: ')) - 1
    func = functions[module][func_index]

    # Execute function
    func(DOC)

    choice = raw_input('"q" to exit or any key to continue: ')
    if choice.lower() == 'q':
        __window__.Close()
        return
    else:
        ui(functions)


if __name__ == '__main__':
    ui(get_functions())
    print('Finished')

