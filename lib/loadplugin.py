import clr
clr.AddReference('mscorlib')
clr.AddReference('System')

import System
import System.Reflection


def load_assembly(path):
    # load the assembly into a byte array
    assemblyBytes = System.IO.File.ReadAllBytes(path)
    pdbPath = path[:-3] + 'pdb'
    if System.IO.File.Exists(pdbPath):
        # load debugging symbols
        pdbBytes = System.IO.File.ReadAllBytes(pdbPath)
        assembly = System.Reflection.Assembly.Load(assemblyBytes, pdbBytes)
    else:
        # no debugging symbols found
        assembly = System.Reflection.Assembly.Load(assemblyBytes)
    # make sure we can resolve assemblies from that directory
    folder = System.IO.Path.GetDirectoryName(path)
    System.AppDomain.CurrentDomain.AssemblyResolve += resolve_assembly_generator(folder)
    return assembly


def resolve_assembly_generator(folder):
    def result(sender, args):
        name = args.Name.split(',')[0]
        try:
            path = System.IO.Path.Combine(folder, name + '.dll')
            if not System.IO.File.Exists(path):
                return None
            return loadAssembly(path)
        except:
            import traceback
            traceback.print_exc()
            return None
    return result

