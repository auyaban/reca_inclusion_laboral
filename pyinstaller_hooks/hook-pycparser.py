from importlib.util import find_spec


hiddenimports = []

for module_name in ("pycparser.lextab", "pycparser.yacctab"):
    if find_spec(module_name) is not None:
        hiddenimports.append(module_name)
