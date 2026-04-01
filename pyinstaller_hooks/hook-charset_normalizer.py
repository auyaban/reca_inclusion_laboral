from importlib.util import find_spec


hiddenimports = []

if find_spec("charset_normalizer.md__mypyc") is not None:
    hiddenimports.append("charset_normalizer.md__mypyc")
