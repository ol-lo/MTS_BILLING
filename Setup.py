from distutils.core import setup
import py2exe
from os import sys, path

try:
    sys.path.append(path.dirname(path.dirname(path.abspath(__file__))))
except NameError:  # py2exe script, not a module
    import sys

    sys.path.append(path.dirname(path.dirname(sys.argv[0])))

import os
import glob
import numpy

def find_data_files(source,target,patterns):
    """Locates the specified data-files and returns the matches
    in a data_files compatible format.

    source is the root of the source data tree.
        Use '' or '.' for current directory.
    target is the root of the target data tree.
        Use '' or '.' for the distribution directory.
    patterns is a sequence of glob-patterns for the
        files you want to copy.
    """
    if glob.has_magic(source) or glob.has_magic(target):
        raise ValueError("Magic not allowed in src, target")
    ret = {}
    for pattern in patterns:
        pattern = os.path.join(source, pattern)
        for file_name in glob.glob(pattern):
            if os.path.isfile(file_name):
                target_path = os.path.join(
                    target,
                    os.path.relpath(
                        file_name, source)
                )
                path_ = os.path.dirname(target_path)
                ret.setdefault(path_, []).append(file_name)
    return sorted(ret.items())
def numpy_dll_paths_fix():
    paths = set()
    np_path = numpy.__path__[0]
    for dirpath, _, filenames in os.walk(np_path):
        for item in filenames:
            if item.endswith('.dll'):
                paths.add(dirpath)

    sys.path.append(*list(paths))

numpy_dll_paths_fix()

setup(
    name="MTS BILLING REPORT",
    version="0.3",
    description="automate report",
    author="Maxim Nalimov",
    console=[
        {
            "script": 'base.py',
            "dest_base": "MTS BILLING REPORT",
            "icon_resources": [(0, "favicon.ico")]
        },
    ],
    options={
        'py2exe': {
            "dll_excludes": ["MSVCP90.dll"]
        }
    },
    data_files=find_data_files('data', '', [
        'config.yaml',
    ]),
)
