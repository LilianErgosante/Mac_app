from setuptools import setup

APP = ['rapport_auto.py']
DATA_FILES = []
OPTIONS = {
    'argv_emulation': True,
    'packages': [
        'tkinter', 'subprocess', 'os', 'pandas', 'matplotlib',
        'docx', 'tabula', 'tabulate', 'statistics', 'numpy'
    ],
    'includes': [
        'tkinter', 'subprocess', 'os', 'pandas', 'matplotlib',
        'docx', 'tabula', 'tabulate', 'statistics', 'numpy'
    ],
    'excludes': [
        'matplotlib.tests', 'numpy.random._examples'
    ]
}

setup(
    app=APP,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app', 'wheel'],
)
