from setuptools import setup

APP = ['CTOS.py']
OPTIONS = {
    'argv_emulation': True,
    'iconfile': 'ctos.icns',
    'resources': ['Picture', 'Themes'],  
}

setup(
    app=APP,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)