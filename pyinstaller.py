import PyInstaller.__main__

PyInstaller.__main__.run([
    'integrate.py',
    '--onedir',      # Use --onefile for a single executable, or --onedir for a folder
    '--windowed',
    '--noconsole',
    '--icon=logo.ico'
])