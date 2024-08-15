from setuptools import setup
import subprocess
import os

APP = ['main.py']
OPTIONS = {
    'iconfile': 'assets/icon.icns',
    'plist': {
        'CFBundleName': 'SAC Cleanup',
        'CFBundleDisplayName': 'SAC Cleanup',
        'CFBundleIdentifier': 'au.edu.vic.woodleigh.saccleanup',
        'CFBundleVersion': '3.0',
        'CFBundleShortVersionString': '3',
    },
    'packages': ['tkinter'],
    'includes': ['time', 'subprocess', 'shutil', 'os'],
}

setup(
    app=APP,
    name='SAC Cleanup',
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)

subprocess.run(['codesign', '--deep', '--signature-size', '9400', '-f', '-s', 'Developer ID Application: Woodleigh School (SMLKBTR495)', 'dist/SAC Cleanup.app'])
