"""
Build timekeeper.app with py2app:

    pip install py2app
    python setup.py py2app

The finished app appears in the dist/ folder.
Drag dist/timekeeper.app to /Applications to install.
"""
from setuptools import setup

APP     = ["timekeeper.py"]
OPTIONS = {
    "argv_emulation": False,
    "packages": ["PySide6", "openpyxl", "objc", "AppKit", "Foundation"],
    "excludes": ["tkinter"],
    "plist": {
        "CFBundleName":             "timekeeper",
        "CFBundleDisplayName":      "timekeeper",
        "CFBundleIdentifier":       "com.nikoskatrakis.timekeeper",
        "CFBundleVersion":          "2.0.0",
        "CFBundleShortVersionString": "2.0.0",
        "NSHighResolutionCapable":  True,
        "LSUIElement":              False,   # show in Dock
    },
}

setup(
    app=APP,
    name="timekeeper",
    options={"py2app": OPTIONS},
    setup_requires=["py2app"],
)
