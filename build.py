#!/usr/bin/env python3
"""Build the .nvda-addon package for Reply to Sender for Microsoft Outlook.

An .nvda-addon file is a zip archive containing the add-on's files with
manifest.ini at the archive root. Run: python build.py
"""

import configparser
import os
import zipfile

ROOT = os.path.dirname(os.path.abspath(__file__))
ADDON_DIR = os.path.join(ROOT, "addon")


def getManifestValue(key):
	parser = configparser.ConfigParser()
	with open(os.path.join(ADDON_DIR, "manifest.ini"), encoding="utf-8") as f:
		parser.read_string("[addon]\n" + f.read())
	return parser["addon"][key].strip('"')


def main():
	name = getManifestValue("name")
	version = getManifestValue("version")
	outFile = os.path.join(ROOT, f"{name}-{version}.nvda-addon")
	with zipfile.ZipFile(outFile, "w", zipfile.ZIP_DEFLATED) as bundle:
		for dirPath, dirNames, fileNames in os.walk(ADDON_DIR):
			dirNames[:] = [d for d in dirNames if d != "__pycache__"]
			for fileName in sorted(fileNames):
				filePath = os.path.join(dirPath, fileName)
				arcName = os.path.relpath(filePath, ADDON_DIR).replace(os.sep, "/")
				bundle.write(filePath, arcName)
	print(f"Built {outFile}")


if __name__ == "__main__":
	main()
