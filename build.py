#!/usr/bin/env python3
"""Build the .nvda-addon package for Spelling and Grammar Check.

An .nvda-addon file is a zip archive containing the add-on's files with
manifest.ini at the archive root. Run: python build.py
"""

import os
import re
import zipfile

ROOT = os.path.dirname(os.path.abspath(__file__))
ADDON_DIR = os.path.join(ROOT, "addon")


def getManifestValue(key):
	# manifest.ini is configobj syntax (triple-quoted multi-line values), which
	# configparser cannot read, so pick out the single-line key directly.
	with open(os.path.join(ADDON_DIR, "manifest.ini"), encoding="utf-8") as f:
		match = re.search(rf'^{key}\s*=\s*"?([^"\r\n]+?)"?\s*$', f.read(), re.MULTILINE)
	if not match:
		raise SystemExit(f"manifest.ini has no {key}")
	return match.group(1)


def main():
	name = getManifestValue("name")
	version = getManifestValue("version")
	outFile = os.path.join(ROOT, f"{name}-{version}.nvda-addon")
	with zipfile.ZipFile(outFile, "w", zipfile.ZIP_DEFLATED) as bundle:
		for dirPath, dirNames, fileNames in os.walk(ADDON_DIR):
			dirNames[:] = sorted(d for d in dirNames if d != "__pycache__")
			for fileName in sorted(fileNames):
				if fileName.endswith(".pyc"):
					continue
				filePath = os.path.join(dirPath, fileName)
				arcName = os.path.relpath(filePath, ADDON_DIR).replace(os.sep, "/")
				bundle.write(filePath, arcName)
	print(f"Built {outFile}")


if __name__ == "__main__":
	main()
