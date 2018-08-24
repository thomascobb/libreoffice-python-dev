#!/bin/env python
# https://gist.github.com/JannieT/6fee66dc821fd0fe6af68063920149f0

from zipfile import ZipFile
import shutil
import sys
import os

if len(sys.argv) < 3:
    print("Usage: {} scriptfile hostdocument".format(sys.argv[0]))
    exit()

MACRO_FILE = sys.argv[1]
DOCUMENT_FILE = sys.argv[2]
MANIFEST_PATH = 'META-INF/manifest.xml'
EMBED_PATH = 'Scripts/python/' + MACRO_FILE
MANIFEST_ENTRY = \
    '<manifest:file-entry manifest:media-type="application/binary" ' + \
    'manifest:full-path="{}"/>'

hasMeta = False
with ZipFile(DOCUMENT_FILE) as bundle:
    # grab the manifest
    manifest = []
    for rawLine in bundle.open('META-INF/manifest.xml'):
        line = rawLine.decode('utf-8')
        if '</manifest:manifest>' in line and MACRO_FILE not in line:
            for path in ['Scripts/', 'Scripts/python/', EMBED_PATH]:
                manifest.append(MANIFEST_ENTRY.format(path))
        manifest.append(line)

    # remove the manifest and script file
    with ZipFile(DOCUMENT_FILE + '.tmp', 'w') as tmp:
        for item in bundle.infolist():
            buf = bundle.read(item.filename)
            if item.filename not in [MANIFEST_PATH, EMBED_PATH]:
                tmp.writestr(item, buf)

os.remove(DOCUMENT_FILE)
shutil.move(DOCUMENT_FILE + '.tmp', DOCUMENT_FILE)

with ZipFile(DOCUMENT_FILE, 'a') as bundle:
    bundle.write(MACRO_FILE, EMBED_PATH)
    bundle.writestr(MANIFEST_PATH, ''.join(manifest))

print("Added the script {} to {}".format(MACRO_FILE, DOCUMENT_FILE))
