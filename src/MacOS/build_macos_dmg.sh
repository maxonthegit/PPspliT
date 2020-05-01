#!/bin/bash

# HOWTO
# - Refresh the PPspliT.ppam resource in the installer bundle application (possibly using Apple's Script Editor)
# - Run this script to generate the DMG file

hdiutil create -size 1m -fs HFS+ -srcfolder PPspliT\ for\ MacOS/  -volname PPspliT -ov -format UDZO PPspliT.dmg

