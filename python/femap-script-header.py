"""This header imports all the necessary packages for using Python with
Femap and connects to the current active model."""

import pythoncom
import pyfemap  # the Python Femap API library
import sys

try:
    existObj = pythoncom.connect(pyfemap.model.CLSID)  # Grabs active model
    app = pyfemap.model(existObj)
except:
    sys.exit("Femap is not open")  # Exits if there is no active femap model
