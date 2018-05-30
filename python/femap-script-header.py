"""This header imports all the necessary packages for using Python with
Femap and connects to the current active model."""

import pythoncom
import pyfemap as pf  # the Python Femap API library
import sys
from pyfemap import constants as fc  # the Femap constants library

try:
    existObj = pythoncom.connect(pf.model.CLSID)  # Grabs active model
    app = pf.model(existObj)
except:
    sys.exit("Femap is not open") #Exits program if there is no active femap model
