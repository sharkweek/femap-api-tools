"""This module provides some useful Python tools to interface with Femap."""

import __main__
import pythoncom
import pyfemap1142  # the Python Femap API library
import sys

try:
    existObj = pythoncom.connect(pyfemap1142.model.CLSID)  # Grabs active model
except:
    sys.exit("Femap not open")  # Exits program if there is no active femap
    # model

app = pyfemap1142.model(existObj)
fc = pyfemap1142.constants

if __name__ != '__main__':
    # Pass imported entities to __main__
    __main__.app = app
    print("Femap model imported as 'app'")
    __main__.fc = fc
    print("Femap constants imported as 'fc'")
    __main__.pf = pyfemap1142
    print("Femap v11.4.2 API library imported as 'pf'")
