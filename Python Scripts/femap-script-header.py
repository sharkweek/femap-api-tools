import pythoncom
import pyfemap as pf
import sys

try:
    existObj = pythoncom.connect(pf.model.CLSID) #Grabs active model
    app = pf.model(existObj)
except:
    sys.exit("femap not open") #Exits program if there is no active femap model
