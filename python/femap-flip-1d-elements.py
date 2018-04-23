import pythoncom
import pyfemap as pf  # the Python Femap API library
import sys
from pyfemap import constants as fc  # the Femap constants library

try:
    existObj = pythoncom.connect(pf.model.CLSID) #Grabs active model
    app = pf.model(existObj)
except:
    sys.exit("Femap is not open") #Exits program if there is no active femap model

app.feAppMessage(fc.FCM_COMMAND, "Flip 1D Element Orientation")

pickSet = app.feSet
pickSet.Select(entityType, clear, prompt)

rc, _, setArray = pickSet.GetArray()  # get the array
del _, pickSet  # delete intermediate variables from cache
if rc == fc.FE_CANCEL:
    response = "User cancelled selection..."
    app.feAppMessage(fc.FCM_ERROR, response)
    print(response)
elif rc == fc.FE_NOT_EXIST:
    response = "No entities of selected type exist."
    app.feAppMessage(fc.FCM_ERROR, response)
    print(response)


for each in setArray:
    elm = app.feElem
    elm.Get(each)
    if elm.topology == 0 or elm.topology == 1:
        elm.node[0], elm.node[1] = elm.node[1], elm.node[0]
    else:
        app.feAppMessage(fc.FCM_NORMAL, "Non-1D element skipped..."
