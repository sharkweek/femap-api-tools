import __main__
import pythoncom
import pyfemap  # the Python Femap API library
import sys

try:
    existObj = pythoncom.connect(pyfemap.model.CLSID) #Grabs active model

except:
    sys.exit("Femap not open") #Exits program if there is no active femap model

app = pyfemap.model(existObj)
fc = pyfemap.constants

if __name__ != '__main__':
    # Pass imported entities to __main__
    __main__.app = app
    print("Femap model imported as 'app'")
    __main__.fc = fc
    print("Femap constants imported as 'fc'")
    __main__.pf = pyfemap
    print("Femap API library imported as 'pf'")

# some useful shorcut functions
def pickEntities(entityType, clear=True, prompt="Select entities..."):
    '''Prompts user to select multiple entities and returns the IDs as an
    array'''

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

    return setArray

def pickID(entityType, prompt="Select entities..."):
    '''Prompts user to pick a single entity and returns the ID'''

    pickSet = app.feSet
    rc, id = pickSet.SelectID(entityType, prompt)

    return id
