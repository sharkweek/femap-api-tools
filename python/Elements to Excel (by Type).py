import pythoncom
import Pyfemap
from Pyfemap import constants  # Sets femap constants to constants object so
# calling constants in constants.FCL_BLACK, not Pyfemap.constants.FCL_BLACK
import sys

import win32com.client as win32
win32.gencache.is_readonly = False
appExcel = win32.Dispatch(
    'Excel.Application')  # Calls and creates new excel application

try:
    existObj = pythoncom.connect(Pyfemap.model.CLSID)  # Grabs active model
    app = Pyfemap.model(existObj)
except:
    sys.exit(
        "femap not open")  # Exits program if there is no active femap model

rc = app.feAppMessage(0, "Python API Example Started")

workBook = appExcel.Workbooks.Add()
workSheet = workBook.Worksheets(1)

workSheet.Cells(1, 1).Value = "Element Type"
workSheet.Cells(1, 2).Value = "Element Count"

col = 2

pElemSet = app.feSet

Msg = "Model Element Summary"

app.feAppMessage(constants.FCL_BLACK, Msg)

string_vals = {1: "L_ROD elements",
               2: "L_BAR elements",
               3: "L_TUBE elements",
               4: "L_LINK elements",
               5: "L_BEAM elements",
               6: "L_SPRING elements",
               7: "L_DOF_SPRING elements",
               8: "L_CURVED_BEAM elements",
               9: "L_GAP elements",
               10: "L_PLOT elements",
               11: "L_SHEAR elements",
               12: "P_SHEAR elements",
               13: "L_MEMBRANE elements",
               14: "P_MEMBRANE elements",
               15: "L_BENDING elements",
               16: "P_BENDING elements",
               17: "L_PLATE elements",
               18: "P_PLATE elements",
               19: "L_PLANE_STRAIN elements",
               20: "P_PLANE_STRAIN elements",
               21: "L_LAMINATE_PLATE elements",
               22: "P_LAMINATE_PLATE elements",
               23: "L_AXISYM elements",
               24: "P_AXISYM elements",
               25: "L_SOLID elements",
               26: "P_SOLID elements",
               27: "L_MASS elements",
               28: "L_MASS_MATRIX elements",
               29: "L_RIGID elements",
               30: "L_STIFF_MATRIX elements",
               31: "L_CURVED_TUBE elements",
               32: "L_PLOT_PLATE elements",
               33: "L_SLIDE_LINE elements",
               34: "L_CONTACT elements",
               35: "L_AXISYM_SHELL elements",
               36: "P_AXISYM_SHELL elements",
               37: "P_BEAM elements",
               38: "L_WELD elements",
               39: "L_SOLID_LAMINATE elements",
               40: "P_SOLID_LAMINATE elements",
               41: "L_SPRING_to_GROUND elements",
               42: "L_DOF_SRPING_to_GROUND elements}"}

for val in string_vals:
    pElemSet.clear()
    pElemSet.AddRule(val, constants.FGD_ELEM_BYTYPE)

    if pElemSet.Count() > 0:
        workSheet.Cells(col, 2).Value = str(pElemSet.Count())
        workSheet.Cells(col, 1).Value = string_vals[val]
        col = col + 1

appExcel.Visible = True
rc = app.feAppMessage(0, "Python API Example Finished")
