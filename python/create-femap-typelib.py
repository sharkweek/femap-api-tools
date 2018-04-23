"""This code is used to convert the femap.tlb file into a python py file.
After running this function the command line will show the directory for where
this file is located. Take this file and copy it into the Lib folder of your
Python installation. To use this, you must have win32com installed. It may be
found at

Build < 221
https://sourceforge.net/projects/pywin32/files/pywin32/Build%20220/

Build >= 221
https://github.com/mhammond/pywin32

Make sure to choose the correct version of python. Alternatively, it can be
installed and kept up to date with your current version by using the
Anaconda distribution of Python, located here:

https://www.anaconda.com/download/
"""

import sys
from win32com.client import makepy
sys.argv = ["makepy", "-o Pyfemap.py", r"C:\<Femap Directory>\femap.tlb"]
makepy.main()
