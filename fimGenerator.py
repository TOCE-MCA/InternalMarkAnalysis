import json
import openpyxl
from openpyxl.styles.borders import Border, Side
import os.path as path
import os
import pandas as pd
# noinspection PyUnresolvedReferences
import pprint
from datetime import datetime
import shutil
from tkinter import filedialog as fd


def setFilePath():
    global fileDir, fileName, filePath
    filetypes = [
        ('Excel Files', '*.xlsx')
    ]
    filePath = fd.askopenfilename(title='Open internal marks excel file', initialdir=str(os.getenv("HOME")),
                                  filetypes=filetypes).replace("\\", "/")
    fileName = path.basename(filePath)
    fileDir = path.dirname(filePath)
    return 1


def main():
    try:
        setFilePath()
    except:
        print("Directory selection failed!\n")
        exit(1)