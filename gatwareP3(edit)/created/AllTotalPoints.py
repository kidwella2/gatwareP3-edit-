
import GatWareGetFromXL
from GatwareP3 import *
import GwareData
from TotalPoints import TotalPoints

from win32com.client import Dispatch

xl = Dispatch("Excel.Application")


class AllTotalPoints:
    """"""
    def __init__(self):
        xl.Workbooks("TimingSpace.xlsx").Worksheets("Sheet1").Activate()
        path = 'c:\\Python27\\'
        xl.Range("B3").Select()
        ShelveName = str(xl.ActiveCell.Value)
        filename = path + ShelveName
        GwareData.database = shelve.open(filename, writeback=False)
        GwareData.InstanceList = GwareData.database['InstanceList']
        DictlistI = [obj.__dict__ for obj in GwareData.InstanceList]
        namelist = []

        for obj in DictlistI:                                   # get all profiles' name
            namelist.append(obj['name'])
        print(namelist)
        xl.Range("B7").Select()
        total = xl.ActiveCell.Value
        TotalPoints.ShelveName = ShelveName                      # initialize self values on TotalPoints to avoid crash
        TotalPoints.DictlistI = DictlistI
        TotalPoints.newx = []
        TotalPoints.newy = []
        TotalPoints.profilename = ''
        for name in namelist:                                   # loop through all profiles
            xl.Range("D7").Select()             # set profile name and total points to excel
            xl.ActiveCell.Value = name
            xl.Range("B7").Select()
            xl.ActiveCell.Value = total
            GatWareGetFromXL.GetXLProfile()
            print('')
            TotalPoints.setTP(self=TotalPoints)
            TotalPoints.writeCSV(self=TotalPoints)
