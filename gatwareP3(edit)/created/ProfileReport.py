
import os.path

from GatwareP3 import *
import GwareData

from win32com.client import Dispatch

xl = Dispatch("Excel.Application")

xlToLeft = 1
xlToRight = 2
xlUp = 3
xlDown = 4
xlAscending = 1
xlYes = 1

class ProfileReport:
    """"""
    def __init__(self):
        self.GenerateReport()

    def GenerateReport(self):
        xl.Workbooks("TimingSpace.xlsx").Worksheets("Sheet1").Activate()
        path = 'c:\\Python27\\'
        xl.Range("B3").Select()
        ShelveName = str(xl.ActiveCell.Value)
        xlname = ShelveName + ".xlsx"
        xlnamelong = path + xlname
        filename = path + ShelveName
        if os.path.isfile('/Python27/' + xlname):  # generate workbook
            xl.Workbooks.Open(xlnamelong)
            xl.Workbooks(xlname).Worksheets("Sheet1").Activate()
        else:
            MakeNewWorkbook()
        GwareData.database = shelve.open(filename, writeback=False)
        GwareData.RawDataList = GwareData.database['RawDataList']
        RawDict = [obj.__dict__ for obj in GwareData.RawDataList]
        RawDict = sorted(RawDict, key=lambda i: i['name'])  # sort list of dict by name
        projName = ['Project Name:', ShelveName]
        initProfile = ['Profile Name:', '', 'Time', 'Displacement', 'Velocity', 'Acceleration']
        xl.Range("B2").Select()
        WriteRowFromSelected(projName)                      # write project name
        for i in range(len(RawDict)):
            xlspot = 5 + (8 * i)
            xl.Range("B" + str(xlspot)).Select()
            WriteColumnFromSelected(initProfile)                # write column headers for profile
            xl.Range("C" + str(xlspot)).Select()
            xl.ActiveCell.Value = RawDict[i]['name']            # write profile names
        numNodes = []
        for i in range(len(RawDict)):                           # get list containing # of nodes per profile
            templist = []
            temp = len(RawDict[i]['RawData'][0])
            for j in range(1, temp + 1):                        # name nodes for importing
                templist.append('Node ' + str(j))
            numNodes.append(templist)
        for i in range(len(RawDict)):                           # write nodes
            xlspot = 6 + (8 * i)
            xl.Range("C" + str(xlspot)).Select()
            WriteRowFromSelected(numNodes[i])
        self.WriteXLRowData(RawDict, 7, 0)                      # write times
        self.WriteXLRowData(RawDict, 8, 1)                      # write displcement
        self.WriteXLRowData(RawDict, 9, 2)                      # write velocity
        self.WriteXLRowData(RawDict, 10, 3)                      # write acceleration
        self.FormatBordersColors(numNodes)
        xl.ActiveWorkbook.SaveAs(xlnamelong)                    # save file



    def WriteXLRowData(self, RawDict, startLine, dataNum):              # writes rows of data for raw data
        for i in range(len(RawDict)):  # write nodes
            xlspot = startLine + (8 * i)
            xl.Range("C" + str(xlspot)).Select()
            WriteRowFromSelected(RawDict[i]['RawData'][dataNum])


    def FormatBordersColors(self, numNodes):                            # handles font size, borders, colors
        xl.Range("B2:C2").Font.Size = 22                                # title font size
        xl.Cells(2, 2).Interior.ColorIndex = 15                         # title background color
        xl.Cells(2, 3).Interior.ColorIndex = 4
        for i in range(1, 5):                                           # creates borders
            xl.Range("B2:C2").Borders(i).Weight = 4
            for j in range(len(numNodes)):
                xlspot = 5 + (8 * j)
                xl.Range(xl.Cells(xlspot, 2), xl.Cells(xlspot + 5, len(numNodes[j]) + 2)).Borders(i).Weight = 4
        for k in range(len(numNodes)):                                  # creates font size and background color
            xlspot = 5 + (8 * k)
            xl.Range(xl.Cells(xlspot, 2), xl.Cells(xlspot + 1, len(numNodes[k]) + 2)).Font.Size = 18  # font size
            xl.Range(xl.Cells(xlspot + 2, 2), xl.Cells(xlspot + 5, 2)).Font.Size = 18
            xl.Range(xl.Cells(xlspot + 2, 3), xl.Cells(xlspot + 5, len(numNodes[k]) + 2)).Font.Size = 16
            xl.Cells(xlspot, 2).Interior.ColorIndex = 15                                             # cell colors
            xl.Cells(xlspot, 3).Interior.ColorIndex = 43
            xl.Range(xl.Cells(xlspot + 2, 2), xl.Cells(xlspot + 2, len(numNodes[k]) + 2)).Interior.ColorIndex = 36
            xl.Range(xl.Cells(xlspot + 3, 2), xl.Cells(xlspot + 3, len(numNodes[k]) + 2)).Interior.ColorIndex = 35
            xl.Range(xl.Cells(xlspot + 4, 2), xl.Cells(xlspot + 4, len(numNodes[k]) + 2)).Interior.ColorIndex = 34
            xl.Range(xl.Cells(xlspot + 5, 2), xl.Cells(xlspot + 5, len(numNodes[k]) + 2)).Interior.ColorIndex = 38
            xl.Range(xl.Cells(xlspot + 1, 3), xl.Cells(xlspot + 1, len(numNodes[k]) + 2)).Interior.ColorIndex = 40
        xl.Columns.AutoFit()                                            # ensures data fits in excel boxes
