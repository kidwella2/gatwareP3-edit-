import sys
from PyQt5 import QtGui, uic, QtCore, QtWidgets
from PyQt5.QtWidgets import QFileDialog, QMessageBox
#from matplotlib.backends.qt_compat import QtCore, QtWidgets
from PyQt5 import *
import shelve
import os
import collections
import numpy as np
import pdb
import csv
#pdb.set_trace()
import ShiftTime
import XaxisBasis
import Npoints
import NodeSteps
import TotalPoints
from GatWareGetFromXL import *
from GatWareProfilePlot import *
from CurrentProfilePanel import *
from GatWareAndyBlender import *
from AllTotalPoints import AllTotalPoints
from ProfileReport import ProfileReport
from uniquer import UniqueSeq
# import GatwareCompile
# import MachineTimingMethodsV3 as MTM
import GwareData

import QtSecondWindowTest
import Isaac
import blendDialog


from win32com.client import Dispatch
xl = Dispatch("Excel.Application")

qtCreatorFile = "GatWare.ui"  # Enter file here.

Ui_MainWindow, QtBaseClass = uic.loadUiType(qtCreatorFile)


class MyApp(QtWidgets.QMainWindow, Ui_MainWindow):

    def __init__(self):
        QtWidgets.QMainWindow.__init__(self)
        Ui_MainWindow.__init__(self)
        self.setupUi(self)

    def activateShelve(self):
        '''
            Called By various processes whenever the database needs to be activated
            for editing.

            Opens the database file referenced by cell B3 in the Excel
            worksheet

            Writeback is True for updating-----
            '''

        xl.Workbooks("TimingSpace.xlsx").Worksheets("Sheet1").Activate()
        path = 'c:\\Python27\\'
        xl.Range("B3").Select()
        ShelveName = str(xl.ActiveCell.Value)
        #suffix = '.db'
        filename = path + ShelveName
        GwareData.database = shelve.open(filename, writeback=True)
        GwareData.InstanceList = GwareData.database['InstanceList']
        GwareData.RawDataList = GwareData.database['RawDataList']

    def getexistingshelve(self):
        '''
        Called from "Load Database" button

        Opens the database file referenced by cell B3 in the Excel
        worksheet

        Not writeback enabled----
        '''
        # print "hello from machine timing methods-->getexistingshelve()"
        #global database, RepeatData, ShelveName
        xl.Workbooks("TimingSpace.xlsx").Worksheets("Sheet1").Activate()
        path = 'c:\\Python27\\'
        xl.Range("B3").Select()
        ShelveName = str(xl.ActiveCell.Value)
        # suffix = '.db'
        filename = path + ShelveName
        GwareData.database = shelve.open(filename, writeback=False)
        GwareData.InstanceList = GwareData.database['InstanceList']
        GwareData.RawDataList = GwareData.database['RawDataList']

    def Indexing(self):
        try:
            import IndexingCollector
            self.window=IndexingCollector.IndexingDialog()
            self.window.show()
        except Exception as e:
            print('Editing Excel')
            print('Catch: ', e.__class__)
            return 1

    def MoveDesigner(self):
        import MoveDesigner
        self.window = MoveDesigner.MatplotlibWidget()
        self.window.show()

    def GetXLProfile(self):
        try:
            GetXLProfile()
        except Exception as e:
            print('Editing Excel')
            print('Catch: ', e.__class__)
            return 1

    def Plot(self):
        '''

            :return:
            '''
        # global Taxis, DispAxis, VelAxis, AccAxis, RawTimes, RawDisplacement
        # pdb.set_trace()

        try:
            GetXLProfile()
        except Exception as e:
            print('Editing Excel')
            print('Catch: ', e.__class__)
            return 1
        plt.close()

        xl = Dispatch("Excel.Application")
        Taxis = []
        DispAxis = []
        VelAxis = []
        AccAxis = []

        segname()
        rcParams['ytick.labelsize'] = 12
        rcParams['xtick.labelsize'] = 12
        rcParams['font.size'] = 12

        for key in GwareData.motionparams.keys():

            if GwareData.segmentname in key:
                obj = GwareData.motionparams[key]
                for t in obj.trange:
                    Taxis.append(t)
                for x in obj.xplot:
                    DispAxis.append(x)
                for v in obj.vplot:
                    VelAxis.append(v)
                for a in obj.aplot:
                    AccAxis.append(a)

        xl.Workbooks("TimingSpace.xlsx").Worksheets("Sheet1").Activate()
        Cell = xl.Range("D9").Activate()

        if Cell is not None:
            # global RawTimes
            xl.Range(xl.Selection, xl.Selection.End(GwareData.xlToRight)).Select()
            inputrange = xl.Selection
            obj = inputrange.Columns()
            GwareData.RawTimes = [item for item in obj[0]]  # print RawTimes

        Cell = xl.Range("D10").Activate()

        if Cell is not None:
            xl.Range(xl.Selection, xl.Selection.End(GwareData.xlToRight)).Select()
            inputrange = xl.Selection
            GwareData.RawDisplacement = [time for time in inputrange]

        # Start Plotting

        fig = plt.figure()  # 11/8/11

        ax1 = plt.subplot(3, 1, 1)
        plt.plot(Taxis, DispAxis, label='%s' % 'Displacement', lw=1, color='g')
        ax1.set_ylabel('Displacement', color='g')
        plt.grid()
        plt.autoscale(enable=True, axis='both', tight=None)
        # plt.legend(loc='best')

        ax2 = plt.subplot(3, 1, 2)
        plt.plot(Taxis, VelAxis, label='%s' % 'Velocity', lw=1, color='b')
        ax2.set_ylabel('Velocity', color='b')
        plt.grid()
        plt.autoscale(enable=True, axis='both', tight=None)
        # plt.legend(loc='best')

        ax3 = plt.subplot(3, 1, 3, sharex=ax1)
        plt.plot(Taxis, AccAxis, label='%s' % 'Acceleration', lw=1, color='r')
        ax3.set_ylabel('Acceleration', color='r')
        plt.autoscale(enable=True, axis='both', tight=None)
        # plt.legend(loc='best')

        for eachvline in GwareData.RawTimes:
            plt.axvline(eachvline, color='r', ls='dashed', lw=1)
        plt.grid(b='on', which='both')
        plt.autoscale(enable=True, axis='both', tight=None)

        def on_key(event):
            sys.stdout.flush()
            ix = event.xdata

            if event.inaxes in [ax3]:
                # print('you pressed', event.key, event.xdata, event.ydata)

                # print('button=%d, x=%d, y=%d, xdata=%f, ydata=%f' % (event.button, event.x, event.y, event.xdata, event.ydata))

                def gohome():
                    xl.Range("D9").Activate()

                def writeValue(acc):
                    tdiff = [abs(event.xdata - i) for i in GwareData.RawTimes]  # look for a small difference
                    m = dict(zip(GwareData.RawTimes, tdiff))
                    mintime = min(m, key=m.get)  # get the time value corresponding to the closest mouse click
                    gohome()
                    indexvalue = (GwareData.RawTimes.index(
                        mintime) + 1)  # get index position for the min time value in timerange
                    xl.ActiveCell.Offset(4, indexvalue).Select()
                    xl.ActiveCell.Value = acc

                if event.key == 'shift':
                    # print("controlkey")
                    acc = Optimize(event.xdata, 120)
                    writeValue(acc)
                    Plot()
                    # return acc
                    # print acc
                    # print "control"

                if event.key == 'alt':
                    acc = event.ydata
                    writeValue(acc)
                    drawn = 1
                    Plot()
                    # Acc=Optimize(event.xdata,120)
                    # print Acc

        cid = fig.canvas.mpl_connect('button_press_event', on_key)

        plt.subplots_adjust(left=0.075, bottom=.05, right=0.95, top=.95, wspace=None, hspace=None)

        segname()

        plt.suptitle(str(GwareData.segmentname))

        PlottingTest.FollowDotCursor(ax1, Taxis, DispAxis)
        PlottingTest.FollowDotCursor(ax2, Taxis, VelAxis)
        PlottingTest.FollowDotCursor(ax3, Taxis, AccAxis)

        # plt.draw() # commented out 11/8
        mng = plt.get_current_fig_manager()
        # mng.window.state('zoomed')
        plt.tight_layout()

        show()  # was plt.show()
        # Plot()

    def CompileAndSave(self):
        '''
            Adds the Current Profiles' Raw Data Values and Calculated X,V & A
            Data to the working memory. These Values are NOT saved until the
            'Store to Shelve' operation is performed.
            ---> Calls the Profile Class
            ---> Calls the GetRawData Class

            @param ProfileName: Name of Active profile from Excel ("D7")
            @type ProfileName: string value
            '''

        # global RawDataList, InstanceList
        namelist = []  # only used for the active profile declared in the Excel sheet
        try:
            xl.Workbooks("TimingSpace.xlsx").Worksheets("Sheet1").Activate()
            xl.Range("D7").Select()  # get profile name
        except Exception as e:
            print('Editing Excel')
            print('Catch: ', e.__class__)
            return 1
        ProfileName = str(xl.ActiveCell.Value)
        namelist.append(ProfileName)
        # Check to see if the local project file exists
        xl.Range("B3").Select()
        path = 'C:\\Python27\\'
        ShelveName = str(xl.ActiveCell.Value)
        suffix = '.dat'
        fullfilename = path + ShelveName + suffix
        filename = path + ShelveName

        if os.path.exists(fullfilename):
            '''If local file exist, open the file and instantiate 'InstanceList and RawDataList'
            '''
            database = shelve.open(filename, writeback=True, protocol=2)
            GwareData.InstanceList = database['InstanceList']
            GwareData.RawDataList = database['RawDataList']

        # print "From Compile/Save process/n"
        # print RawDataList
        # print InstanceList

        if not os.path.exists(fullfilename):
            '''If the database file does not exist already, create it and append the raw data and the Profile
            information to the empty list created at the beginning of this file.
            '''
            database = shelve.open(filename, protocol=2)
            GwareData.InstanceList.append(Profile(ProfileName))
            GwareData.RawDataList.append(GetRawData(ProfileName))

        existingrawdatanames = [obj.name for obj in GwareData.RawDataList]
        existinginstancelistnames = [obj.name for obj in GwareData.InstanceList]

        if not ProfileName in existingrawdatanames:  # On new name, Add to both
            Cell = xl.Range("D9").Select()
            for item in namelist:
                GwareData.InstanceList.append(Profile(ProfileName))
                if Cell is not None:
                    GwareData.RawDataList.append(GetRawData(ProfileName))

        if ProfileName in existingrawdatanames:  # if exists, delete and replace
            poplocale = existingrawdatanames.index(ProfileName)
            GwareData.RawDataList.pop(poplocale)
            GwareData.RawDataList.append(GetRawData(ProfileName))

        if ProfileName in existinginstancelistnames:  # if exists, delete and replace
            poplocale = existinginstancelistnames.index(ProfileName)
            GwareData.InstanceList.pop(poplocale)
            GwareData.InstanceList.append(Profile(ProfileName))

        database['InstanceList'] = GwareData.InstanceList
        database['RawDataList'] = GwareData.RawDataList
        database.sync()  # Added 8/25/16
        database.close()

        # print("Compile")

    def ClearActiveProfile(self):
        try:
            xl.Workbooks("TimingSpace.xlsx").Worksheets("Sheet1").Activate()
            xl.Range("D9").Select()
        except Exception as e:
            print('Editing Excel')
            print('Catch: ', e.__class__)
            return 1
        # sheet=xl.Worksheets("Sheet1").Activate
        # DeclareXLDirections()
        contents = xl.Selection()

        if not contents == None:  # if something there, clear range
            xl.Range(xl.Selection, xl.Selection.End(xlToRight)).Select()
            xl.Range(xl.Selection, xl.Selection.End(xlDown)).Select()
            range = xl.Selection
            range.ClearContents()

        xl.Range("D7").Select()
        xl.Selection.ClearContents()

        xl.Range("A17").Select()
        contents = xl.Selection()

        if not contents == None:  # if something there, clear range
            xl.Range(xl.Selection, xl.Selection.End(xlDown)).Select()
            range = xl.Selection
            range.ClearContents()

        xl.Range("C17").Select()
        contents = xl.Selection()

        if not contents == None:  # if something there, clear range
            xl.Range(xl.Selection, xl.Selection.End(xlDown)).Select()
            xl.Range(xl.Selection, xl.Selection(1, 2)).Select()
            range = xl.Selection
            range.ClearContents()

        # print("ClearProfile")

    def StartingZeros(self):
        '''Enter Zero values into Excel worksheet to begin new profile'''
        try:
            xl.Workbooks("TimingSpace.xlsx").Worksheets("Sheet1").Activate()
            DeclareXLDirections()
        except Exception as e:
            print('Editing Excel')
            print('Catch: ', e.__class__)
            return 1
        xl.Range("D9").Activate()
        cellvalue = (xl.ActiveCell.Value)

        if not isinstance(cellvalue, float):
            xl.ActiveCell.Value = 0
            for cell in range(3):
                xl.ActiveCell.Offset(2, 1).Select()
                xl.ActiveCell.Value = 0

        if isinstance(cellvalue, float):
            QMessageBox.about(self, "Information", "Data Already present")
        print("StartingZeros")

    def WriteXLTables(self):
        """ For the Active profile in the worksheet, Writes Displacement,
        Velocity & Acceleration vs. time into a seperate Excel
        workbook. Each dataset is written to a different sheet"""
        Taxis = []
        DispAxis = []
        VelAxis = []
        AccAxis = []
        try:
            segname()
        except Exception as e:
            print('Editing Excel')
            print('Catch: ', e.__class__)
            return 1

        for key in GwareData.motionparams.keys():
            if GwareData.segmentname in key:
                obj = GwareData.motionparams[key]
                for t in obj.trange:
                    Taxis.append(t)
                for x in obj.xplot:
                    DispAxis.append(x)
                for v in obj.vplot:
                    VelAxis.append(v)
                for a in obj.aplot:
                    AccAxis.append(a)

        from win32com.client import Dispatch
        xl = Dispatch("Excel.Application")
        xl.Visible = 1
        wb = xl.Workbooks.Add()
        xl.Sheets("Sheet1").Name = "Displacement"
        xl.Worksheets.Add()
        xl.Sheets("Sheet2").Name = "Velocity"
        xl.Worksheets.Add()
        xl.Sheets("Sheet3").Name = "Acceleration"

        # ws = wb.Worksheets("Sheet2")

        def lineprinter(seq):
            for obj in seq:
                print(obj)

        tx = zip(Taxis, DispAxis)

        # seq1=UniqueSeq(Taxis,DispAxis)
        # TX=seq1.seqlist

        # tx=set(Tx)
        TX = list(tx)
        TX.sort()

        DispTime = [obj[0] for obj in TX]
        DispAxis = [obj[1] for obj in TX]

        tv = zip(Taxis, VelAxis)
        # tv=set(Tv)
        TV = list(tv)
        TV.sort()

        VelTime = [obj[0] for obj in TV]
        VelAxis = [obj[1] for obj in TV]

        ta = zip(Taxis, AccAxis)
        # ta=set(Ta)
        TA = list(ta)
        TA.sort()

        AccTime = [obj[0] for obj in TA]
        AccAxis = [obj[1] for obj in TA]

        def WriteTime(seq):
            for entry in seq:
                xl.ActiveCell.Value = entry
                xl.ActiveCell.Offset(2, 1).Select()

        def WriteDispl():
            xl.Worksheets("Displacement").Activate()
            WriteTime(DispTime)
            xl.Range("B1").Activate()
            for entry in DispAxis:
                xl.ActiveCell.Value = entry
                xl.ActiveCell.Offset(2, 1).Select()

        def WriteVel():
            xl.Worksheets("Velocity").Activate()
            WriteTime(VelTime)
            xl.Range("B1").Activate()
            for entry in VelAxis:
                xl.ActiveCell.Value = entry
                xl.ActiveCell.Offset(2, 1).Select()

        def WriteAccel():
            xl.Worksheets("Acceleration").Activate()

            WriteTime(AccTime)
            xl.Range("B1").Activate()
            for entry in AccAxis:
                xl.ActiveCell.Value = entry
                xl.ActiveCell.Offset(2, 1).Select()

        WriteAccel()
        WriteVel()
        WriteDispl()

    def WriteVelocity(self):
        """ For the Active profile in the worksheet, Writes Displacement,
                Velocity & Acceleration vs. time into a seperate Excel
                workbook. Each dataset is written to a different sheet"""
        Taxis = []
        VelAxis = []
        try:
            segname()
        except Exception as e:
            print('Editing Excel')
            print('Catch: ', e.__class__)
            return 1

        for key in GwareData.motionparams.keys():
            if GwareData.segmentname in key:
                obj = GwareData.motionparams[key]
                for t in obj.trange:
                    Taxis.append(t)
                for v in obj.vplot:
                    VelAxis.append(v)

        from win32com.client import Dispatch
        xl = Dispatch("Excel.Application")
        wb = xl.Workbooks.Add()
        xl.Sheets("Sheet1").Name = "Velocity"

        tv = zip(Taxis, VelAxis)

        TV = list(tv)
        TV.sort()

        VelTime = [obj[0] for obj in TV]
        VelAxis = [obj[1] for obj in TV]

        def WriteTime(seq):
            for entry in seq:
                xl.ActiveCell.Value = entry
                xl.ActiveCell.Offset(2, 1).Select()

        def WriteVel():
            xl.Worksheets("Velocity").Activate()
            WriteTime(VelTime)
            xl.Range("B1").Activate()
            for entry in VelAxis:
                xl.ActiveCell.Value = entry
                xl.ActiveCell.Offset(2, 1).Select()

        WriteVel()

    def ConsultIsaac(self):
        self.window = Isaac.IsaacBox()
        self.window.show()


    def PortFromGoalSeek(self):
        GwareData.variableList={}

        try:
            xl.Workbooks("TimingSpace.xlsx").Worksheets("Move Worksheet").Activate()
            xl.Range("B9").Select()
        except Exception as e:
            print('Editing Excel')
            print('Catch: ', e.__class__)
            return 1
        Xcv = xl.Selection()
        #print(Xcv)
        DeclareXLDirections()

        def isempty(contents):
            try:
                if isinstance(contents / 0.5, float):
                    GwareData.Datapresent = 1
            except TypeError:
                GwareData.Datapresent = 0
            #print('GwareData from isempty= ', GwareData.Datapresent)

        def EnterValue(RCstart, value):
            xl.Workbooks("TimingSpace.xlsx").Worksheets("Sheet1").Activate()
            xl.Range(RCstart).Select()
            cellcontents = xl.Selection()
            #print(type(cellcontents))
            isempty(cellcontents)

            if not GwareData.Datapresent:
                'If the cell is empty, enter the value'
                xl.ActiveCell.Value = value

            if GwareData.Datapresent:
                # ===============================================================
                # If the starting cell has data in it, check to the right....
                # ===============================================================
                xl.ActiveCell.Offset(1, 2).Select()
                cellcontents = xl.Selection()
                # print type(cellcontents)
                # ===============================================================
                # Check active cell to see if data exists here
                # ===============================================================

                isempty(cellcontents)

                if not GwareData.Datapresent:
                    xl.ActiveCell.Value = value

                if GwareData.Datapresent:
                    # ===========================================================
                    # If active cell has data, go to end of row, and enter value there
                    # ===========================================================

                    xl.Range(RCstart).Select()
                    xl.Selection.End(xlToRight).Select()
                    endingvalue = xl.Selection()
                    xl.ActiveCell.Offset(1, 2).Select()
                    xl.ActiveCell.Value = value + endingvalue
            xl.Workbooks("TimingSpace.xlsx").Worksheets("Move Worksheet").Activate()

        if Xcv <= .001:
            # If the constant velocity distance is about 0, the move is triangular--> enter time and displacement in sheet1
            xl.Range("B16").Select()
            Time = xl.Selection()
            xl.Range("B15").Select()
            Displacement = xl.Selection()

            EnterValue("D9", Time)
            EnterValue("D10", Displacement)
            EnterValue("D11", 0)
            EnterValue("D12", 0)

        if Xcv > .001:
            xl.Range("B8").Select()
            t1 = xl.Selection()
            # Get the first Time value (t1) and enter in sheet1
            EnterValue("D9", t1)

            xl.Range("B7").Select()
            X1 = xl.Selection()
            # Get the first Distance Value (X1) and enter in sheet1
            EnterValue("D10", X1)

            xl.Range("F6").Select()
            # Get max velocity
            Velocity = xl.Selection()
            EnterValue("D11", Velocity)
            EnterValue("D12", 0)

            # Completion of first column value

            xl.Range("B10").Select()
            # Get constant velocity time (tcv) and enter in sheet1
            tcv = xl.Selection()
            EnterValue("D9", tcv)

            EnterValue("D10", Xcv)  # enter constant velocity displacement (previously attained) into sheet1
            xl.Workbooks("TimingSpace.xlsx").Worksheets("Sheet1").Activate()
            xl.Range("D11").Select()

            xl.Selection.End(xlToRight).Select()
            xl.ActiveCell.Offset(1, 2).Select()
            xl.ActiveCell.Value = Velocity
            EnterValue("D12", 0)

            xl.Workbooks("TimingSpace.xlsx").Worksheets("Move Worksheet").Activate()

            xl.Range("B14").Select()
            t2 = xl.Selection()
            EnterValue("D9", t2)

            xl.Range("B13").Select()
            X2 = xl.Selection()
            EnterValue("D10", X2)

            EnterValue("D11", -Velocity)
            EnterValue("D12", 0)

    def SelectWorkingDirectory(self):
        _OutputFolder = str(QFileDialog.getExistingDirectory(self, "Select Directory"))
        self.lineEdit.setText(_OutputFolder)

    def WriteCSV(self):
        path=self.lineEdit.text()
        if not path:
            path='c:/Python27'
        try:
            xl.Workbooks("TimingSpace.xlsx").Worksheets("Sheet1").Activate()
            select('b3')
        except Exception as e:
            print('Editing Excel')
            print('Catch: ', e.__class__)
            return 1
        projname = xl.Selection()
        self.getexistingshelve()

        for thisobj in GwareData.InstanceList:
            global time, disp, name
            #pdb.set_trace()
            time = thisobj.Taxis
            disp = thisobj.DispAxis
            name = thisobj.name
            seq = UniqueSeq(time, disp)  # branches to make list of non-duplicated values
            UniqueList = seq.seqlist
            #path = 'c:\\Python27\\'
            file = path + '/' +projname + '_' + name + '.csv'
            print(file)
            with open(file, 'w', newline='') as f:
                c = csv.writer(f)
                for row in UniqueList:
                    c.writerow(row)
            f.close()

    def WriteVelCSV(self):
        path = self.lineEdit.text()
        if not path:
            path = 'c:/Python27'
        try:
            xl.Workbooks("TimingSpace.xlsx").Worksheets("Sheet1").Activate()
            select('B3')
        except Exception as e:
            print('Editing Excel')
            print('Catch: ', e.__class__)
            return 1
        projname = xl.Selection()
        self.getexistingshelve()

        for thisobj in GwareData.InstanceList:
            global time, vel, name
            # pdb.set_trace()
            time = thisobj.Taxis
            vel = thisobj.VelAxis
            name = thisobj.name
            seq = UniqueSeq(time, vel)  # branches to make list of non-duplicated values
            UniqueList = seq.seqlist
            # path = 'c:\\Python27\\'
            file = path + '/' + projname + '_' + name + 'Vel.csv'
            print(file)
            with open(file, 'w', newline='') as f:
                c = csv.writer(f)
                for row in UniqueList:
                    c.writerow(row)
            f.close()

    def AppendSelected(self):
        from XL_Initializer import SelectEndRight
        try:
            Values = xl.Selection()
        except Exception as e:
            print('Editing Excel')
            print('Catch: ', e.__class__)
            return 1
        try:
            xl.ActiveCell.Offset(1, 0).Select() # go to the left one cell and get that value
            seedtime = xl.ActiveCell.Value
            timelist = list(Values[0])
            timelist.insert(0, seedtime)
            difftimes = np.diff(timelist)
            #print(difftimes)
            appendtimes = [a + max(timelist) for a in difftimes]
            #print(appendtimes)
        except Exception as e:                                           # exception handling for no value selected
            print('Out of Bounds')
            print('Catch: ', e.__class__)
            return 1
            #print('Catch: ', sys.exc_info()[0])


        def WriteSequence(RowStartCell, Seq, truthvalue):
            flag = truthvalue
            for obj in Seq:
                SelectEndRight(RowStartCell)
                if flag:
                    V = xl.ActiveCell()
                if not flag:
                    V = 0
                xl.ActiveCell.Offset(1, 2).Select()
                xl.ActiveCell.Value = obj + V

        WriteSequence("D9", difftimes, 1)

        Disp = list(Values[1])
        Vel = list(Values[2])
        Acc = list(Values[3])

        WriteSequence("D10", Disp, 0)
        WriteSequence("D11", Vel, 0)
        WriteSequence("D12", Acc, 0)
        #print("AppendSelectedData")


    def CumulativeAppend(self):
        from XL_Initializer import SelectEndRight
        try:
            Values = xl.Selection()
        except Exception as e:
            print('Editing Excel')
            print('Catch: ', e.__class__)
            return 1
        try:
            xl.ActiveCell.Offset(1, 0).Select()  # go to the left one cell and get that value
            seedtime = xl.ActiveCell.Value
            timelist = list(Values[0])
            timelist.insert(0, seedtime)
            difftimes = np.diff(timelist)
            # print(difftimes)
            appendtimes = [a + max(timelist) for a in difftimes]
            # print(appendtimes)
        except Exception as e:  # exception handling for no value selected
            print('Out of Bounds')
            print('Catch: ', e.__class__)
            return 1
            # print('Catch: ', sys.exc_info()[0])

        def WriteSequence(RowStartCell, Seq, truthvalue):
            flag = truthvalue
            for obj in Seq:
                SelectEndRight(RowStartCell)
                if flag:
                    V = xl.ActiveCell()
                if not flag:
                    V = 0
                xl.ActiveCell.Offset(1, 2).Select()
                xl.ActiveCell.Value = obj + V

        WriteSequence("D9", difftimes, 1)

        Disp = list(Values[1])
        Disp=[max(Disp) + i for i in Disp]
        Vel = list(Values[2])
        Acc = list(Values[3])

        WriteSequence("D10", Disp, 0)
        WriteSequence("D11", Vel, 0)
        WriteSequence("D12", Acc, 0)
        # print("AppendSelectedData")


    def LoadProfilesToListBox(self):
        # Go get a list of profile names from the database and recurse through below
        self.listWidget.clear()
        try:
            xl.Workbooks("TimingSpace.xlsx").Worksheets("Sheet1").Activate()
            path = 'c:\\Python27\\'
            xl.Range("B3").Select()
        except Exception as e:
            print('Editing Excel')
            print('Catch: ', e.__class__)
            return 1
        ShelveName = str(xl.ActiveCell.Value)
        # suffix = '.db'
        filename = path + ShelveName
        GwareData.database = shelve.open(filename, writeback=False)
        GwareData.InstanceList = GwareData.database['InstanceList']
        GwareData.RawDataList = GwareData.database['RawDataList']

        pos = 0

        for item in GwareData.InstanceList:
            # print(item.name)
            self.listWidget.insertItem(pos, item.name)
            pos += 1
            # self.listWidget.addItem(item.name)


    def ListBoxText(self):
        #
        item = self.listWidget.currentItem()
        print(item.text())
        profile = item.text()
        self.lineEdit_SelectedProfile.setText(str(profile))

    def DeleteProfile(self):
        '''Activates (writeback=true) shelve and deletes items in
                selectionlist'''

        selectionList=[]

        try:
            targetProfile = str(self.lineEdit_SelectedProfile.text())
            targetProfileExists = 1
        except:
            NoProfile = 0

        #print(targetProfile)

        if targetProfileExists:
            selectionList.append(targetProfile)

        self.activateShelve()
        # global database

        if targetProfileExists:

            for obj in GwareData.InstanceList:
                if obj.name == targetProfile:
                    GwareData.InstanceList.pop(GwareData.InstanceList.index(obj))

            for obj in GwareData.RawDataList:
                if obj.name == targetProfile:
                    GwareData.RawDataList.pop(GwareData.RawDataList.index(obj))

            GwareData.database['InstanceList'] = GwareData.InstanceList
            GwareData.database['RawDataList'] = GwareData.RawDataList
            GwareData.database.sync()
            GwareData.database.close()


    def EditProfile(self):

        targetProfile = str(self.lineEdit_SelectedProfile.text())
        targetProfileExists = 1

        #print(targetProfile)

        if targetProfileExists:
            #print("in targetprofile exist")
            #print(GwareData.item)
            # global selectionlist
            Profile = None
            # NoNoList=['ConstantMotion']

            self.getexistingshelve()
            xl.Range("D9").Select()
            cellvalue = xl.Selection()

            if isinstance(cellvalue, float):
                # If something is in cell D9, clear the entire range
                xl.Range(xl.Selection, xl.Selection.End(GwareData.xlToRight)).Select()
                xl.Range(xl.Selection, xl.Selection.End(GwareData.xlDown)).Select()
                range = xl.Selection
                range.ClearContents()

            #target = GwareData.item  # value comes from the list created by the 'CheckListBox1....method'

            ObjList = []

            try:
                NameList = [i.name for i in GwareData.InstanceList]
                position = NameList.index(targetProfile)
                objectFit = GwareData.InstanceList[position].Fit
            except:
                print('No/Wrong Value')
                print('Catch: ', e.__class__)
                return 1

            # All Classes should have a Rawdata variable, exclusion if length of RawData object is 1.

            for obj in GwareData.RawDataList:
                # Put every RawDataList object in ObjList if RawData has more than one Value
                item = obj.__dict__
                if not len(item['RawData']) == 1:
                    ObjList.append(obj)

            for obj in ObjList:
                # If Target in list & Rawdata (from previous)
                if obj.name == targetProfile:
                    if not objectFit[0] in GwareData.NoNoList:
                        Profile = obj

            if Profile:
                Profilename = Profile.name
                data = Profile.RawData
                time = data[0]
                position = data[1]
                velocity = data[2]
                acceleration = data[3]
                xl.Range("D7").Select()  # put name in worksheet

                xl.ActiveCell.Value = Profilename
                xl.Range("D9").Select()  # upper left column of range

                for value in time:
                    # populate time values
                    xl.ActiveCell.Value = value
                    xl.ActiveCell.Offset(1, 2).Select()

                xl.Range("D10").Select()
                for pos in position:  # populate position values
                    xl.ActiveCell.Value = pos
                    xl.ActiveCell.Offset(1, 2).Select()

                xl.Range("D11").Select()
                for vel in velocity:  # populate velocity values
                    xl.ActiveCell.Value = vel
                    xl.ActiveCell.Offset(1, 2).Select()

                xl.Range("D12").Select()
                for acc in acceleration:  # populate acceleration values
                    xl.ActiveCell.Value = acc
                    xl.ActiveCell.Offset(1, 2).Select()

    def Plotting(self):
        try:
            self.window = QtSecondWindowTest.PlotAll()
            self.window.show()
        except Exception as e:
            print('Editing Excel')
            print('Catch: ', e.__class__)
            return 1

    def OpenDatabase(self):
        import string
        name = QFileDialog.getOpenFileName(self, 'Open File')
        name = name[0]
        name = name.translate({ord(ch): ' ' for ch in '/.:'})
        name = name.split()
        try:                                                        # catch if no database selected
            name = name[-2]
        except Exception as e:
            print('Database doesn\'t exist')
            print('Catch: ', e.__class__)
            return 1
        try:
            select('b3')
            xl.Selection.Value = name
        except Exception as e:
            print('Editing Excel')
            print('Catch: ', e.__class__)
            return 1
        #print(name)

    def MergeDatabases(self):
        try:
            xl.Workbooks("TimingSpace.xlsx").Worksheets("Sheet1").Activate()
            path = 'c:\\Python27\\'
            xl.Range("B3").Select()                                     # first database is the current one
        except Exception as e:
            print('Editing Excel')
            print('Catch: ', e.__class__)
            return 1
        ShelveName = str(xl.ActiveCell.Value)
        filename = path + ShelveName
        GwareData.database = shelve.open(filename, writeback=False)
        GwareData.InstanceList = GwareData.database['InstanceList']
        GwareData.RawDataList = GwareData.database['RawDataList']
        DictlistI = [obj.__dict__ for obj in GwareData.InstanceList]
        DictlistR = [obj.__dict__ for obj in GwareData.RawDataList]
        DictlistI = sorted(DictlistI, key=lambda i: i['name'])  # sort list of dict by name
        DictlistR = sorted(DictlistR, key=lambda i: i['name'])  # sort list of dict by name

        name = QFileDialog.getOpenFileName(self, 'Open File')         # locate a database to merge with the current one
        name = name[0]
        name = name.translate({ord(ch): ' ' for ch in '/.:'})
        name = name.split()
        try:  # catch if no database selected
            name = name[-2]
        except Exception as e:
            print('Database doesn\'t exist')
            print('Catch: ', e.__class__)
            return 1
        select('B3')
        xl.Selection.Value = name                                   # open the second database
        filename = path + name
        GwareData.database = shelve.open(filename, writeback=False)
        InstList = GwareData.database['InstanceList']
        RawList = GwareData.database['RawDataList']
        DictlistI2 = [obj.__dict__ for obj in InstList]
        DictlistR2 = [obj.__dict__ for obj in RawList]
        DictlistI2 = sorted(DictlistI2, key=lambda i: i['name'])  # sort list of dict by name
        DictlistR2 = sorted(DictlistR2, key=lambda i: i['name'])  # sort list of dict by name

        for obj in DictlistR2:                                  # alter names to avoid duplicates
            obj['name'] = obj['name'] + '_' + name
            DictlistR.append(obj)
        for obj in DictlistI2:
            obj['name'] = obj['name'] + '_' + name
            DictlistI.append(obj)

        xl.Range("B3").Select()                         # new database to store combined databases
        xl.ActiveCell.Value = ShelveName + '_' + name
        filename = path + ShelveName + '_' + name
        GwareData.database = shelve.open(filename, writeback=True, protocol=2)
        for i in range(len(RawList)):
            GwareData.InstanceList.append(InstList[i])
            GwareData.RawDataList.append(RawList[i])

        existingrawdatanames = [obj['name'] for obj in DictlistR]
        existinginstancelistnames = [obj['name'] for obj in DictlistI]
        print(existingrawdatanames)

        for obj in DictlistR:
            if obj['name'] in existingrawdatanames:  # if exists, delete and replace
                poplocale = existingrawdatanames.index(obj['name'])
                GwareData.RawDataList.pop(poplocale)
                GwareData.RawDataList.append(GetRawData(obj['name']))

            if obj['name'] in existinginstancelistnames:  # if exists, delete and replace
                poplocale = existinginstancelistnames.index(obj['name'])
                GwareData.InstanceList.pop(poplocale)
                GwareData.InstanceList.append(Profile(obj['name']))

        for i in range(len(DictlistR)):                                 # update info
            GwareData.InstanceList[i].__dict__ = DictlistI[i]
            GwareData.RawDataList[i].__dict__ = DictlistR[i]
        GwareData.database['InstanceList'] = GwareData.InstanceList             # update database
        GwareData.database['RawDataList'] = GwareData.RawDataList
        GwareData.database.sync()
        GwareData.database.close()

    def GatwareAndyBlender(self):                                           # accel blend
        self.window = blendDialog.blendDialog()
        self.window.show()
        #GatWareAndyBlender()

    def ShiftProfileTime(self):
        try:
            self.window = ShiftTime.ShiftTime()                                 # shift profiles' time via excel
            self.window.show()
        except Exception as e:
            print('Editing Excel')
            print('Catch: ', e.__class__)
            return 1

    def XaxisBasis(self):                                                # time/degrees for x-axis (dist entered in deg)
        try:
            self.window = XaxisBasis.XaxisBasis()
            self.window.show()
        except Exception as e:
            print('Editing Excel')
            print('Catch: ', e.__class__)
            return 1

    def Npoints(self):                                                      # set points for profile via excel
        try:
            self.window = Npoints.SetNpoints()
            self.window.show()
        except Exception as e:
            print('Editing Excel')
            print('Catch: ', e.__class__)
            return 1

    def TotalPoints(self):                                                  # calc/display/write total points
        try:
            self.window = TotalPoints.TotalPoints()
            self.window.show()
        except Exception as e:
            print('Editing Excel')
            print('Catch: ', e.__class__)
            return 1

    def AllTotalPoints(self):                                               # set all total points in project
        try:
            AllTotalPoints()
        except Exception as e:
            print('Editing Excel')
            print('Catch: ', e.__class__)
            return 1

    def GCW(self):                                                          # save disp csv for all profiles in project
        try:
            GetXLProfile()
        except Exception as e:
            print('Editing Excel')
            print('Catch: ', e.__class__)
            return 1
        self.CompileAndSave()
        self.WriteCSV()


    def GetProfileReports(self):                                            # generates excel file with reports
        try:
            ProfileReport()
        except Exception as e:
            print('Editing Excel')
            print('Catch: ', e.__class__)
            return 1

    def CopyToP1(self):
        try:
            select('D7')
            name = xl.Selection()
        except Exception as e:
            print('Editing Excel')
            print('Catch: ', e.__class__)
            return 1
        select('C36')
        xl.ActiveCell.Value=name
        select('D9')
        SelRange()
        data=xl.Selection()
        cells=['D36','D37','D38','D39']
        for i in range(len(cells)):
            select(cells[i])
            WriteRowFromSelected(data[i])


    def CopyDataToP2(self):
        try:
            select('D7')
            name = xl.Selection()
        except Exception as e:
            print('Editing Excel')
            print('Catch: ', e.__class__)
            return 1
        select('C44')
        xl.ActiveCell.Value = name
        select('D9')
        SelRange()
        data = xl.Selection()
        cells = ['D44', 'D45', 'D46', 'D47']
        for i in range(len(cells)):
            select(cells[i])
            WriteRowFromSelected(data[i])

    def NodeSteps(self):  # num points between nodes
        try:
            self.window = NodeSteps.SetNodeSteps()
            self.window.show()
        except Exception as e:
            print('Editing Excel')
            print('Catch: ', e.__class__)
            return 1

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = MyApp()
    window.show()
    sys.exit(app.exec_())
