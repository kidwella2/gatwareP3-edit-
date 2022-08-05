import sys
from PyQt5 import QtGui, uic, QtCore, QtWidgets
from PyQt5.QtWidgets import QDialog, QApplication, QPushButton, QVBoxLayout
#from matplotlib.backends.qt_compat import QtCore, QtWidgets
import shelve
import matplotlib
#matplotlib.use('WXAgg')
import matplotlib.pyplot as plt
from matplotlib import rcParams
from pylab import figure, show
from datetime import datetime
#from StackedPlotting import plotroutine as pr
from CurrentProfilePanel import *
#from QTWidgetPLot import *
import QTWidgetPLot
import gc
import random
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.backends.backend_qt5agg import NavigationToolbar2QT as NavigationToolbar



from win32com.client import Dispatch
import GwareData

xl = Dispatch("Excel.Application")

qtCreatorFile = "PlotAllDialog.ui" # Enter file here.

Ui_MainWindow, QtBaseClass = uic.loadUiType(qtCreatorFile)

class PlotAll(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self):
        QtWidgets.QMainWindow.__init__(self)
        Ui_MainWindow.__init__(self)
        self.setupUi(self)
        xl.Workbooks("TimingSpace.xlsx").Worksheets("Sheet1").Activate()
        path = 'c:\\Python27\\'
        xl.Range("B3").Select()
        ShelveName = str(xl.ActiveCell.Value)
        suffix = '.dat'
        # filename = path + ShelveName + suffix
        filename = path + ShelveName  # removed suffix 1/6/2021
        GwareData.database = shelve.open(filename, writeback=False)
        GwareData.InstanceList = GwareData.database['InstanceList']
        GwareData.RawDataList = GwareData.database['RawDataList']
        ProfileNames = [obj.name for obj in GwareData.InstanceList]
        self.Namelist = [obj.name for obj in GwareData.InstanceList]
        print(ProfileNames)

    def gettitle(self):
        xl.Range("B3").Select()
        ProjectName = xl.Selection
        return str(ProjectName)

    def plotEverything(self, value):
        self.figure.clear()
        print('pltEverything', value)
        rcParams['ytick.labelsize'] = 7
        rcParams['xtick.labelsize'] = 6
        rcParams['font.size'] = 10
        maxt = []
        ProfilesToPlot = [obj for obj in self.Namelist]
        SelectionList = []
        #print(ProfilesToPlot)
        # Extra cheking to ensure that item is in DataBase
        for thisprofile in ProfilesToPlot:
            for obj in GwareData.InstanceList:
                if obj.name == thisprofile:
                    SelectionList.append(obj)
        # SelectionList has only valid CurrentProfilePanel.Profile instances now
        #print(SelectionList)
        #try:
        for obj in SelectionList:  # Figure out how long the time axis is
            timevalues = obj.Taxis
            maxt.append(timevalues[-1])
                # maxt.sort()
        #except Exception as e:
            #print(len(ProfilesToPlot), 'is too many profiles for MatPlotLib to support')
            #print('Catch: ', e.__class__)
            #plt.close()
            #return 1
        #print(SelectionList)
        tmax = max(maxt)
        maxcount = len(SelectionList)  # how many plots to make?
        pairs = list(zip(GwareData.RawDataList, GwareData.InstanceList))
        selectedpairs = []
        for item in SelectionList:
            for pair in pairs:
                if pair[1] == item:
                    selectedpairs.append(pair)
        # Selected Pairs is the validated list of Raw Data and Instances to be plotted
        subplotlist = []
        axlist = []
        for i in np.arange(1, len(selectedpairs) + 1):
            subplotlist.append(f"{len(selectedpairs)}{1}{i}")
            axlist.append(f"ax{i}")
        # print(subplotlist, axlist)
        pos = 0
        for i in subplotlist:
            thispair = selectedpairs[pos]
            RawData = thispair[0]
            times = RawData.RawData[0]
            vlines = [time for time in times]
            obj = thispair[1]
            #print(obj)
            if value == 'Disp':
                # GwareData.count=0
                DataString = obj.DispAxis
                Yaxis = obj.DispAxis  # Displacement axis from Instance
                titletext = '   Displacement vs Time (Selected Only) '
            if value == 'Vel':
                # GwareData.count = 0
                DataString = obj.VelAxis
                Yaxis = obj.VelAxis  # Displacement axis from Instance
                titletext = '   Velocity vs Time (Selected Only) '
            if value == 'Acc':
                # GwareData.count = 0
                DataString = obj.AccAxis
                Yaxis = obj.AccAxis  # Displacement axis from Instance
                titletext = '   Acceleration vs Time (Selected Only) '
            Ymax = max(DataString)
            Ymin = min(DataString)
            Xaxis = obj.Taxis
            Toggle = [1, 0]
            ax = axlist[pos]
            ncol = len(selectedpairs)                                           # get number of rows and columns
            nrow = 1
            if len(selectedpairs) > 10:                                         # get 2 rows if more than 10 plots
                ncol = int(len(selectedpairs) / 2) + len(selectedpairs) % 2
                nrow = 2
            ax = self.figure.add_subplot(ncol, nrow, pos + 1)                   # num columns, num rows, index
            for eachvline in vlines:
                Toggle.reverse()
                Max = Toggle[0]
                ax.axvline(eachvline, color='r', ls='dashed', linewidth=2)
                s = str(eachvline)
                s = s[:5]
                if Max:
                    ax.text(eachvline, Ymax, s, fontsize=8, bbox=dict(facecolor='white', alpha=.75))
                if not Max:
                    ax.text(eachvline, Ymin, s, fontsize=8, bbox=dict(facecolor='white', alpha=.75))
            #print(ax)
            ax.plot(Xaxis, Yaxis)
            ax.set_xlim(0, tmax)
            ax.set_ylabel(obj.name, rotation='vertical')
            ax.grid(color='#7171C6', linestyle='-', linewidth=.2)  # light grey)
            dandtime = str(datetime.now().strftime('%Y-%m-%d %H:%M'))
            projname = self.gettitle()
            plt.suptitle(projname + titletext + dandtime, fontsize=14)
            pos += 1
        plt.show()
        #self.canvas.draw()


    def plotSelected(self,value):
        self.figure.clear()
        print('pltrtne', value)
        rcParams['ytick.labelsize'] = 7
        rcParams['xtick.labelsize'] = 6
        rcParams['font.size'] = 10
        maxt = []
        xl.Range("C17").Select()
        xl.Range(xl.Selection, xl.Selection.End(GwareData.xlDown)).Select()
        xl.Range(xl.Selection, xl.Selection.End(GwareData.xlToRight)).Select()
        InputRange = xl.Selection
        #print(InputRange)
        Group = [column for column in InputRange.Columns()]
        Profilekey = [str(item[0]) for item in Group]  # Profile names from XL
        #print(Profilekey)
        NewPlotOrder = [item[1] for item in Group]  # Plot order from XL
        # Make dictionary of [plot order,profile name]
        PlotDict = list(zip(NewPlotOrder, Profilekey))
        #print(PlotDict)
        PlotFilter = [item for item in PlotDict if isinstance(item[0], float)]
        # If there is a number assigned to the profile
        PlotFilter.sort()
        #print(PlotFilter)
        ProfilesToPlot = [obj[1] for obj in PlotFilter]
        SelectionList = []
        #print(ProfilesToPlot)
        # Extra cheking to ensure that item is in DataBase
        for thisprofile in ProfilesToPlot:
            for obj in GwareData.InstanceList:
                if obj.name == thisprofile:
                    SelectionList.append(obj)
        # SelectionList has only valid CurrentProfilePanel.Profile instances now
        for obj in SelectionList:  # Figure out how long the time axis is
            timevalues = obj.Taxis
            maxt.append(timevalues[-1])
            # maxt.sort()
        try:                                        # fix issue if no plot is selected
            tmax = max(maxt)
            maxcount = len(SelectionList)  # how many plots to make?
            pairs = list(zip(GwareData.RawDataList, GwareData.InstanceList))
            selectedpairs = []
        except Exception as e:
            print('Plots not selected properly')
            print('Catch: ', e.__class__)
            plt.close()
            return 1
        for item in SelectionList:
            for pair in pairs:
                if pair[1] == item:
                    selectedpairs.append(pair)
        #Selected Pairs is the validated list of Raw Data and Instances to be plotted
        subplotlist=[]
        axlist=[]
        for i in np.arange(1, len(selectedpairs) + 1):
            subplotlist.append(f"{len(selectedpairs)}{1}{i}")
            axlist.append(f"ax{i}")
        #print(subplotlist,axlist)
        pos=0
        for i in subplotlist:
            thispair=selectedpairs[pos]
            RawData = thispair[0]
            times = RawData.RawData[0]
            vlines = [time for time in times]
            obj = thispair[1]
            if value == 'Disp':
                # GwareData.count=0
                DataString = obj.DispAxis
                Yaxis = obj.DispAxis  # Displacement axis from Instance
                titletext = '   Displacement vs Time (Selected Only) '
            if value == 'Vel':
                # GwareData.count = 0
                DataString = obj.VelAxis
                Yaxis = obj.VelAxis  # Displacement axis from Instance
                titletext = '   Velocity vs Time (Selected Only) '
            if value == 'Acc':
                # GwareData.count = 0
                DataString = obj.AccAxis
                Yaxis = obj.AccAxis  # Displacement axis from Instance
                titletext = '   Acceleration vs Time (Selected Only) '
            Ymax = max(DataString)
            Ymin = min(DataString)
            Xaxis = obj.Taxis
            Toggle = [1, 0]
            ax=axlist[pos]
            ncol = len(selectedpairs)  # get number of rows and columns
            nrow = 1
            if len(selectedpairs) > 10:  # get 2 rows if more than 10 plots
                ncol = int(len(selectedpairs) / 2) + len(selectedpairs) % 2
                nrow = 2
            ax = self.figure.add_subplot(ncol, nrow, pos + 1)  # num columns, num rows, index
            for eachvline in vlines:
                Toggle.reverse()
                Max = Toggle[0]
                ax.axvline(eachvline, color='r', ls='dashed', linewidth=2)
                s = str(eachvline)
                s = s[:5]
                if Max:
                    ax.text(eachvline, Ymax, s, fontsize=8, bbox=dict(facecolor='white', alpha=.75))
                if not Max:
                    ax.text(eachvline, Ymin, s, fontsize=8, bbox=dict(facecolor='white', alpha=.75))
            #print(ax)
            ax.plot(Xaxis, Yaxis)
            ax.set_xlim(0, tmax)
            ax.set_ylabel(obj.name, rotation='vertical')
            ax.grid(color='#7171C6', linestyle='-', linewidth=.2)  # light grey)
            dandtime = str(datetime.now().strftime('%Y-%m-%d %H:%M'))
            projname=self.gettitle()
            plt.suptitle(projname + titletext + dandtime, fontsize=14)
            pos+=1
        plt.show()
        self.canvas.draw()


    def plotroutine(self,value):
        self.figure = plt.figure()
        # this is the Canvas Widget that displays the `figure`
        # it takes the `figure` instance as a parameter to __init__
        self.canvas = FigureCanvas(self.figure)
        # this is the Navigation widget
        # it takes the Canvas widget and a parent
        self.toolbar = NavigationToolbar(self.canvas, self)
        # Just some button connected to `plot` method
        # set the layout
        layout = QVBoxLayout()
        layout.addWidget(self.toolbar)
        layout.addWidget(self.canvas)
        self.setLayout(layout)
        if self.selectedOnly:
            self.plotSelected(self.flag)
        if not self.selectedOnly:
            self.plotEverything(self.flag)




    def LoadSpreadsheet(self):
        '''
            Called from "Load Database" button

            Opens the database file referenced by cell B3 in the Excel
            worksheet

            Not writeback enabled----
            '''
        # print "hello from machine timing methods-->getexistingshelve()"
        print("you clicked the LoadspreadsheetButton")
        #global database, RepeatData, ShelveName
        xl.Workbooks("TimingSpace.xlsx").Worksheets("Sheet1").Activate()
        path = 'c:\\Python27\\'
        xl.Range("B3").Select()
        ShelveName = str(xl.ActiveCell.Value)
        suffix = '.dat'
        #filename = path + ShelveName + suffix
        filename = path + ShelveName #removed suffix 1/6/2021
        GwareData.database = shelve.open(filename, writeback=False)
        GwareData.InstanceList = GwareData.database['InstanceList']
        GwareData.RawDataList = GwareData.database['RawDataList']

        ProfileNames = [obj.name for obj in GwareData.InstanceList]

        xl.Range("C17").Select()
        if not xl.Selection == 'None':
            xl.Range(xl.Selection, xl.Selection.End(GwareData.xlDown)).Select()
            InputRange = xl.Selection
            InputRange.ClearContents()
            xl.Range("C17").Select()

        listing = [name for name in ProfileNames]
        listing.sort()

        for thislisting in listing:
            xl.ActiveCell.Value = thislisting
            xl.ActiveCell.Offset(2, 1).Select()

    def PlotDisplacement(self):
        self.plotType()
        self.flag='Disp'
        self.plotroutine(self.flag)

    def PlotVelocity(self):
        self.plotType()
        self.flag = 'Vel'
        self.plotroutine(self.flag)

    def PlotAcceleration(self):
        self.plotType()
        self.flag = 'Acc'
        self.plotroutine(self.flag)

    def plotType(self):
        if self.ChooseProfilebox.isChecked():
            self.selectedOnly = 1
        if not self.ChooseProfilebox.isChecked():
            self.selectedOnly = 0
        print('Status of checkBox= ',self.selectedOnly)


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = PlotAll()
    window.show()
    sys.exit(app.exec_())