
import GatWareGetFromXL
from GatwareP3 import *
import GwareData
from QtSecondWindowTest import PlotAll
from TotalPoints import TotalPoints
from uniquer import UniqueSeq
import numpy as np
from scipy.interpolate import *
import matplotlib.pyplot as plt

from win32com.client import Dispatch

xl = Dispatch("Excel.Application")

xlDown = 4

qtCreatorFile = "Npoints.ui"  # Enter file here.

Ui_MainWindow, QtBaseClass = uic.loadUiType(qtCreatorFile)


class SetNpoints(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self):
        QtWidgets.QMainWindow.__init__(self)
        Ui_MainWindow.__init__(self)
        self.setupUi(self)
        xl.Workbooks("TimingSpace.xlsx").Worksheets("Sheet1").Activate()
        path = 'c:\\Python27\\'
        xl.Range("B3").Select()
        self.ShelveName = str(xl.ActiveCell.Value)
        filename = path + self.ShelveName
        GwareData.database = shelve.open(filename, writeback=False)
        GwareData.InstanceList = GwareData.database['InstanceList']
        GwareData.RawDataList = GwareData.database['RawDataList']
        self.DictlistI = [obj.__dict__ for obj in GwareData.InstanceList]
        self.DictlistR = [obj.__dict__ for obj in GwareData.RawDataList]

    def accept(self):
        xl.Range('C17:H17').Select()  # get profile info from excel
        xl.Range(xl.Selection, xl.Selection.End(GwareData.xlDown)).Select()
        InputRange = xl.Selection
        Group = [column for column in InputRange.Columns()]  # store excel info in columns
        instanceName = [str(item[0]) for item in Group]  # Profile names from XL
        numPoints = [str(item[5]) for item in Group]  # Alter profiles' points from XL
        # Make dictionary of [instance name, num points] (excel)
        ExcelDict = list(zip(instanceName, numPoints))
        print(ExcelDict)
        TotalPoints.ShelveName = self.ShelveName  # initialize self values on TotalPoints to avoid crash
        TotalPoints.DictlistI = self.DictlistI
        TotalPoints.newx = []
        TotalPoints.newy = []
        TotalPoints.profilename = ''

        try:  # handles inappropriate values for number of points
            for obj in ExcelDict:
                if obj[1] != 'None':  # only altered times are changed
                    print('')
                    print('Profile: ', obj[0])
                    print('Total Points: ', obj[1])
                    xl.Range("D7").Select()  # set profile name and total points to excel
                    xl.ActiveCell.Value = obj[0]
                    xl.Range("B7").Select()
                    xl.ActiveCell.Value = obj[1]
                    GatWareGetFromXL.GetXLProfile()
                    TotalPoints.setTP(self=TotalPoints)  # display and write total points
                    TotalPoints.writeCSV(self=TotalPoints)
        except:
            print('No/Wrong Value for Total Points')
            print('Catch: ', e.__class__)
            self.close()
            return 1

    def DisplayAndSave(self):
        profileName = []    # stores profiles' name (inst)
        newName = []
        profileTime = []    # stores profiles' time (inst)
        newTime = []
        profileDisp = []    # stores profiles' displacement (inst)
        newDisp = []
        profileVel = []     # stores profiles' velocity (inst)
        newVel = []
        profileAcc = []     # stores profiles' acceleration (inst)
        newAcc = []
        xl.Range('C17:H17').Select()  # get profile info from excel
        xl.Range(xl.Selection, xl.Selection.End(GwareData.xlDown)).Select()
        InputRange = xl.Selection
        InstDict = self.DictlistI
        InstDict = sorted(InstDict, key=lambda i: i['name'])  # sort list of dict by name
        Group = [column for column in InputRange.Columns()]  # store excel info in columns
        instanceName = [str(item[0]) for item in Group]  # Profile names from XL
        numPoints = [item[5] for item in Group]  # Alter profiles' points from XL
        # Make dictionary of [instance name, num points] (excel)
        ExcelDict = list(zip(instanceName, numPoints))
        print(ExcelDict)
        pos = 0                                         # database position of current profile

        try:  # handles inappropriate values for number of points
            for obj in ExcelDict:
                if obj[1] is not None:  # handles changing time data
                    print('')
                    print('Profile: ', obj[0])
                    print('Total Points: ', obj[1])
                    xl.Range("D7").Select()  # set profile name and total points to excel
                    xl.ActiveCell.Value = obj[0]
                    GatWareGetFromXL.GetXLProfile()
                    totalPoints = int(obj[1])
                    x = []  # x & y values before change
                    y = []
                    y2 = []
                    y3 = []
                    dupx = InstDict[pos]['Taxis']
                    dupy = InstDict[pos]['DispAxis']
                    dupy2 = InstDict[pos]['VelAxis']
                    dupy3 = InstDict[pos]['AccAxis']
                    seq = UniqueSeq(dupx, dupy)  # eliminates duplicate values
                    seq2 = UniqueSeq(dupx, dupy2)
                    seq3 = UniqueSeq(dupx, dupy3)
                    UniqueList = seq.seqlist
                    UniqueList2 = seq2.seqlist
                    UniqueList3 = seq3.seqlist
                    for obj2 in UniqueList:
                        x.append(obj2[0])
                        y.append(obj2[1])
                    for obj2 in UniqueList2:
                        y2.append(obj2[1])
                    for obj2 in UniqueList3:
                        y3.append(obj2[1])
                    profileName.append(InstDict[pos]['name'])           # keep original values in lists
                    profileTime.append(x)
                    profileDisp.append(y)
                    profileVel.append(y2)
                    profileAcc.append(y3)

                    print('Old Points: ', len(x))
                    tck = splrep(x, y)                              # used to change num of points
                    newx = np.linspace(x[0], max(x), totalPoints)
                    print('New Points: ', len(newx))
                    newy = [splev(i, tck) for i in newx]
                    newy = [float(i) for i in newy]
                    tck2 = splrep(x, y2)
                    newy2 = [splev(i, tck2) for i in newx]
                    newy2 = [float(i) for i in newy2]
                    tck3 = splrep(x, y3)
                    newy3 = [splev(i, tck3) for i in newx]
                    newy3 = [float(i) for i in newy3]
                    newName.append(obj[0])                                  # store new values in lists
                    newTime.append(newx)
                    newDisp.append(newy)
                    newVel.append(newy2)
                    newAcc.append(newy3)
                pos += 1
        except:
            print('No/Wrong Value for Total Points')
            print('Catch: ', e.__class__)
            self.close()
            return 1

        fig = plt.figure()                                                  # plot

        ncol = len(newTime)  # get number of rows and columns
        nrow = 3
        axlist = []
        for i in np.arange(1, len(newTime) * 3 + 1):
            axlist.append(f"ax{i}")

        for i in range(len(newTime)):
            tmax = max(newTime[i])
            ax = axlist[i * 3]                                              # plot disp

            ax = fig.add_subplot(ncol, nrow, i * 3 + 1)  # num columns, num rows, index
            plt.suptitle(self.ShelveName, fontsize=14)
            ax.plot(profileTime[i], profileDisp[i], color='b', lw=3, label='Original Disp')
            ax.plot(newTime[i], newDisp[i], color='r', lw=1, label='New Disp')
            ax.grid(color='#7171C6', linestyle='-', linewidth=.2)  # light grey)
            plt.autoscale(enable=True, axis='both', tight=None)
            ax.set_xlim(0, tmax)
            ax.set_ylabel(newName[i], rotation='vertical')
            plt.legend()

            ax = axlist[i * 3 + 1]                                          # plot vel
            ax = fig.add_subplot(ncol, nrow, i * 3 + 2)  # num columns, num rows, index
            ax.plot(profileTime[i], profileVel[i], color='b', lw=3, label='Original Vel')
            ax.plot(newTime[i], newVel[i], color='r', lw=1, label='New Vel')
            ax.grid(color='#7171C6', linestyle='-', linewidth=.2)  # light grey)
            ax.set_xlim(0, tmax)
            ax.set_ylabel(newName[i], rotation='vertical')
            plt.legend()

            ax = axlist[i * 3 + 2]                                          # plot acc
            ax = fig.add_subplot(ncol, nrow, i * 3 + 3)  # num columns, num rows, index
            ax.plot(profileTime[i], profileAcc[i], color='b', lw=3, label='Original Acc')
            ax.plot(newTime[i], newAcc[i], color='r', lw=1, label='New Acc')
            ax.grid(color='#7171C6', linestyle='-', linewidth=.2)  # light grey)
            ax.set_xlim(0, tmax)
            ax.set_ylabel(newName[i], rotation='vertical')
            plt.legend()

        plt.show()

        for i in range(len(InstDict)):                              # store new values in InstDict
            if InstDict[i]['name'] in newName:
                index = newName.index(InstDict[i]['name'])
                InstDict[i]['Taxis'] = newTime[index]
                InstDict[i]['DispAxis'] = newDisp[index]
                InstDict[i]['VelAxis'] = newVel[index]
                InstDict[i]['AccAxis'] = newAcc[index]
        RawDict = self.DictlistR
        RawDict = sorted(RawDict, key=lambda i: i['name'])  # sort list of dict by name
        # save database info
        xl.Workbooks("TimingSpace.xlsx").Worksheets("Sheet1").Activate()
        path = 'c:\\Python27\\'
        xl.Range("B3").Select()
        ShelveName = str(xl.ActiveCell.Value)
        filename = path + ShelveName
        database = shelve.open(filename, writeback=True)
        for i in range(len(InstDict)):                                      # save instance list
            database['InstanceList'][i].__dict__ = InstDict[i]
        for i in range(len(RawDict)):                                       # save raw data list
            database['RawDataList'][i].__dict__ = RawDict[i]
        database.sync()
        database.close()
        self.close()

    def reject(self):
        self.close()

    def LoadProfilesToExcel(self):
        PlotAll.LoadSpreadsheet(self)


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = SetNpoints()
    window.show()
    sys.exit(app.exec_())
