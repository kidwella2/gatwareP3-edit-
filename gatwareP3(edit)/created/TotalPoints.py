
from GatwareP3 import *
import GwareData
from uniquer import UniqueSeq
import numpy as np
from scipy.interpolate import *
import matplotlib.pyplot as plt

from win32com.client import Dispatch
xl = Dispatch("Excel.Application")

qtCreatorFile = "TotalPoints.ui"  # Enter file here.

Ui_MainWindow, QtBaseClass = uic.loadUiType(qtCreatorFile)

class TotalPoints(QtWidgets.QMainWindow, Ui_MainWindow):
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
        self.DictlistI = [obj.__dict__ for obj in GwareData.InstanceList]
        self.newx = []
        self.newy = []
        self.profilename = ''

    def setTP(self):                                    # set total points and graph
        xl.Range("D7").Select()                                 # get profile name from excel
        profileName = str(xl.Selection)
        print('Profile: ', profileName)
        self.profilename = profileName                          # set to use in write
        xl.Range("B7").Select()                                 # get total points from excel
        try:
            totalPoints = int(xl.Selection)
        except Exception as e:
            print('No/Wrong Value for Steps per Node')
            print('Catch: ', e.__class__)
            return 1

        InstDict = self.DictlistI
        InstDict = sorted(InstDict, key=lambda i: i['name'])  # sort list of dict by name
        pos = 0                                             # database position of current profile
        # Loop through InstanceList to acquire profiles' position
        for obj in InstDict:
            if obj['name'] == profileName:
                pos = InstDict.index(obj)
        x = []                                              # x & y values before change
        y = []
        dupx = InstDict[pos]['Taxis']
        dupy = InstDict[pos]['DispAxis']
        seq = UniqueSeq(dupx, dupy)                           # eliminates duplicate values
        UniqueList = seq.seqlist
        for obj in UniqueList:
            x.append(obj[0])
            y.append(obj[1])
        print('Old Points: ', len(x))
        tck = splrep(x, y)                                      # used to change num of points
        newx = np.linspace(x[0], max(x), totalPoints)
        print('New Points: ', len(newx))
        newy = [splev(i, tck) for i in newx]
        newy = [float(i) for i in newy]
        self.newx = newx                                        # set for use in write
        self.newy = newy

        # plt.close()                                      # plot
        plt.figure()

        plt.suptitle(self.profilename, fontsize=14)
        plt.plot(x, y, color='b', lw=3, label='Original')
        plt.plot(newx, newy, color='r', lw=1, label='New')
        plt.grid()
        plt.autoscale(enable=True, axis='both', tight=None)

        plt.legend()
        plt.show()

    def writeCSV(self):
        if len(self.newx) != 0:                             # if num points changed
            path = 'c:/Python27'
            roundy = []
            for y in self.newy:                                 # remove super small nums that should be 0
                if abs(y) < 0.000001:
                    y = 0
                roundy.append(y)
            listCSV = list(zip(self.newx, roundy))           # make list of tuples to write
            file = path + '/' + self.ShelveName + '_' + self.profilename + '(' + str(len(listCSV)) + ').csv'
            print(file)
            with open(file, 'w', newline='') as f:          # write
                c = csv.writer(f)
                for row in listCSV:
                    c.writerow(row)

    def cancel(self):
        self.close()


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = TotalPoints()
    window.show()
    sys.exit(app.exec_())
