
import GatwareP3
from GatwareP3 import *
import GatWareGetFromXL
import GwareData

from win32com.client import Dispatch
xl = Dispatch("Excel.Application")

qtCreatorFile = "NodeSteps.ui"  # Enter file here.

Ui_MainWindow, QtBaseClass = uic.loadUiType(qtCreatorFile)

xlToLeft = 1
xlToRight = 2
xlUp = 3
xlDown = 4


class SetNodeSteps(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self):
        QtWidgets.QMainWindow.__init__(self)
        Ui_MainWindow.__init__(self)
        self.setupUi(self)
        xl.Workbooks("TimingSpace.xlsx").Worksheets("Sheet1").Activate()
        path = 'c:\\Python27\\'
        xl.Range("B3").Select()
        ShelveName = str(xl.ActiveCell.Value)
        filename = path + ShelveName
        GwareData.database = shelve.open(filename, writeback=True)
        GwareData.RawDataList = GwareData.database['RawDataList']
        self.DictlistR = [obj.__dict__ for obj in GwareData.RawDataList]
        self.step = getSteps()

    def accept(self):
        RawDict = self.DictlistR
        try:
            self.step = int(self.lineEdit_Steps.text())                 # get user info for steps per node
            self.setSteps()
        except Exception as e:
            print('No/Wrong Value for Steps per Node')
            print('Catch: ', e.__class__)
            self.close()
            return 1

        for i in range(len(RawDict)):                               # loop through all profiles in project
            print(RawDict[i]['name'])
            xl.Range("D9").Select()                                     # locate and clear current profile data
            xl.Range(xl.Selection, xl.Selection.End(GwareData.xlToRight)).Select()
            xl.Range(xl.Selection, xl.Selection.End(GwareData.xlDown)).Select()
            rangeXL = xl.Selection
            rangeXL.ClearContents()

            Profilename = RawDict[i]['name']                    # assign profile values for xl
            data = RawDict[i]['RawData']
            time = data[0]
            position = data[1]
            velocity = data[2]
            acceleration = data[3]
            xl.Range("D7").Select()  # put name in worksheet

            xl.ActiveCell.Value = Profilename
            xl.Range("D9").Select()  # upper left column of range

            for value in time:                                  # fill xl with project info
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

            GatWareGetFromXL.GetXLProfile()                     # get xl profile and compile
            GatwareP3.MyApp.CompileAndSave(self)
        print('Steps between Nodes: ', getSteps())
        self.close()

    def reject(self):                               # cancel button
        self.close()

    def setSteps(self):                             # update steps per node
        GwareData.steps = self.step


def getSteps():                                     # get current steps per node
    return GwareData.steps


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = SetNodeSteps()
    window.show()
    sys.exit(app.exec_())
