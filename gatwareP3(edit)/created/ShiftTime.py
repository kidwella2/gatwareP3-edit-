
from GatwareP3 import *
import GwareData
from QtSecondWindowTest import PlotAll

from win32com.client import Dispatch
xl = Dispatch("Excel.Application")

xlToLeft = 1
xlToRight = 2
xlUp = 3
xlDown = 4
xlAscending = 1
xlYes = 1

qtCreatorFile = "ShiftTime.ui"  # Enter file here.

Ui_MainWindow, QtBaseClass = uic.loadUiType(qtCreatorFile)

class ShiftTime(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self):
        QtWidgets.QMainWindow.__init__(self)
        Ui_MainWindow.__init__(self)
        self.setupUi(self)
        xl.Workbooks("TimingSpace.xlsx").Worksheets("Sheet1").Activate()
        path = 'c:\\Python27\\'
        xl.Range("B3").Select()
        ShelveName = str(xl.ActiveCell.Value)
        filename = path + ShelveName
        GwareData.database = shelve.open(filename, writeback=False)
        GwareData.InstanceList = GwareData.database['InstanceList']
        GwareData.RawDataList = GwareData.database['RawDataList']
        self.Namelist = [obj.name for obj in GwareData.InstanceList]
        self.DictlistI = [obj.__dict__ for obj in GwareData.InstanceList]
        self.DictlistR = [obj.__dict__ for obj in GwareData.RawDataList]

    def accept(self):
        profileName = []                                    # stores profiles' name (raw)
        profileTime = []                                    # stores profiles' time (raw)
        xl.Range('C17:F17').Select()                            # get profile info from excel
        xl.Range(xl.Selection, xl.Selection.End(GwareData.xlDown)).Select()
        InputRange = xl.Selection
        RawDict = self.DictlistR
        RawDict = sorted(RawDict, key=lambda i: i['name'])  # sort list of dict by name
        # Loop through RawDataList to acquire profiles' name and time
        for item in RawDict:
            profileName.append(item['name'])
            profileTime.append(item['RawData'][0])
        Group = [column for column in InputRange.Columns()]     # store excel info in columns
        instanceName = [str(item[0]) for item in Group]  # Profile names from XL
        addTime = [str(item[3]) for item in Group]  # Alter profile time value from XL
        # Make dictionary of [instance name, add time] (excel)
        ExcelDict = list(zip(instanceName, addTime))
        ExcelFilter = [item for item in ExcelDict if 'None' not in item]        # remove values that don't change
        # Make dictionary of [profileName, profileTime] (raw data)
        RawDict = list(zip(profileName, profileTime))
        ExcelKeys = [item[0] for item in ExcelFilter]           # profile names that change
        numPairs = 0                            # iterate through ExcelFilter
        dictItem = 0                            # hold place of dict item
        alterTime = []                            # hold altered times for raw data
        try:                                # handles inappropriate values for shifting time
            for obj in RawDict:
                alter = False                   # only altered times are changed
                if obj[0] in ExcelKeys:         # handles changing time data
                    alter = True
                    addThis = ExcelFilter[numPairs][1]
                    print('Profile: ', obj[0])
                    numPairs += 1
                    print('Change By: ', addThis)
                    temp = []             # hold time values to be changed
                    for t in obj[1]:
                        t = t + float(addThis)
                        fixT = round(t, 8)    # this line fixes float problem
                        temp.append(fixT)
                if alter:
                    alterTime.append(temp)
                else:
                    alterTime.append(RawDict[dictItem][1])
                dictItem += 1
        except:
            print('No/Wrong Value for Changing Time')
            print('Catch: ', e.__class__)
            self.close()
            return 1
        RawDict = list(zip(profileName, alterTime))
        DictInst = self.DictlistI
        DictInst = sorted(DictInst, key=lambda i: i['name'])  # sort list of dict by name
        count = 0           # interate through addThis values
        for obj in DictInst:
            SelectionList = []
            if addTime[count] != 'None':
                for item in obj['Taxis']:
                    newItem = item + float(addTime[count])
                    SelectionList.append(newItem)
                DictInst[count]["Taxis"] = SelectionList
            count += 1
        # save database info
        xl.Workbooks("TimingSpace.xlsx").Worksheets("Sheet1").Activate()
        path = 'c:\\Python27\\'
        xl.Range("B3").Select()
        ShelveName = str(xl.ActiveCell.Value)
        filename = path + ShelveName
        database = shelve.open(filename, writeback=True)
        for i in range(len(DictInst)):             # save instance list
            database['InstanceList'][i].__dict__ = DictInst[i]
            print(database['InstanceList'][i].__dict__)
        for i in range(len(RawDict)):               # save raw data list
            database['RawDataList'][i].__dict__['name'] = RawDict[i][0]
            database['RawDataList'][i].__dict__['RawData'][0] = RawDict[i][1]
            print(database['RawDataList'][i].__dict__['name'], ': ', database['RawDataList'][i].__dict__['RawData'][0])
        database.sync()
        database.close()
        self.close()

    def reject(self):
        self.close()

    def LoadProfilesToExcel(self):
        PlotAll.LoadSpreadsheet(self)


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = ShiftTime()
    window.show()
    sys.exit(app.exec_())
