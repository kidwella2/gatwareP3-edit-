
from GatwareP3 import *
import GwareData
from QtSecondWindowTest import PlotAll
from uniquer import UniqueSeq

from win32com.client import Dispatch
xl = Dispatch("Excel.Application")

xlToLeft = 1
xlToRight = 2
xlUp = 3
xlDown = 4
xlAscending = 1
xlYes = 1

qtCreatorFile = "XaxisBasis.ui"  # Enter file here.

Ui_MainWindow, QtBaseClass = uic.loadUiType(qtCreatorFile)

class XaxisBasis(QtWidgets.QMainWindow, Ui_MainWindow):
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
        oldMax = []                                          # store largest num on x-axis (raw)
        xl.Range('C17:G17').Select()                            # get profile info from excel
        xl.Range(xl.Selection, xl.Selection.End(GwareData.xlDown)).Select()
        InputRange = xl.Selection
        RawDict = self.DictlistR
        RawDict = sorted(RawDict, key=lambda i: i['name'])  # sort list of dict by name
        InstDict = self.DictlistI
        InstDict = sorted(InstDict, key=lambda i: i['name'])  # sort list of dict by name
        # Loop through RawDataList to acquire profiles' name and time
        for item in RawDict:
            profileName.append(item['name'])
            profileTime.append(item['RawData'][0])
            oldMax.append(max(item['RawData'][0]))
        Group = [column for column in InputRange.Columns()]  # store excel info in columns
        instanceName = [str(item[0]) for item in Group]  # Profile names from XL
        maxX = [str(item[4]) for item in Group]  # Alter profile x-axis value from XL
        # Make dictionary of [instance name, max x value] (excel)
        xlist = list(zip(instanceName, maxX))
        xlistFilter = [item for item in xlist if 'None' not in item]  # remove values that don't change
        # Make dictionary of [profileName, profileTime] (raw data)
        tempRaw = list(zip(profileName, profileTime))
        ExcelKeys = [item[0] for item in xlistFilter]  # profile names that change
        numChange = 0                           # iterate through ExcelFilter
        dictItem = 0                            # hold place of dict item
        alterXaxis = []                           # hold altered times for raw data
        instXaxis = []                          # hold altered times for instance data
        try:
            for obj in tempRaw:
                alter = False
                if obj[0] in ExcelKeys:
                    alter = True
                    newMax = xlistFilter[numChange][1]
                    print('Profile: ', obj[0])
                    print('Change To: ', newMax)
                    numChange += 1
                    temp = []  # hold time values to be changed
                    temp2 = []
                    for t in obj[1]:                                # change raw data
                        alterFraction = float(newMax) / float(oldMax[dictItem])
                        t = t * alterFraction
                        fixT = round(t, 8)  # this line fixes float problem
                        temp.append(fixT)
                    for inst in InstDict[dictItem]['Taxis']:             # change instance data
                        inst = inst * alterFraction
                        temp2.append(inst)
                if alter:
                    alterXaxis.append(temp)
                    instXaxis.append(temp2)
                else:
                    alterXaxis.append(RawDict[dictItem]['RawData'][0])
                    instXaxis.append(InstDict[dictItem]['Taxis'])
                dictItem += 1
        except:
            print('No/Wrong Value for Changing Time')
            print('Catch: ', e.__class__)
            self.close()
            return 1
        for i in range(len(RawDict)):
            InstDict[i]['Taxis'] = instXaxis[i]
            RawDict[i]['RawData'][0] = alterXaxis[i]
        for j in range(len(InstDict)):                              # eliminates duplicate values
            t = []
            x = []
            v = []
            a = []
            dupt = InstDict[j]['Taxis']
            dupx = InstDict[j]['DispAxis']
            dupv = InstDict[j]['VelAxis']
            dupa = InstDict[j]['AccAxis']
            seq = UniqueSeq(dupt, dupx)
            UniqueList = seq.seqlist
            seq2 = UniqueSeq(dupt, dupv)
            UniqueList2 = seq2.seqlist
            seq3 = UniqueSeq(dupt, dupa)
            UniqueList3 = seq3.seqlist
            for obj in UniqueList:
                t.append(obj[0])
                x.append(obj[1])
            for obj2 in UniqueList2:
                v.append(obj2[1])
            for obj3 in UniqueList3:
                a.append(obj3[1])
            InstDict[j]['Taxis'] = t
            InstDict[j]['DispAxis'] = x
            InstDict[j]['VelAxis'] = v
            InstDict[j]['AccAxis'] = a
        # save database info
        xl.Workbooks("TimingSpace.xlsx").Worksheets("Sheet1").Activate()
        path = 'c:\\Python27\\'
        xl.Range("B3").Select()
        ShelveName = str(xl.ActiveCell.Value)
        filename = path + ShelveName
        database = shelve.open(filename, writeback=True)
        for i in range(len(InstDict)):  # save instance list
            database['InstanceList'][i].__dict__ = InstDict[i]
            print(database['InstanceList'][i].__dict__)
        for i in range(len(RawDict)):  # save raw data list
            database['RawDataList'][i].__dict__ = RawDict[i]
            print(database['RawDataList'][i].__dict__)
        database.sync()
        database.close()
        self.close()

    def reject(self):
            self.close()

    def LoadProfilesToExcel(self):
            PlotAll.LoadSpreadsheet(self)


if __name__ == "__main__":
        app = QtWidgets.QApplication(sys.argv)
        window = XaxisBasis()
        window.show()
        sys.exit(app.exec_())
