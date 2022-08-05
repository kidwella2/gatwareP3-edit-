import GwareData
from XLSX_Initializer import *
from win32com.client import Dispatch
from MotionPolys import *
xl = Dispatch("Excel.Application")

#motionparams={}

SeventhOrder=0

def segname():
    '''
    Gets the User entered GwareData.Segment name from Excel @(D7)
    '''
    xl.Workbooks("TimingSpace.xlsx").Worksheets("Sheet1").Activate()
    xl.Range("D7").Select()
    GwareData.segmentname = xl.ActiveCell.Value
    GwareData.segmentname = str(GwareData.segmentname)
    #print(GwareData.segmentname)


def GetXLProfile():  # Gets the data from worksheet
    ''' Checks the worksheet and determines how many segments
    comprise the current profile.

    calls->Datapresent(),GetProfileName(),BuildSegment()

    ---> creates the variable NumberOfSegments
    '''
    #from win32com.client import Dispatch
    #print("HelloFromDataPresent")

    def BuildSegment(Segment):

        ''' Beginning at the active cell, this function loops through
            the worksheet compiling 8 values to pass as arguments into
            the FitPoly class(es). Each loop creates a sequentially numbered child
            of the Main Segment name. Each segment is an instance of
            the Fitpoly class and is stored in the motionparams
            dictionary under the segment name key.
        '''
        #global count, motionparams
        count = 0

        def Offset():
            xl.ActiveCell.Offset(2, 1).Select()

        def SwitchColumns():
            xl.ActiveCell.Offset(-2, 2).Select()

        def ColumnTop():
            xl.ActiveCell.Offset(-2, 1).Select()

        ######################################################################
        #  Loop through profile range and gather (t,x,v,a,T,X,V,A) up to value
        #  of NumberofSegments
        #  Call Fitpoly for each segment and add that segment to the
        #  motionparams dictionary
        ######################################################################

        for loop in np.arange(0, NumberOfSegments, 1):
            #global motionparams
            GwareData.Segment = key + str(count)
            #print(GwareData.Segment)
            t = xl.ActiveCell.Value
            Offset()
            x = xl.ActiveCell.Value
            Offset()
            v = xl.ActiveCell.Value
            Offset()
            a = xl.ActiveCell.Value
            SwitchColumns()
            T = xl.ActiveCell.Value
            Offset()
            X = xl.ActiveCell.Value
            Offset()
            V = xl.ActiveCell.Value
            Offset()
            A = xl.ActiveCell.Value
            variables = [t, x, v, a, T, X, V, A]

            if SeventhOrder:
                j = float(0)
                J = float(0)
                GwareData.motionparams[GwareData.Segment] = SepticFit(t, x, v, a, j, T, X, V, A, J, GwareData.steps)

            elif not SeventhOrder:
                GwareData.motionparams[GwareData.Segment] = QuinticFit(t, x, v, a, T, X, V, A, GwareData.steps)
            #print(GwareData.motionparams.keys())

            ColumnTop()
            count += 1

    xl = Dispatch("Excel.Application")
    xl.Workbooks("TimingSpace.xlsx").Worksheets("Sheet1").Activate()
    xl.Range("D9").Select()
    contents = xl.Selection()

    if not contents == None:
        GwareData.Datapresent = 1

    if contents == None:
        GwareData.Datapresent = 0
    # end of DataPresent()

    if not GwareData.Datapresent:
        print('pass')

    #SeventhOrder = self.checkBox1.GetValue()
    #global SeventhOrder

    SeventhOrder=0

    #make sure that information is available

    if GwareData.Datapresent:
        xl = Dispatch("Excel.Application")
        xl.Workbooks("TimingSpace.xlsx").Worksheets("Sheet1").Activate()
        xl.Range("D9").Activate()  # upper left of profile range
        xl.Range(xl.Selection, xl.Selection.End(xlToRight)).Select()
        ColumnCount = xl.Selection
        NumberOfSegments = len(ColumnCount) - 1  # creates variable
        #print(NumberOfSegments)
        segname()
        key = GwareData.segmentname
        xl.Range("D9").Select()  # select upper left of profile range before leaving
        BuildSegment(GwareData.segmentname)
        #print(key)
    #GwareData.motionparams

if __name__ =='__main__':
    GetXLProfile()