#-----------------------------------------------------------------------------
# Name:        CurrentProfilePanel.py
# Purpose:     
#
# Author:      <Todd Gatman>
#
# Created:     2011/02/04
# RCS-ID:      $Id: CurrentProfilePanel.py $
# Copyright:   (c) 2006
# Licence:     <your licence>
#-----------------------------------------------------------------------------
#Boa:FramePanel:Panel1

import shelve
from types import *
import os
import GwareData
from GwareData import *

#import wx
#from wx.lib.anchors import LayoutAnchors
from win32com.client import Dispatch
import numpy as np
import matplotlib
matplotlib.use('QT5Agg')

#from odict import *
#from WXtrap import trap
from MotionPolys import QuinticFit, SepticFit
#from PhysicsHelper import Frame1 as Physics
#from TriangularIndex import Frame1 as TriangleIndex
#from EllipticalPnP import Frame1 as PnP

import collections
import matplotlib.pyplot as plt

from matplotlib import rcParams

from pylab import show
#import pdb

#global motionparams, RawdataList, InstanceList, steps

#####################################
motionparams = collections.OrderedDict()
InstanceList = []
RawDataList = []
steps = 50


#####################################
xl = Dispatch("Excel.Application")



'''[wxID_PANEL1, wxID_PANEL1ACCPLOT, wxID_PANEL1APPENDSEL, wxID_PANEL1BUTTON1, 
 wxID_PANEL1BUTTON11, wxID_PANEL1BUTTON12, wxID_PANEL1BUTTON5, 
 wxID_PANEL1BUTTON6, wxID_PANEL1CHECKBOX1, wxID_PANEL1CONSULTISAAC, 
 wxID_PANEL1DISPPLOT, wxID_PANEL1GETXL, wxID_PANEL1HELP, 
 wxID_PANEL1MAKETRANSPORT, wxID_PANEL1PANEL1, wxID_PANEL1PANEL2, 
 wxID_PANEL1PANEL3, wxID_PANEL1PANEL4, wxID_PANEL1PANEL5, wxID_PANEL1PANEL6, 
 wxID_PANEL1PLOTTITLE, wxID_PANEL1PORTFROMGOALSEEK, wxID_PANEL1STARTINGZEROS, 
 wxID_PANEL1STATICTEXT1, wxID_PANEL1STATICTEXT2, wxID_PANEL1STATICTEXT3, 
 wxID_PANEL1STATICTEXT4, wxID_PANEL1STATICTEXT5, wxID_PANEL1VELPLOT, 
 wxID_PANEL1WRITEXL, wxID_PANEL1XLABEL, wxID_PANEL1YLABEL, 
] = [wx.NewId() for _init_ctrls in range(32)]'''


class Profile:
    '''
    Compiles disp,velocity and acceleration vs time for each profile
    
    '''
    
    def __init__(self, name):
        self.name = name
        #global segmentname
        self.Taxis = []
        self.DispAxis = []
        self.VelAxis = []
        self.AccAxis = []
        self.Fit = []
        
        xl.Range("D7").Select() #Get profile name
        GwareData.segmentname = xl.ActiveCell.Value
        GwareData.segmentname = str(GwareData.segmentname)
        self.makedata()        
            
    def makedata(self):        
        for key in GwareData.motionparams.keys():
            if GwareData.segmentname in key:
                obj = GwareData.motionparams[key]
                for t in obj.trange:
                    self.Taxis.append(t)
                for x in obj.xplot:
                    self.DispAxis.append(x)
                for v in obj.vplot:
                    self.VelAxis.append(v)       
                for a in obj.aplot:
                    self.AccAxis.append(a)
                for thisfit in obj.fit:
                    self.Fit.append(thisfit)

class ImportedTableInstance:
    '''Class object for the imported XY data
    '''
    def __init__(self, name, Time,Displacement,Velocity,Acceleration):
        self.name = name
        self.Time = Time
        self.Displacement=Displacement
        self.Velocity = Velocity
        self.Acceleration=Acceleration
        #print self.Time
        self.makedata()
    
    def makedata(self):
        self.Taxis = (self.Time)
        self.DispAxis=(self.Displacement)
        self.VelAxis = (self.Velocity)
        self.AccAxis = (self.Acceleration)
        
            
class ImportedTableRawData:
    '''Class object for the Raw imported XY data
    '''
    def __init__(self, name, data):
        self.name = name
        self.RawData = data

    
class GetRawData:
    '''
    Iterates throught the active profile region of the worksheet and creates
    a class instance capturing the current object name and cell values used to
    create the motion profile.
    '''
    def __init__ (self, name):
        self.name = name
        xl.Workbooks("TimingSpace.xlsx").Worksheets("Sheet1").Activate()
        DeclareXLDirections()
        xl.Range("D9").Activate()
        Cell = xl.Range("D9").Value
        if Cell is not None:
            xl.Selection.End(xlToRight).Select()
            xl.Selection.End(xlDown).Select()
            DataRange = xl.Range(xl.Selection, xl.Cells(9, 4))
            self.RawData = [column for column in DataRange.Columns()]
        if Cell == None:
            self.RawData = [0]
            
def makeplotdata():
    '''
    '''

    global Taxis,DispAxis,VelAxis,AccAxis,RawTimes,RawDisplacement
    Taxis = []
    DispAxis = []
    VelAxis = []
    AccAxis = []
        
    segname()
    rcParams['ytick.labelsize'] = 12
    rcParams['xtick.labelsize'] = 12
    rcParams['font.size'] = 12
    
    for key in motionparams.iterkeys():
        
        if segmentname in key:
            obj = motionparams[key]
            for t in obj.trange:
                Taxis.append(t)       
            for x in obj.xplot:
                DispAxis.append(x)
            for v in obj.vplot:
                VelAxis.append(v)
            for a in obj.aplot:
                AccAxis.append(a)
                
    DeclareXLDirections()
        
    xl.Workbooks("TimingSpace.xlsx").Worksheets("Sheet1").Activate()
    Cell = xl.Range("D9").Activate()
    if Cell is not None:
        global RawTimes
        xl.Range(xl.Selection, xl.Selection.End(xlToRight)).Select()
        inputrange = xl.Selection
        obj = inputrange.Columns()
        RawTimes = [item for item in obj[0]]            #print RawTimes
            
        
    Cell = xl.Range("D10").Activate()
    if Cell is not None:
        xl.Range(xl.Selection, xl.Selection.End(xlToRight)).Select()
        inputrange = xl.Selection
        RawDisplacement = [time for time in inputrange]
    

def activateshelf():
    '''
    Called By various processes whenever the database needs to be activated
    for editing.
    Opens the database file referenced by cell B3 in the Excel
    worksheet
    Writeback is True for updating-----
    '''
    GwareData.databaseopen=1
    global database, RawDataList, InstanceList, motionparams
    path = 'c:\\Python27\\'
    xl.Workbooks("TimingSpace.xlsx").Worksheets("Sheet1").Activate()
    xl.Range("B3").Select()
    ShelveName = str(xl.ActiveCell.Value)
    suffix = '.db'
    filename = path + ShelveName + suffix
    database = shelve.open(filename, writeback=True)
    InstanceList = database['InstanceList']
    RawDataList = database['RawDataList']              


def DataPresent(): #Got data?---
    global Datapresent
    xl.Workbooks("TimingSpace.xlsx").Worksheets("Sheet1").Activate()
    #sheet=xl.Worksheets("Sheet1").Activate
    xl.Range("D9").Select()
    contents = xl.Selection()
               
    if not contents == None:
        Datapresent = 1
        
    if contents == None:
        Datapresent = 0
    

def ShelveIt():
    path = 'c:\\Python27\\'
    xl.Workbooks("TimingSpace.xlsx").Worksheets("Sheet1").Activate()
    xl.Range("B3").Select()
    ShelveName = str(xl.ActiveCell.Value)
    #suffix = '.db'
    filename = path + ShelveName
    database = shelve.open(filename,writeback=True)
    database['InstanceList'] = InstanceList
    database['RawDataList'] = RawDataList
    #database['motionparams']=motionparams
    database.sync() #8/25/16
    database.close()

def segname():
    '''
    Gets the User entered Segment name from Excel @(D7)
    '''
    #global segmentname
    xl.Workbooks("TimingSpace.xlsx").Worksheets("Sheet1").Activate()
    xl.Range("D7").Select()
    GwareData.segmentname = xl.ActiveCell.Value
    GwareData.segmentname = str(GwareData.segmentname)

def DeclareXLDirections():
    '''Internal Declaration'''
    global xlToLeft, xlToRight, xlUp, xlDown
    xlToLeft = 1
    xlToRight = 2 
    xlUp = 3 
    xlDown = 4

def GetProfileName():
    ''' 
    Gets the Segment name from the worksheet, and then calls BuildSegment
        ---> called from GetXLProfile
    '''
    global Segment, key
    xl = Dispatch("Excel.Application")
    xl.Workbooks("TimingSpace.xlsx").Worksheets("Sheet1").Activate()
    xl.Range("D7").Select()
    Segment = xl.ActiveCell.Value
    
    if Segment == 'NoName':            
        pass
            
    xl.Range("D7").Select()
    xl.ActiveCell.Value = Segment
    Segment = str(Segment)
    key = Segment
    xl.Range("D9").Select() #select upper left of profile range before leaving
    
    BuildSegment(Segment)
    
def BuildSegment(Segment):
    
    ''' Beginning at the active cell, this function loops through 
        the worksheet compiling 8 values to pass as arguments into 
        the FitPoly class(es). Each loop creates a sequentially numbered child 
        of the Main Segment name. Each segment is an instance of
        the Fitpoly class and is stored in the motionparams 
        dictionary under the segment name key.
    '''
    global count, motionparams
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
        global motionparams       
        Segment = key + str(count)
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
        
        #print variables
        
        #GetCheckBox()
        
        if SeventhOrder:
            j = float(0)
            J = float(0)
            motionparams[Segment] = SepticFit(t, x, v, a, j, T, X, V, A, J, steps)
        elif not SeventhOrder:
            motionparams[Segment] = QuinticFit(t, x, v, a, T, X, V, A, steps)
        
        ColumnTop()
        count += 1


       
            
        
        

        
        
        

    
        
        
        
        
        
            
        
        

    
