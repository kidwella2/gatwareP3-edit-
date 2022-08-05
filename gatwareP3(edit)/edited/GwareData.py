# -*- coding: utf-8 -*-
"""
Created on Tue Sep 06 11:33:16 2016

@author: Gatman.T
"""
from win32com.client import Dispatch

xl = Dispatch("Excel.Application")
global Taxis, DispAxis, VelAxis, AccAxis, RawTimes, RawDisplacement
global segmentname
global Datapresent, Segment, key
global count, SeventhOrder, NumberOfSegments, database
import collections

motionparams = collections.OrderedDict()

database = 0
Datapresent = None
segmentname = 0
profiles = []
item = 0
parameters = []
d = {}
count=1
variableList={}
lawtype=None
law=None

xlToLeft = 1
xlToRight = 2
xlUp = 3
xlDown = 4
InstanceList = []
selectionlist = []
NoNoList = ['ConstantMotion']
databaseopen = 0
RawDataList = []

steps = 60


if __name__ == "__main__":
    pass
