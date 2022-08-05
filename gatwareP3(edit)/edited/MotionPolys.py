# -----------------------------------------------------------------------------
# Name:        MotionPolys.py
# Purpose:
#
# Author:      <Todd Gatman>
#
# Created:     2011/10/13
# RCS-ID:      $Id: MotionPolys.py $
# Copyright:   (c) 2006
# Licence:     <your licence>
# -----------------------------------------------------------------------------


import numpy as np
import matplotlib.pyplot as plt
# from sympy import *
from pylab import figure
from scipy import linalg, matrix
import GwareData
from scipy.interpolate import *
from XLSX_Initializer import deriv
from math import pi,sin

class QuinticFit:
    '''Fifth Order Polynomial Fitting Routine

    Usage ---> "InstanceName=QuinticFit(t, x, v, a, T, X, V, A,steps)
    *************************
    t,T=time
    x,X=Displacement
    v,V=Velocity
    a,A=Acceleration
    *************************
    where t,x,v,a correspond to the initial conditions;
    and   T, X, V, A represent the final conditions for the move profile
    steps=Number of calculation steps in the polynomial
    '''

    def __init__(self, t, x, v, a, T, X, V, A, steps):
        self.steps = steps
        self.t = t
        self.x = x
        self.v = v
        self.a = a

        self.T = T
        self.X = X
        self.V = V
        self.A = A

        self.Solve(t, x, v, a, T, X, V, A, steps)

    def Solve(self, t, x, v, a, T, X, V, A, steps):
        Ar = np.array([[self.t ** 5, self.t ** 4, self.t ** 3, self.t ** 2,
                        self.t, 1],
                       [5 * self.t ** 4, 4 * self.t ** 3, 3 * self.t ** 2,
                        2 * self.t, 1, 0],
                       [20 * self.t ** 3, 12 * self.t ** 2, 6 * self.t, 2, 0, 0],
                       [self.T ** 5, self.T ** 4, self.T ** 3, self.T ** 2,
                        self.T, 1],
                       [5 * self.T ** 4, 4 * self.T ** 3, 3 * self.T ** 2,
                        2 * self.T, 1, 0],
                       [20 * self.T ** 3, 12 * self.T ** 2, 6 * self.T, 2, 0, 0]])

        # self.M=Matrix(Ar)

        B = np.array([self.x, self.v, self.a, self.X, self.V, self.A])
        self.B = B
        C = np.linalg.solve(Ar, B)

        self.trange = np.linspace(self.t, self.T, self.steps)

        # =======================================================================
        # Make data vectors for x,v,a,j from newly solved equations
        # =======================================================================

        def Dispx(t):
            return (C[0] * t ** 5 + C[1] * t ** 4 + C[2] * t ** 3 +
                    C[3] * t ** 2 + C[4] * t + C[5])

        def Vel(t):
            return (5 * C[0] * t ** 4 + 4 * C[1] * t ** 3 + 3 * C[2] * t ** 2 +
                    2 * C[3] * t + C[4])

        def Acc(t):
            return (C[0] * 20 * t ** 3 + C[1] * 12 * t ** 2 + C[2] * 6 * t +
                    C[3] * 2)

        def j(t):
            return C[0] * 60 * t ** 2 + C[1] * 24 * t + C[2] * 6

        self.xplot = [Dispx(time) for time in self.trange]
        self.vplot = [Vel(time) for time in self.trange]
        self.aplot = [Acc(time) for time in self.trange]
        self.jplot = [j(time) for time in self.trange]
        self.pplot=[i*j for i,j in list(zip(self.vplot,self.aplot))]
        self.fit = [('FifthOrder')]
        self.junk = "Junk"
        self.C = C
        self.Solutions = C
        self.Coeff = [C[0], C[1], C[2], C[3], C[4], C[5]]
        self.Disp = (C[0] * t ** 5 + C[1] * t ** 4 + C[2] * t ** 3 +
                     C[3] * t ** 2 + C[4] * t + C[5])
        self.vartable = self.__dict__


class SepticFit:
    '''Seventh Order Polynomial Fitting Routine
    '''

    def __init__(self, t, x, v, a, j, T, X, V, A, J, steps):
        self.steps = steps
        self.ti = t
        self.x = x
        self.v = v
        self.a = a
        self.j = j
        self.Tf = T
        self.X = X
        self.V = V
        self.A = A
        self.J = J

        self.Deltat = self.Tf - self.ti
        self.t = float(0)
        self.T = self.t + self.Deltat
        # print self.t,self.x,self.v,self.a,self.j,self.T,self.X,self.V,self.A,self.J
        self.Solve(t, x, v, a, j, T, X, V, A, J, steps)

    def Solve(self, t, x, v, a, j, T, X, V, A, J, steps):
        Ar = np.array([[self.t ** 7, self.t ** 6, self.t ** 5, self.t ** 4,
                        self.t ** 3, self.t ** 2, self.t, 1],
                       [7 * self.t ** 6, 6 * self.t ** 5, 5 * self.t ** 4,
                        4 * self.t ** 3, 3 * self.t ** 2, 2 * self.t, 1, 0],
                       [42 * self.t ** 5, 30 * self.t ** 4, 20 * self.t ** 3,
                        12 * self.t ** 2, 6 * self.t, 2, 0, 0],
                       [210 * self.t ** 4, 120 * self.t ** 3, 60 * self.t ** 2,
                        24 * self.t, 6, 0, 0, 0],
                       [self.T ** 7, self.T ** 6, self.T ** 5, self.T ** 4,
                        self.T ** 3, self.T ** 2, self.T, 1],
                       [7 * self.T ** 6, 6 * self.T ** 5, 5 * self.T ** 4,
                        4 * self.T ** 3, 3 * self.T ** 2, 2 * self.T, 1, 0],
                       [42 * self.T ** 5, 30 * self.T ** 4, 20 * self.T ** 3,
                        12 * self.T ** 2, 6 * self.T, 2, 0, 0],
                       [210 * self.T ** 4, 120 * self.T ** 3, 60 * self.T ** 2,
                        24 * self.T, 6, 0, 0, 0]])

        B = np.array([self.x, self.v, self.a, self.j, self.X, self.V,
                      self.A, self.J])

        C = np.linalg.solve(Ar, B)
        # for thisc in C:
        # print thisc

        self.trange = np.linspace(self.t, self.T, self.steps)

        def Dispx(t):
            return (C[0] * t ** 7 + C[1] * t ** 6 + C[2] * t ** 5 +
                    C[3] * t ** 4 + C[4] * t ** 3 +
                    C[5] * t ** 2 + C[6] * t + C[7])

        def Vel(t):
            return (7 * C[0] * t ** 6 + 6 * C[1] * t ** 5 + 5 * C[2] * t ** 4 +
                    4 * C[3] * t ** 3 + 3 * C[4] * t ** 2 + 2 * C[5] * t +
                    C[6])

        def Acc(t):
            return (42 * C[0] * t ** 5 + 30 * C[1] * t ** 4 +
                    20 * C[2] * t ** 3 + 12 * C[3] * t ** 2 + 6 * C[4] * t +
                    2 * C[5])

        def Jerk(t):
            return (210 * C[0] * t ** 4 + 120 * C[1] * t ** 3 +
                    60 * C[2] * t ** 2 + 24 * C[3] * t + 6 * C[4])

        self.xplot = [Dispx(time) for time in self.trange]
        self.vplot = [Vel(time) for time in self.trange]
        self.aplot = [Acc(time) for time in self.trange]
        self.jplot = [Jerk(time) for time in self.trange]
        self.pplot = [i * j for i, j in list(zip(self.vplot, self.aplot))]
        self.trange = np.linspace(self.ti, self.Tf, self.steps)
        self.fit = [('SeventhOrder')]
        self.Solutions = C
        self.vartable = self.__dict__


def opti(obj):
    global A, Ar, B, X, Tf, cubby, jerk
    resolution = .1
    cubby = ['lastval']
    Ar = obj.M
    B = list(obj.B)
    A = min(obj.aplot)
    Tf = max(obj.trange)  # Ending time of segment

    def reSolve():
        ''' Resolve the polynomial to get a new set of coefficients '''
        global X
        X = np.linalg.solve(Ar, B)

    def j():
        return X[0] * 60 * Tf ** 2 + X[1] * 24 * Tf + X[2] * 6

    def acc():
        # print X[0]
        return (X[0] * 20 * Tf ** 3 + X[1] * 12 * Tf ** 2 + X[2] * 6 * Tf +
                X[3] * 2)

    reSolve()  # first pass evaluates the initial coefficients

    jerk = j()  # Evaluate jerk a the in the initial state

    if jerk < 0:
        resolution = -.1

    # print jerk

    # B[5]=A # substitute new acceleration value into last position of poly arguments

    while 1:
        # print B
        reSolve()  # resolve with revised terminal acceleration
        # print X

        cubby[0] = B[5]  # store last accel value in cubby
        jerk = j()  # Calculate jerk at terminal pos with new coefficients.
        # print jerk
        B[5] = acc() - resolution  # Put the revised terminal acceleration values in the solution matrix
        if -1 < jerk < 1:
            return B[5]
            break
        # reSolve()
        # print acc()
        # count+=1


class Trap:
    ''' calculates velocity and accelerations for a trapezoidal motion profile given time, distance and Lambda
    -->Lambda is the percentage of the total move time you allow for the acc/dec portion of the motion.
    usage-->your profile name=Trap(move time,move distance,Lambda)
    A move time of 1 second with a Lambda value of 0.5 allocates 0.25 seconds for accel and 0.25 seconds for decel

    '''

    def __init__(self, t, x, Lambda):
        self.t = t
        self.x = x
        self.Lambda = Lambda

        self.v = 2 * self.x / ((self.t - self.t * self.Lambda) + self.t)
        self.t1=(self.t*self.Lambda)/2
        self.t2=self.t-2*self.t1
        self.a=self.v/self.t1
        self.x1=0.5*self.a*self.t1**2
        self.t3=self.t-self.t1
        self.x2=self.x-self.x1
        self.x3=self.x-self.x1
        seg1=QuinticFit(0,0,0,0,self.t1,self.x1,self.v,0,50)
        seg3=QuinticFit(self.t3,self.x3,self.v,0,self.t,self.x,0,0,50)
        self.trange=[i for i in seg1.trange]
        self.xplot=[i for i in seg1.xplot]
        self.vplot = [i for i in seg1.vplot]
        self.aplot = [i for i in seg1.aplot]
        for i in range(len(seg3.trange)):
            self.trange.append(seg3.trange[i])
            self.xplot.append(seg3.xplot[i])
            self.vplot.append(seg3.vplot[i])
            self.aplot.append(seg3.aplot[i])
        self.pplot = [i * j for i, j in list(zip(self.vplot, self.aplot))]

class GeneralTrap:
    ''' calculates velocity and accelerations for a trapezoidal motion profile given time, distance and Lambda
    -->Lambda is the percentage of the total move time you allow for the acc/dec portion of the motion.
    -->C is the percentage of the acceleration time is used for the first accel
    usage-->your profile name=Trap(move time,move distance,Lambda,C)
    A move time of 1 second with a Lambda value of 0.5 and C =0.5 allocates 0.25 seconds for accel and 0.25 seconds for decel
    A move time of 1 second with a Lambda value of 0.5 and C= 0.25 allocates .125s for accel1 and 0.375s for accel2

    '''

    def __init__(self,tstart, t, xstart,x, Lambda,C):
        self.tstart=tstart
        self.xstart=xstart
        self.t = t
        self.x = x
        self.Lambda = Lambda
        self.C=C

        self.t1=(self.t*self.Lambda*self.C)
        self.t2=self.t*self.Lambda-self.t1
        self.t3=self.t-self.t2
        '''print('t1= ',self.t1)
        print('t2= ', self.t2)
        print('t3= ', self.t3)'''
        self.v = -((2 * self.x) / (self.t1-2*self.t+self.t2))
        #print('vel= ',self.v)
        self.a=self.v/self.t1
        #print('acceleration= ',self.a)
        self.x1=0.5*self.a*self.t1**2
        #print('x1= ',self.x1)
        self.x2=self.x1+self.v*(self.t3-self.t1)
        #print('x2= ',self.x2)

        self.seg1=QuinticFit(0,0,0,0,self.t1,self.x1,self.v,0,50)
        self.seg3=QuinticFit(self.t3,self.x2,self.v,0,self.t,self.x,0,0,50)
        self.trange=[i for i in self.seg1.trange]
        self.xplot=[i for i in self.seg1.xplot]
        self.vplot = [i for i in self.seg1.vplot]
        self.aplot = [i for i in self.seg1.aplot]
        for i in range(len(self.seg3.trange)):
            self.trange.append(self.seg3.trange[i])
            self.xplot.append(self.seg3.xplot[i])
            self.vplot.append(self.seg3.vplot[i])
            self.aplot.append(self.seg3.aplot[i])
        self.trange=[i+self.tstart for i in self.trange]
        self.xplot=[i+self.xstart for i in self.xplot]
        self.pplot = [i * j for i, j in list(zip(self.vplot, self.aplot))]


class ModTrap:
    """

    """
    def __init__(self,Time,Disp):
        self.Time=Time
        self.Disp=Disp
        self.h=1
        self.B=1
        self.Th = np.linspace(0, self.B, 100)
        self.hlist=[]
        self.calc()
        self.scale()

    def calc(self):
        for t in self.Th:
            if t<=(self.B/8):
                y=(1/(2+pi))*(2*t-(1/(2*pi))*sin(4*pi*t))
                self.hlist.append(y)
            if (self.B/8<t<(3*self.B/8)):
                y=(1/(2+pi))*((-1/(2*pi))+(2*t)+4*pi*(t-(1/8))**2)
                self.hlist.append(y)
            if (3*self.B/8 < t <= (self.B / 2)):
               #print(t)
               y=(1/(2+pi))*((-pi/2)+2*(1+pi)*t-(1/(2*pi)*sin(4*pi*t-pi)))
               self.hlist.append(y)
            if self.B/2<t<=5*self.B/8:
               t=(1-t)
               y = 1-(1/(2+pi))*((-pi/2)+2*(1+pi)*t-(1/(2*pi)*sin(4*pi*t-pi)))
               self.hlist.append(y)
            if 5*self.B/8<t<=7*self.B/8:
               t = (1 - t)
               y =1-(1/(2+pi))*((-1/(2*pi))+(2*t)+4*pi*(t-(1/8))**2)
               self.hlist.append(y)
            if t>7*self.B/8:
               t = (1 - t)
               y =1-(1/(2+pi))*(2*t-(1/(2*pi))*sin(4*pi*t))
               self.hlist.append(y)

    def scale(self):
        self.trange=[self.Time*i for i in self.Th]
        self.xplot=[self.Disp*i for i in self.hlist]
        tck = splrep(self.trange, self.xplot)
        self.vplot = deriv(self.trange, tck, 1)
        self.aplot = deriv(self.trange, tck, 2)
        self.jplot = deriv(self.trange, tck, 3)
        self.pplot = [i * j for i, j in list(zip(self.vplot, self.aplot))]
        self.fit = [('ModTrap')]


def MotionPlot(t, x, v, a, T, X, V, A, steps, Name):
    '''Use to plot X,V,A for an individual motion profile'''
    plottitle = Name
    # print plottitle
    Name = QuinticFit(t, 0, 0, 0, T, X, V, 0, steps)
    fig = figure(1)
    fig.add_subplot(211)
    plt.plot(Name.trange, Name.xplot, label=plottitle + 'X')
    plt.plot(Name.trange, Name.vplot, label=plottitle + 'V')
    plt.grid(b='on', which='both')
    plt.autoscale(enable=True, axis='both', tight=None)
    loc = 'best'
    plt.legend(loc='best')

    fig.add_subplot(212)
    plt.plot(Name.trange, Name.aplot, label=plottitle + 'A')
    plt.grid(b='on', which='both')
    plt.autoscale(enable=True, axis='both', tight=None)
    loc = 'best'
    plt.legend(loc='best')
    plt.draw()
    plt.show()


if __name__ == '__main__':
    global F
    Time = 0

    t = .2
    T = .36

    x = -.1519
    X = .024
    v = -2.4229
    V = 3.3669

    a = 255.1761
    A = 217.5773
    j = 0
    J = 0

    steps = 100
    # Name='Entertainer'
    # MotionPlot(Time, 0, 0, 0,T, X, V, 0, steps,Name)
    # opti(F)

    #Septic = SepticFit(t, x, v, a, j, T, X, V, A, J, steps)

    # plt.plot(F.trange, F.xplot, label='5th X')
    # plt.plot(F.trange, F.vplot, label='5th V')
    # plt.plot(F.trange, F.aplot, label='5th A')
    # plt.plot(Fifth.trange,Fifth.jplot,label='5th J')

    #plt.plot(Septic.trange, Septic.xplot, label='7th X')
    #plt.plot(Septic.trange, Septic.vplot, label='7th V')
    #plt.plot(Septic.trange, Septic.aplot, label='7th A')
    # plt.plot(Septic.trange,Septic.jplot,label='7th J')
    # y=[(Fifth.Solutions[0]*t**5+Fifth.Solutions[1]*t**4+Fifth.Solutions[2]*t**3) for t in time]

    # plt.plot(Septic.trange, Septic.xplot, label='7th')

    #plt.grid(b='on', which='both')
    #plt.autoscale(enable=True, axis='both', tight=None)
    #loc = 'best'
    #plt.legend(loc='best')
    #plt.draw()
    #plt.show()

    trap=GeneralTrap(0,1,0,1000,1,0.5)
    plt.grid()
    plt.plot(trap.trange, trap.xplot)
    plt.plot(trap.trange, trap.vplot)
    plt.plot(trap.trange, trap.aplot)

