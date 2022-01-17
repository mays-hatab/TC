from mecom import MeCom
import logging
import time

import matplotlib
import xlsxwriter
matplotlib.use('TkAgg')
from matplotlib import pyplot
import pandas as pd

from datetime import datetime

import sys

logfilename = datetime.now().strftime("%Y%b%d_%H%M%S.log")
print ( logfilename)
logfile = open(logfilename, 'w')

import json

with open("config/calibration.json","r") as fp:
    cal_values = json.load(fp)["Unit3"]


with open("profiles/standard_profile.json","r") as fp:
    thermal_profile = json.load(fp)

def temp_conversion(T_requested, tuning):

    T_offset = tuning["T_offset"]
    fudge_factor = tuning["fudge_factor"]
    T_actual = T_offset  + fudge_factor * ( T_requested )
    return T_actual

def inv_conversion(T, tuning):

    T_offset = tuning["T_offset"]
    fudge_factor = tuning["fudge_factor"]
    return (T-T_offset)/fudge_factor

 

port = 'COM7'
#*******************#


COMMAND_TABLE = {
    "loop status": [1200, ""],
    "object temperature": [1000, "degC"],
    "target object temperature": [1010, "degC"],
    "output current": [1020, "A"],
    "output voltage": [1021, "V"],
    "sink temperature": [1001, "degC"],
    "ramp temperature": [1011, "degC"],
}


class dualTEC:

    def __init__(self):
        self._session = None
        self._connect()
        self.data = []
        self.T0 = time.time()
        self.T = []

    def _connect(self):
        self._session = MeCom(port)
        self.address = self._session.identify()
        print ( self.address, file = logfile)

    def session(self):
        if self._session is None:
            self._connect()
        return self._session

    def get(self,pid, instances = [1,2]):
        s = self.session()
        return [ s.get_parameter(
                    parameter_id = pid,
                    address=self.address, 
                    parameter_instance = i)
                    for i in instances ]

    def set(self,pid,value, instances = [2,1]):
        s = self.session()
        for i in instances:
            s.set_parameter(
                    parameter_id = pid,
                    value =value,
                    address=self.address, 
                    parameter_instance = i)

    def setpoint(self, T, ramp=4 ):
        sp_bottom = temp_conversion(float(T), tuning = cal_values["bottom"]) 
        self.set(3003, float(ramp), [2])
        self.set(3000, sp_bottom, [2])

        sp_top = temp_conversion(float(T), tuning = cal_values["top"]) 
        self.set(3003, float(ramp), [1])
        self.set(3000, sp_top, [1])

    # 1200 loop status
    # 1000 obj temp
    # 1010 target obj temp
    # 1020 curr
    # 1021 volt
    # 1001 sink tmp
    # 1011 ramp tmp
    def getdata(self):
        #print ("    loop         objT          targT         curr         volt           sink         ramp")
     #1.00  1.00   32.29 32.61   50.00 50.00   -0.01 -0.01   -0.20 -0.18   40.00 40.00   50.00 50.00 
        s = self.session()
        all_params = []
        for param in [1200,1000,1010,1020,1021,1001,1011]:
            ret = self.get( param )
            print ("{0:5.2f} {1:5.2f} ".format( *ret), end="  ", flush=True, file=logfile)
            all_params += ret
        print(file=logfile)
        self.data.append( all_params )
        self.T.append(time.time() - self.T0)


    # 2010 status
    def enable(self): self.set(2010, 1 )
    def disable(self): self.set(2010, 0 )

    def update(self):
        pyplot.clf()
        objT1 = [x[2] for x in self.data]
        objT2 = [x[3] for x in self. data]
        targT1 = [x[4] for x in self.data]
        targT2 = [x[5] for x in self.data]

        objT1 = [ inv_conversion(x,cal_values["top"]) for x in objT1]
        objT2 = [ inv_conversion(x,cal_values["bottom"]) for x in objT2]

        targT1 = [ inv_conversion(x,cal_values["top"]) for x in targT1]
        targT2 = [ inv_conversion(x,cal_values["bottom"]) for x in targT2]
        pyplot.plot(self.T, objT1,label="objT1")
        pyplot.plot(self.T, objT2,label="objT2")
        pyplot.plot(self.T, targT1,label="targT1")
        pyplot.plot(self.T, targT2,label="targT2")
        pyplot.legend(loc="upper left")
        pyplot.pause(0.1)


    def monitor(self, duration):
        t0 = time.time()
        steptime = 0.50
        for i in range (int(duration/steptime+0.5)):
            time.sleep ( max (0, i*steptime - (time.time() -t0 ) ) )
            t = time.time() - t0
            self.getdata()
            self.update()

#*****SETTINGS*****#

#Update yellow values as needed

#Activation Step
Activation_Temp = thermal_profile["Annealing"]["temp"]
Activation_RampRate = thermal_profile["Annealing"]["ramprate"] 
Activation_Time = thermal_profile["Annealing"]["time"] 

#Cycle Definition
Cycle_Repetition = thermal_profile["cycles"]
#Denaturing Step
Denaturing_Temp = thermal_profile["Denaturing"]["temp"]
Denaturing_RampRate = thermal_profile["Denaturing"]["ramprate"]
Denaturing_Time = thermal_profile["Denaturing"]["time"]

#Annealing Step
Annealing_Temp = thermal_profile["Annealing"]["temp"]
Annealing_RampRate = thermal_profile["Annealing"]["ramprate"]
Annealing_Time = thermal_profile["Annealing"]["time"]

#Equilibration Step
Equilibration_Temp = thermal_profile["Equilibration"]["temp"]
Equilibration_RampRate = thermal_profile["Equilibration"]["ramprate"]
Equilibration_Time = thermal_profile["Equilibration"]["time"]



def holding_time():
    holding_time = ((Denaturing_Temp - Annealing_Temp)/Denaturing_RampRate) + 4.0 
    return holding_time

def PCR(tec):
  
    for cycle in range(Cycle_Repetition): 
        print( "#set 52", file=logfile)
        tec.setpoint(Annealing_Temp, ramp=Annealing_RampRate)
        #tec.setpoint(47.5, ramp=4)
        tec.monitor(holding_time())
        tec.setpoint(Annealing_Temp,ramp=1)
        tec.monitor(Annealing_Time)

        print ( "#set 98 ", file = logfile)
        tec.setpoint(Denaturing_Temp, ramp=Denaturing_RampRate)
        #tec.setpoint(85.5, ramp=4)
        tec.monitor(holding_time())
        tec.setpoint(Denaturing_Temp, ramp=1)
        tec.monitor(Denaturing_Time)


    
def square_wave(tec):
    for ramp in [4,6,8,10,12,14,16]: 
        tec.setpoint(90.0, ramp=ramp)
        tec.monitor(10+40/ramp)

        tec.setpoint(50.0, ramp=ramp)
        tec.monitor(10+40/ramp)

def step_response(tec):

    for setpoint in [60]:
        tec.setpoint(setpoint, ramp=20)
        tec.monitor(20)

        tec.setpoint(50.0, ramp=20)
        tec.monitor(20)



def main():

    tec = dualTEC()

    print ( "#enable", file=logfile)
    tec.enable()

    print ( "#set 99", file=logfile)
    tec.setpoint(Activation_Temp, ramp = Activation_RampRate)
    tec.monitor(Activation_Time)

    PCR(tec)

    print ( "#set 20", file=logfile)
    tec.setpoint(Equilibration_Temp, ramp = Equilibration_RampRate)
    
    tec.monitor(Equilibration_Time)

    print ( "#disable", file=logfile)
    tec.disable()
    tec.monitor(20)

    pyplot.savefig ("graph.png")
    # df = pd.read 
    # df = df.iloc[:,4:]
    # global_num = df.sum()
    # writer = self.ExcelWriter('python_plot.xlsx', engine = 'xlsxwriter')
    # global_num.to_excel(writer, sheet_name = 'Sheet1')
    # worksheet = writer.sheets['Sheet1']
    # worksheet.insert_image('C2','graph.png')
    # writer.save()


if __name__ == "__main__":
    main()
