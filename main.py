from mecom import MeCom
import logging
import time

import matplotlib
import xlsxwriter
matplotlib.use('TkAgg')
from matplotlib import pyplot
import pandas as pd

tuning = {    
    "bottom":{
        "T_offset" : 0.5224,
        "fudge_factor" : 0.8228,  
    },

    "top":{
        "T_offset" : 1.5425,
        "fudge_factor" : 0.7747,
    }
}


def temp_conversion(T_requested, tuning = tuning["bottom"]):

    T_offset = tuning["T_offset"]
    fudge_factor = tuning["fudge_factor"]
    T_actual = T_offset  + fudge_factor * ( T_requested )
    return T_actual
 
#*****SETTINGS*****#

#Update yellow values as needed

#Activation Step
Activation_Temp = 99 #degrees Celcius
Activation_RampRate = 4 #degrees/second
Activation_Time = 60 #seconds

#Cycle Definition
Cycle_Repetition = 42 #times that the cycle will repeat

#Denaturing Step
Denaturing_Temp = 98 #degrees Celcius
Denaturing_RampRate = 4 #degrees/second+
Denaturing_Time = 6 #seconds

#Annealing Step
Annealing_Temp = 60 #degrees Celcius
Annealing_RampRate = 4 #degrees/second
Annealing_Time = 10 #seconds

#Equilibration Step
Equilibration_Temp = 15 #degrees Celcius
Equilibration_RampRate = 4 #degrees/second
Equilibration_Time = 60 #seconds

port = 'COM5'
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
        print ( self.address)

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
        sp_bottom = temp_conversion(float(T), tuning = tuning["bottom"]) 
        self.set(3003, float(ramp))
        self.set(3000, sp_bottom)

        sp_top = temp_conversion(float(T), tuning = tuning["top"]) 
        self.set(3003, float(ramp))
        self.set(3000, sp_top)

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
            print ("{0:5.2f} {1:5.2f} ".format( *ret), end="  ", flush=True)
            all_params += ret
        print()
        self.data.append( all_params )
        self.T.append(time.time() - self.T0)

    # 2010 status
    def enable(self): self.set(2010, 1 )
    def disable(self): self.set(2010, 0 )

    def update(self):
        pyplot.clf()
        pyplot.plot(self.T, [ x[2] for x in self.data],label="objT1")
        pyplot.plot(self.T, [ x[3] for x in self.data],label="objT2")
        pyplot.plot(self.T, [ x[4] for x in self.data],label="targT1")
        pyplot.plot(self.T, [ x[5] for x in self.data],label="targT2")
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




def holding_time():
    holding_time = ((Denaturing_Temp - Annealing_Temp)/Denaturing_RampRate) + 4.0 
    return holding_time

def PCR(tec):
  
    for cycle in range(Cycle_Repetition): 
        print( "#set 52")
        tec.setpoint(Annealing_Temp, ramp=Annealing_RampRate)
        #tec.setpoint(47.5, ramp=4)
        tec.monitor(holding_time())
        tec.setpoint(Annealing_Temp,ramp=1)
        tec.monitor(Annealing_Time)

        print ( "#set 98 ")
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

    print ( "#enable")
    tec.enable()

    print ( "#set 99")
    tec.setpoint(Activation_Temp, ramp = Activation_RampRate)
    tec.monitor(Activation_Time)

    PCR(tec)

    print ( "#set 20")
    tec.setpoint(Equilibration_Temp, ramp = Equilibration_RampRate)
    
    tec.monitor(Equilibration_Time)

    print ( "#disable")
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
