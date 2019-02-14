"""
Title:       Automated data acquisition script to get the efficiency
Description: Used for 3-to-1 ladder converter
Comments: 
Author:      R.Scheuss
Date:        2019-01-26
Version:     0.0
"""
import sys
import os
import time
import dcload
from ntbvisa import *
#import serial
import matplotlib.pyplot as plt

# DC-Load COM-port
DCLOAD_COMPORT = "COM8"
DCLOAD_BAUD = 38400

# parameters for load sweep
startCurrent = 1000     #in mA
endCurrent   = 9000     #in mA
stepSize     =  1000     #in mA
shuntGainIin =    1
shuntGainIout=    1

# Set DMM names
dmm_uin_name =  'TCPIP::128.138.189.186::3490::SOCKET'
dmm_uout_name = 'TCPIP::128.138.189.69::3490::SOCKET'

dmm_iin_name =  'TCPIP::128.138.189.39::3490::SOCKET'
dmm_iout_name = 'TCPIP::128.138.189.162::3490::SOCKET'


def main():
    ###############################################################################
    # Provide Logging Facility
    ###############################################################################
    # create logger instance
    logger = logging.getLogger(__name__)
    logger.setLevel(logging.DEBUG)
    
    # create console handler and set level to debug
    ch = logging.StreamHandler()
    ch.setLevel(chLogLevel)
    
    # create formatter
    formatter = logging.Formatter('%(levelname)s - %(message)s')
    
    # add formatter to ch
    ch.setFormatter(formatter)
    
    # add ch to logger
    logger.addHandler(ch)
    #------------------------------------------------------------------------------
    
    # create instance
    load = dcload.DCLoad()
    
    # list for efficiency values for plot
    efficiency = []
    current    = []
         
    # DMM Config
    dmm_u_config = ['CONF:VOLT:DC',
                          'SENS:VOLT:DC:RANG 100',          #up tp 100V
                          'TRIG:DEL 0',                     #Set the delay between trigger and measurement
                          'TRIG:SOUR IMM',                  #Set meter’s trigger source
                          'SAMP:COUN 1']                    #Set number of samples per trigger
    
    dmm_i_config = ['CONF:CURR:DC',
                          'SENS:CURR:DC:RANG 10',           #up tp 10V
                          'TRIG:DEL 0',                     #Set the delay between trigger and measurement
                          'TRIG:SOUR IMM',                  #Set meter’s trigger source
                          'SAMP:COUN 1']                    #Set number of samples per trigger
                          
    #set up digital multimeter
    time.sleep(1)
    dmm_uin = NTBResource(dmm_uin_name, dmm_u_config)
    time.sleep(1)
    dmm_uout = NTBResource(dmm_uout_name, dmm_u_config)
    time.sleep(1)
    dmm_iin = NTBResource(dmm_iin_name, dmm_i_config)
    time.sleep(1)
    dmm_iout = NTBResource(dmm_iout_name, dmm_i_config)
    time.sleep(1)
    #dmm_v_name = 'TCPIP0::mmies006::INSTR'     NTB-style
        
    setup = NTBSetup([dmm_uin, dmm_uout, dmm_iin, dmm_iout])
    time.sleep(1)
    
    #set up DC load
    print("DC-load, init")
    load.initialize(DCLOAD_COMPORT, DCLOAD_BAUD) # Open a serial connection
    print("DC-load, set remote control", load.setRemoteControl())
    print("DC-load, set max voltage to 15V", load.setMaxVoltage(15))
    print("DC-load, set max power to 300W", load.setMaxPower(300))
    print("DC-load, to constant current mode", load.setMode('cc'))
    print("DC-load, set first current", load.setCCCurrent(startCurrent))
    print("DC-load, turn on", load.turnLoadOn())
    
    # Open log file
    try:
        filename = sys.argv[1]
    except:
        timestr = time.strftime("%Y%m%d-%H%M%S")   
        filename = timestr + ".txt"
        
    if not os.path.isfile(filename):
        with open(filename, 'w') as logdata: 

            # Print header
            row_head = ("Time Uin[V] Iin[A] Pin[W] Uout[V] Iout[A] Pout[W] n[]")
            print(row_head, file=logdata)
            print(row_head)
            
            try:
                for actualCurrent in range(startCurrent, endCurrent + stepSize, stepSize):
                    #print("Set current to %i mA" % actualCurrent)
                    load.setCCCurrent(actualCurrent)
                    time.sleep(2)                               #wait until steady state

                    # Arm and trig the instruments
                    setup.write_all('INIT')
                    #setup.write_all('*TRG')
                            
                    # Read out the measurment values
                    res_uin_raw = dmm_uin.query('FETC?', 'values')[0]
                    res_uout_raw = dmm_uout.query('FETC?', 'values')[0]
                    res_iin_raw = dmm_iin.query('FETC?', 'values')[0]
                    res_iout_raw = dmm_iout.query('FETC?', 'values')[0]
                    
                    res_time = time.strftime('%H:%M:%S')                    
                    
                    #scale if shunt is used
                    res_uin = res_uin_raw 
                    res_uout = res_uout_raw
                    res_iin = res_iin_raw * shuntGainIin
                    res_iout = res_iout_raw * shuntGainIout

                    #calculate power and efficiency
                    res_pin = res_uin * res_iin
                    res_pout = res_uout * res_iout
                    
                    try:                                #if division thru zero
                        res_eff = res_pout / res_pin
                    except:
                        res_eff = 0
                    
                    #store efficiency for plot
                    efficiency.append(res_eff)
                    current.append(actualCurrent/1000)

                    
                    #Save measurements in logfile
                    row_content = '{0} {1} {2} {3} {4} {5} {6} {7}' \
                    .format(res_time, res_uin, res_iin, res_pin, res_uout, res_iout, res_pout, res_eff)
                    print(row_content, file=logdata)
                    
                    row_content_console = ("{0},Uin={1:2.3f}V,Iin={2:1.3f}A,Pin={3:3.3f}W,Uout={4:2.3f}V,Iout={5:2.3f}A,Pout={6:3.3f}W,n={7:1.4f}") \
                    .format(res_time, res_uin, res_iin, res_pin, res_uout, res_iout, res_pout, res_eff)
                    print(row_content_console)
                    
                    
            except KeyboardInterrupt:
                print('Aborted')
                
  
        print("close file and disconnect digital multimeter")    
        logdata.close()
        setup.close_all()
        
        #ramp down current
        print("Ramp down current")
        lastCurrentSetting = int(load.getCCCurrent())

        for actualCurrent in range(lastCurrentSetting, startCurrent-stepSize, -stepSize):
            print("Set current to %i A" % actualCurrent)
            load.setCCCurrent(actualCurrent)
            time.sleep(0.1)
        
        print("turn off load and set local control")
        print(load.turnLoadOff())
        print(load.setLocalControl())
        
        plt.figure("Efficiency")
        plt.title("Efficiency")
        plt.xlabel('Current [A]')
        plt.ylabel('Efficiency []')
        plt.plot(current, efficiency)
        plt.grid(b=True, which='major', color='b', linestyle='-')
        plt.show()
        
            
    else:
        print('file already exists')


if __name__== "__main__":
    main()
