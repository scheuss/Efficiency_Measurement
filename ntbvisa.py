"""
Title:
Description:
Comments:

Author: fkyburz
"""

import logging
import visa
##import ntbdcload

__version__ = '0.0'

__author__ = 'Falk Kyburz'
__license__ = 'Confidential'
__status__ = 'Development'
__date__ = '2017-08-25'

# Provide Logging Facility
chLogLevel = logging.INFO
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
ch = logging.StreamHandler()
ch.setLevel(chLogLevel)
formatter = logging.Formatter('%(levelname)s - %(message)s')
ch.setFormatter(formatter)
logger.addHandler(ch)


def list_resources():
    # Visa resource manager instance for debug
    rm = visa.ResourceManager()
    # Print resources list for debugging
    print(rm.list_resources())


class NTBSetup:

    def __init__(self, resource_list):
        self.resource_list = resource_list
        self.reset_all()
        self.configure_all()

    def reset_all(self):
        for resource in self.resource_list:
            resource.reset()
        logging.debug('All resources reset')

    def configure_all(self):
        self.reset_all()
        for resource in self.resource_list:
            resource.configure()
        logging.debug('All resources configured')
        
    def write_all(self, command):
        for resource in self.resource_list:
            resource.write(command)

    def close_all(self):
        for resource in self.resource_list:
            resource.close()
        logging.debug('All resources closed')


class NTBResource:

    def __init__(self, visa_name, config_message_list,
                 read_termination='\n', write_termination='\n'):

        self.name = visa_name
        self.config = config_message_list
        self.read_termination = read_termination
        self.write_termination = write_termination
        self.resource = None
        self.open()

    def reset(self):
        self.resource.write('*RST')

    def configure(self):
        self.write_multi(self.config)

    def open(self):
        ''' Open instrument with visa interface '''
        logger.debug('Try to open {0}'.format(self.name))
        rm = visa.ResourceManager()
        self.resource = rm.open_resource(self.name)
        self.resource.read_termination = self.read_termination
        self.resource.write_termination = self.write_termination
        logger.debug('Success')
        logger.info(self.resource.query('*IDN?'))

    def close(self):
        self.resource.close()

    def query(self, message, mode = 'string'):
        temp = None
        if mode is 'string':
            temp = self.resource.query(message)
        if mode is 'values':
            temp = self.resource.query_ascii_values(message)
        return temp

    def write(self, message):
        self.resource.write(message)
        
    def read_raw(self):
        return self.resource.read_raw()

    def write_multi(self, message_list):
        ''' Send all messages in a list '''
        for message in message_list:
            self.resource.write(message)


class NTBResourceDCLoad:

    def __init__(self, com_port, baudrate, address, config_function):

        self.com_port = com_port
        self.baudrate = baudrate
        self.address = address
        self.config = config_function
        self.resource = None
        self.open()

    def reset(self):
        self.turnLoadOff()

    def configure(self):
        self.config(self.resource)

    def open(self):
        '''' Open DCLoad '''
        logger.debug('Try to open {0}'.format(self.com_port))
        self.resource = ntbdcload.DCLoad()
        self.resource.initialize(self.com_port, self.baudrate, self.address)
        logger.debug('Success')
        logger.info(self.resource.getProductInformation())

    def close(self):
        self.resource.turnLoadOff()
        self.resource.setLocalControl()
        self.resource.close()

    # Wrapper functions dor DLoad class
    def turnLoadOn(self):
        '''Turns the load on'''
        return self.resource.turnLoadOn()

    def turnLoadOff(self):
        '''Turns the load on'''
        return self.resource.turnLoadOff()

    def setRemoteControl(self):
        '''Sets the load to remote control'''
        return self.resource.setRemoteControl()

    def setLocalControl(self):
        '''Sets the load to local control'''
        return self.resource.setLocalControl()

    def setMaxCurrent(self, current):
        '''Sets the maximum current the load will sink'''
        return self.resource.setMaxCurrent(current)

    def getMaxCurrent(self):
        '''Sets the maximum current the load will sink'''
        return self.resource.getMaxCurrent()

    def setMaxVoltage(self, voltage):
        '''Sets the maximum voltage the load will allow'''
        return self.resource.setMaxVoltage(voltage)

    def getMaxVoltage(self):
        '''Gets the maximum voltage the load will allow'''
        return self.resource.setMaxVoltage()

    def setMaxPower(self, power):
        '''Sets the maximum power the load will allow'''
        return self.resource.setMaxPower(power)

    def getMaxPower(self):
        '''Gets the maximum power the load will allow'''
        return self.resource.getMaxPower()

    def setMode(self, mode):
        '''Sets the mode (constant current, constant voltage, etc.'''
        return self.resource.setMode(mode)

    def getMode(self):
        '''Gets the mode (constant current, constant voltage, etc.'''
        return self.resource.getMode()

    def setCCCurrent(self, current):
        '''Sets the constant current mode's current level'''
        return self.resource.setCCCurrent(current)

    def getCCCurrent(self):
        '''Gets the constant current mode's current level'''
        return self.resource.getCCCurrent()

    def setCVVoltage(self, voltage):
        '''Sets the constant voltage mode's voltage level'''
        return self.resource.setCVVoltage(voltage)

    def getCVVoltage(self):
        '''Gets the constant voltage mode's voltage level'''
        return self.resource.getCVVoltage()

    def setCWPower(self, power):
        '''Sets the constant power mode's power level'''
        return self.resource.setCWPower(power)

    def getCWPower(self):
        '''Gets the constant power mode's power level'''
        return self.resource.getCWPower()

    def setCRResistance(self, resistance):
        '''Sets the constant resistance mode's resistance level'''
        return self.resource.setCRResistance(resistance)

    def getCRResistance(self):
        '''Gets the constant resistance modes resistance level'''
        return self.resource.getCRResistance()

    def setTransient(self, mode, A, A_time_s, B, B_time_s,
                     operation="continuous"):
        '''Sets up the transient operation mode.  mode is one of
           CC", "CV", "CW", or "CR".
        '''
        return self.resource.setTransient(self,
                                          mode, A, A_time_s, B, B_time_s,
                                          operation)

    def getTransient(self, mode):
        '''Gets the transient mode settings'''
        return self.resource.getTransient(mode)

    def setBatteryTestVoltage(self, min_voltage):
        '''Sets the battery test voltage'''
        return self.resource.setBatteryTestVoltage(min_voltage)

    def getBatteryTestVoltage(self):
        '''Gets the battery test voltage'''
        return self.resource.getBatteryTestVoltage()

    def setLoadOnTimer(self, time_in_s):
        '''Sets the time in seconds that the load will be on'''
        return self.resource.setLoadOnTimer(time_in_s)

    def getLoadOnTimer(self):
        '''Gets the time in seconds that the load will be on'''
        return self.resource.getLoadOnTimer()

    def setLoadOnTimerState(self, enabled=0):
        '''Enables or disables the load on timer state'''
        return self.resource.setLoadOnTimerState(enabled)

    def getLoadOnTimerState(self):
        '''Gets the load on timer state'''
        return self.resource.getLoadOnTimerState()

    def setCommunicationAddress(self, address=0):
        '''Sets the communication address.  Note:  this feature is
        not currently supported.  The communication address should always
        be set to 0.
        '''
        return self.resource.setCommunicationAddress(address)

    def enableLocalControl(self):
        '''Enable local control (i.e., key presses work) of the load'''
        return self.resource.enableLocalControl()

    def disableLocalControl(self):
        '''Disable local control of the load'''
        return self.resource.disableLocalControl()

    def setRemoteSense(self, enabled=0):
        '''Enable or disable remote sensing'''
        return self.resource.setRemoteSense(enabled)

    def getRemoteSense(self):
        '''Get the state of remote sensing'''
        return self.resource.getRemoteSense()

    def setTriggerSource(self, source="immediate"):
        '''Set how the instrument will be triggered.
        "immediate" means triggered from the front panel.
        "external" means triggered by a TTL signal on the rear panel.
        "bus" means a software trigger (see TriggerLoad()).
        '''
        return self.resource.setTriggerSource(source)

    def getTriggerSource(self):
        '''Get how the instrument will be triggered'''
        return self.resource.getTriggerSource()

    def triggerLoad(self):
        '''Provide a software trigger.  This is only of use when the trigger
        mode is set to "bus".
        '''
        return self.resource.triggerLoad()

    def saveSettings(self, register=0):
        '''Save instrument settings to a register'''
        return self.resource.saveSettings(register)

    def recallSettings(self, register=0):
        '''Restore instrument settings from a register'''
        return self.resource.recallSettings(register)

    def setFunction(self, function="fixed"):
        '''Set the function (type of operation) of the load.
        function is one of "fixed", "short", "transient", or "battery".
        Note "list" is intentionally left out for now.
        '''
        return self.resource.setFunction(function)

    def getFunction(self):
        '''Get the function (type of operation) of the load'''
        return self.resource.getFunction()

    def getInputValues(self):
        '''Returns voltage in V, current in A, and power in W, op_state byte,
        and demand_state byte.
        '''
        return self.resource.getInputValues()

    def getProductInformation(self):
        '''Returns model number, serial number, and firmware version'''
        return self.resource.getProductInformation()

