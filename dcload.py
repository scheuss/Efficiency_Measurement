'''
Open Source Initiative OSI - The MIT License:Licensing
Tue, 2006-10-31 04:56 - nelson

The MIT License

Copyright (c) 2009 BK Precision

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.


This python module provides a functional interface to a B&K DC load
through the DCLoad object.  This object can also be used as a COM
server by running this module as a script to register it.  All the
DCLoad object methods return strings.  All units into and out of the
DCLoad object's methods are in SI units.

See the documentation file that came with this script.

$RCSfile: dcload.py $
$Revision: 1.0 $
$Date: 2008/05/17 15:57:15 $
$Author:  Don Peterson $
'''

from __future__ import division
import sys
import time
import serial

# from string import join
try:
    from win32com.server.exception import COMException
except:
    pass

# Debugging information is set to stdout by default.  You can change
# the out variable to another method to e.g. write to a different
# stream.
out = sys.stdout.write
nl = "\n"


class InstrumentException(Exception):
    pass


class InstrumentInterface:
    '''Provides the interface to a 26 byte instrument along with utility
    functions.
    '''
    debug = 0  # Set to 1 to see dumps of commands and responses
    length_packet = 26  # Number of bytes in a packet
    convert_current = 1e1  # Convert current in mA to 0.1 mA
    convert_voltage = 1e3  # Convert voltage in V to mV
    convert_power = 1e3  # Convert power in W to mW
    convert_resistance = 1e3  # Convert resistance in ohm to mohm
    to_ms = 1000           # Converts seconds to ms
    # Number of settings storage registers
    lowest_register = 1
    highest_register = 25
    # Values for setting modes of CC, CV, CW, or CR
    modes = {"cc": 0, "cv": 1, "cw": 2, "cr": 3}

    def initialize(self, com_port, baudrate, address=0):
        try:
            self.sp = serial.Serial(com_port, baudrate, timeout=5)
        except:
            print('unable to open port')
        self.address = address
        time.sleep(0.2)

    def close(self):
        self.sp.close()

    def dumpCommand(self, xbytes):
        '''Print out the contents of a 26 byte command.  Example:
            aa .. 20 01 ..   .. .. .. .. ..
            .. .. .. .. ..   .. .. .. .. ..
            .. .. .. .. ..   cb
        '''
        assert(len(xbytes) == self.length_packet)
        header = " "*3
        print(header, end="")
        for i in range(self.length_packet):
            if i % 10 == 0 and i != 0:
                print(nl + header, end="")
            if i % 5 == 0:
                print(" ", end="")
            s = "%02x" % ord(xbytes[i])
            if s == "00":
                # Use the decimal point character if you see an
                # unattractive printout on your machine.
                # s = "."*2
                # The following alternate character looks nicer
                # in a console window on Windows.
                #s = chr(250)*2
                s = '.'*2
            print(s, end="")
        print(nl, end="")

    def commandProperlyFormed(self, cmd):
        '''Return 1 if a command is properly formed; otherwise, return 0.
        '''
        commands = (
            0x20, 0x21, 0x22, 0x23, 0x24, 0x25, 0x26, 0x27, 0x28, 0x29,
            0x2A, 0x2B, 0x2C, 0x2D, 0x2E, 0x2F, 0x30, 0x31, 0x32, 0x33,
            0x34, 0x35, 0x36, 0x37, 0x38, 0x39, 0x3A, 0x3B, 0x3C, 0x3D,
            0x3E, 0x3F, 0x40, 0x41, 0x42, 0x43, 0x44, 0x45, 0x46, 0x47,
            0x48, 0x49, 0x4A, 0x4B, 0x4C, 0x4D, 0x4E, 0x4F, 0x50, 0x51,
            0x52, 0x53, 0x54, 0x55, 0x56, 0x57, 0x58, 0x59, 0x5A, 0x5B,
            0x5C, 0x5D, 0x5E, 0x5F, 0x60, 0x61, 0x62, 0x63, 0x64, 0x65,
            0x66, 0x67, 0x68, 0x69, 0x6A, 0x6B, 0x6C, 0x12
        )
        # Must be proper length
        if len(cmd) != self.length_packet:
            print("Command length = " + str(len(cmd)) + "-- should be " +
                  str(self.length_packet) + nl)
            return 0
        # First character must be 0xaa
        if ord(cmd[0]) != 0xaa:
            print("First byte should be 0xaa" + nl)
            return 0
        # Second character (address) must not be 0xff
        if ord(cmd[1]) == 0xff:
            print("Second byte cannot be 0xff" + nl)
            return 0
        # Third character must be valid command
        byte3 = "%02X" % ord(cmd[2])
        if ord(cmd[2]) not in commands:
            print("Third byte not a valid command:  %s\n" % byte3)
            return 0
        # Calculate checksum and validate it
        checksum = self.calculateChecksum(cmd)
        if checksum != ord(cmd[-1]):
            print("Incorrect checksum" + nl)
            return 0
        return 1

    def calculateChecksum(self, cmd):
        '''Return the sum of the bytes in cmd modulo 256.
        '''
        assert((len(cmd) == self.length_packet - 1) or
               (len(cmd) == self.length_packet))
        checksum = 0
        for i in range(self.length_packet - 1):
            checksum += ord(cmd[i])
        checksum %= 256
        return checksum

    def startCommand(self, byte):
        return chr(0xaa) + chr(self.address) + chr(byte)

    def sendCommand(self, command1):
        '''Sends the command to the serial stream and returns the 26 byte
        response.
        '''
        command = bytes(command1, 'latin-1')
        assert(len(command) == self.length_packet)
        self.sp.write(command)
        response = self.sp.read(self.length_packet)
        assert(len(response) == self.length_packet)
        return str(response, 'latin-1')

    def responseStatus(self, response):
        '''Return a message string about what the response meant.  The
        empty string means the response was OK.
        '''
        responses = {
            0x90: "Wrong checksum",
            0xA0: "Incorrect parameter value",
            0xB0: "Command cannot be carried out",
            0xC0: "Invalid command",
            0x80: "",
        }
        assert(len(response) == self.length_packet)
        assert(ord(response[2]) == 0x12)
        return responses[ord(response[3])]

    def codeInteger(self, value, num_bytes=4):
        '''Construct a little endian string for the indicated value.  Two
        and 4 byte integers are the only ones allowed.
        '''
        assert(num_bytes == 1 or num_bytes == 2 or num_bytes == 4)
        value = int(value)  # Make sure it's an integer
        s = chr(value & 0xff)
        if num_bytes >= 2:
            s += chr((value & (0xff << 8)) >> 8)
            if num_bytes == 4:
                s += chr((value & (0xff << 16)) >> 16)
                s += chr((value & (0xff << 24)) >> 24)
                assert(len(s) == 4)
        return s

    def decodeInteger(self, strng):
        '''Construct an integer from the little endian string. 1, 2, and 4 byte
        strings are the only ones allowed.
        '''
        assert(len(strng) == 1 or len(strng) == 2 or len(strng) == 4)
        n = ord(strng[0])
        if len(strng) >= 2:
            n += (ord(strng[1]) << 8)
            if len(strng) == 4:
                n += (ord(strng[2]) << 16)
                n += (ord(strng[3]) << 24)
        return n

    def getReserved(self, num_used):
        '''Construct a string of nul characters of such length to pad a
        command to one less than the packet size (leaves room for the
        checksum byte.
        '''
        num = self.length_packet - num_used - 1
        assert(num > 0)
        return chr(0)*num

    def printCommandAndResponse(self, cmd, response, cmd_name):
        '''Print the command and its response if debugging is on.
        '''
        assert(cmd_name)
        if self.debug:
            print(cmd_name + " command:" + nl)
            self.dumpCommand(cmd)
            print(cmd_name + " response:" + nl)
            self.dumpCommand(response)

    def getCommand(self, command, value, num_bytes=4):
        '''Construct the command with an integer value of 0, 1, 2, or
        4 bytes.
        '''
        cmd = self.startCommand(command)
        if num_bytes > 0:
            r = num_bytes + 3
            cmd += self.codeInteger(value)[:num_bytes] + self.reserved(r)
        else:
            cmd += self.reserved(0)
        cmd += chr(self.calculateChecksum(cmd))
        assert(self.commandProperlyFormed(cmd))
        return cmd

    def getData(self, data, num_bytes=4):
        '''Extract the little endian integer from the data and return it.
        '''
        assert(len(data) == self.length_packet)
        if num_bytes == 1:
            return ord(data[3])
        elif num_bytes == 2:
            return self.decodeInteger(data[3:5])
        elif num_bytes == 4:
            return self.decodeInteger(data[3:7])
        else:
            raise Exception("Bad number of bytes:  %d" % num_bytes)

    def reserved(self, num_used):
        assert(num_used >= 3 and num_used < self.length_packet - 1)
        return chr(0)*(self.length_packet - num_used - 1)

    def SendIntegerToLoad(self, byte, value, msg, num_bytes=4):
        '''Send the indicated command along with value encoded as an integer
        of the specified size.  Return the instrument's response status.
        '''
        cmd = self.getCommand(byte, value, num_bytes)
        response = self.sendCommand(cmd)
        self.printCommandAndResponse(cmd, response, msg)
        return self.responseStatus(response)

    def getIntegerFromLoad(self, cmd_byte, msg, num_bytes=4):
        '''Construct a command from the byte in cmd_byte, send it, get
        the response, then decode the response into an integer with the
        number of bytes in num_bytes.  msg is the debugging string for
        the printout.  Return the integer.
        '''
        assert(num_bytes == 1 or num_bytes == 2 or num_bytes == 4)
        cmd = self.startCommand(cmd_byte)
        cmd += self.reserved(3)
        cmd += chr(self.calculateChecksum(cmd))
        assert(self.commandProperlyFormed(cmd))
        response = self.sendCommand(cmd)
        self.printCommandAndResponse(cmd, response, msg)
        return self.decodeInteger(response[3:3 + num_bytes])


class DCLoad(InstrumentInterface):
    _reg_clsid_ = "{943E2FA3-4ECE-448A-93AF-9ECAEB49CA1B}"
    _reg_desc_ = "B&K DC Load COM Server"
    _reg_progid_ = "BKServers.DCLoad85xx"  # External name
    _public_attrs_ = ["debug"]
    _public_methods_ = [
        "disableLocalControl",
        "enableLocalControl",
        "getBatteryTestVoltage",
        "getCCCurrent",
        "getCRResistance",
        "getCVVoltage",
        "getCWPower",
        "getFunction",
        "getInputValues",
        "getLoadOnTimer",
        "getLoadOnTimerState",
        "getMaxCurrent",
        "getMaxPower",
        "getMaxVoltage",
        "getMode",
        "getProductInformation",
        "getRemoteSense",
        "getTransient",
        "getTriggerSource",
        "initialize",
        "recallSettings",
        "saveSettings",
        "setBatteryTestVoltage",
        "setCCCurrent",
        "setCRResistance",
        "setCVVoltage",
        "setCWPower",
        "setCommunicationAddress",
        "setFunction",
        "setLoadOnTimer",
        "setLoadOnTimerState",
        "setLocalControl",
        "setMaxCurrent",
        "setMaxPower",
        "setMaxVoltage",
        "setMode",
        "setRemoteControl",
        "setRemoteSense",
        "setTransient",
        "setTriggerSource",
        "timeNow",
        "triggerLoad",
        "turnLoadOff",
        "turnLoadOn",
    ]

    def initialize(self, com_port, baudrate, address=0):
        '''initialize the base class'''
        InstrumentInterface.initialize(self, com_port, baudrate, address)

    def close(self):
        ''' Close the serial port '''
        InstrumentInterface.close(self)

    def timeNow(self):
        '''Returns a string containing the current time'''
        return time.asctime()

    def turnLoadOn(self):
        '''Turns the load on'''
        msg = "Turn load on"
        on = 1
        return self.SendIntegerToLoad(0x21, on, msg, num_bytes=1)

    def turnLoadOff(self):
        '''Turns the load off'''
        msg = "Turn load off"
        off = 0
        return self.SendIntegerToLoad(0x21, off, msg, num_bytes=1)

    def setRemoteControl(self):
        '''Sets the load to remote control'''
        msg = 'Set remote control'
        remote = 1
        return self.SendIntegerToLoad(0x20, remote, msg, num_bytes=1)

    def setLocalControl(self):
        '''Sets the load to local control'''
        msg = "Set local control"
        local = 0
        return self.SendIntegerToLoad(0x20, local, msg, num_bytes=1)

    def setMaxCurrent(self, current):
        '''Sets the maximum current the load will sink'''
        msg = "Set max current"
        return self.SendIntegerToLoad(0x24, current*self.convert_current,
                                      msg, num_bytes=4)

    def getMaxCurrent(self):
        '''Returns the maximum current the load will sink'''
        msg = "Set max current"
        return (self.getIntegerFromLoad(0x25, msg, num_bytes=4) /
                self.convert_current)

    def setMaxVoltage(self, voltage):
        '''Sets the maximum voltage the load will allow'''
        msg = "Set max voltage"
        return self.SendIntegerToLoad(0x22, voltage*self.convert_voltage,
                                      msg, num_bytes=4)

    def getMaxVoltage(self):
        '''Gets the maximum voltage the load will allow'''
        msg = "Get max voltage"
        return (self.getIntegerFromLoad(0x23, msg, num_bytes=4) /
                self.convert_voltage)

    def setMaxPower(self, power):
        '''Sets the maximum power the load will allow'''
        msg = "Set max power"
        return self.SendIntegerToLoad(0x26, power*self.convert_power,
                                      msg, num_bytes=4)

    def getMaxPower(self):
        '''Gets the maximum power the load will allow'''
        msg = "Get max power"
        return (self.getIntegerFromLoad(0x27, msg, num_bytes=4) /
                self.convert_power)

    def setMode(self, mode):
        '''Sets the mode (constant current, constant voltage, etc.'''
        if mode.lower() not in self.modes:
            raise Exception("Unknown mode")
        msg = "Set mode"
        return self.SendIntegerToLoad(0x28, self.modes[mode.lower()],
                                      msg, num_bytes=1)

    def getMode(self):
        '''Gets the mode (constant current, constant voltage, etc.'''
        msg = "Get mode"
        mode = self.getIntegerFromLoad(0x29, msg, num_bytes=1)
        modes_inv = {0: "cc", 1: "cv", 2: "cw", 3: "cr"}
        return modes_inv[mode]

    def setCCCurrent(self, current):
        '''Sets the constant current mode's current level'''
        msg = "Set CC current"
        return self.SendIntegerToLoad(0x2A, current*self.convert_current,
                                      msg, num_bytes=4)

    def getCCCurrent(self):
        '''Gets the constant current mode's current level'''
        msg = "Get CC current"
        return (self.getIntegerFromLoad(0x2B, msg, num_bytes=4) /
                self.convert_current)

    def setCVVoltage(self, voltage):
        '''Sets the constant voltage mode's voltage level'''
        msg = "Set CV voltage"
        return self.SendIntegerToLoad(0x2C, voltage*self.convert_voltage,
                                      msg, num_bytes=4)

    def getCVVoltage(self):
        '''Gets the constant voltage mode's voltage level'''
        msg = "Get CV voltage"
        return (self.getIntegerFromLoad(0x2D, msg, num_bytes=4) /
                self.convert_voltage)

    def setCWPower(self, power):
        '''Sets the constant power mode's power level'''
        msg = "Set CW power"
        return self.SendIntegerToLoad(0x2E, power*self.convert_power,
                                      msg, num_bytes=4)

    def getCWPower(self):
        '''Gets the constant power mode's power level'''
        msg = "Get CW power"
        return (self.getIntegerFromLoad(0x2F, msg, num_bytes=4) /
                self.convert_power)

    def setCRResistance(self, resistance):
        '''Sets the constant resistance mode's resistance level'''
        msg = "Set CR resistance"
        return self.SendIntegerToLoad(0x30, resistance*self.convert_resistance,
                                      msg, num_bytes=4)

    def getCRResistance(self):
        '''Gets the constant resistance modes resistance level'''
        msg = "Get CR resistance"
        return (self.getIntegerFromLoad(0x31, msg, num_bytes=4) /
                self.convert_resistance)

    def setTransient(self, mode, A, A_time_s, B, B_time_s,
                     operation="continuous"):
        '''Sets up the transient operation mode.  mode is one of
           CC", "CV", "CW", or "CR".
        '''
        if mode.lower() not in self.modes:
            raise Exception("Unknown mode")
        opcodes = {"cc": 0x32, "cv": 0x34, "cw": 0x36, "cr": 0x38}
        if mode.lower() == "cc":
            const = self.convert_current
        elif mode.lower() == "cv":
            const = self.convert_voltage
        elif mode.lower() == "cw":
            const = self.convert_power
        else:
            const = self.convert_resistance
        cmd = self.startCommand(opcodes[mode.lower()])
        cmd += self.codeInteger(A*const, num_bytes=4)
        cmd += self.codeInteger(A_time_s*self.to_ms, num_bytes=2)
        cmd += self.codeInteger(B*const, num_bytes=4)
        cmd += self.codeInteger(B_time_s*self.to_ms, num_bytes=2)
        transient_operations = {"continuous": 0, "pulse": 1, "toggled": 2}
        cmd += self.codeInteger(transient_operations[operation], num_bytes=1)
        cmd += self.reserved(16)
        cmd += chr(self.calculateChecksum(cmd))
        assert(self.commandProperlyFormed(cmd))
        response = self.sendCommand(cmd)
        self.printCommandAndResponse(cmd, response, "Set %s transient" % mode)
        return self.responseStatus(response)

    def getTransient(self, mode):
        '''Gets the transient mode settings'''
        if mode.lower() not in self.modes:
            raise Exception("Unknown mode")
        opcodes = {"cc": 0x33, "cv": 0x35, "cw": 0x37, "cr": 0x39}
        cmd = self.startCommand(opcodes[mode.lower()])
        cmd += self.reserved(3)
        cmd += chr(self.calculateChecksum(cmd))
        assert(self.commandProperlyFormed(cmd))
        response = self.sendCommand(cmd)
        self.printCommandAndResponse(cmd, response, "Get %s transient" % mode)
        A = self.decodeInteger(response[3:7])
        A_timer_ms = self.decodeInteger(response[7:9])
        B = self.decodeInteger(response[9:13])
        B_timer_ms = self.decodeInteger(response[13:15])
        operation = self.decodeInteger(response[15])
        time_const = 1e3
        transient_operations_inv = {0: "continuous", 1: "pulse", 2: "toggled"}
        if mode.lower() == "cc":
            return str((A / self.convert_current, A_timer_ms / time_const,
                        B / self.convert_current, B_timer_ms / time_const,
                        transient_operations_inv[operation]))
        elif mode.lower() == "cv":
            return str((A / self.convert_voltage, A_timer_ms / time_const,
                        B / self.convert_voltage, B_timer_ms / time_const,
                        transient_operations_inv[operation]))
        elif mode.lower() == "cw":
            return str((A / self.convert_power, A_timer_ms / time_const,
                        B / self.convert_power, B_timer_ms / time_const,
                        transient_operations_inv[operation]))
        else:
            return str((A / self.convert_resistance, A_timer_ms / time_const,
                        B / self.convert_resistance, B_timer_ms / time_const,
                        transient_operations_inv[operation]))

    def setBatteryTestVoltage(self, min_voltage):
        '''Sets the battery test voltage'''
        msg = "Set battery test voltage"
        return self.SendIntegerToLoad(0x4E, min_voltage*self.convert_voltage,
                                      msg, num_bytes=4)

    def getBatteryTestVoltage(self):
        '''Gets the battery test voltage'''
        msg = "Get battery test voltage"
        return self.getIntegerFromLoad(0x4F, msg,
                                       num_bytes=4)/self.convert_voltage

    def setLoadOnTimer(self, time_in_s):
        '''Sets the time in seconds that the load will be on'''
        msg = "Set load on timer"
        return self.SendIntegerToLoad(0x50, time_in_s, msg, num_bytes=2)

    def getLoadOnTimer(self):
        '''Gets the time in seconds that the load will be on'''
        msg = "Get load on timer"
        return self.getIntegerFromLoad(0x51, msg, num_bytes=2)

    def setLoadOnTimerState(self, enabled=0):
        '''Enables or disables the load on timer state'''
        msg = "Set load on timer state"
        return self.SendIntegerToLoad(0x50, enabled, msg, num_bytes=1)

    def getLoadOnTimerState(self):
        '''Gets the load on timer state'''
        msg = "Get load on timer"
        state = self.getIntegerFromLoad(0x53, msg, num_bytes=1)
        if state == 0:
            return "disabled"
        else:
            return "enabled"

    def setCommunicationAddress(self, address=0):
        '''Sets the communication address.  Note:  this feature is
        not currently supported.  The communication address should always
        be set to 0.
        '''
        msg = "Set communication address"
        return self.SendIntegerToLoad(0x54, address, msg, num_bytes=1)

    def enableLocalControl(self):
        '''Enable local control (i.e., key presses work) of the load'''
        msg = "Enable local control"
        enabled = 1
        return self.SendIntegerToLoad(0x55, enabled, msg, num_bytes=1)

    def disableLocalControl(self):
        '''Disable local control of the load'''
        msg = "Disable local control"
        disabled = 0
        return self.SendIntegerToLoad(0x55, disabled, msg, num_bytes=1)

    def setRemoteSense(self, enabled=0):
        '''Enable or disable remote sensing'''
        msg = "Set remote sense"
        return self.SendIntegerToLoad(0x56, enabled, msg, num_bytes=1)

    def getRemoteSense(self):
        '''Get the state of remote sensing'''
        msg = "Get remote sense"
        return self.getIntegerFromLoad(0x57, msg, num_bytes=1)

    def setTriggerSource(self, source="immediate"):
        '''Set how the instrument will be triggered.
        "immediate" means triggered from the front panel.
        "external" means triggered by a TTL signal on the rear panel.
        "bus" means a software trigger (see TriggerLoad()).
        '''
        trigger = {"immediate": 0, "external": 1, "bus": 2}
        if source not in trigger:
            raise Exception("Trigger type %s not recognized" % source)
        msg = "Set trigger type"
        return self.SendIntegerToLoad(0x54, trigger[source], msg, num_bytes=1)

    def getTriggerSource(self):
        '''Get how the instrument will be triggered'''
        msg = "Get trigger source"
        t = self.getIntegerFromLoad(0x59, msg, num_bytes=1)
        trigger_inv = {0: "immediate", 1: "external", 2: "bus"}
        return trigger_inv[t]

    def triggerLoad(self):
        '''Provide a software trigger.  This is only of use when the trigger
        mode is set to "bus".
        '''
        cmd = self.startCommand(0x5A)
        cmd += self.reserved(3)
        cmd += chr(self.calculateChecksum(cmd))
        assert(self.commandProperlyFormed(cmd))
        response = self.sendCommand(cmd)
        self.printCommandAndResponse(cmd, response,
                                     "Trigger load (trigger = bus)")
        return self.responseStatus(response)

    def saveSettings(self, register=0):
        '''Save instrument settings to a register'''
        assert(self.lowest_register <= register <= self.highest_register)
        msg = "Save to register %d" % register
        return self.SendIntegerToLoad(0x5B, register, msg, num_bytes=1)

    def recallSettings(self, register=0):
        '''Restore instrument settings from a register'''
        assert(self.lowest_register <= register <= self.highest_register)
        cmd = self.getCommand(0x5C, register, num_bytes=1)
        response = self.sendCommand(cmd)
        self.printCommandAndResponse(cmd, response,
                                     "Recall register %d" % register)
        return self.responseStatus(response)

    def setFunction(self, function="fixed"):
        '''Set the function (type of operation) of the load.
        function is one of "fixed", "short", "transient", or "battery".
        Note "list" is intentionally left out for now.
        '''
        msg = "Set function to %s" % function
        functions = {"fixed": 0, "short": 1, "transient": 2, "battery": 4}
        return self.SendIntegerToLoad(0x5D, functions[function],
                                      msg, num_bytes=1)

    def getFunction(self):
        '''Get the function (type of operation) of the load'''
        msg = "Get function"
        fn = self.getIntegerFromLoad(0x5E, msg, num_bytes=1)
        functions_inv = {0: "fixed", 1: "short", 2: "transient", 4: "battery"}
        return functions_inv[fn]

    def getInputValues(self):
        '''Returns voltage in V, current in A, and power in W, op_state byte,
        and demand_state byte.
        '''
        cmd = self.startCommand(0x5F)
        cmd += self.reserved(3)
        cmd += chr(self.calculateChecksum(cmd))
        assert(self.commandProperlyFormed(cmd))
        response = self.sendCommand(cmd)
        self.printCommandAndResponse(cmd, response, "Get input values")
        voltage = self.decodeInteger(response[3:7])/self.convert_voltage
        current = self.decodeInteger(response[7:11])/self.convert_current
        power = self.decodeInteger(response[11:15])/self.convert_power
        op_state = hex(self.decodeInteger(response[15]))
        demand_state = hex(self.decodeInteger(response[16:18]))
        s = [str(voltage) + " V", str(current) + " A",
             str(power) + " W", str(op_state), str(demand_state)]
        return s

    def getProductInformation(self):
        '''Returns model number, serial number, and firmware version'''
        cmd = self.startCommand(0x6A)
        cmd += self.reserved(3)
        cmd += chr(self.calculateChecksum(cmd))
        assert(self.commandProperlyFormed(cmd))
        response = self.sendCommand(cmd)
        self.printCommandAndResponse(cmd, response, "Get product info")
        model = response[3:8]
        fw = hex(ord(response[9]))[2:] + "."
        fw += hex(ord(response[8]))[2:]
        serial_number = response[10:20]
        return ('ITECH,' + str(model) + ',' +
                str(serial_number) + ',Ver.' + str(fw))


def register(pyclass=DCLoad):
    from win32com.server.register import UseCommandLine
    UseCommandLine(pyclass)


def unregister(classid=DCLoad._reg_clsid_):
    from win32com.server.register import UnregisterServer
    UnregisterServer(classid)

# Run this script to register the COM server.  Use the command line
# argument --unregister to unregister the server.

if __name__ == '__main__':
    register()
