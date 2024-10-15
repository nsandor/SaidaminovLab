"""
Library to access the basic functionality of the Keithley 617 Programmable Electrometer by using 
Prologix (GRIP-USB Controller) for communication.
last modified: 2024-06-03
author: Dr. Sergey Dayneko
"""

import serial

class Keithley617:
    def __init__(self, com_port):
        """
        Initialize the connection to the Keithley 617.
        
        Args:
            com_port (str): The COM port to which the Keithley 617 is connected.
        """
        try:
            self.keithley = serial.Serial(port=com_port, baudrate=9600, timeout=1)
            print(f"Connected to Keithley 617 on {com_port}")
        except serial.SerialException as e:
            print(f"Error: {e}")
            self.keithley = None

    def send_command(self, command):
        """
        Send a command to the Keithley 617.

        Args:
            command (str): The command to be sent to the device.
        """
        if self.keithley:
            try:
                self.keithley.write(command.encode())
                self.keithley.write(b'X\n')
            except serial.SerialException as e:
                print(f"Error sending command: {e}")
    
    def set_voltage(self, voltage):
        """
        Set the voltage output of the Keithley 617.

        Args:
            voltage (float): The voltage value to be set.
        """
        self.send_command(f'V{voltage}')
    
    def set_function(self, function):
        """
        Set the measurement function of the Keithley 617.

        Args:
            function (str): The measurement function ('volts', 'amps', 'ohms', 'coulombs', 'external_feedback', 'vi_ohms').
        """
        functions = {'volts': 'F0', 'amps': 'F1', 'ohms': 'F2', 'coulombs': 'F3', 'external_feedback': 'F4', 'vi_ohms': 'F5'}
        if function in functions:
            self.send_command(functions[function])

    def set_range(self, range_code):
        """
        Set the measurement range of the Keithley 617.

        Args:
            range_code (str): The range code as per the device documentation.
        """
        self.send_command(range_code)

    def zero_check(self, state):
        """
        Set the zero check state.

        Args:
            state (str): 'off' to disable zero check, 'on' to enable zero check.
        """
        commands = {'off': 'C0', 'on': 'C1'}
        if state in commands:
            self.send_command(commands[state])

    def zero_correct(self, state):
        """
        Set the zero correct state.

        Args:
            state (str): 'disabled' to disable zero correct, 'enabled' to enable zero correct.
        """
        commands = {'disabled': 'Z0', 'enabled': 'Z1'}
        if state in commands:
            self.send_command(commands[state])

    def baseline_suppression(self, state):
        """
        Set the baseline suppression state.

        Args:
            state (str): 'disabled' to disable suppression, 'enabled' to enable suppression.
        """
        commands = {'disabled': 'N0', 'enabled': 'N1'}
        if state in commands:
            self.send_command(commands[state])

    def display_mode(self, mode):
        """
        Set the display mode.

        Args:
            mode (str): 'electrometer' to set electrometer mode, 'voltage_source' to set voltage source mode.
        """
        commands = {'electrometer': 'D0', 'voltage_source': 'D1'}
        if mode in commands:
            self.send_command(commands[mode])

    def reading_mode(self, mode):
        """
        Set the reading mode.

        Args:
            mode (str): 'electrometer', 'buffer_reading', 'maximum_reading', 'minimum_reading', 'voltage_source'.
        """
        commands = {'electrometer': 'B0', 'buffer_reading': 'B1', 'maximum_reading': 'B2', 'minimum_reading': 'B3', 'voltage_source': 'B4'}
        if mode in commands:
            self.send_command(commands[mode])

    def data_store(self, rate):
        """
        Set the data store conversion rate.

        Args:
            rate (str): 'conversion_rate', 'one_per_second', 'one_per_ten_seconds', 'one_per_minute', 'one_per_ten_minutes', 'one_per_hour', 'trigger_mode', 'disabled'.
        """
        rates = {'conversion_rate': 'Q0', 'one_per_second': 'Q1', 'one_per_ten_seconds': 'Q2', 'one_per_minute': 'Q3', 'one_per_ten_minutes': 'Q4', 'one_per_hour': 'Q5', 'trigger_mode': 'Q6', 'disabled': 'Q7'}
        if rate in rates:
            self.send_command(rates[rate])

    def set_voltage_value(self, value):
        """
        Set the voltage source value.

        Args:
            value (float): The voltage value to be set.
        """
        self.send_command(f'V{value}')

    def source_output(self, state):
        """
        Turn the source output on or off.

        Args:
            state (str): 'off' to turn off the source output, 'on' to turn on the source output.
        """
        commands = {'off': 'O0', 'on': 'O1'}
        if state in commands:
            self.send_command(commands[state])

    def calibrate(self, value):
        """
        Set the calibration value.

        Args:
            value (float): The calibration value to be set.
        """
        self.send_command(f'A{value}')

    def store_calibration(self):
        """
        Store the calibration constants in NVRAM.
        """
        self.send_command('L1')

    def data_format(self, format_type):
        """
        Set the data format.

        Args:
            format_type (str): 'with_prefix', 'without_prefix', 'prefix_suffix'.
        """
        formats = {'with_prefix': 'G0', 'without_prefix': 'G1', 'prefix_suffix': 'G2'}
        if format_type in formats:
            self.send_command(formats[format_type])

    def trigger_mode(self, mode):
        """
        Set the trigger mode.

        Args:
            mode (str): 'continuous_talk', 'one_shot_talk', 'continuous_get', 'one_shot_get', 'continuous_x', 'one_shot_x', 'continuous_external', 'one_shot_external'.
        """
        commands = {'continuous_talk': 'T0', 'one_shot_talk': 'T1', 'continuous_get': 'T2', 'one_shot_get': 'T3', 'continuous_x': 'T4', 'one_shot_x': 'T5', 'continuous_external': 'T6', 'one_shot_external': 'T7'}
        if mode in commands:
            self.send_command(commands[mode])

    def srq(self, condition):
        """
        Set the SRQ condition.

        Args:
            condition (str): 'disable', 'overflow', 'buffer_full', 'done', 'ready', 'error'.
        """
        conditions = {'disable': 'M0', 'overflow': 'M1', 'buffer_full': 'M2', 'done': 'M8', 'ready': 'M16', 'error': 'M32'}
        if condition in conditions:
            self.send_command(conditions[condition])

    def eoi_bus_hold(self, state):
        """
        Set the EOI and bus hold-off state.

        Args:
            state (str): 'enable', 'disable_enable', 'enable_disable', 'disable'.
        """
        states = {'enable': 'K0', 'disable_enable': 'K1', 'enable_disable': 'K2', 'disable': 'K3'}
        if state in states:
            self.send_command(states[state])

    def terminator(self, terminator_type):
        """
        Set the terminator type.

        Args:
            terminator_type (str): 'lf_cr', 'cr_lf', 'ascii', 'none'.
        """
        terminators = {'lf_cr': 'Y(LF CR)', 'cr_lf': 'Y(CR LF)', 'ascii': 'Y(ASCII)', 'none': 'YX'}
        if terminator_type in terminators:
            self.send_command(terminators[terminator_type])

    def status_word(self, word_type):
        """
        Set the status word type.

        Args:
            word_type (str): 'send', 'error', 'data'.
        """
        words = {'send': 'U0', 'error': 'U1', 'data': 'U2'}
        if word_type in words:
            self.send_command(words[word_type])

    def measure_current(self):
        """
        Measure the current and return the value.

        Returns:
            float: The measured current value.
        """
        response = self.keithley.readline().decode().strip()
        return float(response) if response else None

    def disconnect(self):
        """
        Disconnect from the Keithley 617.
        """
        if self.keithley:
            try:
                self.keithley.close()
                print("Disconnected from Keithley 617")
            except serial.SerialException as e:
                print(f"Error disconnecting: {e}")
