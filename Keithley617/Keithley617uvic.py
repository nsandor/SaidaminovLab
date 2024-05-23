import serial
import re
import time
import csv
import numpy as np
import matplotlib.pyplot
from IPython.display import display, clear_output

def connect_keithley_617(com_port):
    try:
        keithley = serial.Serial(port=com_port, baudrate=9600, timeout=1)
        print(f"Connected to Keithley 617 on {com_port}")
        return keithley
    except serial.SerialException as e:
        print(f"Error: {e}")
        return None
    
def set_voltage(keithley, voltage):
    try:
        # Set voltage command
        voltage_input = f'V{voltage}'
        b_voltage = bytes(voltage_input.encode())
        keithley.write(b_voltage)
        keithley.write(b'X\n')
    except serial.SerialException as e:
        print(f"Error setting voltage: {e}")
        
def turn_on_source_output(keithley):
    try:
        # Turn on source output command
        keithley.write(b'O1')
        keithley.write(b'X\n')
    except serial.SerialException as e:
        print(f"Error turning on source output: {e}")
        
def turn_off_source_output(keithley):
    try:
        # Turn off source output command
        keithley.write(b'O0')
        keithley.write(b'X\n')
    except serial.SerialException as e:
        print(f"Error turning off source output: {e}")
        
def measure_current(keithley):
    try:
        # Execute command to read current
        keithley.write(b'F1')
        keithley.write(b'X\n')
        keithley.write(b'G1')
        keithley.write(b'X\n')
        response = keithley.readline().decode().strip()
        return float(response)
    except (serial.SerialException, ValueError) as e:
        print(f"Error measuring current: {e}")
        return None