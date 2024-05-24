import datetime
import os
import sys
import numpy as np
import pyvisa as visa
import time
import csv
import matplotlib.pyplot as plt
from IPython.display import display, clear_output
from pymeasure.instruments import list_resources
import glob
import pandas as pd

""" 
Reference for SCPI commands of Keithley 2450
https://download.tek.com/manual/2450-901-01_D_May_2015_Ref.pdf
"""

### -------------------- Utilities -------------------- ###

## To check the Keithley address, use this code
def get_Keithley_address():
    list_resources()

def make_folder(name):
    # # Check if the project directory exists
    # if os.path.exists(name):
    #     print(f'\033[33mFolder "{name}" already exists.')

    # Create the project directory if it does not exist
    if not os.path.exists(name):
        try:
            os.mkdir(name)
            print(f'\033[33mFolder named "{name}" has been prepared.')
        except OSError:
            print(f"Error: Could not create directory {name}")
            sys.exit(0)
            
def int_ask(question):
    while True:
        A = input(question)
        if A == '0' or A == '1':
            break
        else:
            print('Please input 0 or 1')
            pass
    A = int(A)
    return A

def current_set(currents):
    '''
    input current list
    convert it to a list depending on the order for a realtime plot
    pA, nA, µA, mA, A
    '''
    # Get currents data
    MAX, MIN = abs(max(currents)), abs(min(currents))
    if MAX >= MIN:
        M = MAX
    elif MIN > MAX:
        M = MIN
    
    # Decide the range
    if M < 1e-9:
        RANGE = 'pA'
        currents = [n*1e12 for n in currents]
    elif 1e-9 <= M < 1e-6:
        RANGE = 'nA'
        currents = [n*1e9 for n in currents]
    elif 1e-6 <= M < 1e-3:
        RANGE = 'µA'
        currents = [n*1e6 for n in currents]
    elif 1e-3 <= M < 1:
        RANGE = 'mA'
        currents = [n*1e3 for n in currents]
    else:
        pass
    
    # LABEL NAME
    LABEL = 'Current (' + RANGE + ')'
    return LABEL, currents

def set_IRANGE(RANGE, series):
    """
    input current dataset as DataFrame [pandas.core.series.Series] and range
    convert it to a list depending on the order-pA, nA, µA, mA, A
    """    
    currents = series.to_list()
    # Decide the range
    if RANGE == 'pA':
        currents = [n*1e12 for n in currents]
    elif RANGE == 'nA':
        currents = [n*1e9 for n in currents]
    elif RANGE == 'uA':
        RANGE = 'µA'
        currents = [n*1e6 for n in currents]
    elif RANGE == 'mA':
        currents = [n*1e3 for n in currents]
    else:
        pass
    
    # LABEL NAME
    LABEL = 'Current (' + RANGE + ')'
    return LABEL, currents



### -------------------- Main things -------------------- ###

def project_start(save_dir, operation_mode):
    # Set the project name
    user_list = os.listdir(save_dir)
    for i, user in enumerate(user_list):
        print(f'{i+1:02}: {user}')
    user_ID = int(input('Select your user ID'))
    user_name = user_list[user_ID-1]
    
    # Get the date
    today = datetime.datetime.today().strftime('%y%m%d')

    # Set the project name if you need
    print('\033[31mCreate a new folder (0) or Use an existing folder (1)\033[33m')
    mkdir_option = int_ask('')

    if mkdir_option == 0:
        print('\033[31mEnter project name: \033[33m')
        project_name = input('')
        while project_name == '':
            print('\033[31mError: Project name cannot be empty\nEnter project name: \033[33m')
            project_name = input('')
        # Define the name
        mode_dir = f"{save_dir}/{user_name}/{today}_{project_name}"
        project_dir = f"{save_dir}/{user_name}/{today}_{project_name}/{operation_mode}"
        fig_dir = f"{project_dir}/figures"

    elif mkdir_option == 1:
        folder_list = os.listdir(f"{save_dir}/{user_name}")
        for i, folder in enumerate(folder_list):
            print(f'\033[33m{i+1:02}: {folder}\033[33m')
        print('\033[31mSelect your folder ID\033[33m')
        folder_ID = int(input(''))
        mode_dir = f"{save_dir}/{user_name}/{folder_list[folder_ID-1]}"
        project_dir = f"{save_dir}/{user_name}/{folder_list[folder_ID-1]}/{operation_mode}"
        fig_dir = f"{project_dir}/figures"
    
    # Prepare the directory
    make_folder(mode_dir)
    make_folder(project_dir)
    make_folder(fig_dir)
    
    return project_dir

### PV-SCLC ###
def run_pv_sclc(project_dir,start_log,end_log,log_step,t_ON,t_INT,direction,terminals,address):
    
    # Folder to save the directory
    fig_dir = f"{project_dir}/figures"
    
    # Set the voltage source (V)
    source_log = np.arange(start_log, end_log + log_step, log_step)
    source_voltage = [10**n for n in source_log] # prepare voltage sources

    # # for reverse scan
    source_voltage_R = []
    for x in reversed(source_voltage):
        source_voltage_R.append(x)
        
    # Print the estimated time
    total_time = len(source_voltage) * (t_INT + t_ON)
    if direction == 'B':
        total_time = total_time * 2
    if total_time < 60:
        print(f'Estimated total time is {int(total_time)} sec')
    elif total_time >= 60 and total_time <= 3600:
        print(f'Estimated total time is {int(total_time)/60} min')
    elif total_time > 3600:
        print(f'Estimated total time is {int(total_time)/3600} hrs')

    finish = datetime.datetime.now() + datetime.timedelta(seconds=total_time)
    print(f'Estimated finish time is {str(finish.strftime("%Y/%m/%d, %H:%M:%S"))}')
    
    # Open a connection to the Keithley 2450
    try:
        rm = visa.ResourceManager()
        keithley = rm.open_resource(address)
    except:
        print("Error: Could not connect to instrument")
        sys.exit(0)

    # Make sure if you start or not
    START = input('Press Enter to Start')
    if START == '':
        print("\n Let's get started :)")
        pass
    else:
        sys.exit(0)

    # Initialize the time and current arrays
    times = []
    t_ON_list = []
    t_OFF_list = []
    currents = []
    voltages = []

    num_steps = len(source_voltage)
    if direction == 'B':
        num_steps = num_steps * 2

    keithley.write('*RST') # Rest the Keithley
    keithley.write(":SOUR:VOLT:ILIMIT 0.05") # Set the current limit to 50 mA
    if terminals == 'REAR':
        keithley.write(":ROUT:TERM REAR") # Set to use rear-panel terminals
    keithley.write('TRACE:MAKE "VMEAS", 11; :TRACE:MAKE "CMEAS", 11')
    keithley.write("COUN 1")

    if direction == 'F' or direction == 'B':
        # Start the Forward scan
        start_time = time.perf_counter()
        for i, voltage in enumerate(source_voltage):
            clear_output(wait=True)
            # Set the voltage source
            keithley.write(f'TRACE:CLEAR "CMEAS"; :TRACE:CLEAR "VMEAS"; :SENSE:FUNCtion "CURRent"; :SOURCE:VOLTAGE:LEVEL {voltage}')
            # Set the voltage source output on
            output_start = time.perf_counter()
            keithley.write('OUTPUT ON; :TRACE:TRIG "CMEAS"; :SENSE:FUNCtion "VOLTage"; :TRACE:TRIG "VMEAS"')
            if t_ON > 0:
                while (t_ON - time.perf_counter() + output_start) > 0:
                    pass    
            # Set the voltage source output off
            keithley.write('OUTPUT OFF')
            output_end = time.perf_counter()

            current = float(keithley.query('FETCH? "CMEAS"'))
            voltage = float(keithley.query('FETCH? "VMEAS"'))
            times.append((output_end - start_time))

            t_ON_list.append(output_end - output_start)
            currents.append(current)
            voltages.append(voltage)
            
            # Realtime monitoring
            print(f'Estimated finish time is {str(finish.strftime("%Y/%m/%d, %H:%M:%S"))}')
            print(f'Step: {len(times)}/{num_steps} steps')
            print(f't-ON: {t_ON_list[-1]:.4f} s')
            print(f'Voltage: {voltage:.4g} V')
            print(f'Current: {current:.4g} A')
            
            # Set up the real-time plot
            plt.ion()
            fig = plt.figure(figsize=(12,8))
            plt.rcParams["font.size"] = 20
            plt.plot(voltages, currents, linestyle='-', marker='o', label='Current', color='blue')
            plt.xlabel('Voltage (V)')
            plt.ylabel('Current (A)')
            plt.grid(True)
            display(fig)
            
            # Add interval
            if (t_INT - (time.perf_counter() - output_end)) > 0:
                while (t_INT - (time.perf_counter() - output_end))>0:
                    pass
            else:
                pass

            t_OFF_list.append(time.perf_counter() - output_end)

    if direction == 'R' or direction == 'B':
        # Initialize the time and current arrays
        times_R = []
        t_ON_list_R = []
        t_OFF_list_R = []
        currents_R = []
        voltages_R = []

        # Start the Reverse scan
        start_time_R = time.perf_counter()
        
        for i,voltage in enumerate(source_voltage_R):
            clear_output(wait=True)
            keithley.write(f'TRACE:CLEAR "CMEAS"; :TRACE:CLEAR "VMEAS"; :SENSE:FUNCtion "CURRent"; :SOURCE:VOLTAGE:LEVEL {voltage}')
            # Set the voltage source output on
            output_start = time.perf_counter()
            keithley.write('OUTPUT ON; :TRACE:TRIG "CMEAS"; :SENSE:FUNCtion "VOLTage"; :TRACE:TRIG "VMEAS"')
            if t_ON > 0:
                while (t_ON - time.perf_counter() + output_start) > 0:
                    pass    
            # Set the voltage source output off
            keithley.write('OUTPUT OFF')
            output_end = time.perf_counter()

            current = float(keithley.query('FETCH? "CMEAS"'))
            voltage = float(keithley.query('FETCH? "VMEAS"'))
            t_ON_list_R.append(output_end - output_start)
            times_R.append((output_end - start_time_R))
            currents_R.append(float(current))
            voltages_R.append(float(voltage))
                        
            # Realtime monitoring
            print(f'Estimated finish time is {str(finish.strftime("%Y/%m/%d, %H:%M:%S"))}')
            print(f'Step: {len(times)+len(times_R)} /{num_steps} step')
            print(f't-ON: {t_ON_list_R[-1]:.4f} s')
            print(f'Voltage: {voltage:.4g} V')
            print(f'Current: {current:.4g} A')
            
            # Set up the real-time plot
            plt.ion()
            fig2 = plt.figure(figsize=(12,8))
            plt.rcParams["font.size"] = 20
            plt.plot(voltages, currents, linestyle='-', marker='o', label='FORWARD', color='blue')
            plt.plot(voltages_R, currents_R, linestyle='-', marker='o', label='REVERSE', color='green')
            plt.xlabel('Voltage (V)')
            plt.ylabel('Current (A)')
            plt.legend(frameon=False)
            plt.grid(True)
            display(fig2)
            
            # Add interval
            if (t_INT - (time.perf_counter() - output_end)) > 0:
                while (t_INT - (time.perf_counter() - output_end)) > 0:
                    pass
            else:
                pass

            t_OFF_list_R.append(time.perf_counter() - output_end)
            
    # Close the connection to the Keithley 2450
    keithley.write('TRACE:DEL "VMEAS"')
    keithley.write('TRACE:DEL "CMEAS"')
    keithley.close()
    
    # Set the filename
    file_name = input('Enter file name: ')
    if direction == 'F' or direction == 'B':
        filename = f'{project_dir}/{file_name}-F.csv'
    if direction == 'R' or direction == 'B':
        filename_R = f'{project_dir}/{file_name}-R.csv'
        
    clear_output(wait=True)
    
    # Create a CSV file for saving the data
    with open(filename, 'w', newline='') as csvfile:
        csvwriter = csv.writer(csvfile)
        csvwriter.writerow(['Time (s)', 't-ON (s)', 't-OFF (s)', 'Current (A)', 'Voltage (V)'])
        for i in range(len(times)):
            csvwriter.writerow([times[i], t_ON_list[i], t_OFF_list[i], currents[i], voltages[i]])

    # Check that the CSV file was created successfully
    try:
        with open(filename, 'r') as csvfile:
            pass
    except:
        print("Error: Could not create CSV file")
    
    if not direction == 'F':
        # Create a CSV file for saving the data
        with open(filename_R, 'w', newline='') as csvfile:
            csvwriter = csv.writer(csvfile)
            csvwriter.writerow(['Time (s)', 't-ON (s)', 't-OFF (s)', 'Current (A)', 'Voltage (V)'])
            for i in range(len(times_R)):
                csvwriter.writerow([times_R[i], t_ON_list_R[i], t_OFF_list_R[i], currents_R[i], voltages_R[i]])

        # Check that the CSV file was created successfully
        try:
            with open(filename_R, 'r') as csvfile:
                pass
        except:
            print("Error: Could not create CSV file")

    
    # Figure to save
    fig3 = plt.figure(figsize=(12,8))
    plt.rcParams["font.size"] = 20
    if not direction == 'R':
        plt.plot(voltages, currents, linestyle='-', marker='o', label='FORWARD', color='blue')
    if not direction == 'F':
        plt.plot(voltages_R, currents_R, linestyle='-', marker='o', label='REVERSE', color='green')
    plt.xlabel('Voltage (V)')
    plt.ylabel('Current (A)')
    plt.legend(frameon=False)
    plt.grid(True)
    display(fig3)
    plt.savefig(f'{fig_dir}/{file_name}.jpg', bbox_inches='tight')

    print(f'Program completed at {str(datetime.datetime.now().strftime("%Y/%m/%d, %H:%M:%S"))}')
    
### tI measurement (Constant voltage source) ###
def run_tI(project_dir,source_voltage,delay_time,duration,ILIMIT,terminals,address):
    
    # Folder to save the directory
    fig_dir = f"{project_dir}/figures"
        
    # Make sure if you start or not
    START = input('\nPress Enter to Start')
    if START == '':
        print("\n Let's get started :)")
        pass
    else:
        sys.exit(0)

    # Open a connection to the Keithley 2450
    try:
        rm = visa.ResourceManager()
        keithley = rm.open_resource(address)
    except:
        print("Error: Could not connect to instrument")
        sys.exit(0)
    
    # Setting
    keithley.write('*RST') # Rest the Keithley
    keithley.write(f":SOUR:VOLT:ILIMIT {ILIMIT}") # Set the current limit
    if terminals == 'REAR':
        keithley.write(":ROUT:TERM REAR") # Set to use rear-panel terminals
    keithley.write(f'SOURCE:VOLTAGE:LEVEL {source_voltage}') # Set the voltage source

    # Set the voltage source output on
    keithley.write('OUTPUT ON')

    try:
        # Initialize the time and current arrays
        times = []
        currents = []
        voltages = []

        # Set up the real-time plot
        plt.ion()
        fig = plt.figure(figsize=(12,8))
        plt.rcParams["font.size"] = 20

        # Start the measurement and real-time plot
        start_time = time.perf_counter()
        while (time.perf_counter() - start_time) < duration:
            round_start = time.perf_counter()
            voltage = float(keithley.query('MEASURE:VOLTAGE?'))
            current = float(keithley.query('MEASURE:CURRENT?'))
            times.append((time.perf_counter() - start_time))
            currents.append(current)
            voltages.append(voltage)
            plot_start = time.perf_counter()
            
            clear_output(wait=True)
            
            # Realtime monitor
            print(f'Time: {times[-1]:.2f} s / {duration} s')
            print(f'Voltage: {voltage:.4g} V')
            print(f'Current: {current:.4g} A')
            
            # Plot data
            fig.clf() # initialize
            ylabel, currents_plot = current_set(currents)
            plt.plot(times, currents_plot, linestyle='-', marker='o', color='blue')
            plt.xlabel('Time (s)')
            plt.ylabel(ylabel)
            plt.grid(True)
            display(fig)
            
            print(f'Round time: {time.perf_counter()-round_start} s')
            print(f'Plot time: {time.perf_counter()-plot_start} s')

            if (delay_time - (time.perf_counter() - round_start)) > 0:
                time.sleep(delay_time - (time.perf_counter() - round_start))
                print(f'Real Round time: {time.perf_counter()-round_start} s')
            else:
                pass
    except:
        print(Exception)
        keithley.write('OUTPUT OFF')
        keithley.close()
        sys.exit(0)

    # Set the voltage source output off
    keithley.write('OUTPUT OFF')

    # Close the connection to the Keithley 2450
    keithley.close()

    clear_output(wait=True)
    
    # Show the figure
    fig2 = plt.figure(figsize=(12,8))
    plt.rcParams["font.size"] = 20
    ylabel, currents_plot = current_set(currents)
    plt.plot(times, currents_plot, linestyle='-', marker='o', color='blue')
    plt.xlabel('Time (s)')
    plt.ylabel(ylabel)    
    plt.show()

    # Save the figure
    print(f'\033[33mEnter the file name [XXX], it will be "{project_dir}/XXX.csv"')
    file_name = input('')
    filename = os.path.join(project_dir, file_name + '.csv')
    fig_path = os.path.join(fig_dir, file_name + '.png')
    fig2.savefig(fig_path, transparent = True)
    
    # Create a CSV file for saving the data
    with open(filename, 'w', newline='') as csvfile:
        csvwriter = csv.writer(csvfile)
        csvwriter.writerow(['Time (s)', 'Current (A)', 'Voltage (V)'])
        for i in range(len(times)):
            csvwriter.writerow([times[i], currents[i], voltages[i]])

    # Check that the CSV file was created successfully
    try:
        with open(filename, 'r') as csvfile:
            pass
    except:
        print("Error: Could not create CSV file")
        
    print("Program completed")

    return times, currents_plot

### tI measurement (Constant voltage source) ###
def run_tI_step(project_dir,source_voltages,delay_time,duration,ILIMIT,terminals,address):
    
    # source_voltages = [1,2,5,10,20,50]
    # duration = [20*60, 20*60] # ON time, OFF time (sec)
    
    # Folder to save the directory
    fig_dir = f"{project_dir}/figures"
    
    # Confirm the duration
    total_time = len(source_voltages) * (duration[0] + duration[1]) / 60 # min
    print(f'Total time is {total_time} mins')
    finish_time = datetime.datetime.now() + datetime.timedelta(seconds=int(total_time*60))
    print(f'Estimated finish time is {finish_time.strftime("%Y-%m-%d %H:%M:%S")}')
    
    # Make sure if you start or not
    START = input('\nPress Enter to Start')
    if START == '':
        print("\n Let's get started :)")
        pass
    else:
        sys.exit(0)

    # Open a connection to the Keithley 2450
    try:
        rm = visa.ResourceManager()
        keithley = rm.open_resource(address)
    except:
        print("Error: Could not connect to instrument")
        sys.exit(0)
    
    # Setting
    keithley.write('*RST') # Rest the Keithley
    keithley.write(f":SOUR:VOLT:ILIMIT {ILIMIT}") # Set the current limit
    if terminals == 'REAR':
        keithley.write(":ROUT:TERM REAR") # Set to use rear-panel terminals
    keithley.write(f'SOURCE:VOLTAGE:LEVEL 0') # Set the voltage source

    # Set the voltage source output on
    keithley.write('OUTPUT ON')

    try:
        # Initialize the time and current arrays
        times = []
        currents = []
        voltages = []

        # Set up the real-time plot
        plt.ion()
        fig = plt.figure(figsize=(12,8))
        plt.rcParams["font.size"] = 20

        # Start the measurement and real-time plot
        start_time = time.perf_counter()
        
        for source_voltage in source_voltages:
            # Measure ON current
            keithley.write(f'SOURCE:VOLTAGE:LEVEL {source_voltage}') # Set the voltage source
            round_start_time = time.perf_counter()
            while time.perf_counter() - round_start_time < duration[0]:
                round_start = time.perf_counter()
                voltage = float(keithley.query('MEASURE:VOLTAGE?'))
                current = float(keithley.query('MEASURE:CURRENT?'))
                times.append((time.perf_counter() - start_time))
                currents.append(current)
                voltages.append(voltage)
                plot_start = time.perf_counter()

                clear_output(wait=True)

                # Realtime monitor
                print(f'Time: {times[-1]:.2f} s')
                print(f'Voltage: {voltage:.4g} V')
                print(f'Current: {current:.4g} A')

                # Plot data
                fig.clf() # initialize
                ylabel, currents_plot = current_set(currents)
                plt.plot(times, currents_plot, linestyle='-', marker='o', color='blue')
                plt.xlabel('Time (s)')
                plt.ylabel(ylabel)
                plt.grid(True)
                display(fig)

                print(f'Round time: {time.perf_counter()-round_start} s')
                print(f'Plot time: {time.perf_counter()-plot_start} s')

                if (delay_time - (time.perf_counter() - round_start)) > 0:
                    time.sleep(delay_time - (time.perf_counter() - round_start))
                    print(f'Real Round time: {time.perf_counter()-round_start} s')
                else:
                    pass
                
            # Measure OFF current
            keithley.write(f'SOURCE:VOLTAGE:LEVEL 0') # Set the voltage source
            round_start_time = time.perf_counter()
            while time.perf_counter() - round_start_time < duration[1]:
                round_start = time.perf_counter()
                voltage = float(keithley.query('MEASURE:VOLTAGE?'))
                current = float(keithley.query('MEASURE:CURRENT?'))
                times.append((time.perf_counter() - start_time))
                currents.append(current)
                voltages.append(voltage)
                plot_start = time.perf_counter()

                clear_output(wait=True)

                # Realtime monitor
                print(f'Time: {times[-1]:.2f} s')
                print(f'Voltage: {voltage:.4g} V')
                print(f'Current: {current:.4g} A')

                # Plot data
                fig.clf() # initialize
                ylabel, currents_plot = current_set(currents)
                plt.plot(times, currents_plot, linestyle='-', marker='o', color='blue')
                plt.xlabel('Time (s)')
                plt.ylabel(ylabel)
                plt.grid(True)
                display(fig)

                print(f'Round time: {time.perf_counter()-round_start} s')
                print(f'Plot time: {time.perf_counter()-plot_start} s')

                if (delay_time - (time.perf_counter() - round_start)) > 0:
                    time.sleep(delay_time - (time.perf_counter() - round_start))
                    print(f'Real Round time: {time.perf_counter()-round_start} s')
                else:
                    pass   
    except:
        print(Exception)
        keithley.write('OUTPUT OFF')
        keithley.close()
        sys.exit(0)

    # Set the voltage source output off
    keithley.write('OUTPUT OFF')

    # Close the connection to the Keithley 2450
    keithley.close()

    clear_output(wait=True)
    
    # Show the figure
    fig2 = plt.figure(figsize=(12,8))
    plt.rcParams["font.size"] = 20
    ylabel, currents_plot = current_set(currents)
    plt.plot(times, currents_plot, linestyle='-', marker='o', color='blue')
    plt.xlabel('Time (s)')
    plt.ylabel(ylabel)    
    plt.show()

    # Save the figure
    print(f'\033[33mEnter the file name [XXX], it will be "{project_dir}/XXX.csv"')
    file_name = input('')
    filename = os.path.join(project_dir, file_name + '.csv')
    fig_path = os.path.join(fig_dir, file_name + '.png')
    fig2.savefig(fig_path, transparent = True)
    
    # Create a CSV file for saving the data
    with open(filename, 'w', newline='') as csvfile:
        csvwriter = csv.writer(csvfile)
        csvwriter.writerow(['Time (s)', 'Current (A)', 'Voltage (V)'])
        for i in range(len(times)):
            csvwriter.writerow([times[i], currents[i], voltages[i]])

    # Check that the CSV file was created successfully
    try:
        with open(filename, 'r') as csvfile:
            pass
    except:
        print("Error: Could not create CSV file")
        
    print("Program completed")

    
### IV measurement (Linear Voltage Sweep) ###

def run_IV(project_dir,v_range,step_size,scan_rate,ILIMIT,direction,terminals,address):
    
    start_voltage = v_range[0]
    end_voltage = v_range[1]
    
    if start_voltage > end_voltage:
        step_size = (-1)*step_size
    source_voltage = np.arange(start_voltage, end_voltage+step_size, step_size) # prepare voltage sources
    source_voltage_R = list(source_voltage)
    source_voltage_R.reverse()
    total_time = abs(end_voltage - start_voltage) / scan_rate
    delay_time =  total_time / len(source_voltage)
    if direction == 'B':
        total_time = total_time * 2

    # validation
    if max(source_voltage)*ILIMIT > 5:
        print('invalid limit! output should be less than 5 W')
        sys.exit()

    if total_time < 60:
        print(f'Estimated total time is {int(total_time)} sec')
    elif total_time >= 60 and total_time <= 3600:
        print(f'Estimated total time is {int(total_time)/60} min')
    elif total_time > 3600:
        print(f'Estimated total time is {int(total_time)/3600} hrs')
    
    # Check Parameters
    print(f'Range: from {start_voltage} V to {end_voltage} V')
    print(f'Scan rate: {scan_rate} V/s')

    # Make sure if you start or not
    START = input('Press Enter to Start')
    if START == '':
        print("\n Let's get started :)")
        pass
    else:
        sys.exit(0)

    # Open a connection to the Keithley 2450
    try:
        rm = visa.ResourceManager()
        keithley = rm.open_resource(address)
    except:
        print("Error: Could not connect to instrument")
        sys.exit(0)

    # Set up the real-time plot
    plt.ion()
    fig = plt.figure(figsize=(12,8))
    plt.rcParams["font.size"] = 20
    plt.xlabel('Voltage (V)')
    plt.ylabel('Current (A)')
    
    # Setting
    keithley.write('*RST') # Rest the Keithley
    keithley.write(f":SOUR:VOLT:ILIMIT {ILIMIT}") # Set the current limit
    if terminals == 'REAR':
        keithley.write(":ROUT:TERM REAR") # Set to use rear-panel terminals

    if not direction == 'R':
        # Initialize the time and current arrays
        times = []
        currents = []
        voltages = []
        keithley.write(f'SOURCE:VOLTAGE:LEVEL 0')
        keithley.write('OUTPUT ON')
        # Start the measurement and real-time plot
        start_time = time.perf_counter()
        for i in range(len(source_voltage)):
            round_start = time.perf_counter()
            # Set the voltage source
            keithley.write(f'SOURCE:VOLTAGE:LEVEL {source_voltage[i]}') 
            clear_output(wait=True)
            voltage = float(keithley.query('MEASURE:VOLTAGE?'))
            current = float(keithley.query('MEASURE:CURRENT?'))
            times.append((time.perf_counter() - start_time))
            currents.append(current)
            voltages.append(voltage)
            print(f'Time: {times[-1]:.2f} s')
            print(f'Voltage: {voltage:.4g} V')
            print(f'Current: {current:.4g} A')
            plt.plot(voltages, currents, linestyle='-', marker='o', label='Current', color='blue')
            plt.xlabel('Voltage (V)', fontsize=18)
            plt.ylabel('Current (A)', fontsize=18)
            plt.xticks(fontsize=16)
            plt.yticks(fontsize=16)
            plt.grid(True)
            display(fig)
            if (delay_time - (time.perf_counter() - round_start)) > 0:
                time.sleep(delay_time - (time.perf_counter() - round_start))
            else:
                pass

        clear_output(wait=True)
        # Set the voltage source output off
        keithley.write('OUTPUT OFF')

    # start reverse scan
    if not direction == 'F':
        # Initialize the time and current arrays
        times_R = []
        currents_R = []
        voltages_R = []

        # Set the voltage source output on
        keithley.write(f'SOURCE:VOLTAGE:LEVEL 0') 
        keithley.write('OUTPUT ON')

        # Start the measurement and real-time plot
        start_time = time.perf_counter()
        for i in range(len(source_voltage_R)):
            round_start = time.perf_counter()
            # Set the voltage source
            keithley.write(f'SOURCE:VOLTAGE:LEVEL {source_voltage_R[i]}') 
            clear_output(wait=True)
            voltage = float(keithley.query('MEASURE:VOLTAGE?'))
            current = float(keithley.query('MEASURE:CURRENT?'))
            times_R.append((time.perf_counter() - start_time))
            currents_R.append(current)
            voltages_R.append(voltage)
            print(f'Time: {times_R[-1]:.2f} s')
            print(f'Voltage: {voltage:.4g} V')
            print(f'Current: {current:.4g} A')
            plt.plot(voltages, currents, linestyle='-', marker='o', label='Current', color='blue')
            plt.plot(voltages_R, currents_R, linestyle='-', marker='o', label='Current', color='green')
            plt.xlabel('Voltage (V)', fontsize=18)
            plt.ylabel('Current (A)', fontsize=18)
            plt.xticks(fontsize=16)
            plt.yticks(fontsize=16)
            plt.grid(True)
            display(fig)
            if (delay_time - (time.perf_counter() - round_start)) > 0:
                time.sleep(delay_time - (time.perf_counter() - round_start))
            else:
                pass

        clear_output(wait=True)
        # Set the voltage source output off
        keithley.write('OUTPUT OFF')

    # Close the connection to the Keithley 2450
    keithley.close()
    
    # Set the filename
    file_name = input('Enter file name: ')
    filename = f'{project_dir}/{file_name}-F.csv'
    filename_R = f'{project_dir}/{file_name}-R.csv'
    
    if not direction == 'R':
        # Create a CSV file for saving the data
        with open(filename, 'w', newline='') as csvfile:
            csvwriter = csv.writer(csvfile)
            csvwriter.writerow(['Time (s)', 'Current (A)', 'Voltage (V)'])
            for i in range(len(times)):
                csvwriter.writerow([times[i], currents[i], voltages[i]])

        # Check that the CSV file was created successfully
        try:
            with open(filename, 'r') as csvfile:
                pass
        except:
            print("Error: Could not create CSV file")
    
    if not direction == 'F':
        # Create a CSV file for saving the data
        with open(filename_R, 'w', newline='') as csvfile:
            csvwriter = csv.writer(csvfile)
            csvwriter.writerow(['Time (s)', 'Current (A)', 'Voltage (V)'])
            for i in range(len(times)):
                csvwriter.writerow([times_R[i], currents_R[i], voltages_R[i]])

        # Check that the CSV file was created successfully
        try:
            with open(filename_R, 'r') as csvfile:
                pass
        except:

            print("Error: Could not create CSV file")
    
    # Show the figure
    fig2 = plt.figure(figsize=(12,8))
    plt.rcParams["font.size"] = 20
    if direction == 'B':
        ylabel, currents_plot = current_set(currents+currents_R)
        plt.plot(voltages, currents_plot[:len(currents)], linestyle='-', marker='o', color='blue', label='Forward')
        plt.plot(voltages_R, currents_plot[len(currents):], linestyle='-', marker='o', color='green', label='Reverse')
    if direction == 'F':
        ylabel, currents_plot = current_set(currents)
        plt.plot(voltages, currents_plot, linestyle='-', marker='o', color='blue', label='Forward')
    if direction == 'R':
        ylabel, currents_plot = current_set(currents_R)
        plt.plot(voltages_R, currents_plot, linestyle='-', marker='o', color='green', label='Reverse')
    plt.xlabel('Voltage (V)')
    plt.ylabel(ylabel)
    plt.legend(frameon=False)
    
    # Save the figure
    fig_dir = f"{project_dir}/figures"
    fig_path = os.path.join(fig_dir, file_name + '.png')
    fig2.savefig(fig_path, transparent = True)
    
    print("Program completed")

    
def run_tI_pulse(project_dir, source_voltage,delay_time,t_ON,duration,ILIMIT,terminals,address):

    # Make sure if you start or not
    START = input('\nPress Enter to Start')
    if START == '':
        print("\n Let's get started :)")
        pass
    else:
        sys.exit(0)

    # Open a connection to the Keithley 2450
    try:
        rm = visa.ResourceManager()
        keithley = rm.open_resource(address)
    except:
        print("Error: Could not connect to instrument")
        sys.exit(0)

    # Setting
    keithley.write('*RST') # Reset the Keithley
    keithley.write(f":SOUR:VOLT:ILIMIT {ILIMIT}") # Set the current limit
    
    if terminals == 'REAR':
        keithley.write(":ROUT:TERM REAR") # Set to use rear-panel terminals
    keithley.write(f'SOURCE:VOLTAGE:LEVEL {source_voltage}') # Set the voltage source

    # Initialize the time and current arrays
    times = []
    t_ON_list = []
    currents = []
    voltages = []

    # Set up the real-time plot
    plt.ion()
    fig = plt.figure(figsize=(12,8))
    plt.rcParams["font.size"] = 20
    plt.xlabel('Time (s)')
    plt.ylabel('Current (A)')

    keithley.write('TRACE:MAKE "VMEAS", 11; :TRACE:MAKE "CMEAS", 11')
    keithley.write("COUN 1")
    # Start the measurement and real-time plot
    start_time = time.perf_counter()
    while (time.perf_counter() - start_time) < duration:
        round_start = time.perf_counter()
        keithley.write('TRACE:CLEAR "CMEAS"; :TRACE:CLEAR "VMEAS"; :SENSE:FUNCtion "CURRent"')
        # Set the voltage source output on
        output_start = time.perf_counter()
        keithley.write('OUTPUT ON; :TRACE:TRIG "CMEAS"; :SENSE:FUNCtion "VOLTage"; :TRACE:TRIG "VMEAS"')
        if t_ON > 0:
            while (t_ON - time.perf_counter() + output_start) > 0:
                pass    
        # Set the voltage source output off
        keithley.write('OUTPUT OFF')
        output_end = time.perf_counter()

        current = float(keithley.query('FETCH? "CMEAS"'))
        voltage = float(keithley.query('FETCH? "VMEAS"'))

        times.append((output_end - start_time))
        t_ON_list.append(output_end - output_start)

        currents.append(current)
        voltages.append(voltage)

        plot_start = time.perf_counter()
        clear_output(wait=True)
        print(f'Time: {times[-1]:.2f} s / {duration} s')
        print(f'Voltage: {voltage:.4g} V')
        print(f'Current: {current:.4g} A')
        ylabel, currents_plot = current_set(currents)
        plt.plot(times, currents_plot, linestyle='-', marker='o', color='blue')
        plt.xlabel('Time (s)', fontsize=18)
        plt.ylabel(ylabel, fontsize=18)
        plt.xticks(fontsize=16)
        plt.yticks(fontsize=16)
        plt.grid(True)
        display(fig)
        print(f'Round time: {time.perf_counter()-round_start} s')
        print(f'Plot time: {time.perf_counter()-plot_start} s')

        if (delay_time - (time.perf_counter() - round_start)) > 0:
            time.sleep(delay_time - (time.perf_counter() - round_start))
            print(f'Real Round time: {time.perf_counter()-round_start} s')
        else:
            pass

    # Close the connection to the Keithley 2450
    keithley.write('TRACE:DEL "VMEAS"')
    keithley.write('TRACE:DEL "CMEAS"')
    keithley.close()

    clear_output(wait=True)
    
    # save file, decide file name first
    file_name = input(f'Enter the file name [XXX], it will be "{project_dir}/XXX.csv"\n  [XXX] = ')
    filename = os.path.join(project_dir, file_name + '.csv')
    # Folder to save the directory
    fig_dir = f"{project_dir}/figures"
    fig_path = fig_path = os.path.join(fig_dir, file_name + '.png')
    
    # Save the figure
    fig2 = plt.figure(figsize=(12,8))
    plt.rcParams["font.size"] = 20
    ylabel, currents_plot = current_set(currents)
    plt.plot(times, currents_plot, linestyle='-', marker='o', color='blue')
    plt.xlabel('Time (s)')
    plt.ylabel(ylabel)    
    plt.show()
    fig2.savefig(fig_path, transparent = True)

    # Check if the project directory exists
    overwrite = '' #initialize
    if os.path.exists(filename):
        print(f'\033[33m\nThe file "{file_name}.csv" already exists.')
        overwrite = int_ask('Do you want to overwrite it? Yes(0) or No(1)\n')
        if overwrite == 0:
            print('\033[33mOkay, the program will overwrite the file.\n\033[33m')
            fig_path = os.path.join(fig_dir, file_name + '.png')
        elif overwrite == 1:
            file_name2 = file_name
            while file_name2 == file_name:
                file_name2 = input(f'Enter the file name except for "{file_name}":\n')
            filename = os.path.join(project_dir, file_name2 + '.csv')
            fig_path = os.path.join(fig_dir, file_name2 + '.png')
        
    # Create a CSV file for saving the data
    with open(filename, 'w', newline='') as csvfile:
        csvwriter = csv.writer(csvfile)
        csvwriter.writerow(['Time (s)', 'ON-time (s)', 'Current (A)', 'Voltage (V)'])
        for i in range(len(times)):
            csvwriter.writerow([times[i], t_ON_list[i], currents[i], voltages[i]])

    # Check that the CSV file was created successfully
    try:
        with open(filename, 'r') as csvfile:
            pass
    except:
        print("Error: Could not create CSV file")
        
    print("Program completed")
    
### ---------------- FIGURES -----------------############
def data_list(project_dir):
    
    data_list = glob.glob(f'{project_dir}/*.csv')
    data_list.sort()

    for i, file in enumerate(data_list):
        print(f'{i:02}: {os.path.splitext(os.path.basename(file))[0]}')
        
    return data_list
        
def tI_plot(data_list, plot_index, project_dir):
    
    fig_dir = f"{project_dir}/figures"
    
    plot_list = []
    for i in range(len(plot_index)):
        plot_list.append(data_list[plot_index[i]])

    figure = plt.figure(figsize=(12,8))
    plt.rcParams["font.size"] = 20

    for file in plot_list:
        df = pd.read_csv(file)
        label = os.path.splitext(os.path.basename(file))[0]
        plt.plot(df['Time (s)'], df['Current (A)'], linestyle='-', marker='o', label = label)

    plt.legend(frameon=False)
    # plt.show()

    # Create the project directory if it does not exist
    save = input('save (0) or not (1)?')
    if save == '0':
        fig_name = input('filename')
        fig_com_dir = f"{fig_dir}/combined"
        make_folder(fig_com_dir)
        fig_com_path = os.path.join(fig_com_dir, fig_name + '.png')
        figure.savefig(fig_com_path, transparent = True)
        
def IV_plot(data_list, plot_index, project_dir, xlog = False, ylog = False):
    fig_dir = f"{project_dir}/figures"
    plot_list = []
    for i in range(len(plot_index)):
        plot_list.append(data_list[plot_index[i]])

    figure = plt.figure(figsize=(12,8))
    plt.rcParams["font.size"] = 20

    for file in plot_list:
        df = pd.read_csv(file)
        label = os.path.splitext(os.path.basename(file))[0]
        if ylog:
            y = df['Current (A)'].to_list()
            y = [abs(n) for n in y]
        plt.plot(df['Voltage (V)'], y, linestyle='-', marker='o', label = label)

    plt.legend(frameon=False)
    if xlog:
        plt.xscale('log')
    if ylog:
        plt.yscale('log')
    # plt.show()

    # Create the project directory if it does not exist
    save = input('save (0) or not (1)?')
    if save == '0':
        fig_name = input('filename')
        fig_com_dir = f"{fig_dir}/combined"
        make_folder(fig_com_dir)
        fig_com_path = os.path.join(fig_com_dir, fig_name + '.png')
        figure.savefig(fig_com_path, transparent = True)