import sys
import os

def make_folder(name):
    # Check if the project directory exists
    if os.path.exists(name):
        print(f'\033[33mFolder "{name}" already exists.')

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