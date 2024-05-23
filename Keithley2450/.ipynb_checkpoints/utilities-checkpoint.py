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
