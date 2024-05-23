# Python Script to remote control L100-1S-2 Darwing Peristaltic Pump
# Refer the following website for help
# https://blog.darwin-microfluidics.com/how-to-control-the-longer-l100-1s-2-pump-via-python/
# https://blog.darwin-microfluidics.com/control-command-string-generator-for-longer-peristaltic-pumps/

# The customXOR function is used to strip whitespace characters
# from the PDU and split it by any whitespace in order to compute
# the XOR of the different bytes within the PDU
def customXOR(input_string):
    # strip whitespace characters and split by any whitespace
    hex_values = input_string.strip().split()

    # compute the XOR of the different bytes within the PDU
    result = 0
    for value in hex_values:
        result ^= int(value, 16)

    # return the XOR result of the PDU
    return hex(result)[2:]

def generate_pdu(command_characters,flow_rate,state1,state2):
    # make sure the flow rate is interger
    flow_rate_int = int(flow_rate)
    # convert to hex number
    hex_flow_rate = hex(flow_rate_int)
    # remove the prefix 0x and make 8 digit str
    hex_flow_rate_str = hex_flow_rate[2:].zfill(8)
    # split each 2 digit
    chunks = [hex_flow_rate_str[i:i+2] for i in range(0, len(hex_flow_rate_str), 2)]
    # combine with ''
    formatted_string = ' '.join(chunks)
    # combine all parameters
    all_chunks = [command_characters,chunks[0],chunks[1],chunks[2],chunks[3],state1,state2]
    formatted_string = ' '.join(all_chunks)
    return formatted_string

def generate_fcs(pump_address,length_pdu,pdu):
    xor_pdu = customXOR(pdu)
    # calculate the frame check sequence (hex input and output base)
    frame_check_sequence = "{:x}".format(int(pump_address, 16) ^
    int(length_pdu, 16) ^ int(xor_pdu, 16))
    return frame_check_sequence

# https://blog.darwin-microfluidics.com/control-command-string-generator-for-longer-peristaltic-pumps/
# Set running parameter (flow rate)
def generate_pump_command(flow_rate,state1,state2):
    start_flag = 'E9'
    pdu = generate_pdu(command_characters,flow_rate,state1,state2)
    frame_check_sequence = generate_fcs(pump_address,length_pdu,pdu)

    params = [start_flag,pump_address,length_pdu,pdu,frame_check_sequence]
    hex_values = ' '.join(params).split()
    result = bytearray(int(value, 16) for value in hex_values)

    return hex_values, result

def pump_run(flow_rate,state2):

    